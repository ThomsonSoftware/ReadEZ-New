Option Strict On

' ====================================================================================================
'   ColourStaircase.vb - the automatic colour tests
'
'     ColourStaircase     the classic test (ReadEZ.vb: AutoScreen and CompareWinnersTest), below
'     QuickColourSearch   the quick test (a survey of opposite pairs, then adaptive comparisons
'                         scored with a Bradley-Terry model), at the end of this file
'   Both implement IColourSearch, so the colour step runs either.
'
'   The classic test:
'
'   Thomson Software Solutions' algorithm, as in ReadEZ 2020:
'     Phase 1  three interleaved staircases, starting from colours A, E and H. Each presentation
'              pairs a staircase's current favourite with a colour (1 - 12) it hasn't shown yet,
'              on a random side; the colour preferred becomes the favourite. 11 presentations per
'              staircase, 33 in all.
'     Phase 2  the three favourites are compared (1 with 2, the winner with 3; skipped when two
'              staircases found the same colour), then the winner with grey.
'     Phase 3  lenses and computer only: the winner in the light and the dark lens sets.
'   About 35 presentations. The result: the preferred colour (Colour1), whether grey was preferred
'   to it, the palette (1 overlays, 2 lenses, 3 dark lenses) and how often each colour was chosen
'   (the "spokes" on the colour wheel - one long spoke for a clear preference, many short ones for
'   an inconsistent patient).
'
'   Changes: the response to the last presentation of phase 1 is now counted (ReadEZ 2020 lost it
'   when moving to phase 2), and Repeat simply shows the same pair again.
' ====================================================================================================

Public NotInheritable Class ColourStaircase
    Implements IColourSearch

    Public Const PhaseOnePresentations As Integer = 33

    Public Enum Phase
        Staircases = 1
        Winners = 2
        Saturation = 3
        Done = 4
    End Enum

    Private ReadOnly _set As ReadEZSet
    Private ReadOnly _lenses As Boolean
    Private ReadOnly _rnd As New Random()
    Private ReadOnly _old(2) As Integer
    Private ReadOnly _presented(12, 2) As Integer
    Private _stair As Integer = -1
    Private _winnerStep As Integer

    ''' <summary>How often each colour (0 grey, 1 - 12) was preferred.</summary>
    Public ReadOnly Property Chosen As Integer() = New Integer(12) {}
    Public Property CurrentPhase As Phase = Phase.Staircases
    Public Property Count As Integer

    ' The pair on show: Left/first and Right/second, as colour numbers and sets
    Public Property LeftIndex As Integer
    Public Property RightIndex As Integer
    Public Property LeftSet As ReadEZSet
    Public Property RightSet As ReadEZSet

    ' The answer
    Public Property Winner As Integer
    Public Property GreyPreferred As Boolean
    Public Property PaletteSet As Integer

    Private _oldIndex, _newIndex As Integer
    Private _oldOnLeft As Boolean

    ''' <param name="mode">Overlay (overlay colours) or Lenses / Computer (lens colours, then light or dark).</param>
    Public Sub New(mode As ReadEZTestMode)
        _lenses = mode <> ReadEZTestMode.Overlay
        _set = If(_lenses, ReadEZSet.Lens, ReadEZSet.Overlay)
        PaletteSet = If(_lenses, 2, 1)
        _old(0) = 1 : _presented(1, 0) = 1
        _old(1) = 5 : _presented(5, 1) = 1
        _old(2) = 8 : _presented(8, 2) = 1
        NextStaircasePair()
    End Sub

    Public ReadOnly Property LeftColour As Color
        Get
            Return ReadEZColours.Colour(LeftSet, LeftIndex)
        End Get
    End Property

    Public ReadOnly Property RightColour As Color
        Get
            Return ReadEZColours.Colour(RightSet, RightIndex)
        End Get
    End Property

    ''' <summary>About how many presentations there are in all (for the progress shown).</summary>
    Public ReadOnly Property Expected As Integer
        Get
            Return PhaseOnePresentations + 3 + If(_lenses, 1, 0)
        End Get
    End Property

    Private Sub Show(oldIndex As Integer, newIndex As Integer, Optional oldSet As ReadEZSet? = Nothing, Optional newSet As ReadEZSet? = Nothing)
        _oldIndex = oldIndex
        _newIndex = newIndex
        _oldOnLeft = _rnd.Next(2) = 0
        Dim os = If(oldSet, _set), ns = If(newSet, _set)
        If _oldOnLeft Then
            LeftIndex = oldIndex : LeftSet = os : RightIndex = newIndex : RightSet = ns
        Else
            LeftIndex = newIndex : LeftSet = ns : RightIndex = oldIndex : RightSet = os
        End If
        Count += 1
    End Sub

    Private Sub NextStaircasePair()
        _stair = (_stair + 1) Mod 3
        Dim candidates = Enumerable.Range(1, 12).Where(Function(c) _presented(c, _stair) = 0).ToList()
        If candidates.Count = 0 Then
            ' (Can't happen with 11 per staircase, but never loop for ever)
            candidates = Enumerable.Range(1, 12).Where(Function(c) c <> _old(_stair)).ToList()
        End If
        Dim pick = candidates(_rnd.Next(candidates.Count))
        _presented(pick, _stair) += 1
        Show(_old(_stair), pick)
    End Sub

    ''' <summary>The patient prefers the left / first colour (True) or the right / second (False).</summary>
    Public Sub Respond(preferLeft As Boolean)
        If CurrentPhase = Phase.Done Then Return
        Dim newPreferred = preferLeft <> _oldOnLeft
        Select Case CurrentPhase
            Case Phase.Staircases
                Dim pick = If(newPreferred, _newIndex, _oldIndex)
                Chosen(pick) += 1
                If newPreferred Then _old(_stair) = _newIndex
                If Count < PhaseOnePresentations Then
                    NextStaircasePair()
                Else
                    CurrentPhase = Phase.Winners
                    _winnerStep = 0
                    NextWinnerPair()
                End If

            Case Phase.Winners
                If _winnerStep <= 2 Then
                    ' Comparing staircase winners (steps 1 and 2)
                    Chosen(If(newPreferred, _newIndex, _oldIndex)) += 1
                    If newPreferred Then _old(0) = _newIndex
                    NextWinnerPair()
                Else
                    ' Winner against grey
                    GreyPreferred = newPreferred
                    If newPreferred Then Chosen(0) += 1
                    Winner = _old(0)
                    If _lenses Then
                        CurrentPhase = Phase.Saturation
                        Show(Winner, Winner, ReadEZSet.Lens, ReadEZSet.LensStrong)
                    Else
                        CurrentPhase = Phase.Done
                    End If
                End If

            Case Phase.Saturation
                PaletteSet = If(newPreferred, 3, 2)
                CurrentPhase = Phase.Done
        End Select
    End Sub

    ''' <summary>Winner of staircase 0 against 1, then against 2, then against grey (same colours skipped).</summary>
    Private Sub NextWinnerPair()
        Do
            _winnerStep += 1
            Select Case _winnerStep
                Case 1, 2
                    Dim other = _old(_winnerStep)
                    If other <> _old(0) Then
                        Show(_old(0), other)
                        Return
                    End If
                Case Else
                    _winnerStep = 3
                    Show(_old(0), 0)
                    Return
            End Select
        Loop
    End Sub

    ''' <summary>The colour counts as saved with the result ("Grey:0,A:5,B:1,...").</summary>
    Public Function ScoreText() As String
        Return String.Join(",", Enumerable.Range(0, 13).Select(Function(i) If(i = 0, "Grey", ChrW(AscW("A"c) + i - 1).ToString()) & ":" & Chosen(i)))
    End Function

    ''' <summary>Reads ScoreText back (for the colour wheel of a saved result).</summary>
    Public Shared Function ParseScore(text As String) As Integer()
        Dim counts = New Integer(12) {}
        If String.IsNullOrWhiteSpace(text) Then Return counts
        For Each part In text.Split(","c)
            Dim kv = part.Split(":"c)
            If kv.Length <> 2 Then Continue For
            Dim n As Integer
            If Not Integer.TryParse(kv(1), n) Then Continue For
            Dim key = kv(0).Trim()
            If key.Equals("Grey", StringComparison.OrdinalIgnoreCase) Then
                counts(0) = n
            ElseIf key.Length = 1 AndAlso key(0) >= "A"c AndAlso key(0) <= "L"c Then
                counts(AscW(key(0)) - AscW("A"c) + 1) = n
            End If
        Next
        Return counts
    End Function
#Region "IColourSearch"

    Private ReadOnly Property IColourSearch_IsDone As Boolean Implements IColourSearch.IsDone
        Get
            Return CurrentPhase = Phase.Done
        End Get
    End Property

    Private ReadOnly Property IColourSearch_StageNumber As Integer Implements IColourSearch.StageNumber
        Get
            Return CInt(CurrentPhase)
        End Get
    End Property

    Private ReadOnly Property IColourSearch_Count As Integer Implements IColourSearch.Count
        Get
            Return Count
        End Get
    End Property

    Private ReadOnly Property IColourSearch_Expected As Integer Implements IColourSearch.Expected
        Get
            Return Expected
        End Get
    End Property

    Private ReadOnly Property IColourSearch_Chosen As Integer() Implements IColourSearch.Chosen
        Get
            Return Chosen
        End Get
    End Property

    Private ReadOnly Property IColourSearch_LeftColour As Color Implements IColourSearch.LeftColour
        Get
            Return LeftColour
        End Get
    End Property

    Private ReadOnly Property IColourSearch_RightColour As Color Implements IColourSearch.RightColour
        Get
            Return RightColour
        End Get
    End Property

    Private ReadOnly Property IColourSearch_Winner As Integer Implements IColourSearch.Winner
        Get
            Return Winner
        End Get
    End Property

    Private ReadOnly Property IColourSearch_GreyPreferred As Boolean Implements IColourSearch.GreyPreferred
        Get
            Return GreyPreferred
        End Get
    End Property

    Private ReadOnly Property IColourSearch_PaletteSet As Integer Implements IColourSearch.PaletteSet
        Get
            Return PaletteSet
        End Get
    End Property

    ''' <summary>The classic test always shows its own pairs (pointing at a colour is the quick test's).</summary>
    Private ReadOnly Property IColourSearch_CanJumpTo As Boolean Implements IColourSearch.CanJumpTo
        Get
            Return False
        End Get
    End Property

    Private Sub IColourSearch_Respond(preferLeft As Boolean) Implements IColourSearch.Respond
        Respond(preferLeft)
    End Sub

    Private Sub IColourSearch_JumpTo(index As Integer) Implements IColourSearch.JumpTo
    End Sub

    Private Function IColourSearch_ScoreText() As String Implements IColourSearch.ScoreText
        Return ScoreText()
    End Function

#End Region
End Class

' ====================================================================================================
'   The automatic colour tests share one interface, so the colour step can run either.
' ====================================================================================================

Public Interface IColourSearch
    ReadOnly Property IsDone As Boolean
    ''' <summary>The stage shown in the progress text (1, 2, ...).</summary>
    ReadOnly Property StageNumber As Integer
    ReadOnly Property Count As Integer
    ReadOnly Property Expected As Integer
    ''' <summary>How often each colour (0 grey, 1 - 12) was preferred (the wheel's spokes).</summary>
    ReadOnly Property Chosen As Integer()
    ReadOnly Property LeftColour As Color
    ReadOnly Property RightColour As Color
    ReadOnly Property Winner As Integer
    ReadOnly Property GreyPreferred As Boolean
    ''' <summary>1 overlays, 2 lenses, 3 dark lenses.</summary>
    ReadOnly Property PaletteSet As Integer
    ''' <summary>True if the patient can point at a colour on the wheel (the search goes on from it).</summary>
    ReadOnly Property CanJumpTo As Boolean
    Sub Respond(preferLeft As Boolean)
    ''' <summary>The patient (or the operator for them) picked this colour directly.</summary>
    Sub JumpTo(index As Integer)
    Function ScoreText() As String
End Interface

''' <summary>
''' The quick automatic colour test: adaptive paired comparisons.
'''
'''   Every answer is kept, and after each one every colour gets a "comfort score" fitted to all the
'''   answers so far (a Bradley-Terry model: the bigger the difference in score, the more surely one
'''   colour is preferred to the other). Colours next to each other on the colour circle are
'''   assumed to be liked a little alike, which lets a few answers say something about the
'''   neighbours too, but not so strongly that a patient who likes two separate colours is missed.
'''
'''     Stage 1  Survey (6)     the six opposite pairs (A/G, B/H, C/I, D/J, E/K, F/L), in random
'''                             order and sides: every colour is seen once.
'''     Stage 2  Search         the colour with the best score (the leader) against the colour most
'''                             likely to beat it: a high score, or one that has not been seen much.
'''                             It stops when the leader has stayed the leader for 4 answers in a row
'''                             (at least 12 presentations), or after 22.
'''     Stage 3  Grey (1)       the winner against grey, as in the classic test.
'''     Stage 4  Saturation (1) lenses / computer only: the winner light against dark.
'''
'''   About 14 - 17 presentations instead of about 35. In simulated patients (smooth, sharp, broad,
'''   irregular and two-peaked preferences, with different amounts of inconsistency) it found the
'''   same colour as the classic test within a few per cent, in under half the time; it should
'''   still be checked against the classic test with real patients before it replaces it.
'''
'''   The patient can also point at a colour (JumpTo): it counts as one win against every other
'''   colour and the search carries on from it, so a direct choice is still tested against its
'''   rivals and can be overturned.
''' </summary>
Public NotInheritable Class QuickColourSearch
    Implements IColourSearch

    Public Enum Stage
        Survey = 1
        Search = 2
        Grey = 3
        Saturation = 4
        Done = 5
    End Enum

    ' Tuned in simulation (see the summary)
    Private Const MinPresentations As Integer = 12
    Private Const MaxPresentations As Integer = 22
    Private Const StableAnswers As Integer = 4
    Private Const Smoothness As Double = 0.1          ' how much neighbours on the circle are assumed alike
    Private Const Shrink As Double = 0.05             ' keeps the scores finite
    Private Const Exploration As Double = 1.2         ' how much a colour seen less often is tried

    Private ReadOnly _set As ReadEZSet
    Private ReadOnly _lenses As Boolean
    Private ReadOnly _rnd As New Random()
    Private ReadOnly _chosen As Integer() = New Integer(12) {}
    Private ReadOnly _shown As Integer() = New Integer(12) {}
    Private ReadOnly _survey As New List(Of (A As Integer, B As Integer))
    Private ReadOnly _games As New List(Of (Winner As Integer, Loser As Integer))
    Private _score As Double() = New Double(12) {}

    Private _stage As Stage = Stage.Survey
    Private _count As Integer
    Private _searchCount As Integer
    Private _leader As Integer
    Private _same As Integer
    Private _winner As Integer
    Private _grey As Boolean
    Private _palette As Integer

    Private _leftIndex, _rightIndex As Integer
    Private _leftSet, _rightSet As ReadEZSet

    ''' <summary>The colour the patient pointed at (0 = none).</summary>
    Public Property PointedAt As Integer

    Public Sub New(mode As ReadEZTestMode)
        _lenses = mode <> ReadEZTestMode.Overlay
        _set = If(_lenses, ReadEZSet.Lens, ReadEZSet.Overlay)
        _palette = If(_lenses, 2, 1)
        For i = 1 To 6
            _survey.Add((i, i + 6))
        Next
        For i = _survey.Count - 1 To 1 Step -1
            Dim j = _rnd.Next(i + 1)
            Dim t = _survey(i) : _survey(i) = _survey(j) : _survey(j) = t
        Next
        ShowNext()
    End Sub

#Region "Interface"

    Public ReadOnly Property IsDone As Boolean Implements IColourSearch.IsDone
        Get
            Return _stage = Stage.Done
        End Get
    End Property

    Public ReadOnly Property StageNumber As Integer Implements IColourSearch.StageNumber
        Get
            Return CInt(_stage)
        End Get
    End Property

    Public ReadOnly Property Count As Integer Implements IColourSearch.Count
        Get
            Return _count
        End Get
    End Property

    Public ReadOnly Property Expected As Integer Implements IColourSearch.Expected
        Get
            Return Math.Max(_count + If(_stage = Stage.Done, 0, 1), 15 + 1 + If(_lenses, 1, 0))
        End Get
    End Property

    Public ReadOnly Property Chosen As Integer() Implements IColourSearch.Chosen
        Get
            Return _chosen
        End Get
    End Property

    Public ReadOnly Property LeftColour As Color Implements IColourSearch.LeftColour
        Get
            Return ReadEZColours.Colour(_leftSet, _leftIndex)
        End Get
    End Property

    Public ReadOnly Property RightColour As Color Implements IColourSearch.RightColour
        Get
            Return ReadEZColours.Colour(_rightSet, _rightIndex)
        End Get
    End Property

    Public ReadOnly Property Winner As Integer Implements IColourSearch.Winner
        Get
            Return _winner
        End Get
    End Property

    Public ReadOnly Property GreyPreferred As Boolean Implements IColourSearch.GreyPreferred
        Get
            Return _grey
        End Get
    End Property

    Public ReadOnly Property PaletteSet As Integer Implements IColourSearch.PaletteSet
        Get
            Return _palette
        End Get
    End Property

    Public ReadOnly Property CanJumpTo As Boolean Implements IColourSearch.CanJumpTo
        Get
            Return _stage = Stage.Survey OrElse _stage = Stage.Search
        End Get
    End Property

#End Region

#Region "Scores"

    Private Shared Function Wrap(i As Integer) As Integer
        Return ((i - 1 + 12) Mod 12) + 1
    End Function

    ''' <summary>Fits the comfort scores to all the answers (maximum a posteriori Bradley-Terry, by gradient ascent).</summary>
    Private Sub FitScores()
        Dim v = New Double(12) {}
        Dim grad = New Double(12) {}
        For iteration = 1 To 150
            Array.Clear(grad, 0, grad.Length)
            For Each gm In _games
                Dim pr = 1.0 / (1.0 + Math.Exp(-(v(gm.Winner) - v(gm.Loser))))
                grad(gm.Winner) += 1 - pr
                grad(gm.Loser) -= 1 - pr
            Next
            For i = 1 To 12
                Dim j = Wrap(i + 1)
                Dim d = v(i) - v(j)
                grad(i) -= 2 * Smoothness * d
                grad(j) += 2 * Smoothness * d
                grad(i) -= 2 * Shrink * v(i)
            Next
            For i = 1 To 12
                v(i) += 0.3 * grad(i)
            Next
        Next
        _score = v
    End Sub

    Private Function Best() As Integer
        Dim top = Enumerable.Range(1, 12).Max(Function(i) _score(i))
        Dim ties = Enumerable.Range(1, 12).Where(Function(i) _score(i) >= top - 0.000001).ToList()
        Return ties(_rnd.Next(ties.Count))
    End Function

#End Region

#Region "Presentations"

    Private Sub Show(a As Integer, b As Integer, Optional setA As ReadEZSet? = Nothing, Optional setB As ReadEZSet? = Nothing)
        Dim sa = If(setA, _set), sb = If(setB, _set)
        If _rnd.Next(2) = 0 Then
            _leftIndex = a : _leftSet = sa : _rightIndex = b : _rightSet = sb
        Else
            _leftIndex = b : _leftSet = sb : _rightIndex = a : _rightSet = sa
        End If
        _count += 1
    End Sub

    Private Sub ShowNext()
        Do
            Select Case _stage
                Case Stage.Survey
                    If _survey.Count > 0 Then
                        Dim pr = _survey(0)
                        _survey.RemoveAt(0)
                        Show(pr.A, pr.B)
                        Return
                    End If
                    _stage = Stage.Search
                    _leader = 0
                    _same = 0

                Case Stage.Search
                    FitScores()
                    Dim lead = Best()
                    _same = If(lead = _leader, _same + 1, 0)
                    _leader = lead
                    Dim enough = (_searchCount + 6 >= MinPresentations AndAlso _same >= StableAnswers) OrElse _searchCount + 6 >= MaxPresentations
                    If Not enough Then
                        ' The challenger: the colour most likely to beat the leader (high score or seen little)
                        Dim challenger = Enumerable.Range(1, 12).Where(Function(i) i <> lead).
                                         OrderByDescending(Function(i) _score(i) + Exploration / Math.Sqrt(1 + _shown(i)) + _rnd.NextDouble() * 0.000001).First()
                        _searchCount += 1
                        Show(lead, challenger)
                        Return
                    End If
                    _winner = lead
                    _stage = Stage.Grey

                Case Stage.Grey
                    Show(_winner, 0)
                    Return

                Case Stage.Saturation
                    Show(_winner, _winner, ReadEZSet.Lens, ReadEZSet.LensStrong)
                    Return

                Case Else
                    Return
            End Select
        Loop
    End Sub

#End Region

#Region "Answers"

    Public Sub Respond(preferLeft As Boolean) Implements IColourSearch.Respond
        If _stage = Stage.Done Then Return
        Dim pick = If(preferLeft, _leftIndex, _rightIndex)
        Dim other = If(preferLeft, _rightIndex, _leftIndex)
        Select Case _stage
            Case Stage.Survey, Stage.Search
                _chosen(pick) += 1
                _shown(pick) += 1
                _shown(other) += 1
                _games.Add((pick, other))
            Case Stage.Grey
                _grey = pick = 0
                If _grey Then _chosen(0) += 1
                _stage = If(_lenses, Stage.Saturation, Stage.Done)
            Case Stage.Saturation
                _palette = If(If(preferLeft, _leftSet, _rightSet) = ReadEZSet.LensStrong, 3, 2)
                _stage = Stage.Done
        End Select
        ShowNext()
    End Sub

    ''' <summary>The patient points at a colour: one win against every other colour; the search goes on from it.</summary>
    Public Sub JumpTo(index As Integer) Implements IColourSearch.JumpTo
        If Not CanJumpTo OrElse index < 1 OrElse index > 12 Then Return
        PointedAt = index
        _chosen(index) += 1
        For i = 1 To 12
            If i <> index Then _games.Add((index, i))
        Next
        _survey.Clear()
        _stage = Stage.Search
        _leader = 0
        _same = 0
        _searchCount = Math.Max(_searchCount, MinPresentations - 6 - StableAnswers)  ' a few more answers confirm it
        ShowNext()
    End Sub

    ''' <summary>The colour counts as saved with the result ("Grey:0,A:5,...,Method:Quick").</summary>
    Public Function ScoreText() As String Implements IColourSearch.ScoreText
        Dim parts = Enumerable.Range(0, 13).Select(Function(i) If(i = 0, "Grey", ChrW(AscW("A"c) + i - 1).ToString()) & ":" & _chosen(i)).ToList()
        parts.Add("Method:Quick")
        If PointedAt > 0 Then parts.Add("Pointed:" & ChrW(AscW("A"c) + PointedAt - 1))
        Return String.Join(",", parts)
    End Function

#End Region

End Class
