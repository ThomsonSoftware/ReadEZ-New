Option Strict On

' ====================================================================================================
'   ColourStaircase.vb - the automatic colour test (ReadEZ.vb: AutoScreen and CompareWinnersTest)
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
End Class
