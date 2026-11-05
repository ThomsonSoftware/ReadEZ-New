Option Strict On

' ====================================================================================================
'   TestRun.vb - the steps of a ReadEZ test (Tests screen > a test tile)
'
'   As ReadEZ 2020's step-by-step guide:
'     Baseline       1. Symptoms on white   2. Rate of reading on white
'     Colour test    1. Determination of the preferred colour (automatic or manual)
'                    2. Rate of reading with the colour   3. Symptoms with the colour
'                    (steps 2 and 3 can be left out: Tests > "Continue with reading and symptoms")
'   Colour tests come in three kinds, which differ in how the colour is shown:
'     Overlays   a coloured panel round the text on a white screen (two side by side to compare)
'     Lenses     the whole screen coloured, in a darkened room (the whole visual field coloured)
'     Computer   the whole screen coloured with the room lights on
'   All steps fill in one ReadEZResult, saved when the last step is finished.
'
'   A colour test uses the text settings of the patient's latest baseline (text, font, size, line
'   spacing, rows), so the reading speeds and symptoms can be compared; changing them during the
'   test asks first (TextOptions).
' ====================================================================================================

Public NotInheritable Class TestRun
    Public ReadOnly Property Result As ReadEZResult
    Public ReadOnly Property Steps As New List(Of Func(Of TestStep))
    Private ReadOnly _titles As New List(Of String)
    Public Property Index As Integer

    ''' <summary>True once a colour has been chosen (or for a baseline, which needs none).</summary>
    Public Property HasColour As Boolean

    ''' <summary>The colour test's method: automatic (0) or manual (1).</summary>
    Public Property Method As Integer

    ''' <summary>Colour tests: the baseline whose text settings this test uses (Nothing = none, or a baseline test).</summary>
    Public Property TextBaseline As ReadEZResult

    ''' <summary>The operator agreed to change the text settings from the baseline's for this test.</summary>
    Public Property TextUnlocked As Boolean

    ''' <summary>Grating tests: the operator agreed to change the grating's spatial frequency (Grating card).</summary>
    Public Property GratingUnlocked As Boolean

    ''' <summary>Grating tests: the frequency the test started with (the one found), for "Back to the one found".</summary>
    Public Property GratingFound As Double

    Private Sub New(r As ReadEZResult)
        Result = r
    End Sub

    Private Sub Add(titleKey As String, factory As Func(Of TestStep))
        Steps.Add(factory)
        _titles.Add(titleKey)
    End Sub

    Public Function StepTitles() As List(Of String)
        Return _titles.Select(Function(k) Loc.T(k)).ToList()
    End Function

    Private Shared Function NewResult(mode As ReadEZTestMode) As ReadEZResult
        Dim s = AppHost.Settings
        Dim sz = StimulusArt.ScreenSize
        Return New ReadEZResult With {
            .IsBaseline = mode = ReadEZTestMode.Baseline, .Mode = mode, .Client_ID = AppHost.Session.PatientId,
            .Screener_ID = AppHost.Session.ScreenerId, .Test_Date = Date.Now, .PaletteSet = If(mode = ReadEZTestMode.Overlay, 1, 2),
            .RoR_FontName = s.RoR_FontName, .RoR_FontSize = s.RoR_FontSize, .RoR_LineSpacing = s.RoR_LineSpacing,
            .RoR_NoRows = s.RoR_NoRows, .RoR_TextOptions = s.RoR_TextOptions, .ROR_Text = ReadingTexts.ByIndex(s.RoR_TextOptions).Name,
            .RoR_OverlaySize = s.Overlay_Size, .RoR_PresentationTime = s.Timing, .Test_Width = sz.Width, .Test_Height = sz.Height,
            .Palette = ReadEZColours.CurrentPalette}
    End Function

    ''' <summary>Baseline: symptoms then rate of reading, on white.</summary>
    Public Shared Function Baseline() As TestRun
        Dim t As New TestRun(NewResult(ReadEZTestMode.Baseline)) With {.HasColour = True}
        t.Add("step.symptoms", Function() New SymptomsStep())
        t.Add("step.reading", Function() New ReadingStep())
        Return t
    End Function

    ''' <summary>
    ''' Baseline pattern: the pattern glare test finds the most uncomfortable striped pattern, then
    ''' the pattern symptoms questions with it on white.
    ''' </summary>
    Public Shared Function PatternGlareTest() As TestRun
        Dim t As New TestRun(NewResult(ReadEZTestMode.Baseline)) With {.HasColour = True}
        t.Result.IsPattern = True
        t.Result.PatternDistanceCm = AppHost.Settings.PatternGlareDistanceCm
        t.Add("step.patternGlare", Function() New PatternGlareStep())
        t.Add("step.symptomsPattern", Function() New SymptomsStep())
        Return t
    End Function

    ''' <summary>
    ''' A colour test with the patient's most uncomfortable pattern (from their baseline pattern)
    ''' instead of the text: the colour, then the pattern symptoms questions with it.
    ''' </summary>
    Public Shared Function ColourPattern(mode As ReadEZTestMode, method As Integer, baseline As ReadEZResult) As TestRun
        Dim t As New TestRun(NewResult(mode)) With {.Method = method}
        t.Result.IsPattern = True
        t.Result.PatternCpd = baseline.PatternCpd
        t.Result.PatternDistanceCm = baseline.PatternDistanceCm
        t.Add(If(method = 0, "step.colourAuto", "step.colourManual"), Function() New ColourStep())
        t.Add("step.symptomsColour", Function() New SymptomsStep())
        Return t
    End Function

    ''' <summary>The patient's latest baseline pattern with a frequency found (Nothing if none).</summary>
    Public Shared Function LatestPatternBaseline() As ReadEZResult
        If Not AppHost.Session.HasPatient Then Return Nothing
        Try
            Return AppHost.Db.GetResults(AppHost.Session.PatientId).Where(Function(x) x.IsBaseline AndAlso x.IsPattern AndAlso x.PatternCpd > 0).
                   OrderByDescending(Function(x) x.Test_Date).ThenByDescending(Function(x) x.ID).FirstOrDefault()
        Catch ex As Exception
            Log.Error("Baseline pattern", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>A colour test: the colour, then (optionally) rate of reading and symptoms with it.</summary>
    Public Shared Function Colour(mode As ReadEZTestMode, method As Integer, continueAfter As Boolean) As TestRun
        Dim t As New TestRun(NewResult(mode)) With {.Method = method}
        UseBaselineText(t)
        t.Add(If(method = 0, "step.colourAuto", "step.colourManual"), Function() New ColourStep())
        If continueAfter Then
            t.Add("step.readingColour", Function() New ReadingStep())
            t.Add("step.symptomsColour", Function() New SymptomsStep())
        End If
        Return t
    End Function

    ''' <summary>Copies the text settings of the patient's latest baseline into a colour test.</summary>
    Private Shared Sub UseBaselineText(t As TestRun)
        If Not AppHost.Session.HasPatient Then Return
        Dim b As ReadEZResult = Nothing
        Try
            b = AppHost.Db.GetResults(AppHost.Session.PatientId).Where(Function(x) x.IsBaseline).
                OrderByDescending(Function(x) x.Test_Date).ThenByDescending(Function(x) x.ID).FirstOrDefault()
        Catch ex As Exception
            Log.Error("Baseline text settings", ex)
        End Try
        If b Is Nothing Then Return
        t.TextBaseline = b
        Dim r = t.Result
        r.RoR_TextOptions = b.RoR_TextOptions
        r.ROR_Text = ReadingTexts.ByIndex(b.RoR_TextOptions).Name
        If Not String.IsNullOrWhiteSpace(b.RoR_FontName) Then r.RoR_FontName = b.RoR_FontName
        r.RoR_FontSize = b.RoR_FontSize
        r.RoR_LineSpacing = b.RoR_LineSpacing
        If b.RoR_NoRows > 0 Then r.RoR_NoRows = b.RoR_NoRows
    End Sub

    ''' <summary>
    ''' Translation key of a test's kind, as the results show it: Baseline text, Baseline pattern,
    ''' Overlays, Overlay pattern test ...
    ''' </summary>
    Public Shared Function ResultKey(r As ReadEZResult) As String
        If r.IsPattern Then
            If r.IsBaseline Then Return "kind.baselinePattern"
            Select Case r.Mode
                Case ReadEZTestMode.Overlay : Return "kind.overlayPattern"
                Case ReadEZTestMode.Lenses : Return "kind.lensesPattern"
                Case Else : Return "kind.computerPattern"
            End Select
        End If
        Return If(r.IsBaseline, "kind.baselineText", ModeKey(r.Mode))
    End Function

    ''' <summary>Translation key of a test mode's name.</summary>
    Public Shared Function ModeKey(mode As ReadEZTestMode) As String
        Select Case mode
            Case ReadEZTestMode.Overlay : Return "mode.overlay"
            Case ReadEZTestMode.Lenses : Return "mode.lenses"
            Case ReadEZTestMode.Computer : Return "mode.computer"
            Case Else : Return "mode.baseline"
        End Select
    End Function
End Class
