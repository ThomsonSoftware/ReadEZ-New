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
            .RoR_OverlaySize = s.Overlay_Size, .RoR_PresentationTime = s.Timing, .Test_Width = sz.Width, .Test_Height = sz.Height}
    End Function

    ''' <summary>Baseline: symptoms then rate of reading, on white.</summary>
    Public Shared Function Baseline() As TestRun
        Dim t As New TestRun(NewResult(ReadEZTestMode.Baseline)) With {.HasColour = True}
        t.Add("step.symptoms", Function() New SymptomsStep())
        t.Add("step.reading", Function() New ReadingStep())
        Return t
    End Function

    ''' <summary>A colour test: the colour, then (optionally) rate of reading and symptoms with it.</summary>
    Public Shared Function Colour(mode As ReadEZTestMode, method As Integer, continueAfter As Boolean) As TestRun
        Dim t As New TestRun(NewResult(mode)) With {.Method = method}
        t.Add(If(method = 0, "step.colourAuto", "step.colourManual"), Function() New ColourStep())
        If continueAfter Then
            t.Add("step.readingColour", Function() New ReadingStep())
            t.Add("step.symptomsColour", Function() New SymptomsStep())
        End If
        Return t
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
