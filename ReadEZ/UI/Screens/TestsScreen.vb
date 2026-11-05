Option Strict On

' ====================================================================================================
'   TestsScreen.vb - choose a test (replaces the ribbon's Baseline / Overlays / Lenses / Computer
'   buttons and RZfrmStart)
'
'   Pattern (first)
'     Pattern glare         the baseline pattern: which striped pattern is most uncomfortable (narrowed
'                           down to the worst spatial frequency), then symptoms with it on white
'     Overlays / Lenses / Computer screen - pattern
'                           the colour tests with that pattern instead of the text, then symptoms
'   Text
'     Baseline              symptoms and rate of reading on white
'     Overlays              the colour for an overlay or reading guide
'     Lenses                the colour for tinted lenses (whole screen coloured, darkened room)
'     Computer              the colour for a computer screen (whole screen coloured, room lit)
'
'   With each colour test: the method (automatic or manual) and whether rate of reading and
'   symptoms with the colour follow straight on. Tests switched off in Admin are hidden.
'   A banner says what is missing first (a patient; the colours not checked on this screen).
' ====================================================================================================

Public Class TestsScreen
    Inherits ScreenBase

    Private ReadOnly _flow As New FlowLayoutPanel
    Private ReadOnly _banner As New Label
    Private ReadOnly segMethod As New SegmentedControl
    Private ReadOnly tglContinue As New ToggleSwitch With {.Width = 360}
    Private ReadOnly lblScreener As New Label With {.AutoSize = True, .Font = Fonts.Body}

    Public Sub New()
        TitleKey = "tests.title"
        HelpTopic = "screen.tests"
        SubtitleKey = "tests.subtitle"

        _banner.Dock = DockStyle.Top
        _banner.Height = 36
        _banner.TextAlign = ContentAlignment.MiddleLeft
        _banner.Font = Fonts.Strong
        _banner.Padding = New Padding(16, 0, 0, 0)
        _banner.Cursor = Cursors.Hand
        LabelStyle.SetCustom(_banner)
        AddHandler _banner.Click, Sub() BannerClicked()

        ' Options for the colour tests
        Dim opts As New Card With {.Dock = DockStyle.Top, .Height = 128, .Padding = New Padding(20, 48, 20, 10)}
        opts.SetTitle("tests.options")
        Dim row = Ui.Row(wrap:=True)
        row.Dock = DockStyle.Fill
        segMethod.SetOptions("colour.auto", "colour.manual")
        segMethod.Margin = New Padding(0, 0, 28, 0)
        AddHandler segMethod.SelectedIndexChanged, Sub()
                                                       AppHost.Settings.ColourTestMethod = segMethod.SelectedIndex
                                                       AppHost.SaveSettings()
                                                   End Sub
        row.Controls.Add(Ui.Stack("tests.method", segMethod))
        Loc.Bind(tglContinue, "tests.continue")
        tglContinue.Margin = New Padding(0, 24, 28, 0)
        AddHandler tglContinue.Toggled, Sub()
                                            AppHost.Settings.ContinueAfterColour = tglContinue.IsOn
                                            AppHost.SaveSettings()
                                        End Sub
        row.Controls.Add(tglContinue)
        Dim screenerRow = Ui.Row()
        screenerRow.Margin = New Padding(0, 22, 0, 0)
        lblScreener.Margin = New Padding(0, 8, 8, 0)
        screenerRow.Controls.Add(lblScreener)
        Dim change = ModernButton.Create("tests.changeScreener", ButtonKind.Ghost, Glyphs.Person)
        AddHandler change.Click, Sub()
                                     AppCommands.PickScreener()
                                     UpdateBanner()
                                 End Sub
        screenerRow.Controls.Add(change)
        row.Controls.Add(screenerRow)
        opts.Controls.Add(row)

        _flow.Dock = DockStyle.Fill
        _flow.AutoScroll = True
        _flow.Padding = New Padding(16, 16, 0, 0)
        _flow.BackColor = Color.Transparent
        AddHandler _flow.Resize, Sub() FitRows()

        Content.Controls.Add(_flow)
        Content.Controls.Add(New Panel With {.Dock = DockStyle.Top, .Height = 12, .BackColor = Color.Transparent})
        Content.Controls.Add(opts)
        Content.Controls.Add(_banner)
        AddHandler AppHost.Session.PatientChanged, Sub() UpdateBanner()
        AddHandler AppHost.Session.ScreenerChanged, Sub() UpdateBanner()
        AddHandler ReadEZColours.Changed, Sub() UpdateBanner()
        AddHandler ThemeManager.ThemeChanged, Sub() UpdateBanner()
        Loc.OnChange(Me, AddressOf UpdateBanner)
    End Sub

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        segMethod.SelectedIndex = Math.Max(0, Math.Min(1, AppHost.Settings.ColourTestMethod))
        tglContinue.SetSilently(AppHost.Settings.ContinueAfterColour)
        BuildTiles()
        UpdateBanner()
    End Sub

    ''' <summary>The rows of tiles (one per group), kept as wide as the screen so they wrap.</summary>
    Private ReadOnly _rows As New List(Of FlowLayoutPanel)

    Private Sub BuildTiles()
        _flow.SuspendLayout()
        For Each c In _flow.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        _rows.Clear()
        ' Top to bottom: a heading, then its tiles in a row of their own (which wraps)
        _flow.FlowDirection = FlowDirection.TopDown
        _flow.WrapContents = False
        Dim colourModes = {ReadEZTestMode.Overlay, ReadEZTestMode.Lenses, ReadEZTestMode.Computer}.Where(Function(m) AppHost.Admin.IsTestEnabled(m)).ToList()
        For Each grating In {True, False}
            Dim head = Ui.Section(If(grating, "tests.group.pattern", "tests.group.text"))
            head.Margin = New Padding(4, If(_rows.Count = 0, 0, 16), 0, 8)
            _flow.Controls.Add(head)
            Dim row As New FlowLayoutPanel With {.AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .WrapContents = True,
                                                 .FlowDirection = FlowDirection.LeftToRight, .Margin = Padding.Empty, .Padding = Padding.Empty,
                                                 .BackColor = Color.Transparent}
            Dim tiles As New List(Of (TitleKey As String, DescKey As String, Glyph As String, Mode As ReadEZTestMode))
            If grating Then
                tiles.Add(("mode.patternGlare", "tests.desc.patternGlare", PatternGlareStep.Glyph, ReadEZTestMode.Baseline))
            Else
                tiles.Add(("mode.baseline", "tests.desc.baseline", Glyphs.Font, ReadEZTestMode.Baseline))
            End If
            For Each m In colourModes
                Dim name = If(m = ReadEZTestMode.Overlay, "overlay", If(m = ReadEZTestMode.Lenses, "lenses", "computer"))
                Dim glyph = If(m = ReadEZTestMode.Overlay, Glyphs.Color, If(m = ReadEZTestMode.Lenses, Glyphs.Eye, Glyphs.Picture))
                If grating Then
                    tiles.Add(("mode." & name & "Pattern", "tests.desc." & name & "Pattern", glyph, m))
                Else
                    tiles.Add((TestRun.ModeKey(m), "tests.desc." & name, glyph, m))
                End If
            Next
            For Each t In tiles
                Dim tile = TileButton.Create(t.TitleKey, t.DescKey, t.Glyph)
                tile.Size = New Size(300, 150)
                Dim mode = t.Mode
                If grating Then
                    AddHandler tile.Click, Sub() LaunchPattern(mode)
                Else
                    AddHandler tile.Click, Sub() Launch(mode)
                End If
                row.Controls.Add(tile)
            Next
            _rows.Add(row)
            _flow.Controls.Add(row)
        Next
        Dpi.ScaleNew(_flow, Me)
        ThemeManager.Apply(_flow)
        FitRows()
        _flow.ResumeLayout(True)
    End Sub

    ''' <summary>Each row as wide as the screen (so the tiles wrap onto a second line when there is no room).</summary>
    Private Sub FitRows()
        Dim w = Math.Max(Dpi.Px(Me, 320), _flow.ClientSize.Width - _flow.Padding.Horizontal - Dpi.Px(Me, 8))
        For Each row In _rows
            row.MaximumSize = New Size(w, 0)
        Next
    End Sub

    Private Function ReadyToTest() As Boolean
        If Not AppHost.Session.HasPatient Then
            Ask.Info(FindForm(), "tests.needPatient")
            AppHost.MainWindow?.ShowScreen(MainForm.ScreenPatients)
            Return False
        End If
        ' Every test needs a screener: choose one (or add one) first
        If AppHost.Session.ScreenerId <= 0 Then AppCommands.PickScreener()
        If AppHost.Session.ScreenerId <= 0 Then
            Ask.Info(FindForm(), "tests.needScreener")
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' The pattern tests: the baseline pattern (pattern glare), or a colour test with the pattern found
    ''' there - which needs a baseline pattern first (offered if there is none).
    ''' </summary>
    Private Sub LaunchPattern(mode As ReadEZTestMode)
        If Not ReadyToTest() Then Return
        If mode = ReadEZTestMode.Baseline Then
            AppHost.MainWindow?.RunTest(TestRun.PatternGlareTest())
            Return
        End If
        Dim b = TestRun.LatestPatternBaseline()
        If b Is Nothing Then
            If Ask.YesNo(FindForm(), "tests.needPatternBaseline") Then AppHost.MainWindow?.RunTest(TestRun.PatternGlareTest())
            Return
        End If
        AppHost.MainWindow?.RunTest(TestRun.ColourPattern(mode, AppHost.Settings.ColourTestMethod, b))
    End Sub

    Private Sub Launch(mode As ReadEZTestMode)
        If Not ReadyToTest() Then Return
        Dim run = If(mode = ReadEZTestMode.Baseline, TestRun.Baseline(),
                     TestRun.Colour(mode, AppHost.Settings.ColourTestMethod, AppHost.Settings.ContinueAfterColour))
        AppHost.MainWindow?.RunTest(run)
    End Sub

    Private Sub UpdateBanner()
        Dim p = ThemeManager.Current
        Dim s = AppHost.Session
        Dim name = s.ScreenerName
        lblScreener.Text = If(name = "", Loc.T("tests.noScreener"), Loc.F("tests.screener", name))
        Dim checkedCount = ReadEZColours.CheckedCount()
        If Not s.HasPatient Then
            _banner.Text = Loc.T("tests.banner.noPatient")
            _banner.BackColor = p.WarningSoft
            _banner.ForeColor = p.Ink
            _banner.Tag = "patients"
        ElseIf checkedCount < 3 * ReadEZColours.PerSet Then
            _banner.Text = Loc.F("tests.banner.colours", s.PatientName, checkedCount, 3 * ReadEZColours.PerSet)
            _banner.BackColor = p.WarningSoft
            _banner.ForeColor = p.Ink
            _banner.Tag = "colours"
        Else
            _banner.Text = Loc.F("tests.banner.ready", s.PatientName, Loc.T(If(ScreenLayout.IsTwoScreenMode(), "top.screens.two", "top.screens.one")))
            _banner.BackColor = p.SuccessSoft
            _banner.ForeColor = p.Ink
            _banner.Tag = Nothing
        End If
    End Sub

    Private Sub BannerClicked()
        Select Case TryCast(_banner.Tag, String)
            Case "patients" : AppHost.MainWindow?.ShowScreen(MainForm.ScreenPatients)
            Case "colours" : AppHost.MainWindow?.ShowScreen(MainForm.ScreenColours)
        End Select
    End Sub
End Class
