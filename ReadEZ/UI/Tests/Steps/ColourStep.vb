Option Strict On

' ====================================================================================================
'   ColourStep.vb - determination of the preferred colour (ReadEZ 2020: Colour test, Auto / Manual /
'   Bespoke tabs)
'
'   Overlays   two copies of the text side by side, each on an overlay colour; the patient says
'              which side is more comfortable ("left" / "right", or Left / Right arrow keys).
'   Lenses /   the whole screen in one colour, then the other (each for the presentation time),
'   computer   then white again; the patient says "first" or "second" (Left / Right arrows);
'              Repeat (R) shows the pair again. Lenses: test in a darkened room.
'
'   Automatic  the staircase algorithm (ColourStaircase) chooses the pairs (about 35).
'   Manual     the examiner picks the colours on the wheel(s) and selects the preferred one.
'   Bespoke    any screen colour (colour dialog), saved as a bespoke colour.
' ====================================================================================================

Public NotInheritable Class ColourStep
    Inherits TestStep

    Private _auto As ColourStaircase
    Private ReadOnly segMethod As New SegmentedControl
    Private ReadOnly _autoPanel As FlowLayoutPanel = Ui.Column()
    Private ReadOnly _manualPanel As FlowLayoutPanel = Ui.Column()
    Private ReadOnly lblProgress As New Label With {.AutoSize = True, .Font = Fonts.Strong, .MaximumSize = New Size(310, 0)}
    Private ReadOnly lblResult As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(310, 0)}
    Private ReadOnly btnStart As ModernButton = ModernButton.Create("colour.start", ButtonKind.Primary, Glyphs.Play)
    Private ReadOnly btnLeft As ModernButton
    Private ReadOnly btnRight As ModernButton
    Private ReadOnly btnRepeat As ModernButton = ModernButton.Create("colour.repeat", ButtonKind.Ghost, Glyphs.Refresh)
    Private ReadOnly wheelAuto As New ColourWheel With {.Selectable = False, .Size = New Size(260, 260)}
    Private ReadOnly wheelLeft As New ColourWheel With {.Size = New Size(260, 260)}
    Private ReadOnly wheelRight As New ColourWheel With {.Size = New Size(260, 260)}
    Private ReadOnly sldTiming As New Slider With {.Minimum = 5, .Maximum = 100, .Width = 280, .Height = 30, .ShowValue = True}
    Private ReadOnly btnNext As ModernButton
    Private ReadOnly _timer As New Timer
    Private _successiveStage As Integer        ' 0 none, 1 first colour, 2 second colour
    Private _manualLeft As (SetOf As ReadEZSet, Index As Integer) = (ReadEZSet.Overlay, 1)
    Private _manualRight As (SetOf As ReadEZSet, Index As Integer) = (ReadEZSet.Overlay, 6)
    Private _respondEnabled As Boolean

    Public Sub New()
        btnLeft = ModernButton.Create("colour.left", ButtonKind.Secondary, Glyphs.Back)
        btnRight = ModernButton.Create("colour.right", ButtonKind.Secondary, Glyphs.Forward)
        btnNext = ModernButton.Create("run.nextStep", ButtonKind.Primary, Glyphs.Forward)
        AddHandler _timer.Tick, Sub() SuccessiveTick()
    End Sub

    Public Overrides ReadOnly Property TitleKey As String
        Get
            Return If(Run IsNot Nothing AndAlso Run.Method = 1, "step.colourManual", "step.colourAuto")
        End Get
    End Property

    Public Overrides ReadOnly Property HelpTopic As String
        Get
            Return "test.colour"
        End Get
    End Property

    Private ReadOnly Property Mode As ReadEZTestMode
        Get
            Return Run.Result.Mode
        End Get
    End Property

    Private ReadOnly Property IsOverlay As Boolean
        Get
            Return Mode = ReadEZTestMode.Overlay
        End Get
    End Property

#Region "Panel"

    Public Overrides Function CreatePanel() As Control
        Dim col = Ui.Column()
        Dim intro = Ui.Label(If(IsOverlay, "colour.introOverlay", If(Mode = ReadEZTestMode.Lenses, "colour.introLenses", "colour.introComputer")), muted:=True)
        intro.MaximumSize = New Size(310, 0)
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)
        segMethod.SetOptions("colour.auto", "colour.manual")
        segMethod.SelectedIndex = Run.Method
        segMethod.Margin = New Padding(0, 0, 0, 10)
        AddHandler segMethod.SelectedIndexChanged, Sub() MethodChanged()
        col.Controls.Add(segMethod)

        ' ---- Automatic ----
        lblProgress.Margin = New Padding(0, 0, 0, 8)
        _autoPanel.Controls.Add(lblProgress)
        btnLeft.Text = Loc.T(If(IsOverlay, "colour.left", "colour.first"))
        btnRight.Text = Loc.T(If(IsOverlay, "colour.right", "colour.second"))
        Dim answer = Ui.Row()
        btnLeft.Margin = New Padding(0, 0, 8, 8)
        btnRight.Margin = New Padding(0, 0, 8, 8)
        answer.Controls.Add(btnLeft)
        answer.Controls.Add(btnRight)
        _autoPanel.Controls.Add(answer)
        Dim row2 = Ui.Row()
        btnStart.Margin = New Padding(0, 0, 8, 8)
        row2.Controls.Add(btnStart)
        If Not IsOverlay Then row2.Controls.Add(btnRepeat)
        _autoPanel.Controls.Add(row2)
        _autoPanel.Controls.Add(wheelAuto)
        AddHandler btnStart.Click, Sub() StartAuto()
        AddHandler btnLeft.Click, Sub() Respond(True)
        AddHandler btnRight.Click, Sub() Respond(False)
        AddHandler btnRepeat.Click, Sub() Present()
        wheelAuto.OuterSet = If(IsOverlay, ReadEZSet.Overlay, ReadEZSet.Lens)
        wheelAuto.TwoRings = Not IsOverlay
        col.Controls.Add(_autoPanel)

        ' ---- Manual ----
        wheelLeft.OuterSet = If(IsOverlay, ReadEZSet.Overlay, ReadEZSet.Lens)
        wheelLeft.TwoRings = Not IsOverlay
        wheelRight.OuterSet = ReadEZSet.Overlay
        If IsOverlay Then
            _manualPanel.Controls.Add(Ui.Label("colour.leftColour", muted:=True))
            _manualPanel.Controls.Add(wheelLeft)
            Dim selL = ModernButton.Create("colour.selectLeft", ButtonKind.Secondary, Glyphs.Accept)
            AddHandler selL.Click, Sub() SelectColour(_manualLeft, 1)
            _manualPanel.Controls.Add(selL)
            _manualPanel.Controls.Add(Ui.Label("colour.rightColour", muted:=True))
            _manualPanel.Controls.Add(wheelRight)
            Dim selR = ModernButton.Create("colour.selectRight", ButtonKind.Secondary, Glyphs.Accept)
            AddHandler selR.Click, Sub() SelectColour(_manualRight, 1)
            _manualPanel.Controls.Add(selR)
        Else
            _manualLeft = (ReadEZSet.Lens, 1)
            Dim hint = Ui.Label("colour.ringsHint", muted:=True)
            hint.MaximumSize = New Size(310, 0)
            _manualPanel.Controls.Add(hint)
            _manualPanel.Controls.Add(wheelLeft)
            Dim sel = ModernButton.Create("colour.select", ButtonKind.Secondary, Glyphs.Accept)
            AddHandler sel.Click, Sub() SelectColour(_manualLeft, 1)
            _manualPanel.Controls.Add(sel)
        End If
        Dim bespoke = ModernButton.Create("colour.bespoke", ButtonKind.Ghost, Glyphs.Color)
        AddHandler bespoke.Click, Sub() PickBespoke()
        _manualPanel.Controls.Add(bespoke)
        AddHandler wheelLeft.ColourClicked, Sub(s, i, b)
                                                _manualLeft = (s, i)
                                                ShowManual()
                                            End Sub
        AddHandler wheelRight.ColourClicked, Sub(s, i, b)
                                                 _manualRight = (s, i)
                                                 ShowManual()
                                             End Sub
        col.Controls.Add(_manualPanel)

        ' ---- Presentation time (lenses / computer) ----
        If Not IsOverlay Then
            sldTiming.Value = Math.Max(5, Math.Min(100, AppHost.Settings.Timing \ 100))
            sldTiming.ValueFormat = Function(v) (v / 10.0).ToString("0.0") & " s"
            AddHandler sldTiming.ValueChanged, Sub()
                                                   AppHost.Settings.Timing = sldTiming.Value * 100
                                                   Run.Result.RoR_PresentationTime = AppHost.Settings.Timing
                                                   AppHost.SaveSettings()
                                               End Sub
            col.Controls.Add(Ui.Stack("colour.timing", sldTiming))
        End If

        lblResult.Margin = New Padding(0, 10, 0, 0)
        col.Controls.Add(lblResult)
        col.Controls.Add(TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=IsOverlay, enabledWhen:=Function() _successiveStage = 0))
        btnNext.Text = Loc.T(If(Run.Index = Run.Steps.Count - 1, "run.saveFinish", "run.nextStep"))
        btnNext.Margin = New Padding(0, 14, 0, 0)
        AddHandler btnNext.Click, Sub() Finish()
        col.Controls.Add(btnNext)
        Return col
    End Function

    Private Sub MethodChanged()
        Run.Method = segMethod.SelectedIndex
        AppHost.Settings.ColourTestMethod = Run.Method
        AppHost.SaveSettings()
        StopSuccessive()
        _auto = Nothing
        Host?.RefreshTitleFor(Me)
        UpdatePanels()
        If Run.Method = 1 Then ShowManual() Else ShowWhite()
    End Sub

    Private Sub UpdatePanels()
        _autoPanel.Visible = Run.Method = 0
        _manualPanel.Visible = Run.Method = 1
        Dim running = _auto IsNot Nothing AndAlso _auto.CurrentPhase <> ColourStaircase.Phase.Done
        btnLeft.Visible = running
        btnRight.Visible = running
        btnRepeat.Visible = running AndAlso Not IsOverlay
        btnLeft.Enabled = _respondEnabled
        btnRight.Enabled = _respondEnabled
        btnStart.Text = Loc.T(If(running, "colour.restart", If(_auto Is Nothing, "colour.start", "colour.retest")))
        If _auto Is Nothing Then
            lblProgress.Text = Loc.T("colour.pressStart")
        ElseIf running Then
            lblProgress.Text = Loc.F("colour.progress", CInt(_auto.CurrentPhase), _auto.Count, _auto.Expected)
        Else
            lblProgress.Text = Loc.T("colour.autoDone")
        End If
        lblResult.Text = If(Run.HasColour, Loc.F("colour.chosen", Run.Result.ColourName) &
                            If(Run.Result.GreyPreferred, vbCrLf & Loc.T("colour.greyPreferred"), ""), "")
        lblResult.ForeColor = If(Run.Result.GreyPreferred, ThemeManager.Current.Warning, ThemeManager.Current.Ink)
        btnNext.Enabled = Run.HasColour
        ActionsChanged()
    End Sub

#End Region

#Region "Showing colours"

    Public Overrides Sub Begin()
        ApplyTextSettings()
        Scene.ShowMarks = False
        Run.Result.RoR_PresentationTime = AppHost.Settings.Timing
        MyBase.Begin()
        UpdatePanels()
        If Run.Method = 1 Then ShowManual() Else ShowWhite()
    End Sub

    Public Overrides Sub [End]()
        StopSuccessive()
    End Sub

    Private Sub ShowWhite()
        Scene.Caption = ""
        Scene.Presentation = ReadingPresentation.White
        Scene.Colour1 = Color.White
        Redraw()
    End Sub

    Private Sub ShowManual()
        wheelLeft.Selected = _manualLeft
        wheelRight.Selected = _manualRight
        wheelLeft.Invalidate() : wheelRight.Invalidate()
        Scene.Caption = ""
        If IsOverlay Then
            Scene.Presentation = ReadingPresentation.TwoOverlays
            Scene.Colour1 = ReadEZColours.Colour(_manualLeft.SetOf, _manualLeft.Index)
            Scene.Colour2 = ReadEZColours.Colour(_manualRight.SetOf, _manualRight.Index)
        Else
            Scene.Presentation = ReadingPresentation.FullScreen
            Scene.Colour1 = ReadEZColours.Colour(_manualLeft.SetOf, _manualLeft.Index)
        End If
        Redraw()
    End Sub

    Private Sub SelectColour(c As (SetOf As ReadEZSet, Index As Integer), method As Integer)
        Dim r = Run.Result
        r.Colour1 = c.Index
        r.PaletteSet = If(c.SetOf = ReadEZSet.Overlay, 1, If(c.SetOf = ReadEZSet.Lens, 2, 3))
        Dim sc = ReadEZColours.Colour(c.SetOf, c.Index)
        r.Colour_R = sc.R : r.Colour_G = sc.G : r.Colour_B = sc.B
        r.Screening_Mode = method
        If method <> 0 Then r.GreyPreferred = False : r.AutoScore = ""
        Run.HasColour = True
        UpdatePanels()
        Toast.Show(AppHost.MainWindow, Loc.F("colour.chosen", r.ColourName), ToastKind.Success)
    End Sub

    Private Sub PickBespoke()
        Using dlg As New ColorDialog With {.FullOpen = True, .Color = Scene.Colour1}
            If dlg.ShowDialog(Host?.FindForm()) <> DialogResult.OK Then Return
            Dim r = Run.Result
            r.Colour1 = 0
            r.PaletteSet = If(IsOverlay, 1, 2)
            r.Colour_R = dlg.Color.R : r.Colour_G = dlg.Color.G : r.Colour_B = dlg.Color.B
            r.Screening_Mode = 2
            r.GreyPreferred = False
            Run.HasColour = True
            Scene.Presentation = If(IsOverlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen)
            Scene.Colour1 = dlg.Color
            Redraw()
            UpdatePanels()
        End Using
    End Sub

#End Region

#Region "Automatic"

    Private Sub StartAuto()
        If _auto IsNot Nothing AndAlso _auto.CurrentPhase <> ColourStaircase.Phase.Done AndAlso Not Ask.YesNo(Host?.FindForm(), "colour.confirmRestart") Then Return
        If Mode = ReadEZTestMode.Lenses AndAlso _auto Is Nothing Then Ask.Info(Host?.FindForm(), "colour.darkRoom")
        _auto = New ColourStaircase(Mode)
        wheelAuto.Spokes = _auto.Chosen
        wheelAuto.Selected = Nothing
        Present()
    End Sub

    ''' <summary>Shows the current pair: side by side (overlays) or one after the other (lenses, computer).</summary>
    Private Sub Present()
        If _auto Is Nothing Then Return
        StopSuccessive()
        If IsOverlay Then
            Scene.Caption = ""
            Scene.Presentation = ReadingPresentation.TwoOverlays
            Scene.Colour1 = _auto.LeftColour
            Scene.Colour2 = _auto.RightColour
            _respondEnabled = True
            Redraw()
        Else
            _respondEnabled = False
            _successiveStage = 1
            Scene.Caption = "1"
            Scene.Presentation = ReadingPresentation.FullScreen
            Scene.Colour1 = _auto.LeftColour
            Redraw()
            _timer.Interval = Math.Max(500, AppHost.Settings.Timing)
            _timer.Start()
        End If
        UpdatePanels()
    End Sub

    Private Sub SuccessiveTick()
        If _auto Is Nothing Then _timer.Stop() : Return
        If _successiveStage = 1 Then
            _successiveStage = 2
            Scene.Caption = "2"
            Scene.Colour1 = _auto.RightColour
            Redraw()
        Else
            StopSuccessive()
            ShowWhite()
            _respondEnabled = True
            UpdatePanels()
        End If
    End Sub

    Private Sub StopSuccessive()
        _timer.Stop()
        _successiveStage = 0
    End Sub

    Private Sub Respond(preferLeft As Boolean)
        If _auto Is Nothing OrElse Not _respondEnabled OrElse _auto.CurrentPhase = ColourStaircase.Phase.Done Then Return
        _auto.Respond(preferLeft)
        wheelAuto.Spokes = _auto.Chosen
        wheelAuto.Invalidate()
        If _auto.CurrentPhase = ColourStaircase.Phase.Done Then
            _respondEnabled = False
            Dim r = Run.Result
            r.Colour1 = _auto.Winner
            r.PaletteSet = _auto.PaletteSet
            r.GreyPreferred = _auto.GreyPreferred
            r.AutoScore = _auto.ScoreText()
            Dim sc = ReadEZColours.Colour(r.ColourSet, r.Colour1)
            r.Colour_R = sc.R : r.Colour_G = sc.G : r.Colour_B = sc.B
            r.Screening_Mode = 0
            Run.HasColour = True
            wheelAuto.Selected = (r.ColourSet, r.Colour1)
            ' Show the text in the chosen colour
            Scene.Caption = ""
            Scene.Colour1 = sc
            Scene.Presentation = If(IsOverlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen)
            Redraw()
            UpdatePanels()
            Toast.Show(AppHost.MainWindow, Loc.F("colour.chosen", r.ColourName), ToastKind.Success)
        Else
            Present()
        End If
    End Sub

#End Region

    Public Overrides Function IsComplete() As Boolean
        Return Run.HasColour
    End Function

    Public Overrides Function HandleKey(keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Left
                If Run.Method = 0 Then Respond(True) : Return True
            Case Keys.Right
                If Run.Method = 0 Then Respond(False) : Return True
            Case Keys.R
                If Run.Method = 0 AndAlso Not IsOverlay Then Present() : Return True
            Case Keys.Enter
                If Run.HasColour Then Finish() : Return True
        End Select
        Return False
    End Function

    Public Overrides Function QuickActions() As List(Of QuickAction)
        Dim list As New List(Of QuickAction)
        If Run.Method = 0 Then
            Dim running = _auto IsNot Nothing AndAlso _auto.CurrentPhase <> ColourStaircase.Phase.Done
            If running Then
                list.Add(New QuickAction(If(IsOverlay, "colour.left", "colour.first"), Sub() Respond(True), primary:=True, enabled:=_respondEnabled))
                list.Add(New QuickAction(If(IsOverlay, "colour.right", "colour.second"), Sub() Respond(False), primary:=True, enabled:=_respondEnabled))
                If Not IsOverlay Then list.Add(New QuickAction("colour.repeat", Sub() Present()))
            Else
                list.Add(New QuickAction(If(_auto Is Nothing, "colour.start", "colour.retest"), Sub() StartAuto(), primary:=_auto Is Nothing))
            End If
        End If
        If Run.HasColour Then list.Add(New QuickAction(If(Run.Index = Run.Steps.Count - 1, "run.saveFinish", "run.nextStep"), Sub() Finish(), primary:=True))
        Return list
    End Function

    Public Overrides Sub Dispose()
        _timer.Dispose()
        MyBase.Dispose()
    End Sub
End Class
