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
'   Automatic  Quick (QuickColourSearch, about 12 pairs) or Classic (ColourStaircase, ReadEZ 2020's
'              staircases, about 35) chooses the pairs. The answer can be given with the buttons /
'              arrow keys, or by clicking the preferred side of the text (overlays) or the screen
'              while the preferred colour is showing (lenses, computer) - on the patient's screen or
'              its copy here. Quick test: the patient can also point at a colour on the wheel; the
'              test then checks it against its neighbours and grey.
'   Manual     the examiner picks the colours on the wheel(s) and selects the preferred one.
'   Bespoke    any screen colour (colour dialog), saved as a bespoke colour.
' ====================================================================================================

Public NotInheritable Class ColourStep
    Inherits TestStep

    Private _auto As IColourSearch
    Private ReadOnly segAlgo As New SegmentedControl
    ''' <summary>Lenses / computer: the patient clicked while the first colour was showing.</summary>
    Private _pendingFirst As Boolean
    Private ReadOnly segMethod As New SegmentedControl
    Private ReadOnly _autoPanel As FlowLayoutPanel = Ui.Column()
    Private ReadOnly _manualPanel As FlowLayoutPanel = Ui.Column()
    Private ReadOnly lblProgress As New Label With {.AutoSize = True, .Font = Fonts.Strong}
    Private ReadOnly lblResult As New Label With {.AutoSize = True, .Font = Fonts.Body}
    Private ReadOnly btnStart As ModernButton = ModernButton.Create("colour.start", ButtonKind.Primary, Glyphs.Play)
    Private ReadOnly btnLeft As ModernButton
    Private ReadOnly btnRight As ModernButton
    Private ReadOnly btnRepeat As ModernButton = ModernButton.Create("colour.repeat", ButtonKind.Ghost, Glyphs.Refresh)
    Private ReadOnly wheelAuto As New ColourWheel With {.Selectable = False, .Size = New Size(230, 230)}
    ' Manual: one wheel; overlays mark the left colour (solid ring) and the right colour (dashed ring)
    Private ReadOnly wheelManual As New ColourWheel With {.Size = New Size(230, 230), .Margin = New Padding(0, 2, 0, 4)}
    Private ReadOnly segSide As New SegmentedControl
    Private ReadOnly btnSaveOnly As ModernButton = ModernButton.Create("colour.saveOnly", ButtonKind.Secondary, Glyphs.Save)
    ''' <summary>What to ask the patient now (automatic test).</summary>
    Private ReadOnly lblPrompt As New Label With {.AutoSize = True, .Font = Fonts.Strong}
    Private _timingStack As Control
    ''' <summary>The patient answers on their own screen.</summary>
    Private ReadOnly btnPatient As ModernButton = ModernButton.CreateToggle("colour.patientChooses", Glyphs.Eye)
    ' Feedback for the patient's answer: the choice is shown for a moment, then acted on
    Private ReadOnly _ackTimer As New Timer With {.Interval = 450}
    Private _ackAction As Action
    Private ReadOnly sldTiming As New Slider With {.Minimum = 5, .Maximum = 100, .Height = 30, .ShowValue = True}
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
        AddHandler btnNext.Click, Sub() Finish()
        AddHandler btnSaveOnly.Click, Sub() SaveColourOnly()
        AddHandler _timer.Tick, Sub() SuccessiveTick()
        AddHandler _ackTimer.Tick, Sub()
                                       _ackTimer.Stop()
                                       Scene.Highlight = -1
                                       Dim a = _ackAction
                                       _ackAction = Nothing
                                       a?.Invoke()
                                   End Sub
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
        Ui.FillWidth(col)                                 ' labels wrap at, and sliders / cards stretch to, the panel's width
        Dim intro = Ui.Label(If(IsOverlay, "colour.introOverlay", If(Mode = ReadEZTestMode.Lenses, "colour.introLenses", "colour.introComputer")), muted:=True)
        intro.Font = Fonts.Body
        intro.Margin = New Padding(0, 0, 0, 6)
        col.Controls.Add(intro)
        If Mode = ReadEZTestMode.Lenses Then
            ' Lenses: the darkened room, clearly
            Dim dark = Ui.Label("colour.darkRoomBanner")
            dark.Font = Fonts.Section
            LabelStyle.SetCustom(dark)
            dark.ForeColor = ThemeManager.Current.Warning
            dark.Margin = New Padding(0, 0, 0, 8)
            col.Controls.Add(dark)
        End If
        ' How to answer, point at a colour, use the manual wheel: in a window
        Dim helpKeys As New List(Of String) From {If(IsOverlay, "colour.clickHintOverlay", "colour.clickHintSuccessive"), "colour.pointHint",
                                                  If(IsOverlay, "colour.manualHint", "colour.ringsHint")}
        col.Controls.Add(HelpButton("colour.help.button", "colour.help.title", helpKeys.ToArray()))
        segMethod.SetOptions("colour.auto", "colour.manual")
        segMethod.SelectedIndex = Run.Method
        segMethod.Margin = New Padding(0, 0, 0, 10)
        AddHandler segMethod.SelectedIndexChanged, Sub() MethodChanged()
        col.Controls.Add(segMethod)
        ' The patient can answer on their own screen
        btnPatient.Checked = AppHost.Settings.PatientChooses
        btnPatient.Margin = New Padding(0, 0, 0, 10)
        AddHandler btnPatient.CheckedChanged, Sub()
                                                  AppHost.Settings.PatientChooses = btnPatient.Checked
                                                  AppHost.SaveSettings()
                                                  ApplyPatientChoice()
                                              End Sub
        col.Controls.Add(btnPatient)

        ' ---- Automatic ----
        segAlgo.SetOptions("colour.algo.quick", "colour.algo.classic")
        segAlgo.SelectedIndex = If(AppHost.Settings.ColourAutoQuick, 0, 1)
        segAlgo.Margin = New Padding(0, 0, 0, 8)
        AddHandler segAlgo.SelectedIndexChanged, Sub()
                                                     AppHost.Settings.ColourAutoQuick = segAlgo.SelectedIndex = 0
                                                     AppHost.SaveSettings()
                                                     UpdatePanels()
                                                 End Sub
        _autoPanel.Controls.Add(segAlgo)
        lblProgress.Margin = New Padding(0, 0, 0, 4)
        _autoPanel.Controls.Add(lblProgress)
        LabelStyle.SetCustom(lblPrompt)
        lblPrompt.ForeColor = ThemeManager.Current.Primary
        lblPrompt.Margin = New Padding(0, 0, 0, 8)
        _autoPanel.Controls.Add(lblPrompt)
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
        AddHandler wheelAuto.ColourClicked, Sub(st, i, b) PointAt(i)
        AddHandler btnStart.Click, Sub() StartAuto()
        AddHandler btnLeft.Click, Sub() Respond(True)
        AddHandler btnRight.Click, Sub() Respond(False)
        AddHandler btnRepeat.Click, Sub() Present()
        wheelAuto.OuterSet = If(IsOverlay, ReadEZSet.Overlay, ReadEZSet.Lens)
        wheelAuto.TwoRings = Not IsOverlay
        col.Controls.Add(_autoPanel)

        ' ---- Manual: one wheel ----
        wheelManual.OuterSet = If(IsOverlay, ReadEZSet.Overlay, ReadEZSet.Lens)
        wheelManual.TwoRings = Not IsOverlay
        If IsOverlay Then
            ' Which side a click on the wheel sets (a right-click always sets the right side)
            segSide.SetOptions("colour.leftColour", "colour.rightColour")
            segSide.Margin = New Padding(0, 0, 0, 4)
            _manualPanel.Controls.Add(segSide)
            _manualPanel.Controls.Add(wheelManual)
            Dim choose = Ui.Row(wrap:=True)
            Dim selL = ModernButton.Create("colour.chooseLeft", ButtonKind.Secondary, Glyphs.Accept)
            Dim selR = ModernButton.Create("colour.chooseRight", ButtonKind.Secondary, Glyphs.Accept)
            selL.Margin = New Padding(0, 0, 8, 6)
            selR.Margin = New Padding(0, 0, 8, 6)
            AddHandler selL.Click, Sub() SelectColour(_manualLeft, 1)
            AddHandler selR.Click, Sub() SelectColour(_manualRight, 1)
            choose.Controls.Add(selL)
            choose.Controls.Add(selR)
            _manualPanel.Controls.Add(choose)
        Else
            _manualLeft = (ReadEZSet.Lens, 1)
            _manualPanel.Controls.Add(wheelManual)
            Dim sel = ModernButton.Create("colour.select", ButtonKind.Secondary, Glyphs.Accept)
            sel.Margin = New Padding(0, 0, 0, 6)
            AddHandler sel.Click, Sub() SelectColour(_manualLeft, 1)
            _manualPanel.Controls.Add(sel)
        End If
        ' White (no colour) and any colour (the patient's screen follows the sliders)
        Dim extra = Ui.Row(wrap:=True)
        Dim whiteButton = ModernButton.Create("colour.white", ButtonKind.Ghost, Glyphs.Brightness)
        whiteButton.Margin = New Padding(0, 0, 8, 4)
        AddHandler whiteButton.Click, Sub() SelectWhite()
        extra.Controls.Add(whiteButton)
        Dim bespoke = ModernButton.Create("colour.bespoke", ButtonKind.Ghost, Glyphs.Color)
        AddHandler bespoke.Click, Sub() PickBespoke()
        extra.Controls.Add(bespoke)
        _manualPanel.Controls.Add(extra)
        AddHandler wheelManual.ColourClicked, Sub(s, i, b)
                                                  If IsOverlay AndAlso (b = MouseButtons.Right OrElse segSide.SelectedIndex = 1) Then
                                                      _manualRight = (s, i)
                                                  Else
                                                      _manualLeft = (s, i)
                                                  End If
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
            _timingStack = Ui.Stack("colour.timing", sldTiming)
            col.Controls.Add(_timingStack)                ' automatic test only (UpdatePanels)
        End If

        lblResult.Margin = New Padding(0, 10, 0, 0)
        col.Controls.Add(lblResult)
        If Run.Result.IsPattern Then
            col.Controls.Add(GratingOptions.Create(Me, Sub() Redraw(), enabledWhen:=Function() _successiveStage = 0))
        Else
            col.Controls.Add(TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=IsOverlay, enabledWhen:=Function() _successiveStage = 0))
        End If
        Return col
    End Function

    ''' <summary>Next step (or Save and finish) and, when reading and symptoms follow, "Save the colour only".</summary>
    Public Overrides Function CreateFooter() As Control
        Dim last = Run.Index = Run.Steps.Count - 1
        btnNext.Text = Loc.T(If(last, "run.saveFinish", "run.nextStep"))
        btnNext.IconGlyph = If(last, Glyphs.Save, Glyphs.Forward)
        btnSaveOnly.Visible = Not last
        Return FooterRow(btnNext, btnSaveOnly)
    End Function

    ''' <summary>Saves the colour now, without the rate of reading and symptoms with it.</summary>
    Private Sub SaveColourOnly()
        If Not Run.HasColour Then Return
        SaveNow()
    End Sub

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
        Dim running = _auto IsNot Nothing AndAlso Not _auto.IsDone
        Dim quick = AppHost.Settings.ColourAutoQuick
        segAlgo.Enabled = Not running
        ' Pointing at a colour on the wheel: the quick test, before it starts or while it searches
        wheelAuto.Selectable = quick AndAlso (_auto Is Nothing OrElse _auto.IsDone OrElse _auto.CanJumpTo)
        btnLeft.Visible = running
        btnRight.Visible = running
        btnRepeat.Visible = running AndAlso Not IsOverlay
        btnLeft.Enabled = _respondEnabled
        btnRight.Enabled = _respondEnabled
        btnStart.Text = Loc.T(If(running, "colour.restart", If(_auto Is Nothing, "colour.start", "colour.retest")))
        If _auto Is Nothing Then
            lblProgress.Text = Loc.T("colour.pressStart")
        ElseIf running Then
            lblProgress.Text = Loc.F("colour.progress", _auto.StageNumber, _auto.Count, _auto.Expected)
        Else
            lblProgress.Text = Loc.T("colour.autoDone")
        End If
        lblResult.Text = If(Run.HasColour, Loc.F("colour.chosen", Run.Result.ColourName) &
                            If(Run.Result.GreyPreferred, vbCrLf & Loc.T("colour.greyPreferred"), ""), "")
        lblResult.ForeColor = If(Run.Result.GreyPreferred, ThemeManager.Current.Warning, ThemeManager.Current.Ink)
        btnNext.Enabled = Run.HasColour
        btnSaveOnly.Enabled = Run.HasColour
        If _timingStack IsNot Nothing Then _timingStack.Visible = Run.Method = 0
        btnPatient.Visible = Run.Method = 0 OrElse IsOverlay
        UpdatePrompt()
        ActionsChanged()
    End Sub

    ''' <summary>The question to put to the patient now, and how to give the answer.</summary>
    Private Sub UpdatePrompt()
        Dim running = Run.Method = 0 AndAlso _auto IsNot Nothing AndAlso Not _auto.IsDone
        lblPrompt.Visible = running
        If Not running Then Return
        Dim key As String
        If IsOverlay Then
            key = "colour.ask.overlay"
        ElseIf _successiveStage = 1 Then
            key = "colour.ask.showing1"
        ElseIf _successiveStage = 2 Then
            key = "colour.ask.showing2"
        Else
            key = "colour.ask.successive"
        End If
        lblPrompt.Text = Loc.T(key)
    End Sub

    ''' <summary>
    ''' Overlays: two copies of the text on two colours, "left" and "right" pointed out. Lenses /
    ''' computer: the two colours side by side, numbered, "first ... then second".
    ''' </summary>
    Public Overrides Function InstructionScreen() As InstructionScene
        Dim sc As New InstructionScene(InstructionKeys())
        If IsOverlay Then
            Dim sample = SampleText()
            sample.Presentation = ReadingPresentation.TwoOverlays
            sample.Colour1 = ReadEZColours.Colour(ReadEZSet.Overlay, 1)
            sample.Colour2 = ReadEZColours.Colour(ReadEZSet.Overlay, 7)
            sc.Sample = sample
            sc.Add("instr.overlay.calloutLeft", sample.CopyBounds(0), InstructionScene.Side.Below)
            sc.Add("instr.overlay.calloutRight", sample.CopyBounds(1), InstructionScene.Side.Below)
            sc.Add(If(PatientChooses, "instr.overlay.calloutTap", "instr.overlay.calloutSay"), RectangleF.Empty, InstructionScene.Side.BottomCentre)
        Else
            sc.Sample = New TwoColourSample With {.Colour1 = ReadEZColours.Colour(ReadEZSet.Lens, 1), .Colour2 = ReadEZColours.Colour(ReadEZSet.Lens, 7)}
            sc.Add("instr.successive.callout1", TwoColourSample.Half(1), InstructionScene.Side.Inside)
            sc.Add("instr.successive.callout2", TwoColourSample.Half(2), InstructionScene.Side.Inside)
            sc.Add(If(PatientChooses, "instr.successive.calloutTap", "instr.successive.calloutSay"), RectangleF.Empty, InstructionScene.Side.BottomCentre)
        End If
        Return sc
    End Function

    Public Overrides Function InstructionKeys() As String()
        Select Case Mode
            Case ReadEZTestMode.Overlay
                Return {"instr.colour.title", If(Run.Result.IsPattern, "instr.overlay.1p", "instr.overlay.1"), "instr.overlay.2", "instr.colour.3"}
            Case ReadEZTestMode.Lenses
                Return {"instr.colour.title", "instr.lenses.0", "instr.successive.1", "instr.successive.2", "instr.colour.3"}
            Case Else
                Return {"instr.colour.title", "instr.successive.1", "instr.successive.2", "instr.colour.3"}
        End Select
    End Function

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
        Scene.ChoicePrompt = ""
        Scene.ChoiceButtons = Nothing
        Scene.Presentation = ReadingPresentation.White
        Scene.Colour1 = Color.White
        Redraw()
    End Sub

    Private ReadOnly Property PatientChooses As Boolean
        Get
            Return AppHost.Settings.PatientChooses
        End Get
    End Property

    ''' <summary>
    ''' Lenses / computer, patient chooses: after the two colours, a white screen asking "Which colour
    ''' was more comfortable?" with large buttons: The first, The second, Show again.
    ''' </summary>
    Private Sub ShowChoiceScreen()
        Scene.Caption = ""
        Scene.Presentation = ReadingPresentation.ColourOnly
        Scene.Colour1 = Color.White
        Scene.ChoicePrompt = Loc.T("colour.patient.which")
        Scene.ChoiceButtons = {Loc.T("colour.patient.first"), Loc.T("colour.patient.second"), Loc.T("colour.patient.again")}
        Redraw()
    End Sub

    ''' <summary>Overlays: the hint on the patient's screen to click the side they prefer.</summary>
    Private Function SideHint() As String
        Return If(PatientChooses AndAlso IsOverlay, Loc.T("colour.patient.clickSide"), "")
    End Function

    ''' <summary>"Patient chooses" switched on or off: the patient's screen changes at once.</summary>
    Private Sub ApplyPatientChoice()
        Dim running = Run.Method = 0 AndAlso _auto IsNot Nothing AndAlso Not _auto.IsDone
        If running AndAlso Not IsOverlay Then
            If _successiveStage = 0 AndAlso _respondEnabled Then
                If PatientChooses Then ShowChoiceScreen() Else ShowWhite()
            End If
            Return
        End If
        If running OrElse (Run.Method = 1 AndAlso IsOverlay) Then
            Scene.ChoicePrompt = SideHint()
            Redraw()
        End If
    End Sub

    Private Sub ShowManual()
        wheelManual.Selected = _manualLeft
        wheelManual.Second = If(IsOverlay, CType(_manualRight, (SetOf As ReadEZSet, Index As Integer)?), Nothing)
        wheelManual.Invalidate()
        Scene.Caption = ""
        Scene.ChoiceButtons = Nothing
        Scene.ChoicePrompt = SideHint()
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

    ''' <summary>Manual: white (no colour) as the preferred "colour", shown on the patient's screen.</summary>
    Private Sub SelectWhite()
        Dim r = Run.Result
        r.Colour1 = 0
        r.PaletteSet = If(IsOverlay, 1, 2)
        r.Colour_R = 255 : r.Colour_G = 255 : r.Colour_B = 255
        r.Screening_Mode = 1
        r.GreyPreferred = False
        r.AutoScore = ""
        Run.HasColour = True
        Scene.Caption = ""
        Scene.ChoiceButtons = Nothing
        Scene.Presentation = ReadingPresentation.White
        Scene.Colour1 = Color.White
        Redraw()
        UpdatePanels()
        Toast.Show(AppHost.MainWindow, Loc.F("colour.chosen", r.ColourName), ToastKind.Success)
    End Sub

    ''' <summary>Any colour: the patient's screen changes as the sliders move; Cancel puts back what was shown.</summary>
    Private Sub PickBespoke()
        Dim wasPresentation = Scene.Presentation, wasColour = Scene.Colour1
        Dim preview = Sub(col As Color)
                          Scene.Caption = ""
                          Scene.ChoiceButtons = Nothing
                          Scene.Presentation = If(IsOverlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen)
                          Scene.Colour1 = col
                          Redraw()
                      End Sub
        Using dlg As New BespokeColourDialog(Scene.Colour1, preview)
            If dlg.ShowDialog(Ask.DialogOwner(Host?.FindForm())) <> DialogResult.OK Then
                Scene.Presentation = wasPresentation
                Scene.Colour1 = wasColour
                Redraw()
                Return
            End If
            Dim r = Run.Result
            r.Colour1 = 0
            r.PaletteSet = If(IsOverlay, 1, 2)
            r.Colour_R = dlg.Colour.R : r.Colour_G = dlg.Colour.G : r.Colour_B = dlg.Colour.B
            r.Screening_Mode = 2
            r.GreyPreferred = False
            Run.HasColour = True
            preview(dlg.Colour)
            UpdatePanels()
        End Using
    End Sub

#End Region

#Region "Automatic"

    Private Sub StartAuto()
        If _auto IsNot Nothing AndAlso Not _auto.IsDone AndAlso Not Ask.YesNo(Host?.FindForm(), "colour.confirmRestart") Then Return
        If Mode = ReadEZTestMode.Lenses AndAlso _auto Is Nothing Then Ask.Info(Host?.FindForm(), "colour.darkRoom")
        _auto = If(AppHost.Settings.ColourAutoQuick, CType(New QuickColourSearch(Mode), IColourSearch), New ColourStaircase(Mode))
        wheelAuto.Spokes = _auto.Chosen
        wheelAuto.Selected = Nothing
        Present()
    End Sub

    ''' <summary>Shows the current pair: side by side (overlays) or one after the other (lenses, computer).</summary>
    Private Sub Present()
        If _auto Is Nothing Then Return
        StopSuccessive()
        _pendingFirst = False
        Scene.ChoiceButtons = Nothing
        Scene.ChoicePrompt = SideHint()
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
            UpdatePrompt()
            Scene.Colour1 = _auto.RightColour
            Redraw()
        Else
            StopSuccessive()
            If PatientChooses Then ShowChoiceScreen() Else ShowWhite()
            _respondEnabled = True
            UpdatePanels()
            If _pendingFirst Then
                _pendingFirst = False
                Respond(True)
            End If
        End If
    End Sub

    Private Sub StopSuccessive()
        _timer.Stop()
        _successiveStage = 0
    End Sub

    Private Sub Respond(preferLeft As Boolean)
        If _auto Is Nothing OrElse Not _respondEnabled OrElse _auto.IsDone Then Return
        _auto.Respond(preferLeft)
        wheelAuto.Spokes = _auto.Chosen
        wheelAuto.Invalidate()
        If _auto.IsDone Then
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
            Scene.ChoicePrompt = ""
            Scene.ChoiceButtons = Nothing
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

#Region "Choosing directly"

    ''' <summary>The operator clicks on the copy of the patient's screen.</summary>
    Public Overrides Sub PreviewMouseClick(bounds As Rectangle, e As MouseEventArgs)
        ClickChoice(bounds, e.Location)
    End Sub

    ''' <summary>The patient clicks (or taps on the tablet) on their own screen.</summary>
    Public Overrides Sub PatientScreenClick(bounds As Rectangle, e As MouseEventArgs)
        ClickChoice(bounds, e.Location)
    End Sub

    ''' <summary>
    ''' The patient's answer is shown on their screen at once (the side or button turns green with a
    ''' tick) and on the examiner's (a short note), then acted on after a moment.
    ''' </summary>
    Private Sub Acknowledge(highlight As Integer, noteKey As String, action As Action)
        If _ackAction IsNot Nothing Then Return            ' one answer at a time
        _ackAction = action
        Scene.Highlight = highlight
        Redraw()
        If noteKey <> "" Then Toast.Show(Host?.FindForm(), Loc.F("colour.patient.chose", Loc.T(noteKey)), ToastKind.Info, 1500)
        _ackTimer.Start()
    End Sub

    Public Overrides ReadOnly Property PreviewCursor As Cursor
        Get
            Return If(Run.Method = 0 AndAlso _auto IsNot Nothing AndAlso Not _auto.IsDone, Cursors.Hand, Cursors.Default)
        End Get
    End Property

    ''' <summary>
    ''' A click on the preferred colour: overlays - the side clicked; lenses / computer - the colour
    ''' showing when the screen was clicked (a click during the first colour counts once the second
    ''' has been shown too, so both are always seen).
    ''' </summary>
    Private Sub ClickChoice(bounds As Rectangle, p As Point)
        ' Manual overlays, patient chooses: the side clicked is the colour chosen
        If _ackAction IsNot Nothing Then Return              ' the last answer is still being shown
        If Run.Method = 1 AndAlso IsOverlay Then
            If PatientChooses Then
                Dim isLeft = p.X < bounds.X + bounds.Width \ 2
                Dim c = If(isLeft, _manualLeft, _manualRight)
                Acknowledge(If(isLeft, 0, 1), If(isLeft, "colour.patient.left", "colour.patient.right"), Sub() SelectColour(c, 1))
            End If
            Return
        End If
        If Run.Method <> 0 OrElse _auto Is Nothing OrElse _auto.IsDone Then Return
        ' The patient's buttons (lenses / computer): The first, The second, Show again
        If Scene.ChoiceButtons IsNot Nothing Then
            Select Case Scene.ChoiceAt(bounds, p)
                Case 0 : Acknowledge(0, "colour.patient.first", Sub() Respond(True))
                Case 1 : Acknowledge(1, "colour.patient.second", Sub() Respond(False))
                Case 2 : Acknowledge(2, "colour.patient.again", Sub() Present())
            End Select
            Return
        End If
        If IsOverlay OrElse _successiveStage = 0 Then
            ' Two copies side by side (or the grey / saturation pair after the successive ones were shown)
            If Not _respondEnabled Then Return
            If IsOverlay Then
                Dim isLeft = p.X < bounds.X + bounds.Width \ 2
                Acknowledge(If(isLeft, 0, 1), If(isLeft, "colour.patient.left", "colour.patient.right"), Sub() Respond(isLeft))
            End If
            Return
        End If
        If _successiveStage = 1 Then
            _pendingFirst = True                      ' chosen; the second colour is still shown
        ElseIf _successiveStage = 2 Then
            _pendingFirst = False
            StopSuccessive()
            ShowWhite()
            _respondEnabled = True
            Respond(False)
        End If
    End Sub

    ''' <summary>Quick test: the patient points at a colour on the wheel; the test goes on from it.</summary>
    Private Sub PointAt(index As Integer)
        If Run.Method <> 0 OrElse Not AppHost.Settings.ColourAutoQuick OrElse index < 1 Then Return
        If _auto Is Nothing OrElse _auto.IsDone Then
            If Mode = ReadEZTestMode.Lenses AndAlso _auto Is Nothing Then Ask.Info(Host?.FindForm(), "colour.darkRoom")
            _auto = New QuickColourSearch(Mode)
        End If
        If Not _auto.CanJumpTo Then Return
        _auto.JumpTo(index)
        wheelAuto.Spokes = _auto.Chosen
        wheelAuto.Selected = (If(IsOverlay, ReadEZSet.Overlay, ReadEZSet.Lens), index)
        wheelAuto.Invalidate()
        Present()
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

    Public Overrides Function CreateTextOptions() As Control
        If Run.Result.IsPattern Then Return GratingOptions.Create(Me, Sub() Redraw(), enabledWhen:=Function() _successiveStage = 0)
        Return TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=IsOverlay, enabledWhen:=Function() _successiveStage = 0)
    End Function

    Public Overrides Function QuickActions() As List(Of QuickAction)
        Dim list As New List(Of QuickAction)
        If Run.Method = 0 Then
            Dim running = _auto IsNot Nothing AndAlso Not _auto.IsDone
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
        _ackTimer.Dispose()
        _timer.Dispose()
        MyBase.Dispose()
    End Sub
End Class
