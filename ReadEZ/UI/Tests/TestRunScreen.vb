Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   TestRunScreen.vb - where a ReadEZ test is run (replaces RZfrmTestControl and RZfrmSingle)
'
'     Header : the step's name, Back, "Patient full screen" (one screen), "?"
'     Steps  : the steps of this test (e.g. 1 Colour  2 Rate of reading  3 Symptoms), Skip / Back
'     Centre : the patient's view (PatientView). Two screens: an exact copy of the patient's
'              screen. One screen: the patient reads from this panel (or from the whole screen,
'              F11), as ReadEZ 2020's single-screen mode.
'     Right  : the step's own controls (each step module supplies these)
'
'   A test is a TestRun (TestRun.vb): a list of steps that fill in one ReadEZResult, saved when the
'   last step is finished.
' ====================================================================================================

#Region "Step base"

''' <summary>One step of a test (symptoms, rate of reading, colour choice) on the TestRunScreen.</summary>
Public MustInherit Class TestStep
    Implements IDisposable

    ''' <summary>Translation key of the step's name.</summary>
    Public MustOverride ReadOnly Property TitleKey As String

    ''' <summary>Help topic for the "?" button while this step is showing.</summary>
    Public Overridable ReadOnly Property HelpTopic As String
        Get
            Return "screen.tests"
        End Get
    End Property

    ''' <summary>The test this step belongs to.</summary>
    Public Property Run As TestRun

    ''' <summary>The screen this step is shown on (set before Begin).</summary>
    Public Property Host As TestRunScreen

    ''' <summary>What the patient sees (also drawn in the clinician's view).</summary>
    Public ReadOnly Property Scene As New ReadingScene

    ''' <summary>The step's controls for the right-hand panel.</summary>
    Public MustOverride Function CreatePanel() As Control

    ''' <summary>Buttons for the small control bar shown over the patient's view in one-screen full screen.</summary>
    Public Overridable Function QuickActions() As List(Of QuickAction)
        Return New List(Of QuickAction)
    End Function

    ''' <summary>Called when the step is shown: put the scene on the patient's screen.</summary>
    Public Overridable Sub Begin()
        ShowScene()
    End Sub

    ''' <summary>Called when the step is left (finished, skipped or cancelled).</summary>
    Public Overridable Sub [End]()
    End Sub

    ''' <summary>True if the step has what it needs to go on (otherwise Next asks first).</summary>
    Public Overridable Function IsComplete() As Boolean
        Return True
    End Function

    ''' <summary>Mouse over / click on the patient's view (scene coordinates are worked out by the scene).</summary>
    Public Overridable Sub PreviewMouseMove(bounds As Rectangle, p As Point)
    End Sub

    Public Overridable Sub PreviewMouseClick(bounds As Rectangle, e As MouseEventArgs)
    End Sub

    Public Overridable ReadOnly Property PreviewCursor As Cursor
        Get
            Return Cursors.Default
        End Get
    End Property

    ''' <summary>A key pressed on the test screen or the patient's screen. Return True if used.</summary>
    Public Overridable Function HandleKey(keyData As Keys) As Boolean
        Return False
    End Function

    ''' <summary>Puts the scene on the patient's screen (and redraws the views).</summary>
    Protected Sub ShowScene()
        Stimulus.Window.BackColor = If(Scene.Presentation = ReadingPresentation.FullScreen OrElse Scene.Presentation = ReadingPresentation.ColourOnly,
                                       Scene.Colour1, Color.White)
        Stimulus.Window.Scene = Scene
        Host?.RefreshView()
    End Sub

    ''' <summary>Redraws the patient's screen and the views after the scene was changed.</summary>
    Protected Sub Redraw()
        Stimulus.Window.BackColor = If(Scene.Presentation = ReadingPresentation.FullScreen OrElse Scene.Presentation = ReadingPresentation.ColourOnly,
                                       Scene.Colour1, Color.White)
        If Stimulus.Window.Scene IsNot Scene Then Stimulus.Window.Scene = Scene Else Stimulus.Window.RequestRedraw()
        Host?.RefreshView()
    End Sub

    ''' <summary>Sets the text rows and look from the settings (or from a result being repeated).</summary>
    Protected Sub ApplyTextSettings()
        Dim s = AppHost.Settings
        Dim text = ReadingTexts.ByIndex(s.RoR_TextOptions)
        Scene.Rows = ReadingTexts.RowsFor(text, s.RoR_NoRows, s.RoR_Random = 1)
        Scene.FontName = s.RoR_FontName
        Scene.FontSize = s.RoR_FontSize
        Scene.LineSpacing = s.RoR_LineSpacing
        Scene.OverlaySize = s.Overlay_Size
    End Sub

    ''' <summary>The presentation and colour for reading in this test's colour (white for a baseline).</summary>
    Protected Sub ApplyTestColour()
        Dim r = Run.Result
        If r.IsBaseline OrElse Not Run.HasColour Then
            Scene.Presentation = ReadingPresentation.White
            Scene.Colour1 = Color.White
        Else
            Scene.Colour1 = r.ScreenColour
            Scene.Presentation = If(r.Mode = ReadEZTestMode.Overlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen)
        End If
    End Sub

    ''' <summary>Tells the test this step is done (the next one starts).</summary>
    Protected Sub Finish()
        Host?.NextStep()
    End Sub

    ''' <summary>The step's controls changed state (the control bar is rebuilt).</summary>
    Protected Sub ActionsChanged()
        Host?.RefreshQuickActions()
    End Sub

    Public Overridable Sub Dispose() Implements IDisposable.Dispose
    End Sub
End Class

''' <summary>A button on the one-screen control bar.</summary>
Public NotInheritable Class QuickAction
    Public Property Key As String
    Public Property Action As Action
    Public Property Primary As Boolean
    Public Property Enabled As Boolean = True
    Public Sub New(key As String, action As Action, Optional primary As Boolean = False, Optional enabled As Boolean = True)
        Me.Key = key
        Me.Action = action
        Me.Primary = primary
        Me.Enabled = enabled
    End Sub
End Class

#End Region

Public Class TestRunScreen
    Inherits ScreenBase

    Private _run As TestRun
    Private _step As TestStep
    Private ReadOnly _view As New PatientView
    Private ReadOnly _options As New Card
    Private ReadOnly _optionsHost As New Panel
    Private ReadOnly _steps As New StepStrip
    Private ReadOnly btnFull As ModernButton
    Private ReadOnly btnPrevStep As ModernButton = ModernButton.Create("run.prevStep", ButtonKind.Ghost, Glyphs.Back)
    Private ReadOnly btnSkip As ModernButton = ModernButton.Create("run.skipStep", ButtonKind.Ghost, Glyphs.Forward)
    Private ReadOnly lblMode As New Label With {.AutoSize = True, .Font = Fonts.Small}
    Private _strip As ControlStrip

    Public Sub New()
        TitleKey = "tests.title"
        Dim back = AddHeaderButton("tests.back", ButtonKind.Ghost, Glyphs.Back)
        AddHandler back.Click, Sub() AppHost.MainWindow?.ShowScreen(MainForm.ScreenTests)
        btnFull = AddHeaderButton("run.fullScreen", ButtonKind.Secondary, Glyphs.ZoomIn)
        AddHandler btnFull.Click, Sub() ToggleFullScreen()

        ' Steps bar
        Dim bar As New Panel With {.Dock = DockStyle.Top, .Height = 52, .BackColor = Color.Transparent}
        _steps.Dock = DockStyle.Fill
        Dim right As New FlowLayoutPanel With {.Dock = DockStyle.Right, .AutoSize = True, .WrapContents = False, .BackColor = Color.Transparent, .Padding = New Padding(0, 8, 0, 0)}
        LabelStyle.SetMuted(lblMode)
        lblMode.Margin = New Padding(0, 9, 16, 0)
        right.Controls.Add(lblMode)
        right.Controls.Add(btnPrevStep)
        right.Controls.Add(btnSkip)
        AddHandler btnPrevStep.Click, Sub() PreviousStep()
        AddHandler btnSkip.Click, Sub() NextStep()
        bar.Controls.Add(_steps)
        bar.Controls.Add(right)

        _options.Dock = DockStyle.Right
        _options.Width = 360
        _options.Padding = New Padding(16, 52, 16, 12)
        _options.TitleBanner = True
        _options.SetTitle("tests.options")
        _optionsHost.Dock = DockStyle.Fill
        _optionsHost.AutoScroll = True
        _optionsHost.BackColor = Color.Transparent
        _options.Controls.Add(_optionsHost)
        Dim gap As New Panel With {.Dock = DockStyle.Right, .Width = 12, .BackColor = Color.Transparent}

        _view.Dock = DockStyle.Fill
        AddHandler _view.SceneMouseMove, Sub(b, p) _step?.PreviewMouseMove(b, p)
        AddHandler _view.SceneMouseClick, Sub(b, e) _step?.PreviewMouseClick(b, e)
        _view.CursorFor = Function() If(_step?.PreviewCursor, Cursors.Default)

        Content.Controls.Add(_view)
        Content.Controls.Add(gap)
        Content.Controls.Add(_options)
        Content.Controls.Add(bar)

        AddHandler Stimulus.Window.KeyPressedOnScreen, AddressOf KeyFromPatientScreen
        Loc.OnChange(Me, Sub() UpdateModeLabel())
    End Sub

#Region "Run and steps"

    ''' <summary>The test being run. Set before the screen is shown.</summary>
    Public Sub StartRun(run As TestRun)
        EndCurrentStep()
        _run = run
        _run.Index = 0
        ShowStep()
    End Sub

    Private Sub ShowStep()
        EndCurrentStep()
        For Each c As Control In _optionsHost.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        If _run Is Nothing OrElse _run.Index >= _run.Steps.Count Then Return
        _step = _run.Steps(_run.Index).Invoke()
        _step.Run = _run
        _step.Host = Me
        TitleKey = _step.TitleKey
        _options.SetTitle(_step.TitleKey)
        Dim panel = _step.CreatePanel()
        If panel IsNot Nothing Then
            panel.Dock = DockStyle.Top
            _optionsHost.Controls.Add(panel)
            Dpi.ScaleNew(panel, _optionsHost)
            ThemeManager.Apply(panel)
        End If
        _view.Scene = _step.Scene
        _steps.SetSteps(_run.StepTitles(), _run.Index)
        btnPrevStep.Visible = _run.Index > 0
        btnSkip.Visible = _run.Index < _run.Steps.Count - 1
        RefreshTitle()
        _step.Begin()
        RefreshQuickActions()
        RefreshView()
    End Sub

    Private Sub EndCurrentStep()
        If _step Is Nothing Then Return
        Try
            _step.End()
        Catch ex As Exception
            Log.Error("End step", ex)
        End Try
        _step.Dispose()
        _step = Nothing
    End Sub

    ''' <summary>Goes on to the next step, or saves the test after the last one.</summary>
    Public Sub NextStep()
        If _run Is Nothing OrElse _step Is Nothing Then Return
        If Not _step.IsComplete() AndAlso Not Ask.YesNo(FindForm(), "run.stepNotComplete") Then Return
        If _run.Index < _run.Steps.Count - 1 Then
            _run.Index += 1
            BeginInvoke(Sub() ShowStep())
            Return
        End If
        BeginInvoke(Sub() FinishRun())
    End Sub

    Private Sub PreviousStep()
        If _run Is Nothing OrElse _run.Index = 0 Then Return
        _run.Index -= 1
        ShowStep()
    End Sub

    Private Sub FinishRun()
        EndCurrentStep()
        SetFullScreen(False)
        Dim r = _run.Result
        Try
            r.Test_Date = If(r.ID > 0, r.Test_Date, Date.Now)
            AppHost.Db.SaveResult(r)
            AppHost.Session.NotifyResultSaved(r)
            AppHost.Session.SelectResult(r.ID)
            Toast.Show(FindForm(), Loc.T("run.saved"), ToastKind.Success)
        Catch ex As Exception
            Log.Error("Save test", ex)
            Ask.Error(FindForm(), "common.saveFailed", ex.Message)
            Return
        End Try
        _run = Nothing
        Stimulus.ShowLogo()
        AppHost.MainWindow?.ShowScreen(MainForm.ScreenResults)
    End Sub

    ''' <summary>The step's name changed (e.g. automatic / manual colour test).</summary>
    Public Sub RefreshTitleFor(st As TestStep)
        If st IsNot _step Then Return
        TitleKey = st.TitleKey
        _options.SetTitle(st.TitleKey)
        RefreshTitle()
        RefreshQuickActions()
    End Sub

    Public ReadOnly Property IsRunning As Boolean
        Get
            Return _run IsNot Nothing
        End Get
    End Property

#End Region

#Region "Activation"

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        Dim s = AppHost.Session
        SetSubtitleText(If(s.HasPatient, s.PatientName, ""))
        Stimulus.Window.SetInUse("test", True)
        UpdateModeLabel()
        If _step IsNot Nothing Then _step.Begin()
        If Not ScreenLayout.IsTwoScreenMode() AndAlso AppHost.Settings.SingleScreenFullWindow Then BeginInvoke(Sub() SetFullScreen(True))
    End Sub

    Public Overrides Function CanLeave() As Boolean
        If _run IsNot Nothing Then
            If Not Ask.YesNoWarning(FindForm(), "run.confirmLeave") Then Return False
            EndCurrentStep()
            _run = Nothing
        End If
        Return True
    End Function

    Public Overrides Sub OnDeactivated()
        SetFullScreen(False)
        Stimulus.Window.SetInUse("test", False)
        Stimulus.ShowLogo()
        MyBase.OnDeactivated()
    End Sub

    Protected Overrides Function CurrentHelpTopic() As String
        Return If(_step?.HelpTopic, "screen.tests")
    End Function

    Private Sub UpdateModeLabel()
        Dim two = ScreenLayout.IsTwoScreenMode()
        lblMode.Text = Loc.T(If(two, "run.mode.two", "run.mode.one"))
        btnFull.Visible = Not two
        _view.Caption = Loc.T(If(two, "run.view.copy", "run.view.patient"))
    End Sub

    ''' <summary>Redraws the patient's view.</summary>
    Public Sub RefreshView()
        _view.Invalidate()
        _strip?.Invalidate()
    End Sub

    Public Sub RefreshQuickActions()
        _strip?.SetActions(If(_step?.QuickActions(), New List(Of QuickAction)), If(_step Is Nothing, "", Loc.T(_step.TitleKey)))
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveHandler Stimulus.Window.KeyPressedOnScreen, AddressOf KeyFromPatientScreen
            _strip?.Dispose()
            _step?.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

#Region "One screen: patient full screen"

    Private Sub ToggleFullScreen()
        SetFullScreen(Not Stimulus.Window.IsInFront)
    End Sub

    ''' <summary>One screen: the patient's screen in front with a small control bar (ReadEZ 2020's single-screen mode).</summary>
    Private Sub SetFullScreen(full As Boolean)
        If full AndAlso ScreenLayout.IsTwoScreenMode() Then Return
        If Not full AndAlso Not Stimulus.Window.IsInFront Then Return
        If full Then
            If _strip Is Nothing OrElse _strip.IsDisposed Then
                _strip = New ControlStrip()
                AddHandler _strip.ExitClicked, Sub() SetFullScreen(False)
                AddHandler _strip.KeyPressed, Sub(k) HandleKey(k)
            End If
            Stimulus.Window.SetFront(True)
            RefreshQuickActions()
            _strip.ShowOver(Stimulus.Window)
        Else
            _strip?.Hide()
            Stimulus.Window.SetFront(False)
            Stimulus.Window.SendToBack()
            FindForm()?.Activate()
        End If
    End Sub

#End Region

#Region "Keys"

    Private Sub KeyFromPatientScreen(keyData As Keys)
        If Not Visible Then Return
        HandleKey(keyData)
    End Sub

    Private Function HandleKey(keyData As Keys) As Boolean
        If keyData = Keys.F11 OrElse (keyData = Keys.Escape AndAlso Stimulus.Window.IsInFront) Then
            If Not ScreenLayout.IsTwoScreenMode() Then ToggleFullScreen()
            Return True
        End If
        Return _step IsNot Nothing AndAlso _step.HandleKey(keyData)
    End Function

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        ' Keys typed into a box stay there
        Dim focused = FindForm()?.ActiveControl
        Dim typing = TypeOf focused Is TextBoxBase OrElse TypeOf focused Is NumericUpDown OrElse TypeOf focused Is ComboBox
        If Visible AndAlso (Not typing OrElse keyData = Keys.F11) AndAlso HandleKey(keyData) Then Return True
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

#End Region

End Class

''' <summary>The patient's screen drawn in the clinician's window, at the screen's shape.</summary>
Public NotInheritable Class PatientView
    Inherits Control
    Implements IThemed

    Public Property Scene As IStimulusScene
    Public Property Caption As String = ""
    Public Property CursorFor As Func(Of Cursor)

    Public Event SceneMouseMove(bounds As Rectangle, p As Point)
    Public Event SceneMouseClick(bounds As Rectangle, e As MouseEventArgs)

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        AddHandler Stimulus.Window.Changed, Sub() Invalidate()
    End Sub

    ''' <summary>Where the patient's screen is drawn (its shape kept), below the caption.</summary>
    Public Function ScreenBounds() As Rectangle
        Dim sz = StimulusArt.ScreenSize
        Dim top = 26
        Dim availW = Math.Max(10, Width - 2), availH = Math.Max(10, Height - top - 2)
        Dim k = Math.Min(availW / CDbl(sz.Width), availH / CDbl(sz.Height))
        Dim w = CInt(sz.Width * k), h = CInt(sz.Height * k)
        Return New Rectangle((Width - w) \ 2, top + (availH - h) \ 2, w, h)
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        TextRenderer.DrawText(g, Caption, Fonts.SmallStrong, New Rectangle(0, 0, Width, 22), p.Muted, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
        Dim r = ScreenBounds()
        g.SetClip(r)
        If Stimulus.IsOpen Then
            Stimulus.Window.RenderPreview(g, r)
        Else
            Scene?.Paint(g, r, True)
        End If
        g.ResetClip()
        Using pen As New Pen(p.Border, 2)
            g.DrawRectangle(pen, r.X - 1, r.Y - 1, r.Width + 1, r.Height + 1)
        End Using
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim r = ScreenBounds()
        If r.Contains(e.Location) Then
            RaiseEvent SceneMouseMove(r, e.Location)
            Cursor = If(CursorFor?.Invoke(), Cursors.Default)
        Else
            Cursor = Cursors.Default
        End If
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        Dim r = ScreenBounds()
        If r.Contains(e.Location) Then RaiseEvent SceneMouseClick(r, e)
    End Sub

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property
End Class

''' <summary>The steps of a test as numbered chips ("1 Colour  ›  2 Rate of reading  ›  3 Symptoms").</summary>
Public NotInheritable Class StepStrip
    Inherits Control
    Implements IThemed

    Private _titles As New List(Of String)
    Private _current As Integer

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
    End Sub

    Public Sub SetSteps(titles As List(Of String), current As Integer)
        _titles = titles
        _current = current
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim x = 0
        Dim d = LogicalToDeviceUnits(26)
        Dim y = (Height - d) \ 2
        For i = 0 To _titles.Count - 1
            Dim done = i < _current, now = i = _current
            Dim back = If(now, p.Primary, If(done, p.Success, p.SurfaceAlt))
            Using b As New SolidBrush(back)
                g.FillEllipse(b, x, y, d, d)
            End Using
            TextRenderer.DrawText(g, If(done, ChrW(&H2713).ToString(), (i + 1).ToString()), Fonts.SmallStrong, New Rectangle(x, y, d, d),
                                  If(now OrElse done, Color.White, p.Muted), TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            x += d + LogicalToDeviceUnits(8)
            Dim f = If(now, Fonts.Strong, Fonts.Body)
            Dim w = TextRenderer.MeasureText(_titles(i), f).Width
            TextRenderer.DrawText(g, _titles(i), f, New Rectangle(x, 0, w + 4, Height), If(now, p.Ink, p.Muted), TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
            x += w + LogicalToDeviceUnits(10)
            If i < _titles.Count - 1 Then
                TextRenderer.DrawText(g, "›", Fonts.Section, New Rectangle(x, 0, 20, Height), p.Muted, TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
                x += LogicalToDeviceUnits(24)
            End If
        Next
    End Sub

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property
End Class

''' <summary>
''' One screen, patient full screen: a small bar at the bottom of the patient's screen with the
''' step's buttons (as ReadEZ 2020's control bar under the test screen). Esc / F11 or "Controls"
''' goes back to the full test screen.
''' </summary>
Public NotInheritable Class ControlStrip
    Inherits Form

    Private ReadOnly _row As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .WrapContents = False, .Padding = New Padding(10, 8, 10, 8), .BackColor = Color.Transparent}
    Private ReadOnly _title As New Label With {.AutoSize = True, .Font = Fonts.Strong, .ForeColor = Color.White, .Margin = New Padding(4, 7, 16, 0)}

    Public Event ExitClicked As EventHandler
    Public Event KeyPressed(keyData As Keys)

    Public Sub New()
        FormBorderStyle = FormBorderStyle.None
        StartPosition = FormStartPosition.Manual
        ShowInTaskbar = False
        TopMost = True
        KeyPreview = True
        AutoScaleMode = AutoScaleMode.None
        BackColor = Brand.Navy
        Opacity = 0.94
        Font = Fonts.Body
        ThemeManager.SetRole(Me, SurfaceRole.Stimulus)
        Controls.Add(_row)
    End Sub

    Public Sub SetActions(actions As List(Of QuickAction), title As String)
        _row.SuspendLayout()
        For Each c In _row.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        _title.Text = title
        _row.Controls.Add(_title)
        For Each a In actions
            Dim act = a
            Dim b = ModernButton.Create(a.Key, If(a.Primary, ButtonKind.Secondary, ButtonKind.OnDark))
            b.Enabled = a.Enabled
            b.Margin = New Padding(0, 0, 8, 0)
            AddHandler b.Click, Sub() act.Action.Invoke()
            _row.Controls.Add(b)
        Next
        Dim back = ModernButton.Create("run.showControls", ButtonKind.OnDark, Glyphs.ZoomOut)
        back.Margin = New Padding(16, 0, 0, 0)
        AddHandler back.Click, Sub() RaiseEvent ExitClicked(Me, EventArgs.Empty)
        _row.Controls.Add(back)
        _row.ResumeLayout(True)
        Dpi.ScaleNew(_row, Me)
        ThemeManager.Apply(_row)
        FitWidth()
    End Sub

    Private _owner As Form

    Public Sub ShowOver(window As Form)
        _owner = window
        If Not Visible Then Show(window)
        FitWidth()
        BringToFront()
    End Sub

    Private Sub FitWidth()
        If _owner Is Nothing Then Return
        Dim w = Math.Min(_owner.Width - 40, _row.Controls.Cast(Of Control)().Sum(Function(c) c.Width + c.Margin.Horizontal) + 40)
        Dim h = Dpi.Px(Me, 54)
        Bounds = New Rectangle(_owner.Left + (_owner.Width - w) \ 2, _owner.Bottom - h - 16, w, h)
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        RaiseEvent KeyPressed(keyData)
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property
End Class
