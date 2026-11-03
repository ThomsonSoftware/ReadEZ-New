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

    ''' <summary>
    ''' The patient's instructions, shown on the patient's screen before the step starts: the title's
    ''' translation key, then one key per paragraph. Nothing = none.
    ''' </summary>
    Public Overridable Function InstructionKeys() As String()
        Return Nothing
    End Function

    ''' <summary>
    ''' The patient's screen while the instructions are read (as Eye Dynamix's sample screens: what
    ''' the test will look like, with numbered labels). Nothing = the instructions as text.
    ''' </summary>
    Public Overridable Function InstructionScreen() As InstructionScene
        Return Nothing
    End Function

    ''' <summary>The test's reading text as a sample (its text settings, on white).</summary>
    Protected Function SampleText() As ReadingScene
        Dim s = AppHost.Settings
        Dim r = Run.Result
        Dim sample As New ReadingScene With {.FontName = r.RoR_FontName, .FontSize = r.RoR_FontSize, .LineSpacing = r.RoR_LineSpacing,
                                             .OverlaySize = If(r.RoR_OverlaySize > 0, r.RoR_OverlaySize, s.Overlay_Size)}
        sample.Rows = ReadingTexts.RowsFor(ReadingTexts.ByIndex(r.RoR_TextOptions), r.RoR_NoRows, s.RoR_Random = 1)
        Return sample
    End Function

    ''' <summary>Buttons kept at the bottom of the panel, always visible (Next step ...). Nothing = none.</summary>
    Public Overridable Function CreateFooter() As Control
        Return Nothing
    End Function

    ''' <summary>A row of footer buttons.</summary>
    Protected Shared Function FooterRow(ParamArray buttons As Control()) As Control
        Dim row As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = False,
                                             .BackColor = Color.Transparent, .Margin = New Padding(0), .Padding = New Padding(0)}
        For Each b In buttons
            b.Margin = New Padding(0, 0, 8, 0)
            row.Controls.Add(b)
        Next
        Return row
    End Function

    ''' <summary>A "How to ..." button that shows instructions (translation keys, one paragraph each) in a window.</summary>
    Protected Function HelpButton(buttonKey As String, titleKey As String, ParamArray lineKeys As String()) As ModernButton
        Dim b = ModernButton.Create(buttonKey, ButtonKind.Ghost, Glyphs.Help)
        b.Margin = New Padding(0, 0, 0, 8)
        AddHandler b.Click, Sub()
                                Using d As New InstructionsDialog(titleKey, lineKeys)
                                    d.ShowDialog(Host?.FindForm())
                                End Using
                            End Sub
        Return b
    End Function

    ''' <summary>The "Next step" / "Save and finish" button.</summary>
    Protected Function NextButton() As ModernButton
        Dim last = Run.Index = Run.Steps.Count - 1
        Dim b = ModernButton.Create(If(last, "run.saveFinish", "run.nextStep"), ButtonKind.Primary, If(last, Glyphs.Save, Glyphs.Forward))
        AddHandler b.Click, Sub() Finish()
        Return b
    End Function

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

    ''' <summary>A click on the patient's own screen (bounds = the whole screen).</summary>
    Public Overridable Sub PatientScreenClick(bounds As Rectangle, e As MouseEventArgs)
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

    ''' <summary>
    ''' Sets the text rows and look from the test (the settings when it began; a colour test the
    ''' latest baseline's, so the two can be compared).
    ''' </summary>
    Protected Sub ApplyTextSettings()
        Dim s = AppHost.Settings
        Dim r = Run.Result
        Dim text = ReadingTexts.ByIndex(r.RoR_TextOptions)
        Scene.Rows = ReadingTexts.RowsFor(text, r.RoR_NoRows, s.RoR_Random = 1)
        Scene.FontName = r.RoR_FontName
        Scene.FontSize = r.RoR_FontSize
        Scene.LineSpacing = r.RoR_LineSpacing
        Scene.OverlaySize = If(r.RoR_OverlaySize > 0, r.RoR_OverlaySize, s.Overlay_Size)
    End Sub

    ''' <summary>Saves the test now, leaving out the steps after this one (e.g. the colour only).</summary>
    Protected Sub SaveNow()
        Host?.SaveNow()
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

''' <summary>
''' The patient's instructions on the patient's screen, as Eye Dynamix's sample screens: a sample of
''' what the test will show (the text, the two overlays, the two colours ...), a title at the top and
''' numbered labels in blue with arrows to the parts that matter ("1. Start reading here"). Drawn in
''' the patient's screen pixels and scaled, so the examiner's copy is identical. Without a sample,
''' the title and paragraphs are shown on white.
''' </summary>
Public NotInheritable Class InstructionScene
    Implements IStimulusScene

    ''' <summary>Where a label goes relative to the part it points at.</summary>
    Public Enum Side
        Left
        Right
        Above
        Below
        Inside
        BottomCentre
    End Enum

    Public NotInheritable Class Callout
        Public TextKey As String
        ''' <summary>The part pointed at (patient's screen pixels); Empty = no arrow.</summary>
        Public Target As RectangleF
        Public Placement As Side
        Public Sub New(textKey As String, target As RectangleF, placement As Side)
            Me.TextKey = textKey : Me.Target = target : Me.Placement = placement
        End Sub
    End Class

    ' Eye Dynamix's colours: title dark grey, labels and arrows blue
    Private Shared ReadOnly TitleInk As Color = Color.FromArgb(60, 60, 70)
    Private Shared ReadOnly LabelInk As Color = Color.FromArgb(40, 110, 200)

    Private ReadOnly _keys As String()

    ''' <summary>The title's key (keys(0) when Nothing).</summary>
    Public Property TitleKey As String
    ''' <summary>What the test will look like, drawn underneath (Nothing = white with the paragraphs).</summary>
    Public Property Sample As IStimulusScene
    Public ReadOnly Property Callouts As New List(Of Callout)

    Public Sub New(keys As String())
        _keys = keys
    End Sub

    Public Sub Add(textKey As String, target As RectangleF, placement As Side)
        Callouts.Add(New Callout(textKey, target, placement))
    End Sub

    Public Sub Paint(g As Graphics, bounds As Rectangle, preview As Boolean) Implements IStimulusScene.Paint
        If Sample Is Nothing Then
            PaintText(g, bounds)
            Return
        End If
        Sample.Paint(g, bounds, False)
        Dim sz = StimulusArt.ScreenSize
        Dim k = bounds.Width / CSng(Math.Max(1, sz.Width))
        Dim state = g.Save()
        g.TranslateTransform(bounds.X, bounds.Y)
        g.ScaleTransform(k, k)
        Try
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
            Dim h = CSng(sz.Height)
            Using titleFont As New Font("Segoe UI Semibold", h * 0.045F, GraphicsUnit.Pixel),
                  labelFont As New Font("Segoe UI Semibold", h * 0.03F, GraphicsUnit.Pixel),
                  titleBrush As New SolidBrush(TitleInk), labelBrush As New SolidBrush(LabelInk),
                  backing As New SolidBrush(Color.FromArgb(225, 255, 255, 255)),
                  arrow As New Pen(LabelInk, Math.Max(2.0F, h * 0.004F)) With {.EndCap = Drawing2D.LineCap.ArrowAnchor},
                  centre As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                ' Title at the top
                Dim title = Loc.T(If(TitleKey, _keys(0)))
                Dim tr As New RectangleF(sz.Width * 0.1F, h * 0.03F, sz.Width * 0.8F, h * 0.08F)
                Plate(g, MeasureCentred(g, title, titleFont, tr), backing)
                g.DrawString(title, titleFont, titleBrush, tr, centre)
                ' Numbered labels with arrows
                Dim n = 0
                For Each c In Callouts
                    n += 1
                    Dim text = n & ".  " & Loc.T(c.TextKey)
                    Dim lr = LabelRect(g, text, labelFont, c, sz)
                    Plate(g, lr, backing)
                    g.DrawString(text, labelFont, labelBrush, lr, centre)
                    If Not c.Target.IsEmpty AndAlso c.Placement <> Side.Inside AndAlso c.Placement <> Side.BottomCentre Then
                        Dim a = ArrowEnds(lr, RectangleF.Inflate(c.Target, h * 0.012F, h * 0.012F), c.Placement)
                        g.DrawLine(arrow, a.Item1, a.Item2)
                    End If
                Next
            End Using
        Finally
            g.Restore(state)
        End Try
    End Sub

    ''' <summary>A soft white plate behind text, so it reads on any colour.</summary>
    Private Shared Sub Plate(g As Graphics, r As RectangleF, b As Brush)
        Dim p = r.Height * 0.25F
        Using path = Draw.RoundedRect(RectangleF.Inflate(r, p, p * 0.4F), r.Height * 0.3F)
            g.FillPath(b, path)
        End Using
    End Sub

    Private Shared Function MeasureCentred(g As Graphics, text As String, f As Font, box As RectangleF) As RectangleF
        Dim s = g.MeasureString(text, f, CInt(box.Width))
        Return New RectangleF(box.X + (box.Width - s.Width) / 2, box.Y + (box.Height - s.Height) / 2, s.Width, s.Height)
    End Function

    ''' <summary>Where a label goes: beside, above, below or inside the part, or centred at the bottom.</summary>
    Private Shared Function LabelRect(g As Graphics, text As String, f As Font, c As Callout, sz As Size) As RectangleF
        Dim w = sz.Width, h = sz.Height
        Dim t = c.Target
        Dim gap = h * 0.06F
        Dim maxW As Single
        Select Case c.Placement
            Case Side.Left : maxW = Math.Max(w * 0.12F, t.X - gap * 1.5F)
            Case Side.Right : maxW = Math.Max(w * 0.12F, w - t.Right - gap * 1.5F)
            Case Side.Inside : maxW = Math.Max(w * 0.12F, t.Width * 0.85F)
            Case Else : maxW = w * 0.7F
        End Select
        Dim s = g.MeasureString(text, f, CInt(maxW))
        Dim x, y As Single
        Select Case c.Placement
            Case Side.Left
                x = t.X - gap - s.Width : y = t.Y + t.Height / 2 - s.Height / 2
            Case Side.Right
                x = t.Right + gap : y = t.Y + t.Height / 2 - s.Height / 2
            Case Side.Above
                x = t.X + t.Width / 2 - s.Width / 2 : y = t.Y - gap - s.Height
            Case Side.Below
                x = t.X + t.Width / 2 - s.Width / 2 : y = t.Bottom + gap
            Case Side.Inside
                x = t.X + t.Width / 2 - s.Width / 2 : y = t.Y + t.Height / 2 - s.Height / 2
            Case Else
                x = w / 2.0F - s.Width / 2 : y = h * 0.9F - s.Height / 2
        End Select
        ' Keep it on the screen
        x = Math.Max(h * 0.02F, Math.Min(w - s.Width - h * 0.02F, x))
        y = Math.Max(h * 0.12F, Math.Min(h - s.Height - h * 0.02F, y))
        Return New RectangleF(x, y, s.Width, s.Height)
    End Function

    Private Shared Function ArrowEnds(label As RectangleF, target As RectangleF, placement As Side) As Tuple(Of PointF, PointF)
        Select Case placement
            Case Side.Left
                Dim y = label.Y + label.Height / 2
                Return Tuple.Create(New PointF(label.Right + 6, y), New PointF(target.X, Math.Max(target.Y, Math.Min(target.Bottom, y))))
            Case Side.Right
                Dim y = label.Y + label.Height / 2
                Return Tuple.Create(New PointF(label.X - 6, y), New PointF(target.Right, Math.Max(target.Y, Math.Min(target.Bottom, y))))
            Case Side.Above
                Dim x = label.X + label.Width / 2
                Return Tuple.Create(New PointF(x, label.Bottom + 6), New PointF(Math.Max(target.X, Math.Min(target.Right, x)), target.Y))
            Case Else
                Dim x = label.X + label.Width / 2
                Return Tuple.Create(New PointF(x, label.Y - 6), New PointF(Math.Max(target.X, Math.Min(target.Right, x)), target.Bottom))
        End Select
    End Function

    ''' <summary>No sample: the title and the paragraphs, large, on white.</summary>
    Private Sub PaintText(g As Graphics, bounds As Rectangle)
        g.FillRectangle(Brushes.White, bounds)
        If _keys Is Nothing OrElse _keys.Length = 0 Then Return
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        Dim h = CSng(bounds.Height)
        Dim w = bounds.Width * 0.66F
        Dim x = bounds.X + (bounds.Width - w) / 2
        Using titleFont As New Font("Segoe UI Semibold", Math.Max(6.0F, h * 0.055F), GraphicsUnit.Pixel),
              bodyFont As New Font("Segoe UI", Math.Max(5.0F, h * 0.034F), GraphicsUnit.Pixel),
              ink As New SolidBrush(TitleInk), blue As New SolidBrush(LabelInk)
            Dim parts As New List(Of (Text As String, Font As Font, Brush As Brush, Gap As Single))
            parts.Add((Loc.T(If(TitleKey, _keys(0))), titleFont, ink, h * 0.04F))
            For i = 1 To _keys.Length - 1
                parts.Add((i & ".  " & Loc.T(_keys(i)), bodyFont, blue, h * 0.022F))
            Next
            Dim heights = parts.Select(Function(pt) g.MeasureString(pt.Text, pt.Font, CInt(w)).Height).ToList()
            Dim total = heights.Sum() + parts.Sum(Function(pt) pt.Gap)
            Dim y = bounds.Y + Math.Max(h * 0.06F, (h - total) / 2)
            For i = 0 To parts.Count - 1
                g.DrawString(parts(i).Text, parts(i).Font, parts(i).Brush, New RectangleF(x, y, w, heights(i) + 2))
                y += heights(i) + parts(i).Gap
            Next
        End Using
    End Sub
End Class

''' <summary>Lenses / computer sample: the two colours side by side, numbered 1 and 2.</summary>
Public NotInheritable Class TwoColourSample
    Implements IStimulusScene

    Public Property Colour1 As Color
    Public Property Colour2 As Color

    Public Sub Paint(g As Graphics, bounds As Rectangle, preview As Boolean) Implements IStimulusScene.Paint
        Dim half = bounds.Width \ 2
        Using b1 As New SolidBrush(Colour1), b2 As New SolidBrush(Colour2)
            g.FillRectangle(b1, bounds.X, bounds.Y, half, bounds.Height)
            g.FillRectangle(b2, bounds.X + half, bounds.Y, bounds.Width - half, bounds.Height)
        End Using
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        Using f As New Font("Segoe UI Semibold", Math.Max(8.0F, bounds.Height * 0.2F), GraphicsUnit.Pixel), b As New SolidBrush(Color.FromArgb(70, 40, 40, 40)),
              fmt As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString("1", f, b, New RectangleF(bounds.X, bounds.Y, half, bounds.Height * 0.6F), fmt)
            g.DrawString("2", f, b, New RectangleF(bounds.X + half, bounds.Y, bounds.Width - half, bounds.Height * 0.6F), fmt)
        End Using
    End Sub

    ''' <summary>The half holding colour 1 or 2 (patient's screen pixels).</summary>
    Public Shared Function Half(n As Integer) As RectangleF
        Dim sz = StimulusArt.ScreenSize
        Dim w = sz.Width / 2.0F
        Return New RectangleF(If(n = 1, 0, w), 0, w, sz.Height)
    End Function
End Class

''' <summary>Instructions for a test step, shown in a window (the step's "How to ..." button).</summary>
Public NotInheritable Class InstructionsDialog
    Inherits DialogBase

    Public Sub New(titleKey As String, lineKeys As String())
        SetTitle(titleKey)
        ClientSize = New Size(600, 120 + 52 * Math.Max(1, lineKeys.Length))
        Dim col = Ui.Column(4)
        col.Dock = DockStyle.Fill
        col.AutoSize = False
        col.AutoScroll = True
        For Each k In lineKeys
            Dim l = Ui.Label(k)
            l.Font = Fonts.Body
            l.MaximumSize = New Size(540, 0)
            l.Margin = New Padding(0, 0, 0, 10)
            col.Controls.Add(l)
        Next
        Body.Controls.Add(col)
        Dim close = AddFooterButton("common.close", ButtonKind.Primary)
        close.DialogResult = DialogResult.OK
        AcceptButton = close
        CancelButton = close
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
    Private ReadOnly _footer As New Panel With {.Dock = DockStyle.Bottom, .Height = 58, .BackColor = Color.Transparent, .Padding = New Padding(0, 12, 0, 2)}
    Private ReadOnly _steps As New StepStrip
    Private ReadOnly btnFull As ModernButton
    Private ReadOnly btnPrevStep As ModernButton = ModernButton.Create("run.prevStep", ButtonKind.Ghost, Glyphs.Back)
    Private ReadOnly btnSkip As ModernButton = ModernButton.Create("run.skipStep", ButtonKind.Ghost, Glyphs.Forward)
    Private ReadOnly lblMode As New Label With {.AutoSize = True, .Font = Fonts.Small}
    Private _strip As ControlStrip
    ' Patient instructions before a step
    Private _inInstructions As Boolean
    Private _instructionScene As InstructionScene
    Private ReadOnly _instructed As New HashSet(Of Integer)

    Public Sub New()
        TitleKey = "tests.title"
        UseProminentSubtitle()                            ' the patient's name, large, beside the title
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
        _options.Width = 400
        _options.Padding = New Padding(16, 52, 16, 12)
        _options.TitleBanner = True
        _options.SetTitle("tests.options")
        _optionsHost.Dock = DockStyle.Fill
        _optionsHost.AutoScroll = True
        _optionsHost.BackColor = Color.Transparent
        _options.Controls.Add(_optionsHost)
        _options.Controls.Add(_footer)                     ' always visible under the step's controls
        AddHandler _footer.Paint, Sub(o, e)
                                      Using pen As New Pen(ThemeManager.Current.Border)
                                          e.Graphics.DrawLine(pen, 0, 2, _footer.Width, 2)
                                      End Using
                                  End Sub
        Dim gap As New Panel With {.Dock = DockStyle.Right, .Width = 12, .BackColor = Color.Transparent}

        _view.Dock = DockStyle.Fill
        AddHandler _view.SceneMouseMove, Sub(b, p) If Not _inInstructions Then _step?.PreviewMouseMove(b, p)
        AddHandler _view.SceneMouseClick, Sub(b, e) If Not _inInstructions Then _step?.PreviewMouseClick(b, e)
        _view.CursorFor = Function() If(_step?.PreviewCursor, Cursors.Default)

        Content.Controls.Add(_view)
        Content.Controls.Add(gap)
        Content.Controls.Add(_options)
        Content.Controls.Add(bar)

        AddHandler Stimulus.Window.KeyPressedOnScreen, AddressOf KeyFromPatientScreen
        AddHandler Stimulus.Window.ClickedOnScreen, AddressOf ClickFromPatientScreen
        Loc.OnChange(Me, Sub() UpdateModeLabel())
    End Sub

#Region "Run and steps"

    ''' <summary>The test being run. Set before the screen is shown.</summary>
    Public Sub StartRun(run As TestRun)
        EndCurrentStep()
        _run = run
        _run.Index = 0
        _instructed.Clear()
        ShowStep()
    End Sub

    Private Sub ShowStep()
        EndCurrentStep()
        For Each c As Control In _optionsHost.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        For Each c As Control In _footer.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        If _run Is Nothing OrElse _run.Index >= _run.Steps.Count Then Return
        _step = _run.Steps(_run.Index).Invoke()
        _step.Run = _run
        _step.Host = Me
        TitleKey = _step.TitleKey
        _options.SetTitle(_step.TitleKey)
        _steps.SetSteps(_run.StepTitles(), _run.Index)
        btnPrevStep.Visible = _run.Index > 0
        btnSkip.Visible = _run.Index < _run.Steps.Count - 1
        RefreshTitle()
        ' The patient's instructions first (once per step of a test)
        Dim keys = _step.InstructionKeys()
        If keys IsNot Nothing AndAlso keys.Length > 0 AndAlso AppHost.Settings.ShowPatientInstructions AndAlso Not _instructed.Contains(_run.Index) Then
            ShowInstructions(keys)
            Return
        End If
        StartStep()
    End Sub

    ''' <summary>Builds the step's controls and starts it.</summary>
    Private Sub StartStep()
        _inInstructions = False
        ClearPanel()
        PlacePanel(_step.CreatePanel())
        PlaceFooter(_step.CreateFooter())
        _view.Scene = _step.Scene
        _view.ZoomToText = False                          ' a step that wants the scoring copy turns it on in Begin
        _step.Begin()
        RefreshQuickActions()
        RefreshView()
    End Sub

    Private Sub ClearPanel()
        For Each c As Control In _optionsHost.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        For Each c As Control In _footer.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
    End Sub

    Private Sub PlacePanel(panel As Control)
        If panel Is Nothing Then Return
        LargerText(panel)
        panel.Dock = DockStyle.Top
        _optionsHost.Controls.Add(panel)
        Dpi.ScaleNew(panel, _optionsHost)
        ThemeManager.Apply(panel)
    End Sub

    Private Sub PlaceFooter(footer As Control)
        If footer IsNot Nothing Then
            footer.Dock = DockStyle.Fill
            _footer.Controls.Add(footer)
            Dpi.ScaleNew(footer, _footer)
            ThemeManager.Apply(footer)
        End If
        _footer.Visible = footer IsNot Nothing
    End Sub

#Region "Patient instructions"

    ''' <summary>
    ''' Shows the step's instructions on the patient's screen (large, on white) and, in the panel, the
    ''' same text for the examiner with a Begin button (Space / Enter).
    ''' </summary>
    Private Sub ShowInstructions(keys As String())
        _inInstructions = True
        ClearPanel()
        Dim col = Ui.Column()
        Ui.FillWidth(col)
        ' As Eye Dynamix: what to say, in large type (the patient's screen shows the sample)
        Dim lead = Ui.Label("instr.panel", muted:=True)
        lead.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(lead)
        For i = 1 To keys.Length - 1
            Dim l = Ui.Label(keys(i))
            l.Font = Fonts.Section
            LabelStyle.SetCustom(l)
            l.ForeColor = ThemeManager.Current.Ink
            l.Margin = New Padding(0, 0, 0, 12)
            col.Controls.Add(l)
        Next
        Dim always = ThemedCheckBox.Create("instr.always")
        always.Checked = AppHost.Settings.ShowPatientInstructions
        always.Margin = New Padding(0, 12, 0, 0)
        AddHandler always.CheckedChanged, Sub()
                                              AppHost.Settings.ShowPatientInstructions = always.Checked
                                              AppHost.SaveSettings()
                                          End Sub
        col.Controls.Add(always)
        PlacePanel(col)
        Dim begin = ModernButton.Create("instr.start", ButtonKind.Primary, Glyphs.Play)
        AddHandler begin.Click, Sub() BeginAfterInstructions()
        Dim row As New FlowLayoutPanel With {.FlowDirection = FlowDirection.LeftToRight, .WrapContents = False, .BackColor = Color.Transparent}
        row.Controls.Add(begin)
        PlaceFooter(row)
        _instructionScene = Nothing
        Try
            _instructionScene = _step.InstructionScreen()
        Catch ex As Exception
            Log.Error("Instruction screen", ex)
        End Try
        If _instructionScene Is Nothing Then _instructionScene = New InstructionScene(keys)
        ShowInstructionScene()
        RefreshQuickActions()
    End Sub

    Private Sub ShowInstructionScene()
        Stimulus.Window.BackColor = Color.White
        Stimulus.Window.Scene = _instructionScene
        _view.ZoomToText = False
        _view.Scene = _instructionScene
        RefreshView()
    End Sub

    Private Sub BeginAfterInstructions()
        If Not _inInstructions OrElse _run Is Nothing OrElse _step Is Nothing Then Return
        _instructed.Add(_run.Index)
        StartStep()
    End Sub

#End Region

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
        If Not _inInstructions AndAlso Not _step.IsComplete() AndAlso Not Ask.YesNo(FindForm(), "run.stepNotComplete") Then Return
        If _run.Index < _run.Steps.Count - 1 Then
            _run.Index += 1
            BeginInvoke(Sub() ShowStep())
            Return
        End If
        BeginInvoke(Sub() FinishRun())
    End Sub

    ''' <summary>
    ''' The step panels are read at a distance during a test: small text (field labels, notes,
    ''' units) is shown at the normal size, and number boxes a size larger.
    ''' </summary>
    Private Shared Sub LargerText(c As Control)
        If TypeOf c Is Label Then
            If c.Font Is Fonts.Small Then c.Font = Fonts.Body
            If c.Font Is Fonts.SmallStrong Then c.Font = Fonts.Strong
        ElseIf TypeOf c Is NumericUpDown Then
            c.Font = Fonts.Section
        ElseIf TypeOf c Is Slider AndAlso c.Font Is Fonts.Small Then
            c.Font = Fonts.Body                           ' the value shown beside the slider
        End If
        For Each child As Control In c.Controls
            LargerText(child)
        Next
    End Sub

    ''' <summary>Saves the test as it is now (the remaining steps are left out).</summary>
    Public Sub SaveNow()
        If _run Is Nothing OrElse _step Is Nothing Then Return
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
        TabletScreen.SetListening(True)                   ' a tablet: watch it for the patient's taps
        UpdateModeLabel()
        If _step IsNot Nothing Then
            If _inInstructions Then ShowInstructionScene() Else _step.Begin()
        End If
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
        TabletScreen.SetListening(False)
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
        _view.ZoomCaption = Loc.T("run.view.scoring")
        If Not two Then _view.ZoomToText = False
    End Sub

    ''' <summary>
    ''' Two screens only: the centre shows the reading text enlarged to fill it (sharp enough to mark
    ''' the transcript on) instead of an exact, small copy of the patient's screen. With one screen the
    ''' patient reads from the centre panel, so it always stays an exact copy.
    ''' </summary>
    Public Property ScoringView As Boolean
        Get
            Return _view.ZoomToText
        End Get
        Set(value As Boolean)
            _view.ZoomToText = value AndAlso ScreenLayout.IsTwoScreenMode()
        End Set
    End Property

    ''' <summary>Redraws the patient's view.</summary>
    Public Sub RefreshView()
        _view.Invalidate()
        _strip?.Invalidate()
    End Sub

    Public Sub RefreshQuickActions()
        Dim actions = If(_inInstructions, New List(Of QuickAction) From {New QuickAction("instr.start", Sub() BeginAfterInstructions(), primary:=True)},
                         If(_step?.QuickActions(), New List(Of QuickAction)))
        _strip?.SetActions(actions, If(_step Is Nothing, "", Loc.T(_step.TitleKey)))
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveHandler Stimulus.Window.KeyPressedOnScreen, AddressOf KeyFromPatientScreen
            RemoveHandler Stimulus.Window.ClickedOnScreen, AddressOf ClickFromPatientScreen
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

    Private Sub ClickFromPatientScreen(bounds As Rectangle, e As MouseEventArgs)
        If Not Visible OrElse _inInstructions Then Return
        _step?.PatientScreenClick(bounds, e)
    End Sub

    Private Sub KeyFromPatientScreen(keyData As Keys)
        If Not Visible Then Return
        HandleKey(keyData)
    End Sub

    Private Function HandleKey(keyData As Keys) As Boolean
        If keyData = Keys.F11 OrElse (keyData = Keys.Escape AndAlso Stimulus.Window.IsInFront) Then
            If Not ScreenLayout.IsTwoScreenMode() Then ToggleFullScreen()
            Return True
        End If
        If _inInstructions Then
            If keyData = Keys.Space OrElse keyData = Keys.Enter Then BeginAfterInstructions() : Return True
            Return False
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
    ''' <summary>Caption shown while the text is enlarged for scoring.</summary>
    Public Property ZoomCaption As String = ""
    Public Property CursorFor As Func(Of Cursor)

    Private _zoom As Boolean

    ''' <summary>True: a reading text is drawn enlarged to fill the view (for scoring) rather than as an exact copy of the patient's screen.</summary>
    Public Property ZoomToText As Boolean
        Get
            Return _zoom
        End Get
        Set(value As Boolean)
            If _zoom = value Then Return
            _zoom = value
            Invalidate()
        End Set
    End Property

    Private Const CaptionHeight As Integer = 26

    ''' <summary>The part of the view below the caption.</summary>
    Private Function ViewArea() As Rectangle
        Dim top = CaptionHeight
        Return New Rectangle(1, top, Math.Max(10, Width - 2), Math.Max(10, Height - top - 2))
    End Function

    Public Event SceneMouseMove(bounds As Rectangle, p As Point)
    Public Event SceneMouseClick(bounds As Rectangle, e As MouseEventArgs)

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        AddHandler Stimulus.Window.Changed, Sub() Invalidate()
    End Sub

    ''' <summary>Where the patient's screen is drawn (its shape kept), below the caption.</summary>
    ''' <remarks>
    ''' When the text is enlarged for scoring this is the (larger than the view) rectangle the whole
    ''' patient's screen would cover so that the text block fills the view; the scene is drawn into it
    ''' and clipped to the view, and the scene's WordAt works with it unchanged.
    ''' </remarks>
    Public Function ScreenBounds() As Rectangle
        Dim sz = StimulusArt.ScreenSize
        Dim area = ViewArea()
        Dim top = area.Top, availW = area.Width, availH = area.Height
        Dim k = Math.Min(availW / CDbl(sz.Width), availH / CDbl(sz.Height))
        Dim zoomed = ZoomedBounds(area, k)
        If zoomed.HasValue Then Return zoomed.Value
        Dim w = CInt(sz.Width * k), h = CInt(sz.Height * k)
        Return New Rectangle((Width - w) \ 2, top + (availH - h) \ 2, w, h)
    End Function

    ''' <summary>True while the text is actually being drawn enlarged.</summary>
    Public ReadOnly Property IsZoomed As Boolean
        Get
            Dim sz = StimulusArt.ScreenSize
            Dim area = ViewArea()
            Return ZoomedBounds(area, Math.Min(area.Width / CDbl(sz.Width), area.Height / CDbl(sz.Height))).HasValue
        End Get
    End Property

    Private Function ZoomedBounds(area As Rectangle, fitScale As Double) As Rectangle?
        If Not _zoom Then Return Nothing
        Dim rs = TryCast(Scene, ReadingScene)
        If rs Is Nothing OrElse rs.Presentation = ReadingPresentation.TwoOverlays OrElse rs.Presentation = ReadingPresentation.ColourOnly Then Return Nothing
        Dim tb = rs.TextBounds()
        If tb.Width <= 0 OrElse tb.Height <= 0 Then Return Nothing
        Dim kt = Math.Min(area.Width / CDbl(tb.Width), area.Height / CDbl(tb.Height))
        If kt <= fitScale Then Return Nothing             ' the text already fills the view
        Dim sz = StimulusArt.ScreenSize
        Dim cx = area.Left + area.Width / 2.0, cy = area.Top + area.Height / 2.0
        Dim x = cx - (tb.X + tb.Width / 2.0) * kt, y = cy - (tb.Y + tb.Height / 2.0) * kt
        Return New Rectangle(CInt(Math.Round(x)), CInt(Math.Round(y)), CInt(Math.Round(sz.Width * kt)), CInt(Math.Round(sz.Height * kt)))
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        Dim zoomed = IsZoomed
        TextRenderer.DrawText(g, If(zoomed, ZoomCaption, Caption), Fonts.SmallStrong, New Rectangle(0, 0, Width, 22), If(zoomed, p.Primary, p.Muted),
                              TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        Dim r = ScreenBounds()
        Dim frame = If(zoomed, ViewArea(), r)            ' enlarged: r is bigger than the view, so draw into it clipped to the view
        g.SetClip(frame)
        If Stimulus.IsOpen Then
            Stimulus.Window.RenderPreview(g, r)
        Else
            Scene?.Paint(g, r, True)
        End If
        g.ResetClip()
        Using pen As New Pen(If(zoomed, p.Primary, p.Border), 2)
            g.DrawRectangle(pen, frame.X - 1, frame.Y - 1, frame.Width + 1, frame.Height + 1)
        End Using
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim r = ScreenBounds()
        If r.Contains(e.Location) AndAlso (Not IsZoomed OrElse ViewArea().Contains(e.Location)) Then
            RaiseEvent SceneMouseMove(r, e.Location)
            Cursor = If(CursorFor?.Invoke(), Cursors.Default)
        Else
            Cursor = Cursors.Default
        End If
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        Dim r = ScreenBounds()
        If r.Contains(e.Location) AndAlso (Not IsZoomed OrElse ViewArea().Contains(e.Location)) Then RaiseEvent SceneMouseClick(r, e)
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
