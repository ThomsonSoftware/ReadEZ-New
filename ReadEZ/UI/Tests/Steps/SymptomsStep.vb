Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   SymptomsStep.vb - symptoms on white (baseline) or with the chosen colour
'
'   As ReadEZ 2020: the patient looks at the passage of text, and after enough time for any
'   symptoms to develop is asked four questions, each graded 0 (none) to 5 (severe):
'     1. Do the words appear to move, wobble or flicker?
'     2. Do the words look too close together?
'     3. Does the page look too bright or dazzling?
'     4. Does it hurt your eyes to look at the page?
'   plus any other symptoms (notes). The grades are clicked on a 0 - 5 scale (or typed 0 - 5 for
'   the question that is highlighted; Up / Down moves between questions).
' ====================================================================================================

Public NotInheritable Class SymptomsStep
    Inherits TestStep

    Public Const Questions As Integer = 4
    Private ReadOnly _scales As New List(Of GradeScale)
    Private ReadOnly txtNotes As New TextField(multiline:=True) With {.Width = 310, .Height = 80}
    Private _current As Integer

    Public Overrides ReadOnly Property TitleKey As String
        Get
            Return If(Run IsNot Nothing AndAlso Not Run.Result.IsBaseline, "step.symptomsColour", "step.symptoms")
        End Get
    End Property

    Public Overrides ReadOnly Property HelpTopic As String
        Get
            Return "test.symptoms"
        End Get
    End Property

    Public Overrides Function CreatePanel() As Control
        Dim col = Ui.Column()
        Dim intro = Ui.Label("symptoms.intro", muted:=True)
        intro.MaximumSize = New Size(310, 0)
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)
        Dim legend = Ui.Label("symptoms.scale", muted:=True)
        legend.Font = Fonts.Small
        legend.Margin = New Padding(0, 0, 0, 6)
        col.Controls.Add(legend)
        For q = 0 To Questions - 1
            Dim index = q
            Dim l = Ui.Label("symptoms.q" & (q + 1))
            l.Font = Fonts.Strong
            l.MaximumSize = New Size(310, 0)
            l.Margin = New Padding(0, 8, 0, 4)
            col.Controls.Add(l)
            Dim s As New GradeScale With {.Size = New Size(310, 40), .Grade = Run.Result.Symptoms(q)}
            AddHandler s.GradeChanged, Sub()
                                           Run.Result.Symptoms(index) = s.Grade
                                           Run.Result.SymptomsDone = True
                                           _current = Math.Min(Questions - 1, index + 1)
                                           HighlightCurrent()
                                       End Sub
            AddHandler s.Click, Sub()
                                    _current = index
                                    HighlightCurrent()
                                End Sub
            _scales.Add(s)
            col.Controls.Add(s)
        Next
        col.Controls.Add(Ui.Stack("symptoms.other", txtNotes))
        txtNotes.Text = Run.Result.Symptoms_Notes
        AddHandler txtNotes.Box.TextChanged, Sub()
                                                 Run.Result.Symptoms_Notes = txtNotes.Text
                                                 If txtNotes.Text <> "" Then Run.Result.SymptomsDone = True
                                             End Sub
        col.Controls.Add(TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=Run.Result.Mode = ReadEZTestMode.Overlay))
        Dim nextBtn = ModernButton.Create(If(IsLastStep(), "run.saveFinish", "run.nextStep"), ButtonKind.Primary, If(IsLastStep(), Glyphs.Save, Glyphs.Forward))
        nextBtn.Margin = New Padding(0, 14, 0, 0)
        AddHandler nextBtn.Click, Sub() Finish()
        col.Controls.Add(nextBtn)
        HighlightCurrent()
        Return col
    End Function

    Private Function IsLastStep() As Boolean
        Return Run.Index = Run.Steps.Count - 1
    End Function

    Private Sub HighlightCurrent()
        For i = 0 To _scales.Count - 1
            _scales(i).Highlighted = i = _current
        Next
    End Sub

    Public Overrides Sub Begin()
        ApplyTextSettings()
        ApplyTestColour()
        Scene.ShowMarks = False
        MyBase.Begin()
    End Sub

    Public Overrides Function IsComplete() As Boolean
        Return Run.Result.SymptomsDone
    End Function

    Public Overrides Function HandleKey(keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.D0 To Keys.D5, Keys.NumPad0 To Keys.NumPad5
                Dim g = If(keyData >= Keys.NumPad0, keyData - Keys.NumPad0, keyData - Keys.D0)
                If _current >= 0 AndAlso _current < _scales.Count Then _scales(_current).SetGrade(CInt(g))
                Return True
            Case Keys.Up
                _current = Math.Max(0, _current - 1) : HighlightCurrent() : Return True
            Case Keys.Down
                _current = Math.Min(Questions - 1, _current + 1) : HighlightCurrent() : Return True
            Case Keys.Enter
                Finish() : Return True
        End Select
        Return False
    End Function

    Public Overrides Function QuickActions() As List(Of QuickAction)
        Return New List(Of QuickAction) From {New QuickAction(If(IsLastStep(), "run.saveFinish", "run.nextStep"), Sub() Finish(), primary:=True)}
    End Function
End Class

''' <summary>A 0 (none) to 5 (severe) scale: six boxes, green to red.</summary>
Public NotInheritable Class GradeScale
    Inherits Control
    Implements IThemed

    Private _grade As Integer
    Private _hover As Integer = -1
    Private _highlighted As Boolean
    Public Event GradeChanged As EventHandler

    Private Shared ReadOnly Colours As Color() = {
        Color.FromArgb(46, 160, 67), Color.FromArgb(125, 180, 60), Color.FromArgb(200, 190, 40),
        Color.FromArgb(235, 150, 40), Color.FromArgb(225, 95, 45), Color.FromArgb(200, 40, 40)}

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Cursor = Cursors.Hand
        Margin = New Padding(0, 0, 0, 2)
    End Sub

    Public Property Grade As Integer
        Get
            Return _grade
        End Get
        Set(value As Integer)
            _grade = Math.Max(0, Math.Min(5, value))
            Invalidate()
        End Set
    End Property

    Public Property Highlighted As Boolean
        Get
            Return _highlighted
        End Get
        Set(value As Boolean)
            _highlighted = value
            Invalidate()
        End Set
    End Property

    Public Sub SetGrade(g As Integer)
        Grade = g
        RaiseEvent GradeChanged(Me, EventArgs.Empty)
    End Sub

    Public Shared Function GradeColour(g As Integer) As Color
        Return Colours(Math.Max(0, Math.Min(5, g)))
    End Function

    Private Function BoxAt(x As Integer) As Integer
        Return Math.Max(0, Math.Min(5, x * 6 \ Math.Max(1, Width)))
    End Function

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim h = BoxAt(e.X)
        If h <> _hover Then _hover = h : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _hover = -1
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        SetGrade(BoxAt(e.X))
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim w = (Width - 2) / 6.0F
        For i = 0 To 5
            Dim r As New RectangleF(1 + i * w + 2, 2, w - 4, Height - 4)
            Dim isOn = i = _grade
            Using path = Draw.RoundedRect(r, 6), b As New SolidBrush(If(isOn, Colours(i), If(i = _hover, Draw.Mix(p.Surface, Colours(i), 0.25), p.SurfaceAlt))),
                  pen As New Pen(If(isOn, Colours(i), If(_highlighted, p.Primary, p.Border)), If(_highlighted AndAlso Not isOn, 1.5F, 1.0F))
                g.FillPath(b, path)
                g.DrawPath(pen, path)
            End Using
            TextRenderer.DrawText(g, i.ToString(), Fonts.Strong, Rectangle.Round(r), If(isOn, Color.White, p.Ink),
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
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

''' <summary>The text settings shared by the reading steps (they change the settings and the text at once).</summary>
Public NotInheritable Class TextOptions
    Private Sub New()
    End Sub

    Public Shared Function Create(owner As TestStep, changed As Action, showOverlaySize As Boolean, Optional enabledWhen As Func(Of Boolean) = Nothing) As Control
        Dim s = AppHost.Settings
        Dim card As New CollapsibleCard With {.Width = 320, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Margin = New Padding(0, 14, 0, 0)}
        card.SetTitle("text.options")
        card.Expanded = False
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        Dim cboText As New ThemedComboBox With {.Width = 280, .DropDownStyle = ComboBoxStyle.DropDownList}
        For Each t In ReadingTexts.All()
            cboText.Items.Add(t.Name)
        Next
        cboText.SelectedIndex = Math.Max(0, Math.Min(cboText.Items.Count - 1, s.RoR_TextOptions))
        Dim fnt As New FontButton With {.Width = 280, .SelectedFontName = s.RoR_FontName}
        Dim size As New Slider With {.Minimum = 0, .Maximum = 30, .Width = 280, .Height = 30, .ShowValue = True, .Value = s.RoR_FontSize}
        size.ValueFormat = Function(v) (v + 8).ToString() & " pt"
        Dim spacing As New Slider With {.Minimum = 0, .Maximum = 20, .Width = 280, .Height = 30, .ShowValue = True, .Value = s.RoR_LineSpacing}
        spacing.ValueFormat = Function(v) (0.8 + v / 10.0).ToString("0.0")
        Dim rows As New Slider With {.Minimum = 1, .Maximum = 20, .Width = 280, .Height = 30, .ShowValue = True, .Value = s.RoR_NoRows}
        Dim overlay As New Slider With {.Minimum = 0, .Maximum = 10, .Width = 280, .Height = 30, .ShowValue = True, .Value = s.Overlay_Size}
        Dim mixed = ThemedCheckBox.Create("text.mixed")
        mixed.Checked = s.RoR_Random = 1
        col.Controls.Add(Ui.Stack("text.text", cboText))
        col.Controls.Add(Ui.Stack("text.font", fnt))
        col.Controls.Add(Ui.Stack("text.size", size))
        col.Controls.Add(Ui.Stack("text.spacing", spacing))
        col.Controls.Add(Ui.Stack("text.rows", rows))
        col.Controls.Add(mixed)
        If showOverlaySize Then col.Controls.Add(Ui.Stack("text.overlaySize", overlay))
        card.Controls.Add(col)

        Dim apply = Sub()
                        If enabledWhen IsNot Nothing AndAlso Not enabledWhen() Then Return
                        s.RoR_TextOptions = Math.Max(0, cboText.SelectedIndex)
                        s.RoR_FontName = fnt.SelectedFontName
                        s.RoR_FontSize = size.Value
                        s.RoR_LineSpacing = spacing.Value
                        s.RoR_NoRows = rows.Value
                        s.RoR_Random = If(mixed.Checked, 1, 0)
                        s.Overlay_Size = overlay.Value
                        AppHost.SaveSettings()
                        Dim r = owner.Run.Result
                        r.RoR_TextOptions = s.RoR_TextOptions
                        r.ROR_Text = ReadingTexts.ByIndex(s.RoR_TextOptions).Name
                        r.RoR_FontName = s.RoR_FontName : r.RoR_FontSize = s.RoR_FontSize : r.RoR_LineSpacing = s.RoR_LineSpacing
                        r.RoR_NoRows = s.RoR_NoRows : r.RoR_OverlaySize = s.Overlay_Size
                        Dim text = ReadingTexts.ByIndex(s.RoR_TextOptions)
                        owner.Scene.Rows = ReadingTexts.RowsFor(text, s.RoR_NoRows, s.RoR_Random = 1)
                        owner.Scene.FontName = s.RoR_FontName
                        owner.Scene.FontSize = s.RoR_FontSize
                        owner.Scene.LineSpacing = s.RoR_LineSpacing
                        owner.Scene.OverlaySize = s.Overlay_Size
                        changed()
                    End Sub
        AddHandler cboText.SelectedIndexChanged, Sub() apply()
        AddHandler fnt.SelectedFontChanged, Sub() apply()
        AddHandler size.ValueChanged, Sub() apply()
        AddHandler spacing.ValueChanged, Sub() apply()
        AddHandler rows.ValueChanged, Sub() apply()
        AddHandler overlay.ValueChanged, Sub() apply()
        AddHandler mixed.CheckedChanged, Sub() apply()
        Return card
    End Function
End Class
