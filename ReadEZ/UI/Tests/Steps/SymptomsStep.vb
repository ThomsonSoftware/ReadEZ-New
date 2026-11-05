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
    Private ReadOnly txtNotes As New TextField(multiline:=True) With {.Height = 80}
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

    ''' <summary>A pattern test: the pattern questions (how the stripes look and feel) instead of the text ones.</summary>
    Private ReadOnly Property IsPattern As Boolean
        Get
            Return Run IsNot Nothing AndAlso Run.Result.IsPattern
        End Get
    End Property

    ''' <summary>Question n (1 - 4): "symptoms.q1" ... or, for a pattern, "symptoms.pq1" ...</summary>
    Private Function QuestionKey(n As Integer) As String
        Return If(IsPattern, "symptoms.pq", "symptoms.q") & n
    End Function

    Public Overrides Function CreatePanel() As Control
        Dim col = Ui.Column()
        Ui.FillWidth(col)                                 ' everything below stretches to the panel's width
        Dim intro = Ui.Label(If(IsPattern, "symptoms.pintro", "symptoms.intro"), muted:=True)
        intro.Font = Fonts.Body
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)
        Dim legend = Ui.Label("symptoms.scale", muted:=True)
        legend.Font = Fonts.Small
        legend.Margin = New Padding(0, 0, 0, 6)
        col.Controls.Add(legend)
        For q = 0 To Questions - 1
            Dim index = q
            Dim l = Ui.Label(QuestionKey(q + 1))
            l.Font = Fonts.Strong
            l.Margin = New Padding(0, 8, 0, 4)
            col.Controls.Add(l)
            Dim s As New GradeScale With {.Height = 40, .Grade = Run.Result.Symptoms(q)}
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
        If IsPattern Then
            col.Controls.Add(GratingOptions.Create(Me, Sub() Redraw()))    ' its spatial frequency (asks first)
        Else
            col.Controls.Add(TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=Run.Result.Mode = ReadEZTestMode.Overlay))
        End If
        HighlightCurrent()
        Return col
    End Function

    Public Overrides Function InstructionKeys() As String()
        If IsPattern Then Return {"instr.psymptoms.title", "instr.psymptoms.1", "instr.psymptoms.2", "instr.psymptoms.3"}
        Return {"instr.symptoms.title", "instr.symptoms.1", "instr.symptoms.2", "instr.symptoms.3"}
    End Function

    ''' <summary>The text as it will be shown (white, or with the colour), the patient told to look at it.</summary>
    Public Overrides Function InstructionScreen() As InstructionScene
        Dim sample = SampleText()
        Dim r = Run.Result
        If Not r.IsBaseline AndAlso Run.HasColour Then
            sample.Colour1 = r.ScreenColour
            sample.Presentation = If(r.Mode = ReadEZTestMode.Overlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen)
        End If
        Dim sc As New InstructionScene(InstructionKeys()) With {.Sample = sample}
        sc.Add(If(IsPattern, "instr.psymptoms.callout1", "instr.symptoms.callout1"), sample.CopyBounds(0), InstructionScene.Side.Left)
        sc.Add(If(IsPattern, "instr.psymptoms.callout2", "instr.symptoms.callout2"), RectangleF.Empty, InstructionScene.Side.BottomCentre)
        Return sc
    End Function

    Public Overrides Function CreateFooter() As Control
        Return FooterRow(NextButton())
    End Function

    Private Function IsLastStep() As Boolean
        Return Run.Index = Run.Steps.Count - 1
    End Function

    Private Sub HighlightCurrent()
        For i = 0 To _scales.Count - 1
            _scales(i).Highlighted = i = _current
        Next
        ActionsChanged()                       ' one screen: the bar shows the question being asked
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

    ''' <summary>
    ''' One screen, patient full screen: the bar asks the questions - ◀ / ▶ to move between them,
    ''' the question (1/4 ...), its grade 0 - 5 (the grade given is highlighted), then Next step.
    ''' Choosing a grade moves on to the next question, as on the test screen.
    ''' </summary>
    Public Overrides Function QuickActions() As List(Of QuickAction)
        Dim list As New List(Of QuickAction)
        If _scales.Count = Questions Then
            Dim q = Math.Max(0, Math.Min(Questions - 1, _current))
            list.Add(New QuickAction("", Sub() MoveQuestion(-1), enabled:=q > 0) With {.Text = "◀", .GroupStart = True})
            list.Add(New QuickAction("", Nothing) With {.Text = Loc.F("symptoms.strip.question", q + 1, Questions, Loc.T(QuestionKey(q + 1)))})
            list.Add(New QuickAction("", Sub() MoveQuestion(1), enabled:=q < Questions - 1) With {.Text = "▶"})
            Dim given = If(Run.Result.SymptomsDone, Run.Result.Symptoms(q), -1)
            For g = 0 To 5
                Dim grade = g
                list.Add(New QuickAction("", Sub() _scales(q).SetGrade(grade), primary:=grade = given) With {
                         .Text = g.ToString(), .GroupStart = g = 0, .Colour = GradeScale.GradeColour(g)})
            Next
        End If
        list.Add(New QuickAction(If(IsLastStep(), "run.saveFinish", "run.nextStep"), Sub() Finish(), primary:=True) With {.GroupStart = True})
        Return list
    End Function

    Public Overrides Function CreateTextOptions() As Control
        If IsPattern Then Return GratingOptions.Create(Me, Sub() Redraw())
        Return TextOptions.Create(Me, Sub() Redraw(), showOverlaySize:=Run.Result.Mode = ReadEZTestMode.Overlay)
    End Function

    Private Sub MoveQuestion(delta As Integer)
        _current = Math.Max(0, Math.Min(Questions - 1, _current + delta))
        HighlightCurrent()
    End Sub
End Class

''' <summary>A 0 (none) to 5 (severe) scale: six boxes, green to red.</summary>
Public NotInheritable Class GradeScale
    Inherits Control
    Implements IThemed, IFillWidth

    Private _grade As Integer
    Private _hover As Integer = -1
    Private _highlighted As Boolean
    Public Event GradeChanged As EventHandler

    ''' <summary>ReadEZ 2020's grade colours: 0 white (none) through pale yellow to orange (5, severe).</summary>
    Private Shared ReadOnly Colours As Color() = {
        Color.FromArgb(255, 255, 255), Color.FromArgb(253, 227, 167), Color.FromArgb(251, 201, 122),
        Color.FromArgb(249, 174, 88), Color.FromArgb(246, 142, 61), Color.FromArgb(242, 107, 33)}

    ''' <summary>Text on the grade boxes (dark: every grade colour is light enough).</summary>
    Public Shared ReadOnly BoxInk As Color = Color.FromArgb(40, 40, 40)

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
            ' As ReadEZ 2020: every box in its grade colour, the chosen grade ringed in black.
            ' Hover = a thin blue outline only, so it can't be mistaken for an answer.
            Dim isHover = i = _hover AndAlso Not isOn
            ' The chosen grade: full colour, thick black frame, bold number; the others paler
            Dim penColour = If(isOn, Color.Black, If(isHover OrElse _highlighted, p.Primary, Draw.Mix(Colours(i), Color.Black, 0.15)))
            Dim penWidth = If(isOn, 3.0F, If(isHover, 2.0F, If(_highlighted, 1.5F, 1.0F)))
            Dim fill = If(isOn, Colours(i), Draw.Mix(Colours(i), p.Surface, 0.45))
            Using path = Draw.RoundedRect(If(isOn, RectangleF.Inflate(r, -1, -1), r), 6), b As New SolidBrush(fill),
                  pen As New Pen(penColour, penWidth)
                g.FillPath(b, path)
                g.DrawPath(pen, path)
            End Using
            TextRenderer.DrawText(g, i.ToString(), If(isOn, Fonts.Section, Fonts.Body), Rectangle.Round(r), If(isOn, Color.Black, Color.FromArgb(110, 110, 110)),
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
        Dim run = owner.Run
        Dim res = run.Result
        Dim baseline = run.TextBaseline                 ' colour test: the baseline's settings are used
        Dim card As New CollapsibleCard With {.Width = 320, .FitHeight = True, .Margin = New Padding(0, 14, 0, 0)}
        card.SetTitle("text.options")
        If baseline IsNot Nothing Then card.TitleSuffix = "  ·  " & Loc.F("text.asBaseline", baseline.Test_Date.ToString("d", Loc.Culture))
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        Ui.FillWidth(col)
        Dim cboText As New ThemedComboBox With {.Width = 280, .DropDownStyle = ComboBoxStyle.DropDownList}
        For Each t In ReadingTexts.All()
            cboText.Items.Add(t.Name)
        Next
        cboText.SelectedIndex = Math.Max(0, Math.Min(cboText.Items.Count - 1, res.RoR_TextOptions))
        Dim fnt As New FontButton With {.Width = 280, .SelectedFontName = res.RoR_FontName}
        Dim size As New Slider With {.Minimum = 0, .Maximum = 30, .Height = 30, .ShowValue = True, .Value = res.RoR_FontSize}
        size.ValueFormat = Function(v) (v + 8).ToString() & " pt"
        Dim spacing As New Slider With {.Minimum = 0, .Maximum = 20, .Height = 30, .ShowValue = True, .Value = res.RoR_LineSpacing}
        spacing.ValueFormat = Function(v) (0.8 + v / 10.0).ToString("0.0")
        Dim rows As New Slider With {.Minimum = 1, .Maximum = 20, .Height = 30, .ShowValue = True, .Value = Math.Max(1, Math.Min(20, res.RoR_NoRows))}
        Dim overlay As New Slider With {.Minimum = 0, .Maximum = 10, .Height = 30, .ShowValue = True, .Value = If(res.RoR_OverlaySize > 0, res.RoR_OverlaySize, s.Overlay_Size)}
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
        card.Expanded = False                             ' collapsed at first (after the body is added, so it is hidden)

        ' Colour test: opening the settings warns first - they match the baseline so the two compare
        Dim asking = False
        AddHandler card.ExpandedChanged, Sub()
                                             If Not card.Expanded OrElse baseline Is Nothing OrElse run.TextUnlocked OrElse asking Then Return
                                             asking = True
                                             Try
                                                 If Ask.YesNoWarning(card.FindForm(), "text.changeWarning", baseline.Test_Date.ToString("d", Loc.Culture)) Then
                                                     run.TextUnlocked = True
                                                 Else
                                                     card.Expanded = False
                                                 End If
                                             Finally
                                                 asking = False
                                             End Try
                                         End Sub

        Dim apply = Sub()
                        If enabledWhen IsNot Nothing AndAlso Not enabledWhen() Then Return
                        ' This test
                        res.RoR_TextOptions = Math.Max(0, cboText.SelectedIndex)
                        res.ROR_Text = ReadingTexts.ByIndex(res.RoR_TextOptions).Name
                        res.RoR_FontName = fnt.SelectedFontName
                        res.RoR_FontSize = size.Value
                        res.RoR_LineSpacing = spacing.Value
                        res.RoR_NoRows = rows.Value
                        res.RoR_OverlaySize = overlay.Value
                        s.RoR_Random = If(mixed.Checked, 1, 0)
                        ' The settings for the next tests too - except in a colour test, which keeps to its baseline's
                        If baseline Is Nothing Then
                            s.RoR_TextOptions = res.RoR_TextOptions
                            s.RoR_FontName = res.RoR_FontName
                            s.RoR_FontSize = res.RoR_FontSize
                            s.RoR_LineSpacing = res.RoR_LineSpacing
                            s.RoR_NoRows = res.RoR_NoRows
                        End If
                        s.Overlay_Size = overlay.Value
                        AppHost.SaveSettings()
                        Dim text = ReadingTexts.ByIndex(res.RoR_TextOptions)
                        owner.Scene.Rows = ReadingTexts.RowsFor(text, res.RoR_NoRows, s.RoR_Random = 1)
                        owner.Scene.FontName = res.RoR_FontName
                        owner.Scene.FontSize = res.RoR_FontSize
                        owner.Scene.LineSpacing = res.RoR_LineSpacing
                        owner.Scene.OverlaySize = res.RoR_OverlaySize
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
