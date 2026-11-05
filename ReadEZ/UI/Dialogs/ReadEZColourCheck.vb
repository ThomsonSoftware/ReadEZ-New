Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   ReadEZColourCheck.vb - checks the ReadEZ colours on the patient's screen
'
'   (ReadEZ: Colour Calibration, where the colour was matched with red / green / blue sliders.)
'   The check fills the patient's screen, because the match only holds for the screen it is made on.
'     Overlays  the overlay is laid on white paper and held, in the room's normal light, against
'               the right-hand edge of the screen, beside a patch of the screen colour; the screen
'               colour should look the same as the overlay.
'     Lenses    the lens is held over a white window 50 mm across; the screen's version of the lens
'               colour is shown beside the window (compare) or in a ring of circles round it (find).
'
'     1. Compare  the stored colour. If it matches: "Looks the same" (the colour is marked as
'                 checked on this screen, with the date) and the next colour that hasn't been
'                 checked comes up.
'     2. Find     if not: eight strips (overlays) or circles (lenses) in a shuffled order - the
'                 present colour (twice) and six small changes of it (lighter, darker, more colour,
'                 less colour, and a hue shift either way). Click the one that matches the overlay
'                 or lens best. Choosing the present colour halves the size of the changes; choosing a
'                 change moves to it. After a few choices the changes are too small to see and the
'                 match is shown for a final compare. The changes are equal steps to the eye
'                 (OKLab), so every choice is about equally hard and no slider work is needed.
'
'   Sliders for hue, amount of colour and lightness (OKLCh, as the bespoke colour) adjust the colour
'   directly, the comparison following as they move; "ReadEZ default" goes back to the palette's
'   own colour. Nothing is stored until Save. Keys: 1 - 8 choose a
'   strip, Enter = looks the same, Backspace = undo, Esc = cancel.
'   Sizes on the patient's screen are in millimetres (Settings > Screens), so the strips are the
'   same size whatever the screen.
' ====================================================================================================

Public NotInheritable Class ReadEZColourCheck
    Inherits Form

    ''' <summary>Opens the check on the patient's screen, starting at a colour code (or the first one not checked).</summary>
    Public Shared Sub Run(owner As IWin32Window, startCode As String)
        If Not ScreenLayout.HasStimulusScreen Then
            If Not Ask.YesNo(owner, "readez.check.oneScreen") Then Return
        End If
        ReadEZColours.ForgetScreen()
        Using f As New ReadEZColourCheck(startCode)
            f.ShowDialog(owner)
        End Using
    End Sub

    Private Enum Mode
        Compare
        Find
    End Enum

    Private Const Candidates As Integer = 8
    Private Const StartStep As Double = 0.03         ' OKLab; about 1.5 just-noticeable differences
    Private Const EndStep As Double = 0.006          ' below about a third of a just-noticeable difference
    Private Const NudgeStep As Double = 0.008

    Private ReadOnly _colours As Color()()
    Private ReadOnly _checked As Boolean()()
    Private _set As ReadEZSet = ReadEZSet.Overlay
    Private _index As Integer = 1
    Private _mode As Mode = Mode.Compare
    Private _justFound As Boolean

    ' The search
    Private _cur As OkLab
    Private _step As Double
    Private _round As Integer
    Private ReadOnly _options As New List(Of (Lab As OkLab, IsCurrent As Boolean))
    Private ReadOnly _undo As New Stack(Of (Lab As OkLab, StepSize As Double, Round As Integer))
    Private ReadOnly _rnd As New Random()
    Private _hoverOption As Integer = -1
    Private _optionRects As New List(Of RectangleF)
    Private _optionRound As Boolean             ' lens candidates are circles
    Private _dirty As Boolean
    Private _syncing As Boolean

    ' Controls
    Private ReadOnly _side As New Panel
    Private ReadOnly _view As New Canvas
    Private ReadOnly segSet As New SegmentedControl
    Private ReadOnly _swatches As New Canvas
    Private ReadOnly lblName As New Label
    Private ReadOnly lblState As New Label
    Private ReadOnly lblHelp As New Label
    Private ReadOnly lblIntro As New Label
    Private ReadOnly btnSame As ModernButton
    Private ReadOnly btnFind As ModernButton
    Private ReadOnly btnUndo As ModernButton
    Private ReadOnly btnStop As ModernButton
    Private ReadOnly btnDefault As ModernButton
    Private ReadOnly btnSave As ModernButton
    Private ReadOnly btnCancel As ModernButton
    Private ReadOnly _nudges As New List(Of ModernButton)
    Private ReadOnly _sliders As New OkLchSliders(340)
    Private ReadOnly _diagram As New CheckDiagram
    Private _fromSliders As Boolean

    Private Shared ReadOnly Surround As Color = Color.FromArgb(38, 38, 38)
    Private Shared ReadOnly SurroundText As Color = Color.FromArgb(185, 185, 185)

    Private Sub New(startCode As String)
        ' Working copy of the colours
        _colours = New Color(2)() {}
        _checked = New Boolean(2)() {}
        For Each s In ReadEZColours.Sets
            _colours(CInt(s)) = New Color(ReadEZColours.PerSet - 1) {}
            _checked(CInt(s)) = New Boolean(ReadEZColours.PerSet - 1) {}
            For i = 0 To ReadEZColours.PerSet - 1
                _colours(CInt(s))(i) = ReadEZColours.Colour(s, i)
            Next
        Next

        SuspendLayout()
        AutoScaleMode = AutoScaleMode.None
        FormBorderStyle = FormBorderStyle.None
        StartPosition = FormStartPosition.Manual
        ShowInTaskbar = False
        TopMost = True
        KeyPreview = True
        Font = Fonts.Body
        Icon = AppHost.AppIcon
        Text = Loc.T("readez.check.title")
        Bounds = PatientScreen()
        ThemeManager.SetRole(Me, SurfaceRole.Surface)

        ' ---- Controls (left) ----
        _side.Dock = DockStyle.Left
        _side.Width = 400
        _side.AutoScroll = True
        _side.Padding = New Padding(24, 20, 20, 20)
        ThemeManager.SetRole(_side, SurfaceRole.Surface)
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top                ' inside the panel's padding
        col.MaximumSize = New Size(350, 0)
        col.Controls.Add(Ui.Title("readez.check.title"))
        ' How to hold the overlay (or lens): a small picture
        _diagram.Size = New Size(340, 150)
        _diagram.Margin = New Padding(0, 4, 0, 6)
        col.Controls.Add(_diagram)
        Dim screenLbl = Ui.Label("", muted:=True)
        screenLbl.Text = Loc.F("readez.check.screen", ReadEZColours.ScreenName)
        screenLbl.MaximumSize = New Size(350, 0)
        col.Controls.Add(screenLbl)
        lblIntro.AutoSize = True
        lblIntro.Font = Fonts.Body
        lblIntro.MaximumSize = New Size(350, 0)
        lblIntro.Margin = New Padding(0, 8, 0, 12)
        col.Controls.Add(lblIntro)

        segSet.SetOptions("readez.overlays", "readez.lightTint", "readez.darkTint")
        col.Controls.Add(segSet)
        _swatches.Size = New Size(13 * 26, 48)
        _swatches.Cursor = Cursors.Hand
        _swatches.BackColor = Color.Transparent
        _swatches.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(_swatches)

        For Each l In {lblName, lblState, lblHelp, lblIntro}
            LabelStyle.SetCustom(l)          ' coloured in UpdateUi
        Next
        lblName.AutoSize = True
        lblName.Font = Fonts.Section
        lblName.Margin = New Padding(0, 4, 0, 2)
        col.Controls.Add(lblName)
        lblState.AutoSize = True
        lblState.Font = Fonts.Small
        lblState.MaximumSize = New Size(350, 0)
        lblState.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(lblState)
        lblHelp.AutoSize = True
        lblHelp.MaximumSize = New Size(350, 0)
        lblHelp.Margin = New Padding(0, 0, 0, 12)
        col.Controls.Add(lblHelp)

        btnSame = ModernButton.Create("readez.check.same", ButtonKind.Primary, Glyphs.Accept)
        btnFind = ModernButton.Create("readez.check.find", ButtonKind.Secondary, Glyphs.Search)
        btnUndo = ModernButton.Create("readez.check.undo", ButtonKind.Secondary, Glyphs.Back)
        btnStop = ModernButton.Create("readez.check.stop", ButtonKind.Ghost, Glyphs.Cancel)
        Dim main = Ui.Row(wrap:=True)
        main.MaximumSize = New Size(360, 0)
        For Each b In {btnSame, btnFind, btnUndo, btnStop}
            b.Margin = New Padding(0, 0, 8, 8)
            main.Controls.Add(b)
        Next
        col.Controls.Add(main)

        ' Direct adjustment
        col.Controls.Add(Ui.Label("readez.check.adjust", muted:=True))
        _sliders.Margin = New Padding(0, 4, 0, 0)
        AddHandler _sliders.ColourChanged, Sub() SlidersMoved()
        col.Controls.Add(_sliders)
        btnDefault = ModernButton.Create("readez.check.default", ButtonKind.Ghost, Glyphs.Refresh)
        btnDefault.Margin = New Padding(0, 2, 0, 20)
        col.Controls.Add(btnDefault)

        btnSave = ModernButton.Create("readez.check.save", ButtonKind.Primary, Glyphs.Save)
        btnCancel = ModernButton.Create("common.cancel", ButtonKind.Secondary)
        Dim foot = Ui.Row()
        btnSave.Margin = New Padding(0, 0, 8, 0)
        foot.Controls.Add(btnSave)
        foot.Controls.Add(btnCancel)
        col.Controls.Add(foot)
        _side.Controls.Add(col)

        ' ---- The comparison (right) ----
        _view.Dock = DockStyle.Fill
        _view.BackColor = Surround
        ThemeManager.SetRole(_view, SurfaceRole.Stimulus)
        Controls.Add(_view)
        Controls.Add(_side)
        ResumeLayout(False)

        AddHandler segSet.SelectedIndexChanged, Sub() If Not _syncing Then SelectColour(ReadEZColours.Sets(segSet.SelectedIndex), FirstToCheck(ReadEZColours.Sets(segSet.SelectedIndex)))
        AddHandler _swatches.Paint, AddressOf PaintSwatches
        AddHandler _swatches.MouseClick, Sub(s, e)
                                             Dim i = e.X \ 26
                                             If i >= 0 AndAlso i < ReadEZColours.PerSet Then SelectColour(_set, i)
                                         End Sub
        AddHandler _view.Paint, AddressOf PaintView
        AddHandler _view.MouseMove, Sub(s, e)
                                        Dim h = OptionAt(e.Location)
                                        If h <> _hoverOption Then _hoverOption = h : _view.Invalidate()
                                        _view.Cursor = If(h >= 0, Cursors.Hand, Cursors.Default)
                                    End Sub
        AddHandler _view.MouseClick, Sub(s, e)
                                         Dim h = OptionAt(e.Location)
                                         If h >= 0 Then Choose(h)
                                     End Sub
        AddHandler btnSame.Click, Sub() LooksTheSame()
        AddHandler btnFind.Click, Sub() StartFind()
        AddHandler btnUndo.Click, Sub() Undo()
        AddHandler btnStop.Click, Sub() StopFind()
        AddHandler btnDefault.Click, Sub()
                                         SetColour(ReadEZColours.DefaultColour(_set, _index))
                                         If _mode = Mode.Find Then StartFind()
                                     End Sub
        AddHandler btnSave.Click, Sub()
                                      ReadEZColours.Store(_colours, _checked)
                                      DialogResult = DialogResult.OK
                                  End Sub
        AddHandler btnCancel.Click, Sub() Close()

        ' Start at the chosen colour, or the first one not checked yet
        Dim s0 As ReadEZSet, i0 As Integer
        If ReadEZColours.TryParse(startCode, s0, i0) Then
            _set = s0 : _index = i0
        Else
            _set = ReadEZSet.Overlay : _index = FirstToCheck(ReadEZSet.Overlay)
        End If
        SelectColour(_set, _index)
    End Sub

    Private Shared Function PatientScreen() As Rectangle
        Dim r = AppHost.Settings.StimulusBounds
        If Not ScreenLayout.IsOnScreen(r) Then r = ScreenLayout.PrimaryLeadingScreens().Last().Bounds
        Return r
    End Function

    Protected Overrides Sub OnLoad(e As EventArgs)
        Dpi.ScaleOnce(Me)
        ThemeManager.Apply(Me)
        _view.BackColor = Surround
        Bounds = PatientScreen()
        MyBase.OnLoad(e)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If DialogResult <> DialogResult.OK AndAlso _dirty AndAlso Not Ask.YesNo(Me, "readez.check.discard") Then
            e.Cancel = True
            Return
        End If
        MyBase.OnFormClosing(e)
    End Sub

#Region "Choosing a colour to check"

    ''' <summary>The first colour of a set not checked on this screen (A first, grey last).</summary>
    Private Function FirstToCheck(s As ReadEZSet) As Integer
        For Each i In Enumerable.Range(1, ReadEZColours.PerSet - 1).Concat({0})
            If Not IsChecked(s, i) Then Return i
        Next
        Return 1
    End Function

    Private Function IsChecked(s As ReadEZSet, i As Integer) As Boolean
        Return _checked(CInt(s))(i) OrElse (ReadEZColours.CheckedOn(s, i).HasValue AndAlso
                                            _colours(CInt(s))(i).ToArgb() = ReadEZColours.Colour(s, i).ToArgb())
    End Function

    Private Sub SelectColour(s As ReadEZSet, i As Integer)
        _set = s
        _index = i
        If segSet.SelectedIndex <> CInt(s) Then
            _syncing = True
            segSet.SelectedIndex = CInt(s)
            _syncing = False
        End If
        _mode = Mode.Compare
        _justFound = False
        _undo.Clear()
        UpdateUi()
    End Sub

    ''' <summary>Marks the colour as matching and moves on to the next one not yet checked.</summary>
    Private Sub LooksTheSame()
        _checked(CInt(_set))(_index) = True
        _dirty = True
        ' Next one in this set, then the other sets
        For Each s In {_set}.Concat(ReadEZColours.Sets.Where(Function(x) x <> _set))
            For Each i In Enumerable.Range(1, ReadEZColours.PerSet - 1).Concat({0})
                If Not IsChecked(s, i) Then
                    SelectColour(s, i)
                    Return
                End If
            Next
        Next
        UpdateUi()   ' all checked
    End Sub

    Private Sub SetColour(c As Color)
        If _colours(CInt(_set))(_index).ToArgb() = c.ToArgb() Then Return
        _colours(CInt(_set))(_index) = c
        _checked(CInt(_set))(_index) = False
        _dirty = True
        UpdateUi()
    End Sub

#End Region

#Region "Finding the match"

    Private Sub StartFind()
        _cur = OkLab.FromColor(_colours(CInt(_set))(_index))
        _step = StartStep
        _round = 0
        _undo.Clear()
        _mode = Mode.Find
        _justFound = False
        NewRound()
    End Sub

    Private Sub StopFind()
        _mode = Mode.Compare
        UpdateUi()
    End Sub

    ''' <summary>The present colour twice and six changes of it, shuffled.</summary>
    Private Sub NewRound()
        _round += 1
        _options.Clear()
        _options.Add((_cur, True))
        _options.Add((_cur, True))
        ' Directions: lightness, colourfulness (away from grey) and hue (around grey)
        Dim c = _cur.Chroma
        Dim ua = If(c > 0.01, _cur.A / c, 1.0), ub = If(c > 0.01, _cur.B / c, 0.0)
        Dim va = -ub, vb = ua
        Dim d = _step
        _options.Add((New OkLab(_cur.L + d, _cur.A, _cur.B), False))
        _options.Add((New OkLab(_cur.L - d, _cur.A, _cur.B), False))
        _options.Add((New OkLab(_cur.L, _cur.A + d * ua, _cur.B + d * ub), False))
        _options.Add((New OkLab(_cur.L, Math.Max(0, c - d) * ua, Math.Max(0, c - d) * ub), False))
        _options.Add((New OkLab(_cur.L, _cur.A + d * va, _cur.B + d * vb), False))
        _options.Add((New OkLab(_cur.L, _cur.A - d * va, _cur.B - d * vb), False))
        ' Shuffle
        For i = _options.Count - 1 To 1 Step -1
            Dim j = _rnd.Next(i + 1)
            Dim t = _options(i) : _options(i) = _options(j) : _options(j) = t
        Next
        _hoverOption = -1
        UpdateUi()
    End Sub

    Private Sub Choose(k As Integer)
        If _mode <> Mode.Find OrElse k < 0 OrElse k >= _options.Count Then Return
        _undo.Push((_cur, _step, _round))
        Dim o = _options(k)
        If o.IsCurrent Then
            _step /= 2
        Else
            ' Move to it (as the screen will show it)
            _cur = OkLab.FromColor(o.Lab.ToColor())
        End If
        SetColour(_cur.ToColor())
        If _step < EndStep Then
            ' Found: compare once more
            _mode = Mode.Compare
            _justFound = True
            UpdateUi()
        Else
            NewRound()
        End If
    End Sub

    Private Sub Undo()
        If _undo.Count = 0 Then Return
        Dim u = _undo.Pop()
        _cur = u.Lab
        _step = u.StepSize
        _round = u.Round - 1
        SetColour(_cur.ToColor())
        _mode = Mode.Find
        NewRound()
    End Sub

    ''' <summary>The sliders moved: the colour follows (and, while finding, the choices are made round it).</summary>
    Private Sub SlidersMoved()
        _fromSliders = True
        Try
            SetColour(_sliders.Colour)
        Finally
            _fromSliders = False
        End Try
        If _mode = Mode.Find Then
            _cur = OkLab.FromColor(_colours(CInt(_set))(_index))
            NewRound()
        End If
    End Sub

    ''' <summary>Lighter / darker (axis 0), more / less colour (1), hue (2).</summary>
    Private Sub Nudge(axis As Integer, sign As Integer)
        Dim lab = OkLab.FromColor(_colours(CInt(_set))(_index))
        Dim c = lab.Chroma
        Dim ua = If(c > 0.01, lab.A / c, 1.0), ub = If(c > 0.01, lab.B / c, 0.0)
        Select Case axis
            Case 0 : lab.L += sign * NudgeStep
            Case 1
                Dim nc = Math.Max(0, c + sign * NudgeStep)
                lab.A = nc * ua : lab.B = nc * ub
            Case 2
                lab.A += sign * NudgeStep * -ub : lab.B += sign * NudgeStep * ua
        End Select
        SetColour(lab.ToColor())
        If _mode = Mode.Find Then
            _cur = OkLab.FromColor(_colours(CInt(_set))(_index))
            NewRound()
        End If
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.D1 To Keys.D8
                Choose(CInt(keyData) - CInt(Keys.D1)) : Return True
            Case Keys.NumPad1 To Keys.NumPad8
                Choose(CInt(keyData) - CInt(Keys.NumPad1)) : Return True
            Case Keys.Enter
                If _mode = Mode.Compare Then LooksTheSame()
                Return True
            Case Keys.Back
                Undo() : Return True
            Case Keys.Escape
                If _mode = Mode.Find Then StopFind() Else Close()
                Return True
        End Select
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

#End Region

#Region "Display"

    Private Sub UpdateUi()
        lblName.Text = ReadEZColours.Name(_set, _index)
        Dim c = _colours(CInt(_set))(_index)
        If Not _fromSliders AndAlso _sliders.Colour.ToArgb() <> c.ToArgb() Then _sliders.Colour = c   ' (the sliders show the colour)
        Dim def = ReadEZColours.DefaultColour(_set, _index)
        Dim diff = OkLab.Distance(OkLab.FromColor(c), OkLab.FromColor(def))
        Dim state As String
        If _checked(CInt(_set))(_index) Then
            state = Loc.T("readez.check.stateMatched")
        Else
            Dim whenDone = ReadEZColours.CheckedOn(_set, _index)
            If whenDone.HasValue AndAlso c.ToArgb() = ReadEZColours.Colour(_set, _index).ToArgb() Then
                state = Loc.F("readez.checkedOn", whenDone.Value.ToShortDateString())
            ElseIf ReadEZColours.CheckedElsewhere(_set, _index) Then
                state = Loc.T("readez.check.stateOtherScreen")
            Else
                state = Loc.T("readez.notChecked")
            End If
        End If
        state &= "   ·   " & $"RGB {c.R}, {c.G}, {c.B}"
        ' Far from ReadEZ's own colour: probably the wrong filter, or a screen setting (night light ...)
        If diff > 0.08 Then state &= vbCrLf & Loc.T("readez.check.farFromDefault")
        lblState.Text = state
        lblState.ForeColor = If(diff > 0.08, ThemeManager.Current.Warning, ThemeManager.Current.Muted)

        Dim allDone = ReadEZColours.Sets.All(Function(s) Enumerable.Range(0, ReadEZColours.PerSet).All(Function(i) IsChecked(s, i)))
        If _mode = Mode.Find Then
            lblHelp.Text = Loc.F("readez.check.helpFind", _round)
        ElseIf _justFound Then
            lblHelp.Text = Loc.T("readez.check.helpFound")
        ElseIf allDone Then
            lblHelp.Text = Loc.T("readez.check.helpAllDone")
        Else
            lblHelp.Text = Loc.T("readez.check.helpCompare")
        End If
        lblHelp.ForeColor = ThemeManager.Current.Ink
        ' How to hold the filter: an overlay on paper beside the screen, a lens over the window
        lblIntro.Text = Loc.T(If(_set = ReadEZSet.Overlay, "readez.check.introOverlay", "readez.check.introLens"))
        lblIntro.ForeColor = ThemeManager.Current.Ink
        lblName.ForeColor = ThemeManager.Current.Ink
        _diagram.ForLens = _set <> ReadEZSet.Overlay
        _diagram.Tint = c
        _diagram.Invalidate()

        btnSame.Visible = _mode = Mode.Compare
        btnFind.Visible = _mode = Mode.Compare
        btnUndo.Visible = _mode = Mode.Find
        btnUndo.Enabled = _undo.Count > 0
        btnStop.Visible = _mode = Mode.Find
        _swatches.Invalidate()
        _view.Invalidate()
    End Sub

    Private Sub PaintSwatches(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.PixelOffsetMode = PixelOffsetMode.HighQuality
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        Dim p = ThemeManager.Current
        Using f As New Font("Segoe UI", 7.5F), tb As New SolidBrush(p.Muted), sel As New Pen(p.Primary, 3),
              tick As New Pen(p.Success, 2), edge As New Pen(Color.FromArgb(110, p.Ink), 1)
            Dim sf As New StringFormat With {.Alignment = StringAlignment.Center}
            For i = 0 To ReadEZColours.PerSet - 1
                Dim r As New RectangleF(i * 26 + 3, 3, 20, 20)
                Using b As New SolidBrush(_colours(CInt(_set))(i))
                    g.FillEllipse(b, r)
                End Using
                g.DrawEllipse(edge, r)
                If i = _index Then g.DrawEllipse(sel, RectangleF.Inflate(r, 2.5F, 2.5F))
                g.DrawString(If(i = 0, "–", ReadEZColours.Letter(i)), f, tb, New RectangleF(i * 26, 25, 26, 14), sf)
                If IsChecked(_set, i) Then
                    ' Small tick under the letter
                    Dim x = i * 26 + 13.0F, y = 42.0F
                    g.DrawLines(tick, {New PointF(x - 4, y - 1), New PointF(x - 1, y + 2), New PointF(x + 4, y - 3)})
                End If
            Next
        End Using
    End Sub

    ''' <summary>Pixels per millimetre on the patient's screen (Settings > Screens).</summary>
    Private Function PxPerMm() As Single
        Dim s = AppHost.Settings
        Return If(s.Screen_Width > 0 AndAlso s.Stimulus_Width > 0, s.Stimulus_Width / CSng(s.Screen_Width), 96 / 25.4F)
    End Function

    Private Function OptionAt(p As Point) As Integer
        If _mode <> Mode.Find Then Return -1
        For k = 0 To _optionRects.Count - 1
            Dim r = _optionRects(k)
            If _optionRound Then
                Dim dx = p.X - (r.X + r.Width / 2), dy = p.Y - (r.Y + r.Height / 2)
                If dx * dx + dy * dy <= r.Width * r.Width / 4 Then Return k
            ElseIf r.Contains(p) Then
                Return k
            End If
        Next
        Return -1
    End Function

    Private Sub PaintView(sender As Object, e As PaintEventArgs)
        Dim g = e.Graphics
        g.Clear(Surround)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        Dim w = _view.ClientSize.Width, h = _view.ClientSize.Height
        If w < 50 OrElse h < 50 Then Return
        _optionRects = New List(Of RectangleF)
        _optionRound = _set <> ReadEZSet.Overlay
        Dim mm = PxPerMm()
        Dim current = _colours(CInt(_set))(_index)

        Using tb As New SolidBrush(SurroundText), f As New Font("Segoe UI", 13.0F), fb As New Font("Segoe UI Semibold", 15.0F),
              hl As New Pen(Color.White, 3)
            Dim centreBottom As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Far}
            Dim progressTop As Single

            If _set = ReadEZSet.Overlay Then
                ' ---- Overlay: the colour against the right-hand edge of the screen, where the overlay
                ' (on white paper, in the room's light) is held beside it. 80 x 128 mm. ----
                Dim pw = 80 * mm, ph = 128 * mm
                Dim k = Math.Min(1.0F, Math.Min(w * 0.45F / pw, h * 0.7F / ph))
                pw *= k : ph *= k
                Dim patch As New RectangleF(w - pw, (h - ph) / 2, pw, ph)
                If _mode = Mode.Compare Then
                    Using b As New SolidBrush(current)
                        g.FillRectangle(b, patch)
                    End Using
                Else
                    ' Eight strips, each touching the edge of the screen (2 px apart)
                    Dim pitch = ph / Candidates
                    For i = 0 To Candidates - 1
                        Dim r As New RectangleF(patch.X, patch.Y + i * pitch, pw, pitch - If(i < Candidates - 1, 2, 0))
                        _optionRects.Add(r)
                        Using b As New SolidBrush(_options(i).Lab.ToColor())
                            g.FillRectangle(b, r)
                        End Using
                        Dim num = (i + 1).ToString()
                        Dim ns = g.MeasureString(num, fb)
                        g.DrawString(num, fb, tb, r.X - ns.Width - 14, r.Y + (r.Height - ns.Height) / 2)
                        If i = _hoverOption Then
                            ' Outlined on the three sides away from the screen edge
                            g.DrawLines(hl, {New PointF(r.Right, r.Y), New PointF(r.X, r.Y), New PointF(r.X, r.Bottom), New PointF(r.Right, r.Bottom)})
                        End If
                    Next
                End If
                ' Arrow and words pointing to the edge
                Dim textBox As New RectangleF(patch.X - 330, patch.Y - 110, 320 + pw, 96)
                g.DrawString(Loc.T(If(_mode = Mode.Compare, "readez.check.paperHere", "readez.check.paperPick")), f, tb, textBox,
                             New StringFormat With {.Alignment = StringAlignment.Far, .LineAlignment = StringAlignment.Far})
                progressTop = patch.Bottom + 30
                DrawProgress(g, f, tb, New RectangleF(patch.X - 40, progressTop, pw + 30, 8))
            Else
                ' ---- Lens: a white window 50 mm across, held behind the lens; the screen's version of the
                ' lens colour beside it (compare) or round it (find). ----
                Dim win = 50 * mm
                Dim cand = 34 * mm                       ' candidates (find)
                Dim ringR = win / 2 + 8 * mm + cand / 2  ' centre of the window to the centre of a candidate
                Dim span = If(_mode = Mode.Compare, 2 * win + 14 * mm, 2 * (ringR + cand / 2) + 20 * mm)
                Dim k = Math.Min(1.0F, Math.Min(w * 0.85F, h * 0.72F) / span)
                win *= k : cand *= k : ringR *= k
                Dim cx = w / 2.0F, cy = h / 2.0F
                Dim window As RectangleF
                If _mode = Mode.Compare Then
                    Dim gap = 14 * mm * k
                    window = New RectangleF(cx - gap / 2 - win, cy - win / 2, win, win)
                    Dim sim As New RectangleF(cx + gap / 2, cy - win / 2, win, win)
                    g.FillEllipse(Brushes.White, window)
                    Using b As New SolidBrush(current)
                        g.FillEllipse(b, sim)
                    End Using
                    g.DrawString(Loc.T("readez.check.lensWindow"), f, tb, New RectangleF(window.X - 80, window.Y - 100, win + 160, 88), centreBottom)
                    g.DrawString(ReadEZColours.Name(_set, _index), f, tb, New RectangleF(sim.X - 80, sim.Y - 100, win + 160, 88), centreBottom)
                    progressTop = window.Bottom + 30
                Else
                    window = New RectangleF(cx - win / 2, cy - win / 2, win, win)
                    g.FillEllipse(Brushes.White, window)
                    For i = 0 To Candidates - 1
                        Dim a = -Math.PI / 2 + i * 2 * Math.PI / Candidates
                        Dim px = cx + CSng(ringR * Math.Cos(a)), py = cy + CSng(ringR * Math.Sin(a))
                        Dim r As New RectangleF(px - cand / 2, py - cand / 2, cand, cand)
                        _optionRects.Add(r)
                        Using b As New SolidBrush(_options(i).Lab.ToColor())
                            g.FillEllipse(b, r)
                        End Using
                        If i = _hoverOption Then g.DrawEllipse(hl, RectangleF.Inflate(r, 2, 2))
                        ' Number just outside the circle
                        Dim nr = ringR + cand / 2 + 16
                        Dim num = (i + 1).ToString()
                        Dim ns = g.MeasureString(num, fb)
                        g.DrawString(num, fb, tb, cx + CSng(nr * Math.Cos(a)) - ns.Width / 2, cy + CSng(nr * Math.Sin(a)) - ns.Height / 2)
                    Next
                    g.DrawString(Loc.T("readez.check.lensPick"), f, tb,
                                 New RectangleF(cx - 300, cy - ringR - cand / 2 - 140, 600, 96), centreBottom)
                    progressTop = cy + ringR + cand / 2 + 50
                End If
                DrawProgress(g, f, tb, New RectangleF(cx - 150, progressTop, 300, 8))
            End If

            ' Reminder about the screen and the light
            g.DrawString(Loc.T(If(_set = ReadEZSet.Overlay, "readez.check.lightTip", "readez.check.screenTip")), Fonts.Small, tb,
                         New RectangleF(20, h - 44, w - 40, 40),
                         New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Far})
        End Using
    End Sub

    ''' <summary>While finding: how far the search has got (the changes shrink from the start step to the end step).</summary>
    Private Sub DrawProgress(g As Graphics, f As Font, tb As Brush, bar As RectangleF)
        If _mode <> Mode.Find Then Return
        Dim done = CSng(Math.Min(1, Math.Max(0, Math.Log(StartStep / _step) / Math.Log(StartStep / EndStep))))
        Using back As New SolidBrush(Color.FromArgb(70, 70, 70)), fill As New SolidBrush(Color.FromArgb(150, 200, 255))
            g.FillRectangle(back, bar)
            g.FillRectangle(fill, bar.X, bar.Y, bar.Width * done, bar.Height)
        End Using
        g.DrawString(Loc.T("readez.check.progress"), f, tb, New RectangleF(bar.X - 100, bar.Bottom + 6, bar.Width + 200, 30),
                     New StringFormat With {.Alignment = StringAlignment.Center})
    End Sub

#End Region
End Class

''' <summary>
''' The colour check's little picture of how to compare: a monitor showing the screen colour at its
''' right-hand edge, with the overlay on a white card held beside it under a lamp (room light) - or,
''' for lenses, the lens held over the white window in the middle of the screen.
''' Drawn in a 340 x 150 box, scaled to the control.
''' </summary>
Public NotInheritable Class CheckDiagram
    Inherits Control
    Implements IThemed

    ''' <summary>The colour being checked (the screen's patch and the filter's tint).</summary>
    Public Property Tint As Color = Color.FromArgb(255, 219, 122)
    Public Property ForLens As Boolean

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        Dim k = Math.Min(Width / 340.0F, Height / 150.0F)
        Dim state = g.Save()
        g.ScaleTransform(k, k)
        Try
            ' The monitor: bezel, dark surround (as the check), stand
            Dim mon As New RectangleF(10, 18, 200, 112)
            Using bezel As New SolidBrush(Color.FromArgb(45, 48, 54)), stand As New SolidBrush(Color.FromArgb(150, 155, 162))
                g.FillRectangle(stand, mon.X + mon.Width / 2 - 9, mon.Bottom, 18, 10)
                Using basePath = Draw.RoundedRect(New RectangleF(mon.X + mon.Width / 2 - 34, mon.Bottom + 9, 68, 6), 3)
                    g.FillPath(stand, basePath)
                End Using
                Using path = Draw.RoundedRect(mon, 6)
                    g.FillPath(bezel, path)
                End Using
            End Using
            Dim scr = RectangleF.Inflate(mon, -6, -6)
            Using surround As New SolidBrush(Color.FromArgb(38, 38, 38))
                g.FillRectangle(surround, scr)
            End Using
            Using f As New Font("Segoe UI", 9.5F, GraphicsUnit.Pixel), ink As New SolidBrush(p.Muted)
                If ForLens Then
                    PaintLens(g, scr, f, ink)
                Else
                    PaintOverlay(g, mon, scr, f, ink)
                End If
            End Using
        Finally
            g.Restore(state)
        End Try
    End Sub

    ''' <summary>The screen colour at the right-hand edge; the overlay on a white card beside it, under a lamp.</summary>
    Private Sub PaintOverlay(g As Graphics, mon As RectangleF, scr As RectangleF, f As Font, ink As Brush)
        Dim patch As New RectangleF(scr.Right - 44, scr.Y + scr.Height / 2 - 30, 44, 60)
        Using b As New SolidBrush(Tint)
            g.FillRectangle(b, patch)
        End Using
        ' The white card, held against the monitor's right edge, with the overlay on it
        Dim card As New RectangleF(mon.Right + 4, patch.Y - 8, 58, 76)
        Using shadow As New SolidBrush(Color.FromArgb(40, 0, 0, 0))
            g.FillRectangle(shadow, card.X + 3, card.Y + 3, card.Width, card.Height)
        End Using
        g.FillRectangle(Brushes.White, card)
        Using pen As New Pen(Color.FromArgb(170, 170, 170))
            g.DrawRectangle(pen, card.X, card.Y, card.Width, card.Height)
        End Using
        Dim film = RectangleF.Inflate(card, -6, -8)
        Using b As New SolidBrush(Color.FromArgb(200, Tint)), pen As New Pen(Draw.Mix(Tint, Color.Black, 0.25), 1.2F)
            g.FillRectangle(b, film)
            g.DrawRectangle(pen, film.X, film.Y, film.Width, film.Height)
        End Using
        ' A lamp above, its light falling on the card (and the screen's edge)
        Dim lamp As New PointF(card.X + card.Width / 2 + 34, 8)
        Using cone As New GraphicsPath()
            cone.AddPolygon({New PointF(lamp.X - 8, lamp.Y + 8), New PointF(lamp.X + 8, lamp.Y + 8),
                             New PointF(card.Right + 10, card.Bottom), New PointF(card.X - 14, card.Bottom)})
            Using light As New LinearGradientBrush(New PointF(0, lamp.Y), New PointF(0, card.Bottom),
                                                   Color.FromArgb(150, 255, 236, 160), Color.FromArgb(20, 255, 236, 160))
                g.FillPath(light, cone)
            End Using
        End Using
        Using arm As New Pen(Color.FromArgb(110, 115, 125), 2.5F), head As New SolidBrush(Color.FromArgb(90, 95, 105)),
              bulb As New SolidBrush(Color.FromArgb(255, 240, 170))
            g.DrawLine(arm, lamp.X + 26, 0, lamp.X + 8, lamp.Y)
            Using shade As New GraphicsPath()
                shade.AddPolygon({New PointF(lamp.X - 6, lamp.Y - 4), New PointF(lamp.X + 6, lamp.Y - 4),
                                  New PointF(lamp.X + 11, lamp.Y + 8), New PointF(lamp.X - 11, lamp.Y + 8)})
                g.FillPath(head, shade)
            End Using
            g.FillEllipse(bulb, lamp.X - 5, lamp.Y + 5, 10, 6)
        End Using
        ' Labels
        Dim fmt As New StringFormat With {.Alignment = StringAlignment.Center}
        g.DrawString(Loc.T("readez.diagram.screen"), f, ink, New RectangleF(scr.X, mon.Bottom + 17, scr.Width, 14), fmt)
        g.DrawString(Loc.T("readez.diagram.card"), f, ink, New RectangleF(card.X - 30, card.Bottom + 6, card.Width + 60, 28), fmt)
        Dim lightFmt As New StringFormat With {.Alignment = StringAlignment.Near}
        g.DrawString(Loc.T("readez.diagram.light"), f, ink, New PointF(lamp.X + 14, lamp.Y + 6), lightFmt)
    End Sub

    ''' <summary>Lenses: the white window in the middle of the screen, the lens colour beside it, the lens held over the window.</summary>
    Private Sub PaintLens(g As Graphics, scr As RectangleF, f As Font, ink As Brush)
        Dim cx = scr.X + scr.Width / 2, cy = scr.Y + scr.Height / 2
        Dim r = 22.0F
        g.FillEllipse(Brushes.White, cx - r, cy - r, 2 * r, 2 * r)
        Using b As New SolidBrush(Tint)
            g.FillEllipse(b, cx + r + 10, cy - r * 0.8F, r * 1.6F, r * 1.6F)
        End Using
        ' The lens, held slightly in front of the window (a tinted disc in a thin frame)
        Using b As New SolidBrush(Color.FromArgb(170, Tint)), pen As New Pen(Color.FromArgb(70, 70, 80), 2)
            g.FillEllipse(b, cx - r - 4, cy - r - 2, 2 * r + 8, 2 * r + 8)
            g.DrawEllipse(pen, cx - r - 4, cy - r - 2, 2 * r + 8, 2 * r + 8)
        End Using
        Dim fmt As New StringFormat With {.Alignment = StringAlignment.Near}
        g.DrawString(Loc.T("readez.diagram.lens"), f, ink, New RectangleF(scr.Right + 14, scr.Y + 10, 110, 90), fmt)
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
