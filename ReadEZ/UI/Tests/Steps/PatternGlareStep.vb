Option Strict On

Imports System.Globalization

' ====================================================================================================
'   PatternGlareStep.vb - the pattern glare test: the first step of the baseline pattern test
'   (Tests > Pattern > Pattern glare); symptoms with that pattern on white follow (SymptomsStep).
'
'   Two square patches of horizontal stripes (square-wave gratings, black and white, full contrast)
'   on a mid-grey of the same average brightness, one each side of the screen, with different
'   spatial frequencies (cycles per degree at the viewing distance). The patient says which is the
'   MORE uncomfortable to look at - or that they are about the same - and two new frequencies are
'   shown, until the most uncomfortable frequency has been narrowed down.
'
'   How it narrows down (a "ternary search" for the peak of the discomfort curve, on a log scale):
'     The range starts at 0.5 c/deg (or 3 cycles across a patch) to 16 c/deg (or 2 pixels per
'     stripe, the finest the screen can show). The two patterns are at one third and two thirds
'     of the range (in octaves), sides at random. The less uncomfortable one's outer third is
'     dropped; "about the same" keeps the middle third. The range shrinks by a third each time:
'     about 7 choices take 5 octaves down to a third of an octave. The result is the middle of
'     what is left. Each stripe is a whole number of pixels, so the frequencies recorded are the
'     ones actually shown.
'
'   The cycles per degree depend on the viewing distance (asked for, 40 cm by default) and the
'   screen's size (Settings > General > the patient's screen width in mm).
'   The frequency found is the test's PatternCpd (with PatternDistanceCm); the search itself goes in
'   AutoScore ("PG;lo=...;hi=...;n=...;done=...;trials=..."). The colour pattern tests (overlays,
'   lenses, computer) then show this pattern instead of the reading text (ReadingScene.PatternStripe).
' ====================================================================================================

#Region "The patient's screen"

''' <summary>
''' Two patches of horizontal black-and-white stripes on mid-grey (the grey has the stripes' average
''' brightness). Drawn in the patient's screen pixels; each stripe a whole number of pixels.
''' </summary>
Public NotInheritable Class GratingScene
    Implements IStimulusScene

    ''' <summary>Mid-grey of the same brightness as the stripes on average (50 % luminance in sRGB).</summary>
    Public Shared ReadOnly Grey As Color = Color.FromArgb(186, 186, 186)
    Private Shared ReadOnly Marker As Color = Color.FromArgb(40, 110, 200)

    ''' <summary>Height of one stripe (pixels) on the left / right; 0 = no patch.</summary>
    Public Property LeftStripe As Integer
    Public Property RightStripe As Integer
    ''' <summary>The choice just made, ringed for a moment: 0 left, 1 right, 2 both ("about the same"); -1 none.</summary>
    Public Property Highlight As Integer = -1
    ''' <summary>The patient's own "About the same" button at the bottom (patient chooses / tablet).</summary>
    Public Property ShowSameButton As Boolean
    ''' <summary>Text in the middle instead of the patches (e.g. "Thank you").</summary>
    Public Property Message As String = ""

    ''' <summary>A patch (patient's screen pixels): side 0 left, 1 right. Square, the same size each side.</summary>
    Public Shared Function PatchRect(sz As Size, side As Integer) As Rectangle
        Dim s = CInt(Math.Min(sz.Height * 0.46, sz.Width * 0.36))
        Dim cx = CInt(If(side = 0, sz.Width * 0.27, sz.Width * 0.73))
        Return New Rectangle(cx - s \ 2, (sz.Height - s) \ 2, s, s)
    End Function

    Private Shared Function SameRect(sz As Size) As Rectangle
        Dim w = CInt(sz.Width * 0.22), h = CInt(sz.Height * 0.075)
        Return New Rectangle((sz.Width - w) \ 2, CInt(sz.Height * 0.87), w, h)
    End Function

    Public Sub Paint(g As Graphics, bounds As Rectangle, preview As Boolean) Implements IStimulusScene.Paint
        Using b As New SolidBrush(Grey)
            g.FillRectangle(b, bounds)
        End Using
        Dim sz = StimulusArt.ScreenSize
        Dim k = bounds.Width / CSng(Math.Max(1, sz.Width))
        Dim map = Function(rc As Rectangle) New RectangleF(bounds.X + rc.X * k, bounds.Y + rc.Y * k, rc.Width * k, rc.Height * k)
        If Message <> "" Then
            Using f As New Font("Segoe UI Semibold", Math.Max(6.0F, sz.Height * 0.05F * k), GraphicsUnit.Pixel),
                  fmt As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center},
                  ink As New SolidBrush(Color.FromArgb(50, 50, 60))
                g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
                g.DrawString(Message, f, ink, bounds, fmt)
            End Using
            Return
        End If
        Dim oldSmoothing = g.SmoothingMode
        g.SmoothingMode = Drawing2D.SmoothingMode.None
        For side = 0 To 1
            Dim stripe = If(side = 0, LeftStripe, RightStripe)
            If stripe > 0 Then DrawGrating(g, map(PatchRect(sz, side)), stripe * k)
        Next
        ' The choice just made: a blue frame round it
        If Highlight >= 0 Then
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using pen As New Pen(Marker, Math.Max(2.0F, sz.Height * 0.012F * k))
                For side = 0 To 1
                    If Highlight = side OrElse Highlight = 2 Then
                        Dim r = map(PatchRect(sz, side))
                        Dim pad = sz.Height * 0.02F * k
                        g.DrawRectangle(pen, r.X - pad, r.Y - pad, r.Width + 2 * pad, r.Height + 2 * pad)
                    End If
                Next
            End Using
        End If
        g.SmoothingMode = oldSmoothing
        If ShowSameButton Then
            Dim r = map(SameRect(sz))
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using path = Draw.RoundedRect(r, r.Height * 0.25F), b As New SolidBrush(If(Highlight = 2, Marker, Color.FromArgb(80, 80, 90))),
                  f As New Font("Segoe UI Semibold", Math.Max(6.0F, r.Height * 0.38F), GraphicsUnit.Pixel),
                  fmt As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                g.FillPath(b, path)
                g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
                g.DrawString(Loc.T("pg.patient.same"), f, Brushes.White, r, fmt)
            End Using
            g.SmoothingMode = oldSmoothing
        End If
    End Sub

    ''' <summary>Horizontal stripes, black first, each stripeHeight high (exact pixels on the patient's screen).</summary>
    Private Shared Sub DrawGrating(g As Graphics, r As RectangleF, stripeHeight As Single)
        g.FillRectangle(Brushes.White, r)
        If stripeHeight <= 0 Then Return
        Dim y = r.Y
        While y < r.Bottom
            g.FillRectangle(Brushes.Black, r.X, y, r.Width, Math.Min(stripeHeight, r.Bottom - y))
            y += 2 * stripeHeight
        End While
    End Sub

    ''' <summary>What a click at p is on: 0 the left patch, 1 the right, 2 "About the same", -1 nothing.</summary>
    Public Function HitTest(bounds As Rectangle, p As Point) As Integer
        Dim sz = StimulusArt.ScreenSize
        Dim k = bounds.Width / CSng(Math.Max(1, sz.Width))
        If k <= 0 Then Return -1
        Dim sp As New Point(CInt((p.X - bounds.X) / k), CInt((p.Y - bounds.Y) / k))
        If ShowSameButton AndAlso SameRect(sz).Contains(sp) Then Return 2
        For side = 0 To 1
            Dim r = PatchRect(sz, side)
            r.Inflate(r.Width \ 10, r.Height \ 10)
            If r.Contains(sp) Then Return side
        Next
        Return -1
    End Function
End Class

#End Region

#Region "The step"

Public NotInheritable Class PatternGlareStep
    Inherits TestStep

    ''' <summary>Three horizontal lines (Segoe MDL2): the test's symbol.</summary>
    Public Const Glyph As String = ChrW(&HE700)
    ''' <summary>Stop when the range is a third of an octave or less, or after this many choices.</summary>
    Private Const StopOctaves As Double = 0.34
    Private Const MaxChoices As Integer = 12
    Private Const MinCpd As Double = 0.5, MaxCpd As Double = 16.0

    Private ReadOnly _scene As New GratingScene
    Private ReadOnly _random As New Random()
    Private ReadOnly _flash As New Timer With {.Interval = 450}
    Private ReadOnly nudDistance As NumericUpDown = Ui.NumberBox(20, 200, 40, width:=110)
    Private ReadOnly lblProgress As New Label With {.AutoSize = True, .Font = Fonts.Strong, .Margin = New Padding(0, 6, 0, 4)}
    Private ReadOnly lblRange As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(330, 0), .Margin = New Padding(0, 0, 0, 8)}
    Private ReadOnly lblResult As New Label With {.AutoSize = True, .Font = Fonts.Title, .MaximumSize = New Size(330, 0), .Margin = New Padding(0, 8, 0, 4)}
    Private ReadOnly lblTrials As New Label With {.AutoSize = True, .Font = Fonts.Small, .MaximumSize = New Size(330, 0), .Margin = New Padding(0, 6, 0, 0)}
    Private ReadOnly btnLeft As ModernButton = ModernButton.Create("pg.left", ButtonKind.Primary, Glyphs.Back)
    Private ReadOnly btnRight As ModernButton = ModernButton.Create("pg.right", ButtonKind.Primary, Glyphs.Forward)
    Private ReadOnly btnSame As ModernButton = ModernButton.Create("pg.same", ButtonKind.Secondary)

    ' The search (octaves: log2 of cycles per degree)
    Private _lo, _hi, _x1, _x2 As Double
    Private _lowOnLeft As Boolean
    Private _leftCpd, _rightCpd As Double
    Private _done, _busy, _started As Boolean
    Private ReadOnly _choices As New List(Of String)
    Private _pendingSide As Integer = -1

    Public Overrides ReadOnly Property TitleKey As String
        Get
            Return "step.patternGlare"
        End Get
    End Property

    Public Overrides ReadOnly Property ViewScene As IStimulusScene
        Get
            Return _scene
        End Get
    End Property

    Public Sub New()
        nudDistance.Value = Math.Max(nudDistance.Minimum, Math.Min(nudDistance.Maximum, AppHost.Settings.PatternGlareDistanceCm))
        AddHandler _flash.Tick, Sub()
                                    _flash.Stop()
                                    ApplyChoice(_pendingSide)
                                End Sub
    End Sub

    ''' <summary>The result as it stands (copied into the test by UpdateResult).</summary>
    Private NotInheritable Class Estimate
        Public PeakCpd, LowCpd, HighCpd As Double
        Public DistanceCm As Integer
        Public Completed As Boolean
        ''' <summary>Each choice: "left c/deg,right c/deg,L|R|S", separated by spaces.</summary>
        Public Trials As String = ""
    End Class

    Private ReadOnly Result As New Estimate

#Region "Panel"

    Public Overrides Function CreatePanel() As Control
        Dim col = Ui.Column()
        Ui.FillWidth(col)
        Dim lblAsk = Ui.Label("pg.ask")
        lblAsk.Font = Fonts.Section
        lblAsk.MaximumSize = New Size(340, 0)
        lblAsk.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(lblAsk)

        Dim answers = Ui.Row(wrap:=True)
        For Each b In {btnLeft, btnRight, btnSame}
            b.Margin = New Padding(0, 0, 8, 8)
            answers.Controls.Add(b)
        Next
        AddHandler btnLeft.Click, Sub() Choose(0)
        AddHandler btnRight.Click, Sub() Choose(1)
        AddHandler btnSame.Click, Sub() Choose(2)
        col.Controls.Add(answers)
        Dim lblKeys = Ui.Label("pg.keys", muted:=True)
        lblKeys.MaximumSize = New Size(340, 0)
        col.Controls.Add(lblKeys)

        col.Controls.Add(lblProgress)
        LabelStyle.SetMuted(lblRange)
        col.Controls.Add(lblRange)
        col.Controls.Add(lblResult)

        ' Viewing distance (sets the cycles per degree): before the first choice
        AddHandler nudDistance.ValueChanged, Sub()
                                                 AppHost.Settings.PatternGlareDistanceCm = CInt(nudDistance.Value)
                                                 AppHost.SaveSettings()
                                                 If _choices.Count = 0 Then Restart()
                                             End Sub
        col.Controls.Add(Ui.Stack("pg.distance", nudDistance))
        Dim distHint = Ui.Label("pg.distanceHint", muted:=True)
        distHint.MaximumSize = New Size(340, 0)
        col.Controls.Add(distHint)

        Dim actions = Ui.Row(wrap:=True)
        actions.Margin = New Padding(0, 10, 0, 0)
        Dim restartButton = ModernButton.Create("pg.restart", ButtonKind.Ghost, Glyphs.Refresh)
        AddHandler restartButton.Click, Sub() Restart()
        restartButton.Margin = New Padding(0, 0, 8, 8)
        actions.Controls.Add(restartButton)
        Dim how = HelpButton("pg.how.button", "pg.how.title", "pg.how.1", "pg.how.2", "pg.how.3", "pg.how.4", "pg.how.5")
        actions.Controls.Add(how)
        col.Controls.Add(actions)
        LabelStyle.SetMuted(lblTrials)
        col.Controls.Add(lblTrials)
        ShowState()
        Return col
    End Function

    Public Overrides Function CreateFooter() As Control
        Return FooterRow(NextButton())
    End Function

    Private Sub ShowState()
        Dim n = _choices.Count
        lblProgress.Text = If(_done, Loc.T("pg.finished"), Loc.F("pg.choice", n + 1))
        lblRange.Text = Loc.F("pg.range", Cpd(Math.Pow(2, _lo)), Cpd(Math.Pow(2, _hi))) & vbCrLf &
                        If(_done, "", Loc.F("pg.showing", Cpd(_leftCpd), Cpd(_rightCpd)))
        lblResult.Visible = _done
        If _done Then lblResult.Text = Loc.F("pg.result", Cpd(Result.PeakCpd))
        lblTrials.Text = String.Join(vbCrLf, _choices.Select(Function(c, i) $"{i + 1}.  {c}"))
        Dim answering = _started AndAlso Not _done AndAlso Not _busy
        btnLeft.Enabled = answering
        btnRight.Enabled = answering
        btnSame.Enabled = answering
        nudDistance.Enabled = n = 0
    End Sub

    ''' <summary>"3.2" (cycles per degree, one decimal place; whole numbers from 10 up).</summary>
    Friend Shared Function Cpd(v As Double) As String
        Return If(v >= 10, v.ToString("0", Loc.Culture), v.ToString("0.0", Loc.Culture))
    End Function

#End Region

#Region "Instructions"

    Public Overrides Function InstructionKeys() As String()
        Return {"instr.pg.title", "instr.pg.1", "instr.pg.2", "instr.pg.3"}
    End Function

    ''' <summary>Two example patterns, with "1. Look at this one", "2. ... and this one", "3. Which is more uncomfortable?".</summary>
    Public Overrides Function InstructionScreen() As InstructionScene
        Dim ppd = PixelsPerDegree()
        Dim sample As New GratingScene With {.LeftStripe = StripeFor(1.0, ppd), .RightStripe = StripeFor(4.0, ppd)}
        Dim sz = StimulusArt.ScreenSize
        Dim sc As New InstructionScene(InstructionKeys()) With {.Sample = sample}
        sc.Add("instr.pg.callout1", GratingScene.PatchRect(sz, 0), InstructionScene.Side.Below)
        sc.Add("instr.pg.callout2", GratingScene.PatchRect(sz, 1), InstructionScene.Side.Below)
        sc.Add("instr.pg.callout3", RectangleF.Empty, InstructionScene.Side.BottomCentre)
        Return sc
    End Function

#End Region

#Region "The search"

    ''' <summary>Pixels per degree at the viewing distance being used.</summary>
    Private Function PixelsPerDegree() As Double
        Return PixelsPerDegreeAt(CDbl(nudDistance.Value))
    End Function

    ''' <summary>Pixels per degree at a viewing distance (the patient's screen's pixels per mm from Settings).</summary>
    Public Shared Function PixelsPerDegreeAt(distanceCm As Double) As Double
        Return StimulusArt.PxPerMm * Math.Max(1.0, distanceCm) * 10 * Math.Tan(Math.PI / 180)
    End Function

    ''' <summary>A pattern test's stripe height on the patient's screen now (0 = not a pattern test).</summary>
    Public Shared Function StripeOf(r As ReadEZResult) As Integer
        If r Is Nothing OrElse Not r.IsPattern OrElse r.PatternCpd <= 0 Then Return 0
        Return StripeFor(r.PatternCpd, PixelsPerDegreeAt(If(r.PatternDistanceCm > 0, r.PatternDistanceCm, 40)))
    End Function

    ''' <summary>The search recorded in a baseline pattern test's AutoScore: the range left and the number of choices.</summary>
    Public Shared Function ParseScore(score As String) As (Low As Double, High As Double, Choices As Integer, Completed As Boolean)
        Dim lo = 0.0, hi = 0.0, n = 0, done = False
        If score Is Nothing OrElse Not score.StartsWith("PG;", StringComparison.Ordinal) Then Return (0, 0, 0, False)
        For Each part In score.Split(";"c)
            Dim eq = part.IndexOf("="c)
            If eq <= 0 Then Continue For
            Dim name = part.Substring(0, eq), value = part.Substring(eq + 1)
            Select Case name
                Case "lo" : Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, lo)
                Case "hi" : Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, hi)
                Case "n" : Integer.TryParse(value, n)
                Case "done" : done = value = "1"
            End Select
        Next
        Return (lo, hi, n, done)
    End Function

    ''' <summary>The stripe height in whole pixels nearest to a frequency (a cycle is two stripes).</summary>
    Public Shared Function StripeFor(cpd As Double, ppd As Double) As Integer
        Return Math.Max(1, CInt(Math.Round(ppd / (2 * cpd))))
    End Function

    Private Shared Function CpdOf(stripe As Integer, ppd As Double) As Double
        Return ppd / (2.0 * Math.Max(1, stripe))
    End Function

    ''' <summary>Starts (again): the whole range, the first two patterns.</summary>
    Private Sub Restart()
        _flash.Stop()
        _busy = False
        _done = False
        _started = True
        _choices.Clear()
        _scene.Message = ""
        _scene.Highlight = -1
        Dim ppd = PixelsPerDegree()
        Dim side = GratingScene.PatchRect(StimulusArt.ScreenSize, 0).Height
        ' At least 3 cycles across a patch; at least 2 pixels per stripe
        Dim lo = Math.Max(MinCpd, 3 * ppd / Math.Max(1, side))
        Dim hi = Math.Min(MaxCpd, ppd / 4)
        If hi < lo * 2 Then hi = lo * 2
        _lo = Math.Log(lo, 2)
        _hi = Math.Log(hi, 2)
        UpdateResult(completed:=False)
        NextPair()
    End Sub

    ''' <summary>The next two patterns, at a third and two thirds of the range (sides at random) - or the result.</summary>
    Private Sub NextPair()
        _scene.Highlight = -1
        Dim ppd = PixelsPerDegree()
        If _hi - _lo <= StopOctaves OrElse _choices.Count >= MaxChoices Then
            Complete()
            Return
        End If
        _x1 = _lo + (_hi - _lo) / 3
        _x2 = _lo + 2 * (_hi - _lo) / 3
        Dim s1 = StripeFor(Math.Pow(2, _x1), ppd), s2 = StripeFor(Math.Pow(2, _x2), ppd)
        If s1 = s2 Then
            Complete()                                  ' the screen cannot show them differently any more
            Return
        End If
        _lowOnLeft = _random.Next(2) = 0
        _scene.LeftStripe = If(_lowOnLeft, s1, s2)
        _scene.RightStripe = If(_lowOnLeft, s2, s1)
        _leftCpd = CpdOf(_scene.LeftStripe, ppd)
        _rightCpd = CpdOf(_scene.RightStripe, ppd)
        _scene.ShowSameButton = AppHost.Settings.PatientChooses OrElse TabletScreen.IsActive
        ShowGratings()
        ShowState()
        ActionsChanged()
    End Sub

    ''' <summary>A choice (0 left, 1 right, 2 about the same): ringed on the screen for a moment, then applied.</summary>
    Private Sub Choose(side As Integer)
        If Not _started OrElse _done OrElse _busy OrElse side < 0 Then Return
        _busy = True
        _pendingSide = side
        _scene.Highlight = side
        ShowGratings()
        ShowState()
        _flash.Start()
    End Sub

    Private Sub ApplyChoice(side As Integer)
        _busy = False
        If _done Then Return
        Dim mark = If(side = 0, "L", If(side = 1, "R", "S"))
        Dim chosenLow = (side = 0 AndAlso _lowOnLeft) OrElse (side = 1 AndAlso Not _lowOnLeft)
        If side = 2 Then
            _lo = _x1 : _hi = _x2                       ' about the same: the peak is between them
        ElseIf chosenLow Then
            _hi = _x2                                   ' the lower one is worse: the peak is below the higher one
        Else
            _lo = _x1                                   ' the higher one is worse: the peak is above the lower one
        End If
        Dim chosenText = If(side = 2, Loc.T("pg.trial.same"), Loc.F("pg.trial.chose", Cpd(If(side = 0, _leftCpd, _rightCpd))))
        _choices.Add(Loc.F("pg.trial", Cpd(_leftCpd), Cpd(_rightCpd), chosenText))
        Dim list = If(Result.Trials = "", New List(Of String), Result.Trials.Split(" "c).ToList())
        list.Add(String.Format(CultureInfo.InvariantCulture, "{0:0.00},{1:0.00},{2}", _leftCpd, _rightCpd, mark))
        Result.Trials = String.Join(" ", list)
        UpdateResult(completed:=False)
        NextPair()
    End Sub

    ''' <summary>Narrowed down: the result is the middle of the range left (in octaves).</summary>
    Private Sub Complete()
        _done = True
        UpdateResult(completed:=True)
        _scene.LeftStripe = 0
        _scene.RightStripe = 0
        _scene.ShowSameButton = False
        _scene.Message = Loc.T("pg.patient.thanks")
        ShowGratings()
        ShowState()
        ActionsChanged()
        System.Media.SystemSounds.Asterisk.Play()
        Host?.ShowExaminerMessage(Loc.F("pg.result", Cpd(Result.PeakCpd)))
    End Sub

    ''' <summary>The result as it stands (a test stopped early keeps its estimate so far).</summary>
    Private Sub UpdateResult(completed As Boolean)
        Dim r = Result
        r.PeakCpd = Math.Pow(2, (_lo + _hi) / 2)
        r.LowCpd = Math.Pow(2, _lo)
        r.HighCpd = Math.Pow(2, _hi)
        r.DistanceCm = CInt(nudDistance.Value)
        r.Completed = completed
        If Not completed AndAlso _choices.Count = 0 Then r.Trials = ""
        ' Into the test: the worst frequency, the distance, and the search
        Dim t = Run.Result
        t.IsPattern = True
        t.PatternCpd = r.PeakCpd
        t.PatternDistanceCm = r.DistanceCm
        t.AutoScore = String.Format(CultureInfo.InvariantCulture, "PG;lo={0:0.000};hi={1:0.000};n={2};done={3};trials={4}",
                                    r.LowCpd, r.HighCpd, _choices.Count, If(completed, 1, 0), r.Trials)
    End Sub

#End Region

#Region "Step"

    Private Sub ShowGratings()
        Stimulus.Window.BackColor = GratingScene.Grey
        If Stimulus.Window.Scene IsNot _scene Then Stimulus.Window.Scene = _scene Else Stimulus.Window.RequestRedraw()
        Host?.RefreshView()
    End Sub

    Public Overrides Sub Begin()
        If Not _started Then Restart() Else ShowGratings()
        ShowState()
    End Sub

    Public Overrides Sub [End]()
        _flash.Stop()
        _busy = False
    End Sub

    Public Overrides Function IsComplete() As Boolean
        Return _done
    End Function

    Public Overrides Sub PreviewMouseClick(bounds As Rectangle, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then Choose(_scene.HitTest(bounds, e.Location))
    End Sub

    ''' <summary>The patient (or, with one screen, the examiner) clicks or taps the more uncomfortable pattern.</summary>
    Public Overrides Sub PatientScreenClick(bounds As Rectangle, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then Choose(_scene.HitTest(bounds, e.Location))
    End Sub

    Public Overrides ReadOnly Property PreviewCursor As Cursor
        Get
            Return If(_done, Cursors.Default, Cursors.Hand)
        End Get
    End Property

    ''' <summary>Left / Right arrow: that side is worse; Down or Space: about the same; Enter: finish once done.</summary>
    Public Overrides Function HandleKey(keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Left : Choose(0) : Return True
            Case Keys.Right : Choose(1) : Return True
            Case Keys.Down, Keys.Space : Choose(2) : Return True
            Case Keys.Enter
                If _done Then Finish()
                Return True
        End Select
        Return False
    End Function

    Public Overrides Function QuickActions() As List(Of QuickAction)
        Dim answering = _started AndAlso Not _done AndAlso Not _busy
        Dim list As New List(Of QuickAction)
        If _done Then
            list.Add(New QuickAction("", Nothing) With {.Text = Loc.T("pg.bar.result"), .Value = Cpd(Result.PeakCpd), .GroupStart = True})
            list.Add(New QuickAction("", Nothing) With {.Text = Loc.T("pg.bar.unit")})
        Else
            list.Add(New QuickAction("", Nothing) With {.Text = Loc.T("pg.bar.choice"), .Value = (_choices.Count + 1).ToString(), .GroupStart = True})
            list.Add(New QuickAction("", Sub() Choose(0), primary:=True, enabled:=answering) With {.Text = Loc.T("pg.left") & "  (←)", .GroupStart = True})
            list.Add(New QuickAction("", Sub() Choose(2), enabled:=answering) With {.Text = Loc.T("pg.same") & "  (↓)"})
            list.Add(New QuickAction("", Sub() Choose(1), primary:=True, enabled:=answering) With {.Text = Loc.T("pg.right") & "  (→)"})
        End If
        list.Add(New QuickAction("pg.restart", Sub() Restart()) With {.GroupStart = True})
        list.Add(New QuickAction(If(Run.Index = Run.Steps.Count - 1, "run.saveFinish", "run.nextStep"), Sub() Finish(), primary:=_done, enabled:=_done) With {.GroupStart = True})
        Return list
    End Function

    Public Overrides Sub Dispose()
        _flash.Dispose()
        MyBase.Dispose()
    End Sub

#End Region

End Class

#End Region
