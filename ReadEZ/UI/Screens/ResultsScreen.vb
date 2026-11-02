Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   ResultsScreen.vb - a patient's results (replaces RZfrmResults' Summary panel and RZfrmResultsDetail)
'
'     Left    the patient's tests, newest first
'     Right   the selected test:
'               - the colour chosen (swatch, name, test, method; the automatic test's colour wheel
'                 with its spokes) or "white" for a baseline
'               - symptoms: the four grades, beside the baseline's for a colour test
'               - rate of reading: words per minute, errors, rows skipped, beside the baseline's
'               - change with colour: the difference in reading speed (words / min and %) and in
'                 symptoms compared with the baseline (the last one before the colour test)
'     Header  Report, Letter (parents / teacher), Show on the patient's screen, Notes, Delete
' ====================================================================================================

Public Class ResultsScreen
    Inherits ScreenBase

    Private ReadOnly _list As New ResultsList
    Private ReadOnly _detail As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True, .BackColor = Color.Transparent}
    Private ReadOnly lblEmpty As New Label With {.Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter}
    Private _results As New List(Of ReadEZResult)
    Private _current As ReadEZResult
    Private btnReport, btnLetter, btnShow, btnNotes, btnDelete As ModernButton

    Public Sub New()
        TitleKey = "results.title"
        HelpTopic = "screen.results"
        btnDelete = AddHeaderButton("common.delete", ButtonKind.Ghost, Glyphs.Delete)
        btnNotes = AddHeaderButton("results.notes", ButtonKind.Ghost, Glyphs.Edit)
        btnShow = AddHeaderButton("results.show", ButtonKind.Ghost, Glyphs.Eye)
        btnLetter = AddHeaderButton("results.letter", ButtonKind.Secondary, Glyphs.Report)
        btnReport = AddHeaderButton("results.report", ButtonKind.Primary, Glyphs.Print)
        AddHandler btnReport.Click, Sub() If _current IsNot Nothing Then ReadEZReport.ShowReport(FindForm(), _current, BaselineFor(_current))
        AddHandler btnLetter.Click, Sub() If _current IsNot Nothing Then ReadEZReport.ShowLetter(FindForm(), _current, BaselineFor(_current))
        AddHandler btnShow.Click, Sub() ShowOnPatientScreen()
        AddHandler btnNotes.Click, Sub()
                                       AppCommands.EditResult()
                                       Reload()
                                   End Sub
        AddHandler btnDelete.Click, Sub()
                                        AppCommands.DeleteResult()
                                        Reload()
                                    End Sub

        Dim split As New SplitContainer With {.Dock = DockStyle.Fill, .SplitterWidth = 12, .FixedPanel = FixedPanel.Panel1, .BackColor = Color.Transparent}
        ThemeManager.SetRole(split, SurfaceRole.Canvas)
        Dim placed = False
        AddHandler split.SizeChanged, Sub()
                                          If placed OrElse split.Width < 400 Then Return
                                          placed = True
                                          split.SplitterDistance = Math.Min(380, CInt(split.Width * 0.3))
                                      End Sub
        Dim listCard As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(10, 48, 10, 10)}
        listCard.SetTitle("tests.list.title")
        _list.Dock = DockStyle.Fill
        listCard.Controls.Add(_list)
        split.Panel1.Controls.Add(listCard)
        LabelStyle.SetMuted(lblEmpty)
        Loc.Bind(lblEmpty, "results.none")
        split.Panel2.Controls.Add(_detail)
        split.Panel2.Controls.Add(lblEmpty)
        Content.Controls.Add(split)

        AddHandler _list.ResultSelected, Sub()
                                             Dim r = _list.SelectedResult
                                             If r IsNot Nothing Then AppHost.Session.SelectResult(r.ID)
                                             ShowResult(r)
                                         End Sub
        AddHandler _detail.ClientSizeChanged, Sub() FitCards()
        AddHandler AppCommands.DataChanged, Sub(c) If Visible Then Reload()
        Loc.OnChange(Me, Sub() If Visible Then ShowResult(_current))
    End Sub

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        Reload()
    End Sub

    Private Sub Reload()
        Dim s = AppHost.Session
        SetSubtitleText(If(s.HasPatient, s.PatientName, Loc.T("top.noPatient")))
        _results = If(s.HasPatient, AppHost.Db.GetResults(s.PatientId), New List(Of ReadEZResult))
        _list.SetResults(_results, s.ResultId)
        lblEmpty.Visible = _results.Count = 0
        _detail.Visible = _results.Count > 0
        ShowResult(_list.SelectedResult)
    End Sub

    ''' <summary>The baseline to compare a colour test with: the last one before it, or else the newest.</summary>
    Private Function BaselineFor(r As ReadEZResult) As ReadEZResult
        If r Is Nothing OrElse r.IsBaseline Then Return Nothing
        Dim baselines = _results.Where(Function(x) x.IsBaseline).ToList()
        Return If(baselines.Where(Function(b) b.Test_Date <= r.Test_Date).OrderByDescending(Function(b) b.Test_Date).FirstOrDefault(),
                  baselines.OrderByDescending(Function(b) b.Test_Date).FirstOrDefault())
    End Function

#Region "Detail"

    Private Sub ShowResult(r As ReadEZResult)
        _current = r
        For Each b In {btnReport, btnLetter, btnShow, btnNotes, btnDelete}
            b.Enabled = r IsNot Nothing
        Next
        _detail.SuspendLayout()
        For Each c In _detail.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        If r IsNot Nothing Then
            Dim baseline = BaselineFor(r)
            _detail.Controls.Add(SummaryCard(r))
            If Not r.IsBaseline AndAlso baseline IsNot Nothing Then _detail.Controls.Add(ChangeCard(r, baseline))
            _detail.Controls.Add(SymptomsCard(r, baseline))
            _detail.Controls.Add(ReadingCard(r, baseline))
            Dpi.ScaleNew(_detail, Me)
            ThemeManager.Apply(_detail)
        End If
        _detail.ResumeLayout(True)
        FitCards()
    End Sub

    Private Sub FitCards()
        Dim w = Math.Max(500, _detail.ClientSize.Width - 24)
        For Each c As Control In _detail.Controls
            c.Width = w
        Next
    End Sub

    Private Shared Function NewCard(titleKey As String, height As Integer) As Card
        Dim c As New Card With {.Height = height, .Padding = New Padding(20, 48, 20, 12), .Margin = New Padding(0, 0, 0, 12)}
        c.SetTitle(titleKey)
        Return c
    End Function

    Private Function SummaryCard(r As ReadEZResult) As Control
        Dim card = NewCard(If(r.IsBaseline, "results.baseline", "results.colourResult"), If(r.IsBaseline OrElse r.Screening_Mode <> 0, 190, 300))
        Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.Transparent}
        Dim sw As New ColourSwatch With {.Colour = r.ScreenColour, .Size = New Size(110, 110), .Location = New Point(0, 0),
                                         .Mode = r.Mode}
        host.Controls.Add(sw)
        Dim lines As New List(Of String)
        lines.Add(Loc.T(TestRun.ModeKey(r.Mode)) & "  ·  " & r.Test_Date.ToString("f", Loc.Culture))
        If Not r.IsBaseline Then
            lines.Add(Loc.F("results.method", Loc.T(If(r.Screening_Mode = 0, "colour.auto", If(r.Screening_Mode = 1, "colour.manual", "results.bespoke")))))
            lines.Add($"RGB {r.ScreenColour.R}, {r.ScreenColour.G}, {r.ScreenColour.B}")
        End If
        Dim screener = AppHost.Db.ScreenerName(r.Screener_ID)
        If screener <> "" Then lines.Add(Loc.F("results.screener", screener))
        If r.ROR_Text <> "" Then lines.Add(Loc.F("results.text", r.ROR_Text, r.RoR_FontName, r.RoR_FontSize + 8, r.RoR_NoRows))
        If r.ImportedFrom <> "" Then lines.Add(Loc.F("results.imported", r.ImportedFrom))
        If r.Notes <> "" Then lines.Add(Loc.T("report.notes") & ": " & r.Notes)
        Dim name As New Label With {.AutoSize = True, .Font = Fonts.Title, .Location = New Point(130, 0), .Text = If(r.IsBaseline, Loc.T("result.white"), r.ColourName)}
        Dim info As New Label With {.AutoSize = True, .Font = Fonts.Body, .Location = New Point(132, 38), .MaximumSize = New Size(560, 0), .Text = String.Join(vbCrLf, lines)}
        LabelStyle.SetMuted(info)
        host.Controls.Add(name)
        host.Controls.Add(info)
        If r.GreyPreferred Then
            Dim warn As New Label With {.AutoSize = True, .Font = Fonts.Strong, .Location = New Point(132, 150), .Text = Loc.T("colour.greyPreferred")}
            LabelStyle.SetCustom(warn)
            warn.ForeColor = ThemeManager.Current.Warning
            host.Controls.Add(warn)
        End If
        If Not r.IsBaseline AndAlso r.Screening_Mode = 0 AndAlso r.AutoScore <> "" Then
            Dim wheel As New ColourWheel With {.Selectable = False, .Size = New Size(230, 230), .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                                               .OuterSet = If(r.Mode = ReadEZTestMode.Overlay, ReadEZSet.Overlay, ReadEZSet.Lens),
                                               .TwoRings = r.Mode <> ReadEZTestMode.Overlay, .Spokes = ColourStaircase.ParseScore(r.AutoScore),
                                               .Selected = (r.ColourSet, r.Colour1)}
            AddHandler host.Resize, Sub() wheel.Left = host.Width - wheel.Width
            host.Controls.Add(wheel)
        End If
        card.Controls.Add(host)
        Return card
    End Function

    Private Function ChangeCard(r As ReadEZResult, b As ReadEZResult) As Control
        Dim card = NewCard("results.change", 150)
        Dim v As New KeyResultsView With {.Dock = DockStyle.Fill}
        v.Items.AddRange(ChangeItems(r, b))
        card.Controls.Add(v)
        Return card
    End Function

    ''' <summary>The change with colour as key results (shared with the report).</summary>
    Public Shared Function ChangeItems(r As ReadEZResult, b As ReadEZResult) As List(Of KeyResultsView.Item)
        Dim items As New List(Of KeyResultsView.Item)
        If r.HasReading AndAlso b.HasReading Then
            Dim diff = r.ROR - b.ROR
            Dim pct = If(b.ROR > 0, 100.0 * diff / b.ROR, 0)
            Dim lvl = If(pct >= 5, Brand.Level.Good, If(pct > -5, Brand.Level.Caution, Brand.Level.Concern))
            items.Add(New KeyResultsView.Item(Loc.T("results.change.reading"), $"{If(diff >= 0, "+", "")}{diff} {Loc.T("reading.wpmUnit")}  ({If(pct >= 0, "+", "")}{pct:0}%)",
                                              Loc.T(If(pct >= 5, "results.change.readingGood", If(pct > -5, "results.change.readingSame", "results.change.readingWorse"))), lvl))
            Dim de = r.ROR_NoErrors - b.ROR_NoErrors
            items.Add(New KeyResultsView.Item(Loc.T("results.change.errors"), $"{If(de > 0, "+", "")}{de}", Loc.F("results.change.fromTo", b.ROR_NoErrors, r.ROR_NoErrors),
                                              If(de < 0, Brand.Level.Good, If(de = 0, Brand.Level.Info, Brand.Level.Caution))))
        End If
        If r.SymptomsDone AndAlso b.SymptomsDone Then
            Dim ds = r.SymptomTotal - b.SymptomTotal
            items.Add(New KeyResultsView.Item(Loc.T("results.change.symptoms"), $"{If(ds > 0, "+", "")}{ds}", Loc.F("results.change.fromTo", b.SymptomTotal, r.SymptomTotal),
                                              If(ds < 0, Brand.Level.Good, If(ds = 0, Brand.Level.Info, Brand.Level.Concern))))
        End If
        Return items
    End Function

    Private Function SymptomsCard(r As ReadEZResult, b As ReadEZResult) As Control
        Dim card = NewCard("results.symptoms", 260 + If(r.Symptoms_Notes <> "", 30, 0))
        Dim chart As New SymptomsChart With {.Dock = DockStyle.Fill, .Result = r, .Baseline = b}
        card.Controls.Add(chart)
        If r.Symptoms_Notes <> "" Then
            Dim notes As New Label With {.Dock = DockStyle.Bottom, .Height = 26, .Font = Fonts.Body, .Text = Loc.T("symptoms.other") & " " & r.Symptoms_Notes}
            LabelStyle.SetMuted(notes)
            card.Controls.Add(notes)
        End If
        Return card
    End Function

    Private Function ReadingCard(r As ReadEZResult, b As ReadEZResult) As Control
        Dim card = NewCard("results.reading", 200)
        Dim v As New ReadingSummaryView With {.Dock = DockStyle.Fill, .Result = r, .Baseline = b}
        card.Controls.Add(v)
        If r.ROR_Transcript.Trim() <> "" Then
            Dim t = ModernButton.Create("results.transcript", ButtonKind.Ghost, Glyphs.List)
            t.Dock = DockStyle.Bottom
            AddHandler t.Click, Sub() ShowTranscript(r)
            card.Controls.Add(t)
        End If
        Return card
    End Function

    Private Sub ShowTranscript(r As ReadEZResult)
        Dim text = ReadingTexts.ByIndex(r.RoR_TextOptions)
        Dim rows = ReadingTexts.RowsFor(text, r.RoR_NoRows, False)
        Dim marks = r.ROR_Transcript.Split(ChrW(10))
        Dim lines As New List(Of String)
        For i = 0 To rows.Count - 1
            lines.Add(rows(i))
            lines.Add(If(i < marks.Length, marks(i), ""))
        Next
        Using d As New TranscriptDialog(String.Join(vbCrLf, lines))
            d.ShowDialog(FindForm())
        End Using
    End Sub

    ''' <summary>Shows the text in the test's colour on the patient's screen (to try it again).</summary>
    Private Sub ShowOnPatientScreen()
        If _current Is Nothing Then Return
        Dim r = _current
        Dim scene As New ReadingScene With {.FontName = r.RoR_FontName, .FontSize = r.RoR_FontSize, .LineSpacing = r.RoR_LineSpacing,
                                            .OverlaySize = If(r.RoR_OverlaySize > 0, r.RoR_OverlaySize, AppHost.Settings.Overlay_Size),
                                            .Colour1 = r.ScreenColour,
                                            .Presentation = If(r.IsBaseline, ReadingPresentation.White, If(r.Mode = ReadEZTestMode.Overlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen))}
        scene.Rows = ReadingTexts.RowsFor(ReadingTexts.ByIndex(r.RoR_TextOptions), r.RoR_NoRows, AppHost.Settings.RoR_Random = 1)
        Stimulus.Window.BackColor = If(scene.Presentation = ReadingPresentation.FullScreen, scene.Colour1, Color.White)
        Stimulus.Window.Scene = scene
        Toast.Show(FindForm(), Loc.T(If(ScreenLayout.IsTwoScreenMode(), "results.shown", "results.shownOne")), ToastKind.Info)
        If Not ScreenLayout.IsTwoScreenMode() Then
            Using d As New PatientPreviewDialog()
                d.ShowDialog(FindForm())
            End Using
            Stimulus.ShowLogo()
        End If
    End Sub

    Public Overrides Sub OnDeactivated()
        Stimulus.ShowLogo()
        MyBase.OnDeactivated()
    End Sub

#End Region

End Class

''' <summary>A large colour swatch: a panel round text for overlays, a filled square for lenses / computer.</summary>
Public NotInheritable Class ColourSwatch
    Inherits Control
    Implements IThemed

    Public Property Colour As Color = Color.White
    Public Property Mode As ReadEZTestMode

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim r As New RectangleF(1, 1, Width - 3, Height - 3)
        Using path = Draw.RoundedRect(r, 10), b As New SolidBrush(If(Mode = ReadEZTestMode.Overlay, Color.White, Colour)), pen As New Pen(p.Border)
            g.FillPath(b, path)
            If Mode = ReadEZTestMode.Overlay Then
                Using ob As New SolidBrush(Colour)
                    g.FillRectangle(ob, RectangleF.Inflate(r, -14, -14))
                End Using
            End If
            g.DrawPath(pen, path)
        End Using
        ' A few lines of "text"
        Using tp As New Pen(Color.FromArgb(90, 40, 40, 40), 3)
            For i = 0 To 3
                Dim y = r.Y + 30 + i * 14
                g.DrawLine(tp, r.X + 24, y, r.Right - 24 - (i Mod 2) * 14, y)
            Next
        End Using
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

''' <summary>Key results as tinted panels with a coloured edge (as the report's key results).</summary>
Public NotInheritable Class KeyResultsView
    Inherits Control
    Implements IThemed

    Public NotInheritable Class Item
        Public Label, Value, Note As String
        Public Level As Brand.Level
        Public Sub New(label As String, value As String, note As String, level As Brand.Level)
            Me.Label = label : Me.Value = value : Me.Note = note : Me.Level = level
        End Sub
    End Class

    Public ReadOnly Property Items As New List(Of Item)

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        If Items.Count = 0 Then Return
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        Dim w = (Width - (Items.Count - 1) * 12.0F) / Items.Count
        For i = 0 To Items.Count - 1
            Dim it = Items(i)
            Brand.DrawKeyResult(g, New RectangleF(i * (w + 12), 0, w, Height - 4), it.Label, it.Value, it.Note, it.Level, ThemeManager.Current.Ink, ThemeManager.Current.Muted, ThemeManager.Current.Surface, 18.0F)
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

''' <summary>The four symptom grades as bars (0 - 5), beside the baseline's grades for a colour test.</summary>
Public NotInheritable Class SymptomsChart
    Inherits Control
    Implements IThemed

    Public Property Result As ReadEZResult
    Public Property Baseline As ReadEZResult
    ''' <summary>Draw for print (white background, fixed colours).</summary>
    Public Property ForPrint As Boolean

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Render(e.Graphics, ClientRectangle)
    End Sub

    Public Sub Render(g As Graphics, bounds As Rectangle)
        Dim p = ThemeManager.Current
        Dim ink = If(ForPrint, Brand.Charcoal, p.Ink), muted = If(ForPrint, Brand.Grey, p.Muted), grid = If(ForPrint, Brand.TintBorder, p.GridLine)
        g.FillRectangle(New SolidBrush(If(ForPrint, Color.White, Draw.ParentBack(Me))), bounds)
        If Result Is Nothing Then Return
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        If Not Result.SymptomsDone Then
            TextRenderer.DrawText(g, Loc.T("results.notDone"), Fonts.Body, bounds, muted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        Dim compare = Baseline IsNot Nothing AndAlso Baseline.SymptomsDone AndAlso Not Result.IsBaseline
        Dim labelW = CInt(bounds.Width * 0.42)
        Dim rowH = (bounds.Height - 30) / 4.0F
        Dim x0 = bounds.X + labelW + 10, barMax = bounds.Width - labelW - 60
        ' Scale 0 - 5
        Using gp As New Pen(grid)
            For v = 0 To 5
                Dim x = x0 + barMax * v / 5.0F
                g.DrawLine(gp, x, bounds.Y, x, bounds.Bottom - 26)
                TextRenderer.DrawText(g, v.ToString(), Fonts.Small, New Rectangle(CInt(x) - 10, bounds.Bottom - 24, 20, 18), muted, TextFormatFlags.HorizontalCenter)
            Next
        End Using
        For q = 0 To 3
            Dim y = bounds.Y + q * rowH
            TextRenderer.DrawText(g, Loc.T("symptoms.short" & (q + 1)), Fonts.Body, New Rectangle(bounds.X, CInt(y), labelW, CInt(rowH)), ink,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.WordBreak)
            Dim barH = If(compare, rowH * 0.3F, rowH * 0.45F)
            Dim by = y + (rowH - If(compare, barH * 2 + 4, barH)) / 2
            If compare Then
                DrawBar(g, x0, by, barMax * Baseline.Symptoms(q) / 5.0F, barH, Color.FromArgb(190, 190, 190), Baseline.Symptoms(q), muted)
                by += barH + 4
            End If
            DrawBar(g, x0, by, barMax * Result.Symptoms(q) / 5.0F, barH, If(Result.IsBaseline, Color.FromArgb(150, 160, 175), Darker(Result.ScreenColour)), Result.Symptoms(q), ink)
        Next
        If compare Then
            Dim lx = bounds.Right - 230
            TextRenderer.DrawText(g, "■ " & Loc.T("results.legend.baseline") & "   ", Fonts.Small, New Point(lx, bounds.Bottom - 24), Color.FromArgb(150, 150, 150))
            TextRenderer.DrawText(g, "■ " & Loc.T("results.legend.colour"), Fonts.Small, New Point(lx + 110, bounds.Bottom - 24), Darker(Result.ScreenColour))
        End If
    End Sub

    ''' <summary>A screen colour dark enough to see as a bar on white.</summary>
    Public Shared Function Darker(c As Color) As Color
        Dim lab = OkLab.FromColor(c)
        lab.L = Math.Min(lab.L, 0.72)
        Dim k = If(lab.Chroma > 0.001, Math.Max(1.0, 0.09 / lab.Chroma), 1.0)
        lab.A *= k : lab.B *= k
        Return lab.ToColor()
    End Function

    Private Shared Sub DrawBar(g As Graphics, x As Single, y As Single, w As Single, h As Single, colour As Color, value As Integer, ink As Color)
        Using b As New SolidBrush(colour)
            If w > 0 Then g.FillRectangle(b, x, y, w, h)
        End Using
        TextRenderer.DrawText(g, value.ToString(), Fonts.SmallStrong, New Point(CInt(x + w + 6), CInt(y + h / 2 - 8)), ink)
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

''' <summary>Rate of reading: big words-per-minute figures (baseline and colour) and the counts.</summary>
Public NotInheritable Class ReadingSummaryView
    Inherits Control
    Implements IThemed

    Public Property Result As ReadEZResult
    Public Property Baseline As ReadEZResult

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        If Result Is Nothing Then Return
        If Not Result.HasReading Then
            TextRenderer.DrawText(g, Loc.T("results.notDone"), Fonts.Body, ClientRectangle, p.Muted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        Dim cols As New List(Of (Title As String, R As ReadEZResult))
        If Baseline IsNot Nothing AndAlso Baseline.HasReading AndAlso Not Result.IsBaseline Then cols.Add((Loc.T("results.legend.baseline"), Baseline))
        cols.Add((If(Result.IsBaseline, Loc.T("result.white"), Result.ColourName), Result))
        Dim w = Width / CSng(Math.Max(2, cols.Count))
        For i = 0 To cols.Count - 1
            Dim c = cols(i)
            Dim x = CInt(i * w)
            TextRenderer.DrawText(g, c.Title, Fonts.Strong, New Point(x, 0), p.Muted)
            TextRenderer.DrawText(g, c.R.ROR.ToString(), Fonts.Display, New Point(x, 20), p.Ink)
            Dim numW = TextRenderer.MeasureText(c.R.ROR.ToString(), Fonts.Display).Width
            TextRenderer.DrawText(g, Loc.T("reading.wpmUnit"), Fonts.Body, New Point(x + numW, 42), p.Muted)
            Dim detail = Loc.F("results.readingDetail", c.R.ROR_NoWordsRead, c.R.ROR_RecordTime, c.R.ROR_NoErrors, c.R.ROR_RowsSkipped)
            TextRenderer.DrawText(g, detail, Fonts.Body, New Rectangle(x, 78, CInt(w) - 10, 60), p.Muted, TextFormatFlags.WordBreak)
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

''' <summary>The transcript of a rate-of-reading test (each row, with its marks underneath).</summary>
Public NotInheritable Class TranscriptDialog
    Inherits DialogBase

    Public Sub New(text As String)
        SetTitle("results.transcript", "results.transcript.subtitle")
        ClientSize = New Size(820, 640)
        Dim box As New TextBox With {.Dock = DockStyle.Fill, .Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Both, .WordWrap = False,
                                     .Font = New Font("Consolas", 10.5F), .Text = text}
        Body.Controls.Add(box)
        Dim close = AddFooterButton("common.close", ButtonKind.Primary)
        close.DialogResult = DialogResult.OK
        AcceptButton = close
        CancelButton = close
    End Sub
End Class

''' <summary>One screen: the patient's screen shown large in a window (Results > Show on the patient's screen).</summary>
Public NotInheritable Class PatientPreviewDialog
    Inherits Form

    Public Sub New()
        Text = AppHost.ProductTitle & " - " & Loc.T("stimulus.title")
        Icon = AppHost.AppIcon
        FormBorderStyle = FormBorderStyle.None
        WindowState = FormWindowState.Maximized
        StartPosition = FormStartPosition.Manual
        Bounds = AppHost.Settings.StimulusBounds
        KeyPreview = True
        SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw, True)
        AddHandler KeyDown, Sub() Close()
        AddHandler MouseClick, Sub() Close()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Stimulus.Window.RenderPreview(e.Graphics, ClientRectangle)
        TextRenderer.DrawText(e.Graphics, Loc.T("results.closePreview"), Fonts.Small, New Point(12, ClientSize.Height - 24), Color.Gray)
    End Sub
End Class
