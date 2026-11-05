Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   ResultsScreen.vb - a patient's results (replaces RZfrmResults' Summary panel and RZfrmResultsDetail)
'
'     Left    Baseline results: the patient's baselines, newest first (the newest is selected)
'             Results with colour: tabs Overlays / Lenses / Computer screen, each listing those
'             tests, newest first
'     Right   one page, no scrolling, as ReadEZ 2020's Summary:
'               1  the selected baseline      test information | symptoms score | reading performance
'               2  the selected colour test   test information (and the automatic test's colour
'                                             wheel) | symptoms score | reading performance
'               3  change with colour         each symptom on a worse -5 ... +5 better scale |
'                                             change in reading speed, errors and rows skipped
'     Header  Report, Letter and Show on the patient's screen use the colour test (with the selected
'             baseline), or the baseline when there is no colour test; Notes and Delete act on the
'             test last clicked in the lists.
' ====================================================================================================

Public Class ResultsScreen
    Inherits ScreenBase

    Private Shared ReadOnly ColourModes As ReadEZTestMode() = {ReadEZTestMode.Overlay, ReadEZTestMode.Lenses, ReadEZTestMode.Computer}

    Private ReadOnly _baseList As New ResultsList
    Private ReadOnly _colourList As New ResultsList
    Private ReadOnly segMode As New SegmentedControl
    Private ReadOnly lblNoBaseline As New Label With {.Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter}
    Private ReadOnly lblNoColour As New Label With {.Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter}
    Private ReadOnly _right As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 3, .BackColor = Color.Transparent, .Margin = New Padding(0), .Padding = New Padding(0)}
    Private ReadOnly cardBase As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(18, 48, 18, 12), .Margin = New Padding(0, 0, 0, 10)}
    Private ReadOnly cardColour As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(18, 48, 18, 12), .Margin = New Padding(0, 0, 0, 10)}
    Private ReadOnly cardChange As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(18, 48, 18, 12), .Margin = New Padding(0)}
    Private ReadOnly viewBase As New ResultPanelView With {.Dock = DockStyle.Fill, .EmptyKey = "results.noBaseline"}
    Private ReadOnly viewColour As New ResultPanelView With {.Dock = DockStyle.Fill, .EmptyKey = "results.noColourTests"}
    Private ReadOnly viewChange As New ChangePanelView With {.Dock = DockStyle.Fill}
    Private ReadOnly lblEmpty As New Label With {.Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter}
    Private _results As New List(Of ReadEZResult)
    Private _baselines As New List(Of ReadEZResult)
    Private _colourTests As New List(Of ReadEZResult)
    Private _baseline As ReadEZResult              ' the baseline selected on the left
    Private _colour As ReadEZResult                ' the colour test selected on the left
    Private _current As ReadEZResult               ' the test last clicked (Notes, Delete)
    Private _patientId As Integer = -1
    Private _loading As Boolean
    Private btnReport, btnLetter, btnShow, btnNotes, btnDelete As ModernButton
    Private ReadOnly btnEditBase, btnEditColour As ModernButton

    Public Sub New()
        TitleKey = "results.title"
        HelpTopic = "screen.results"
        UseProminentSubtitle()                            ' the patient's name, large, beside the title
        btnDelete = AddHeaderButton("common.delete", ButtonKind.Ghost, Glyphs.Delete)
        btnNotes = AddHeaderButton("results.notes", ButtonKind.Ghost, Glyphs.Edit)
        btnShow = AddHeaderButton("results.show", ButtonKind.Ghost, Glyphs.Eye)
        btnLetter = AddHeaderButton("results.letter", ButtonKind.Secondary, Glyphs.Report)
        btnReport = AddHeaderButton("results.report", ButtonKind.Primary, Glyphs.Print)
        AddHandler btnReport.Click, Sub()
                                        Dim r = ReportTarget()
                                        If r IsNot Nothing AndAlso r.IsPattern Then
                                            Ask.Info(FindForm(), "results.patternNoReport")
                                            Return
                                        End If
                                        If r IsNot Nothing Then ReadEZReport.ShowReport(FindForm(), r, BaselineFor(r))
                                    End Sub
        AddHandler btnLetter.Click, Sub()
                                        Dim r = ReportTarget()
                                        If r IsNot Nothing AndAlso r.IsPattern Then
                                            Ask.Info(FindForm(), "results.patternNoReport")
                                            Return
                                        End If
                                        If r IsNot Nothing Then ReadEZReport.ShowLetter(FindForm(), r, BaselineFor(r))
                                    End Sub
        AddHandler btnShow.Click, Sub() ShowOnPatientScreen()
        AddHandler btnNotes.Click, Sub()
                                       SelectCurrent()
                                       AppCommands.EditResult()
                                       Reload()
                                   End Sub
        AddHandler btnDelete.Click, Sub()
                                        SelectCurrent()
                                        AppCommands.DeleteResult()
                                        Reload()
                                    End Sub

        Dim split As New SplitContainer With {.Dock = DockStyle.Fill, .SplitterWidth = 12, .FixedPanel = FixedPanel.Panel1, .BackColor = Color.Transparent}
        ThemeManager.SetRole(split, SurfaceRole.Canvas)
        Dim placed = False
        AddHandler split.SizeChanged, Sub()
                                          If placed OrElse split.Width < 400 Then Return
                                          placed = True
                                          split.SplitterDistance = Math.Min(380, CInt(split.Width * 0.24))
                                      End Sub

        ' ---- Left: baseline results above, results with colour (by tabs) below ----
        Dim left As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .BackColor = Color.Transparent, .Margin = New Padding(0), .Padding = New Padding(0)}
        left.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        left.RowStyles.Add(New RowStyle(SizeType.Percent, 40))
        left.RowStyles.Add(New RowStyle(SizeType.Percent, 60))

        Dim baseCard As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(10, 48, 10, 10), .Margin = New Padding(0, 0, 0, 10)}
        baseCard.SetTitle("results.baselines")
        _baseList.Dock = DockStyle.Fill
        LabelStyle.SetMuted(lblNoBaseline)
        Loc.Bind(lblNoBaseline, "results.noBaseline")
        baseCard.Controls.Add(_baseList)
        baseCard.Controls.Add(lblNoBaseline)
        left.Controls.Add(baseCard, 0, 0)

        Dim colourCard As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(10, 48, 10, 10), .Margin = New Padding(0)}
        colourCard.SetTitle("results.colourTests")
        segMode.SetOptions("mode.overlay", "mode.lenses", "mode.computer")
        segMode.Dock = DockStyle.Top
        Dim gap As New Panel With {.Dock = DockStyle.Top, .Height = 8, .BackColor = Color.Transparent}
        _colourList.Dock = DockStyle.Fill
        LabelStyle.SetMuted(lblNoColour)
        Loc.Bind(lblNoColour, "results.noColourTests")
        colourCard.Controls.Add(_colourList)
        colourCard.Controls.Add(lblNoColour)
        colourCard.Controls.Add(gap)
        colourCard.Controls.Add(segMode)
        left.Controls.Add(colourCard, 0, 1)
        split.Panel1.Controls.Add(left)

        ' ---- Right: baseline / colour test / change with colour, filling the height ----
        _right.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        _right.RowStyles.Add(New RowStyle(SizeType.Percent, 32))
        _right.RowStyles.Add(New RowStyle(SizeType.Percent, 32))
        _right.RowStyles.Add(New RowStyle(SizeType.Percent, 36))   ' room for the summary
        cardBase.SetTitle("results.panel.baseline")
        cardColour.SetTitle("results.panel.colour")
        cardChange.SetTitle("results.panel.change")
        ' Edit the test selected in each list
        btnEditBase = AddEditButton(baseCard, Sub() EditTest(_baseline))
        btnEditColour = AddEditButton(colourCard, Sub() EditTest(_colour))
        cardBase.Controls.Add(viewBase)
        cardColour.Controls.Add(viewColour)
        cardChange.Controls.Add(viewChange)
        _right.Controls.Add(cardBase, 0, 0)
        _right.Controls.Add(cardColour, 0, 1)
        _right.Controls.Add(cardChange, 0, 2)
        LabelStyle.SetMuted(lblEmpty)
        Loc.Bind(lblEmpty, "results.none")
        split.Panel2.Controls.Add(_right)
        split.Panel2.Controls.Add(lblEmpty)
        Content.Controls.Add(split)

        AddHandler viewBase.TranscriptClicked, Sub() If _baseline IsNot Nothing Then ShowTranscript(_baseline)
        AddHandler viewColour.TranscriptClicked, Sub() If _colour IsNot Nothing Then ShowTranscript(_colour)
        AddHandler _baseList.ResultSelected, Sub()
                                                 If _loading Then Return
                                                 _baseline = _baseList.SelectedResult
                                                 _current = _baseline
                                                 ShowPanels()
                                             End Sub
        AddHandler _colourList.ResultSelected, Sub()
                                                   If _loading Then Return
                                                   _colour = _colourList.SelectedResult
                                                   _current = _colour
                                                   MatchBaseline()
                                                   ShowPanels()
                                               End Sub
        ' Clicking the test already selected makes it the one Notes / Delete act on
        AddHandler _baseList.Click, Sub()
                                        If _loading OrElse _baseline Is Nothing Then Return
                                        _current = _baseline
                                        SelectCurrent()
                                    End Sub
        AddHandler _colourList.Click, Sub()
                                          If _loading OrElse _colour Is Nothing Then Return
                                          _current = _colour
                                          SelectCurrent()
                                      End Sub
        AddHandler segMode.SelectedIndexChanged, Sub()
                                                     If _loading Then Return
                                                     FillColourList(-1)
                                                     MatchBaseline()
                                                     _current = If(_colour, _baseline)
                                                     ShowPanels()
                                                 End Sub
        AddHandler AppCommands.DataChanged, Sub(c) If Visible Then Reload()
        Loc.OnChange(Me, Sub() If Visible Then ShowPanels())
    End Sub

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        Reload()
    End Sub

    ''' <summary>An "Edit" button at the right of a list's title (edits the test selected in that list).</summary>
    Private Shared Function AddEditButton(card As Card, onClick As Action) As ModernButton
        Dim b = ModernButton.Create("results.edit", ButtonKind.Secondary, Glyphs.Edit)
        b.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        b.TabStop = False
        AddHandler b.Click, Sub() onClick()
        card.Controls.Add(b)
        b.BringToFront()
        Dim place = Sub()
                        Dim head = Dpi.Px(card, 40)
                        b.Location = New Point(card.Width - b.Width - Dpi.Px(card, 10), Math.Max(Dpi.Px(card, 3), (head - b.Height) \ 2))
                    End Sub
        AddHandler card.Resize, Sub() place()
        AddHandler b.SizeChanged, Sub() place()
        place()
        Return b
    End Function

    ''' <summary>Edits a test's symptoms, rate of reading and notes, then shows the change.</summary>
    Private Sub EditTest(r As ReadEZResult)
        If r Is Nothing Then Return
        _current = r
        SelectCurrent()
        Using d As New ResultEditDialog(r)
            If d.ShowDialog(FindForm()) <> DialogResult.OK Then Return
        End Using
        Toast.Show(FindForm(), Loc.T("results.edit.saved"), ToastKind.Success)
        AppCommands.NotifyDataChanged(False)          ' every screen showing the test reloads it (this one too)
    End Sub

#Region "Lists"

    Private Sub Reload()
        Dim s = AppHost.Session
        SetSubtitleText(If(s.HasPatient, s.PatientName, Loc.T("top.noPatient")))
        _results = If(s.HasPatient, AppHost.Db.GetResults(s.PatientId), New List(Of ReadEZResult))
        _baselines = _results.Where(Function(x) x.IsBaseline).OrderByDescending(Function(x) x.Test_Date).ThenByDescending(Function(x) x.ID).ToList()
        _colourTests = _results.Where(Function(x) Not x.IsBaseline).OrderByDescending(Function(x) x.Test_Date).ThenByDescending(Function(x) x.ID).ToList()
        Dim newPatient = Not s.HasPatient OrElse s.PatientId <> _patientId
        _patientId = If(s.HasPatient, s.PatientId, -1)
        Dim chosen = _results.FirstOrDefault(Function(x) x.ID = s.ResultId)
        Dim keep = Function(old As ReadEZResult) If(newPatient OrElse old Is Nothing, -1, old.ID)

        _loading = True
        Try
            ' Baseline: the one chosen elsewhere (e.g. just tested), else the one already selected, else the newest
            Dim baseId = If(chosen IsNot Nothing AndAlso chosen.IsBaseline, chosen.ID, keep(_baseline))
            If Not _baselines.Any(Function(b) b.ID = baseId) Then baseId = If(_baselines.Count > 0, _baselines(0).ID, -1)
            _baseList.SetResults(_baselines, baseId)
            _baseline = _baseList.SelectedResult
            _baseList.Visible = _baselines.Count > 0
            lblNoBaseline.Visible = _baselines.Count = 0

            ' Colour: the tab of the test chosen elsewhere / already selected, else the first tab with tests
            Dim colourPick = If(chosen IsNot Nothing AndAlso Not chosen.IsBaseline, chosen,
                                If(newPatient, Nothing, _colourTests.FirstOrDefault(Function(x) _colour IsNot Nothing AndAlso x.ID = _colour.ID)))
            Dim mode = ReadEZTestMode.Overlay
            If colourPick IsNot Nothing Then
                mode = colourPick.Mode
            Else
                For Each m In ColourModes
                    Dim mm = m
                    If _colourTests.Any(Function(x) x.Mode = mm) Then mode = mm : Exit For
                Next
            End If
            segMode.SelectedIndex = Math.Max(0, Array.IndexOf(ColourModes, mode))
            FillColourList(If(colourPick Is Nothing, -1, colourPick.ID))
            If chosen Is Nothing OrElse Not chosen.IsBaseline Then MatchBaseline()
        Finally
            _loading = False
        End Try
        _current = If(chosen IsNot Nothing AndAlso (chosen Is _baseline OrElse chosen Is _colour), chosen, If(_colour, _baseline))
        lblEmpty.Visible = _results.Count = 0
        _right.Visible = _results.Count > 0
        ShowPanels()
    End Sub

    ''' <summary>The colour tests of the selected tab (Overlays / Lenses / Computer screen), newest first.</summary>
    Private Sub FillColourList(selectId As Integer)
        Dim mode = ColourModes(Math.Max(0, Math.Min(ColourModes.Length - 1, segMode.SelectedIndex)))
        Dim list = _colourTests.Where(Function(x) x.Mode = mode).ToList()
        Dim was = _loading
        _loading = True
        Try
            _colourList.SetResults(list, selectId)
        Finally
            _loading = was
        End Try
        _colour = _colourList.SelectedResult
        _colourList.Visible = list.Count > 0
        lblNoColour.Visible = list.Count = 0
    End Sub

    ''' <summary>
    ''' A colour test is compared with a baseline of the same kind (a text test with a baseline text, a
    ''' pattern test with a baseline pattern): if the baseline selected is the other kind, the latest
    ''' one of the right kind (before the test, else the newest) is selected instead.
    ''' </summary>
    Private Sub MatchBaseline()
        If _colour Is Nothing Then Return
        If _baseline IsNot Nothing AndAlso _baseline.IsPattern = _colour.IsPattern Then Return
        Dim c = _colour
        Dim kind = _baselines.Where(Function(b) b.IsPattern = c.IsPattern).ToList()
        If kind.Count = 0 Then Return
        Dim pick = If(kind.FirstOrDefault(Function(b) b.Test_Date <= c.Test_Date), kind(0))
        Dim was = _loading
        _loading = True
        Try
            _baseList.SetResults(_baselines, pick.ID)
        Finally
            _loading = was
        End Try
        _baseline = _baseList.SelectedResult
    End Sub

    ''' <summary>The test the report, letter and "show" use: the colour test, or the baseline if there is none.</summary>
    Private Function ReportTarget() As ReadEZResult
        Return If(_colour, _baseline)
    End Function

    ''' <summary>The baseline to compare a colour test with: the one selected, else the last one before it, else the newest.</summary>
    Private Function BaselineFor(r As ReadEZResult) As ReadEZResult
        If r Is Nothing OrElse r.IsBaseline Then Return Nothing
        If _baseline IsNot Nothing AndAlso _baseline.IsPattern = r.IsPattern Then Return _baseline
        Dim kind = _baselines.Where(Function(b) b.IsPattern = r.IsPattern).ToList()
        Return If(kind.Where(Function(b) b.Test_Date <= r.Test_Date).FirstOrDefault(), kind.FirstOrDefault())
    End Function

#End Region

#Region "Panels"

    Private Sub ShowPanels()
        viewBase.Result = _baseline
        viewColour.Result = _colour
        viewChange.SetTests(_baseline, _colour)
        cardBase.TitleSuffix = If(_baseline Is Nothing, "", ":   " & Loc.T(TestRun.ResultKey(_baseline)) & "  ·  " & _baseline.Test_Date.ToString("g", Loc.Culture))
        cardColour.TitleSuffix = If(_colour Is Nothing, "", ":   " & Loc.T(TestRun.ResultKey(_colour)) & "  ·  " & _colour.Test_Date.ToString("g", Loc.Culture))
        cardBase.Invalidate()
        cardColour.Invalidate()
        btnEditBase.Visible = _baseline IsNot Nothing
        btnEditColour.Visible = _colour IsNot Nothing
        SelectCurrent()
    End Sub

    ''' <summary>Makes the test last clicked the session's test (Notes and Delete act on it) and updates the buttons.</summary>
    Private Sub SelectCurrent()
        If _current IsNot Nothing Then AppHost.Session.SelectResult(_current.ID)
        Dim target = ReportTarget()
        btnReport.Enabled = target IsNot Nothing
        btnLetter.Enabled = target IsNot Nothing
        btnShow.Enabled = target IsNot Nothing
        btnNotes.Enabled = _current IsNot Nothing
        btnDelete.Enabled = _current IsNot Nothing
    End Sub

    ''' <summary>The change with colour as key results (used by the report).</summary>
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
        Dim r = ReportTarget()
        If r Is Nothing Then Return
        Dim scene As New ReadingScene With {.FontName = r.RoR_FontName, .FontSize = r.RoR_FontSize, .LineSpacing = r.RoR_LineSpacing,
                                            .OverlaySize = If(r.RoR_OverlaySize > 0, r.RoR_OverlaySize, AppHost.Settings.Overlay_Size),
                                            .Colour1 = r.ScreenColour,
                                            .Presentation = If(r.IsBaseline, ReadingPresentation.White, If(r.Mode = ReadEZTestMode.Overlay, ReadingPresentation.Overlay, ReadingPresentation.FullScreen))}
        scene.Rows = ReadingTexts.RowsFor(ReadingTexts.ByIndex(r.RoR_TextOptions), r.RoR_NoRows, AppHost.Settings.RoR_Random = 1)
        scene.PatternStripe = PatternGlareStep.StripeOf(r)           ' a pattern test: its pattern
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

''' <summary>
''' One test on the Results page, laid out as ReadEZ 2020's summary:
'''   Test information (the patient's screen as it looked, the colour, the text settings)
'''   [the automatic colour test's wheel]  |  Symptoms score (each question 5 ... 0, the grade boxed)
'''   |  Reading performance (speed, errors, rows skipped, words read) and a Transcript button.
''' Sizes follow the panel, so the page never needs to scroll.
''' </summary>
Public NotInheritable Class ResultPanelView
    Inherits Control
    Implements IThemed

    Private _result As ReadEZResult
    Private _screener As String = ""
    Private ReadOnly _wheel As New ColourWheel With {.Selectable = False, .Visible = False}
    Private ReadOnly _transcript As ModernButton = ModernButton.Create("results.transcript", ButtonKind.Secondary, Glyphs.List)

    ''' <summary>Translation key of the text shown when there is no test.</summary>
    Public Property EmptyKey As String = "results.notDone"
    Public Event TranscriptClicked As EventHandler

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        _transcript.Visible = False
        AddHandler _transcript.Click, Sub() RaiseEvent TranscriptClicked(Me, EventArgs.Empty)
        Controls.Add(_wheel)
        Controls.Add(_transcript)
    End Sub

    Public Property Result As ReadEZResult
        Get
            Return _result
        End Get
        Set(value As ReadEZResult)
            _result = value
            _screener = If(value Is Nothing, "", AppHost.Db.ScreenerName(value.Screener_ID))
            Dim r = value
            Dim wheel = r IsNot Nothing AndAlso Not r.IsBaseline AndAlso r.Screening_Mode = 0 AndAlso r.AutoScore <> ""
            If wheel Then
                _wheel.OuterSet = If(r.Mode = ReadEZTestMode.Overlay, ReadEZSet.Overlay, ReadEZSet.Lens)
                _wheel.TwoRings = r.Mode <> ReadEZTestMode.Overlay
                _wheel.Spokes = ColourStaircase.ParseScore(r.AutoScore)
                _wheel.Selected = (r.ColourSet, r.Colour1)
            End If
            _wheel.Visible = wheel
            _transcript.Visible = r IsNot Nothing AndAlso r.HasReading AndAlso r.ROR_Transcript.Trim() <> ""
            PerformLayout()
            _wheel.Invalidate()
            Invalidate()
        End Set
    End Property

    Private Function D(v As Integer) As Integer
        Return LogicalToDeviceUnits(v)
    End Function

    ''' <summary>The columns: test information, the wheel (if any), symptoms, reading.</summary>
    Private Function Columns() As (Info As Rectangle, Wheel As Rectangle, Symptoms As Rectangle, Reading As Rectangle)
        Dim gap = D(24)
        ' Every panel has the same columns (test information | wheel | symptoms | reading), so the
        ' baseline and the colour test line up exactly; the baseline simply has no wheel in its slot.
        Dim avail = Math.Max(10, Width - 3 * gap)
        Dim sympW = CInt(avail * 0.34)
        Dim readW = CInt(avail * 0.22)
        Dim wheelW = Math.Min(Height, CInt(avail * 0.14))
        Dim infoW = Math.Max(10, avail - sympW - readW - wheelW)
        Dim info As New Rectangle(0, 0, infoW, Height)
        Dim wh As New Rectangle(info.Right + gap, (Height - wheelW) \ 2, wheelW, wheelW)
        Dim symp As New Rectangle(info.Right + gap + wheelW + gap, 0, sympW, Height)
        Dim read As New Rectangle(symp.Right + gap, 0, readW, Height)
        Return (info, wh, symp, read)
    End Function

    Protected Overrides Sub OnLayout(e As LayoutEventArgs)
        MyBase.OnLayout(e)
        Dim c = Columns()
        If _wheel.Visible Then _wheel.Bounds = c.Wheel
        If _transcript.Visible Then _transcript.Location = New Point(c.Reading.X, Math.Max(0, Height - _transcript.Height))
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        PerformLayout()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        If _result Is Nothing Then
            TextRenderer.DrawText(g, Loc.T(EmptyKey), Fonts.Body, ClientRectangle, p.Muted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        Dim c = Columns()
        PaintInfo(g, c.Info, p)
        PaintSymptoms(g, c.Symptoms, p)
        PaintReading(g, c.Reading, p)
    End Sub

    ''' <summary>The first of the fonts that fits the text in the box (the last one if none does).</summary>
    Friend Shared Function FitFont(g As Graphics, text As String, box As Size, ParamArray fonts As Font()) As Font
        For Each f In fonts
            Dim sz = TextRenderer.MeasureText(g, text, f, Size.Empty, TextFormatFlags.NoPadding)
            If sz.Width <= box.Width AndAlso sz.Height <= box.Height Then Return f
        Next
        Return fonts(fonts.Length - 1)
    End Function

    ''' <summary>
    ''' A row's label: on one line if it fits, on two if the row is tall enough, otherwise in the
    ''' small font (cut short with "..." only as a last resort) - never clipped half-way through a line.
    ''' </summary>
    Friend Shared Sub DrawRowLabel(g As Graphics, text As String, r As Rectangle, colour As Color)
        Dim one = TextRenderer.MeasureText(g, text, Fonts.Body, Size.Empty, TextFormatFlags.NoPadding)
        Dim flags = TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding
        If one.Width <= r.Width Then
            TextRenderer.DrawText(g, text, Fonts.Body, r, colour, flags)
        ElseIf r.Height >= one.Height * 2 + 2 Then
            TextRenderer.DrawText(g, text, Fonts.Body, r, colour, flags Or TextFormatFlags.WordBreak)
        Else
            Dim small = TextRenderer.MeasureText(g, text, Fonts.Small, Size.Empty, TextFormatFlags.NoPadding)
            If small.Width <= r.Width OrElse r.Height < small.Height * 2 + 2 Then
                TextRenderer.DrawText(g, text, Fonts.Small, r, colour, flags Or TextFormatFlags.EndEllipsis)
            Else
                TextRenderer.DrawText(g, text, Fonts.Small, r, colour, flags Or TextFormatFlags.WordBreak)
            End If
        End If
    End Sub

    ''' <summary>A small heading at the top of a column.</summary>
    Private Function Heading(g As Graphics, r As Rectangle, key As String, p As Palette) As Integer
        Dim h = D(26)
        TextRenderer.DrawText(g, Loc.T(key).ToUpper(Loc.Culture), Fonts.Strong, New Rectangle(r.X, r.Y, r.Width, h), p.Primary,
                              TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        Return r.Y + h + D(6)
    End Function

#Region "Test information"

    Private Sub PaintInfo(g As Graphics, col As Rectangle, p As Palette)
        Dim r = _result
        Dim top = Heading(g, col, "results.head.info", p)
        Dim avail = col.Bottom - top
        ' The patient's screen: a monitor 4:3 with a stand, the text as it was shown
        Dim chipH = D(34)
        Dim monW = Math.Min(CInt(col.Width * 0.42), CInt((avail - chipH - D(8)) * 4 / 3.0 / 1.18))
        monW = Math.Max(D(60), monW)
        Dim monH = CInt(monW * 0.75)
        Dim mon As New Rectangle(col.X, top, monW, monH)
        DrawMonitor(g, mon, r)
        ' The colour (or "White") under it
        Dim chip As New Rectangle(col.X, mon.Bottom + CInt(monH * 0.18) + D(8), monW, chipH)
        Using path = Draw.RoundedRect(New RectangleF(chip.X, chip.Y, chip.Width, chip.Height), D(6)),
              b As New SolidBrush(p.SurfaceAlt), pen As New Pen(p.Border)
            g.FillPath(b, path)
            g.DrawPath(pen, path)
        End Using
        Dim sw As New Rectangle(chip.X + D(8), chip.Y + (chipH - D(16)) \ 2, D(16), D(16))
        Using b As New SolidBrush(r.ScreenColour), pen As New Pen(p.Border)
            g.FillEllipse(b, sw)
            g.DrawEllipse(pen, sw)
        End Using
        TextRenderer.DrawText(g, If(r.IsBaseline, Loc.T("result.white"), r.ColourName), Fonts.Strong,
                              New Rectangle(sw.Right + D(6), chip.Y, chip.Right - sw.Right - D(10), chip.Height), p.Ink,
                              TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        ' The settings, beside the monitor
        Dim lines As New List(Of String)
        Dim quick = r.AutoScore.Contains("Method:Quick")
        If Not r.IsBaseline Then lines.Add(Loc.F("results.method", Loc.T(If(r.Screening_Mode = 0, If(quick, "colour.autoQuick", "colour.auto"), If(r.Screening_Mode = 1, "colour.manual", "results.bespoke")))))
        Dim changedAt = r.AutoScore.IndexOf("ChangedFrom:", StringComparison.Ordinal)
        If Not r.IsBaseline AndAlso changedAt >= 0 Then
            Dim was = r.AutoScore.Substring(changedAt + "ChangedFrom:".Length).Split(","c)(0)
            lines.Add(Loc.F("results.colourChanged", was))
        End If
        If r.IsPattern Then
            lines.Add(Loc.F("results.cfg.pattern", PatternGlareStep.Cpd(r.PatternCpd)))
            If r.PatternDistanceCm > 0 Then lines.Add(Loc.F("results.cfg.distance", r.PatternDistanceCm))
        Else
            If r.ROR_Text <> "" Then lines.Add(Loc.F("results.cfg.text", r.ROR_Text))
            lines.Add(Loc.F("results.cfg.font", r.RoR_FontName, r.RoR_FontSize + 8))
            lines.Add(Loc.F("results.cfg.spacing", (0.8 + r.RoR_LineSpacing / 10.0).ToString("0.0", Loc.Culture)))
            lines.Add(Loc.F("results.cfg.rows", r.RoR_NoRows))
        End If
        If _screener <> "" Then lines.Add(Loc.F("results.screener", _screener))
        If r.GreyPreferred Then lines.Add(Loc.T("colour.greyPreferred"))
        If r.ImportedFrom <> "" Then lines.Add(Loc.F("results.imported", r.ImportedFrom))
        Dim tx = mon.Right + D(16)
        Dim tw = Math.Max(10, col.Right - tx)
        Dim lineH = TextRenderer.MeasureText("Ag", Fonts.Body).Height + D(3)
        Dim y = top
        For Each l In lines
            If y + lineH > col.Bottom Then Exit For
            Dim warn = r.GreyPreferred AndAlso l = Loc.T("colour.greyPreferred")
            TextRenderer.DrawText(g, l, Fonts.Body, New Rectangle(tx, y, tw, lineH), If(warn, p.Warning, p.Muted),
                                  TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
            y += lineH
        Next
    End Sub

    ''' <summary>A monitor showing the text as it was presented (white, an overlay round the text, or the whole screen in colour).</summary>
    Private Sub DrawMonitor(g As Graphics, mon As Rectangle, r As ReadEZResult)
        Dim bezel = Color.FromArgb(45, 48, 54)
        Using path = Draw.RoundedRect(New RectangleF(mon.X, mon.Y, mon.Width, mon.Height), D(6)), b As New SolidBrush(bezel)
            g.FillPath(b, path)
        End Using
        ' Stand
        Dim neckW = mon.Width \ 8, neckH = CInt(mon.Height * 0.1)
        Using b As New SolidBrush(Color.FromArgb(150, 155, 162))
            g.FillRectangle(b, mon.X + (mon.Width - neckW) \ 2, mon.Bottom, neckW, neckH)
            Using path = Draw.RoundedRect(New RectangleF(mon.X + mon.Width * 0.32F, mon.Bottom + neckH, mon.Width * 0.36F, mon.Height * 0.06F), D(3))
                g.FillPath(b, path)
            End Using
        End Using
        Dim scr = Rectangle.Inflate(mon, -D(6), -D(6))
        Dim full = Not r.IsBaseline AndAlso r.Mode <> ReadEZTestMode.Overlay
        Using b As New SolidBrush(If(full, r.ScreenColour, Color.White))
            g.FillRectangle(b, scr)
        End Using
        ' The block of text in the middle, with the overlay round it
        Dim tb As New RectangleF(scr.X + scr.Width * 0.36F, scr.Y + scr.Height * 0.25F, scr.Width * 0.28F, scr.Height * 0.5F)
        If Not r.IsBaseline AndAlso r.Mode = ReadEZTestMode.Overlay Then
            Using b As New SolidBrush(r.ScreenColour)
                g.FillRectangle(b, RectangleF.Inflate(tb, scr.Width * 0.08F, scr.Height * 0.1F))
            End Using
        End If
        If r.IsPattern Then
            ' The pattern: a square of black stripes (the white ones are the background or the overlay)
            Dim side = Math.Min(tb.Width, tb.Height)
            Dim pb As New RectangleF(tb.X + (tb.Width - side) / 2, tb.Y + (tb.Height - side) / 2, side, side)
            Dim stripe = side / 10
            For i = 0 To 4
                g.FillRectangle(Brushes.Black, pb.X, pb.Y + i * 2 * stripe, pb.Width, stripe)
            Next
            Return
        End If
        Using pen As New Pen(Color.FromArgb(150, 60, 60, 60), Math.Max(1.0F, scr.Height / 60.0F))
            Dim rows = 9
            For i = 0 To rows - 1
                Dim y = tb.Y + i * tb.Height / (rows - 1)
                g.DrawLine(pen, tb.X, y, tb.Right - If(i Mod 3 = 2, tb.Width * 0.2F, 0), y)
            Next
        End Using
    End Sub

#End Region

#Region "Symptoms"

    Private Sub PaintSymptoms(g As Graphics, col As Rectangle, p As Palette)
        Dim r = _result
        Dim top = Heading(g, col, "results.head.symptoms", p)
        If Not r.SymptomsDone Then
            TextRenderer.DrawText(g, Loc.T("results.notDone"), Fonts.Section, New Rectangle(col.X, top, col.Width, col.Bottom - top), p.Muted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        Dim notesH = If(r.Symptoms_Notes <> "", D(24), 0)
        Dim scaleH = D(20)
        Dim labelW = CInt(col.Width * 0.46)
        Dim gap = D(7)
        Dim rowsTop = top + scaleH
        Dim rowH = (col.Bottom - rowsTop - notesH) / 4.0F
        Dim box = Math.Max(D(16), Math.Min(D(30), Math.Min(CInt(rowH) - D(10), (col.Width - labelW - D(10) - 5 * gap) \ 6)))
        Dim x0 = col.X + labelW + D(10)
        ' "Severe" over 5, "None" over 0
        TextRenderer.DrawText(g, Loc.T("results.scale.severe"), Fonts.Body, New Rectangle(x0 - D(20), top, box + D(40), scaleH), p.Muted,
                              TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
        Dim lastX = x0 + 5 * (box + gap)
        TextRenderer.DrawText(g, Loc.T("results.scale.none"), Fonts.Body, New Rectangle(lastX - D(20), top, box + D(40), scaleH), p.Muted,
                              TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
        For q = 0 To 3
            Dim y = rowsTop + q * rowH
            TextRenderer.DrawText(g, Loc.T(If(r.IsPattern, "symptoms.pshort", "symptoms.short") & (q + 1)), Fonts.Body, New Rectangle(col.X, CInt(y), labelW, CInt(rowH)), p.Ink,
                                  TextFormatFlags.Right Or TextFormatFlags.VerticalCenter Or TextFormatFlags.WordBreak)
            Dim by = y + (rowH - box) / 2
            For i = 0 To 5
                Dim grade = 5 - i
                Dim br As New RectangleF(x0 + i * (box + gap), by, box, box)
                PaintGradeBox(g, br, GradeScale.GradeColour(grade), grade.ToString(), grade = r.Symptoms(q), p)
            Next
        Next
        If notesH > 0 Then
            TextRenderer.DrawText(g, Loc.T("symptoms.other") & " " & r.Symptoms_Notes, Fonts.Body, New Rectangle(col.X, col.Bottom - notesH, col.Width, notesH), p.Muted,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        End If
    End Sub

    ''' <summary>A coloured box with its number; the chosen one is ringed (as ReadEZ 2020).</summary>
    Friend Shared Sub PaintGradeBox(g As Graphics, r As RectangleF, colour As Color, text As String, chosen As Boolean, p As Palette)
        PaintScaleBox(g, r, colour, text, chosen, p)
    End Sub

    ''' <summary>
    ''' A box of a colour scale (as ReadEZ 2020): the chosen one in full colour, a little larger, with a
    ''' thick black frame and a bold number; the others paler, so the chosen one stands out.
    ''' </summary>
    Friend Shared Sub PaintScaleBox(g As Graphics, r As RectangleF, colour As Color, text As String, chosen As Boolean, p As Palette)
        Dim box = If(chosen, RectangleF.Inflate(r, 2, 2), r)
        Dim radius = Math.Max(2.0F, box.Height / 7.0F)
        Dim fill = If(chosen, colour, Draw.Mix(colour, p.Surface, 0.45))
        Using path = Draw.RoundedRect(box, radius), b As New SolidBrush(fill),
              edge As New Pen(If(chosen, Color.Black, Draw.Mix(colour, Color.Black, 0.15)), If(chosen, 3.0F, 1.0F))
            g.FillPath(b, path)
            g.DrawPath(edge, path)
        End Using
        Using f As New Font(Fonts.Strong.FontFamily, Math.Max(10.0F, r.Height * If(chosen, 0.52F, 0.42F)), If(chosen, FontStyle.Bold, FontStyle.Regular), GraphicsUnit.Pixel)
            TextRenderer.DrawText(g, text, f, Rectangle.Round(box), If(chosen, Color.Black, Color.FromArgb(110, 110, 110)),
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
        End Using
    End Sub

#End Region

#Region "Reading"

    Private Sub PaintReading(g As Graphics, col As Rectangle, p As Palette)
        Dim r = _result
        Dim top = Heading(g, col, If(r.IsPattern, "results.head.pattern", "results.head.reading"), p)
        Dim bottom = If(_transcript.Visible, col.Bottom - _transcript.Height - D(8), col.Bottom)
        If r.IsPattern Then
            PaintRows(g, col, top, bottom, PatternRows(r), p)
            Return
        End If
        If Not r.HasReading Then
            TextRenderer.DrawText(g, Loc.T("results.notDone"), Fonts.Section, New Rectangle(col.X, top, col.Width, bottom - top), p.Muted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        Dim rows As New List(Of (Label As String, Value As String, Big As Boolean)) From {
            (Loc.T("results.row.wpm"), r.ROR.ToString(), True),
            (Loc.T("results.row.errors"), r.ROR_NoErrors.ToString(), False),
            (Loc.T("results.row.skipped"), r.ROR_RowsSkipped.ToString(), False),
            (Loc.T("results.row.wordsTime"), Loc.F("results.row.wordsTimeValue", r.ROR_NoWordsRead, r.ROR_RecordTime), False)}
        PaintRows(g, col, top, bottom, rows, p)
    End Sub

    ''' <summary>A pattern test's facts: the worst spatial frequency (and, for the baseline pattern, how it was found).</summary>
    Private Shared Function PatternRows(r As ReadEZResult) As List(Of (Label As String, Value As String, Big As Boolean))
        Dim rows As New List(Of (Label As String, Value As String, Big As Boolean)) From {
            (Loc.T("results.row.worstSf"), If(r.PatternCpd > 0, Loc.F("results.row.cpdValue", PatternGlareStep.Cpd(r.PatternCpd)), "–"), True)}
        If r.PatternDistanceCm > 0 Then rows.Add((Loc.T("results.row.distance"), Loc.F("results.row.cmValue", r.PatternDistanceCm), False))
        Dim found = PatternGlareStep.ParseScore(r.AutoScore)
        If found.Choices > 0 Then
            rows.Add((Loc.T("results.row.cpdRange"), PatternGlareStep.Cpd(found.Low) & " – " & PatternGlareStep.Cpd(found.High), False))
            rows.Add((Loc.T("results.row.choices"), found.Choices.ToString() & If(found.Completed, "", " " & Loc.T("results.row.stoppedEarly")), False))
        End If
        Return rows
    End Function

    ''' <summary>Rows of a label and a value in a box, filling the column.</summary>
    Private Sub PaintRows(g As Graphics, col As Rectangle, top As Integer, bottom As Integer, rows As List(Of (Label As String, Value As String, Big As Boolean)), p As Palette)
        If rows.Count = 0 Then Return
        Dim rowH = (bottom - top) / CSng(rows.Count)
        ' The value boxes as wide as the widest value needs (within reason), the labels get the rest
        Dim need = rows.Max(Function(x) TextRenderer.MeasureText(g, x.Value, Fonts.Section, Size.Empty, TextFormatFlags.NoPadding).Width) + D(18)
        Dim valW = Math.Max(D(64), Math.Min(need, CInt(col.Width * 0.5)))
        Dim valX = col.Right - valW
        For i = 0 To rows.Count - 1
            Dim row = rows(i)
            Dim y = top + i * rowH
            DrawRowLabel(g, row.Label, New Rectangle(col.X, CInt(y), valX - col.X - D(8), CInt(rowH)), p.Ink)
            Dim bh = Math.Max(D(18), Math.Min(CInt(rowH) - D(3), If(row.Big, D(42), D(32))))
            Dim vb As New RectangleF(valX, y + (rowH - bh) / 2, valW, bh)
            Using path = Draw.RoundedRect(vb, D(5)), b As New SolidBrush(p.SurfaceAlt), pen As New Pen(p.Border)
                g.FillPath(b, path)
                g.DrawPath(pen, path)
            End Using
            Dim inner As New Size(CInt(vb.Width) - D(6), CInt(vb.Height))
            Dim f = If(row.Big, FitFont(g, row.Value, inner, Fonts.Title, Fonts.Section, Fonts.Strong, Fonts.Body),
                                FitFont(g, row.Value, inner, Fonts.Section, Fonts.Strong, Fonts.Body, Fonts.Small))
            TextRenderer.DrawText(g, row.Value, f, Rectangle.Round(vb), p.Ink,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding Or TextFormatFlags.EndEllipsis)
        Next
    End Sub

#End Region

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return True
        End Get
    End Property
End Class

''' <summary>
''' Change with colour (ReadEZ 2020's "Change in symptoms and reading performance with colour"):
''' each symptom on a scale from -5 (worse) to +5 (better), the change ringed, and the change in
''' reading speed, errors and rows skipped. Change in a symptom = baseline grade - grade with colour.
''' </summary>
Public NotInheritable Class ChangePanelView
    Inherits Control
    Implements IThemed

    Private _baseline, _result As ReadEZResult

    ' ReadEZ 2020's colours: -5 ... -1 worse (the grade colours, orange to pale yellow), 0 no change
    ' (white), +1 ... +5 better (pale to darker green)
    Private Shared ReadOnly Scale As Color() = {
        GradeScale.GradeColour(5), GradeScale.GradeColour(4), GradeScale.GradeColour(3), GradeScale.GradeColour(2), GradeScale.GradeColour(1),
        Color.FromArgb(255, 255, 255),
        Color.FromArgb(226, 239, 218), Color.FromArgb(197, 224, 180), Color.FromArgb(169, 209, 142), Color.FromArgb(140, 192, 100), Color.FromArgb(112, 173, 71)}

    Public Shared Function ChangeColour(change As Integer) As Color
        Return Scale(Math.Max(-5, Math.Min(5, change)) + 5)
    End Function

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
    End Sub

    Public Sub SetTests(baseline As ReadEZResult, result As ReadEZResult)
        _baseline = baseline
        _result = result
        Invalidate()
    End Sub

    Private Function D(v As Integer) As Integer
        Return LogicalToDeviceUnits(v)
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        If _baseline Is Nothing OrElse _result Is Nothing Then
            TextRenderer.DrawText(g, Loc.T("results.change.needBoth"), Fonts.Body, ClientRectangle, p.Muted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.WordBreak)
            Return
        End If
        If _baseline.IsPattern <> _result.IsPattern Then
            TextRenderer.DrawText(g, Loc.T("results.change.mismatch"), Fonts.Body, ClientRectangle, p.Muted, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.WordBreak)
            Return
        End If
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        ' The summary across the top, the details underneath
        Dim summaryH = D(50)
        PaintSummary(g, New Rectangle(0, 0, Width, summaryH), p)
        Dim top = summaryH + D(12)
        If _result.IsPattern Then
            ' Pattern tests: symptoms with and without colour only
            PaintSymptoms(g, New Rectangle(0, top, Width, Height - top), p)
            Return
        End If
        Dim gap = D(24)
        Dim sympW = CInt((Width - gap) * 0.6)
        PaintSymptoms(g, New Rectangle(0, top, sympW, Height - top), p)
        PaintReading(g, New Rectangle(sympW + gap, top, Width - sympW - gap, Height - top), p)
    End Sub

    Private Enum Verdict
        Better
        Same
        Worse
        NotBoth
    End Enum

    ''' <summary>Symptoms: better if the total (0 - 20) is lower with colour, worse if higher.</summary>
    Private Function SymptomsVerdict() As Verdict
        If Not _baseline.SymptomsDone OrElse Not _result.SymptomsDone Then Return Verdict.NotBoth
        Dim fewer = _baseline.SymptomTotal - _result.SymptomTotal
        Return If(fewer > 0, Verdict.Better, If(fewer < 0, Verdict.Worse, Verdict.Same))
    End Function

    ''' <summary>Reading: better if 5 % or more faster with colour, worse if 5 % or more slower (as the report).</summary>
    Private Function ReadingVerdict(ByRef pct As Double) As Verdict
        pct = 0
        If Not _baseline.HasReading OrElse Not _result.HasReading OrElse _baseline.ROR <= 0 Then Return Verdict.NotBoth
        pct = 100.0 * (_result.ROR - _baseline.ROR) / _baseline.ROR
        Return If(pct >= 5, Verdict.Better, If(pct <= -5, Verdict.Worse, Verdict.Same))
    End Function

    ''' <summary>Two clear boxes: "Symptoms: Better with colour" and "Reading: Worse with colour" (or no change).</summary>
    Private Sub PaintSummary(g As Graphics, r As Rectangle, p As Palette)
        Dim gap = D(16)
        Dim w = (r.Width - gap) \ 2
        Dim pct As Double
        Dim sv = SymptomsVerdict()
        Dim rv = ReadingVerdict(pct)
        Dim symDetail = If(sv = Verdict.NotBoth, "", Loc.F("results.summary.symptomsDetail", _baseline.SymptomTotal, _result.SymptomTotal))
        If _result.IsPattern Then
            PaintVerdict(g, r, Loc.T("results.summary.symptoms"), sv, symDetail, p)    ' (no reading with the pattern)
            Return
        End If
        Dim readDetail = If(rv = Verdict.NotBoth, "", Loc.F("results.summary.readingDetail", _baseline.ROR, _result.ROR, If(pct >= 0, "+", "") & pct.ToString("0", Loc.Culture) & "%"))
        PaintVerdict(g, New Rectangle(r.X, r.Y, w, r.Height), Loc.T("results.summary.symptoms"), sv, symDetail, p)
        PaintVerdict(g, New Rectangle(r.X + w + gap, r.Y, r.Width - w - gap, r.Height), Loc.T("results.summary.reading"), rv, readDetail, p)
    End Sub

    Private Sub PaintVerdict(g As Graphics, r As Rectangle, label As String, v As Verdict, detail As String, p As Palette)
        Dim colour = If(v = Verdict.Better, p.Success, If(v = Verdict.Worse, p.Danger, p.Muted))
        Dim text = Loc.T(If(v = Verdict.Better, "results.summary.better", If(v = Verdict.Worse, "results.summary.worse",
                            If(v = Verdict.Same, "results.summary.same", "results.summary.notBoth"))))
        Dim symbol = If(v = Verdict.Better, ChrW(&H2713).ToString(), If(v = Verdict.Worse, ChrW(&H2715).ToString(), If(v = Verdict.Same, "=", "–")))
        Using path = Draw.RoundedRect(New RectangleF(r.X + 0.5F, r.Y + 0.5F, r.Width - 1, r.Height - 1), D(8)),
              fill As New SolidBrush(Draw.Mix(p.Surface, colour, 0.12)), edge As New Pen(Draw.Mix(p.Surface, colour, 0.55), 1.5F)
            g.FillPath(fill, path)
            g.DrawPath(edge, path)
        End Using
        ' A round badge with a tick, a cross or "="
        Dim badgeSize = r.Height - D(16)
        Dim badge As New Rectangle(r.X + D(10), r.Y + (r.Height - badgeSize) \ 2, badgeSize, badgeSize)
        Using b As New SolidBrush(colour)
            g.FillEllipse(b, badge)
        End Using
        TextRenderer.DrawText(g, symbol, Fonts.Section, badge, Color.White, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
        ' "Symptoms:"  "Better with colour"           detail at the right
        Dim x = badge.Right + D(12)
        Dim flags = TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding
        Dim lab = label & ":"
        Dim labW = TextRenderer.MeasureText(g, lab, Fonts.Section, Size.Empty, TextFormatFlags.NoPadding).Width
        TextRenderer.DrawText(g, lab, Fonts.Section, New Rectangle(x, r.Y, labW + 2, r.Height), p.Ink, flags)
        x += labW + D(8)
        Dim detailW = If(detail = "", 0, TextRenderer.MeasureText(g, detail, Fonts.Body, Size.Empty, TextFormatFlags.NoPadding).Width + D(14))
        TextRenderer.DrawText(g, text, Fonts.Section, New Rectangle(x, r.Y, Math.Max(10, r.Right - x - detailW - D(6)), r.Height),
                              If(v = Verdict.NotBoth, p.Muted, colour), flags Or TextFormatFlags.EndEllipsis)
        If detail <> "" Then
            TextRenderer.DrawText(g, detail, Fonts.Body, New Rectangle(r.Right - detailW, r.Y, detailW - D(10), r.Height), p.Muted,
                                  TextFormatFlags.Right Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
        End If
    End Sub

    Private Function Heading(g As Graphics, r As Rectangle, key As String, p As Palette) As Integer
        Dim h = D(26)
        TextRenderer.DrawText(g, Loc.T(key).ToUpper(Loc.Culture), Fonts.Strong, New Rectangle(r.X, r.Y, r.Width, h), p.Primary,
                              TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        Return r.Y + h + D(6)
    End Function

    Private Sub PaintSymptoms(g As Graphics, col As Rectangle, p As Palette)
        Dim b = _baseline, c = _result
        Dim top = Heading(g, col, "results.head.symptomsChange", p)
        If Not b.SymptomsDone OrElse Not c.SymptomsDone Then
            TextRenderer.DrawText(g, Loc.T("results.change.notBoth"), Fonts.Section, New Rectangle(col.X, top, col.Width, col.Bottom - top), p.Muted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        Dim scaleH = D(20)
        Dim labelW = CInt(col.Width * 0.32)
        Dim fromToW = D(70)
        Dim gap = D(7)
        Dim rowsTop = top + scaleH
        Dim rowH = (col.Bottom - rowsTop) / 4.0F
        Dim box = Math.Max(D(16), Math.Min(D(30), Math.Min(CInt(rowH) - D(10), (col.Width - labelW - fromToW - D(20) - 10 * gap) \ 11)))
        Dim x0 = col.X + labelW + D(10)
        Dim zeroX = x0 + 5 * (box + gap)
        TextRenderer.DrawText(g, Loc.T("results.change.worse").ToUpper(Loc.Culture), Fonts.SmallStrong, New Rectangle(x0, top, zeroX - x0 - gap, scaleH), p.Muted,
                              TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
        TextRenderer.DrawText(g, Loc.T("results.change.better").ToUpper(Loc.Culture), Fonts.SmallStrong, New Rectangle(zeroX + box + gap, top, 5 * (box + gap), scaleH), p.Muted,
                              TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
        Dim ftX = x0 + 11 * (box + gap) + D(8)
        TextRenderer.DrawText(g, Loc.T("results.change.fromToHead"), Fonts.Body, New Rectangle(ftX, top, col.Right - ftX, scaleH), p.Muted,
                              TextFormatFlags.Left Or TextFormatFlags.Bottom Or TextFormatFlags.EndEllipsis)
        For q = 0 To 3
            Dim y = rowsTop + q * rowH
            TextRenderer.DrawText(g, Loc.T(If(c.IsPattern, "symptoms.pshort", "symptoms.short") & (q + 1)), Fonts.Body, New Rectangle(col.X, CInt(y), labelW, CInt(rowH)), p.Ink,
                                  TextFormatFlags.Right Or TextFormatFlags.VerticalCenter Or TextFormatFlags.WordBreak)
            Dim change = b.Symptoms(q) - c.Symptoms(q)
            Dim by = y + (rowH - box) / 2
            For v = -5 To 5
                Dim br As New RectangleF(x0 + (v + 5) * (box + gap), by, box, box)
                Dim chosen = v = change
                Dim col2 = ChangeColour(v)
                ResultPanelView.PaintScaleBox(g, br, col2, If(v > 0, "+" & v, v.ToString()), chosen, p)
            Next
            TextRenderer.DrawText(g, $"{b.Symptoms(q)} → {c.Symptoms(q)}", Fonts.Body, New Rectangle(ftX, CInt(y), col.Right - ftX, CInt(rowH)), p.Muted,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter)
        Next
    End Sub

    Private Sub PaintReading(g As Graphics, col As Rectangle, p As Palette)
        Dim b = _baseline, c = _result
        Dim top = Heading(g, col, "results.head.readingChange", p)
        If Not b.HasReading OrElse Not c.HasReading Then
            TextRenderer.DrawText(g, Loc.T("results.change.notBoth"), Fonts.Section, New Rectangle(col.X, top, col.Width, col.Bottom - top), p.Muted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
            Return
        End If
        Dim dw = c.ROR - b.ROR
        Dim pct = If(b.ROR > 0, 100.0 * dw / b.ROR, 0.0)
        Dim de = c.ROR_NoErrors - b.ROR_NoErrors
        Dim ds = c.ROR_RowsSkipped - b.ROR_RowsSkipped
        Dim rows As New List(Of (Label As String, Value As String, Note As String, Colour As Color, Big As Boolean)) From {
            (Loc.T("results.row.wpm"), $"{Signed(dw)}  ({If(pct >= 0, "+", "")}{pct:0}%)", $"{b.ROR} → {c.ROR}", If(pct >= 5, p.Success, If(pct > -5, p.Warning, p.Danger)), True),
            (Loc.T("results.row.errors"), Signed(de), $"{b.ROR_NoErrors} → {c.ROR_NoErrors}", If(de < 0, p.Success, If(de = 0, p.Muted, p.Danger)), False),
            (Loc.T("results.row.skipped"), Signed(ds), $"{b.ROR_RowsSkipped} → {c.ROR_RowsSkipped}", If(ds < 0, p.Success, If(ds = 0, p.Muted, p.Danger)), False)}
        Dim noteH = 0
        Dim rowH = (col.Bottom - top - noteH) / CSng(rows.Count)
        ' The value boxes and the "from -> to" notes as wide as they need (within reason), the labels get the rest
        Dim need = rows.Max(Function(x) TextRenderer.MeasureText(g, x.Value, Fonts.Section, Size.Empty, TextFormatFlags.NoPadding).Width) + D(18)
        Dim valW = Math.Max(D(64), Math.Min(need, CInt(col.Width * 0.42)))
        Dim ftW = Math.Min(CInt(col.Width * 0.25), rows.Max(Function(x) TextRenderer.MeasureText(g, x.Note, Fonts.Body).Width) + D(12))
        Dim valX = col.Right - valW - ftW
        For i = 0 To rows.Count - 1
            Dim row = rows(i)
            Dim y = top + i * rowH
            ResultPanelView.DrawRowLabel(g, row.Label, New Rectangle(col.X, CInt(y), valX - col.X - D(8), CInt(rowH)), p.Ink)
            Dim bh = Math.Max(D(18), Math.Min(CInt(rowH) - D(3), If(row.Big, D(42), D(32))))
            Dim vb As New RectangleF(valX, y + (rowH - bh) / 2, valW, bh)
            Using path = Draw.RoundedRect(vb, D(5)), fill As New SolidBrush(Draw.Mix(p.Surface, row.Colour, 0.12)), pen As New Pen(Draw.Mix(p.Surface, row.Colour, 0.5))
                g.FillPath(fill, path)
                g.DrawPath(pen, path)
            End Using
            Dim inner As New Size(CInt(vb.Width) - D(6), CInt(vb.Height))
            Dim f = If(row.Big, ResultPanelView.FitFont(g, row.Value, inner, Fonts.Title, Fonts.Section, Fonts.Strong, Fonts.Body),
                                ResultPanelView.FitFont(g, row.Value, inner, Fonts.Section, Fonts.Strong, Fonts.Body, Fonts.Small))
            TextRenderer.DrawText(g, row.Value, f, Rectangle.Round(vb), row.Colour,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding Or TextFormatFlags.EndEllipsis)
            TextRenderer.DrawText(g, row.Note, Fonts.Body, New Rectangle(CInt(vb.Right) + D(8), CInt(y), ftW - D(8), CInt(rowH)), p.Muted,
                                  TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
        Next
    End Sub

    Private Shared Function Signed(v As Integer) As String
        Return If(v > 0, "+" & v, v.ToString())
    End Function

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
''' Edit a test after it was saved: the preferred colour (colour tests: on the wheel, or a bespoke
''' colour), the symptom grades (and whether symptoms were recorded), other symptoms, the
''' rate-of-reading counts (words per minute is worked out again) and the notes. A colour changed
''' here is noted with the result ("changed after the test, was ...").
''' </summary>
Public NotInheritable Class ResultEditDialog
    Inherits DialogBase

    Private ReadOnly _r As ReadEZResult
    Private ReadOnly _scales As New List(Of GradeScale)
    Private ReadOnly chkSymptoms As ThemedCheckBox = ThemedCheckBox.Create("results.edit.symptomsDone")
    Private ReadOnly txtOther As New TextField With {.Width = 530}
    Private ReadOnly nudTime As NumericUpDown = Ui.NumberBox(0, 600, 0, width:=110)
    Private ReadOnly nudWords As NumericUpDown = Ui.NumberBox(0, 2000, 0, width:=110)
    Private ReadOnly nudErrors As NumericUpDown = Ui.NumberBox(0, 500, 0, width:=110)
    Private ReadOnly nudSkipped As NumericUpDown = Ui.NumberBox(0, 20, 0, width:=110)
    Private ReadOnly lblWpm As New Label With {.AutoSize = True, .Font = Fonts.Section, .Margin = New Padding(0, 4, 0, 12)}
    Private ReadOnly txtNotes As New TextField(multiline:=True) With {.Width = 530, .Height = 70}
    Private ReadOnly wheel As New ColourWheel With {.Size = New Size(250, 250), .Margin = New Padding(0, 4, 0, 4)}
    Private ReadOnly lblColour As New Label With {.AutoSize = True, .Font = Fonts.Section, .Margin = New Padding(0, 0, 0, 4)}
    Private ReadOnly _originalColour As String
    Private _colourChanged As Boolean

    Public Sub New(r As ReadEZResult)
        _r = r.Clone()
        SetTitle("results.edit.title")
        SetSubtitleText(Loc.T(TestRun.ResultKey(r)) & If(r.IsBaseline, "", "  ·  " & r.ColourName) & "  ·  " & r.Test_Date.ToString("g", Loc.Culture))
        ClientSize = New Size(620, 800)

        Dim col = Ui.Column(4)
        col.AutoSize = False
        col.AutoScroll = True
        col.Dock = DockStyle.Fill

        ' Colour (colour tests)
        _originalColour = r.ColourName
        If Not r.IsBaseline Then
            col.Controls.Add(Ui.Section("results.edit.colour"))
            col.Controls.Add(lblColour)
            Dim hint = Ui.Label("results.edit.colourHint", muted:=True)
            hint.MaximumSize = New Size(540, 0)
            col.Controls.Add(hint)
            wheel.OuterSet = If(r.Mode = ReadEZTestMode.Overlay, ReadEZSet.Overlay, ReadEZSet.Lens)
            wheel.TwoRings = r.Mode <> ReadEZTestMode.Overlay
            If r.Screening_Mode = 0 AndAlso r.AutoScore <> "" Then wheel.Spokes = ColourStaircase.ParseScore(r.AutoScore)
            If r.Screening_Mode <> 2 Then wheel.Selected = (r.ColourSet, r.Colour1)
            AddHandler wheel.ColourClicked, Sub(st, i, b) PickColour(st, i)
            col.Controls.Add(wheel)
            Dim bespoke = ModernButton.Create("colour.bespoke", ButtonKind.Ghost, Glyphs.Color)
            bespoke.Margin = New Padding(0, 0, 0, 8)
            AddHandler bespoke.Click, Sub() PickBespoke()
            col.Controls.Add(bespoke)
            ShowColour()
        End If

        ' Symptoms
        col.Controls.Add(Ui.Section("results.symptoms"))
        chkSymptoms.Checked = r.SymptomsDone
        chkSymptoms.Margin = New Padding(0, 0, 0, 6)
        col.Controls.Add(chkSymptoms)
        Dim legend = Ui.Label("symptoms.scale", muted:=True)
        col.Controls.Add(legend)
        For q = 0 To 3
            Dim l = Ui.Label("symptoms.q" & (q + 1))
            l.Font = Fonts.Strong
            l.MaximumSize = New Size(540, 0)
            l.Margin = New Padding(0, 8, 0, 4)
            col.Controls.Add(l)
            Dim s As New GradeScale With {.Size = New Size(380, 40), .Grade = r.Symptoms(q)}
            AddHandler s.GradeChanged, Sub() chkSymptoms.Checked = True
            _scales.Add(s)
            col.Controls.Add(s)
        Next
        txtOther.Text = r.Symptoms_Notes
        col.Controls.Add(Ui.Stack("symptoms.other", txtOther))

        ' Rate of reading
        col.Controls.Add(Ui.Section("results.reading"))
        Dim grid As New TableLayoutPanel With {.ColumnCount = 2, .AutoSize = True, .BackColor = Color.Transparent, .Margin = New Padding(0)}
        grid.Controls.Add(Ui.Stack("reading.seconds", nudTime), 0, 0)
        grid.Controls.Add(Ui.Stack("reading.words", nudWords), 1, 0)
        grid.Controls.Add(Ui.Stack("reading.errors", nudErrors), 0, 1)
        grid.Controls.Add(Ui.Stack("reading.skipped", nudSkipped), 1, 1)
        col.Controls.Add(grid)
        col.Controls.Add(lblWpm)
        nudTime.Value = Math.Min(nudTime.Maximum, r.ROR_RecordTime)
        nudWords.Value = Math.Min(nudWords.Maximum, r.ROR_NoWordsRead)
        nudErrors.Value = Math.Min(nudErrors.Maximum, r.ROR_NoErrors)
        nudSkipped.Value = Math.Min(nudSkipped.Maximum, r.ROR_RowsSkipped)
        For Each n In {nudTime, nudWords, nudSkipped}
            AddHandler n.ValueChanged, Sub() UpdateWpm()
        Next
        UpdateWpm()

        ' Notes
        txtNotes.Text = r.Notes
        col.Controls.Add(Ui.Stack("results.notes", txtNotes))
        Body.Controls.Add(col)

        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        AddHandler save.Click, Sub() SaveClicked()
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
    End Sub

    Private Sub ShowColour()
        lblColour.Text = Loc.F("results.edit.current", _r.ColourName)
    End Sub

    ''' <summary>A colour clicked on the wheel (outer ring: overlay or light lens; inner ring: dark lens).</summary>
    Private Sub PickColour(st As ReadEZSet, index As Integer)
        If index < 1 OrElse index > 12 Then Return
        _r.Colour1 = index
        _r.PaletteSet = If(st = ReadEZSet.Overlay, 1, If(st = ReadEZSet.Lens, 2, 3))
        Dim c = ReadEZColours.Colour(st, index)
        _r.Colour_R = c.R : _r.Colour_G = c.G : _r.Colour_B = c.B
        If _r.Screening_Mode = 2 Then _r.Screening_Mode = 1        ' no longer a bespoke colour
        _r.GreyPreferred = False
        wheel.Selected = (st, index)
        wheel.Invalidate()
        _colourChanged = True
        ShowColour()
    End Sub

    Private Sub PickBespoke()
        Using dlg As New ColorDialog With {.FullOpen = True, .Color = _r.ScreenColour}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            _r.Colour1 = 0
            _r.PaletteSet = If(_r.Mode = ReadEZTestMode.Overlay, 1, 2)
            _r.Colour_R = dlg.Color.R : _r.Colour_G = dlg.Color.G : _r.Colour_B = dlg.Color.B
            _r.Screening_Mode = 2
            _r.GreyPreferred = False
            wheel.Selected = Nothing
            wheel.Invalidate()
            _colourChanged = True
            ShowColour()
        End Using
    End Sub

    Private Sub UpdateWpm()
        Dim wpm = ReadEZResult.WordsPerMinute(CInt(nudWords.Value), CInt(nudSkipped.Value), CInt(nudTime.Value))
        lblWpm.Text = Loc.F("results.edit.wpm", If(nudTime.Value > 0 AndAlso nudWords.Value > 0, wpm.ToString(), "–"))
    End Sub

    Private Sub SaveClicked()
        _r.SymptomsDone = chkSymptoms.Checked
        For q = 0 To _scales.Count - 1
            _r.Symptoms(q) = _scales(q).Grade
        Next
        _r.Symptoms_Notes = txtOther.Text.Trim()
        _r.ROR_RecordTime = CInt(nudTime.Value)
        _r.ROR_NoWordsRead = CInt(nudWords.Value)
        _r.ROR_NoErrors = CInt(nudErrors.Value)
        _r.ROR_RowsSkipped = CInt(nudSkipped.Value)
        _r.UpdateRate()
        _r.Notes = txtNotes.Text.Trim()
        ' Keep a note of the colour the test found when it is changed here (the first time only)
        If _colourChanged AndAlso _r.ColourName <> _originalColour AndAlso Not _r.AutoScore.Contains("ChangedFrom:") Then
            _r.AutoScore = If(_r.AutoScore = "", "", _r.AutoScore & ",") & "ChangedFrom:" & _originalColour.Replace(",", " ").Replace(":", " ")
        End If
        Try
            AppHost.Db.SaveResult(_r)
            DialogResult = DialogResult.OK
        Catch ex As Exception
            Log.Error("Edit test", ex)
            Ask.Error(Me, "common.saveFailed", ex.Message)
        End Try
    End Sub
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
