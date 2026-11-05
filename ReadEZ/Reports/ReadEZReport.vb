Option Strict On

Imports System.IO

' ====================================================================================================
'   ReadEZReport.vb - the printed reports (replaces RZfrmReports and its HTML templates)
'
'   Report  patient and test details, the colour chosen (key result), change with colour,
'           symptoms (chart) and rate of reading (table) - with the baseline beside the colour test.
'   Letter  to a parent / carer or a teacher: the opening and closing paragraphs come from the data
'           folder's "ReadEZ Templates" (ParentOpeningPara.txt ... TeacherClosingPara.txt, as in
'           ReadEZ 2020) with {Firstname}, {Surname}, {DateOfTest} and {Result} filled in; the
'           results go in between.
'   Both use ReportBuilder (the same look as Eye Dynamix's reports) and open in the preview, which
'   prints, saves as PDF / Word and e-mails.
' ====================================================================================================

Public NotInheritable Class ReadEZReport
    Private Sub New()
    End Sub

    Public Shared Sub ShowReport(owner As IWin32Window, r As ReadEZResult, baseline As ReadEZResult)
        Try
            Using rb As New ReportBuilder("report.title")
                rb.AddPatientBlock(r.Client_ID, TestLines(r))
                AddResults(rb, r, baseline)
                rb.ShowPreview(owner)
            End Using
        Catch ex As Exception
            Log.Error("Report", ex)
            Ask.Error(owner, "common.errorWithDetail", ex.Message)
        End Try
    End Sub

    Public Shared Sub ShowLetter(owner As IWin32Window, r As ReadEZResult, baseline As ReadEZResult)
        Dim choice = Ask.YesNoCancel(owner, "letter.which")
        If choice = DialogResult.Cancel Then Return
        Dim prefix = If(choice = DialogResult.Yes, "Parent", "Teacher")
        Try
            Using rb As New ReportBuilder(If(prefix = "Parent", "letter.titleParent", "letter.titleTeacher"))
                Dim p = AppHost.Db.GetPatientById(r.Client_ID)
                rb.AddParagraph(Fill(Template(prefix & "OpeningPara.txt"), p, r))
                AddResults(rb, r, baseline)
                rb.AddParagraph(Fill(Template(prefix & "ClosingPara.txt"), p, r))
                Dim screener = AppHost.Db.ScreenerName(r.Screener_ID)
                If screener <> "" Then rb.AddParagraph(Loc.T("letter.yours") & vbCrLf & vbCrLf & screener)
                rb.ShowPreview(owner)
            End Using
        Catch ex As Exception
            Log.Error("Letter", ex)
            Ask.Error(owner, "common.errorWithDetail", ex.Message)
        End Try
    End Sub

    Private Shared Function TestLines(r As ReadEZResult) As List(Of String)
        Dim l As New List(Of String) From {
            Loc.T("report.test") & ": " & Loc.T(TestRun.ResultKey(r)),
            Loc.T("report.date") & ": " & r.Test_Date.ToString("g", Loc.Culture)}
        Dim screener = AppHost.Db.ScreenerName(r.Screener_ID)
        If screener <> "" Then l.Add(Loc.T("report.screener") & ": " & screener)
        If r.ROR_Text <> "" Then l.Add(Loc.T("report.text") & ": " & r.ROR_Text)
        If r.Notes <> "" Then l.Add(Loc.T("report.notes") & ": " & r.Notes)
        Return l
    End Function

    ''' <summary>The colour, change with colour, symptoms and reading.</summary>
    Private Shared Sub AddResults(rb As ReportBuilder, r As ReadEZResult, baseline As ReadEZResult)
        If Not r.IsBaseline Then
            Dim note = Loc.T(TestRun.ModeKey(r.Mode)) & "  ·  RGB " & r.ScreenColour.R & ", " & r.ScreenColour.G & ", " & r.ScreenColour.B &
                       If(r.GreyPreferred, vbCrLf & Loc.T("colour.greyPreferred"), "")
            rb.AddKeyResult(Loc.T("report.colourChosen"), r.ColourName, note, If(r.GreyPreferred, Brand.Level.Caution, Brand.Level.Info))
            rb.AddImage(Swatch(r), maxHeight:=90, border:=False)
            If baseline IsNot Nothing Then
                For Each it In ResultsScreen.ChangeItems(r, baseline)
                    rb.AddKeyResult(it.Label, it.Value, it.Note, it.Level)
                Next
            End If
        End If

        rb.AddSection(Loc.T("results.symptoms"))
        If r.SymptomsDone Then
            Dim rows As New List(Of String())
            Dim compare = baseline IsNot Nothing AndAlso baseline.SymptomsDone AndAlso Not r.IsBaseline
            For q = 0 To 3
                rows.Add(If(compare, {Loc.T("symptoms.q" & (q + 1)), baseline.Symptoms(q).ToString(), r.Symptoms(q).ToString()},
                                     {Loc.T("symptoms.q" & (q + 1)), r.Symptoms(q).ToString()}))
            Next
            rows.Add(If(compare, {Loc.T("report.total"), baseline.SymptomTotal.ToString(), r.SymptomTotal.ToString()},
                                 {Loc.T("report.total"), r.SymptomTotal.ToString()}))
            rb.AddTable(If(compare, {Loc.T("report.question"), Loc.T("results.legend.baseline"), r.ColourName}, {Loc.T("report.question"), Loc.T("report.grade")}),
                        rows, If(compare, {3.0F, 1, 1}, {4.0F, 1}))
            rb.AddImage(SymptomsPicture(r, If(compare, baseline, Nothing)), maxHeight:=230, border:=False)
            If r.Symptoms_Notes <> "" Then rb.AddTextBox(Loc.T("symptoms.other"), {r.Symptoms_Notes})
        Else
            rb.AddParagraph(Loc.T("results.notDone"))
        End If

        rb.AddSection(Loc.T("results.reading"))
        If r.HasReading Then
            Dim compare = baseline IsNot Nothing AndAlso baseline.HasReading AndAlso Not r.IsBaseline
            Dim cols = If(compare, {baseline, r}, {r})
            Dim headers = {Loc.T("report.measure")}.Concat(cols.Select(Function(c) If(c.IsBaseline, Loc.T("results.legend.baseline"), c.ColourName))).ToArray()
            Dim rows As New List(Of String()) From {
                {Loc.T("report.wpm")}.Concat(cols.Select(Function(c) c.ROR.ToString())).ToArray(),
                {Loc.T("reading.words")}.Concat(cols.Select(Function(c) c.ROR_NoWordsRead.ToString())).ToArray(),
                {Loc.T("reading.seconds")}.Concat(cols.Select(Function(c) c.ROR_RecordTime.ToString())).ToArray(),
                {Loc.T("reading.errors")}.Concat(cols.Select(Function(c) c.ROR_NoErrors.ToString())).ToArray(),
                {Loc.T("reading.skipped")}.Concat(cols.Select(Function(c) c.ROR_RowsSkipped.ToString())).ToArray()}
            rb.AddTable(headers, rows, If(compare, {3.0F, 1, 1}, {4.0F, 1}))
        Else
            rb.AddParagraph(Loc.T("results.notDone"))
        End If
    End Sub

    Private Shared Function Swatch(r As ReadEZResult) As Image
        Dim bmp As New Bitmap(600, 90)
        Using g = Graphics.FromImage(bmp)
            g.Clear(Color.White)
            Dim rect As New Rectangle(225, 5, 150, 80)
            Using b As New SolidBrush(r.ScreenColour), p As New Pen(Brand.TintBorder)
                g.FillRectangle(b, rect)
                g.DrawRectangle(p, rect)
            End Using
        End Using
        Return bmp
    End Function

    Private Shared Function SymptomsPicture(r As ReadEZResult, baseline As ReadEZResult) As Image
        Dim bmp As New Bitmap(900, 260)
        bmp.SetResolution(96, 96)
        Using g = Graphics.FromImage(bmp), chart As New SymptomsChart With {.Result = r, .Baseline = baseline, .ForPrint = True}
            chart.Render(g, New Rectangle(0, 0, bmp.Width, bmp.Height))
        End Using
        Return bmp
    End Function

#Region "Letter templates"

    Private Shared Function Template(name As String) As String
        Dim file = Path.Combine(AppHost.Settings.DataDirectory, "ReadEZ Templates", name)
        Try
            If IO.File.Exists(file) Then Return IO.File.ReadAllText(file).Trim()
        Catch ex As Exception
            Log.Warn("Letter template " & file & ": " & ex.Message)
        End Try
        Return Loc.T("letter.default." & Path.GetFileNameWithoutExtension(name))
    End Function

    Private Shared Function Fill(text As String, p As PatientStructure?, r As ReadEZResult) As String
        Dim first = If(p.HasValue, p.Value.Firstname, "")
        Dim sur = If(p.HasValue, p.Value.Surname, "")
        Dim result = If(r.IsBaseline, Loc.T("letter.result.baseline"), Loc.F(If(r.Mode = ReadEZTestMode.Overlay, "letter.result.overlay", "letter.result.lens"), r.ColourName))
        Return text.Replace("{Firstname}", first).Replace("{Surname}", sur).Replace("{DateOfTest}", r.Test_Date.ToString("D", Loc.Culture)).
                    Replace("{Result}", result).Replace("  ", " ")
    End Function

#End Region
End Class
