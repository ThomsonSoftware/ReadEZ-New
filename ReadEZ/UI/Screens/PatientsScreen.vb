Option Strict On

Imports System.Data

' ====================================================================================================
'   PatientsScreen.vb - choose a patient and see their tests (as Eye Dynamix's Patients screen;
'   replaces ReadEZ 2020's database list)
'
'     +----------------------------+  +-----------------------------------------+
'     | Search  [Organisation v]   |  | Patient details                         |
'     | Patients grid              |  +-----------------------------------------+
'     |                            |  | Tests (date, test, colour, wpm, ...)    |
'     | [Edit] [New patient]       |  | [View results] [Start a test]           |
'     +----------------------------+  +-----------------------------------------+
' ====================================================================================================

Public Class PatientsScreen
    Inherits ScreenBase

    Private ReadOnly txtSearch As New TextField
    Private ReadOnly cboOrg As New ThemedComboBox With {.Width = 220, .DropDownStyle = ComboBoxStyle.DropDownList}
    Private ReadOnly grid As New DataGridView
    Private ReadOnly tests As New ResultsList
    Private ReadOnly searchTimer As New Timer With {.Interval = 250}
    Private ReadOnly lblDetails As New Label
    Private Shared ReadOnly ListFont As New Font("Segoe UI", 10.75F)
    Private Shared ReadOnly NameFont As New Font("Segoe UI Semibold", 13.5F)
    Private ReadOnly lblName As New Label
    Private ReadOnly lblCount As New Label
    Private ReadOnly lblEmpty As New Label
    Private btnEdit, btnAdd, btnView, btnTest As ModernButton
    Private _loading As Boolean
    Private _patientCount As Integer
    Private _orgs As New List(Of OrganisationStructure)

    Public Sub New()
        TitleKey = "patients.title"
        HelpTopic = "screen.patients"
        SubtitleKey = "patients.subtitle"

        Dim split As New SplitContainer With {.Dock = DockStyle.Fill, .SplitterWidth = 12, .FixedPanel = FixedPanel.Panel1, .BackColor = Color.Transparent}
        ThemeManager.SetRole(split, SurfaceRole.Canvas)
        Dim placed = False
        AddHandler split.SizeChanged, Sub()
                                          If placed OrElse split.Width < 400 Then Return
                                          placed = True
                                          split.SplitterDistance = CInt(split.Width * 0.45)
                                      End Sub
        split.Panel1.Controls.Add(BuildPatientList())
        split.Panel2.Controls.Add(BuildPatientDetail())
        Content.Controls.Add(split)

        AddHandler searchTimer.Tick, Sub()
                                         searchTimer.Stop()
                                         LoadPatients()
                                     End Sub
        AddHandler txtSearch.Box.TextChanged, Sub()
                                                  searchTimer.Stop()
                                                  searchTimer.Start()
                                              End Sub
        AddHandler cboOrg.SelectedIndexChanged, Sub()
                                                    If _loading Then Return
                                                    AppHost.Settings.OrganisationFilter = If(cboOrg.SelectedIndex <= 0, -1, _orgs(cboOrg.SelectedIndex - 1).ID)
                                                    LoadPatients()
                                                End Sub
        AddHandler grid.SelectionChanged, AddressOf GridSelectionChanged
        AddHandler grid.ColumnHeaderMouseClick, AddressOf SortByColumn
        AddHandler grid.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then AppCommands.EditPatient()
        AddHandler tests.ResultSelected, Sub() TestSelected()
        AddHandler tests.ResultOpened, Sub() ViewResults()
        AddHandler AppHost.Session.ResultSaved, Sub(r) If Visible Then ShowPatient()
        AddHandler AppCommands.DataChanged, AddressOf OnDataChanged
        Loc.OnChange(Me, Sub()
                             lblCount.Text = Loc.F("patients.count", _patientCount)
                             If AppHost.Session.HasPatient Then ShowPatient()
                         End Sub)
    End Sub

#Region "Layout"

    Private Function BuildPatientList() As Control
        Dim card As New Card With {.Dock = DockStyle.Fill, .Padding = New Padding(16)}
        Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 4, .BackColor = Color.Transparent}
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim searchRow As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .AutoSize = True, .BackColor = Color.Transparent, .Margin = New Padding(0)}
        searchRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        searchRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        txtSearch.Dock = DockStyle.Fill
        txtSearch.SetPlaceholder("patients.search")
        cboOrg.Margin = New Padding(8, 2, 0, 0)
        searchRow.Controls.Add(txtSearch, 0, 0)
        searchRow.Controls.Add(cboOrg, 1, 0)
        layout.Controls.Add(searchRow, 0, 0)

        lblCount.AutoSize = True
        lblCount.Font = Fonts.Small
        lblCount.Margin = New Padding(0, 0, 0, 6)
        LabelStyle.SetMuted(lblCount)
        layout.Controls.Add(lblCount, 0, 1)

        ThemeManager.StyleGrid(grid)
        grid.DefaultCellStyle.Font = ListFont
        grid.ColumnHeadersDefaultCellStyle.Font = Fonts.Body
        grid.RowTemplate.Height += 4
        grid.ColumnHeadersHeight += 2
        grid.Dock = DockStyle.Fill
        grid.ReadOnly = True
        grid.AutoGenerateColumns = False
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        AddColumn("Name", "patients.col.name", 36)
        AddColumn("DoB", "patients.col.dob", 18)
        AddColumn("Organisation", "patients.col.organisation", 24)
        AddColumn("ResultCount", "patients.col.tests", 10)
        AddColumn("LastTestDate", "patients.col.lastTest", 16)
        layout.Controls.Add(grid, 0, 2)

        Dim buttons = Ui.Row()
        buttons.Margin = New Padding(0, 12, 0, 0)
        btnEdit = ModernButton.Create("patients.edit", ButtonKind.Secondary, Glyphs.Edit)
        btnAdd = ModernButton.Create("patients.add", ButtonKind.Primary, Glyphs.Add)
        AddHandler btnEdit.Click, Sub() AppCommands.EditPatient()
        AddHandler btnAdd.Click, Sub() AppCommands.NewPatient()
        buttons.Controls.Add(btnEdit)
        buttons.Controls.Add(btnAdd)
        layout.Controls.Add(buttons, 0, 3)

        card.Controls.Add(layout)
        Return card
    End Function

    Private Sub AddColumn(name As String, key As String, weight As Single)
        Dim c As New DataGridViewTextBoxColumn With {.Name = name, .DataPropertyName = name, .FillWeight = weight, .SortMode = DataGridViewColumnSortMode.Programmatic}
        grid.Columns.Add(c)
        Loc.Bind(c, key)
    End Sub

    Private Function BuildPatientDetail() As Control
        Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .BackColor = Color.Transparent}
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        Dim summary As New Card With {.Dock = DockStyle.Top, .Height = 130, .Margin = New Padding(0, 0, 0, 12)}
        lblName.AutoSize = True
        lblName.Font = NameFont
        lblName.Location = New Point(18, 16)
        lblDetails.AutoSize = True
        lblDetails.Font = ListFont
        lblDetails.Location = New Point(18, 50)
        LabelStyle.SetMuted(lblDetails)
        summary.Controls.Add(lblName)
        summary.Controls.Add(lblDetails)
        layout.Controls.Add(summary, 0, 0)

        Dim recCard As New Card With {.Dock = DockStyle.Fill, .Margin = New Padding(0)}
        recCard.SetTitle("tests.list.title")
        Dim inner As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .BackColor = Color.Transparent}
        inner.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        inner.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        tests.Dock = DockStyle.Fill
        lblEmpty.Dock = DockStyle.Fill
        lblEmpty.TextAlign = ContentAlignment.MiddleCenter
        lblEmpty.Visible = False
        LabelStyle.SetMuted(lblEmpty)
        Loc.Bind(lblEmpty, "tests.list.none")
        Dim host As New Panel With {.Dock = DockStyle.Fill, .BackColor = Color.Transparent}
        host.Controls.Add(tests)
        host.Controls.Add(lblEmpty)
        inner.Controls.Add(host, 0, 0)

        Dim buttons = Ui.Row()
        buttons.Margin = New Padding(0, 12, 0, 0)
        btnView = ModernButton.Create("results.view", ButtonKind.Primary, Glyphs.Chart)
        btnTest = ModernButton.Create("patients.startTest", ButtonKind.Secondary, Glyphs.Play)
        AddHandler btnView.Click, Sub() ViewResults()
        AddHandler btnTest.Click, Sub() AppHost.MainWindow?.ShowScreen(MainForm.ScreenTests)
        buttons.Controls.Add(btnView)
        buttons.Controls.Add(btnTest)
        inner.Controls.Add(buttons, 0, 1)
        recCard.Controls.Add(inner)
        layout.Controls.Add(recCard, 0, 1)
        Return layout
    End Function

#End Region

#Region "Loading"

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        FillOrganisations()
        LoadPatients()
        txtSearch.Box.Focus()
    End Sub

    Private Sub FillOrganisations()
        _loading = True
        Try
            _orgs = AppHost.Db.GetOrganisations()
            cboOrg.Items.Clear()
            cboOrg.Items.Add(Loc.T("patients.allOrganisations"))
            For Each o In _orgs
                cboOrg.Items.Add(o.Name)
            Next
            Dim i = _orgs.FindIndex(Function(o) o.ID = AppHost.Settings.OrganisationFilter)
            cboOrg.SelectedIndex = If(i >= 0, i + 1, 0)
            cboOrg.Visible = _orgs.Count > 0
        Finally
            _loading = False
        End Try
    End Sub

    Private Sub LoadPatients()
        _loading = True
        Try
            Dim filter = If(cboOrg.SelectedIndex <= 0, -1, _orgs(cboOrg.SelectedIndex - 1).ID)
            Dim dt = AppHost.Db.SearchPatients(txtSearch.Text, filter)
            dt.Columns.Add("Name", GetType(String))
            dt.Columns.Add("DoBDate", GetType(Date))
            dt.Columns.Add("LastTestDate", GetType(Date))
            For Each r As DataRow In dt.Rows
                r("Name") = $"{ReadEZDatabase.Str(r("Surname"))}, {ReadEZDatabase.Str(r("Firstname"))}".Trim(" "c, ","c)
                Dim d = ReadEZDatabase.DateOf(r("DoB"))
                r("DoBDate") = If(Ages.IsKnown(d), CObj(d), DBNull.Value)
                Dim lt = ReadEZDatabase.DateOf(r("LastTest"))
                r("LastTestDate") = If(lt > Date.MinValue, CObj(lt.Date), DBNull.Value)
            Next
            For Each n In {"DoB", "LastTestDate"}
                Dim c = grid.Columns(n)
                c.DefaultCellStyle.Format = "d"
                c.DefaultCellStyle.FormatProvider = Loc.Culture
            Next
            grid.Columns("DoB").DataPropertyName = "DoBDate"
            grid.DataSource = dt
            _patientCount = dt.Rows.Count
            lblCount.Text = Loc.F("patients.count", _patientCount)
            SelectRowOf(If(AppHost.Session.HasPatient, AppHost.Session.PatientId, AppHost.Settings.LastPatientId))
        Catch ex As Exception
            Log.Error("LoadPatients", ex)
        Finally
            _loading = False
        End Try
        GridSelectionChanged(Nothing, EventArgs.Empty)
    End Sub

    Private Sub SelectRowOf(patientId As Integer)
        grid.ClearSelection()
        For Each row As DataGridViewRow In grid.Rows
            If ReadEZDatabase.Int(DirectCast(row.DataBoundItem, DataRowView)("ID")) = patientId Then
                row.Selected = True
                grid.CurrentCell = row.Cells(0)
                Exit For
            End If
        Next
    End Sub

    Private Sub SortByColumn(sender As Object, e As DataGridViewCellMouseEventArgs)
        If e.Button <> MouseButtons.Left OrElse e.ColumnIndex < 0 Then Return
        Dim col = grid.Columns(e.ColumnIndex)
        Dim dir = If(grid.SortedColumn Is col AndAlso grid.SortOrder = SortOrder.Ascending,
                     ComponentModel.ListSortDirection.Descending, ComponentModel.ListSortDirection.Ascending)
        _loading = True
        Try
            grid.Sort(col, dir)
            SelectRowOf(AppHost.Session.PatientId)
        Finally
            _loading = False
        End Try
    End Sub

    Private Function SelectedPatientId() As Integer
        If grid.CurrentRow Is Nothing OrElse Not grid.CurrentRow.Selected Then Return -1
        Dim drv = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
        Return If(drv Is Nothing, -1, ReadEZDatabase.Int(drv("ID")))
    End Function

    Private Sub GridSelectionChanged(sender As Object, e As EventArgs)
        If _loading Then Return
        Dim id = SelectedPatientId()
        If id > 0 AndAlso id <> AppHost.Session.PatientId Then
            AppHost.Session.SelectPatient(AppHost.Db.GetPatientById(id))
        End If
        ShowPatient()
    End Sub

    Private Sub ShowPatient()
        Dim s = AppHost.Session
        btnEdit.Enabled = s.HasPatient
        btnTest.Enabled = s.HasPatient
        If Not s.HasPatient Then
            lblName.Text = Loc.T("top.noPatient")
            lblDetails.Text = Loc.T("patients.selectHint")
            tests.SetResults(New List(Of ReadEZResult), -1)
            lblEmpty.Visible = False
            TestSelected()
            Return
        End If
        Dim p = s.Patient
        lblName.Text = s.PatientName
        Dim lines As New List(Of String) From {TopBar.PatientDetail(p)}
        Dim org = AppHost.Db.GetOrganisation(p.Organisation_ID)
        If org.HasValue Then lines.Add(org.Value.Name)
        Dim address = String.Join(", ", {p.Address1, p.Address2, p.Town, p.County, p.Postcode}.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))
        If address <> "" Then lines.Add(address)
        Dim contact = String.Join("  ·  ", {p.Tel, p.Mobile, p.Email}.Where(Function(x) Not String.IsNullOrWhiteSpace(x)))
        If contact <> "" Then lines.Add(contact)
        lblDetails.Text = String.Join(Environment.NewLine, lines.Where(Function(x) x <> ""))
        Dim results = AppHost.Db.GetResults(p.ID)
        tests.SetResults(results, AppHost.Session.ResultId)
        lblEmpty.Visible = results.Count = 0
        tests.Visible = results.Count > 0
        TestSelected()
    End Sub

    Private Sub TestSelected()
        Dim r = tests.SelectedResult
        btnView.Enabled = AppHost.Session.HasPatient AndAlso tests.Count > 0
        If r IsNot Nothing Then AppHost.Session.SelectResult(r.ID)
    End Sub

#End Region

    Private Sub ViewResults()
        If Not AppHost.Session.HasPatient Then Return
        Dim r = tests.SelectedResult
        If r IsNot Nothing Then AppHost.Session.SelectResult(r.ID)
        AppHost.MainWindow?.ShowScreen(MainForm.ScreenResults)
    End Sub

    Private Sub OnDataChanged(clearSearch As Boolean)
        If IsDisposed Then Return
        If clearSearch Then
            searchTimer.Stop()
            txtSearch.Text = ""
        End If
        If Not Visible Then Return
        FillOrganisations()
        LoadPatients()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveHandler AppCommands.DataChanged, AddressOf OnDataChanged
            searchTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

End Class

''' <summary>Search-and-pick a patient.</summary>
Public Class PatientPickerDialog
    Inherits DialogBase

    Private ReadOnly txtSearch As New TextField
    Private ReadOnly list As New ListBox
    Private _table As DataTable

    Public Property SelectedPatientId As Integer = -1
    Public Property SelectedPatientName As String = ""

    Public Sub New(titleKey As String)
        SetTitle(titleKey)
        ClientSize = New Size(460, 520)
        Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2}
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        txtSearch.Dock = DockStyle.Fill
        txtSearch.SetPlaceholder("patients.search")
        list.Dock = DockStyle.Fill
        list.BorderStyle = BorderStyle.FixedSingle
        list.Font = Fonts.Body
        list.IntegralHeight = False
        list.ItemHeight = 22
        layout.Controls.Add(txtSearch, 0, 0)
        layout.Controls.Add(list, 0, 1)
        Body.Controls.Add(layout)

        Dim ok = AddFooterButton("common.ok", ButtonKind.Primary)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AcceptButton = ok
        AddHandler ok.Click, Sub() Accept()
        AddHandler list.DoubleClick, Sub() Accept()
        AddHandler txtSearch.Box.TextChanged, Sub() Fill()
        Fill()
    End Sub

    Private Sub Fill()
        _table = AppHost.Db.SearchPatients(txtSearch.Text)
        list.BeginUpdate()
        list.Items.Clear()
        For Each r As DataRow In _table.Rows
            Dim d = ReadEZDatabase.DateOf(r("DoB"))
            list.Items.Add($"{ReadEZDatabase.Str(r("Surname"))}, {ReadEZDatabase.Str(r("Firstname"))}" &
                           If(Ages.IsKnown(d), "   (" & d.ToString("d", Loc.Culture) & ")", ""))
        Next
        list.EndUpdate()
        If list.Items.Count > 0 Then list.SelectedIndex = 0
    End Sub

    Private Sub Accept()
        If list.SelectedIndex < 0 Then
            Ask.Info(Me, "patients.pick.none")
            Return
        End If
        Dim r = _table.Rows(list.SelectedIndex)
        SelectedPatientId = ReadEZDatabase.Int(r("ID"))
        SelectedPatientName = $"{ReadEZDatabase.Str(r("Firstname"))} {ReadEZDatabase.Str(r("Surname"))}"
        DialogResult = DialogResult.OK
    End Sub
End Class

''' <summary>A patient's tests: date, test, colour swatch and name, words per minute, symptoms.</summary>
Public NotInheritable Class ResultsList
    Inherits ListBox
    Implements IThemed

    Private _results As New List(Of ReadEZResult)
    Public Event ResultSelected As EventHandler
    Public Event ResultOpened As EventHandler

    Public Sub New()
        DrawMode = DrawMode.OwnerDrawFixed
        ItemHeight = 44
        IntegralHeight = False
        BorderStyle = BorderStyle.None
        Font = Fonts.Body
        AddHandler SelectedIndexChanged, Sub() RaiseEvent ResultSelected(Me, EventArgs.Empty)
        AddHandler DoubleClick, Sub() If SelectedIndex >= 0 Then RaiseEvent ResultOpened(Me, EventArgs.Empty)
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return _results.Count
        End Get
    End Property

    Public Sub SetResults(results As List(Of ReadEZResult), selectId As Integer)
        _results = results
        BeginUpdate()
        Items.Clear()
        For Each r In results
            Items.Add(r.ID)
        Next
        EndUpdate()
        Dim i = results.FindIndex(Function(x) x.ID = selectId)
        If i < 0 AndAlso results.Count > 0 Then i = 0
        SelectedIndex = i
    End Sub

    Public ReadOnly Property SelectedResult As ReadEZResult
        Get
            Return If(SelectedIndex >= 0 AndAlso SelectedIndex < _results.Count, _results(SelectedIndex), Nothing)
        End Get
    End Property

    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        If e.Index < 0 OrElse e.Index >= _results.Count Then Return
        Dim p = ThemeManager.Current
        Dim r = _results(e.Index)
        Dim g = e.Graphics
        Dim sel = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Using b As New SolidBrush(If(sel, p.PrimarySoft, p.Surface))
            g.FillRectangle(b, e.Bounds)
        End Using
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        Dim sw As New Rectangle(e.Bounds.X + 10, e.Bounds.Y + 9, 26, 26)
        Using b As New SolidBrush(r.ScreenColour), pen As New Pen(p.Border)
            g.FillEllipse(b, sw)
            g.DrawEllipse(pen, sw)
        End Using
        Dim x = sw.Right + 12
        Dim title = Loc.T(TestRun.ModeKey(r.Mode)) & If(r.IsBaseline, "", "  ·  " & r.ColourName)
        TextRenderer.DrawText(g, title, Fonts.Strong, New Rectangle(x, e.Bounds.Y + 4, e.Bounds.Width - x - 140, 20), p.Ink, TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
        Dim parts As New List(Of String) From {r.Test_Date.ToString("g", Loc.Culture)}
        If r.HasReading Then parts.Add(Loc.F("results.wpmShort", r.ROR))
        If r.SymptomsDone Then parts.Add(Loc.F("results.symptomsShort", r.SymptomTotal))
        If r.ImportedFrom <> "" Then parts.Add(r.ImportedFrom)
        TextRenderer.DrawText(g, String.Join("  ·  ", parts), Fonts.Small, New Rectangle(x, e.Bounds.Y + 23, e.Bounds.Width - x - 10, 18), p.Muted, TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
        Using pen As New Pen(p.GridLine)
            g.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1)
        End Using
    End Sub

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        BackColor = pal.Surface
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property
End Class
