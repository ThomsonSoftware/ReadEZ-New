Option Strict On

' ====================================================================================================
'   PatientDialog.vb - add or edit a patient (replaces RZfrmClients)
'   Same fields and checks as before: first and family name required, date of birth not in the
'   future, e-mail address must look valid if given.
' ====================================================================================================

Public Class PatientDialog
    Inherits DialogBase

    Private ReadOnly _id As Integer
    Private ReadOnly cmbTitle As New ComboBox
    Private ReadOnly txtFirst As New TextField
    Private ReadOnly txtSurname As New TextField
    Private ReadOnly dtpDoB As New DateTimePicker
    Private ReadOnly txtReference As New TextField
    Private ReadOnly txtAddress1 As New TextField
    Private ReadOnly txtAddress2 As New TextField
    Private ReadOnly txtTown As New TextField
    Private ReadOnly txtCounty As New TextField
    Private ReadOnly txtPostcode As New TextField
    Private ReadOnly txtTel As New TextField
    Private ReadOnly txtMobile As New TextField
    Private ReadOnly cboOrg As New ThemedComboBox With {.Width = 270, .DropDownStyle = ComboBoxStyle.DropDownList}
    Private ReadOnly _orgs As List(Of OrganisationStructure)
    Private ReadOnly txtEmail As New TextField

    ''' <summary>ID of the saved patient (new or existing).</summary>
    Public Property SavedPatientId As Integer = -1

    ''' <param name="patientId">-1 to add a new patient.</param>
    Public Sub New(patientId As Integer)
        _id = patientId
        If patientId < 0 Then
            SetTitle("patient.dialog.addTitle", "patient.dialog.subtitle")
        Else
            SetTitle("patient.dialog.editTitle", "patient.dialog.subtitle")
        End If
        ClientSize = New Size(640, 720)
        _orgs = AppHost.Db.GetOrganisations()
        cboOrg.Items.Add(Loc.T("patient.noOrganisation"))
        For Each o In _orgs
            cboOrg.Items.Add(o.Name)
        Next
        cboOrg.SelectedIndex = 0

        cmbTitle.Items.AddRange({"Mr", "Mrs", "Ms", "Miss", "Mx", "Dr", "Prof", "Sir", "Lord", "Revd"})
        cmbTitle.Font = Fonts.Body
        cmbTitle.FlatStyle = FlatStyle.Flat
        cmbTitle.Width = 120
        dtpDoB.Format = DateTimePickerFormat.Short
        dtpDoB.Font = Fonts.Body
        dtpDoB.Width = 170
        dtpDoB.MaxDate = Date.Today
        dtpDoB.Value = Date.Today.AddYears(-30)
        ' Unticked = not known. Without this, a patient with no date of birth was saved with
        ' the default (30 years ago) the first time they were edited.
        dtpDoB.ShowCheckBox = True
        dtpDoB.Checked = False

        Dim grid = NewGrid(2)
        grid.Controls.Add(Ui.Stack("patient.title", cmbTitle), 0, 0)
        grid.Controls.Add(Ui.Stack("patient.dob", dtpDoB), 1, 0)
        grid.Controls.Add(Ui.Stack("patient.firstname", txtFirst, 270), 0, 1)
        grid.Controls.Add(Ui.Stack("patient.surname", txtSurname, 270), 1, 1)
        grid.Controls.Add(Ui.Stack("patient.reference", txtReference, 270), 0, 2)
        grid.Controls.Add(Ui.Stack("patient.tel", txtTel, 270), 1, 2)
        grid.Controls.Add(Ui.Stack("patient.email", txtEmail, 270), 0, 3)
        grid.Controls.Add(Ui.Stack("patient.mobile", txtMobile, 270), 1, 3)
        grid.Controls.Add(Ui.Stack("patient.address1", txtAddress1, 270), 0, 4)
        grid.Controls.Add(Ui.Stack("patient.address2", txtAddress2, 270), 1, 4)
        grid.Controls.Add(Ui.Stack("patient.town", txtTown, 270), 0, 5)
        grid.Controls.Add(Ui.Stack("patient.county", txtCounty, 270), 1, 5)
        grid.Controls.Add(Ui.Stack("patient.postcode", txtPostcode, 160), 0, 6)
        grid.Controls.Add(Ui.Stack("patient.organisation", cboOrg), 1, 6)
        Body.Controls.Add(grid)

        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        AddHandler save.Click, AddressOf SaveClicked
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AcceptButton = save

        If patientId >= 0 Then LoadPatient()
        AddHandler Shown, Sub() txtFirst.Box.Focus()
    End Sub

    Private Sub LoadPatient()
        Dim found = AppHost.Db.GetPatientById(_id)
        If Not found.HasValue Then Return
        Dim p = found.Value
        cmbTitle.Text = p.Title
        txtFirst.Text = p.Firstname
        txtSurname.Text = p.Surname
        If Ages.IsKnown(p.DoB) AndAlso p.DoB > dtpDoB.MinDate AndAlso p.DoB <= dtpDoB.MaxDate Then
            dtpDoB.Value = p.DoB
            dtpDoB.Checked = True
        End If
        txtReference.Text = p.Reference
        txtAddress1.Text = p.Address1
        txtAddress2.Text = p.Address2
        txtTown.Text = p.Town
        txtCounty.Text = p.County
        txtPostcode.Text = p.Postcode
        txtTel.Text = p.Tel
        txtEmail.Text = p.Email
        txtMobile.Text = p.Mobile
        Dim oi = _orgs.FindIndex(Function(o) o.ID = p.Organisation_ID)
        cboOrg.SelectedIndex = If(oi >= 0, oi + 1, 0)
        SetSubtitleText(Loc.F("patient.dialog.id", _id))
    End Sub

    Private Sub SaveClicked(sender As Object, e As EventArgs)
        txtFirst.Invalid = txtFirst.Text.Trim() = ""
        txtSurname.Invalid = txtSurname.Text.Trim() = ""
        If txtFirst.Invalid OrElse txtSurname.Invalid Then
            Ask.Info(Me, "patient.validation.names")
            Return
        End If
        If dtpDoB.Checked AndAlso dtpDoB.Value.Date > Date.Today Then
            Ask.Warn(Me, "patient.validation.dob")
            dtpDoB.Focus()
            Return
        End If
        txtEmail.Invalid = txtEmail.Text.Trim() <> "" AndAlso Not IsValidEmail(txtEmail.Text.Trim())
        If txtEmail.Invalid Then
            Ask.Warn(Me, "patient.validation.email")
            txtEmail.Box.Focus()
            Return
        End If

        Dim p As New PatientStructure With {
            .ID = _id,
            .Title = cmbTitle.Text.Trim(),
            .Firstname = txtFirst.Text.Trim(),
            .Surname = txtSurname.Text.Trim(),
            .DoB = If(dtpDoB.Checked, dtpDoB.Value.Date, Ages.UnknownDoB),
            .Reference = txtReference.Text.Trim(),
            .Address1 = txtAddress1.Text.Trim(),
            .Address2 = txtAddress2.Text.Trim(),
            .Town = txtTown.Text.Trim(),
            .County = txtCounty.Text.Trim(),
            .Postcode = txtPostcode.Text.Trim(),
            .Tel = txtTel.Text.Trim(),
            .Email = txtEmail.Text.Trim(),
            .Mobile = txtMobile.Text.Trim(),
            .Organisation_ID = If(cboOrg.SelectedIndex > 0, _orgs(cboOrg.SelectedIndex - 1).ID, 0)
        }
        Try
            If _id < 0 Then
                SavedPatientId = AppHost.Db.SavePatient(p)
            Else
                AppHost.Db.UpdatePatient(_id, p)
                SavedPatientId = _id
            End If
            DialogResult = DialogResult.OK
        Catch ex As Exception
            Log.Error("Save patient", ex)
            Ask.Error(Me, "common.saveFailed", ex.Message)
        End Try
    End Sub

    Private Shared Function IsValidEmail(email As String) As Boolean
        Try
            Return New Net.Mail.MailAddress(email).Address = email
        Catch
            Return False
        End Try
    End Function
End Class
