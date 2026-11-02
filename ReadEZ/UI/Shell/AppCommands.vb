Option Strict On

' ====================================================================================================
'   AppCommands.vb - patient, test and database commands shared by the menu and the screens
' ====================================================================================================

Public NotInheritable Class AppCommands
    Private Sub New()
    End Sub

    ''' <summary>Patients or tests were added, changed, moved or deleted (clearSearch: show all patients).</summary>
    Public Shared Event DataChanged(clearSearch As Boolean)

    Private Shared ReadOnly Property Owner As Form
        Get
            Return AppHost.MainWindow
        End Get
    End Property

    Public Shared Sub NotifyDataChanged(clearSearch As Boolean)
        RaiseEvent DataChanged(clearSearch)
    End Sub

#Region "Patients"

    Public Shared Sub NewPatient()
        Using f As New PatientDialog(-1)
            If f.ShowDialog(Owner) <> DialogResult.OK Then Return
            AppHost.Session.SelectPatient(AppHost.Db.GetPatientById(f.SavedPatientId))
        End Using
        RaiseEvent DataChanged(True)
        AppHost.MainWindow?.ShowScreen(MainForm.ScreenPatients)
    End Sub

    Public Shared Sub EditPatient()
        If Not AppHost.Session.HasPatient Then Return
        Using f As New PatientDialog(AppHost.Session.PatientId)
            If f.ShowDialog(Owner) <> DialogResult.OK Then Return
            AppHost.Session.ReloadPatient()
        End Using
        RaiseEvent DataChanged(False)
    End Sub

    Public Shared Sub DeletePatient()
        Dim s = AppHost.Session
        If Not s.HasPatient Then Return
        Dim name = s.PatientName
        If Not Ask.YesNoWarning(Owner, "patients.delete.confirm1", name) Then Return
        If Not Ask.YesNoWarning(Owner, "patients.delete.confirm2", name) Then Return
        Try
            Dim n = AppHost.Db.DeletePatient(s.PatientId)
            s.SelectPatient(Nothing)
            RaiseEvent DataChanged(False)
            Toast.Show(Owner, If(n > 0, Loc.F("patients.delete.doneWithTests", name, n), Loc.F("patients.delete.done", name)), ToastKind.Success)
            LeaveDeletedResult()
        Catch ex As Exception
            Log.Error("DeletePatient", ex)
            Ask.Error(Owner, "common.errorWithDetail", ex.Message)
        End Try
    End Sub

    Public Shared Sub EditOrganisations()
        Using f As New OrganisationsDialog()
            f.ShowDialog(Owner)
        End Using
        RaiseEvent DataChanged(False)
    End Sub

    Public Shared Sub EditScreeners()
        Using f As New ScreenersDialog()
            f.ShowDialog(Owner)
        End Using
        AppHost.MainWindow?.RefreshStatus()
    End Sub

    ''' <summary>Top bar: choose who is running the tests (or edit the list).</summary>
    Public Shared Sub PickScreener()
        Using f As New ScreenersDialog(picking:=True)
            If f.ShowDialog(Owner) = DialogResult.OK Then AppHost.Session.SelectScreener(f.SelectedScreenerId)
        End Using
        AppHost.MainWindow?.RefreshStatus()
    End Sub

#End Region

#Region "Tests"

    Public Shared Sub EditResult()
        If Not AppHost.Session.HasResult Then Return
        Dim r = AppHost.Db.LoadResult(AppHost.Session.ResultId)
        If r Is Nothing Then Return
        Dim text = Ask.Text(Owner, "results.notes", "results.notes.prompt", If(r.Notes, ""))
        If text Is Nothing Then Return
        Try
            AppHost.Db.UpdateNotes(r.ID, text.Trim())
            RaiseEvent DataChanged(False)
        Catch ex As Exception
            Log.Error("Update notes", ex)
            Ask.Error(Owner, "common.errorWithDetail", ex.Message)
        End Try
    End Sub

    Public Shared Sub MoveResult()
        If Not AppHost.Session.HasResult Then Return
        Dim r = AppHost.Db.LoadResult(AppHost.Session.ResultId)
        If r Is Nothing Then Return
        Using picker As New PatientPickerDialog("results.move.title")
            If picker.ShowDialog(Owner) <> DialogResult.OK Then Return
            If picker.SelectedPatientId = r.Client_ID Then
                Ask.Info(Owner, "results.move.same")
                Return
            End If
            If Not Ask.YesNo(Owner, "results.move.confirm", picker.SelectedPatientName) Then Return
            Try
                AppHost.Db.MoveResult(r.ID, picker.SelectedPatientId)
                AppHost.Session.SelectResult(-1)
                RaiseEvent DataChanged(False)
                Toast.Show(Owner, Loc.F("results.move.done", picker.SelectedPatientName), ToastKind.Success)
            Catch ex As Exception
                Log.Error("MoveResult", ex)
                Ask.Error(Owner, "common.errorWithDetail", ex.Message)
            End Try
        End Using
    End Sub

    Public Shared Sub DeleteResult()
        If Not AppHost.Session.HasResult Then Return
        If Not Ask.YesNoWarning(Owner, "results.delete.confirm") Then Return
        Try
            AppHost.Db.DeleteResult(AppHost.Session.ResultId)
            AppHost.Session.SelectResult(-1)
            RaiseEvent DataChanged(False)
            Toast.Show(Owner, Loc.T("results.delete.done"), ToastKind.Success)
        Catch ex As Exception
            Log.Error("DeleteResult", ex)
            Ask.Error(Owner, "common.errorWithDetail", ex.Message)
        End Try
    End Sub

    Private Shared Sub LeaveDeletedResult()
        Dim main = AppHost.MainWindow
        If main IsNot Nothing AndAlso main.CurrentScreenId = MainForm.ScreenResults Then main.ShowScreen(MainForm.ScreenPatients)
    End Sub

#End Region

End Class

''' <summary>Database commands: import a ReadEZ 2020 database.</summary>
Public NotInheritable Class DatabaseCommands
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Imports a ReadEZ 2020 database (ReadEZ2020.db) into the database in use. Patients and tests
    ''' already imported from the same file are not added again.
    ''' </summary>
    Public Shared Sub ImportOldDatabase(owner As IWin32Window)
        Using dlg As New OpenFileDialog With {.Filter = Loc.T("import.filter"), .Title = Loc.T("import.pick"), .CheckFileExists = True}
            Dim found = LegacyDatabaseImport.FindOldDatabase()
            If found IsNot Nothing Then
                dlg.InitialDirectory = IO.Path.GetDirectoryName(found)
                dlg.FileName = IO.Path.GetFileName(found)
            End If
            If dlg.ShowDialog(owner) <> DialogResult.OK Then Return
            If Not LegacyDatabaseImport.IsLegacyDatabase(dlg.FileName) Then
                Ask.Warn(owner, "import.notOld", dlg.FileName)
                Return
            End If
            ' A backup first, so the import can be undone with File > Restore
            Try
                If Not String.IsNullOrWhiteSpace(AppHost.Settings.BackupFolder) Then
                    BackupService.Create(AppHost.Settings.BackupFolder, BackupKind.BeforeRestore, includeStimuli:=False)
                End If
            Catch ex As Exception
                Log.Warn("Backup before import: " & ex.Message)
            End Try
            Dim f = TryCast(owner, Form)
            If f IsNot Nothing Then f.Cursor = Cursors.WaitCursor
            Try
                Dim sum = LegacyDatabaseImport.Import(dlg.FileName, AppHost.Db)
                Ask.Info(owner, "import.done", sum.Patients, sum.Tests, sum.AlreadyThere, sum.Skipped)
                AppCommands.NotifyDataChanged(True)
            Catch ex As Exception
                Log.Error("Import " & dlg.FileName, ex)
                Ask.Error(owner, "import.failed", dlg.FileName, ex.Message)
            Finally
                If f IsNot Nothing Then f.Cursor = Cursors.Default
            End Try
        End Using
    End Sub
End Class
