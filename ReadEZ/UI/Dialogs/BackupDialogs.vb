Option Strict On

Imports System.IO

' ====================================================================================================
'   BackupDialogs.vb - Back up now, Restore from backup, Move database (File menu and Settings > Data)
'
'     BackupCommands         the commands themselves (shared by the menu, Settings and closing)
'     BackupProgressDialog   "Backing up…" while a backup, restore or move runs in the background
'     RestoreBackupDialog    the backups in the backup folder (newest first) to choose from
' ====================================================================================================

Public NotInheritable Class BackupCommands
    Private Sub New()
    End Sub

    ''' <summary>Asks for the backup folder (and remembers it). Nothing if cancelled.</summary>
    Public Shared Function PickFolder(owner As Form) As String
        Using dlg As New FolderBrowserDialog With {.Description = Loc.T("backup.pickFolder"), .UseDescriptionForTitle = True}
            Dim current = AppHost.Settings.BackupFolder
            If Not String.IsNullOrWhiteSpace(current) AndAlso Directory.Exists(current) Then dlg.SelectedPath = current
            If dlg.ShowDialog(owner) <> DialogResult.OK Then Return Nothing
            If IsInsideDataFolder(dlg.SelectedPath) Then Ask.Warn(owner, "backup.sameFolderWarning")
            AppHost.Settings.BackupFolder = dlg.SelectedPath
            AppHost.SaveSettings()
            Return dlg.SelectedPath
        End Using
    End Function

    ''' <summary>True if the folder is the database's own folder (a backup there is lost with the disk).</summary>
    Private Shared Function IsInsideDataFolder(folder As String) As Boolean
        Try
            Dim dbFolder = Path.GetFullPath(Path.GetDirectoryName(AppHost.Settings.Local_Database)).TrimEnd(Path.DirectorySeparatorChar)
            Return String.Equals(Path.GetFullPath(folder).TrimEnd(Path.DirectorySeparatorChar), dbFolder, StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>File > Back up now.</summary>
    Public Shared Function BackUpNow(owner As Form) As Boolean
        Dim folder = AppHost.Settings.BackupFolder
        If String.IsNullOrWhiteSpace(folder) Then
            folder = PickFolder(owner)
            If folder Is Nothing Then Return False
        End If
        Dim info As BackupInfo = Nothing
        Dim include = AppHost.Settings.BackupIncludeStimuli
        Dim err = BackupProgressDialog.Run(owner, "backup.backingUp",
                                           Sub(report) info = BackupService.Create(folder, BackupKind.Manual, include, report))
        AppHost.SaveSettings()
        If err IsNot Nothing Then
            Ask.Error(owner, "backup.failed", err.Message)
            Return False
        End If
        If owner IsNot Nothing Then Toast.Show(owner, Loc.F("backup.done", Path.GetFileName(info.FilePath)), ToastKind.Success)
        Return True
    End Function

    ''' <summary>Backs up while the program closes, if switched on and something changed. Never stops the close.</summary>
    Public Shared Sub BackUpOnClose(owner As Form)
        If Not BackupService.ShouldBackUpOnClose() Then Return
        Dim folder = AppHost.Settings.BackupFolder
        Dim include = AppHost.Settings.BackupIncludeStimuli
        Dim err = BackupProgressDialog.Run(owner, "backup.backingUp",
                                           Sub(report) BackupService.Create(folder, BackupKind.OnClose, include, report))
        AppHost.SaveSettings()
        If err IsNot Nothing Then
            Log.Error("Backup on close", err)
            Ask.Warn(owner, "backup.closeFailed", folder, err.Message)
        End If
    End Sub

    ''' <summary>File > Restore from backup.</summary>
    Public Shared Sub Restore(owner As Form)
        Using dlg As New RestoreBackupDialog()
            If dlg.ShowDialog(owner) <> DialogResult.OK OrElse dlg.Chosen Is Nothing Then Return
            Dim chosen = dlg.Chosen
            Dim stimuli = dlg.RestoreStimuli
            If Not Ask.YesNoWarning(owner, "backup.restore.confirm", chosen.Created.ToString("g", Loc.Culture)) Then Return
            Dim err = BackupProgressDialog.Run(owner, "backup.restoring", Sub(report) BackupService.Restore(chosen, stimuli, report))
            AppHost.SaveSettings()
            If err IsNot Nothing Then
                Log.Error("Restore", err)
                Ask.Error(owner, "backup.restore.failed", err.Message)
                ReopenDatabase()
                Return
            End If
            ReopenDatabase()
            Ask.Info(owner, "backup.restore.done", chosen.Created.ToString("g", Loc.Culture))
        End Using
    End Sub

    ''' <summary>Settings > Data > Move database.</summary>
    Public Shared Function MoveDatabase(owner As Form) As Boolean
        Using dlg As New FolderBrowserDialog With {.Description = Loc.T("backup.move.pick"), .UseDescriptionForTitle = True}
            If dlg.ShowDialog(owner) <> DialogResult.OK Then Return False
            Dim folder = dlg.SelectedPath
            If Not Ask.YesNo(owner, "backup.move.confirm", AppHost.Settings.Local_Database, folder) Then Return False
            Dim newPath As String = Nothing
            Dim err = BackupProgressDialog.Run(owner, "backup.moving", Sub(report) newPath = BackupService.MoveDatabase(folder))
            If err IsNot Nothing Then
                Log.Error("Move database", err)
                Ask.Error(owner, "backup.move.failed", err.Message)
                Return False
            End If
            ForgetPatient()
            AppCommands.NotifyDataChanged(True)
            BackupService.NotifyDatabaseReplaced()
            Ask.Info(owner, "backup.move.done", newPath)
            Return True
        End Using
    End Function

    ''' <summary>After a restore the patients are different: open the database again and start from no patient.</summary>
    Private Shared Sub ReopenDatabase()
        Try
            AppHost.Db = New ReadEZDatabase(AppHost.Settings.Local_Database)
        Catch ex As Exception
            Log.Error("Reopen database", ex)
        End Try
        ForgetPatient()
        AppCommands.NotifyDataChanged(True)
        BackupService.NotifyDatabaseReplaced()
        AppHost.MainWindow?.ShowScreen(MainForm.ScreenPatients)
    End Sub

    Private Shared Sub ForgetPatient()
        AppHost.Session.SelectPatient(Nothing)
        AppHost.Settings.LastPatientId = -1
        AppHost.SaveSettings()
    End Sub
End Class

''' <summary>A small window that runs work in the background and shows what it is doing.</summary>
Public NotInheritable Class BackupProgressDialog
    Inherits DialogBase

    Private ReadOnly lblStatus As New Label With {.AutoSize = True, .Font = Fonts.Body}
    Private ReadOnly _work As Action(Of Action(Of String))
    Private _error As Exception

    Private Sub New(titleKey As String, work As Action(Of Action(Of String)))
        _work = work
        SetTitle(titleKey, "backup.wait")
        ClientSize = New Size(460, 200)
        ControlBox = False
        Footer.Visible = False
        lblStatus.Location = New Point(28, 16)
        LabelStyle.SetMuted(lblStatus)
        Body.Controls.Add(lblStatus)
    End Sub

    ''' <summary>Runs the work (on another thread) while the window is shown. Returns the error, or Nothing.</summary>
    Public Shared Function Run(owner As IWin32Window, titleKey As String, work As Action(Of Action(Of String))) As Exception
        Using dlg As New BackupProgressDialog(titleKey, work)
            dlg.ShowDialog(owner)
            Return dlg._error
        End Using
    End Function

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        UseWaitCursor = True
        Dim progress As New Progress(Of String)(Sub(t) lblStatus.Text = t)
        Dim reporter As IProgress(Of String) = progress
        Try
            Await Task.Run(Sub() _work(Sub(t) reporter.Report(t)))
        Catch ex As Exception
            _error = ex
        End Try
        UseWaitCursor = False
        DialogResult = DialogResult.OK
        Close()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        ' Not closed by the user (Alt+F4) while the work runs
        If DialogResult <> DialogResult.OK Then e.Cancel = True
        MyBase.OnFormClosing(e)
    End Sub
End Class

''' <summary>Lists the backups in the backup folder; the chosen one is restored by BackupCommands.Restore.</summary>
Public NotInheritable Class RestoreBackupDialog
    Inherits DialogBase

    Private ReadOnly grid As New DataGridView
    Private ReadOnly txtFolder As New TextField
    Private ReadOnly chkStimuli As ThemedCheckBox = ThemedCheckBox.Create("backup.restore.stimuli")
    Private ReadOnly lblNone As Label
    Private ReadOnly btnRestore As ModernButton
    Private _folder As String
    Private _list As New List(Of BackupInfo)

    Public Property Chosen As BackupInfo
    Public ReadOnly Property RestoreStimuli As Boolean
        Get
            Return chkStimuli.Checked AndAlso chkStimuli.Enabled
        End Get
    End Property

    Public Sub New()
        SetTitle("backup.restore.title", "backup.restore.subtitle")
        ClientSize = New Size(940, 560)

        Dim top = Ui.Column()
        top.Dock = DockStyle.Top
        Dim row = Ui.Row()
        txtFolder.Width = 560
        txtFolder.Box.ReadOnly = True
        row.Controls.Add(Ui.Stack("backup.folder", txtFolder))
        Dim browse = ModernButton.Create("backup.otherFolder", ButtonKind.Secondary, Glyphs.Folder)
        browse.Margin = New Padding(8, 22, 0, 0)
        AddHandler browse.Click, Sub() OtherFolder()
        row.Controls.Add(browse)
        top.Controls.Add(row)
        lblNone = Ui.Label("backup.restore.none", muted:=True)
        lblNone.Margin = New Padding(0, 0, 0, 8)
        top.Controls.Add(lblNone)

        ThemeManager.StyleGrid(grid)
        grid.Dock = DockStyle.Fill
        grid.ReadOnly = True
        grid.MultiSelect = False
        grid.AllowUserToAddRows = False
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddColumn("When", "backup.col.when", 22)
        AddColumn("Kind", "backup.col.kind", 18)
        AddColumn("Patients", "backup.col.patients", 10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        AddColumn("Recordings", "backup.col.recordings", 12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        AddColumn("Stimuli", "backup.col.stimuli", 10)
        AddColumn("Size", "backup.col.size", 10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        AddColumn("Computer", "backup.col.computer", 18)

        Dim bottom As New Panel With {.Dock = DockStyle.Bottom, .Height = 40, .BackColor = Color.Transparent}
        chkStimuli.Location = New Point(0, 10)
        chkStimuli.Checked = True
        bottom.Controls.Add(chkStimuli)

        Body.Controls.Add(grid)
        Body.Controls.Add(bottom)
        Body.Controls.Add(top)
        grid.BringToFront()

        btnRestore = AddFooterButton("backup.restore.button", ButtonKind.Primary, Glyphs.Restore)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AddHandler btnRestore.Click, Sub() Choose()
        AddHandler grid.CellDoubleClick, Sub(s, e) If e.RowIndex >= 0 Then Choose()
        AddHandler grid.SelectionChanged, Sub() UpdateChoice()

        Fill(AppHost.Settings.BackupFolder)
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        ' The grid selects its first row when it is shown: select the newest backup again
        If grid.Rows.Count > 0 Then
            grid.ClearSelection()
            grid.Rows(0).Selected = True
        End If
        UpdateChoice()
    End Sub

    Private Function AddColumn(name As String, key As String, weight As Single) As DataGridViewColumn
        Dim c As New DataGridViewTextBoxColumn With {.Name = name, .FillWeight = Math.Max(1, weight)}
        grid.Columns.Add(c)
        Loc.Bind(c, key)
        Return c
    End Function

    Private Sub Fill(folder As String)
        _folder = If(folder, "")
        txtFolder.Text = _folder
        _list = BackupService.List(_folder)
        grid.Rows.Clear()
        For Each b In _list
            Dim i = grid.Rows.Add(b.Created.ToString("g", Loc.Culture), KindText(b.Kind), b.Patients.ToString("N0"), b.Recordings.ToString("N0"),
                                  If(b.HasStimuli, Loc.T("common.yes"), Loc.T("common.no")), SizeText(b.SizeBytes), b.Computer)
            grid.Rows(i).Tag = b
        Next
        lblNone.Visible = _list.Count = 0
        If _list.Count > 0 Then grid.Rows(0).Selected = True
        UpdateChoice()
    End Sub

    Private Shared Function KindText(k As BackupKind) As String
        Select Case k
            Case BackupKind.OnClose : Return Loc.T("backup.kind.onClose")
            Case BackupKind.BeforeRestore : Return Loc.T("backup.kind.beforeRestore")
            Case Else : Return Loc.T("backup.kind.manual")
        End Select
    End Function

    Private Shared Function SizeText(bytes As Long) As String
        If bytes >= 1024L * 1024 * 1024 Then Return (bytes / 1024.0 / 1024 / 1024).ToString("0.0") & " GB"
        If bytes >= 1024L * 1024 Then Return (bytes / 1024.0 / 1024).ToString("0.0") & " MB"
        Return Math.Max(1, bytes \ 1024).ToString("N0") & " KB"
    End Function

    Private ReadOnly Property Selected As BackupInfo
        Get
            If grid.SelectedRows.Count = 0 Then Return Nothing
            Return TryCast(grid.SelectedRows(0).Tag, BackupInfo)
        End Get
    End Property

    Private Sub UpdateChoice()
        Dim b = Selected
        btnRestore.Enabled = b IsNot Nothing
        chkStimuli.Enabled = b IsNot Nothing AndAlso b.HasStimuli
    End Sub

    Private Sub OtherFolder()
        Using dlg As New FolderBrowserDialog With {.Description = Loc.T("backup.pickRestoreFolder"), .UseDescriptionForTitle = True}
            If Directory.Exists(_folder) Then dlg.SelectedPath = _folder
            If dlg.ShowDialog(Me) = DialogResult.OK Then Fill(dlg.SelectedPath)
        End Using
    End Sub

    Private Sub Choose()
        Dim b = Selected
        If b Is Nothing Then Return
        Chosen = b
        DialogResult = DialogResult.OK
    End Sub
End Class
