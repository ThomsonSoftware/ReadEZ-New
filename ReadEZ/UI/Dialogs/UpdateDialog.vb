Option Strict On

Imports System.Threading

' ====================================================================================================
'   UpdateDialog.vb - "A new version of ReadEZ is available"
'
'   Opened at start-up when the server has a newer version (not one the user chose to skip), and
'   from Help > Check for updates. Shows this version, the new version and what has changed;
'   "Install update" downloads the installer (with progress), checks it, then closes ReadEZ,
'   installs and starts it again. Settings, licence, databases and stimuli are kept.
' ====================================================================================================

Public Class UpdateDialog
    Inherits DialogBase

    Private _info As UpdateInfo
    Private ReadOnly lblVersions As New Label
    Private ReadOnly txtNotes As New TextField(multiline:=True)
    Private ReadOnly lblNotes As Label = Ui.Label("update.notes", muted:=True)
    Private ReadOnly bar As New ProgressBar With {.Width = 540, .Height = 10, .Style = ProgressBarStyle.Continuous, .Visible = False,
                                                  .Margin = New Padding(0, 4, 0, 6)}
    Private ReadOnly lblStatus As New Label
    Private ReadOnly chkAtStart As ThemedCheckBox = ThemedCheckBox.Create("update.checkAtStart")
    Private ReadOnly btnInstall As ModernButton
    Private ReadOnly btnSkip As ModernButton
    Private ReadOnly btnClose As ModernButton
    Private _cts As CancellationTokenSource
    Private _busy As Boolean

    ''' <summary>info = Nothing: check the server now (Help > Check for updates).</summary>
    Public Sub New(Optional info As UpdateInfo = Nothing)
        _info = info
        SetTitle("update.title", "update.subtitle")
        ClientSize = New Size(640, 470)

        ' One column that fills the window (no scrolling): version line, notes (fill), progress,
        ' status, start-up check box. The notes are hidden when there is no update to show.
        Body.AutoScroll = False
        Dim col As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 6, .BackColor = Color.Transparent}
        col.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        For i = 0 To 5
            col.RowStyles.Add(New RowStyle(If(i = 2, SizeType.Percent, SizeType.AutoSize), 100))
        Next
        lblVersions.AutoSize = True
        lblVersions.MaximumSize = New Size(560, 0)
        lblVersions.Font = Fonts.Strong
        lblVersions.Margin = New Padding(0, 4, 0, 10)
        LabelStyle.SetCustom(lblVersions)
        lblVersions.ForeColor = ThemeManager.Current.Ink
        col.Controls.Add(lblVersions, 0, 0)
        lblNotes.Margin = New Padding(0, 0, 0, 4)
        col.Controls.Add(lblNotes, 0, 1)
        txtNotes.Dock = DockStyle.Fill
        txtNotes.Margin = New Padding(0, 0, 0, 10)
        txtNotes.Box.ReadOnly = True
        col.Controls.Add(txtNotes, 0, 2)
        bar.Dock = DockStyle.Top
        col.Controls.Add(bar, 0, 3)
        lblStatus.AutoSize = True
        lblStatus.MaximumSize = New Size(560, 0)
        lblStatus.Margin = New Padding(0, 0, 0, 8)
        LabelStyle.SetMuted(lblStatus)
        col.Controls.Add(lblStatus, 0, 4)
        chkAtStart.Checked = AppHost.Settings.Update_CheckAtStart
        AddHandler chkAtStart.CheckedChanged, Sub()
                                                  AppHost.Settings.Update_CheckAtStart = chkAtStart.Checked
                                                  AppHost.SaveSettings()
                                              End Sub
        col.Controls.Add(chkAtStart, 0, 5)
        Body.Controls.Add(col)

        btnInstall = AddFooterButton("update.install", ButtonKind.Primary, Glyphs.Export)
        btnClose = AddFooterButton("update.later", ButtonKind.Secondary)
        btnSkip = AddFooterButton("update.skip", ButtonKind.Ghost)
        AddHandler btnInstall.Click, Async Sub() Await InstallAsync()
        AddHandler btnClose.Click, Sub() CloseOrCancel()
        AddHandler btnSkip.Click, Sub()
                                      If _info IsNot Nothing Then
                                          AppHost.Settings.Update_SkipVersion = _info.VersionText
                                          AppHost.SaveSettings()
                                      End If
                                      DialogResult = DialogResult.Cancel
                                  End Sub
        ShowInfo()
    End Sub

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        If _info Is Nothing Then Await CheckNowAsync()
    End Sub

    Private Async Function CheckNowAsync() As Task
        SetBusy(True)
        lblStatus.Text = Loc.T("update.checking")
        Try
            _info = Await UpdateService.CheckAsync()
            If _info Is Nothing Then
                lblVersions.Text = Loc.F("update.upToDate", AppHost.VersionNumber)
                lblStatus.Text = ""
            Else
                ShowInfo()
            End If
        Catch ex As Exception
            Log.Warn("Update check failed: " & ex.Message)
            lblVersions.Text = Loc.F("update.thisVersion", AppHost.VersionNumber)
            lblStatus.Text = Loc.F("update.checkFailed", ex.Message)
        Finally
            SetBusy(False)
        End Try
    End Function

    Private Sub ShowInfo()
        Dim has = _info IsNot Nothing
        If has Then
            lblVersions.Text = Loc.F("update.available", _info.VersionText, AppHost.VersionNumber) &
                               If(_info.Released = "", "", Environment.NewLine & Loc.F("update.released", _info.Released))
            txtNotes.Text = If(_info.Notes, "").Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf)
            lblStatus.Text = Loc.T("update.keeps")
        Else
            lblVersions.Text = Loc.F("update.thisVersion", AppHost.VersionNumber)
        End If
        lblNotes.Visible = has AndAlso txtNotes.Text.Trim() <> ""
        txtNotes.Visible = lblNotes.Visible
        btnInstall.Visible = has
        btnSkip.Visible = has
        btnClose.Text = Loc.T(If(has, "update.later", "common.close"))
    End Sub

    Private Sub SetBusy(busy As Boolean)
        _busy = busy
        btnInstall.Enabled = Not busy
        btnSkip.Enabled = Not busy
        UseWaitCursor = busy
    End Sub

    Private Sub CloseOrCancel()
        If _busy AndAlso _cts IsNot Nothing Then
            _cts.Cancel()
            Return
        End If
        DialogResult = DialogResult.Cancel
    End Sub

    Private Async Function InstallAsync() As Task
        If _info Is Nothing OrElse _busy Then Return
        SetBusy(True)
        btnClose.Text = Loc.T("common.cancel")
        bar.Value = 0
        bar.Visible = True
        lblStatus.Text = Loc.F("update.downloading", _info.VersionText)
        _cts = New CancellationTokenSource()
        Try
            Dim progress As New Progress(Of Integer)(Sub(p)
                                                         bar.Value = Math.Max(0, Math.Min(100, p))
                                                         lblStatus.Text = Loc.F("update.downloadingPercent", _info.VersionText, p)
                                                     End Sub)
            Dim msi = Await UpdateService.DownloadAsync(_info, progress, _cts.Token)
            If Not Ask.YesNo(Me, "update.readyToInstall", _info.VersionText) Then
                lblStatus.Text = Loc.T("update.notInstalled")
                Return
            End If
            UpdateService.StartInstall(msi)
            AppHost.SaveSettings()
            DialogResult = DialogResult.OK
            ' Close the main window; the script installs once the program has closed
            Dim main = Application.OpenForms.OfType(Of MainForm)().FirstOrDefault()
            If main IsNot Nothing Then main.BeginInvoke(Sub() main.Close())
        Catch ex As OperationCanceledException
            lblStatus.Text = Loc.T("update.cancelled")
        Catch ex As Exception
            Log.Error("Update download", ex)
            lblStatus.Text = Loc.F("update.failed", ex.Message)
        Finally
            bar.Visible = False
            _cts?.Dispose()
            _cts = Nothing
            If Not IsDisposed Then
                SetBusy(False)
                btnClose.Text = Loc.T("update.later")
            End If
        End Try
    End Function

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If _busy AndAlso _cts IsNot Nothing Then _cts.Cancel()
        MyBase.OnFormClosing(e)
    End Sub
End Class
