Option Strict On

' ====================================================================================================
'   LicenceDialog.vb - product licence window (replaces frmOnLineActivation, frmOfflineActivation
'   and frmMarketingMessage)
'
'   Two pages, chosen with the switch at the top:
'     Online   - e-mail + licence key, Activate / Check / Deactivate (Thomson licence server)
'     Offline  - serial number, activation code, Deactivate, 3-day trial
'
'   StartUp mode : shown before the main window, only when the program may not run (as D-Chart:
'                  LicenceService.NeedsLicenceWindow). "Continue" is only available when the program
'                  may run; "Quit" closes the program. If the licence turns out to be fine the window
'                  closes by itself.
'   Info mode    : opened from the licence chip in the top bar (or Help > Licence). Shows the
'                  licence, and allows it to be checked, changed or deactivated.
'
'   All server calls are asynchronous so the window never freezes while the server is contacted.
' ====================================================================================================

Public Enum LicenceDialogMode
    StartUp
    Info
End Enum

Public Class LicenceDialog
    Inherits DialogBase

    Private ReadOnly _mode As LicenceDialogMode
    Private ReadOnly segKind As New SegmentedControl
    Private ReadOnly pageOnline As FlowLayoutPanel = Ui.Column()
    Private ReadOnly pageOffline As FlowLayoutPanel = Ui.Column()
    Private ReadOnly lblMessage As New Label
    Private ReadOnly btnContinue As ModernButton
    Private ReadOnly btnQuit As ModernButton
    Private _busy As Boolean
    Private _canRun As Boolean
    Private _quit As Boolean
    Private ReadOnly lblMarketing As New Label

    ' Online
    Private ReadOnly txtEmail As New TextField
    Private ReadOnly txtKey As New TextField
    Private ReadOnly lblLicensedTo As New Label
    Private ReadOnly btnActivate As ModernButton = ModernButton.Create("licence.activate", ButtonKind.Primary, Glyphs.Accept)
    Private ReadOnly btnCheck As ModernButton = ModernButton.Create("licence.check", ButtonKind.Secondary, Glyphs.Refresh)
    Private ReadOnly btnDeactivateOnline As ModernButton = ModernButton.Create("licence.deactivate", ButtonKind.Ghost, Glyphs.Cancel)

    ' Offline
    Private ReadOnly lblSerial As New Label
    Private ReadOnly txtCode As New TextField
    Private ReadOnly lblOfflineExpiry As New Label
    Private ReadOnly btnActivateOffline As ModernButton = ModernButton.Create("licence.activate", ButtonKind.Primary, Glyphs.Accept)
    Private ReadOnly btnDeactivateOffline As ModernButton = ModernButton.Create("licence.deactivate", ButtonKind.Ghost, Glyphs.Cancel)
    Private ReadOnly btnTrial As ModernButton = ModernButton.Create("licence.trial", ButtonKind.Ghost, Glyphs.Info)

    Public Sub New(mode As LicenceDialogMode)
        _mode = mode
        SetTitle("licence.title", "licence.subtitle")
        ClientSize = New Size(640, 600)
        If mode = LicenceDialogMode.StartUp Then
            StartPosition = FormStartPosition.CenterScreen
            ShowInTaskbar = True
        End If

        Dim stack = Ui.Column()
        stack.Dock = DockStyle.Top
        segKind.SetOptions("licence.online", "licence.offline")
        segKind.Margin = New Padding(0, 4, 0, 16)
        stack.Controls.Add(segKind)
        BuildOnline()
        BuildOffline()
        stack.Controls.Add(pageOnline)
        stack.Controls.Add(pageOffline)

        lblMessage.AutoSize = True
        lblMessage.MaximumSize = New Size(580, 0)
        lblMessage.Margin = New Padding(0, 12, 0, 0)
        lblMessage.Font = Fonts.Strong
        LabelStyle.SetCustom(lblMessage)
        stack.Controls.Add(lblMessage)
        lblMarketing.AutoSize = True
        lblMarketing.MaximumSize = New Size(580, 0)
        lblMarketing.Margin = New Padding(0, 10, 0, 0)
        LabelStyle.SetMuted(lblMarketing)
        stack.Controls.Add(lblMarketing)
        Body.Controls.Add(stack)

        If mode = LicenceDialogMode.StartUp Then
            btnContinue = AddFooterButton("licence.continue", ButtonKind.Primary, Glyphs.Forward)
            btnQuit = AddFooterButton("licence.quit", ButtonKind.Secondary)
            AddHandler btnContinue.Click, Sub() DialogResult = DialogResult.OK
            AddHandler btnQuit.Click, Sub()
                                           _quit = True
                                           DialogResult = DialogResult.Cancel
                                       End Sub
        Else
            btnContinue = AddFooterButton("common.close", ButtonKind.Primary)
            btnQuit = btnContinue
            AddHandler btnContinue.Click, Sub() DialogResult = DialogResult.OK
        End If

        segKind.SelectedIndex = If(LicenceService.IsOffline, 1, 0)
        AddHandler segKind.SelectedIndexChanged, Sub() SwitchKind()

        ' Service shortcut kept from the previous version: Shift + right-click on the header
        ' fills in a 12-month offline code for this computer.
        AddHandler Header.MouseDown, Sub(s, e)
                                         If e.Button = MouseButtons.Right AndAlso (ModifierKeys And Keys.Shift) = Keys.Shift AndAlso LicenceService.IsOffline Then
                                             txtCode.Text = LicenceService.GenerateServiceCode()
                                         End If
                                     End Sub
    End Sub

#Region "Layout"

    Private Sub BuildOnline()
        Dim intro = Ui.Label("licence.onlineIntro", muted:=True)
        intro.MaximumSize = New Size(580, 0)
        pageOnline.Controls.Add(intro)
        txtEmail.Width = 420
        pageOnline.Controls.Add(Ui.Stack("licence.email", txtEmail))
        Dim row = Ui.Row()
        txtKey.Width = 420
        row.Controls.Add(Ui.Stack("licence.key", txtKey))
        Dim paste = ModernButton.Create("licence.paste", ButtonKind.Secondary, Glyphs.Copy)
        paste.Margin = New Padding(0, 22, 0, 0)
        AddHandler paste.Click, Sub() If Clipboard.ContainsText() Then txtKey.Text = Clipboard.GetText().Trim()
        row.Controls.Add(paste)
        pageOnline.Controls.Add(row)

        lblLicensedTo.AutoSize = True
        lblLicensedTo.Margin = New Padding(0, 0, 0, 8)
        LabelStyle.SetMuted(lblLicensedTo)
        pageOnline.Controls.Add(lblLicensedTo)

        Dim buttons = Ui.Row()
        btnActivate.Margin = New Padding(0, 0, 8, 0)
        btnCheck.Margin = New Padding(0, 0, 8, 0)
        buttons.Controls.Add(btnActivate)
        buttons.Controls.Add(btnCheck)
        buttons.Controls.Add(btnDeactivateOnline)
        pageOnline.Controls.Add(buttons)

        AddHandler btnActivate.Click, Async Sub() Await ActivateOnlineAsync()
        AddHandler btnCheck.Click, Async Sub() Await CheckOnlineAsync()
        AddHandler btnDeactivateOnline.Click, Async Sub() Await DeactivateOnlineAsync()
        AddHandler txtEmail.Box.TextChanged, Sub() UpdateOnlineButtons()
        AddHandler txtKey.Box.TextChanged, Sub() UpdateOnlineButtons()
    End Sub

    Private Sub BuildOffline()
        Dim intro = Ui.Label("licence.offlineIntro", muted:=True)
        intro.MaximumSize = New Size(580, 0)
        pageOffline.Controls.Add(intro)

        Dim serialRow = Ui.Row()
        Dim serialStack = Ui.Column()
        serialStack.Controls.Add(Ui.Label("licence.serial", muted:=True))
        lblSerial.AutoSize = True
        lblSerial.Font = Fonts.Section
        serialStack.Controls.Add(lblSerial)
        serialRow.Controls.Add(serialStack)
        Dim copy = ModernButton.Create("licence.copy", ButtonKind.Secondary, Glyphs.Copy)
        copy.Margin = New Padding(16, 14, 0, 0)
        AddHandler copy.Click, Sub() Clipboard.SetText(lblSerial.Text)
        serialRow.Controls.Add(copy)
        pageOffline.Controls.Add(serialRow)

        Dim codeRow = Ui.Row()
        txtCode.Width = 300
        txtCode.Box.MaxLength = 16
        txtCode.Box.CharacterCasing = CharacterCasing.Upper
        codeRow.Controls.Add(Ui.Stack("licence.code", txtCode))
        Dim paste = ModernButton.Create("licence.paste", ButtonKind.Secondary, Glyphs.Copy)
        paste.Margin = New Padding(0, 22, 8, 0)
        AddHandler paste.Click, Sub() If Clipboard.ContainsText() Then txtCode.Text = Clipboard.GetText().Trim()
        codeRow.Controls.Add(paste)
        btnActivateOffline.Margin = New Padding(0, 22, 0, 0)
        codeRow.Controls.Add(btnActivateOffline)
        pageOffline.Controls.Add(codeRow)

        lblOfflineExpiry.AutoSize = True
        lblOfflineExpiry.Margin = New Padding(0, 0, 0, 8)
        LabelStyle.SetMuted(lblOfflineExpiry)
        pageOffline.Controls.Add(lblOfflineExpiry)

        Dim buttons = Ui.Row()
        btnDeactivateOffline.Margin = New Padding(0, 0, 8, 0)
        buttons.Controls.Add(btnDeactivateOffline)
        buttons.Controls.Add(btnTrial)
        pageOffline.Controls.Add(buttons)

        AddHandler txtCode.Box.TextChanged, Sub() btnActivateOffline.Enabled = Not _busy AndAlso txtCode.Text.Trim().Length = 16
        AddHandler btnActivateOffline.Click, Sub() ActivateOffline()
        AddHandler btnDeactivateOffline.Click, Sub() DeactivateOffline()
        AddHandler btnTrial.Click, Sub() StartTrial()
    End Sub

#End Region

#Region "Start"

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Dim d = LicenceService.Data
        txtEmail.Text = d.OL_Email
        txtKey.Text = d.OL_LicenceCode
        ShowPage()

        If Not LicenceService.HostnameMatches Then
            ' Copied file or renamed computer: activate again here (this used to lock the window
            ' with only Quit available, so even a renamed computer needed support to recover)
            LicenceService.ResetForThisComputer()
            ShowPage()
            If LicenceService.IsOffline Then RefreshOffline(fromStart:=False)    ' shows the new serial number
            SetMessage("licence.otherComputer", ToastKind.Error)
            SetCanRun(False)
            UpdateOnlineButtons()
            Return
        End If

        If LicenceService.IsOffline Then
            RefreshOffline(fromStart:=_mode = LicenceDialogMode.StartUp)
        ElseIf LicenceService.HasOnlineActivation Then
            Await CheckOnlineAsync(fromStart:=_mode = LicenceDialogMode.StartUp)
        Else
            SetMessage("licence.notActivated", ToastKind.Info)
            SetCanRun(False)
        End If
        UpdateOnlineButtons()
    End Sub

    Private Sub SwitchKind()
        Dim offline = segKind.SelectedIndex = 1
        If offline = LicenceService.IsOffline Then Return
        LicenceService.SwitchMode(offline)
        SetMessage("", ToastKind.Info)
        ShowPage()
        If offline Then RefreshOffline(fromStart:=False) Else SetCanRun(False)
        UpdateOnlineButtons()
    End Sub

    Private Sub ShowPage()
        pageOnline.Visible = Not LicenceService.IsOffline
        pageOffline.Visible = LicenceService.IsOffline
        ' The kind can only be changed while nothing is activated (as before).
        segKind.Enabled = Not (LicenceService.HasOnlineActivation OrElse LicenceService.Data.Act_Registered)
    End Sub

#End Region

#Region "Online"

    Private Function ReadKey(ByRef key As Guid) As Boolean
        If Guid.TryParse(txtKey.Text.Trim(), key) Then Return True
        SetMessage("licence.keyFormat", ToastKind.Error)
        Return False
    End Function

    Private Async Function ActivateOnlineAsync() As Task
        Dim key As Guid
        If Not ReadKey(key) Then Return
        SetBusy(True, "licence.connecting")
        Try
            Dim r = Await LicenceApiClient.ActivateAsync(txtEmail.Text.Trim(), key)
            Dim days = LicenceService.DaysUntil(r.Termination)
            If r.LicenceStatus = LicenceStatusType.Activated AndAlso days >= 0 Then
                Dim d = LicenceService.Data
                d.OL_Registered = True
                d.OL_LicenceCode = txtKey.Text.Trim()
                d.OL_ActivationClientToken = r.ActivationToken.ToString()
                d.OL_Email = txtEmail.Text.Trim()
                d.OL_ExpiryDate = r.Termination
                d.OL_PhoneHomeDays = r.NextCheckRequiredDays
                d.OL_LastCheckDate = Date.Now
                If r.LicencedTo <> "" Then d.OL_LicencedTo = r.LicencedTo
                d.Demo_Count = 0
                If LicenceService.Save() Then
                    ShowMarketing(r.MarketingMessage)
                    ReportValid()
                Else
                    SetMessage("licence.writeFailed", ToastKind.Error)
                    Try
                        Await LicenceApiClient.DeactivateAsync(key, r.ActivationToken)
                    Catch ex As LicenceServerUnavailableException
                    End Try
                    LicenceService.Data.OL_ActivationClientToken = ""
                End If
            ElseIf r.LicenceStatus = LicenceStatusType.Activated Then
                SetMessageText(Loc.F("licence.expiredOn", r.Termination.ToShortDateString()), ToastKind.Error)
                Try
                    Await LicenceApiClient.DeactivateAsync(key, r.ActivationToken)
                Catch ex As LicenceServerUnavailableException
                End Try
                SetCanRun(False)
            Else
                SetMessageText(StatusText(r), ToastKind.Error)
                SetCanRun(False)
            End If
        Catch ex As LicenceServerUnavailableException
            SetMessage("licence.serverUnavailable", ToastKind.Error)
        Finally
            SetBusy(False)
            ShowPage()
        End Try
    End Function

    Private Async Function CheckOnlineAsync(Optional fromStart As Boolean = False) As Task
        Dim d = LicenceService.Data
        Dim key As Guid, token As Guid
        If Not Guid.TryParse(d.OL_LicenceCode, key) OrElse Not Guid.TryParse(d.OL_ActivationClientToken, token) Then
            SetMessage("licence.keyFormat", ToastKind.Error)
            SetCanRun(False)
            Return
        End If
        SetBusy(True, "licence.checking")
        Try
            Dim r = Await LicenceApiClient.CheckAsync(key, token)
            If r.LicenceStatus = LicenceStatusType.Valid Then
                d.OL_LastCheckDate = Date.Now
                d.OL_LicencedTo = r.LicencedTo
                d.OL_ExpiryDate = r.Termination
                d.OL_PhoneHomeDays = r.NextCheckRequiredDays
                If Not LicenceService.Save() Then SetMessage("licence.writeFailed", ToastKind.Warning)
                ShowMarketing(r.MarketingMessage)
                ReportValid()
                ' At start-up the window closes as soon as the licence is found to be valid (an
                ' expiry within a month is shown on the licence chip in the top bar instead).
                If fromStart AndAlso r.MarketingMessage = "" Then DialogResult = DialogResult.OK
            Else
                ' Not valid any more - try to activate again with the saved details (as before).
                SetBusy(False)
                Await ActivateOnlineAsync()
            End If
        Catch ex As LicenceServerUnavailableException
            Dim left = LicenceService.OfflineGraceDaysLeft
            If left > 1 Then
                SetMessageText(Loc.F("licence.graceDays", left), ToastKind.Warning)
                SetCanRun(True)
            ElseIf left = 1 Then
                SetMessage("licence.graceLastDay", ToastKind.Warning)
                SetCanRun(True)
            Else
                SetMessage("licence.serverUnavailable", ToastKind.Error)
                SetCanRun(False)
            End If
        Finally
            SetBusy(False)
            ShowPage()
            UpdateOnlineButtons()
        End Try
    End Function

    Private Async Function DeactivateOnlineAsync() As Task
        If Not Ask.YesNoWarning(Me, "licence.confirmDeactivateOnline") Then Return
        Dim d = LicenceService.Data
        Dim key As Guid, token As Guid
        If Not Guid.TryParse(d.OL_LicenceCode, key) OrElse Not Guid.TryParse(d.OL_ActivationClientToken, token) Then Return
        SetBusy(True, "licence.connecting")
        Try
            Dim r = Await LicenceApiClient.DeactivateAsync(key, token)
            If r.LicenceStatus = LicenceStatusType.Invalid Then
                SetMessageText(StatusText(r), ToastKind.Error)
            Else
                d.OL_ActivationClientToken = ""
                d.OL_Registered = False
                SetMessage(If(LicenceService.Save(), "licence.deactivated", "licence.deactivatedNotSaved"), ToastKind.Success)
                SetCanRun(False)
                MustCloseProgram()
            End If
        Catch ex As LicenceServerUnavailableException
            SetMessage("licence.deactivateNeedsInternet", ToastKind.Error)
        Finally
            SetBusy(False)
            ShowPage()
            UpdateOnlineButtons()
        End Try
    End Function

    ''' <summary>Shows the "valid until" message or the expiry warning. Returns True when close to expiry.</summary>
    Private Function ReportValid() As Boolean
        Dim d = LicenceService.Data
        Dim days = LicenceService.DaysUntil(d.OL_ExpiryDate)
        lblLicensedTo.Text = If(d.OL_LicencedTo = "", "", Loc.F("licence.licensedTo", d.OL_LicencedTo))
        SetCanRun(True)
        If days <= LicenceService.WarnDaysBeforeExpiry Then
            SetMessageText(ExpiryWarning(days, d.OL_ExpiryDate), ToastKind.Warning)
            Return True
        End If
        SetMessageText(Loc.F("licence.validUntil", d.OL_ExpiryDate.ToShortDateString()), ToastKind.Success)
        Return False
    End Function

    Private Sub UpdateOnlineButtons()
        Dim d = LicenceService.Data
        Dim sameAsSaved = txtKey.Text.Trim() = d.OL_LicenceCode AndAlso txtEmail.Text.Trim() = d.OL_Email
        Dim active = LicenceService.HasOnlineActivation
        btnActivate.Visible = Not (active AndAlso sameAsSaved)
        btnCheck.Visible = active AndAlso sameAsSaved
        btnDeactivateOnline.Visible = active
        btnActivate.Enabled = Not _busy AndAlso txtEmail.Text.Trim() <> "" AndAlso txtKey.Text.Trim() <> "" AndAlso LicenceService.HostnameMatches
        btnCheck.Enabled = Not _busy
        btnDeactivateOnline.Enabled = Not _busy
    End Sub

    Private Shared Function StatusText(r As ActivationStatus) As String
        Dim name = Loc.T("licence.status." & r.LicenceStatus.ToString().ToLowerInvariant())
        Return If(r.Message = "", name, name & ": " & r.Message)
    End Function

    Private Sub ShowMarketing(text As String)
        LicenceService.MarketingMessage = If(text, "").Trim()
        lblMarketing.Text = LicenceService.MarketingMessage
    End Sub

#End Region

#Region "Offline"

    Private Sub RefreshOffline(fromStart As Boolean)
        Dim d = LicenceService.Data
        lblSerial.Text = LicenceService.OfflineSerial()
        Dim expiry = LicenceService.ExpiryFromCode(d.Act_ActivationCode)
        lblOfflineExpiry.Text = If(expiry.HasValue, Loc.F("licence.expiryDate", expiry.Value.ToShortDateString()), "")
        btnDeactivateOffline.Visible = d.Act_Registered
        btnTrial.Visible = Not d.Act_Registered AndAlso _mode = LicenceDialogMode.StartUp

        If d.Act_Registered AndAlso Not LicenceService.OfflineCodeIsValid() Then
            d.Act_Registered = False
            LicenceService.Save()
        End If
        If d.Act_Registered AndAlso expiry.HasValue Then
            Dim days = LicenceService.DaysUntil(expiry.Value)
            If days < 0 Then
                d.Act_Registered = False
                LicenceService.Save()
                RefreshOffline(False)
                SetMessageText(Loc.F("licence.offlineExpired", expiry.Value.ToShortDateString()), ToastKind.Error)
                SetCanRun(False)
                Return
            End If
            SetCanRun(True)
            If days <= LicenceService.WarnDaysBeforeExpiry Then
                SetMessageText(ExpiryWarning(days, expiry.Value), ToastKind.Warning)
            Else
                SetMessageText(Loc.F("licence.validUntil", expiry.Value.ToShortDateString()), ToastKind.Success)
            End If
            If fromStart Then DialogResult = DialogResult.OK
        Else
            If Not fromStart OrElse lblMessage.Text = "" Then SetMessage("licence.offlineNotActivated", ToastKind.Info)
            SetCanRun(False)
        End If
        ShowPage()
    End Sub

    Private Sub ActivateOffline()
        Dim expiry As Date
        Select Case LicenceService.ActivateOffline(txtCode.Text, expiry)
            Case LicenceService.OfflineResult.WrongCode
                SetMessage("licence.wrongCode", ToastKind.Error)
            Case LicenceService.OfflineResult.Expired
                SetMessageText(Loc.F("licence.codeOutdated", expiry.ToShortDateString()), ToastKind.Error)
            Case LicenceService.OfflineResult.SaveFailed
                SetMessage("licence.writeFailed", ToastKind.Error)
            Case LicenceService.OfflineResult.Activated
                txtCode.Text = ""
                RefreshOffline(fromStart:=False)
                SetMessageText(Loc.F("licence.offlineActivated", expiry.ToShortDateString()), ToastKind.Success)
        End Select
    End Sub

    Private Sub DeactivateOffline()
        If Not Ask.YesNoWarning(Me, "licence.confirmDeactivateOffline") Then Return
        Dim code = LicenceService.DeactivateOffline()
        If code Is Nothing Then
            SetMessage("licence.writeFailed", ToastKind.Error)
            Return
        End If
        RefreshOffline(fromStart:=False)
        SetMessageText(Loc.F("licence.deactivationCode", code), ToastKind.Warning)
        SetCanRun(False)
        MustCloseProgram()
    End Sub

    Private Sub StartTrial()
        Dim left = LicenceService.UseTrial()
        If left > 0 Then
            Ask.Info(Me, "licence.trialDaysLeft", left)
            _canRun = True
            DialogResult = DialogResult.OK
        Else
            Ask.Error(Me, "licence.trialExpired")
        End If
    End Sub

#End Region

#Region "Helpers"

    Private Shared Function ExpiryWarning(days As Integer, expiry As Date) As String
        If days = 0 Then Return Loc.F("licence.expiresTomorrow", expiry.ToShortDateString())
        Return Loc.F("licence.expiresIn", days + 1, expiry.ToShortDateString())
    End Function

    Private Sub SetMessage(key As String, kind As ToastKind)
        SetMessageText(If(key = "", "", Loc.T(key)), kind)
    End Sub

    Private Sub SetMessageText(text As String, kind As ToastKind)
        Dim p = ThemeManager.Current
        lblMessage.Text = text
        lblMessage.ForeColor = If(kind = ToastKind.Error, p.Danger, If(kind = ToastKind.Warning, p.Warning, If(kind = ToastKind.Success, p.Success, p.Ink)))
    End Sub

    Private Sub SetBusy(busy As Boolean, Optional key As String = "")
        _busy = busy
        UseWaitCursor = busy
        If busy AndAlso key <> "" Then SetMessage(key, ToastKind.Info)
        segKind.Enabled = Not busy
        txtEmail.Enabled = Not busy
        txtKey.Enabled = Not busy
        btnContinue.Enabled = Not busy AndAlso (_canRun OrElse _mode = LicenceDialogMode.Info)
        UpdateOnlineButtons()
    End Sub

    Private Sub SetCanRun(value As Boolean)
        _canRun = value
        btnContinue.Enabled = (value OrElse _mode = LicenceDialogMode.Info) AndAlso Not _busy
        If _mode = LicenceDialogMode.StartUp Then AcceptButton = If(value, btnContinue, Nothing)
    End Sub

    ''' <summary>After deactivating from Help > Licence the program closes with this window (as before).</summary>
    Private _closeProgram As Boolean

    Private Sub MustCloseProgram()
        If _mode <> LicenceDialogMode.Info Then Return
        _closeProgram = True
        lblMarketing.Text = Loc.T("licence.willClose")
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        MyBase.OnFormClosed(e)
        If _closeProgram Then
            Dim main = AppHost.MainWindow
            If main IsNot Nothing AndAlso Not main.IsDisposed Then main.BeginInvoke(Sub() main.Close())
        End If
    End Sub

    Private Sub Lock(locked As Boolean)
        pageOnline.Enabled = Not locked
        pageOffline.Enabled = Not locked
        segKind.Enabled = Not locked
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If _busy AndAlso e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Return
        End If
        If _mode = LicenceDialogMode.StartUp Then
            ' Quit always closes the program. The window's X counts as Continue when the program may run.
            If _quit OrElse Not _canRun Then
                DialogResult = DialogResult.Cancel
            Else
                DialogResult = DialogResult.OK
            End If
        End If
        MyBase.OnFormClosing(e)
    End Sub

#End Region

End Class
