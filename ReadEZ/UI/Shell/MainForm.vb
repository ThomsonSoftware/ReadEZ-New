Option Strict On

' ====================================================================================================
'   MainForm.vb - the clinician's window
'
'   Layout
'     +------------------------------------------------------------------------+
'     |  Menu: File · Edit · Settings · About · Help                           |
'     +--------+---------------------------------------------------------------+
'     |        |  TopBar: patient · screener · licence · colours · screens · theme|
'     |  Nav   +---------------------------------------------------------------+
'     |  rail  |                                                               |
'     |        |  Current screen (Patients / Colours / Tests / Results ...)    |
'     |        |                                                               |
'     +--------+---------------------------------------------------------------+
'
'   The same window as Eye Dynamix. Replaces RZfrmResults (ReadEZ 2020's main window with its
'   ribbon); each page is its own ScreenBase in UI\Screens and this form only switches between them.
' ====================================================================================================

Public Class MainForm
    Inherits Form

    Private ReadOnly _nav As New NavRail
    Private ReadOnly _top As New TopBar
    Private ReadOnly _menu As New MainMenu
    Private ReadOnly _host As New Panel
    Private ReadOnly _screens As New Dictionary(Of String, ScreenBase)
    Private ReadOnly _factories As New Dictionary(Of String, Func(Of ScreenBase))
    Private _current As ScreenBase
    Private _currentId As String = ""

    ' Screen ids
    ''' <summary>Item in the menu on the left that opens the online help (not a screen).</summary>
    Private Const NavOnlineHelp As String = "onlineHelp"
    Public Const ScreenPatients As String = "patients"
    Public Const ScreenColours As String = "colours"
    Public Const ScreenTests As String = "tests"
    Public Const ScreenResults As String = "results"
    Public Const ScreenSettings As String = "settings"
    Public Const ScreenHelp As String = "help"
    Public Const ScreenRun As String = "run"

    Public Sub New()
        SuspendLayout()
        AutoScaleMode = AutoScaleMode.None   ' scaled once below (see Dpi.vb)
        Text = AppHost.ProductTitle
        Icon = AppHost.AppIcon
        Font = Fonts.Body
        KeyPreview = True
        ThemeManager.SetRole(Me, SurfaceRole.Canvas)

        BuildNavigation()
        RegisterScreens()

        _host.Dock = DockStyle.Fill
        ThemeManager.SetRole(_host, SurfaceRole.Canvas)

        Controls.Add(_host)
        Controls.Add(_top)
        Controls.Add(_nav)
        Controls.Add(_menu)          ' added last so it docks first: right across the window
        MainMenuStrip = _menu
        ResumeLayout(False)

        AddHandler _top.PatientClicked, Sub() ShowScreen(ScreenPatients)
        AddHandler _top.ColoursClicked, Sub() ShowScreen(ScreenColours)
        AddHandler _top.ScreensClicked, Sub() ShowScreen(ScreenSettings)
        AddHandler _top.ScreenerClicked, Sub() AppCommands.PickScreener()
        AddHandler _top.LicenceClicked, Sub() ShowLicence()
        AddHandler AppHost.Session.PatientChanged, Sub() UpdateNavAvailability()
        AddHandler AppHost.Session.ResultChanged, Sub() UpdateNavAvailability()

        RestoreWindow()
        ThemeManager.Apply(Me)
    End Sub

#Region "Navigation"

    Private Sub BuildNavigation()
        _nav.AddItem(ScreenPatients, "nav.patients", Glyphs.People)
        _nav.AddItem(ScreenTests, "nav.tests", Glyphs.Eye)
        _nav.AddItem(ScreenResults, "nav.results", Glyphs.Chart)
        _nav.AddItem(ScreenColours, "nav.colours", Glyphs.Color)
        _nav.AddItem(ScreenSettings, "nav.settings", Glyphs.Settings, atBottom:=True)
        _nav.AddItem(ScreenHelp, "nav.help", Glyphs.Help, atBottom:=True)
        _nav.AddItem(NavOnlineHelp, "menu.onlineHelp", Glyphs.Globe, atBottom:=True)     ' opens the online help (Admin > Online help)
        _nav.Collapsed = AppHost.Settings.NavCollapsed
        _nav.Logo = AppHost.LogoImage            ' Thomson Software Solutions, on a white panel at the foot of the rail
        AddHandler _nav.ItemClicked, Sub(s, id)
                                         If id = NavOnlineHelp Then MainMenu.OpenOnlineHelp() Else ShowScreen(id)
                                     End Sub
        AddHandler _nav.CollapsedChanged, Sub()
                                              AppHost.Settings.NavCollapsed = _nav.Collapsed
                                              AppHost.SaveSettings()
                                          End Sub
    End Sub

    ''' <summary>
    ''' Each screen is created the first time it is opened. Modules not yet ported show a
    ''' placeholder; replace the factory when the module is added.
    ''' </summary>
    Private Sub RegisterScreens()
        _factories(ScreenPatients) = Function() New PatientsScreen()
        _factories(ScreenColours) = Function() New ColoursScreen()
        _factories(ScreenTests) = Function() New TestsScreen()
        _factories(ScreenResults) = Function() New ResultsScreen()
        _factories(ScreenSettings) = Function() New SettingsScreen()
        _factories(ScreenHelp) = Function() New HelpScreen()
        _factories(ScreenRun) = Function() New TestRunScreen()
    End Sub

    Public ReadOnly Property CurrentScreenId As String
        Get
            Return _currentId
        End Get
    End Property

    ''' <summary>The screen showing now (Nothing before the first one is shown).</summary>
    Public ReadOnly Property CurrentScreen As ScreenBase
        Get
            Return _current
        End Get
    End Property

    ''' <summary>Switches to a screen. Returns False if the current screen refused to close.</summary>
    Public Function ShowScreen(id As String) As Boolean
        If id = _currentId Then Return True
        If _current IsNot Nothing AndAlso Not _current.CanLeave() Then Return False

        Dim screen = GetOrCreate(id)
        If screen Is Nothing Then Return False

        _host.SuspendLayout()
        If _current IsNot Nothing Then
            _current.Visible = False
            _current.OnDeactivated()
        End If
        _current = screen
        _currentId = id
        screen.Visible = True
        screen.BringToFront()
        _host.ResumeLayout(True)
        _nav.SelectedId = If(id = ScreenRun, ScreenTests, id)    ' a running test belongs to "Run tests"
        screen.OnActivated()
        Return True
    End Function

    ''' <summary>Opens the test screen with a test (its steps run one after the other).</summary>
    Public Sub RunTest(t As TestRun)
        If t Is Nothing Then Return
        If _currentId = ScreenRun Then
            If Not ShowScreen(ScreenTests) Then Return     ' leave the running test cleanly first
        End If
        Dim run = DirectCast(GetOrCreate(ScreenRun), TestRunScreen)
        run.StartRun(t)
        ShowScreen(ScreenRun)
    End Sub

    Private Function GetOrCreate(id As String) As ScreenBase
        Dim screen As ScreenBase = Nothing
        If _screens.TryGetValue(id, screen) Then Return screen
        Dim factory As Func(Of ScreenBase) = Nothing
        If Not _factories.TryGetValue(id, factory) Then Return Nothing
        screen = factory()
        Dpi.ScaleOnce(screen)
        _screens(id) = screen
        screen.Visible = False
        _host.Controls.Add(screen)
        ThemeManager.Apply(screen)
        Return screen
    End Function

    Private Sub UpdateNavAvailability()
        _nav.SetEnabled(ScreenResults, True)
    End Sub

    ''' <summary>Runs an action on the UI thread (events from the tracker arrive on its own thread).</summary>
    Public Sub OnUiThread(action As Action)
        If IsDisposed Then Return
        If InvokeRequired Then
            BeginInvoke(action)
        Else
            action()
        End If
    End Sub

#End Region

#Region "Window position"

    ''' <summary>
    ''' Start-up placement. The control screen from Settings decides which monitor the window opens
    ''' on. The last size and position (and maximised state) are reused only if they are on that
    ''' same monitor, so changing the control screen in Settings always wins.
    ''' The window handle is created on the target monitor BEFORE scaling, so the interface is
    ''' scaled for that monitor's DPI and not the primary monitor's.
    ''' </summary>
    Private Sub RestoreWindow()
        Dim s = AppHost.Settings
        StartPosition = FormStartPosition.Manual
        Dim target = ControlTarget()
        Dim remembered = s.MainWindowBounds
        Dim useRemembered = ScreenLayout.IsOnScreen(remembered) AndAlso
                            ScreenLayout.ScreenOf(remembered).DeviceName = ScreenLayout.ScreenOf(target).DeviceName
        Dim r = If(useRemembered, remembered, target)
        Me.Bounds = r
        If Not IsHandleCreated Then CreateHandle()      ' DeviceDpi now belongs to the target monitor
        Dpi.ScaleOnce(Me)
        MinimumSize = New Size(Math.Min(Dpi.Px(Me, 1100), target.Width), Math.Min(Dpi.Px(Me, 680), target.Height))
        Me.Bounds = r                                   ' scaling changes the size; put it back
        If s.MainWindowMaximised Then WindowState = FormWindowState.Maximized
    End Sub

    ''' <summary>The control-screen rectangle from Settings, or the primary working area if it is not usable.</summary>
    Private Shared Function ControlTarget() As Rectangle
        Dim c = AppHost.Settings.ControlBounds
        Return If(ScreenLayout.IsOnScreen(c), c, Screen.PrimaryScreen.WorkingArea)
    End Function

    ''' <summary>Moves the window to the control screen after it has been changed in Settings.</summary>
    Public Sub PlaceOnControlScreen()
        Dim s = AppHost.Settings
        Dim target = ControlTarget()
        Dim oldDpi = DeviceDpi
        Dim wasMax = WindowState = FormWindowState.Maximized
        s.MainWindowBounds = Rectangle.Empty
        ' A maximised window must be restored before it can move; it is maximised again on the new monitor.
        If WindowState <> FormWindowState.Normal Then WindowState = FormWindowState.Normal
        MinimumSize = New Size(Math.Min(MinimumSize.Width, target.Width), Math.Min(MinimumSize.Height, target.Height))
        Bounds = target
        If wasMax Then WindowState = FormWindowState.Maximized
        s.MainWindowMaximised = wasMax
        If DeviceDpi <> oldDpi Then Toast.Show(Me, Loc.T("settings.restartForScaling"), ToastKind.Info)
    End Sub

    ''' <summary>
    ''' Sizes in this program are scaled once (see Dpi.vb), so when the window crosses to a monitor
    ''' with different scaling it keeps its own size instead of the size Windows suggests.
    ''' </summary>
#If Not COMPILECHECK Then
    Protected Overrides Sub OnDpiChanged(e As DpiChangedEventArgs)
        e.Cancel = True
        MyBase.OnDpiChanged(e)
    End Sub
#End If

    Private Sub SaveWindow()
        Dim s = AppHost.Settings
        s.MainWindowMaximised = WindowState = FormWindowState.Maximized
        s.MainWindowBounds = If(WindowState = FormWindowState.Normal, Bounds, RestoreBounds)
    End Sub

#End Region

#Region "Lifetime"

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        SplashForm.Finish(Me)                 ' stays above this window until it has shown for 2 s, then fades
        ShowScreen(ScreenPatients)
        Dim unused = Stimulus.Window          ' the patient's screen is shown from the start, as in Eye Dynamix
        Activate()
        WarmUpInBackground()
        CheckLicenceInBackground()
    End Sub

    ''' <summary>
    ''' Work that isn't needed to show the first screen, done once the window is up:
    '''   - reading texts, letter paragraphs and report pictures copied into the data folder (first run)
    ''' </summary>
    Private Shared Sub WarmUpInBackground()
        Task.Run(Sub() StandardFiles.EnsureInstalled())
    End Sub

    ''' <summary>
    ''' As D-Chart: an online licence is checked with the server once the window is showing. If the
    ''' server says it is no longer valid the licence window is shown, and the program closes unless
    ''' it is activated again.
    ''' </summary>
    Private Async Sub CheckLicenceInBackground()
        Dim ok As Boolean
        Try
            ok = Await LicenceService.CheckOnlineInBackgroundAsync()
        Catch ex As Exception
            Log.Warn("Background licence check: " & ex.Message)
            ok = True
        End Try
        If IsDisposed Then Return
        If Not ok Then
            Using dlg As New LicenceDialog(LicenceDialogMode.StartUp)
                If dlg.ShowDialog(Me) <> DialogResult.OK Then
                    _top.RefreshLicence()
                    Close()
                    Return
                End If
            End Using
        End If
        _top.RefreshLicence()
        CheckForUpdatesInBackground()
    End Sub

    ''' <summary>
    ''' Looks for a new version on the web (Settings.Update_CheckAtStart). If there is one, and it
    ''' is not a version the user chose to skip, offers to install it. Silent if the server can't
    ''' be reached.
    ''' </summary>
    Private Async Sub CheckForUpdatesInBackground()
        If Not AppHost.Settings.Update_CheckAtStart Then Return
        Dim info As UpdateInfo
        Try
            info = Await UpdateService.CheckAsync()
        Catch ex As Exception
            Log.Warn("Update check at start-up: " & ex.Message)
            Return
        End Try
        If IsDisposed Then Return
        If info Is Nothing Then
            Log.Info("Update check at start-up: no newer version")
            Return
        End If
        If String.Equals(info.VersionText, AppHost.Settings.Update_SkipVersion, StringComparison.OrdinalIgnoreCase) Then
            Log.Info($"Update check at start-up: {info.VersionText} available but skipped by the user")
            Return
        End If
        ' Offered when nothing else is in the way: not while a test is running or another window is
        ' open - the splash screen counts, as it is owned by this window while it fades. Before, the
        ' offer was simply dropped if the check finished while the splash was still showing, which
        ' was nearly always, so updates were never offered at start-up. Now it waits (up to 10 min).
        Dim busy = Function() _currentId = ScreenRun OrElse OwnedForms.Any(Function(f) f.Visible)
        Dim waited = 0
        While busy() AndAlso waited < 600
            Await Task.Delay(1000)
            waited += 1
            If IsDisposed Then Return
        End While
        If busy() Then
            Log.Info($"Update check at start-up: {info.VersionText} available, not offered (the program was busy)")
            Return
        End If
        Log.Info($"Update check at start-up: offering {info.VersionText}")
        ShowUpdate(info)
    End Sub

    ''' <summary>Help > Check for updates (info = Nothing checks now), or the start-up offer.</summary>
    Public Sub ShowUpdate(Optional info As UpdateInfo = Nothing)
        Using dlg As New UpdateDialog(info)
            dlg.ShowDialog(Me)
        End Using
    End Sub

    ''' <summary>About > About ReadEZ.</summary>
    Public Sub ShowAbout()
        Using dlg As New AboutDialog()
            dlg.ShowDialog(Me)
        End Using
    End Sub

    ''' <summary>Refreshes the top bar's screener, colour and screen chips.</summary>
    Public Sub RefreshStatus()
        _top.RefreshStatus()
    End Sub

    ''' <summary>Licence chip in the top bar: shows the licence, to check, change or deactivate it.</summary>
    Public Sub ShowLicence()
        Using dlg As New LicenceDialog(LicenceDialogMode.Info)
            dlg.ShowDialog(Me)
        End Using
        _top.RefreshLicence()
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If _current IsNot Nothing AndAlso Not _current.CanLeave() Then
            e.Cancel = True
            Return
        End If
        SaveWindow()
        AppHost.SaveSettings()
        ' Backup on closing (Settings > Data), if anything changed since the last one
        If Not e.Cancel Then
            Try
                BackupCommands.BackUpOnClose(Me)
            Catch ex As Exception
                Log.Error("Backup on close", ex)
            End Try
        End If
        MyBase.OnFormClosing(e)
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Stimulus.CloseWindow()
        MyBase.OnFormClosed(e)
    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Control Or Keys.D1 : ShowScreen(ScreenPatients) : Return True
            Case Keys.Control Or Keys.D2 : ShowScreen(ScreenTests) : Return True
            Case Keys.Control Or Keys.D3 : ShowScreen(ScreenResults) : Return True
            Case Keys.Control Or Keys.D4 : ShowScreen(ScreenColours) : Return True
            Case Keys.F1
                ' Help for the screen showing, in the help file; the Help screen if it isn't installed
                If Not HelpFile.Show(Me, If(_current?.HelpTopicNow(), "")) Then ShowScreen(ScreenHelp)
                Return True
        End Select
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

#End Region

End Class
