Option Strict On

' ====================================================================================================
'   MainMenu.vb - the menu bar across the top of the main window, and the commands it runs
'
'     File       Back up now, Restore a backup, Import a ReadEZ 2020 database, Exit
'     Edit       New patient, Edit patient, Delete patient,
'                Edit notes of the test, Move the test to another patient, Delete the test,
'                Organisations, Screeners
'     Settings   Program options (the Settings screen), Colour calibration
'     About      About ReadEZ, Licence, Check for updates
'     Help       the help videos in groups (HelpCatalog), the help file and the Help screen
'
'   The patient and result commands live in AppCommands so that the menu and the buttons on the
'   Patients and Results screens do exactly the same thing. "The test" is the one selected
'   (Session.ResultId).
' ====================================================================================================

Public NotInheritable Class MainMenu
    Inherits MenuStrip

    Private ReadOnly _editMenu As New ToolStripMenuItem
    Private ReadOnly _newPatient As New ToolStripMenuItem
    Private ReadOnly _editPatient As New ToolStripMenuItem
    Private ReadOnly _deletePatient As New ToolStripMenuItem
    Private ReadOnly _editResult As New ToolStripMenuItem
    Private ReadOnly _moveResult As New ToolStripMenuItem
    Private ReadOnly _deleteResult As New ToolStripMenuItem

    Public Sub New()
        Dock = DockStyle.Top
        Font = Fonts.Body
        Padding = New Padding(8, 3, 0, 3)
        GripStyle = ToolStripGripStyle.Hidden
        Renderer = New ThemedToolStripRenderer()

        ' File
        Dim file = TopMenu("menu.file")
        MenuItem(file, "menu.backupNow", Sub() BackupCommands.BackUpNow(AppHost.MainWindow))
        MenuItem(file, "menu.restoreBackup", Sub() BackupCommands.Restore(AppHost.MainWindow))
        file.DropDownItems.Add(New ToolStripSeparator)
        MenuItem(file, "menu.importOld", Sub() DatabaseCommands.ImportOldDatabase(AppHost.MainWindow))
        file.DropDownItems.Add(New ToolStripSeparator)
        Dim exitItem = MenuItem(file, "menu.exit", Sub() AppHost.MainWindow?.Close())
        exitItem.ShortcutKeyDisplayString = "Alt+F4"

        ' Edit
        Loc.Bind(_editMenu, "menu.edit")
        Items.Add(_editMenu)
        AddTo(_editMenu, _newPatient, "menu.newPatient", Sub() AppCommands.NewPatient())
        AddTo(_editMenu, _editPatient, "menu.editPatient", Sub() AppCommands.EditPatient())
        AddTo(_editMenu, _deletePatient, "menu.deletePatient", Sub() AppCommands.DeletePatient())
        _editMenu.DropDownItems.Add(New ToolStripSeparator)
        AddTo(_editMenu, _editResult, "menu.editResult", Sub() AppCommands.EditResult())
        AddTo(_editMenu, _moveResult, "menu.moveResult", Sub() AppCommands.MoveResult())
        AddTo(_editMenu, _deleteResult, "menu.deleteResult", Sub() AppCommands.DeleteResult())
        _editMenu.DropDownItems.Add(New ToolStripSeparator)
        MenuItem(_editMenu, "menu.organisations", Sub() AppCommands.EditOrganisations())
        MenuItem(_editMenu, "menu.screeners", Sub() AppCommands.EditScreeners())
        AddHandler _editMenu.DropDownOpening, Sub() UpdateEditItems()

        ' Settings: the program's own options and the colour calibration
        Dim settings = TopMenu("nav.settings")
        MenuItem(settings, "menu.programOptions", Sub() AppHost.MainWindow?.ShowScreen(MainForm.ScreenSettings))
        MenuItem(settings, "menu.colourCalibration", Sub() ReadEZColourCheck.Run(AppHost.MainWindow, ""))

        ' About
        Dim about = TopMenu("menu.about")
        MenuItem(about, "menu.aboutProgram", Sub() AppHost.MainWindow?.ShowAbout())
        MenuItem(about, "help.licence", Sub() AppHost.MainWindow?.ShowLicence())
        MenuItem(about, "menu.checkUpdates", Sub() AppHost.MainWindow?.ShowUpdate())

        ' Help: the help topics in their groups (HelpCatalog), then the Help screen. Rebuilt when the
        ' topics are changed in Admin or the language changes (topic titles can be typed in Admin).
        _help = TopMenu("nav.help")
        BuildHelpMenu()
        AddHandler _help.DropDownOpening, Sub() MarkVideos(_help)
        AddHandler HelpCatalog.TopicsChanged, Sub() BuildHelpMenu()
        AddHandler Loc.LanguageChanged, Sub() BuildHelpMenu()
    End Sub

    Private _help As ToolStripMenuItem

    Private Sub BuildHelpMenu()
        If _help Is Nothing Then Return
        For Each item In _help.DropDownItems.Cast(Of ToolStripItem)().ToList()
            item.Dispose()
        Next
        _help.DropDownItems.Clear()
        For Each g In HelpCatalog.Groups
            Dim sub1 As New ToolStripMenuItem(g.Title)
            _help.DropDownItems.Add(sub1)
            AddHelpTopics(sub1, g.Children)
        Next
        _help.DropDownItems.Add(New ToolStripSeparator)
        ' The full help: online (address set in Admin) and, if it is installed, the CHM file
        MenuItem(_help, "menu.onlineHelp", Sub() OpenOnlineHelp())
        ' F1 opens the help file at the topic for the screen showing, or the Help screen without it
        If HelpFile.Available Then
            Dim contents = MenuItem(_help, "menu.helpContents", Sub() HelpFile.Show(AppHost.MainWindow))
            contents.ShortcutKeyDisplayString = "F1"
        End If
        Dim helpItem = MenuItem(_help, "menu.helpScreen", Sub() AppHost.MainWindow?.ShowScreen(MainForm.ScreenHelp))
        If Not HelpFile.Available Then helpItem.ShortcutKeyDisplayString = "F1"
    End Sub

    ''' <summary>Opens the online help (Admin > Online help); says where to set it if it isn't set.</summary>
    Public Shared Sub OpenOnlineHelp()
        Dim url = NormaliseUrl(AppHost.Admin.OnlineHelpUrl)
        If url = "" Then
            Ask.Info(AppHost.MainWindow, "help.online.notSet")
            Return
        End If
        AppHost.OpenUrl(url)
    End Sub

    ''' <summary>A web address with https:// added if no scheme was typed ("" stays "").</summary>
    Public Shared Function NormaliseUrl(url As String) As String
        Dim u = If(url, "").Trim()
        If u = "" Then Return ""
        If Not u.Contains("://") Then u = "https://" & u
        Return u
    End Function

    Private Shared Function TopicItem(topic As HelpTopic) As ToolStripMenuItem
        Dim item As New ToolStripMenuItem(topic.Title) With {.Tag = topic}
        AddHandler item.Click, Sub()
                                   Try
                                       HelpCatalog.OpenVideo(AppHost.MainWindow, topic.Id)
                                   Catch ex As Exception
                                       Log.Error("Help menu", ex)
                                   End Try
                               End Sub
        Return item
    End Function

    ''' <summary>
    ''' One item per topic; a topic with sub-topics becomes a sub-menu that starts with the topic
    ''' itself. Clicking plays the video (or says there isn't one yet).
    ''' </summary>
    Private Shared Sub AddHelpTopics(parent As ToolStripMenuItem, topics As IEnumerable(Of HelpTopic))
        For Each t In topics
            Dim topic = t
            If topic.Children.Count = 0 Then
                parent.DropDownItems.Add(TopicItem(topic))
            Else
                Dim sub1 As New ToolStripMenuItem(topic.Title)
                parent.DropDownItems.Add(sub1)
                sub1.DropDownItems.Add(TopicItem(topic))
                sub1.DropDownItems.Add(New ToolStripSeparator)
                AddHelpTopics(sub1, topic.Children)
            End If
        Next
    End Sub

    ''' <summary>Topics without a video yet say so (and are drawn muted by ThemedToolStripRenderer).</summary>
    Private Shared Sub MarkVideos(menu As ToolStripMenuItem)
        For Each item In menu.DropDownItems.OfType(Of ToolStripMenuItem)()
            Dim topic = TryCast(item.Tag, HelpTopic)
            If topic IsNot Nothing Then
                item.ShortcutKeyDisplayString = If(topic.HasVideo, "", Loc.T("help.comingSoon"))
                item.ShowShortcutKeys = Not topic.HasVideo
            End If
            If item.HasDropDownItems Then MarkVideos(item)
        Next
    End Sub

    Private Function TopMenu(key As String) As ToolStripMenuItem
        Dim m As New ToolStripMenuItem
        Loc.Bind(m, key)
        Items.Add(m)
        Return m
    End Function

    Private Shared Function MenuItem(parent As ToolStripMenuItem, key As String, action As Action) As ToolStripMenuItem
        Dim i As New ToolStripMenuItem
        AddTo(parent, i, key, action)
        Return i
    End Function

    Private Shared Sub AddTo(parent As ToolStripMenuItem, i As ToolStripMenuItem, key As String, action As Action)
        Loc.Bind(i, key)
        AddHandler i.Click, Sub()
                                Try
                                    action()
                                Catch ex As Exception
                                    Log.Error("Menu " & key, ex)
                                    Ask.Error(AppHost.MainWindow, "common.errorWithDetail", ex.Message)
                                End Try
                            End Sub
        parent.DropDownItems.Add(i)
    End Sub

    ''' <summary>Only what makes sense for the current patient and selection is enabled.</summary>
    Private Sub UpdateEditItems()
        ' Nothing is changed while a test is running
        Dim idle = AppHost.MainWindow Is Nothing OrElse AppHost.MainWindow.CurrentScreenId <> MainForm.ScreenRun
        Dim hasPatient = AppHost.Session.HasPatient
        Dim hasResult = AppHost.Session.HasResult
        _newPatient.Enabled = idle
        _editPatient.Enabled = idle AndAlso hasPatient
        _deletePatient.Enabled = idle AndAlso hasPatient
        _editResult.Enabled = idle AndAlso hasResult
        _moveResult.Enabled = idle AndAlso hasResult
        _deleteResult.Enabled = idle AndAlso hasResult
    End Sub
End Class

''' <summary>About > About ReadEZ.</summary>
Public NotInheritable Class AboutDialog
    Inherits DialogBase

    Public Sub New()
        SetTitle("about.title")
        ' The start-up label (SplashScreen.vb), with the UDI and build underneath
        ClientSize = New Size(SplashArt.LogicalSize.Width + 48, SplashArt.LogicalSize.Height + 76 + 68 + 52)
        Body.Padding = New Padding(24, 0, 24, 8)
        Dim label As New SplashArtPanel With {.Dock = DockStyle.Top, .Height = SplashArt.LogicalSize.Height}
        Dim text As New Label With {.Dock = DockStyle.Fill, .AutoSize = False, .Font = Fonts.Small, .TextAlign = ContentAlignment.MiddleLeft}
        LabelStyle.SetMuted(text)
        Loc.OnChange(text, Sub() text.Text = Loc.F("about.details", AppHost.UdiText, GetType(AboutDialog).Assembly.GetName().Version))
        Body.Controls.Add(text)
        Body.Controls.Add(label)

        Dim close = AddFooterButton("common.close", ButtonKind.Primary)
        close.DialogResult = DialogResult.Cancel
        CancelButton = close
        AcceptButton = close
        Dim eula = AddFooterButton("eula.button", ButtonKind.Secondary, Glyphs.Report)
        AddHandler eula.Click, Sub()
                                   Using dlg As New EulaDialog()
                                       dlg.ShowDialog(Me)
                                   End Using
                               End Sub
    End Sub
End Class
