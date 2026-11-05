Option Strict On

Imports System.Data.SQLite
Imports System.Drawing.Drawing2D
Imports System.IO

' ====================================================================================================
'   SettingsScreen.vb - program settings (replaces RZOptions, as Eye Dynamix's Settings screen)
'
'   Sections: General (report, size of the patient's screen) · Appearance and language · Data
'             (database, data folder, backup, import from ReadEZ 2020) · Screens (one or two) ·
'             Test setup (reading texts, timings, ReadEZ colours) · Admin
'
'   Changes are applied when Save is pressed. Unlike the previous version the program no longer
'   restarts itself: the theme and language change at once, and a different database is opened
'   in place.
' ====================================================================================================

Public Class SettingsScreen
    Inherits ScreenBase

    Private ReadOnly _tabs As New SegmentedControl
    Private ReadOnly _pages As New List(Of Control)
    Private ReadOnly _host As New Panel

    ' General
    Private ReadOnly txtPractice As New TextField
    Private ReadOnly chkLogo As ThemedCheckBox = ThemedCheckBox.Create("settings.showLogo")
    Private ReadOnly txtLogo As New TextField
    Private ReadOnly pctLogo As New PictureBox
    Private ReadOnly nudLogoScale As NumericUpDown = Ui.NumberBox(10, 400, 100)
    Private ReadOnly nudScreenW As NumericUpDown = Ui.NumberBox(100, 3000, 500)
    Private ReadOnly nudScreenH As NumericUpDown = Ui.NumberBox(100, 2000, 300)

    ' Appearance and language
    Private ReadOnly segTheme As New SegmentedControl
    Private ReadOnly cmbLanguage As New ThemedComboBox

    ' Data
    Private ReadOnly txtDatabase As New TextField
    Private ReadOnly txtDataDir As New TextField

    ' Backup (saved as soon as changed)
    Private ReadOnly txtBackupFolder As New TextField
    Private ReadOnly segBackupWhen As New SegmentedControl
    Private ReadOnly chkBackupStimuli As ThemedCheckBox = ThemedCheckBox.Create("backup.includeStimuli")
    Private ReadOnly lblLastBackup As New Label With {.AutoSize = True, .Font = Fonts.Body}
    Private _loadingBackup As Boolean

    ' Reading texts (the data folder's own; the built-in texts are listed but can't be changed)
    Private ReadOnly lstReading As New ListBox With {.Width = 320, .Height = 190, .IntegralHeight = False, .BorderStyle = BorderStyle.FixedSingle, .Font = Fonts.Body}
    ' Timings
    Private ReadOnly nudReadTime As NumericUpDown = Ui.NumberBox(10, 300, 60)
    Private ReadOnly nudTiming As NumericUpDown = Ui.NumberBox(0.5D, 10, 3, decimals:=1)
    Private ReadOnly chkBeep As ThemedCheckBox = ThemedCheckBox.Create("settings.beep")
    Private ReadOnly segScreens As New SegmentedControl
    Private ReadOnly cboOneMonitor As New ThemedComboBox With {.Width = 300, .DropDownStyle = ComboBoxStyle.DropDownList}
    Private _oneMonitorStack As Control
    ' Two screens / control screen and tablet: which monitor is which
    Private ReadOnly cboControlMon As New ThemedComboBox With {.Width = 300, .DropDownStyle = ComboBoxStyle.DropDownList}
    Private ReadOnly cboPatientMon As New ThemedComboBox With {.Width = 300, .DropDownStyle = ComboBoxStyle.DropDownList}
    Private _controlStack, _patientStack, _tabletPanel As Control
    Private _controlCard, _patientCard As Card
    Private _syncing, _loadingScreens As Boolean
    ''' <summary>Set while Tablet devices is open from here, so the choices not yet saved are kept.</summary>
    Private _keepEdits As Boolean
    ' Settings > Screens > "Screens to use"
    Private Const ModeOne As Integer = 0, ModeTwo As Integer = 1, ModeTablet As Integer = 2

    ' ReadEZ colours
    Private ReadOnly lblReadEZ As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(560, 0), .Margin = New Padding(0, 8, 0, 6)}

    ' Screens
    Private ReadOnly preview As New ScreenArrangementPreview
    Private ReadOnly ctl As New RectEditor(fullScreen:=False)
    Private ReadOnly stim As New RectEditor(fullScreen:=True)
    Private ReadOnly lblMode As New Label With {.AutoSize = True, .Font = Fonts.Strong, .MaximumSize = New Size(620, 0), .Margin = New Padding(0, 8, 0, 0)}
    Private ReadOnly lblTablet As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(620, 0), .Margin = New Padding(0, 6, 0, 8)}

    Public Sub New()
        TitleKey = "settings.title"
        HelpTopic = "screen.settings"
        SubtitleKey = "settings.subtitle"

        Dim save = AddHeaderButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim revert = AddHeaderButton("settings.revert", ButtonKind.Secondary, Glyphs.Refresh)
        AddHandler save.Click, Sub() SaveAll()
        AddHandler revert.Click, Sub() LoadAll()

        _tabs.SetOptions("settings.tab.general", "settings.tab.appearance", "settings.tab.data", "settings.tab.screens", "settings.tab.testSetup", "settings.tab.admin")
        _tabs.Dock = DockStyle.Top
        AddHandler _tabs.SelectedIndexChanged, Sub() ShowPage(_tabs.SelectedIndex)
        Dim tabRow As New Panel With {.Dock = DockStyle.Top, .Height = 50, .BackColor = Color.Transparent}
        _tabs.Location = New Point(0, 0)
        _tabs.Dock = DockStyle.None
        tabRow.Controls.Add(_tabs)

        _host.Dock = DockStyle.Fill
        _host.AutoScroll = True
        _host.BackColor = Color.Transparent

        _pages.Add(BuildGeneral())
        _pages.Add(BuildAppearance())
        _pages.Add(BuildData())
        _pages.Add(BuildScreens())
        ' Shown top to bottom: timings, reading texts, ReadEZ colours (Page shows its cards last first)
        _pages.Add(Page(BuildReadEZ(), BuildTexts(), BuildTimings()))
        _pages.Add(BuildAdmin())
        For Each p In _pages
            p.Visible = False
            _host.Controls.Add(p)
        Next

        Content.Controls.Add(_host)
        Content.Controls.Add(tabRow)
        ShowPage(0)

        ' The theme and language can also be changed from the top bar while this screen is open
        AddHandler ThemeManager.ThemeChanged, Sub() If segTheme.SelectedIndex <> CInt(ThemeManager.Mode) Then segTheme.SelectedIndex = CInt(ThemeManager.Mode)
        AddHandler Loc.LanguageChanged, Sub()
                                            _settingLanguage = True
                                            cmbLanguage.SelectedIndex = CInt(Loc.Language)
                                            _settingLanguage = False
                                        End Sub
    End Sub

    Private Sub ShowPage(i As Integer)
        For n = 0 To _pages.Count - 1
            _pages(n).Visible = n = i
        Next
    End Sub

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        If _keepEdits Then
            ' Back from Tablet devices: keep what was being set up here
            _keepEdits = False
            ShowTablet()
            ShowMode()
        Else
            LoadAll()
        End If
    End Sub

    ''' <summary>Shows the Screens tab (top bar's screens chip, Tablet devices' Back).</summary>
    Public Sub ShowScreensTab()
        _tabs.SelectedIndex = 3
        ShowPage(3)
    End Sub

    Private Shared Function PageCard(titleKey As String) As Card
        Dim c As New Card With {.Dock = DockStyle.Top, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                                .Padding = New Padding(20, 44, 20, 16), .Margin = New Padding(0, 0, 0, 12), .MinimumSize = New Size(600, 0)}
        c.SetTitle(titleKey)
        Return c
    End Function

    Private Shared Function Page(ParamArray cards As Control()) As Control
        Dim p As New Panel With {.Dock = DockStyle.Top, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink, .BackColor = Color.Transparent}
        For i = cards.Length - 1 To 0 Step -1
            Dim spacer As New Panel With {.Dock = DockStyle.Top, .Height = 12, .BackColor = Color.Transparent}
            p.Controls.Add(spacer)
            p.Controls.Add(cards(i))
            cards(i).BringToFront()
        Next
        Return p
    End Function

#Region "Pages"

    Private Function BuildGeneral() As Control
        Dim report = PageCard("settings.section.report")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        col.Controls.Add(Ui.Stack("settings.practiceName", txtPractice, 420))
        col.Controls.Add(chkLogo)
        Dim logoRow = Ui.Row()
        logoRow.Controls.Add(Ui.Stack("settings.logoFile", txtLogo, 420))
        Dim browse = ModernButton.Create("common.browse", ButtonKind.Secondary, Glyphs.Open)
        browse.Margin = New Padding(0, 22, 0, 0)
        AddHandler browse.Click, Sub() BrowseLogo()
        logoRow.Controls.Add(browse)
        col.Controls.Add(logoRow)
        col.Controls.Add(Ui.Stack("settings.logoScale", nudLogoScale))
        pctLogo.Size = New Size(220, 80)
        pctLogo.SizeMode = PictureBoxSizeMode.Zoom
        pctLogo.BorderStyle = BorderStyle.None
        col.Controls.Add(pctLogo)
        report.Controls.Add(col)

        Dim display = PageCard("settings.section.stimulusScreen")
        Dim dcol = Ui.Column()
        dcol.Dock = DockStyle.Top
        Dim hint = Ui.Label("settings.screenSize.hint", muted:=True)
        hint.MaximumSize = New Size(560, 0)
        dcol.Controls.Add(hint)
        Dim sizeRow = Ui.Row()
        sizeRow.Controls.Add(Ui.Stack("settings.screenWidthMm", nudScreenW))
        sizeRow.Controls.Add(Ui.Stack("settings.screenHeightMm", nudScreenH))
        dcol.Controls.Add(sizeRow)
        display.Controls.Add(dcol)
        Return Page(report, display)
    End Function

    ''' <summary>
    ''' The texts used by the custom test: reading texts (Reading, one line per row) and documents
    ''' (Text, RTF). Moved here from the test's options so they are looked after in one place.
    ''' </summary>
    Private Function BuildTexts() As Control
        Dim card = PageCard("settings.section.readingTexts")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        Dim intro = Ui.Label("settings.texts.intro", muted:=True)
        intro.MaximumSize = New Size(620, 0)
        intro.Margin = New Padding(0, 0, 0, 8)
        col.Controls.Add(intro)
        lstReading.BackColor = ThemeManager.Current.InputBack
        lstReading.ForeColor = ThemeManager.Current.Ink
        Dim newText = ModernButton.Create("settings.texts.new", ButtonKind.Secondary, Glyphs.Add)
        Dim editText = ModernButton.Create("settings.texts.edit", ButtonKind.Secondary, Glyphs.Edit)
        Dim delText = ModernButton.Create("settings.texts.delete", ButtonKind.Ghost, Glyphs.Delete)
        col.Controls.Add(ListWithButtons("settings.texts.reading", lstReading, newText, editText, delText))
        AddHandler newText.Click, Sub() EditReadingText(Nothing)
        AddHandler editText.Click, Sub() EditSelectedText()
        AddHandler lstReading.DoubleClick, Sub() EditSelectedText()
        AddHandler delText.Click, Sub()
                                      Dim t = SelectedUserText()
                                      If t Is Nothing Then Return
                                      If DeleteFile(ReadingTexts.UserFolder, t & ".txt") Then FillTexts()
                                  End Sub
        card.Controls.Add(col)
        Return card
    End Function

    ''' <summary>Reading time, beep, presentation time, one-screen full window.</summary>
    Private Function BuildTimings() As Control
        Dim card = PageCard("settings.section.timings")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        Dim row = Ui.Row()
        row.Controls.Add(Ui.Stack("settings.readTime", nudReadTime))
        row.Controls.Add(Ui.Stack("settings.timing", nudTiming))
        col.Controls.Add(row)
        col.Controls.Add(chkBeep)
        card.Controls.Add(col)
        Return card
    End Function

    ''' <summary>The selected text if it is one of the data folder's own (built-in texts can't be changed).</summary>
    Private Function SelectedUserText() As String
        Dim i = lstReading.SelectedIndex
        If i < ReadingTexts.BuiltInCount Then
            If i >= 0 Then Ask.Info(FindForm(), "settings.texts.builtIn")
            Return Nothing
        End If
        Return CStr(lstReading.SelectedItem)
    End Function

    Private Sub EditSelectedText()
        Dim t = SelectedUserText()
        If t IsNot Nothing Then EditReadingText(t)
    End Sub

    Private Shared Function ListWithButtons(labelKey As String, list As ListBox, ParamArray buttons As ModernButton()) As Control
        Dim r = Ui.Row()
        r.Margin = New Padding(0, 0, 32, 8)
        r.Controls.Add(Ui.Stack(labelKey, list))
        Dim c = Ui.Column()
        c.Margin = New Padding(0, 22, 0, 0)
        For Each b In buttons
            b.Margin = New Padding(0, 0, 0, 6)
            c.Controls.Add(b)
        Next
        r.Controls.Add(c)
        Return r
    End Function

    Private Sub FillTexts()
        Dim keep = lstReading.SelectedIndex
        lstReading.Items.Clear()
        For Each t In ReadingTexts.All()
            lstReading.Items.Add(t.Name)
        Next
        If keep >= 0 AndAlso keep < lstReading.Items.Count Then lstReading.SelectedIndex = keep
    End Sub

    Private Sub EditReadingText(name As String)
        Dim text = ""
        If name IsNot Nothing Then
            Try
                text = File.ReadAllText(Path.Combine(ReadingTexts.UserFolder, name & ".txt"))
            Catch
            End Try
        End If
        Using d As New ReadingTextDialog(name, text)
            If d.ShowDialog(FindForm()) <> DialogResult.OK Then Return
            Try
                Directory.CreateDirectory(ReadingTexts.UserFolder)
                File.WriteAllText(Path.Combine(ReadingTexts.UserFolder, d.TextName & ".txt"), d.TextBody)
                FillTexts()
                lstReading.SelectedItem = d.TextName
            Catch ex As Exception
                Ask.Error(FindForm(), "common.errorWithDetail", ex.Message)
            End Try
        End Using
    End Sub

    Private Function DeleteFile(folder As String, name As String) As Boolean
        If Not Ask.YesNoWarning(FindForm(), "settings.texts.confirmDelete", name) Then Return False
        Try
            File.Delete(Path.Combine(folder, name))
            Return True
        Catch ex As Exception
            Ask.Error(FindForm(), "common.errorWithDetail", ex.Message)
            Return False
        End Try
    End Function


    Private _settingLanguage As Boolean

    Private Function BuildAppearance() As Control
        ' Language of the program (moved here from the top bar): applied as soon as it is chosen
        Dim language = PageCard("settings.section.language")
        Dim colLang = Ui.Column()
        colLang.Dock = DockStyle.Top
        cmbLanguage.Width = 220
        For Each lang In [Enum].GetValues(Of AppLanguage)()
            cmbLanguage.Items.Add(Loc.NativeName(lang))
        Next
        AddHandler cmbLanguage.SelectedIndexChanged, Sub() LanguagePicked()
        colLang.Controls.Add(Ui.Stack("settings.language", cmbLanguage))
        Dim langNote = Ui.Label("settings.languageNote", muted:=True)
        langNote.MaximumSize = New Size(560, 0)
        langNote.Margin = New Padding(0, 0, 0, 14)
        colLang.Controls.Add(langNote)
        language.Controls.Add(colLang)

        Dim look = PageCard("settings.section.appearance")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        segTheme.SetOptions("theme.light", "theme.dark", "theme.blue")
        AddHandler segTheme.SelectedIndexChanged, Sub() ThemeManager.SetMode(CType(segTheme.SelectedIndex, ThemeMode))
        col.Controls.Add(Ui.Label("settings.theme", muted:=True))
        col.Controls.Add(segTheme)
        look.Controls.Add(col)

        ' (Page shows its cards last first: Appearance at the top, then Language)
        Return Page(language, look)
    End Function

    ''' <summary>Changes the program's language straight away (as the top bar's language box did).</summary>
    Private Sub LanguagePicked()
        If _settingLanguage OrElse cmbLanguage.SelectedIndex < 0 Then Return
        Dim lang = CType(cmbLanguage.SelectedIndex, AppLanguage)
        If lang = Loc.Language Then Return
        AppHost.Settings.Language = lang.ToString()
        AppHost.SaveSettings()
        Loc.SetLanguage(lang)
    End Sub

    Private Function BuildData() As Control
        Dim main = PageCard("settings.section.mainDb")
        main.Controls.Add(FileRow(txtDatabase, Sub() BrowseDatabase(txtDatabase), Sub() NewDatabase(txtDatabase, isQuest:=False)))
        Dim importCard = PageCard("settings.section.importOld")
        Dim icol = Ui.Column()
        icol.Dock = DockStyle.Top
        Dim ihint = Ui.Label("settings.importOld.hint", muted:=True)
        ihint.MaximumSize = New Size(620, 0)
        icol.Controls.Add(ihint)
        Dim importOld = ModernButton.Create("menu.importOld", ButtonKind.Secondary, Glyphs.Folder)
        importOld.Margin = New Padding(0, 10, 0, 0)
        AddHandler importOld.Click, Sub() DatabaseCommands.ImportOldDatabase(FindForm())
        icol.Controls.Add(importOld)
        importCard.Controls.Add(icol)
        Dim data = PageCard("settings.section.dataFolder")
        data.Controls.Add(FileRow(txtDataDir, Sub() BrowseFolder(txtDataDir), Nothing))
        Dim importRow = Ui.Row()
        importRow.Dock = DockStyle.Top
        Dim import = ModernButton.Create("settings.importStandard", ButtonKind.Secondary, Glyphs.Copy)
        import.Margin = New Padding(0, 10, 12, 0)
        AddHandler import.Click, Sub() ImportStandardFiles()
        importRow.Controls.Add(import)
        Dim hint = Ui.Label("settings.importStandard.hint", muted:=True)
        hint.MaximumSize = New Size(520, 0)
        hint.Margin = New Padding(0, 14, 0, 0)
        importRow.Controls.Add(hint)
        data.Controls.Add(importRow)
        importRow.BringToFront()

        ' Move the patient database
        Dim moveRow = Ui.Row()
        moveRow.Dock = DockStyle.Top
        Dim move = ModernButton.Create("backup.move", ButtonKind.Secondary, Glyphs.Move)
        move.Margin = New Padding(0, 10, 12, 0)
        AddHandler move.Click, Sub()
                                   If BackupCommands.MoveDatabase(FindForm()) Then
                                       txtDatabase.Text = AppHost.Settings.Local_Database
                                       ShowLastBackup()
                                   End If
                               End Sub
        moveRow.Controls.Add(move)
        Dim moveHint = Ui.Label("backup.move.hint", muted:=True)
        moveHint.MaximumSize = New Size(520, 0)
        moveHint.Margin = New Padding(0, 14, 0, 0)
        moveRow.Controls.Add(moveHint)
        main.Controls.Add(moveRow)
        moveRow.BringToFront()

        Return Page(main, importCard, data, BuildBackup())
    End Function

    ''' <summary>Backup folder, when to back up, Back up now and Restore.</summary>
    Private Function BuildBackup() As Control
        Dim card = PageCard("settings.section.backup")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        col.Controls.Add(FileRow(txtBackupFolder, Sub()
                                                      Dim f = BackupCommands.PickFolder(FindForm())
                                                      If f IsNot Nothing Then
                                                          txtBackupFolder.Text = f
                                                          ShowLastBackup()
                                                      End If
                                                  End Sub, Nothing))
        txtBackupFolder.Box.ReadOnly = True

        segBackupWhen.SetOptions("backup.whenClose", "backup.whenManual")
        segBackupWhen.Margin = New Padding(0, 12, 0, 6)
        AddHandler segBackupWhen.SelectedIndexChanged, Sub()
                                                           If _loadingBackup Then Return
                                                           AppHost.Settings.BackupOnClose = segBackupWhen.SelectedIndex = 0
                                                           AppHost.SaveSettings()
                                                       End Sub
        col.Controls.Add(Ui.Stack("backup.when", segBackupWhen))
        AddHandler chkBackupStimuli.CheckedChanged, Sub()
                                                        If _loadingBackup Then Return
                                                        AppHost.Settings.BackupIncludeStimuli = chkBackupStimuli.Checked
                                                        AppHost.SaveSettings()
                                                    End Sub
        col.Controls.Add(chkBackupStimuli)
        Dim hint As New Label With {.AutoSize = True, .Font = Fonts.Small, .MaximumSize = New Size(620, 0), .Margin = New Padding(0, 6, 0, 10)}
        LabelStyle.SetMuted(hint)
        Loc.BindF(hint, "backup.hint", BackupService.KeepCount)
        col.Controls.Add(hint)

        Dim row = Ui.Row()
        Dim now = ModernButton.Create("backup.now", ButtonKind.Primary, Glyphs.Save)
        now.Margin = New Padding(0, 0, 8, 0)
        AddHandler now.Click, Sub()
                                  BackupCommands.BackUpNow(FindForm())
                                  txtBackupFolder.Text = AppHost.Settings.BackupFolder
                                  ShowLastBackup()
                              End Sub
        Dim restore = ModernButton.Create("backup.restore", ButtonKind.Secondary, Glyphs.Restore)
        restore.Margin = New Padding(0, 0, 16, 0)
        AddHandler restore.Click, Sub()
                                      BackupCommands.Restore(FindForm())
                                      LoadAll()
                                  End Sub
        row.Controls.Add(now)
        row.Controls.Add(restore)
        lblLastBackup.Margin = New Padding(0, 9, 0, 0)
        LabelStyle.SetMuted(lblLastBackup)
        row.Controls.Add(lblLastBackup)
        col.Controls.Add(row)
        card.Controls.Add(col)
        AddHandler Loc.LanguageChanged, Sub() ShowLastBackup()
        Return card
    End Function

    Private Sub LoadBackup()
        _loadingBackup = True
        Dim s = AppHost.Settings
        txtBackupFolder.Text = s.BackupFolder
        segBackupWhen.SelectedIndex = If(s.BackupOnClose, 0, 1)
        chkBackupStimuli.Checked = s.BackupIncludeStimuli
        _loadingBackup = False
        ShowLastBackup()
    End Sub

    Private Sub ShowLastBackup()
        Try
            Dim newest = BackupService.List(AppHost.Settings.BackupFolder).FirstOrDefault()
            lblLastBackup.Text = If(newest Is Nothing, Loc.T("backup.lastNone"), Loc.F("backup.last", newest.Created.ToString("g", Loc.Culture)))
        Catch ex As Exception
            lblLastBackup.Text = ""
        End Try
    End Sub

    ''' <summary>Copies stimuli, sounds and report pictures from another data folder (missing files only).</summary>
    Private Sub ImportStandardFiles()
        Using dlg As New FolderBrowserDialog With {.Description = Loc.T("settings.importStandard.pick"), .UseDescriptionForTitle = True}
            If dlg.ShowDialog(FindForm()) <> DialogResult.OK Then Return
            Dim source = dlg.SelectedPath
            If Not StandardFiles.LooksLikeDataFolder(source) Then
                Ask.Warn(FindForm(), "settings.importStandard.notData", String.Join(", ", StandardFiles.Folders))
                Return
            End If
            Try
                Cursor = Cursors.WaitCursor
                Dim target = txtDataDir.Text.Trim()
                If target = "" Then target = AppHost.Settings.DataDirectory
                Dim n = StandardFiles.CopyMissing(source, target)
                Toast.Show(FindForm(), Loc.F("settings.importStandard.done", n), ToastKind.Success)
            Catch ex As Exception
                Log.Error("Import standard files", ex)
                Ask.Error(FindForm(), "common.errorWithDetail", ex.Message)
            Finally
                Cursor = Cursors.Default
            End Try
        End Using
    End Sub

    Private Shared Function FileRow(field As TextField, browse As Action, create As Action) As Control
        Dim row = Ui.Row()
        row.Dock = DockStyle.Top
        field.Width = 520
        field.Margin = New Padding(0, 0, 8, 0)
        row.Controls.Add(field)
        Dim b = ModernButton.Create("common.browse", ButtonKind.Secondary, Glyphs.Open)
        AddHandler b.Click, Sub() browse()
        row.Controls.Add(b)
        If create IsNot Nothing Then
            Dim n = ModernButton.Create("common.new", ButtonKind.Secondary, Glyphs.Add)
            AddHandler n.Click, Sub() create()
            row.Controls.Add(n)
        End If
        Return row
    End Function

    ''' <summary>
    ''' Screens: one screen (which monitor), two screens (which monitor is the control screen and
    ''' which the patient's), or the control screen and a tablet (the tablet chosen in Tablet
    ''' devices becomes the patient's screen).
    ''' </summary>
    Private Function BuildScreens() As Control
        Dim arrangement = PageCard("settings.section.screens")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        segScreens.SetOptions("settings.screens.modeOne", "settings.screens.modeTwo", "settings.screens.modeTablet")
        segScreens.Margin = New Padding(0, 0, 0, 10)
        AddHandler segScreens.SelectedIndexChanged, Sub() ModeChanged()
        col.Controls.Add(Ui.Stack("settings.screens.use", segScreens))

        ' Which monitor: one screen - the monitor for everything; otherwise the control screen (and the patient's)
        FillMonitors()
        _oneMonitorStack = Ui.Stack("settings.screens.oneMonitor", cboOneMonitor)
        _controlStack = Ui.Stack("settings.screens.controlMonitor", cboControlMon)
        _patientStack = Ui.Stack("settings.screens.patientMonitor", cboPatientMon)
        Dim pickers = Ui.Row(wrap:=True)
        For Each st In {_oneMonitorStack, _controlStack, _patientStack}
            st.Margin = New Padding(0, 0, 16, 6)
            pickers.Controls.Add(st)
        Next
        col.Controls.Add(pickers)
        AddHandler cboOneMonitor.SelectedIndexChanged, Sub()
                                                           preview.Invalidate()
                                                           ShowMode()
                                                       End Sub
        AddHandler cboControlMon.SelectedIndexChanged, Sub()
                                                           If _syncing Then Return
                                                           Dim i = cboControlMon.SelectedIndex
                                                           If i >= 0 AndAlso MonitorIndex(ctl.Value) <> i Then ctl.Value = ScreenLayout.ScreenRect(i, False)
                                                       End Sub
        AddHandler cboPatientMon.SelectedIndexChanged, Sub()
                                                           If _syncing Then Return
                                                           Dim i = cboPatientMon.SelectedIndex
                                                           If i >= 0 AndAlso MonitorIndex(stim.Value) <> i Then stim.Value = ScreenLayout.ScreenRect(i, True)
                                                       End Sub

        ' Control screen and tablet: what the tablet is for, its state, and the devices
        Dim tcol = Ui.Column()
        tcol.Margin = New Padding(0, 4, 0, 8)
        Dim becomes = Ui.Label("settings.tablet.becomes")
        becomes.Font = Fonts.Strong
        becomes.MaximumSize = New Size(620, 0)
        tcol.Controls.Add(becomes)
        Dim about = Ui.Label("settings.tablet.about", muted:=True)
        about.MaximumSize = New Size(620, 0)
        about.Margin = New Padding(0, 6, 0, 0)
        tcol.Controls.Add(about)
        LabelStyle.SetMuted(lblTablet)
        tcol.Controls.Add(lblTablet)
        Dim devices = ModernButton.Create("settings.tablet.devices", ButtonKind.Primary, DevicesScreen.DevicesGlyph)
        AddHandler devices.Click, Sub()
                                      Dim main = AppHost.MainWindow
                                      If main Is Nothing Then Return
                                      _keepEdits = True
                                      If Not main.ShowScreen(MainForm.ScreenDevices) Then _keepEdits = False
                                  End Sub
        tcol.Controls.Add(devices)
        _tabletPanel = tcol
        col.Controls.Add(tcol)
        AddHandler TabletScreen.Changed, Sub() If Not IsDisposed Then ShowTablet()

        preview.Size = New Size(620, 170)
        ' One screen: both windows on the monitor chosen; a tablet: no patient's window on the computer
        preview.ControlRect = Function() If(segScreens.SelectedIndex = ModeOne, ChosenMonitor().WorkingArea, ctl.Value)
        preview.StimulusRect = Function()
                                   Select Case segScreens.SelectedIndex
                                       Case ModeOne : Return ChosenMonitor().Bounds
                                       Case ModeTablet : Return Rectangle.Empty
                                       Case Else : Return stim.Value
                                   End Select
                               End Function
        col.Controls.Add(preview)
        col.Controls.Add(lblMode)
        arrangement.Controls.Add(col)

        ' Exact positions (two screens; the control screen also with a tablet)
        _controlCard = PageCard("settings.section.controlScreen")
        ctl.Dock = DockStyle.Top
        AddHandler ctl.Changed, Sub()
                                    SyncPickers()
                                    preview.Invalidate()
                                    ShowMode()
                                End Sub
        _controlCard.Controls.Add(ctl)

        _patientCard = PageCard("settings.section.patientScreen")
        Dim col2 = Ui.Column()
        col2.Dock = DockStyle.Top
        AddHandler stim.Changed, Sub()
                                     SyncPickers()
                                     preview.Invalidate()
                                     ShowMode()
                                 End Sub
        col2.Controls.Add(stim)
        _patientCard.Controls.Add(col2)
        ModeChanged()
        Return Page(arrangement, _controlCard, _patientCard)
    End Function

    ''' <summary>Shows what the mode chosen needs: the monitor pickers, the tablet, the position cards.</summary>
    Private Sub ModeChanged()
        If _patientCard Is Nothing Then Return
        Dim m = segScreens.SelectedIndex
        _oneMonitorStack.Visible = m = ModeOne
        _controlStack.Visible = m <> ModeOne
        _patientStack.Visible = m = ModeTwo
        _tabletPanel.Visible = m = ModeTablet
        _controlCard.Visible = m <> ModeOne
        _patientCard.Visible = m = ModeTwo
        ' Two screens chosen with both windows on one monitor: put the patient's on another one
        If m = ModeTwo AndAlso Not _loadingScreens Then
            Dim screens = ScreenLayout.PrimaryLeadingScreens()
            Dim ci = Math.Max(0, MonitorIndex(ctl.Value))
            If screens.Count > 1 AndAlso MonitorIndex(stim.Value) = ci Then
                stim.Value = ScreenLayout.ScreenRect(If(ci = 0, 1, 0), True)
            End If
        End If
        SyncPickers()
        ShowTablet()
        preview.Invalidate()
        ShowMode()
    End Sub

    ''' <summary>The control and patient pickers follow the positions (which may also be typed in).</summary>
    Private Sub SyncPickers()
        _syncing = True
        Try
            Dim ci = MonitorIndex(ctl.Value), pi = MonitorIndex(stim.Value)
            If ci < cboControlMon.Items.Count Then cboControlMon.SelectedIndex = ci
            If pi < cboPatientMon.Items.Count Then cboPatientMon.SelectedIndex = pi
        Finally
            _syncing = False
        End Try
    End Sub

    ''' <summary>Which monitor (0 = screen 1, as numbered in the picture) holds r; -1 if none.</summary>
    Private Shared Function MonitorIndex(r As Rectangle) As Integer
        If Not ScreenLayout.IsOnScreen(r) Then Return -1
        Dim dev = ScreenLayout.ScreenOf(r).DeviceName
        Return ScreenLayout.PrimaryLeadingScreens().FindIndex(Function(x) x.DeviceName = dev)
    End Function

    ''' <summary>"Screen 1 (1920 × 1080, main)", "Screen 2 ..." - numbered as in the picture below.</summary>
    Private Sub FillMonitors()
        _syncing = True
        Try
            Dim screens = ScreenLayout.PrimaryLeadingScreens()
            For Each cbo In {cboOneMonitor, cboControlMon, cboPatientMon}
                cbo.Items.Clear()
                For i = 0 To screens.Count - 1
                    Dim b = screens(i).Bounds
                    cbo.Items.Add(Loc.F(If(screens(i).Primary, "settings.screens.monitorMain", "settings.screens.monitor"), i + 1, b.Width, b.Height))
                Next
            Next
            cboOneMonitor.SelectedIndex = Math.Max(0, Math.Min(cboOneMonitor.Items.Count - 1, AppHost.Settings.OneScreenMonitor))
        Finally
            _syncing = False
        End Try
    End Sub

    Private Function ChosenMonitor() As Screen
        Dim screens = ScreenLayout.PrimaryLeadingScreens()
        Dim i = cboOneMonitor.SelectedIndex
        Return If(i >= 0 AndAlso i < screens.Count, screens(i), screens(0))
    End Function

    Private Sub ShowTablet()
        Dim s = AppHost.Settings
        If s.CompanyGuid = Guid.Empty Then
            lblTablet.Text = Loc.T("settings.tablet.noPractice")
        ElseIf TabletScreen.IsActive Then
            Dim sz = TabletScreen.ScreenSize
            lblTablet.Text = Loc.F("settings.tablet.active", TabletScreen.DeviceName, sz.Width, sz.Height)
        ElseIf s.UseTablet Then
            lblTablet.Text = Loc.T("settings.tablet.waiting")
        Else
            lblTablet.Text = Loc.T("settings.tablet.notSaved")
        End If
    End Sub

    ''' <summary>ReadEZ colours: checked on the patient's screen (as saved: save new screen positions first).</summary>
    Private Function BuildReadEZ() As Control
        Dim readez = PageCard("settings.section.readez")
        Dim col3 = Ui.Column()
        col3.Dock = DockStyle.Top
        Dim about = Ui.Label("settings.readez.about", muted:=True)
        about.MaximumSize = New Size(560, 0)
        col3.Controls.Add(about)
        col3.Controls.Add(lblReadEZ)
        Dim check = ModernButton.Create("settings.readez.check", ButtonKind.Secondary, Glyphs.Calibrate)
        AddHandler check.Click, Sub()
                                    ReadEZColourCheck.Run(FindForm(), "")
                                    ShowReadEZStatus()
                                End Sub
        col3.Controls.Add(check)
        readez.Controls.Add(col3)
        Return readez
    End Function

    ''' <summary>What the choices mean, in one line under the picture.</summary>
    Private Sub ShowMode()
        Select Case segScreens.SelectedIndex
            Case ModeOne
                lblMode.Text = Loc.F("settings.mode.oneChosenOn", Math.Max(0, cboOneMonitor.SelectedIndex) + 1)
            Case ModeTablet
                lblMode.Text = Loc.F("settings.mode.tabletOn", Math.Max(0, MonitorIndex(ctl.Value)) + 1)
            Case Else
                Dim ci = MonitorIndex(ctl.Value), pi = MonitorIndex(stim.Value)
                If Screen.AllScreens.Length < 2 Then
                    lblMode.Text = Loc.T("settings.mode.twoOnlyOne")
                ElseIf ci < 0 OrElse pi < 0 OrElse ci = pi Then
                    lblMode.Text = Loc.T("settings.mode.twoSame")
                Else
                    lblMode.Text = Loc.F("settings.mode.twoOn", ci + 1, pi + 1)
                End If
        End Select
    End Sub

    Private Sub ShowReadEZStatus()
        ReadEZColours.ForgetScreen()
        Dim n = ReadEZColours.CheckedCount(), total = 3 * ReadEZColours.PerSet
        lblReadEZ.Text = Loc.F("readez.check.screen", ReadEZColours.ScreenName) & vbCrLf &
                         If(n = total, Loc.T("readez.allChecked"), Loc.F("readez.someChecked", n, total))
    End Sub

    Private Function BuildAdmin() As Control
        Dim admin = PageCard("settings.section.admin")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        Dim info = Ui.Label("settings.admin.info", muted:=True)
        info.MaximumSize = New Size(620, 0)
        col.Controls.Add(info)
        Dim open = ModernButton.Create("settings.admin.open", ButtonKind.Primary, Glyphs.Lock)
        open.Margin = New Padding(0, 10, 0, 0)
        AddHandler open.Click, Sub() OpenAdmin()
        col.Controls.Add(open)
        admin.Controls.Add(col)
        Return Page(admin)
    End Function

#End Region

#Region "Load and save"

    Private Sub LoadAll()
        FillTexts()
        ShowReadEZStatus()
        Dim s = AppHost.Settings
        txtPractice.Text = s.Practice_Name
        chkLogo.Checked = s.ShowLogo
        txtLogo.Text = s.LogoFilename
        nudLogoScale.Value = Clamp(nudLogoScale, s.LogoScale)
        nudScreenW.Value = Clamp(nudScreenW, s.Screen_Width)
        nudScreenH.Value = Clamp(nudScreenH, s.Screen_Height)
        nudReadTime.Value = Clamp(nudReadTime, s.RoR_ReadingTime)
        nudTiming.Value = CDec(Math.Max(0.5, Math.Min(10, s.Timing / 1000.0)))
        chkBeep.Checked = s.RoR_Beep
        FillMonitors()
        ShowLogoPreview()

        segTheme.SelectedIndex = CInt(ThemeManager.Mode)
        _settingLanguage = True
        cmbLanguage.SelectedIndex = CInt(s.UiLanguage)
        _settingLanguage = False

        txtDatabase.Text = s.Local_Database
        txtDataDir.Text = s.DataDirectory
        LoadBackup()

        ctl.Value = s.ControlBounds
        stim.Value = s.StimulusBounds
        _loadingScreens = True
        segScreens.SelectedIndex = If(s.UseTablet, ModeTablet, If(s.ForceOneScreen OrElse Screen.AllScreens.Length < 2, ModeOne, ModeTwo))
        _loadingScreens = False
        ModeChanged()
    End Sub

    Private Shared Function Clamp(n As NumericUpDown, v As Integer) As Decimal
        Return Math.Max(n.Minimum, Math.Min(n.Maximum, v))
    End Function

    Private Sub SaveAll()
        Dim s = AppHost.Settings
        ' Opening a path that doesn't exist silently creates an empty database (a typo or an
        ' unconnected network drive would look like "all patients gone"), so ask before saving anything.
        Dim newDb = txtDatabase.Text.Trim()
        Dim dbChanged = Not String.Equals(newDb, s.Local_Database, StringComparison.OrdinalIgnoreCase)
        If dbChanged AndAlso Not File.Exists(newDb) AndAlso Not Ask.YesNoWarning(FindForm(), "settings.db.notFoundCreate", newDb) Then Return

        ' Database first: if it can't be opened nothing else is changed (the values would otherwise
        ' stay in memory and be written by the next settings save).
        If dbChanged Then
            Try
                Directory.CreateDirectory(Path.GetDirectoryName(newDb))
                If ReadEZDatabase.IsLegacyFile(newDb) Then
                    ' A ReadEZ 2020 database: it is imported into the database in use, not opened
                    Ask.Info(FindForm(), "settings.db.isOld")
                    txtDatabase.Text = s.Local_Database
                    Return
                End If
                AppHost.Db = New ReadEZDatabase(newDb)
                s.Local_Database = newDb
                ' Patient IDs belong to the old database: forget the last patient
                AppHost.Session.SelectPatient(Nothing)
                s.LastPatientId = -1
            Catch ex As Exception
                Log.Error("Open database", ex)
                Ask.Error(FindForm(), "db.openFailed", newDb, ex.Message)
                Return
            End Try
        End If

        s.Practice_Name = txtPractice.Text.Trim()
        s.ShowLogo = chkLogo.Checked
        s.LogoFilename = txtLogo.Text.Trim()
        s.LogoScale = CInt(nudLogoScale.Value)
        s.Screen_Width = CInt(nudScreenW.Value)
        s.Screen_Height = CInt(nudScreenH.Value)
        s.RoR_ReadingTime = CInt(nudReadTime.Value)
        s.Timing = CInt(nudTiming.Value * 1000)
        s.RoR_Beep = chkBeep.Checked
        Dim mode = segScreens.SelectedIndex
        s.ForceOneScreen = mode = ModeOne
        s.OneScreenMonitor = Math.Max(0, cboOneMonitor.SelectedIndex)
        Dim useTablet = mode = ModeTablet
        Dim tabletChanged = useTablet <> s.UseTablet

        s.Theme = CType(segTheme.SelectedIndex, ThemeMode)

        Dim c = ctl.Value
        ' One screen: the control window goes on the monitor chosen
        If s.ForceOneScreen AndAlso ScreenLayout.ScreenOf(c).DeviceName <> ChosenMonitor().DeviceName Then
            c = ChosenMonitor().WorkingArea
            ctl.Value = c
        End If
        Dim controlMoved = c <> s.ControlBounds
        s.Control_Left = c.Left : s.Control_Top = c.Top : s.Control_Width = c.Width : s.Control_Height = c.Height
        Dim t = stim.Value
        s.Stimulus_Left = t.Left : s.Stimulus_Top = t.Top : s.Stimulus_Width = t.Width : s.Stimulus_Height = t.Height

        s.DataDirectory = txtDataDir.Text.Trim()

        ' Language last, so the confirmation appears in the new language
        Dim lang = CType(Math.Max(0, cmbLanguage.SelectedIndex), AppLanguage)
        s.Language = lang.ToString()
        AppHost.SaveSettings()
        If tabletChanged Then TabletScreen.SetEnabled(useTablet)     ' the tablet becomes (or stops being) the patient's screen
        If Stimulus.IsOpen Then
            Stimulus.Window.PlaceOnStimulusScreen()
            Stimulus.Window.ResumeTopMost()      ' after an emergency Ctrl+S
        End If
        Dim main = TryCast(FindForm(), MainForm)
        If controlMoved AndAlso main IsNot Nothing Then main.PlaceOnControlScreen()
        main?.Activate()                         ' the control window in front (the patient's view stays behind it)
        ReadEZColours.ForgetScreen()
        main?.RefreshStatus()
        ThemeManager.SetMode(s.Theme)
        If lang <> Loc.Language Then Loc.SetLanguage(lang)
        Toast.Show(FindForm(), Loc.T("settings.saved"), ToastKind.Success)
    End Sub

    Public Overrides Function CanLeave() As Boolean
        ' Theme preview is applied live; put it back if the user leaves without saving
        If ThemeManager.Mode <> AppHost.Settings.Theme Then ThemeManager.SetMode(AppHost.Settings.Theme)
        Return True
    End Function

#End Region

#Region "Pickers"

    Private Sub ShowLogoPreview()
        Dim old = pctLogo.Image
        pctLogo.Image = Nothing
        old?.Dispose()
        Try
            If File.Exists(txtLogo.Text) Then
                ' Copy into a new bitmap so the file is not kept open (and GDI+ doesn't need the stream)
                Using fs = File.OpenRead(txtLogo.Text), img = Image.FromStream(fs)
                    pctLogo.Image = New Bitmap(img)
                End Using
            End If
        Catch ex As Exception
            Log.Warn("Logo preview: " & ex.Message)
        End Try
    End Sub

    Private Sub BrowseLogo()
        Using dlg As New OpenFileDialog With {.Filter = "Images (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"}
            If File.Exists(txtLogo.Text) Then dlg.InitialDirectory = Path.GetDirectoryName(txtLogo.Text)
            If dlg.ShowDialog(FindForm()) = DialogResult.OK Then
                txtLogo.Text = dlg.FileName
                ShowLogoPreview()
            End If
        End Using
    End Sub

    Private Sub BrowseDatabase(field As TextField)
        Using dlg As New OpenFileDialog With {.Filter = "SQLite database (*.db)|*.db|All files (*.*)|*.*", .CheckFileExists = True}
            If File.Exists(field.Text) Then dlg.InitialDirectory = Path.GetDirectoryName(field.Text)
            If dlg.ShowDialog(FindForm()) = DialogResult.OK Then field.Text = dlg.FileName
        End Using
    End Sub

    Private Sub NewDatabase(field As TextField, isQuest As Boolean)
        Using dlg As New SaveFileDialog With {.Filter = "SQLite database (*.db)|*.db", .DefaultExt = "db", .AddExtension = True,
                                             .OverwritePrompt = False, .FileName = "ReadEZ.db"}
            If dlg.ShowDialog(FindForm()) <> DialogResult.OK Then Return
            If File.Exists(dlg.FileName) Then
                ' Never overwrite an existing database
                If Not Ask.YesNo(FindForm(), "settings.db.existsUse") Then Return
            Else
                SQLiteConnection.CreateFile(dlg.FileName)
                Dim created As New ReadEZDatabase(dlg.FileName)
                Toast.Show(FindForm(), Loc.F("settings.db.created", dlg.FileName), ToastKind.Success)
            End If
            field.Text = dlg.FileName
        End Using
    End Sub

    Private Sub BrowseFolder(field As TextField)
        Using dlg As New FolderBrowserDialog With {.SelectedPath = field.Text}
            If dlg.ShowDialog(FindForm()) = DialogResult.OK Then field.Text = dlg.SelectedPath
        End Using
    End Sub

    Private Sub OpenAdmin()
        Dim pw = Ask.Text(FindForm(), "admin.password.title", "admin.password.prompt", password:=True)
        If pw Is Nothing Then Return
        If Not AppHost.Admin.CheckPassword(pw) Then
            Ask.Warn(FindForm(), "admin.password.wrong")
            Return
        End If
        Using f As New AdminDialog()
            f.ShowDialog(FindForm())
            ' Otherwise a later Save would write this screen's old values over the imported ones
            If f.SettingsImported Then LoadAll()
        End Using
    End Sub

#End Region

End Class

''' <summary>Left/Top/Width/Height boxes with one "use screen N" button per attached screen.</summary>
Public Class RectEditor
    Inherits FlowLayoutPanel

    Private ReadOnly nudL As NumericUpDown = Ui.NumberBox(-20000, 20000, 0)
    Private ReadOnly nudT As NumericUpDown = Ui.NumberBox(-20000, 20000, 0)
    Private ReadOnly nudW As NumericUpDown = Ui.NumberBox(0, 20000, 0)
    Private ReadOnly nudH As NumericUpDown = Ui.NumberBox(0, 20000, 0)
    Private _loading As Boolean

    Public Event Changed As EventHandler

    Private ReadOnly _fullScreen As Boolean

    ''' <param name="fullScreen">True for the patient screen: "Use screen N" takes the whole screen, taskbar included.</param>
    Public Sub New(fullScreen As Boolean)
        _fullScreen = fullScreen
        AutoSize = True
        AutoSizeMode = AutoSizeMode.GrowAndShrink
        WrapContents = True
        BackColor = Color.Transparent
        Margin = New Padding(0)
        Controls.Add(Ui.Stack("settings.rect.left", nudL))
        Controls.Add(Ui.Stack("settings.rect.top", nudT))
        Controls.Add(Ui.Stack("settings.rect.width", nudW))
        Controls.Add(Ui.Stack("settings.rect.height", nudH))
        For Each n In {nudL, nudT, nudW, nudH}
            AddHandler n.ValueChanged, Sub() If Not _loading Then RaiseEvent Changed(Me, EventArgs.Empty)
        Next
        Dim buttons = Ui.Row()
        buttons.Margin = New Padding(8, 22, 0, 0)
        Dim screens = ScreenLayout.PrimaryLeadingScreens()
        For i = 0 To screens.Count - 1
            Dim index = i
            Dim b = ModernButton.Create("", ButtonKind.Secondary, Glyphs.Tracker)
            Loc.OnChange(b, Sub() b.Text = Loc.F("settings.rect.useScreen", index + 1))
            AddHandler b.Click, Sub() Value = ScreenLayout.ScreenRect(index, _fullScreen)
            buttons.Controls.Add(b)
        Next
        Controls.Add(buttons)
    End Sub

    Public Property Value As Rectangle
        Get
            Return New Rectangle(CInt(nudL.Value), CInt(nudT.Value), CInt(nudW.Value), CInt(nudH.Value))
        End Get
        Set(r As Rectangle)
            _loading = True
            nudL.Value = Math.Max(nudL.Minimum, Math.Min(nudL.Maximum, r.Left))
            nudT.Value = Math.Max(nudT.Minimum, Math.Min(nudT.Maximum, r.Top))
            nudW.Value = Math.Max(nudW.Minimum, Math.Min(nudW.Maximum, r.Width))
            nudH.Value = Math.Max(nudH.Minimum, Math.Min(nudH.Maximum, r.Height))
            _loading = False
            RaiseEvent Changed(Me, EventArgs.Empty)
        End Set
    End Property
End Class

''' <summary>Miniature of all attached screens showing where the control and patient windows go.</summary>
Public Class ScreenArrangementPreview
    Inherits Control
    Implements IThemed

    Public Property ControlRect As Func(Of Rectangle)
    Public Property StimulusRect As Func(Of Rectangle)

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Margin = New Padding(0, 0, 0, 10)
        Loc.OnChange(Me, Sub() Invalidate())
    End Sub

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim all = ScreenLayout.AllScreensRect()
        If all.Width = 0 OrElse all.Height = 0 Then Return
        Dim scale = Math.Min((Width - 10) / all.Width, (Height - 26) / all.Height)
        Dim ox = (Width - all.Width * scale) / 2
        Dim map = Function(r As Rectangle) New RectangleF(CSng(ox + (r.X - all.X) * scale), CSng((r.Y - all.Y) * scale + 2), CSng(r.Width * scale), CSng(r.Height * scale))
        Dim screens = ScreenLayout.PrimaryLeadingScreens()
        For i = 0 To screens.Count - 1
            Dim r = map(screens(i).Bounds)
            Using b As New SolidBrush(p.SurfaceAlt), pen As New Pen(p.Border, 2)
                g.FillRectangle(b, r)
                g.DrawRectangle(pen, r.X, r.Y, r.Width, r.Height)
            End Using
            TextRenderer.DrawText(g, Loc.F("settings.screenN", i + 1), Fonts.SmallStrong, Rectangle.Round(r), p.Muted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.Bottom)
        Next
        If ControlRect IsNot Nothing Then DrawWindow(g, map(ControlRect.Invoke()), p.Primary, Loc.T("settings.controlWindow"))
        If StimulusRect IsNot Nothing Then
            Dim sr = StimulusRect.Invoke()
            If Not sr.IsEmpty Then DrawWindow(g, map(sr), p.Success, Loc.T("settings.patientWindow"))
        End If
    End Sub

    Private Shared Sub DrawWindow(g As Graphics, r As RectangleF, colour As Color, caption As String)
        Using pen As New Pen(colour, 3), b As New SolidBrush(Color.FromArgb(40, colour))
            g.FillRectangle(b, r)
            g.DrawRectangle(pen, r.X + 1.5F, r.Y + 1.5F, r.Width - 3, r.Height - 3)
        End Using
        TextRenderer.DrawText(g, caption, Fonts.SmallStrong, Rectangle.Round(r), colour, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
    End Sub
End Class
