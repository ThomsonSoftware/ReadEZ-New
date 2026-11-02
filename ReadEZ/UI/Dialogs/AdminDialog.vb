Option Strict On

Imports System.IO

' ====================================================================================================
'   AdminDialog.vb - administrator settings (as Eye Dynamix)
'
'   Sections
'     Colour tests   - which colour tests are offered: overlays, lenses, computer
'     Releases       - build the signed installer and update.json (PublishUpdateDialog)
'     Help topics    - add, delete and reorder help topics and set their videos (HelpTopicsDialog)
'     Online help    - the web address of the online help (Help > Online help)
'     Maintenance    - change the password, import ReadEZ 2020's settings (RZSettings.ini) and
'                      database (ReadEZ2020.db), translation check
' ====================================================================================================

Public Class AdminDialog
    Inherits DialogBase

    ' Online help
    Private ReadOnly txtHelpUrl As New TextField

    ' Colour tests
    Private ReadOnly chkOverlay As ThemedCheckBox = ThemedCheckBox.Create("mode.overlay")
    Private ReadOnly chkLenses As ThemedCheckBox = ThemedCheckBox.Create("mode.lenses")
    Private ReadOnly chkComputer As ThemedCheckBox = ThemedCheckBox.Create("mode.computer")

    Public Sub New()
        SetTitle("admin.title", "admin.subtitle")
        ClientSize = New Size(800, 760)

        Dim stack As New FlowLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .FlowDirection = FlowDirection.TopDown,
                                               .WrapContents = False, .BackColor = Color.Transparent}
        stack.Controls.Add(BuildModules())
        stack.Controls.Add(BuildReleases())
        stack.Controls.Add(BuildHelpVideos())
        stack.Controls.Add(BuildOnlineHelp())
        stack.Controls.Add(BuildMaintenance())
        Body.Controls.Add(stack)

        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AddHandler save.Click, Sub() SaveAndClose()

        LoadValues()
    End Sub

#Region "Layout"

    Private Const SectionWidth As Integer = 720
    Private Const ContentWidth As Integer = SectionWidth - 36

    ''' <summary>
    ''' A titled card of fixed width whose height follows its content. The content is placed
    ''' below the title (a Panel ignores Padding for undocked children, so before this the content
    ''' sat at the top-left corner, over the title and the card's border).
    ''' </summary>
    Private Shared Function Section(titleKey As String, content As Control) As Card
        Dim c As New Card With {.AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                                .MinimumSize = New Size(SectionWidth, 0), .MaximumSize = New Size(SectionWidth, 0),
                                .Padding = New Padding(18, 44, 18, 16), .Margin = New Padding(0, 0, 0, 14)}
        c.SetTitle(titleKey)
        content.Margin = New Padding(0)
        content.Location = New Point(c.Padding.Left, c.Padding.Top)
        c.Controls.Add(content)
        Return c
    End Function

    Private Function BuildModules() As Control
        Dim row = Ui.Row()
        For Each c In {chkOverlay, chkLenses, chkComputer}
            c.Margin = New Padding(0, 0, 24, 0)
            row.Controls.Add(c)
        Next
        Return Section("admin.modules.title", row)
    End Function

    Private Function BuildMaintenance() As Control
        Dim row = Ui.Row(wrap:=True)
        row.MaximumSize = New Size(ContentWidth, 0)
        Dim pw = ModernButton.Create("admin.password.change", ButtonKind.Secondary, Glyphs.Lock)
        AddHandler pw.Click, Sub() ChangePassword()
        Dim import = ModernButton.Create("admin.import", ButtonKind.Secondary, Glyphs.Folder)
        AddHandler import.Click, Sub() ImportPrevious()
        Dim importDb = ModernButton.Create("admin.importDb", ButtonKind.Secondary, Glyphs.Folder)
        AddHandler importDb.Click, Sub() DatabaseCommands.ImportOldDatabase(Me)
        Dim check = ModernButton.Create("admin.translations", ButtonKind.Secondary, Glyphs.Globe)
        AddHandler check.Click, Sub() CheckTranslations()
        row.Controls.Add(pw)
        row.Controls.Add(import)
        row.Controls.Add(importDb)
        row.Controls.Add(check)
        Return Section("admin.maintenance.title", row)
    End Function

    ''' <summary>Help videos: the link for each topic in the Help menu, Help screen and "?" buttons.</summary>
    Private Function BuildHelpVideos() As Control
        Dim col = Ui.Column()
        Dim intro = Ui.Label("admin.help.intro", muted:=True)
        intro.MaximumSize = New Size(ContentWidth, 0)
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)
        Dim edit = ModernButton.Create("admin.help.edit", ButtonKind.Secondary, Glyphs.Video)
        AddHandler edit.Click, Sub()
                                   Using dlg As New HelpTopicsDialog()
                                       dlg.ShowDialog(Me)
                                   End Using
                               End Sub
        col.Controls.Add(edit)
        Return Section("admin.help.title", col)
    End Function

    ''' <summary>Online help: the address of the help built with HelpNDoc (Help > Online help opens it).</summary>
    Private Function BuildOnlineHelp() As Control
        Dim col = Ui.Column()
        Dim intro = Ui.Label("admin.onlineHelp.intro", muted:=True)
        intro.MaximumSize = New Size(ContentWidth, 0)
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)
        Dim row = Ui.Row()
        row.Controls.Add(Ui.Stack("admin.onlineHelp.url", txtHelpUrl, 480))
        Dim test = ModernButton.Create("admin.onlineHelp.open", ButtonKind.Secondary, Glyphs.Globe)
        test.Margin = New Padding(8, 22, 0, 0)
        AddHandler test.Click, Sub()
                                   Dim url = MainMenu.NormaliseUrl(txtHelpUrl.Text)
                                   If url <> "" Then AppHost.OpenUrl(url)
                               End Sub
        row.Controls.Add(test)
        col.Controls.Add(row)
        Return Section("admin.onlineHelp.title", col)
    End Function

    ''' <summary>Releases: build the signed installer and update.json (PublishUpdateDialog).</summary>
    Private Function BuildReleases() As Control
        Dim col = Ui.Column()
        Dim intro = Ui.Label("admin.releases.intro", muted:=True)
        intro.MaximumSize = New Size(ContentWidth, 0)
        intro.Margin = New Padding(0, 0, 0, 10)
        col.Controls.Add(intro)
        Dim publish = ModernButton.Create("publish.open", ButtonKind.Primary, Glyphs.Export)
        AddHandler publish.Click, Sub()
                                      Using dlg As New PublishUpdateDialog()
                                          dlg.ShowDialog(Me)
                                      End Using
                                  End Sub
        col.Controls.Add(publish)
        Return Section("admin.releases.title", col)
    End Function

#End Region

#Region "Values"

    Private Sub LoadValues()
        Dim a = AppHost.Admin
        txtHelpUrl.Text = If(a.OnlineHelpUrl, "")
        chkOverlay.Checked = a.OverlayTest_Enabled
        chkLenses.Checked = a.LensTest_Enabled
        chkComputer.Checked = a.ComputerTest_Enabled
    End Sub

    Private Sub StoreValues()
        Dim a = AppHost.Admin
        a.OnlineHelpUrl = MainMenu.NormaliseUrl(txtHelpUrl.Text)
        a.OverlayTest_Enabled = chkOverlay.Checked
        a.LensTest_Enabled = chkLenses.Checked
        a.ComputerTest_Enabled = chkComputer.Checked
    End Sub

    Private Sub SaveAndClose()
        StoreValues()
        AppHost.SaveAdmin()
        Log.Info("Admin settings saved")
        DialogResult = DialogResult.OK
    End Sub

#End Region

#Region "Maintenance"

    Private Sub ChangePassword()
        Dim first = Ask.Text(Me, "admin.password.change", "admin.password.new", password:=True)
        If String.IsNullOrEmpty(first) Then Return
        If first.Length < 6 Then
            Ask.Warn(Me, "admin.password.tooShort")
            Return
        End If
        Dim second = Ask.Text(Me, "admin.password.change", "admin.password.repeat", password:=True)
        If first <> second Then
            Ask.Warn(Me, "admin.password.mismatch")
            Return
        End If
        AppHost.Admin.SetPassword(first)
        AppHost.SaveAdmin()
        Toast.Show(Me, Loc.T("admin.password.changed"), ToastKind.Success)
    End Sub

    ''' <summary>Settings of ReadEZ 2020 (RZSettings.ini): practice, screens, text settings and screen colours.</summary>
    Private Sub ImportPrevious()
        Using dlg As New OpenFileDialog With {.Filter = "RZSettings.ini|RZSettings.ini|*.ini|*.ini", .Title = Loc.T("admin.import.pick")}
            Dim first = LegacyImport.CandidateIniFiles().FirstOrDefault(Function(f) File.Exists(f))
            If first IsNot Nothing Then dlg.InitialDirectory = Path.GetDirectoryName(first)
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim s As AppSettings
            Try
                s = LegacyImport.FromIni(dlg.FileName)
            Catch ex As Exception
                Log.Error("Import settings", ex)
                Ask.Warn(Me, "admin.import.none")
                Return
            End Try
            Dim old = AppHost.Settings
            ' Keep this version's interface choices and database (the old database is imported separately)
            s.Theme = old.Theme
            s.NavCollapsed = old.NavCollapsed
            s.Language = old.Language
            s.Local_Database = old.Local_Database
            s.LastPatientId = old.LastPatientId
            s.ScreenerID = old.ScreenerID
            s.FillMissing()
            AppHost.Settings = s
            AppHost.SaveSettings()
            If Stimulus.IsOpen Then Stimulus.Window.PlaceOnStimulusScreen()
            TryCast(Owner, MainForm)?.PlaceOnControlScreen()
            SettingsImported = True
            Ask.Info(Me, "admin.import.done")
        End Using
    End Sub

    ''' <summary>True after settings were imported (the Settings screen must reload its fields).</summary>
    Public Property SettingsImported As Boolean

    Private Sub CheckTranslations()
        Dim lines As New List(Of String)
        For Each lang In {AppLanguage.Spanish, AppLanguage.German, AppLanguage.Italian}
            Dim missing = Loc.MissingFor(lang)
            lines.Add($"{Loc.NativeName(lang)}: {If(missing.Count = 0, "✓", missing.Count.ToString() & " – " & String.Join(", ", missing.Take(8)) & If(missing.Count > 8, " …", ""))}")
        Next
        MessageBox.Show(Me, String.Join(Environment.NewLine, lines), Loc.T("admin.translations"), MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

#End Region

End Class
