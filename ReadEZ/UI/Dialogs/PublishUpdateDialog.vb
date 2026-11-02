Option Strict On

Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions

' ====================================================================================================
'   PublishUpdateDialog.vb - Admin > Publish update (used only on the computer that builds releases)
'
'   One button does what used to be several manual steps:
'     1. writes the new version into <Version> in ReadEZ\ReadEZ.vbproj ("Increase" adds 1
'        to the last number)
'     2. publishes ReadEZ (self-contained, 64-bit) into ReadEZSetup\obj\publish\Release
'     3. signs ReadEZ.exe and
'        accepts a process of that name signed with the Thomson certificate - and signs it and
'        ReadEZ.dll with signtool and the .pfx
'     4. builds ReadEZSetup.wixproj from that folder (AppPrePublished=true) and signs the MSI
'     5. copies the MSI to the output folder as ReadEZ-Setup-<version>.msi, writes update.json
'        next to it (version, file, size, SHA-256, notes) and opens the folder, ready to upload
'   Starter database: a database with example patients and results can be chosen. It is
'   copied (with SQLite's backup, so a database that is open is copied whole) into the published
'   program as Starter\Database\ReadEZ.db. ReadEZ installs it only the first time it runs on
'   a computer, and never over an existing database (StandardFiles), so updates leave the
'   user's patients alone.
'   dotnet and signtool are started directly rather than from MSBuild Exec tasks: those run
'   temporary .cmd files, which AVG's Behavior Shield blocked as "IDP.Generic".
'   "Use an existing installer" does step 5 for an MSI built some other way.
'
'   The certificate password is kept encrypted for this Windows user (DPAPI), and given to signtool
'   directly, so it never appears in the settings file or the build log.
' ====================================================================================================

Public Class PublishUpdateDialog
    Inherits DialogBase

    Private ReadOnly txtUrl As New TextField
    Private ReadOnly lblServer As New Label
    Private ReadOnly txtSolution As New TextField
    Private ReadOnly txtVersion As New TextField
    Private ReadOnly txtNotes As New TextField(multiline:=True)
    Private ReadOnly txtPfx As New TextField
    Private ReadOnly txtPassword As New TextField
    Private ReadOnly txtSignTool As New TextField
    Private ReadOnly txtTimestamp As New TextField
    Private ReadOnly txtOutput As New TextField
    Private ReadOnly txtStarter As New TextField
    Private ReadOnly txtLog As New TextBox With {.Multiline = True, .ReadOnly = True, .ScrollBars = ScrollBars.Both, .WordWrap = False,
                                                .Font = New Font("Consolas", 8.5F), .BorderStyle = BorderStyle.None}
    Private ReadOnly btnBuild As ModernButton
    Private ReadOnly btnExisting As ModernButton
    Private ReadOnly btnClose As ModernButton
    Private _busy As Boolean

    Public Sub New()
        SetTitle("publish.title", "publish.subtitle")
        ClientSize = New Size(960, 870)
        ' Everything fits the window: two cards side by side, the build output below them
        Body.AutoScroll = False

        ' ----- Release (left) -----
        Dim release = FieldColumn()
        Dim checkUrl = ModernButton.Create("publish.readServer", ButtonKind.Secondary, Glyphs.Refresh)
        AddHandler checkUrl.Click, Async Sub() Await ReadServerAsync()
        release.Controls.Add(FieldRow("publish.url", txtUrl, checkUrl))
        _tip.SetToolTip(txtUrl.Box, Loc.T("publish.urlNote"))
        lblServer.AutoSize = True
        lblServer.Margin = New Padding(0, 0, 0, 10)
        lblServer.Font = Fonts.SmallStrong
        LabelStyle.SetMuted(lblServer)
        release.Controls.Add(lblServer)
        release.Controls.Add(FieldRow("publish.solution", txtSolution, BrowseButton(txtSolution, folder:=True, filter:="")))
        Dim inc = ModernButton.Create("publish.increase", ButtonKind.Secondary, Glyphs.Add)
        AddHandler inc.Click, Sub() IncreaseVersion()
        _tip.SetToolTip(txtVersion.Box, Loc.T("publish.versionNote"))
        _tip.SetToolTip(inc, Loc.T("publish.versionNote"))
        release.Controls.Add(Ui.Label("publish.version", muted:=True))
        Dim ver = Ui.Row()
        txtVersion.Width = 140
        txtVersion.Margin = New Padding(0, 0, 8, 8)
        inc.Margin = New Padding(0)
        ver.Controls.Add(txtVersion)
        ver.Controls.Add(inc)
        release.Controls.Add(ver)
        Dim notesLabel = Ui.Label("publish.notes", muted:=True)
        release.Controls.Add(notesLabel)
        txtNotes.Dock = DockStyle.Fill
        txtNotes.Margin = New Padding(0)
        release.Controls.Add(txtNotes)
        release.RowStyles(release.RowCount - 1).SizeType = SizeType.Percent
        release.RowStyles(release.RowCount - 1).Height = 100

        ' ----- Signing and output (right) -----
        Dim signing = FieldColumn()
        signing.Controls.Add(FieldRow("publish.pfx", txtPfx, BrowseButton(txtPfx, folder:=False, filter:="Certificate (*.pfx;*.p12)|*.pfx;*.p12|All files (*.*)|*.*")))
        txtPassword.Box.UseSystemPasswordChar = True
        signing.Controls.Add(FieldRow("publish.password", txtPassword, Nothing))
        signing.Controls.Add(FieldRow("publish.timestamp", txtTimestamp, Nothing))
        signing.Controls.Add(FieldRow("publish.signtool", txtSignTool, BrowseButton(txtSignTool, folder:=False, filter:="signtool.exe|signtool.exe|Programs (*.exe)|*.exe")))
        signing.Controls.Add(FieldRow("publish.output", txtOutput, BrowseButton(txtOutput, folder:=True, filter:="")))
        ' Starter database (new installations only; blank = an empty database)
        signing.Controls.Add(FieldRow("publish.starterDb", txtStarter, BrowseButton(txtStarter, folder:=False, filter:="ReadEZ database (*.db)|*.db|All files (*.*)|*.*")))
        _tip.SetToolTip(txtStarter.Box, Loc.T("publish.starterDbNote"))

        ' ----- Build output (bottom) -----
        txtLog.Dock = DockStyle.Fill

        Dim grid As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 2, .BackColor = Color.Transparent}
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
        grid.RowStyles.Add(New RowStyle(SizeType.Absolute, 500))
        grid.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        grid.Controls.Add(SectionCard("publish.section.release", release, New Padding(0, 0, 6, 12)), 0, 0)
        grid.Controls.Add(SectionCard("publish.section.signing", signing, New Padding(6, 0, 0, 12)), 1, 0)
        Dim logCard = SectionCard("publish.log", txtLog, New Padding(0))
        grid.Controls.Add(logCard, 0, 1)
        grid.SetColumnSpan(logCard, 2)
        Body.Controls.Add(grid)

        btnBuild = AddFooterButton("publish.build", ButtonKind.Primary, Glyphs.Export)
        btnClose = AddFooterButton("common.close", ButtonKind.Secondary)
        btnExisting = AddFooterButton("publish.existing", ButtonKind.Ghost, Glyphs.Open)
        AddHandler btnBuild.Click, Async Sub() Await BuildAndPublishAsync()
        AddHandler btnExisting.Click, Sub() PublishExisting()
        AddHandler btnClose.Click, Sub() If Not _busy Then DialogResult = DialogResult.Cancel

        LoadValues()
    End Sub

    Private ReadOnly _tip As New ToolTip With {.AutoPopDelay = 15000}

    ''' <summary>A titled card filling its cell; the content is docked so it sits inside the padding.</summary>
    Private Shared Function SectionCard(titleKey As String, content As Control, margin As Padding) As Card
        Dim c As New Card With {.Dock = DockStyle.Fill, .Margin = margin, .Padding = New Padding(16, 0, 16, 14)}
        c.SetTitle(titleKey)
        content.Dock = DockStyle.Fill
        c.Controls.Add(content)
        Return c
    End Function

    ''' <summary>One column of fields, each as wide as the column.</summary>
    Private Shared Function FieldColumn() As TableLayoutPanel
        Dim t As New TableLayoutPanel With {.ColumnCount = 1, .BackColor = Color.Transparent, .GrowStyle = TableLayoutPanelGrowStyle.AddRows}
        t.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        AddHandler t.ControlAdded, Sub(s, e)
                                       t.RowCount = t.Controls.Count
                                       While t.RowStyles.Count < t.RowCount
                                           t.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                                       End While
                                       If e.Control.Dock = DockStyle.None Then e.Control.Dock = DockStyle.Top
                                   End Sub
        Return t
    End Function

    ''' <summary>Label above a field that fills the width, with an optional button on its right.</summary>
    Private Shared Function FieldRow(labelKey As String, field As TextField, button As Control) As TableLayoutPanel
        Dim t As New TableLayoutPanel With {.ColumnCount = 2, .RowCount = 2, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                                            .BackColor = Color.Transparent, .Margin = New Padding(0, 0, 0, 4), .Dock = DockStyle.Top}
        t.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        t.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        t.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        t.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Dim l = Ui.Label(labelKey, muted:=True)
        t.Controls.Add(l, 0, 0)
        t.SetColumnSpan(l, 2)
        field.Dock = DockStyle.Fill
        field.Margin = New Padding(0, 0, 0, 8)
        t.Controls.Add(field, 0, 1)
        If button IsNot Nothing Then
            button.Margin = New Padding(8, 0, 0, 8)
            button.Anchor = AnchorStyles.Top Or AnchorStyles.Left
            t.Controls.Add(button, 1, 1)
        End If
        Return t
    End Function

    Private Function BrowseButton(field As TextField, folder As Boolean, filter As String) As Control
        Dim browse = ModernButton.Create("common.browse", ButtonKind.Secondary, Glyphs.Open)
        AddHandler browse.Click, Sub()
                                     If folder Then
                                         Using dlg As New FolderBrowserDialog With {.SelectedPath = field.Text}
                                             If dlg.ShowDialog(Me) = DialogResult.OK Then field.Text = dlg.SelectedPath
                                         End Using
                                     Else
                                         Using dlg As New OpenFileDialog With {.Filter = filter, .FileName = field.Text}
                                             If dlg.ShowDialog(Me) = DialogResult.OK Then field.Text = dlg.FileName
                                         End Using
                                     End If
                                 End Sub
        Return browse
    End Function

    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Await ReadServerAsync()
    End Sub

#Region "Values"

    Private Sub LoadValues()
        Dim a = AppHost.Admin
        txtUrl.Text = UpdateService.UpdateUrl
        txtSolution.Text = If(a.Publish_SolutionFolder <> "" AndAlso Directory.Exists(a.Publish_SolutionFolder), a.Publish_SolutionFolder, FindSolutionFolder())
        txtOutput.Text = If(a.Publish_OutputFolder <> "", a.Publish_OutputFolder,
                            If(txtSolution.Text <> "", Path.Combine(txtSolution.Text, "Releases"), ""))
        txtPfx.Text = a.Publish_PfxFile
        txtPassword.Text = Dpapi.Unprotect(a.Publish_PfxPassword)
        txtSignTool.Text = If(a.Publish_SignTool <> "" AndAlso File.Exists(a.Publish_SignTool), a.Publish_SignTool, FindSignTool())
        txtTimestamp.Text = If(a.Publish_TimestampUrl = "", "http://timestamp.digicert.com", a.Publish_TimestampUrl)
        txtStarter.Text = a.Publish_StarterDatabase
        txtVersion.Text = If(ReadProjectVersion(), AppHost.VersionNumber)
    End Sub

    Private Sub StoreValues()
        Dim a = AppHost.Admin
        Dim url = txtUrl.Text.Trim()
        a.UpdateUrl = If(String.Equals(url, UpdateService.DefaultUpdateUrl, StringComparison.OrdinalIgnoreCase), "", url)
        a.Publish_SolutionFolder = txtSolution.Text.Trim()
        a.Publish_OutputFolder = txtOutput.Text.Trim()
        a.Publish_PfxFile = txtPfx.Text.Trim()
        a.Publish_PfxPassword = Dpapi.Protect(txtPassword.Text)
        a.Publish_SignTool = txtSignTool.Text.Trim()
        a.Publish_TimestampUrl = txtTimestamp.Text.Trim()
        a.Publish_StarterDatabase = txtStarter.Text.Trim()
        AppHost.SaveAdmin()
    End Sub

    ''' <summary>The folder with ReadEZ.sln above the running program (a Visual Studio build).</summary>
    Private Shared Function FindSolutionFolder() As String
        Dim d = New DirectoryInfo(AppHost.AppFolder)
        While d IsNot Nothing
            If File.Exists(Path.Combine(d.FullName, "ReadEZ.sln")) OrElse File.Exists(Path.Combine(d.FullName, "ReadEZ.slnx")) Then Return d.FullName
            d = d.Parent
        End While
        Return ""
    End Function

    ''' <summary>The newest signtool.exe from the Windows SDK.</summary>
    Private Shared Function FindSignTool() As String
        Dim kits = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Windows Kits", "10", "bin")
        Try
            If Directory.Exists(kits) Then
                Dim best = Directory.GetDirectories(kits).
                    Select(Function(d) (Dir:=d, Ver:=UpdateService.ParseVersion(Path.GetFileName(d)))).
                    Where(Function(x) x.Ver IsNot Nothing AndAlso File.Exists(Path.Combine(x.Dir, "x64", "signtool.exe"))).
                    OrderByDescending(Function(x) x.Ver).FirstOrDefault()
                If best.Dir IsNot Nothing Then Return Path.Combine(best.Dir, "x64", "signtool.exe")
                Dim flat = Path.Combine(kits, "x64", "signtool.exe")
                If File.Exists(flat) Then Return flat
            End If
            Dim ack = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Windows Kits", "10", "App Certification Kit", "signtool.exe")
            If File.Exists(ack) Then Return ack
        Catch ex As Exception
            Log.Warn("FindSignTool: " & ex.Message)
        End Try
        Return ""
    End Function

#End Region

#Region "Version"

    Private ReadOnly Property ProjectFile As String
        Get
            Return Path.Combine(txtSolution.Text.Trim(), "ReadEZ", "ReadEZ.vbproj")
        End Get
    End Property

    Private ReadOnly Property SetupProject As String
        Get
            Return Path.Combine(txtSolution.Text.Trim(), "ReadEZSetup", "ReadEZSetup.wixproj")
        End Get
    End Property

    Private Shared ReadOnly VersionTag As New Regex("<Version>\s*([^<]*?)\s*</Version>", RegexOptions.IgnoreCase)

    Private Function ReadProjectVersion() As String
        Try
            If Not File.Exists(ProjectFile) Then Return Nothing
            Dim m = VersionTag.Match(File.ReadAllText(ProjectFile))
            Return If(m.Success, m.Groups(1).Value, Nothing)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>Writes the version into ReadEZ.vbproj (keeping the rest of the file as it is).</summary>
    Private Sub WriteProjectVersion(version As String)
        Dim bytes = File.ReadAllBytes(ProjectFile)
        Dim bom = bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF
        Dim text = File.ReadAllText(ProjectFile)
        If Not VersionTag.IsMatch(text) Then Throw New InvalidDataException("No <Version> in " & ProjectFile)
        text = VersionTag.Replace(text, "<Version>" & version & "</Version>", 1)
        File.WriteAllText(ProjectFile, text, New UTF8Encoding(bom))
    End Sub

    Private Sub IncreaseVersion()
        Dim v = UpdateService.ParseVersion(txtVersion.Text)
        If v Is Nothing Then v = UpdateService.CurrentVersion
        txtVersion.Text = UpdateService.VersionText(If(v.Revision > 0,
                                                       New Version(v.Major, v.Minor, v.Build, v.Revision + 1),
                                                       New Version(v.Major, v.Minor, v.Build + 1)))
    End Sub

    Private _serverVersion As Version

    Private Async Function ReadServerAsync() As Task
        lblServer.Text = Loc.T("update.checking")
        Try
            Dim info = Await UpdateService.ReadManifestAsync(txtUrl.Text.Trim())
            _serverVersion = info.Version
            lblServer.Text = Loc.F("publish.serverHas", info.VersionText, info.Released)
        Catch ex As Exception
            _serverVersion = Nothing
            lblServer.Text = Loc.F("publish.serverUnread", ex.Message)
        End Try
    End Function

#End Region

#Region "Build and publish"

    Private Sub AppendLog(line As String)
        If IsDisposed Then Return
        If InvokeRequired Then
            BeginInvoke(Sub() AppendLog(line))
            Return
        End If
        txtLog.AppendText(line & Environment.NewLine)
    End Sub

    Private Sub SetBusy(busy As Boolean)
        _busy = busy
        btnBuild.Enabled = Not busy
        btnExisting.Enabled = Not busy
        btnClose.Enabled = Not busy
        UseWaitCursor = busy
    End Sub

    ''' <summary>Checks the version and folders before building. Returns the version, or Nothing.</summary>
    Private Function CheckBeforePublish(needSource As Boolean) As Version
        Dim v = UpdateService.ParseVersion(txtVersion.Text)
        If v Is Nothing Then
            Ask.Warn(Me, "publish.badVersion")
            Return Nothing
        End If
        If needSource AndAlso Not File.Exists(SetupProject) Then
            Ask.Warn(Me, "publish.noSetupProject", SetupProject)
            Return Nothing
        End If
        If txtOutput.Text.Trim() = "" Then
            Ask.Warn(Me, "publish.noOutput")
            Return Nothing
        End If
        If _serverVersion IsNot Nothing AndAlso v <= _serverVersion AndAlso
           Not Ask.YesNoWarning(Me, "publish.notNewer", UpdateService.VersionText(v), UpdateService.VersionText(_serverVersion)) Then
            Return Nothing
        End If
        Return v
    End Function

    ''' <summary>The program's file name in the installer.</summary>
    Private Const ReleasedExeName As String = "ReadEZ.exe"

    ''' <summary>
    ''' Runs a program and shows its output in the log. Programs are started directly (not through
    ''' MSBuild's temporary .cmd files, which antivirus programs such as AVG block as suspicious).
    ''' shownArgs is what the log shows (without the certificate password).
    ''' </summary>
    Private Async Function RunAsync(exe As String, args As String, shownArgs As String) As Task(Of Integer)
        AppendLog("> " & Path.GetFileName(exe) & " " & shownArgs)
        Dim psi As New ProcessStartInfo(exe, args) With {
            .UseShellExecute = False, .CreateNoWindow = True, .RedirectStandardOutput = True, .RedirectStandardError = True,
            .WorkingDirectory = txtSolution.Text.Trim(), .StandardOutputEncoding = Encoding.UTF8, .StandardErrorEncoding = Encoding.UTF8}
        Using p As New Process With {.StartInfo = psi, .EnableRaisingEvents = True}
            AddHandler p.OutputDataReceived, Sub(s, e) If e.Data IsNot Nothing Then AppendLog(e.Data)
            AddHandler p.ErrorDataReceived, Sub(s, e) If e.Data IsNot Nothing Then AppendLog(e.Data)
            p.Start()
            p.BeginOutputReadLine()
            p.BeginErrorReadLine()
            Await p.WaitForExitAsync()
            Return p.ExitCode
        End Using
    End Function

    ''' <summary>Signs files with signtool and the .pfx certificate. Returns True if it worked.</summary>
    Private Async Function SignAsync(description As String, ParamArray files As String()) As Task(Of Boolean)
        Dim pfx = txtPfx.Text.Trim()
        Dim ts = txtTimestamp.Text.Trim()
        Dim common = $"sign /f ""{pfx}"" {{0}} /fd SHA256" & If(ts = "", "", $" /tr ""{ts}"" /td SHA256") &
                     $" /d ""{description}"" " & String.Join(" ", files.Select(Function(f) $"""{f}"""))
        Dim pw = If(txtPassword.Text = "", "", $"/p ""{txtPassword.Text}""")
        Dim code = Await RunAsync(txtSignTool.Text.Trim(), String.Format(common, pw), String.Format(common, If(pw = "", "", "/p ********")))
        Return code = 0
    End Function

    Private Async Function BuildAndPublishAsync() As Task
        If _busy Then Return
        Dim v = CheckBeforePublish(needSource:=True)
        If v Is Nothing Then Return
        Dim pfx = txtPfx.Text.Trim()
        If pfx = "" Then
            If Not Ask.YesNoWarning(Me, "publish.unsignedConfirm") Then Return
        ElseIf Not File.Exists(pfx) Then
            Ask.Warn(Me, "publish.noPfx", pfx)
            Return
        End If
        If pfx <> "" AndAlso Not File.Exists(txtSignTool.Text.Trim()) Then
            Ask.Warn(Me, "publish.noSignTool")
            Return
        End If
        If Not CheckStarterDatabase() Then Return

        StoreValues()
        txtLog.Clear()
        SetBusy(True)
        Try
            Dim versionText = UpdateService.VersionText(v)
            If ReadProjectVersion() <> versionText Then
                WriteProjectVersion(versionText)
                AppendLog($"<Version> in ReadEZ.vbproj set to {versionText}")
            End If
            Dim setupDir = Path.GetDirectoryName(SetupProject)
            Dim published = Path.Combine(setupDir, "obj", "publish", "Release")
            Dim started = Date.Now

            ' 1. Publish ReadEZ (self-contained, 64-bit)
            If Directory.Exists(published) Then Directory.Delete(published, recursive:=True)
            Dim restoreArgs = $"restore ""{ProjectFile}"" -r win-x64 -nologo -v:m"
            If Await RunAsync("dotnet", restoreArgs, restoreArgs) <> 0 Then
                Ask.Warn(Me, "publish.buildFailed")
                Return
            End If
            Dim pubArgs = $"publish ""{ProjectFile}"" -c Release -r win-x64 --self-contained true -p:DebugType=None -p:DebugSymbols=false -nologo -v:m -o ""{published}"""
            If Await RunAsync("dotnet", pubArgs, pubArgs) <> 0 Then
                Ask.Warn(Me, "publish.buildFailed")
                Return
            End If
            Dim builtVersion = UpdateService.Normalise(AssemblyName.GetAssemblyName(Path.Combine(published, "ReadEZ.dll")).Version)
            If builtVersion <> v Then
                AppendLog($"The program was built as {UpdateService.VersionText(builtVersion)}, not {versionText}.")
                Ask.Warn(Me, "publish.versionMismatch", UpdateService.VersionText(builtVersion), versionText)
                Return
            End If

            ' Starter database (installed on new installations only)
            AddStarterDatabase(published)

            ' 2. Sign the program
            Dim exe = Path.Combine(published, ReleasedExeName)
            If pfx <> "" Then
                If Not Await SignAsync("ReadEZ", exe, Path.Combine(published, "ReadEZ.dll")) Then
                    Ask.Warn(Me, "publish.notSigned")
                    Return
                End If
            End If
            Dim exeSig = CodeSigning.Check(exe)
            AppendLog($"{ReleasedExeName}: " & Describe(exeSig))
            If pfx <> "" AndAlso Not exeSig.IsIntact Then
                Ask.Warn(Me, "publish.notSigned")
                Return
            End If

            ' 3. Pack the installer from that folder (the setup project doesn't publish or sign again)
            Dim wixArgs = $"build ""{SetupProject}"" -c Release --no-incremental -nologo -v:m -p:AppPrePublished=true"
            If Await RunAsync("dotnet", wixArgs, wixArgs) <> 0 Then
                Ask.Warn(Me, "publish.buildFailed")
                Return
            End If
            Dim msi = Directory.EnumerateFiles(Path.Combine(setupDir, "bin"), "*.msi", SearchOption.AllDirectories).
                          Select(Function(f) New FileInfo(f)).Where(Function(f) f.LastWriteTime >= started.AddMinutes(-1)).
                          OrderByDescending(Function(f) f.LastWriteTime).FirstOrDefault()
            If msi Is Nothing Then
                Ask.Warn(Me, "publish.noMsi")
                Return
            End If

            ' 4. Sign the installer
            If pfx <> "" AndAlso Not Await SignAsync("ReadEZ installer", msi.FullName) Then
                Ask.Warn(Me, "publish.notSigned")
                Return
            End If

            ' 5. Copy to the output folder with update.json
            WriteRelease(msi.FullName, v)
        Catch ex As Exception
            Log.Error("Publish update", ex)
            AppendLog("Error: " & ex.Message)
            Ask.Warn(Me, "publish.failed", ex.Message)
        Finally
            If Not IsDisposed Then SetBusy(False)
        End Try
    End Function

    ''' <summary>
    ''' Checks the chosen starter database: it must exist and be an ReadEZ database,
    ''' and the user confirms it holds only example data (it goes to every new installation).
    ''' </summary>
    Private Function CheckStarterDatabase() As Boolean
        Dim f = txtStarter.Text.Trim()
        If f = "" Then Return True
        If Not File.Exists(f) Then
            Ask.Warn(Me, "publish.noStarter", f)
            Return False
        End If
        Try
            Dim counts = CountStarter(f)
            Return Ask.YesNoWarning(Me, "publish.starterConfirm", counts.Patients, counts.Results)
        Catch ex As Exception
            Ask.Warn(Me, "publish.starterBad", f, ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>Patients and recordings in a database (read only; fails if it isn't one).</summary>
    Private Shared Function CountStarter(file As String) As (Patients As Integer, Results As Integer)
        Using c As New System.Data.SQLite.SQLiteConnection($"Data Source={file};Version=3;Read Only=True;")
            c.Open()
            Using cmd = c.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM Patients;"
                Dim p = Convert.ToInt32(cmd.ExecuteScalar())
                cmd.CommandText = "SELECT COUNT(*) FROM Tests;"
                Dim r = Convert.ToInt32(cmd.ExecuteScalar())
                Return (p, r)
            End Using
        End Using
    End Function

    ''' <summary>Copies the starter database into the published program (Starter\Database\ReadEZ.db).</summary>
    Private Sub AddStarterDatabase(published As String)
        Dim source = txtStarter.Text.Trim()
        Dim target = Path.Combine(published, "Starter", "Database", "ReadEZ.db")
        If source = "" Then
            AppLog(If(File.Exists(target), "Starter database: the one in ReadEZ\Deploy is used",
                                         "No starter database: new installations start with an empty database"))
            Return
        End If
        Directory.CreateDirectory(Path.GetDirectoryName(target))
        If File.Exists(target) Then File.Delete(target)
        ' SQLite's backup copies the whole database, even while ReadEZ has it open
        Using src As New System.Data.SQLite.SQLiteConnection($"Data Source={source};Version=3;Read Only=True;"),
              dst As New System.Data.SQLite.SQLiteConnection($"Data Source={target};Version=3;")
            src.Open()
            dst.Open()
            src.BackupDatabase(dst, "main", "main", -1, Nothing, 0)
        End Using
        System.Data.SQLite.SQLiteConnection.ClearAllPools()
        Dim counts = CountStarter(target)
        AppLog($"Starter database: {source} ({counts.Patients} patients, {counts.Results} tests) -> Starter\Database\ReadEZ.db (new installations only)")
    End Sub

    Private Sub AppLog(line As String)
        AppendLog(line)
    End Sub

    Private Shared Function Describe(sig As SignatureInfo) As String
        Return sig.Describe()
    End Function

    ''' <summary>"Use an existing installer": an MSI built another way, with the version in the box.</summary>
    Private Sub PublishExisting()
        If _busy Then Return
        Dim v = CheckBeforePublish(needSource:=False)
        If v Is Nothing Then Return
        Using dlg As New OpenFileDialog With {.Filter = "Windows Installer (*.msi)|*.msi"}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Return
            StoreValues()
            txtLog.Clear()
            Try
                WriteRelease(dlg.FileName, v)
            Catch ex As Exception
                Log.Error("Publish update", ex)
                Ask.Warn(Me, "publish.failed", ex.Message)
            End Try
        End Using
    End Sub

    ''' <summary>Copies the MSI to the output folder, writes update.json and opens the folder.</summary>
    Private Sub WriteRelease(msiSource As String, v As Version)
        Dim versionText = UpdateService.VersionText(v)
        Dim msiSig = CodeSigning.Check(msiSource)
        AppendLog("Installer: " & Describe(msiSig))
        If msiSig.IsSigned AndAlso Not msiSig.IsIntact Then
            ' Installed (signed) copies refuse an installer that was changed after signing
            If Not Ask.YesNoWarning(Me, "publish.signatureInvalid", $"0x{msiSig.ErrorCode:X8}") Then Return
        End If

        Dim outDir = txtOutput.Text.Trim()
        Directory.CreateDirectory(outDir)
        Dim fileName = $"ReadEZ-Setup-{versionText}.msi"
        Dim target = Path.Combine(outDir, fileName)
        If Not String.Equals(Path.GetFullPath(msiSource), Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase) Then
            File.Copy(msiSource, target, overwrite:=True)
        End If
        Dim size = New FileInfo(target).Length
        Dim hash = UpdateService.FileSha256(target)

        Dim manifest As New Dictionary(Of String, Object) From {
            {"product", "ReadEZ"},
            {"version", versionText},
            {"released", Date.Today.ToString("yyyy-MM-dd")},
            {"file", fileName},
            {"size", size},
            {"sha256", hash},
            {"notes", txtNotes.Text.Trim()}}
        Dim json = JsonSerializer.Serialize(manifest, New JsonSerializerOptions With {
            .WriteIndented = True, .TypeInfoResolver = New Serialization.Metadata.DefaultJsonTypeInfoResolver(),
            .Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping})
        Dim manifestFile = Path.Combine(outDir, UpdateService.ManifestFileName)
        File.WriteAllText(manifestFile, json, New UTF8Encoding(False))
        ' Check that it reads back as the program will read it
        UpdateService.ParseManifest(json, txtUrl.Text.Trim())

        AppendLog($"Written {target} ({size:N0} bytes)")
        AppendLog($"Written {manifestFile}")
        AppendLog($"Upload both files to the folder of {txtUrl.Text.Trim()}")
        Log.Info($"Published update {versionText}: {target}")
        Process.Start(New ProcessStartInfo("explorer.exe", $"/select,""{target}""") With {.UseShellExecute = True})
        Ask.Info(Me, "publish.done", versionText, outDir)
    End Sub

#End Region

End Class

''' <summary>Encrypts a secret for the current Windows user (Data Protection API).</summary>
Friend NotInheritable Class Dpapi
    Private Sub New()
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure DATA_BLOB
        Public cbData As Integer
        Public pbData As IntPtr
    End Structure

    <DllImport("crypt32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function CryptProtectData(ByRef pDataIn As DATA_BLOB, szDataDescr As String, pOptionalEntropy As IntPtr,
                                             pvReserved As IntPtr, pPromptStruct As IntPtr, dwFlags As Integer, ByRef pDataOut As DATA_BLOB) As Boolean
    End Function

    <DllImport("crypt32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Shared Function CryptUnprotectData(ByRef pDataIn As DATA_BLOB, pszDataDescr As IntPtr, pOptionalEntropy As IntPtr,
                                               pvReserved As IntPtr, pPromptStruct As IntPtr, dwFlags As Integer, ByRef pDataOut As DATA_BLOB) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Private Shared Function LocalFree(hMem As IntPtr) As IntPtr
    End Function

    Private Const CRYPTPROTECT_UI_FORBIDDEN As Integer = 1

    Public Shared Function Protect(text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Try
            Return Convert.ToBase64String(Run(Encoding.UTF8.GetBytes(text), protect:=True))
        Catch ex As Exception
            Log.Warn("DPAPI protect: " & ex.Message)
            Return ""
        End Try
    End Function

    Public Shared Function Unprotect(stored As String) As String
        If String.IsNullOrEmpty(stored) Then Return ""
        Try
            Return Encoding.UTF8.GetString(Run(Convert.FromBase64String(stored), protect:=False))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Shared Function Run(data As Byte(), protect As Boolean) As Byte()
        Dim input As New DATA_BLOB With {.cbData = data.Length, .pbData = Marshal.AllocHGlobal(data.Length)}
        Dim output As New DATA_BLOB
        Try
            Marshal.Copy(data, 0, input.pbData, data.Length)
            Dim ok = If(protect,
                        CryptProtectData(input, "ReadEZ", IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, CRYPTPROTECT_UI_FORBIDDEN, output),
                        CryptUnprotectData(input, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, CRYPTPROTECT_UI_FORBIDDEN, output))
            If Not ok Then Throw New ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
            Dim result(output.cbData - 1) As Byte
            Marshal.Copy(output.pbData, result, 0, output.cbData)
            Return result
        Finally
            Marshal.FreeHGlobal(input.pbData)
            If output.pbData <> IntPtr.Zero Then LocalFree(output.pbData)
        End Try
    End Function
End Class
