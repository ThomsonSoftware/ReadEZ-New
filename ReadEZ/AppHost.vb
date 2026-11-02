Option Strict On

Imports System.IO
Imports System.Reflection

' ====================================================================================================
'   AppHost.vb - start-up and program-wide services
'
'   Start-up order:
'     1. Settings (imported from ReadEZ 2020's RZSettings.ini the first time)
'     2. Admin settings (password, help and update addresses)
'     3. Theme and language
'     4. Product licence (activation)
'     5. Database
'     6. Main window
'
'   Everything that used to be a loose global in the ReadEZ module is reached through AppHost:
'   AppHost.Settings, AppHost.Admin, AppHost.Db, AppHost.Session.
' ====================================================================================================

Public Module AppHost

    Public Const ProductTitle As String = "ReadEZ"
    Public Const ShortTitle As String = "ReadEZ"
    ''' <summary>Program version (from &lt;Version&gt; in ReadEZ.vbproj), e.g. "3.0.0".</summary>
    Public ReadOnly Property VersionNumber As String
        Get
            Return UpdateService.VersionText(UpdateService.CurrentVersion)
        End Get
    End Property

    ''' <summary>Unique Device Identifier printed on reports and the About box.</summary>
    Public Const UdiText As String = ""

    Public Property Settings As AppSettings
    Public Property Admin As AdminSettings
    ''' <summary>
    ''' True when ReadEZ has never run on this computer (no settings file, and none from the
    ''' ReadEZ 2020): only then is the starter database installed (StandardFiles).
    ''' </summary>
    Public Property FirstRun As Boolean
    Public Property Db As ReadEZDatabase
    Public Property Session As New SessionState

    Private _icon As Icon
    Private _logo As Image

#Region "Entry point"

    <STAThread>
    Public Sub Main()
#If Not COMPILECHECK Then
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2)
#End If
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)
        AddHandler Application.ThreadException, Sub(s, e) ShowError(e.Exception)
        AddHandler AppDomain.CurrentDomain.UnhandledException, Sub(s, e) Log.Error("Unhandled", TryCast(e.ExceptionObject, Exception))

        ' The splash first, so something shows at once; it closes itself once the main window is up
        SplashForm.ShowNow()
        Log.Info($"{ProductTitle} {VersionNumber} starting")

        ' 1-2. Settings
        If Not LoadSettings() Then Return
        ' Starter databases now (the database opens below); stimuli, sounds and report pictures are
        ' copied in the background once the main window is showing (MainForm.OnShown)
        StandardFiles.EnsureStarterDatabases()

        ' 3. Theme and language
        ThemeManager.SetMode(Settings.Theme)
        Loc.SetLanguage(Settings.UiLanguage)

        ' 4. Product licence
        If Not LicenceService.EnsureLicensed() Then
            SplashForm.CloseNow()
            Return
        End If

        ' 5. Database
        If Not OpenDatabase() Then Return

        Application.AddMessageFilter(New EmergencyShow())   ' Ctrl+S: bring the control window back
        Application.Run(New MainForm())

        Log.Info("Closed")
    End Sub

#End Region

#Region "Settings"

    Private Function LoadSettings() As Boolean
        Try
            Settings = AppSettings.TryLoad()
            If Settings Is Nothing Then
                ' First run of this version: bring settings across from the previous version if possible.
                Settings = LegacyImport.TryImportSettings()
                If Settings Is Nothing Then
                    Settings = AppSettings.CreateDefault()
                    FirstRun = True
                End If
                Settings.FillMissing()
                Settings.Save()
            End If
            Admin = AdminSettings.TryLoad()
            If Admin Is Nothing Then
                Admin = If(LegacyImport.TryImportAdmin(), New AdminSettings())
                Admin.Save()
            End If
            Return True
        Catch ex As Exception
            Log.Error("LoadSettings", ex)
            SplashForm.CloseNow()
            MessageBox.Show(ProductTitle & " could not read its settings file:" & Environment.NewLine & AppSettings.SettingsFile &
                            Environment.NewLine & Environment.NewLine & ex.Message,
                            ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>Saves settings, reporting (not throwing) any failure.</summary>
    Public Sub SaveSettings()
        Try
            Settings.Save()
        Catch ex As Exception
            Log.Error("SaveSettings", ex)
        End Try
    End Sub

    Public Sub SaveAdmin()
        Try
            Admin.Save()
        Catch ex As Exception
            Log.Error("SaveAdmin", ex)
        End Try
    End Sub

#End Region

#Region "Database"

    Private Function OpenDatabase() As Boolean
        Do
            Try
                Dim path = Settings.Local_Database
                ' A ReadEZ 2020 database chosen as the database: bring it into a new one beside it
                If ReadEZDatabase.IsLegacyFile(path) Then
                    Dim oldFile = path
                    path = IO.Path.Combine(IO.Path.GetDirectoryName(oldFile), "ReadEZ.db")
                    Settings.Local_Database = path
                    SaveSettings()
                    Db = New ReadEZDatabase(path)
                    ImportOldDatabase(oldFile, askFirst:=False)
                    Return True
                End If
                Directory.CreateDirectory(IO.Path.GetDirectoryName(path))
                Db = New ReadEZDatabase(path)
                ' First run: offer to bring in the patients and results of ReadEZ 2020
                If FirstRun OrElse Db.PatientCount() = 0 Then
                    Dim oldDb = LegacyDatabaseImport.FindOldDatabase()
                    If oldDb IsNot Nothing AndAlso Db.GetMeta("ImportedFrom." & IO.Path.GetFileName(oldDb)) = "" Then ImportOldDatabase(oldDb, askFirst:=True)
                End If
                Return True
            Catch ex As Exception
                Log.Error("OpenDatabase", ex)
                SplashForm.CloseNow()
                Dim answer = MessageBox.Show(Loc.F("db.openFailed", Settings.Local_Database, ex.Message), ProductTitle,
                                             MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error)
                Select Case answer
                    Case DialogResult.Retry
                        ' loop
                    Case DialogResult.Ignore
                        Using dlg As New OpenFileDialog With {.Filter = "ReadEZ database (*.db)|*.db|All files (*.*)|*.*", .CheckFileExists = False}
                            If dlg.ShowDialog() <> DialogResult.OK Then Return False
                            Settings.Local_Database = dlg.FileName
                            SaveSettings()
                        End Using
                    Case Else
                        Return False
                End Select
            End Try
        Loop
    End Function

    ''' <summary>Imports a ReadEZ 2020 database into the open database (asking first if wanted).</summary>
    Public Sub ImportOldDatabase(oldFile As String, askFirst As Boolean, Optional owner As IWin32Window = Nothing)
        SplashForm.CloseNow()
        If askFirst AndAlso MessageBox.Show(owner, Loc.F("import.offer", oldFile), ProductTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        Try
            Dim sum = LegacyDatabaseImport.Import(oldFile, Db)
            MessageBox.Show(owner, Loc.F("import.done", sum.Patients, sum.Tests, sum.AlreadyThere, sum.Skipped), ProductTitle,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Log.Error("Import " & oldFile, ex)
            MessageBox.Show(owner, Loc.F("import.failed", oldFile, ex.Message), ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

#Region "Folders and resources"

    ''' <summary>Folder the program runs from.</summary>
    Public ReadOnly Property AppFolder As String
        Get
            Return AppContext.BaseDirectory
        End Get
    End Property

    ''' <summary>Default data folder: Documents\ReadEZ Data (database, reading texts, report letters and pictures).</summary>
    Public ReadOnly Property DefaultDataFolder As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ReadEZ Data")
        End Get
    End Property

    Public Function FileInAppFolder(ParamArray parts As String()) As String
        Dim all = New List(Of String) From {AppFolder}
        all.AddRange(parts)
        Return Path.Combine(all.ToArray())
    End Function

    Public ReadOnly Property AppIcon As Icon
        Get
            If _icon Is Nothing Then
                Using s = Assembly.GetExecutingAssembly().GetManifestResourceStream("ReadEZ.ico")
                    If s IsNot Nothing Then _icon = New Icon(s)
                End Using
            End If
            Return _icon
        End Get
    End Property

    Public ReadOnly Property LogoImage As Image
        Get
            If _logo Is Nothing Then _logo = LoadImageResource("ThomsonLogo.png")
            Return _logo
        End Get
    End Property

    Public Function LoadImageResource(name As String) As Image
        Dim s = Assembly.GetExecutingAssembly().GetManifestResourceStream(name)
        If s Is Nothing Then Return Nothing
        Return Image.FromStream(s)
    End Function

#End Region

#Region "Helpers"

    Public Sub OpenUrl(target As String)
        Try
            Process.Start(New ProcessStartInfo(target) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show(Loc.F("common.couldNotOpen", target) & Environment.NewLine & ex.Message, ProductTitle,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    Public Sub ShowError(ex As Exception)
        Log.Error("ThreadException", ex)
        SplashForm.CloseNow()
        MessageBox.Show(Loc.T("common.somethingWentWrong") & Environment.NewLine & Environment.NewLine & ex.Message,
                        ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>The main window, once it exists.</summary>
    Public ReadOnly Property MainWindow As MainForm
        Get
            Return Application.OpenForms.OfType(Of MainForm)().FirstOrDefault()
        End Get
    End Property

#End Region

End Module
