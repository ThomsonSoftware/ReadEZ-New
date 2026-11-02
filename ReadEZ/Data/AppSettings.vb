Option Strict On

Imports System.IO

' ====================================================================================================
'   AppSettings.vb
'   Program settings. Replaces REZSettingsStructure + RZSettings.ini. Stored as readable JSON in:
'
'       %AppData%\ReadEZ\Settings.json
'
'   Field names are the same as ReadEZ 2020 where the setting existed there, so the old
'   RZSettings.ini can be imported (LegacyImport). The ReadEZ colours themselves are kept in
'   %AppData%\ReadEZ\ReadEZ Colours.json (ReadEZColours.vb), with the date each was checked.
' ====================================================================================================

Public NotInheritable Class AppSettings

#Region "Practice branding"
    Public Practice_Name As String = ""
    Public ShowLogo As Boolean = True
    Public LogoFilename As String = ""
    Public LogoScale As Integer = 60
#End Region

#Region "Interface"
    ''' <summary>"English", "Spanish", "German" or "Italian".</summary>
    Public Language As String = "English"
    Public Theme As ThemeMode = ThemeMode.Light
    Public NavCollapsed As Boolean = False
    Public MainWindowBounds As Rectangle = Rectangle.Empty
    Public MainWindowMaximised As Boolean = True
    Public LastPatientId As Integer = -1
    ''' <summary>The person running the tests (Tester table; 0 = none chosen).</summary>
    Public ScreenerID As Integer = 0
    ''' <summary>Patients screen: show only this organisation's patients (-1 = all).</summary>
    Public OrganisationFilter As Integer = -1
#End Region

#Region "Updates"
    ''' <summary>Look for a new version on the web each time the program starts.</summary>
    Public Update_CheckAtStart As Boolean = True
    ''' <summary>A version the user chose to skip (not offered again at start-up).</summary>
    Public Update_SkipVersion As String = ""
#End Region

#Region "File system paths"
    ''' <summary>The ReadEZ database (same format as ReadEZ 2020's ReadEZ2020.db).</summary>
    Public Local_Database As String = ""
    ''' <summary>Data folder: reading texts, report letters and pictures.</summary>
    Public DataDirectory As String = ""
#End Region

#Region "Backup"
    Public BackupFolder As String = ""
    Public BackupOnClose As Boolean = False
    ''' <summary>Include the data folder (texts, letters, pictures) as well as the database.</summary>
    Public BackupIncludeStimuli As Boolean = True
    Public BackupLastStamp As String = ""
#End Region

#Region "Window positions"
    ' Control window (the clinician's window)
    Public Control_Top As Integer
    Public Control_Left As Integer
    Public Control_Height As Integer
    Public Control_Width As Integer
    ' Patient's screen (Test_* in ReadEZ 2020)
    Public Stimulus_Top As Integer
    Public Stimulus_Left As Integer
    Public Stimulus_Height As Integer
    Public Stimulus_Width As Integer
    ''' <summary>Physical width of the patient's screen in mm (for sizes on the screen in mm).</summary>
    Public Screen_Width As Integer = 500
    ''' <summary>Physical height of the patient's screen in mm.</summary>
    Public Screen_Height As Integer = 300
#End Region

#Region "Reading text (symptoms and rate of reading)"
    ''' <summary>Index of the text (ReadingTexts: built-in texts first, then the data folder's own).</summary>
    Public RoR_TextOptions As Integer = 0
    ''' <summary>0 = rows in order, 1 = rows in a mixed order (ReadEZ 2020: RoR_Random).</summary>
    Public RoR_Random As Integer = 0
    ''' <summary>Line spacing in tenths added to 0.8 (6 = 1.4 x the line height).</summary>
    Public RoR_LineSpacing As Integer = 6
    ''' <summary>Font size setting; the text is drawn at this + 8 points (as ReadEZ 2020).</summary>
    Public RoR_FontSize As Integer = 5
    Public RoR_FontName As String = "Times New Roman"
    Public RoR_NoRows As Integer = 20
    ''' <summary>Overlay mode: how far the colour extends round the text (0 - 10).</summary>
    Public Overlay_Size As Integer = 5
    ''' <summary>Lenses / computer: how long each colour is shown (ms).</summary>
    Public Timing As Integer = 3000
    ''' <summary>Rate of reading: how long the patient reads (s).</summary>
    Public RoR_ReadingTime As Integer = 60
    ''' <summary>Rate of reading: sound when the time is up.</summary>
    Public RoR_Beep As Boolean = True
#End Region

#Region "Tests"
    ''' <summary>The colour test used last: 0 automatic, 1 manual.</summary>
    Public ColourTestMethod As Integer = 0
    ''' <summary>After choosing a colour, go straight on to rate of reading and symptoms with it.</summary>
    Public ContinueAfterColour As Boolean = True
    ''' <summary>One screen: show the patient's text in the whole window (F11) while testing.</summary>
    Public SingleScreenFullWindow As Boolean = False
#End Region

#Region "Derived values"

    Public ReadOnly Property UiLanguage As AppLanguage
        Get
            Return Loc.Parse(Language)
        End Get
    End Property

    Public ReadOnly Property StimulusBounds As Rectangle
        Get
            Return New Rectangle(Stimulus_Left, Stimulus_Top, Stimulus_Width, Stimulus_Height)
        End Get
    End Property

    Public ReadOnly Property ControlBounds As Rectangle
        Get
            Return New Rectangle(Control_Left, Control_Top, Control_Width, Control_Height)
        End Get
    End Property

    ''' <summary>The text font size in points.</summary>
    Public ReadOnly Property TextPoints As Single
        Get
            Return RoR_FontSize + 8
        End Get
    End Property

#End Region

#Region "Defaults"

    ''' <summary>Settings for a first run, using the attached screens.</summary>
    Public Shared Function CreateDefault() As AppSettings
        Dim s As New AppSettings()
        Dim dataRoot = AppHost.DefaultDataFolder
        s.DataDirectory = dataRoot
        s.Local_Database = Path.Combine(dataRoot, "Database", "ReadEZ.db")
        s.LogoFilename = Path.Combine(dataRoot, "Report images", "ReadEZ-logo.png")
        s.ApplyDefaultScreens()
        Return s
    End Function

    ''' <summary>
    ''' Control window on the primary screen; patient's screen on the last screen. With one screen
    ''' both are on it (one-screen mode), as in ReadEZ 2020 and Eye Dynamix.
    ''' </summary>
    Public Sub ApplyDefaultScreens()
        Dim primary = Screen.PrimaryScreen.WorkingArea
        Control_Top = primary.Top
        Control_Left = primary.Left
        Control_Height = primary.Height
        Control_Width = primary.Width
        Dim stim = ScreenLayout.PrimaryLeadingScreens().Last().Bounds
        Stimulus_Top = stim.Top
        Stimulus_Left = stim.Left
        Stimulus_Height = stim.Height
        Stimulus_Width = stim.Width
    End Sub

#End Region

#Region "Load and save"

    Public Shared ReadOnly Property SettingsFolder As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ReadEZ")
        End Get
    End Property

    Public Shared ReadOnly Property SettingsFile As String
        Get
            Return Path.Combine(SettingsFolder, "Settings.json")
        End Get
    End Property

    ''' <summary>Loads the settings file. Returns Nothing if there isn't one yet.</summary>
    Public Shared Function TryLoad() As AppSettings
        If Not File.Exists(SettingsFile) Then Return Nothing
        Dim s As AppSettings
        Try
            s = SettingsJson.Read(Of AppSettings)(File.ReadAllText(SettingsFile))
        Catch ex As Exception
            SettingsJson.SetAsideDamagedFile(SettingsFile, ex)
            Return Nothing
        End Try
        s.FillMissing()
        Return s
    End Function

    Public Sub Save()
        SettingsJson.WriteFileSafely(SettingsFile, SettingsJson.Write(Me))
    End Sub

    ''' <summary>Fills blank or out-of-range values (for example from an older settings file).</summary>
    Public Sub FillMissing()
        Dim d = CreateDefault()
        If String.IsNullOrWhiteSpace(Local_Database) Then Local_Database = d.Local_Database
        If String.IsNullOrWhiteSpace(DataDirectory) Then DataDirectory = d.DataDirectory
        If String.IsNullOrWhiteSpace(Language) Then Language = "English"
        If Stimulus_Width <= 0 OrElse Stimulus_Height <= 0 OrElse Control_Width <= 0 Then ApplyDefaultScreens()
        If Screen_Width <= 0 Then Screen_Width = 500
        If Screen_Height <= 0 Then Screen_Height = 300
        If String.IsNullOrWhiteSpace(RoR_FontName) Then RoR_FontName = "Times New Roman"
        RoR_FontSize = Math.Max(0, Math.Min(40, RoR_FontSize))
        RoR_LineSpacing = Math.Max(0, Math.Min(30, RoR_LineSpacing))
        RoR_NoRows = Math.Max(1, Math.Min(20, RoR_NoRows))
        Overlay_Size = Math.Max(0, Math.Min(10, Overlay_Size))
        If Timing < 500 Then Timing = 3000
        If RoR_ReadingTime <= 0 Then RoR_ReadingTime = 60
    End Sub

    Public Function Clone() As AppSettings
        Return SettingsJson.Read(Of AppSettings)(SettingsJson.Write(Me))
    End Function

#End Region

End Class
