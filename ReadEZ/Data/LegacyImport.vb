Option Strict On

Imports System.IO

' ====================================================================================================
'   LegacyImport.vb - brings the settings of ReadEZ 2020 across (once, on the first start)
'
'   ReadEZ 2020 kept its settings in RZSettings.ini next to the program ([ReadEZ Pro] section):
'   database, practice name and logo, window positions, the reading-text settings and the 39
'   screen colours (OLA ... LSGREY). They are read from the first RZSettings.ini found in the usual
'   places, or from a file chosen in Admin > Import from ReadEZ 2020.
'
'   The colours are brought across as they were matched in ReadEZ 2020's colour calibration, but
'   are marked as not checked: they should be checked again with Colour calibration, which also
'   records the screen they were checked on.
' ====================================================================================================

Public NotInheritable Class LegacyImport
    Private Sub New()
    End Sub

    Private Shared ReadOnly Inv As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture

    ''' <summary>Where ReadEZ 2020 may have been installed.</summary>
    Public Shared Function CandidateIniFiles() As IEnumerable(Of String)
        Dim list As New List(Of String) From {Path.Combine(AppHost.AppFolder, "RZSettings.ini")}
        For Each root In {Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                          Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                          Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}
            If String.IsNullOrEmpty(root) Then Continue For
            For Each company In {"Thomson Software Solutions", "Thomson Software Solutions Ltd", ""}
                For Each product In {"ReadEZ 2020", "ReadEZ", "ReadEZ Pro"}
                    list.Add(Path.Combine(root, company, product, "RZSettings.ini"))
                Next
            Next
        Next
        Return list.Distinct(StringComparer.OrdinalIgnoreCase)
    End Function

    ''' <summary>Settings from ReadEZ 2020, or Nothing if it isn't found.</summary>
    Public Shared Function TryImportSettings() As AppSettings
        For Each f In CandidateIniFiles()
            Try
                If File.Exists(f) Then
                    Dim s = FromIni(f)
                    Log.Info("Settings imported from " & f)
                    Return s
                End If
            Catch ex As Exception
                Log.Error("Import " & f, ex)
            End Try
        Next
        Return Nothing
    End Function

    ''' <summary>There are no administrator settings in ReadEZ 2020.</summary>
    Public Shared Function TryImportAdmin() As AdminSettings
        Return Nothing
    End Function

    ''' <summary>Reads an RZSettings.ini into new settings (and stores its screen colours).</summary>
    Public Shared Function FromIni(file As String) As AppSettings
        Dim ini = ReadIni(file)
        Dim s = AppSettings.CreateDefault()
        Dim get1 = Function(key As String) As String
                       Dim v As String = Nothing
                       Return If(ini.TryGetValue(key, v), v.Trim(), "")
                   End Function
        Dim getInt = Sub(key As String, ByRef target As Integer)
                         Dim n As Integer
                         If Integer.TryParse(get1(key), Globalization.NumberStyles.Integer, Inv, n) Then target = n
                     End Sub
        Dim getBool = Sub(key As String, ByRef target As Boolean)
                          Dim b As Boolean
                          If Boolean.TryParse(get1(key), b) Then target = b
                      End Sub

        Dim db = get1("Local_Database")
        If db <> "" AndAlso IO.File.Exists(db) Then s.Local_Database = db
        Dim dataDir = get1("DataDirectory")
        If dataDir <> "" AndAlso Directory.Exists(dataDir) Then s.DataDirectory = dataDir
        If get1("Practice_Name") <> "" Then s.Practice_Name = get1("Practice_Name")
        getBool("ShowLogo", s.ShowLogo)
        Dim logo = get1("LogoFilename")
        If logo <> "" AndAlso IO.File.Exists(logo) Then s.LogoFilename = logo
        getInt("LogoScale", s.LogoScale)
        getInt("ScreenerID", s.ScreenerID)

        ' The patient's screen was the "Test" window
        Dim t As New Rectangle
        getInt("Test_Left", t.X) : getInt("Test_Top", t.Y) : getInt("Test_Width", t.Width) : getInt("Test_Height", t.Height)
        If ScreenLayout.IsOnScreen(t) Then
            s.Stimulus_Left = t.X : s.Stimulus_Top = t.Y : s.Stimulus_Width = t.Width : s.Stimulus_Height = t.Height
        End If
        getInt("Screen_Width", s.Screen_Width)

        getInt("RoR_Random", s.RoR_Random)
        getInt("RoR_LineSpacing", s.RoR_LineSpacing)
        getInt("RoR_FontSize", s.RoR_FontSize)
        If get1("RoR_FontName") <> "" Then s.RoR_FontName = get1("RoR_FontName")
        getInt("RoR_NoRows", s.RoR_NoRows)
        getInt("RoR_TextOptions", s.RoR_TextOptions)
        getInt("Overlay_Size", s.Overlay_Size)
        getInt("Timing", s.Timing)
        s.FillMissing()

        ImportColours(ini)
        Return s
    End Function

    ''' <summary>The 39 screen colours of RZSettings.ini (OLA ... OLL, OLGREY, LA ..., LSA ...), stored as not checked.</summary>
    Private Shared Sub ImportColours(ini As Dictionary(Of String, String))
        Dim prefixes = {"OL", "L", "LS"}
        Dim colours = New Color(2)() {}
        Dim checkedNow = New Boolean(2)() {}
        Dim found = 0
        For setIndex = 0 To 2
            colours(setIndex) = New Color(ReadEZColours.PerSet - 1) {}
            checkedNow(setIndex) = New Boolean(ReadEZColours.PerSet - 1) {}
            For i = 0 To ReadEZColours.PerSet - 1
                Dim key = prefixes(setIndex) & If(i = 0, "GREY", ChrW(AscW("A"c) + i - 1).ToString())
                Dim c = ReadEZColours.Colour(CType(setIndex, ReadEZSet), i)
                Dim v As String = Nothing
                If ini.TryGetValue(key, v) Then
                    Dim parts = v.Split(","c)
                    Dim r, g, b As Integer
                    If parts.Length = 3 AndAlso Integer.TryParse(parts(0).Trim(), r) AndAlso Integer.TryParse(parts(1).Trim(), g) AndAlso
                       Integer.TryParse(parts(2).Trim(), b) Then
                        c = Color.FromArgb(255, Math.Max(0, Math.Min(255, r)), Math.Max(0, Math.Min(255, g)), Math.Max(0, Math.Min(255, b)))
                        found += 1
                    End If
                End If
                colours(setIndex)(i) = c
            Next
        Next
        If found > 0 Then ReadEZColours.Store(colours, checkedNow)
    End Sub

    Private Shared Function ReadIni(file As String) As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each line In IO.File.ReadAllLines(file)
            Dim l = line.Trim()
            If l = "" OrElse l.StartsWith(";") OrElse l.StartsWith("[") Then Continue For
            Dim eq = l.IndexOf("="c)
            If eq <= 0 Then Continue For
            d(l.Substring(0, eq).Trim()) = l.Substring(eq + 1)
        Next
        Return d
    End Function
End Class
