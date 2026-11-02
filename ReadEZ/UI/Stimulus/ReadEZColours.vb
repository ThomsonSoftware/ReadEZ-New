Option Strict On

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.Json
Imports System.Text.Json.Nodes

' ====================================================================================================
'   ReadEZColours.vb - the ReadEZ overlay and lens colours (as in ReadEZ: Default.pal)
'
'   Three sets of 13 colours: grey (index 0) and A - L (1 - 12):
'     Overlays                 the ReadEZ overlays
'     Lenses                   the ReadEZ lenses (palette set 2 in ReadEZ)
'     Lenses (strong)          the higher-saturation lenses (palette set 3 in ReadEZ)
'
'   The screen colour that matches a real overlay or lens depends on the screen, so each colour can
'   be checked (and adjusted) on the patient's screen (ReadEZColourCheck). The colours, and which
'   screen and date each was checked on, are kept in %AppData%\ReadEZ\ReadEZ Colours.json.
'   The screen is identified by the monitor's hardware id and resolution, so a colour checked on
'   one monitor shows as not checked when another monitor is used for the patient.
'
'   A chosen colour is recorded with the stimulus as a code in stimulus_String2:
'     "O:A" overlay A, "L:C" lens C, "S:F" strong lens F, "O:Grey" ...
'   and its actual screen colour in stimulus_backcolour (so playback shows what the patient saw).
' ====================================================================================================

Public Enum ReadEZSet
    Overlay = 0
    Lens = 1
    LensStrong = 2
End Enum

Public NotInheritable Class ReadEZColours
    Private Sub New()
    End Sub

    ''' <summary>Grey and A - L.</summary>
    Public Const PerSet As Integer = 13
    Public Shared ReadOnly Sets As ReadEZSet() = {ReadEZSet.Overlay, ReadEZSet.Lens, ReadEZSet.LensStrong}

    Private Shared ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    ' ReadEZ's own palette (Palettes\Default.pal), grey first then A - L
    Private Shared ReadOnly Defaults As Color()() = {
        New Color() {Rgb(220, 220, 220), Rgb(255, 219, 122), Rgb(255, 226, 95), Rgb(239, 224, 100), Rgb(211, 222, 134),
                     Rgb(197, 223, 187), Rgb(184, 202, 222), Rgb(186, 178, 215), Rgb(224, 181, 227), Rgb(247, 170, 208),
                     Rgb(255, 162, 185), Rgb(255, 181, 182), Rgb(255, 173, 100)},
        New Color() {Rgb(220, 220, 220), Rgb(245, 223, 152), Rgb(253, 249, 182), Rgb(237, 254, 195), Rgb(208, 253, 214),
                     Rgb(212, 239, 221), Rgb(199, 228, 252), Rgb(215, 225, 254), Rgb(226, 208, 253), Rgb(246, 205, 254),
                     Rgb(254, 199, 227), Rgb(237, 182, 190), Rgb(253, 208, 168)},
        New Color() {Rgb(220, 220, 220), Rgb(222, 187, 79), Rgb(249, 241, 148), Rgb(234, 251, 142), Rgb(179, 252, 199),
                     Rgb(175, 242, 241), Rgb(158, 208, 251), Rgb(179, 203, 252), Rgb(204, 177, 254), Rgb(245, 180, 253),
                     Rgb(254, 167, 208), Rgb(230, 136, 154), Rgb(254, 170, 109)}}

    Private Class Entry
        Public Colour As Color
        Public CheckedOn As Date?
        Public Screen As String = ""
    End Class

    Private Shared _entries As Entry()()
    Private Shared _screenId As String
    Private Shared _screenName As String

    ''' <summary>Raised when a colour is changed (after a check).</summary>
    Public Shared Event Changed As EventHandler

    Private Shared Function Rgb(r As Integer, g As Integer, b As Integer) As Color
        Return Color.FromArgb(255, r, g, b)
    End Function

#Region "Colours"

    Public Shared Function DefaultColour(s As ReadEZSet, index As Integer) As Color
        Return Defaults(CInt(s))(index)
    End Function

    ''' <summary>The colour to show on this computer's patient screen.</summary>
    Public Shared Function Colour(s As ReadEZSet, index As Integer) As Color
        EnsureLoaded()
        Return _entries(CInt(s))(index).Colour
    End Function

    ''' <summary>When the colour was last checked on the screen now used for the patient (Nothing if not).</summary>
    Public Shared Function CheckedOn(s As ReadEZSet, index As Integer) As Date?
        EnsureLoaded()
        Dim e = _entries(CInt(s))(index)
        If e.CheckedOn.HasValue AndAlso e.Screen = ScreenId Then Return e.CheckedOn
        Return Nothing
    End Function

    ''' <summary>Checked, but on a different screen from the one now used for the patient.</summary>
    Public Shared Function CheckedElsewhere(s As ReadEZSet, index As Integer) As Boolean
        EnsureLoaded()
        Dim e = _entries(CInt(s))(index)
        Return e.CheckedOn.HasValue AndAlso e.Screen <> ScreenId
    End Function

    ''' <summary>How many of the 39 colours have been checked on this screen.</summary>
    Public Shared Function CheckedCount() As Integer
        Dim n = 0
        For Each s In Sets
            For i = 0 To PerSet - 1
                If CheckedOn(s, i).HasValue Then n += 1
            Next
        Next
        Return n
    End Function

    ''' <summary>Stores the colours from a check (checked = matched on this screen today) and saves them.</summary>
    Public Shared Sub Store(colours As Color()(), checkedNow As Boolean()())
        EnsureLoaded()
        For Each s In Sets
            For i = 0 To PerSet - 1
                Dim e = _entries(CInt(s))(i)
                e.Colour = colours(CInt(s))(i)
                If checkedNow(CInt(s))(i) Then
                    e.CheckedOn = Date.Today
                    e.Screen = ScreenId
                End If
            Next
        Next
        Save()
        RaiseEvent Changed(Nothing, EventArgs.Empty)
    End Sub

#End Region

#Region "Names and codes"

    ''' <summary>"A" ... "L", or the word grey.</summary>
    Public Shared Function Letter(index As Integer) As String
        If index = 0 Then Return Loc.T("readez.grey")
        Return ChrW(AscW("A"c) + index - 1)
    End Function

    ''' <summary>"Overlay C", "Lens C", "Strong lens C".</summary>
    Public Shared Function Name(s As ReadEZSet, index As Integer) As String
        Select Case s
            Case ReadEZSet.Overlay : Return Loc.F("readez.overlayName", Letter(index))
            Case ReadEZSet.Lens : Return Loc.F("readez.lensName", Letter(index))
            Case Else : Return Loc.F("readez.strongName", Letter(index))
        End Select
    End Function

    Public Shared Function Code(s As ReadEZSet, index As Integer) As String
        Dim prefix = If(s = ReadEZSet.Overlay, "O", If(s = ReadEZSet.Lens, "L", "S"))
        Return prefix & ":" & If(index = 0, "Grey", ChrW(AscW("A"c) + index - 1).ToString())
    End Function

    Public Shared Function TryParse(text As String, ByRef s As ReadEZSet, ByRef index As Integer) As Boolean
        If String.IsNullOrEmpty(text) OrElse text.Length < 3 OrElse text(1) <> ":"c Then Return False
        Select Case text(0)
            Case "O"c : s = ReadEZSet.Overlay
            Case "L"c : s = ReadEZSet.Lens
            Case "S"c : s = ReadEZSet.LensStrong
            Case Else : Return False
        End Select
        Dim rest = text.Substring(2)
        If rest = "Grey" Then index = 0 : Return True
        If rest.Length = 1 AndAlso rest(0) >= "A"c AndAlso rest(0) <= "L"c Then index = AscW(rest(0)) - AscW("A"c) + 1 : Return True
        Return False
    End Function

    ''' <summary>The name for a recorded code ("" if it isn't one).</summary>
    Public Shared Function NameForCode(code As String) As String
        Dim s As ReadEZSet, i As Integer
        Return If(TryParse(code, s, i), Name(s, i), "")
    End Function

#End Region

#Region "The patient's screen"

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure DISPLAY_DEVICE
        Public cb As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public DeviceName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Public DeviceString As String
        Public StateFlags As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Public DeviceID As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Public DeviceKey As String
    End Structure

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function EnumDisplayDevices(lpDevice As String, iDevNum As Integer, ByRef lpDisplayDevice As DISPLAY_DEVICE, dwFlags As Integer) As Boolean
    End Function

    ''' <summary>The monitor used for the patient: its hardware id (e.g. "DEL4106") and resolution.</summary>
    Public Shared ReadOnly Property ScreenId As String
        Get
            IdentifyScreen()
            Return _screenId
        End Get
    End Property

    ''' <summary>The monitor's name and resolution, for display.</summary>
    Public Shared ReadOnly Property ScreenName As String
        Get
            IdentifyScreen()
            Return _screenName
        End Get
    End Property

    ''' <summary>Looks the monitor up again (after the screens are changed in Settings).</summary>
    Public Shared Sub ForgetScreen()
        _screenId = Nothing
    End Sub

    Private Shared Sub IdentifyScreen()
        If _screenId IsNot Nothing Then Return
        Dim sc = ScreenLayout.ScreenOf(AppHost.Settings.StimulusBounds)
        Dim size = $"{sc.Bounds.Width}x{sc.Bounds.Height}"
        Dim hw = sc.DeviceName, model = sc.DeviceName.Replace("\\.\", "")
        Try
            Dim dd As New DISPLAY_DEVICE With {.cb = Marshal.SizeOf(GetType(DISPLAY_DEVICE))}
            If EnumDisplayDevices(sc.DeviceName, 0, dd, 0) Then
                ' DeviceID is like MONITOR\DEL4106\{...}\0001: the middle part names the model
                Dim parts = If(dd.DeviceID, "").Split("\"c)
                If parts.Length > 1 AndAlso parts(1) <> "" Then hw = parts(1)
                If Not String.IsNullOrWhiteSpace(dd.DeviceString) Then model = dd.DeviceString.Trim()
            End If
        Catch ex As Exception
            Log.Warn("Monitor name: " & ex.Message)
        End Try
        _screenId = hw & " " & size
        _screenName = $"{model} ({hw}), {sc.Bounds.Width} × {sc.Bounds.Height}"
    End Sub

#End Region

#Region "File"

    Private Shared ReadOnly Property FilePath As String
        Get
            Return IO.Path.Combine(AppSettings.SettingsFolder, "ReadEZ Colours.json")
        End Get
    End Property

    Private Shared ReadOnly SetNames As String() = {"Overlays", "Lenses", "StrongLenses"}

    Private Shared Sub EnsureLoaded()
        If _entries IsNot Nothing Then Return
        Dim entries = New Entry(2)() {}
        For s = 0 To 2
            entries(s) = New Entry(PerSet - 1) {}
            For i = 0 To PerSet - 1
                entries(s)(i) = New Entry With {.Colour = Defaults(s)(i)}
            Next
        Next
        Try
            If IO.File.Exists(FilePath) Then
                Dim root = TryCast(JsonNode.Parse(IO.File.ReadAllText(FilePath)), JsonObject)
                For s = 0 To 2
                    Dim arr = TryCast(root?(SetNames(s)), JsonArray)
                    If arr Is Nothing Then Continue For
                    For i = 0 To Math.Min(PerSet, arr.Count) - 1
                        Dim o = TryCast(arr(i), JsonObject)
                        If o Is Nothing Then Continue For
                        Dim hex = o("Rgb")?.GetValue(Of String)()
                        If hex IsNot Nothing AndAlso hex.Length = 7 AndAlso hex(0) = "#"c Then
                            entries(s)(i).Colour = Color.FromArgb(Integer.Parse(hex.Substring(1), NumberStyles.HexNumber, Inv) Or &HFF000000)
                        End If
                        Dim d = o("Checked")?.GetValue(Of String)()
                        Dim whenDone As Date
                        If d IsNot Nothing AndAlso Date.TryParseExact(d, "yyyy-MM-dd", Inv, DateTimeStyles.None, whenDone) Then entries(s)(i).CheckedOn = whenDone
                        entries(s)(i).Screen = If(o("Screen")?.GetValue(Of String)(), "")
                    Next
                Next
            End If
        Catch ex As Exception
            Log.Error("ReadEZ colours", ex)
        End Try
        _entries = entries
    End Sub

    Private Shared Sub Save()
        Try
            Dim root As New JsonObject()
            For s = 0 To 2
                Dim arr As New JsonArray()
                For i = 0 To PerSet - 1
                    Dim e = _entries(s)(i)
                    Dim o As New JsonObject()
                    o("Name") = Code(CType(s, ReadEZSet), i)
                    o("Rgb") = $"#{e.Colour.R:X2}{e.Colour.G:X2}{e.Colour.B:X2}"
                    If e.CheckedOn.HasValue Then
                        o("Checked") = e.CheckedOn.Value.ToString("yyyy-MM-dd", Inv)
                        o("Screen") = e.Screen
                    End If
                    arr.Add(o)
                Next
                root(SetNames(s)) = arr
            Next
            SettingsJson.WriteFileSafely(FilePath, root.ToJsonString(New JsonSerializerOptions With {.WriteIndented = True}))
        Catch ex As Exception
            Log.Error("Save ReadEZ colours", ex)
        End Try
    End Sub

#End Region
End Class

''' <summary>
''' Perceptually even colour space (OKLab, Björn Ottosson) for adjusting a colour in small, equal
''' looking steps: lighter / darker, more / less colour, and hue.
''' </summary>
Public Structure OkLab
    Public L, A, B As Double

    Public Sub New(lightness As Double, aAxis As Double, bAxis As Double)
        L = lightness : A = aAxis : B = bAxis
    End Sub

    Public ReadOnly Property Chroma As Double
        Get
            Return Math.Sqrt(A * A + B * B)
        End Get
    End Property

    Public Shared Function FromColor(c As Color) As OkLab
        Dim r = ToLinear(c.R), g = ToLinear(c.G), bl = ToLinear(c.B)
        Dim lc = Math.Cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * bl)
        Dim mc = Math.Cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * bl)
        Dim sc = Math.Cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * bl)
        Return New OkLab(0.2104542553 * lc + 0.793617785 * mc - 0.0040720468 * sc,
                         1.9779984951 * lc - 2.428592205 * mc + 0.4505937099 * sc,
                         0.0259040371 * lc + 0.7827717662 * mc - 0.808675766 * sc)
    End Function

    ''' <summary>The nearest screen colour (channels clipped to 0 - 255).</summary>
    Public Function ToColor() As Color
        ' (VB names aren't case-sensitive, so these can't be called l, m, s beside the fields L, A, B)
        Dim lc = Math.Pow(Me.L + 0.3963377774 * Me.A + 0.2158037573 * Me.B, 3)
        Dim mc = Math.Pow(Me.L - 0.1055613458 * Me.A - 0.0638541728 * Me.B, 3)
        Dim sc = Math.Pow(Me.L - 0.0894841775 * Me.A - 1.291485548 * Me.B, 3)
        Return Color.FromArgb(255,
                              ToByte(4.0767416621 * lc - 3.3077115913 * mc + 0.2309699292 * sc),
                              ToByte(-1.2684380046 * lc + 2.6097574011 * mc - 0.3413193965 * sc),
                              ToByte(-0.0041960863 * lc - 0.7034186147 * mc + 1.707614701 * sc))
    End Function

    ''' <summary>Distance between two colours (about 0.02 is just noticeable side by side).</summary>
    Public Shared Function Distance(x As OkLab, y As OkLab) As Double
        Return Math.Sqrt((x.L - y.L) ^ 2 + (x.A - y.A) ^ 2 + (x.B - y.B) ^ 2)
    End Function

    Private Shared Function ToLinear(v As Byte) As Double
        Dim c = v / 255.0
        Return If(c <= 0.04045, c / 12.92, ((c + 0.055) / 1.055) ^ 2.4)
    End Function

    Private Shared Function ToByte(lin As Double) As Integer
        lin = Math.Max(0, Math.Min(1, lin))
        Dim c = If(lin <= 0.0031308, 12.92 * lin, 1.055 * lin ^ (1 / 2.4) - 0.055)
        Return CInt(Math.Round(Math.Max(0, Math.Min(1, c)) * 255))
    End Function
End Structure
