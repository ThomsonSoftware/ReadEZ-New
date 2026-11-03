Option Strict On

Imports System.Globalization

' ====================================================================================================
'   DeviceModels.vb - a phone or tablet registered on the Remote Test server (from Near Chart 3)
'     DeviceInfo   the device: id, name, size in pixels, whether it has connected
'     TouchInfo    the last place the screen was touched (and the screen's size then)
' ====================================================================================================

Public NotInheritable Class DeviceInfo
    Public Property Id As Guid
    Public Property Name As String = ""
    Public Property Width As Integer
    Public Property Height As Integer
    Public Property Calibrated As Boolean
    Public Property Linked As Boolean
    Public Property FullScreen As Boolean
    Public Property ClientConnected As Date?
    Public Property PixelsPerMM As Single?
    Public Property PtsPerCM As Single?
    Public Property PtsPerCMLogMAR As Single?
    Public Property ViewingDistanceNear As Single?
    Public Property ViewingDistanceFar As Single?
    Public Property RedColour As String
    Public Property GreenColour As String
    Public Property BackgroundColour As String
    Public Property DuoChromeRed As String
    Public Property DuoChromeGreen As String

    ''' <summary>True once the phone/tablet has opened the link at least once (it has reported its size).</summary>
    Public ReadOnly Property HasConnected As Boolean
        Get
            Return ClientConnected.HasValue AndAlso Width > 0 AndAlso Height > 0
        End Get
    End Property

    Public ReadOnly Property IsLandscape As Boolean
        Get
            Return Width > Height
        End Get
    End Property

    ''' <summary>Same rule as before: more than a million pixels is a tablet.</summary>
    Public ReadOnly Property IsTablet As Boolean
        Get
            Return CLng(Width) * Height > 1000000
        End Get
    End Property

    ' Calibration values with the same defaults the previous version fell back on
    Public ReadOnly Property PxPerMM As Single
        Get
            Return If(PixelsPerMM.HasValue AndAlso PixelsPerMM.Value > 0, PixelsPerMM.Value, 10.0F)
        End Get
    End Property

    Public ReadOnly Property PtsCM As Single
        Get
            Return If(PtsPerCM.HasValue AndAlso PtsPerCM.Value > 0, PtsPerCM.Value, 92.0F)
        End Get
    End Property

    Public ReadOnly Property PtsCMLogMAR As Single
        Get
            Return If(PtsPerCMLogMAR.HasValue AndAlso PtsPerCMLogMAR.Value > 0, PtsPerCMLogMAR.Value, 71.0F)
        End Get
    End Property

    Public ReadOnly Property NearCm As Single
        Get
            Return If(ViewingDistanceNear.HasValue AndAlso ViewingDistanceNear.Value > 0, ViewingDistanceNear.Value, 40.0F)
        End Get
    End Property

    Public ReadOnly Property FarCm As Single
        Get
            Return If(ViewingDistanceFar.HasValue AndAlso ViewingDistanceFar.Value > 0, ViewingDistanceFar.Value, 300.0F)
        End Get
    End Property

    ''' <summary>Reads "r:g:b" as stored on the server.</summary>
    Public Shared Function ParseColour(value As String, fallback As Color) As Color
        If String.IsNullOrWhiteSpace(value) Then Return fallback
        Dim p = value.Split(":"c)
        If p.Length < 3 Then Return fallback
        Dim r, g, b As Integer
        If Integer.TryParse(p(0), NumberStyles.Integer, CultureInfo.InvariantCulture, r) AndAlso
           Integer.TryParse(p(1), NumberStyles.Integer, CultureInfo.InvariantCulture, g) AndAlso
           Integer.TryParse(p(2), NumberStyles.Integer, CultureInfo.InvariantCulture, b) Then
            Return Color.FromArgb(Clamp(r), Clamp(g), Clamp(b))
        End If
        Return fallback
    End Function

    Public Shared Function FormatColour(c As Color) As String
        Return $"{c.R}:{c.G}:{c.B}"
    End Function

    Private Shared Function Clamp(v As Integer) As Integer
        Return Math.Max(0, Math.Min(255, v))
    End Function
End Class

Public NotInheritable Class TouchInfo
    Public Property X As Integer
    Public Property Y As Integer
    Public Property Width As Integer
    Public Property Height As Integer
    Public Property FullScreen As Boolean
End Class
