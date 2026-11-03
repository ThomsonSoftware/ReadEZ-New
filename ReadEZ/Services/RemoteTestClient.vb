Option Strict On

Imports System.Globalization
Imports System.IO
Imports System.Net.Http
Imports System.Xml.Linq

' ====================================================================================================
'   RemoteTestClient.vb - the Remote Test "ImageSync" web service
'
'   Taken unchanged from Thomson Near Chart 3, so ReadEZ links phones and tablets exactly as the
'   Near Chart does (same server, same practice ID, same device links).
'
'   The previous version used a WCF service reference (Connected Services\ImageSync). .NET 8 has no
'   built-in WCF client, so the same SOAP 1.1 calls (BasicHttpBinding, document/literal) are made
'   directly with HttpClient. Operation names, parameter names and namespaces are exactly those of
'   ImageSync.wsdl, so the server sees the same messages as before.
'
'   UploadImage takes a System.Drawing.Bitmap. WCF sent it as an ISerializable object: a single
'   member "Data" holding the image file bytes (PNG). That XML is reproduced in UploadImageAsync.
'
'   Every call is asynchronous; the window never freezes while the server is contacted.
' ====================================================================================================

''' <summary>Raised when the server cannot be reached (no internet, server down).</summary>
Public Class RemoteServerUnavailableException
    Inherits Exception
    Public Sub New(inner As Exception)
        MyBase.New(If(inner?.Message, "Remote Test server unavailable"), inner)
    End Sub
End Class

''' <summary>Raised when the server replies with a SOAP fault.</summary>
Public Class RemoteServerFaultException
    Inherits Exception
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub
End Class

Public NotInheritable Class RemoteTestClient

    Private Shared ReadOnly Http As New HttpClient With {.Timeout = TimeSpan.FromSeconds(30)}
    Private Shared ReadOnly NsSoap As XNamespace = "http://schemas.xmlsoap.org/soap/envelope/"
    Private Shared ReadOnly NsTemp As XNamespace = "http://tempuri.org/"
    Private Shared ReadOnly NsDto As XNamespace = "http://schemas.datacontract.org/2004/07/Thomson.Database2"
    Private Shared ReadOnly NsXsi As XNamespace = "http://www.w3.org/2001/XMLSchema-instance"
    Private Shared ReadOnly NsXsd As XNamespace = "http://www.w3.org/2001/XMLSchema"
    Private Shared ReadOnly NsDrawing As XNamespace = "http://schemas.datacontract.org/2004/07/System.Drawing"
    Private Shared ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    Private ReadOnly _company As Guid

    ''' <summary>The URL that answered last; tried first next time.</summary>
    Private Shared _workingUrl As String = ""

    Public Sub New(companyId As Guid)
        _company = companyId
    End Sub

    Public ReadOnly Property CompanyId As Guid
        Get
            Return _company
        End Get
    End Property

#Region "Operations"

    Public Async Function ValidateCompanyAsync() As Task(Of Boolean)
        Dim r = Await CallAsync("ValidateCompany", P("companyId", _company)).ConfigureAwait(False)
        Return ParseBool(r)
    End Function

    Public Async Function DeviceGetAllAsync() As Task(Of List(Of DeviceInfo))
        Dim r = Await CallAsync("DeviceGetAll", P("companyId", _company)).ConfigureAwait(False)
        Dim list As New List(Of DeviceInfo)
        If r Is Nothing Then Return list
        For Each el In r.Elements(NsDto + "DeviceDTO")
            Dim d = ParseDevice(el)
            If d IsNot Nothing Then list.Add(d)
        Next
        Return list
    End Function

    Public Async Function DeviceSettingsAsync(deviceId As Guid) As Task(Of DeviceInfo)
        Return ParseDevice(Await CallAsync("DeviceSettings", P("deviceId", deviceId)).ConfigureAwait(False))
    End Function

    Public Async Function DeviceAddAsync(name As String) As Task(Of DeviceInfo)
        Return ParseDevice(Await CallAsync("DeviceAdd", P("companyId", _company), P("deviceName", name)).ConfigureAwait(False))
    End Function

    Public Async Function DeviceUpdateNameAsync(deviceId As Guid, name As String) As Task(Of DeviceInfo)
        Return ParseDevice(Await CallAsync("DeviceUpdateName", P("companyId", _company), P("deviceId", deviceId), P("deviceName", name)).ConfigureAwait(False))
    End Function

    Public Async Function DeviceDeleteAsync(deviceId As Guid) As Task(Of Boolean)
        Return ParseBool(Await CallAsync("DeviceDelete", P("companyId", _company), P("deviceId", deviceId)).ConfigureAwait(False))
    End Function

    Public Async Function DeviceClearCacheAsync(deviceId As Guid) As Task(Of Boolean)
        Return ParseBool(Await CallAsync("DeviceClearCache", P("companyId", _company), P("deviceId", deviceId)).ConfigureAwait(False))
    End Function

    ''' <summary>Any of the DeviceUpdate... calls that take a float (PtsPerCM, PixelsPerMM, ViewingDistanceNear/Far, ptsPerCMLogMAR).</summary>
    Public Async Function DeviceUpdateAsync(operation As String, deviceId As Guid, value As Single) As Task
        Await CallAsync(operation, P("companyId", _company), P("deviceId", deviceId), P("value", value)).ConfigureAwait(False)
    End Function

    ''' <summary>Any of the DeviceUpdate... calls that take a string (colours "r:g:b").</summary>
    Public Async Function DeviceUpdateAsync(operation As String, deviceId As Guid, value As String) As Task
        Await CallAsync(operation, P("companyId", _company), P("deviceId", deviceId), P("value", value)).ConfigureAwait(False)
    End Function

    ''' <summary>DeviceUpdateCalibrated / DeviceUpdateLinked.</summary>
    Public Async Function DeviceUpdateAsync(operation As String, deviceId As Guid, value As Boolean) As Task
        Await CallAsync(operation, P("companyId", _company), P("deviceId", deviceId), P("value", value)).ConfigureAwait(False)
    End Function

    Public Async Function DeviceTouchedValueAsync(deviceId As Guid) As Task(Of TouchInfo)
        Dim r = Await CallAsync("DeviceTouchedValue", P("companyId", _company), P("deviceId", deviceId)).ConfigureAwait(False)
        If r Is Nothing OrElse IsNil(r) Then Return Nothing
        Return New TouchInfo With {
            .X = ParseInt(r.Element(NsDto + "X")),
            .Y = ParseInt(r.Element(NsDto + "Y")),
            .Width = ParseInt(r.Element(NsDto + "Width")),
            .Height = ParseInt(r.Element(NsDto + "Height")),
            .FullScreen = ParseBool(r.Element(NsDto + "FullScreen"))}
    End Function

    ''' <summary>Uploads a picture for a device and returns the server's image id.</summary>
    Public Async Function UploadImageAsync(png As Byte(), deviceId As Guid) As Task(Of Guid)
        ' Same XML as WCF wrote for an ISerializable Bitmap: <image><Data i:type="b:base64Binary">...</Data></image>
        Dim image As New XElement(NsTemp + "image",
            New XAttribute(XNamespace.Xmlns + "a", NsDrawing),
            New XAttribute(XNamespace.Xmlns + "i", NsXsi),
            New XElement(XNamespace.None + "Data",
                New XAttribute(XNamespace.Xmlns + "b", NsXsd),
                New XAttribute(NsXsi + "type", "b:base64Binary"),
                Convert.ToBase64String(png)))
        Dim r = Await CallAsync("UploadImage", image, P("deviceId", deviceId)).ConfigureAwait(False)
        Dim g As Guid
        If r IsNot Nothing AndAlso Guid.TryParse(r.Value, g) Then Return g
        Return Guid.Empty
    End Function

    Public Async Function SelectImageAsync(imageId As Guid, deviceId As Guid) As Task(Of Boolean)
        Return ParseBool(Await CallAsync("SelectImage", P("imageGuid", imageId), P("deviceId", deviceId)).ConfigureAwait(False))
    End Function

    ''' <summary>Encodes a bitmap as PNG (what Bitmap.Save wrote for an in-memory bitmap before).</summary>
    Public Shared Function ToPng(bmp As Bitmap) As Byte()
        Using ms As New MemoryStream()
            bmp.Save(ms, Imaging.ImageFormat.Png)
            Return ms.ToArray()
        End Using
    End Function

#End Region

#Region "SOAP plumbing"

    Private Shared Function P(name As String, value As Object) As XElement
        Dim text As String
        Select Case True
            Case TypeOf value Is Guid : text = DirectCast(value, Guid).ToString()
            Case TypeOf value Is Boolean : text = If(DirectCast(value, Boolean), "true", "false")
            Case TypeOf value Is Single : text = DirectCast(value, Single).ToString("R", Inv)
            Case Else : text = Convert.ToString(value, Inv)
        End Select
        Return New XElement(NsTemp + name, text)
    End Function

    ''' <summary>The configured URL first, then the same address with the other scheme (http/https).</summary>
    Private Shared Function Endpoints() As List(Of String)
        Dim configured = If(AppHost.Settings?.ServerUrl, AppSettings.DefaultServerUrl)
        If String.IsNullOrWhiteSpace(configured) Then configured = AppSettings.DefaultServerUrl
        Dim other = If(configured.StartsWith("https:", StringComparison.OrdinalIgnoreCase),
                       "http:" & configured.Substring(6), "https:" & configured.Substring(configured.IndexOf(":"c) + 1))
        Dim list As New List(Of String)
        If _workingUrl <> "" AndAlso (_workingUrl = configured OrElse _workingUrl = other) Then list.Add(_workingUrl)
        For Each u In {configured, other}
            If Not list.Contains(u) Then list.Add(u)
        Next
        Return list
    End Function

    ''' <summary>Sends one operation and returns the &lt;operationResult&gt; element (Nothing if empty).</summary>
    Private Async Function CallAsync(operation As String, ParamArray args As XElement()) As Task(Of XElement)
        Dim envelope As New XDocument(
            New XElement(NsSoap + "Envelope",
                New XAttribute(XNamespace.Xmlns + "s", NsSoap),
                New XElement(NsSoap + "Body",
                    New XElement(NsTemp + operation, args))))
        Dim body = envelope.ToString(SaveOptions.DisableFormatting)

        Dim lastError As Exception = Nothing
        For Each url In Endpoints()
            Try
                Using req As New HttpRequestMessage(HttpMethod.Post, url)
                    req.Content = New StringContent(body, Text.Encoding.UTF8, "text/xml")
                    req.Headers.Add("SOAPAction", $"""http://tempuri.org/IImageSync/{operation}""")
                    Using resp = Await Http.SendAsync(req).ConfigureAwait(False)
                        Dim xml = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                        If Not resp.IsSuccessStatusCode AndAlso Not xml.Contains("Envelope") Then
                            lastError = New HttpRequestException($"HTTP {CInt(resp.StatusCode)} {resp.ReasonPhrase}")
                            Continue For
                        End If
                        Dim doc = XDocument.Parse(xml)
                        _workingUrl = url
                        Dim fault = doc.Descendants(NsSoap + "Fault").FirstOrDefault()
                        If fault IsNot Nothing Then
                            Dim msg = If(fault.Element("faultstring")?.Value, "SOAP fault")
                            Log.Warn($"Remote Test {operation}: fault {msg}")
                            Throw New RemoteServerFaultException(msg)
                        End If
                        Return doc.Descendants(NsTemp + (operation & "Result")).FirstOrDefault()
                    End Using
                End Using
            Catch ex As HttpRequestException
                lastError = ex
            Catch ex As TaskCanceledException
                lastError = ex
            Catch ex As Xml.XmlException
                lastError = ex      ' not a SOAP reply (e.g. a proxy or Wi-Fi sign-in page)
            End Try
        Next
        Log.Warn($"Remote Test {operation}: {lastError?.Message}")
        Throw New RemoteServerUnavailableException(lastError)
    End Function

#End Region

#Region "Parsing"

    Private Shared Function IsNil(el As XElement) As Boolean
        Return el IsNot Nothing AndAlso String.Equals(CStr(el.Attribute(NsXsi + "nil")), "true", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function ParseBool(el As XElement) As Boolean
        Return el IsNot Nothing AndAlso String.Equals(el.Value.Trim(), "true", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function ParseInt(el As XElement) As Integer
        Dim v As Integer
        If el IsNot Nothing AndAlso Integer.TryParse(el.Value, NumberStyles.Integer, Inv, v) Then Return v
        Return 0
    End Function

    Private Shared Function ParseSingle(el As XElement) As Single?
        If el Is Nothing OrElse IsNil(el) Then Return Nothing
        Dim v As Single
        If Single.TryParse(el.Value, NumberStyles.Float, Inv, v) Then Return v
        Return Nothing
    End Function

    Private Shared Function ParseString(el As XElement) As String
        If el Is Nothing OrElse IsNil(el) Then Return Nothing
        Return el.Value
    End Function

    Private Shared Function ParseDevice(el As XElement) As DeviceInfo
        If el Is Nothing OrElse IsNil(el) Then Return Nothing
        Dim v = Function(name As String) el.Element(NsDto + name)
        Dim d As New DeviceInfo With {
            .Name = If(ParseString(v("Name")), ""),
            .Width = ParseInt(v("Width")),
            .Height = ParseInt(v("Height")),
            .Calibrated = ParseBool(v("Calibrated")),
            .Linked = ParseBool(v("Linked")),
            .FullScreen = ParseBool(v("FullScreen")),
            .PixelsPerMM = ParseSingle(v("PixelsPerMM")),
            .PtsPerCM = ParseSingle(v("PtsPerCM")),
            .PtsPerCMLogMAR = ParseSingle(v("PtsPerCMLogMAR")),
            .ViewingDistanceNear = ParseSingle(v("ViewingDistanceNear")),
            .ViewingDistanceFar = ParseSingle(v("ViewingDistanceFar")),
            .RedColour = ParseString(v("RedColour")),
            .GreenColour = ParseString(v("GreenColour")),
            .BackgroundColour = ParseString(v("BackgroundColour")),
            .DuoChromeRed = ParseString(v("DuoChromeRed")),
            .DuoChromeGreen = ParseString(v("DuoChromeGreen"))}
        Dim g As Guid
        If Guid.TryParse(ParseString(v("Id")), g) Then d.Id = g
        Dim connected = v("ClientConnected")
        Dim dt As Date
        If connected IsNot Nothing AndAlso Not IsNil(connected) AndAlso
           Date.TryParse(connected.Value, Inv, DateTimeStyles.RoundtripKind, dt) Then d.ClientConnected = dt
        Return d
    End Function

#End Region

End Class
