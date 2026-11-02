Option Strict On

Imports System.Globalization
Imports System.IO
Imports System.Net.Http
Imports System.Runtime.Serialization
Imports System.Xml.Linq

' ====================================================================================================
'   LicenceService.vb - product licence (activation)
'
'   Replaces Licence.vb, frmOnLineActivation and frmOfflineActivation. The rules are unchanged:
'
'   ONLINE licence  (Thomson licence server, e-mail + 32-digit licence key)
'     - First use: Activate  -> the server returns an activation token, expiry date and the
'       number of days the program may run without checking in ("phone home" days).
'     - Every start: Check   -> if the server can't be reached the program still runs until the
'       phone-home days since the last successful check have been used up.
'     - Deactivate frees the licence for another computer.
'
'   OFFLINE licence (activation code issued by Thomson Software Solutions)
'     - The program shows a serial number (RZ-xxxxxxxxxxxxxx). Thomson returns a 16-character
'       activation code "AAAAAAAA-BBBBBBB": the first part is derived from the serial number and
'       the second part encodes the expiry date.
'     - Deactivate gives a de-activation code (VO-xxxxx9) to send to Thomson.
'     - A 3-day trial is available.
'
'   Storage: %AppData%\ReadEZ\Licence.json. ReadEZ 2020's tssRZ.dll (BinaryFormatter) is imported
'   once if found next to the program or in ReadEZ 2020's folder. As before the file records the
'   computer name; a file copied from another computer is refused. Same server, product token and
'   offline codes as ReadEZ 2020, so existing licences carry over.
' ====================================================================================================

#Region "Licence data"

''' <summary>The saved licence state. Field names match the previous LicenceStructure.</summary>
Public Class LicenceData
    ''' <summary>True = offline activation code, False = online licence server.</summary>
    Public LicenceMode As Boolean

    ' Offline
    Public Act_Registered As Boolean
    Public Act_Registration_Date As String = ""
    Public Demo_Count As Integer
    Public Demo_Date As Date
    Public Act_ActivationCode As String = ""
    Public Act_FirstUse As Boolean
    Public Act_SerialNo As String = ""

    ' Online
    Public OL_Registered As Boolean
    Public OL_LicenceCode As String = ""
    Public OL_ActivationClientToken As String = ""
    Public OL_Email As String = ""
    Public OL_ExpiryDate As Date
    Public OL_PhoneHomeDays As Integer
    Public OL_LastCheckDate As Date
    Public OL_LicencedTo As String = ""

    Public Hostname As String = ""

    ''' <summary>Checksum of all the values above, so the file can't be edited by hand.</summary>
    Public Signature As String = ""
End Class

#End Region

Public NotInheritable Class LicenceService
    Private Sub New()
    End Sub

    ''' <summary>ReadEZ's product on the Thomson licence server (as ReadEZ 2020).</summary>
    Public Const ProductToken As String = "a5fd2c72-12c2-4f0b-8301-b9759f829bb1"
    Public Const OfflineSerialPrefix As String = "RZ-"
    Public Const WarnDaysBeforeExpiry As Integer = 31
    Public Const TrialDays As Integer = 3

    Public Shared Property Data As New LicenceData
    ''' <summary>False if the licence file was made on another computer.</summary>
    Public Shared Property HostnameMatches As Boolean = True
    ''' <summary>Message from the licence server to show to the user (may be empty).</summary>
    Public Shared Property MarketingMessage As String = ""

    Private Shared ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    Public Shared ReadOnly Property FilePath As String
        Get
            Return Path.Combine(AppSettings.SettingsFolder, "Licence.json")
        End Get
    End Property

    ''' <summary>The previous version's licence file: next to this program, or in the previous install folder.</summary>
    Private Shared ReadOnly Property LegacyFilePath As String
        Get
            Dim folders As New List(Of String) From {AppHost.AppFolder}
            For Each f In LegacyImport.CandidateIniFiles()
                Dim parent = Path.GetDirectoryName(f)
                If parent IsNot Nothing Then folders.Add(parent)
            Next
            For Each f In folders
                Dim file = Path.Combine(f, "tssRZ.dll")
                If IO.File.Exists(file) Then Return file
            Next
            Return Path.Combine(AppHost.AppFolder, "tssRZ.dll")
        End Get
    End Property

#Region "Start-up"

    ''' <summary>
    ''' Start-up check, as D-Chart: the licence window is only shown when the program may not run
    ''' (not activated, activated on another computer, expired, or offline too long). Otherwise the
    ''' program starts straight away and the online licence is checked in the background once the
    ''' main window is showing (CheckOnlineInBackgroundAsync). Expiry within a month is shown on
    ''' the licence chip in the top bar rather than by stopping at start-up.
    ''' </summary>
    Public Shared Function EnsureLicensed() As Boolean
        Load()
        ' An activated online licence that simply hasn't been checked recently (the server often allows
        ' only 1 day between checks, so any gap of two days in use brought up the licence window even
        ' though the licence was valid) is checked again quietly while the splash shows. The licence
        ' window then only appears when there is no licence, it has expired, the server says it is no
        ' longer valid, or it can't be checked and the offline grace days have run out.
        If NeedsLicenceWindow() AndAlso CanRenewWithoutWindow() Then RenewQuietly()
        If NeedsLicenceWindow() Then
            SplashForm.CloseNow()           ' the licence window would be hidden behind it
            Using dlg As New LicenceDialog(LicenceDialogMode.StartUp)
                Return dlg.ShowDialog() = DialogResult.OK
            End Using
        End If
        Return True
    End Function

    ''' <summary>True for an activated online licence on this computer that has not reached its expiry date.</summary>
    Private Shared Function CanRenewWithoutWindow() As Boolean
        Dim d = Data
        If d Is Nothing OrElse Not HostnameMatches OrElse d.LicenceMode OrElse Not HasOnlineActivation Then Return False
        Return d.OL_ExpiryDate = Date.MinValue OrElse DaysUntil(d.OL_ExpiryDate) >= 0
    End Function

    ''' <summary>Checks the online licence with the server at start-up (at most 10 seconds).</summary>
    Private Shared Sub RenewQuietly()
        Try
            Dim check = Task.Run(Of Boolean)(Function() CheckOnlineInBackgroundAsync())
            If check.Wait(TimeSpan.FromSeconds(10)) Then
                Log.Info("Start-up licence check: " & If(check.Result, "valid", "not valid"))
            Else
                Log.Warn("Start-up licence check timed out")
            End If
        Catch ex As Exception
            Log.Warn("Start-up licence check: " & ex.Message)
        End Try
    End Sub

    ''' <summary>True when the program may not run without the licence window (from the saved file only).</summary>
    Public Shared Function NeedsLicenceWindow() As Boolean
        Dim d = Data
        If d Is Nothing OrElse Not HostnameMatches Then Return True
        If d.LicenceMode Then
            If Not d.Act_Registered OrElse Not OfflineCodeIsValid() Then Return True
            Dim exp = ExpiryFromCode(d.Act_ActivationCode)
            Return Not exp.HasValue OrElse DaysUntil(exp.Value) < 0
        End If
        If Not HasOnlineActivation Then Return True
        If d.OL_ExpiryDate <> Date.MinValue AndAlso DaysUntil(d.OL_ExpiryDate) < 0 Then Return True
        Return OfflineGraceDaysLeft <= 0
    End Function

    ''' <summary>Days until the licence expires, or Nothing when there is no activated licence with an expiry date.</summary>
    Public Shared Function DaysToExpiry() As Integer?
        Dim d = Data
        If d Is Nothing Then Return Nothing
        If d.LicenceMode Then
            If Not d.Act_Registered Then Return Nothing
            Dim exp = ExpiryFromCode(d.Act_ActivationCode)
            If Not exp.HasValue Then Return Nothing
            Return DaysUntil(exp.Value)
        End If
        If Not HasOnlineActivation OrElse d.OL_ExpiryDate = Date.MinValue Then Return Nothing
        Return DaysUntil(d.OL_ExpiryDate)
    End Function

    ''' <summary>The expiry date of the activated licence, if known.</summary>
    Public Shared Function ExpiryDate() As Date?
        Dim d = Data
        If d Is Nothing Then Return Nothing
        If d.LicenceMode Then Return If(d.Act_Registered, ExpiryFromCode(d.Act_ActivationCode), Nothing)
        If Not HasOnlineActivation OrElse d.OL_ExpiryDate = Date.MinValue Then Return Nothing
        Return d.OL_ExpiryDate
    End Function

    ''' <summary>
    ''' Checks an online licence with the server after start-up. False only when the server says
    ''' the licence is no longer valid; when the server cannot be reached the program carries on
    ''' within its offline grace days.
    ''' </summary>
    Public Shared Async Function CheckOnlineInBackgroundAsync() As Task(Of Boolean)
        Dim d = Data
        If d Is Nothing Then Return False
        If d.LicenceMode OrElse Not HasOnlineActivation Then Return True
        Dim key As Guid, token As Guid
        If Not Guid.TryParse(d.OL_LicenceCode, key) OrElse Not Guid.TryParse(d.OL_ActivationClientToken, token) Then Return False
        Try
            Dim r = Await LicenceApiClient.CheckAsync(key, token)
            If r.LicenceStatus = LicenceStatusType.Valid Then
                d.OL_LastCheckDate = Date.Now
                d.OL_LicencedTo = r.LicencedTo
                d.OL_ExpiryDate = r.Termination
                d.OL_PhoneHomeDays = r.NextCheckRequiredDays
                MarketingMessage = If(r.MarketingMessage, "")
                Save()
                Return True
            End If
            Log.Warn("Background licence check: licence status " & r.LicenceStatus.ToString())
            Return False
        Catch ex As LicenceServerUnavailableException
            Log.Warn("Background licence check: server unavailable, " & OfflineGraceDaysLeft & " day(s) of grace left")
            Return True
        Catch ex As Exception
            Log.Warn("Background licence check failed: " & ex.Message)
            Return True
        End Try
    End Function

#End Region

#Region "Load / save"

    Public Shared Sub Load()
        Data = Nothing
        Try
            If File.Exists(FilePath) Then
                Data = SettingsJson.Read(Of LicenceData)(File.ReadAllText(FilePath))
                If Not SignatureMatches(Data) Then
                    ' Edited by hand (or damaged): the activation must be confirmed again.
                    Log.Warn("Licence file checksum does not match - activation must be confirmed again")
                    Data.Act_Registered = False
                    Data.OL_LastCheckDate = Date.MinValue
                    Save()
                End If
            ElseIf File.Exists(LegacyFilePath) Then
                Data = ImportLegacy(LegacyFilePath)
                If Data IsNot Nothing Then
                    Log.Info("Imported licence from " & LegacyFilePath)
                    Save()
                End If
            End If
        Catch ex As Exception
            SettingsJson.SetAsideDamagedFile(FilePath, ex)
        End Try
        If Data Is Nothing Then Data = New LicenceData With {.LicenceMode = False}

        If String.IsNullOrEmpty(Data.Hostname) Then
            Data.Hostname = Environment.MachineName
        Else
            HostnameMatches = String.Equals(Data.Hostname, Environment.MachineName, StringComparison.OrdinalIgnoreCase)
        End If
    End Sub

    ''' <summary>Saves the licence. Returns False (and logs) if the file can't be written.</summary>
    Public Shared Function Save() As Boolean
        Try
            Data.Signature = ComputeSignature(Data)
            SettingsJson.WriteFileSafely(FilePath, SettingsJson.Write(Data))
            Return True
        Catch ex As Exception
            Log.Error("Licence save", ex)
            Return False
        End Try
    End Function

    ' The key ties the checksum to this program and this computer. It is not a secret from a
    ' determined attacker, but it stops the licence from being changed in a text editor.
    Private Shared Function ComputeSignature(d As LicenceData) As String
        Dim saved = d.Signature
        d.Signature = ""
        Try
            Dim json = SettingsJson.Write(d)
            Dim key = Text.Encoding.UTF8.GetBytes("ReadEZ|tssRZ|" & ProductToken & "|" & Environment.MachineName.ToUpperInvariant())
            Using h As New Security.Cryptography.HMACSHA256(key)
                Return Convert.ToBase64String(h.ComputeHash(Text.Encoding.UTF8.GetBytes(json)))
            End Using
        Finally
            d.Signature = saved
        End Try
    End Function

    Private Shared Function SignatureMatches(d As LicenceData) As Boolean
        Return Not String.IsNullOrEmpty(d.Signature) AndAlso d.Signature = ComputeSignature(d)
    End Function

#Disable Warning SYSLIB0011 ' read once from the previous version's file
    Private Shared Function ImportLegacy(file As String) As LicenceData
        Try
            Using fs = IO.File.OpenRead(file)
                If fs.Length = 0 Then Return Nothing
                Dim bf As New Formatters.Binary.BinaryFormatter With {.Binder = New LegacyLicenceBinder()}
                Dim o = bf.Deserialize(fs)
                If Not TypeOf o Is LegacyLicenceStructure Then Return Nothing
                Dim l = DirectCast(o, LegacyLicenceStructure)
                Dim d As New LicenceData With {
                    .LicenceMode = l.LicenceMode, .Act_Registered = l.Act_Registered,
                    .Act_Registration_Date = If(l.Act_Registration_Date, ""), .Demo_Count = l.Demo_Count,
                    .Act_ActivationCode = If(l.Act_ActivationCode, ""), .Act_FirstUse = l.Act_FirstUse,
                    .Act_SerialNo = If(l.Act_SerialNo, ""), .OL_Registered = l.OL_Registered,
                    .OL_LicenceCode = If(l.OL_LicenceCode, ""), .OL_ActivationClientToken = If(l.OL_ActivationClientToken, ""),
                    .OL_Email = If(l.OL_Email, ""), .OL_ExpiryDate = l.OL_ExpiryDate, .OL_PhoneHomeDays = l.OL_PhoneHomeDays,
                    .OL_LastCheckDate = l.OL_LastCheckDate, .OL_LicencedTo = If(l.OL_LicencedTo, ""), .Hostname = If(l.Hostname, "")}
                Dim demo As Date
                If Not String.IsNullOrEmpty(l.Demo_Date) AndAlso Date.TryParse(l.Demo_Date, CultureInfo.CurrentCulture, DateTimeStyles.None, demo) Then d.Demo_Date = demo
                Return d
            End Using
        Catch ex As Exception
            Log.Error("Licence import", ex)
            Return Nothing
        End Try
    End Function
#Enable Warning SYSLIB0011

    Private NotInheritable Class LegacyLicenceBinder
        Inherits SerializationBinder
        Public Overrides Function BindToType(assemblyName As String, typeName As String) As Type
            If typeName.EndsWith("LicenceStructure", StringComparison.Ordinal) Then Return GetType(LegacyLicenceStructure)
            Return Type.GetType($"{typeName}, {assemblyName}", throwOnError:=False)
        End Function
    End Class

#End Region

#Region "Status"

    ''' <summary>Whole days from now until the date (negative once it has passed), counted as before (DateDiff).</summary>
    Public Shared Function DaysUntil(d As Date) As Integer
        Return CInt(Math.Truncate((d - Date.Now).TotalDays))
    End Function

    Public Shared ReadOnly Property IsOffline As Boolean
        Get
            Return Data.LicenceMode
        End Get
    End Property

    Public Shared ReadOnly Property HasOnlineActivation As Boolean
        Get
            Return Data.OL_ActivationClientToken <> ""
        End Get
    End Property

    ''' <summary>Days the program may still run without reaching the licence server.</summary>
    Public Shared ReadOnly Property OfflineGraceDaysLeft As Integer
        Get
            Return Data.OL_PhoneHomeDays - CInt(Math.Floor((Date.Today - Data.OL_LastCheckDate.Date).TotalDays))
        End Get
    End Property

    ''' <summary>
    ''' The licence file was saved on another computer (copied, or this computer was renamed).
    ''' The activation isn't valid here, so it is cleared and the licence is moved to this computer:
    ''' the user can activate again instead of being locked out with no way forward.
    ''' (A new offline serial number is made, so old offline codes don't work; the old online
    ''' activation is dropped, so the file can't be copied to run on a second computer.)
    ''' </summary>
    Public Shared Sub ResetForThisComputer()
        Log.Warn($"Licence was saved on '{Data.Hostname}', this computer is '{Environment.MachineName}': activation cleared")
        Data.Act_Registered = False
        Data.Act_ActivationCode = ""
        Data.Act_FirstUse = True
        Data.OL_ActivationClientToken = ""
        Data.OL_LastCheckDate = Date.MinValue
        Data.Hostname = Environment.MachineName
        HostnameMatches = True
        Save()
    End Sub

    Public Shared Sub SwitchMode(offline As Boolean)
        Data.Act_Registered = False
        Data.OL_ActivationClientToken = ""
        Data.LicenceMode = offline
        Save()
    End Sub

#End Region

#Region "Offline activation codes"

    ''' <summary>The serial number shown to the user, creating one on first use.</summary>
    Public Shared Function OfflineSerial() As String
        If Data.Act_SerialNo.Length < 8 Then Data.Act_FirstUse = True : Data.Act_Registration_Date = ""
        If Data.Act_FirstUse Then
            Data.Demo_Count = 5
            Data.Act_Registration_Date = Date.Today.ToString("yyyy-MM-dd", Inv)
            Data.Act_SerialNo = Date.Now.ToString("ss4mm5hh6yy2mm3dd", Inv)   ' same pattern as before
            Data.Act_FirstUse = False
            Save()
        End If
        Return OfflineSerialPrefix & Data.Act_SerialNo
    End Function

    ''' <summary>First half of a valid activation code for this computer's serial number.</summary>
    Private Shared Function Serial1() As String
        Dim sb As New Text.StringBuilder
        For i = 9 To 2 Step -1
            Dim digit = If(i <= Data.Act_SerialNo.Length AndAlso Char.IsDigit(Data.Act_SerialNo(i - 1)), AscW(Data.Act_SerialNo(i - 1)) - 48, 0)
            sb.Append(ChrW(digit + 68))
        Next
        Return sb.ToString()
    End Function

    ''' <summary>True if the saved activation code was issued for this computer's serial number.</summary>
    Public Shared Function OfflineCodeIsValid() As Boolean
        Dim code = If(Data.Act_ActivationCode, "")
        Return code.Length >= 16 AndAlso code.StartsWith(Serial1(), StringComparison.Ordinal)
    End Function

    ''' <summary>Decodes the expiry date from the second half of an activation code (7 characters).</summary>
    Public Shared Function ExpiryFromCode(code As String) As Date?
        If code Is Nothing OrElse code.Length < 7 Then Return Nothing
        Dim s = code.Substring(code.Length - 7)
        Dim d = Function(i As Integer) AscW(s(i - 1)) - 68
        Try
            Dim yy = d(1) * 10 + d(2)
            Dim dd = d(3) * 10 + d(4)
            Dim mm = d(6) * 10 + d(7)
            Return New Date(2000 + yy, mm, dd)
        Catch
            Return Nothing
        End Try
    End Function

    Public Enum OfflineResult
        Activated
        WrongCode
        Expired
        SaveFailed
    End Enum

    Public Shared Function ActivateOffline(code As String, ByRef expiry As Date) As OfflineResult
        code = If(code, "").Trim()
        If code.Length < 16 OrElse Not code.StartsWith(Serial1(), StringComparison.Ordinal) Then Return OfflineResult.WrongCode
        Dim exp = ExpiryFromCode(code)
        If Not exp.HasValue Then Return OfflineResult.WrongCode
        expiry = exp.Value
        If exp.Value <= Date.Now Then Return OfflineResult.Expired
        Data.Act_Registered = True
        Data.Act_ActivationCode = Serial1() & "-" & code.Substring(code.Length - 7)
        Return If(Save(), OfflineResult.Activated, OfflineResult.SaveFailed)
    End Function

    ''' <summary>Removes the offline activation and returns the de-activation code (Nothing if it could not be saved).</summary>
    Public Shared Function DeactivateOffline() As String
        Dim old = (Data.Act_ActivationCode, Data.Act_FirstUse, Data.Act_Registered)
        Data.Act_ActivationCode = ""
        Data.Act_FirstUse = True
        Data.Act_Registered = False
        Dim serial = Data.Act_SerialNo
        If Save() Then Return "VO-" & serial.Substring(0, Math.Min(5, serial.Length)) & "9"
        Data.Act_ActivationCode = old.Item1 : Data.Act_FirstUse = old.Item2 : Data.Act_Registered = old.Item3
        Return Nothing
    End Function

    ''' <summary>Builds a 12-month code for this computer (Shift + right-click on the licence window header, as before).</summary>
    Public Shared Function GenerateServiceCode() As String
        Dim expDigits = Date.Now.AddMonths(12).ToString("yydd0MM", Inv)
        Dim sb As New Text.StringBuilder
        For Each ch In expDigits
            sb.Append(ChrW(AscW(ch) - 48 + 68))
        Next
        Return Serial1() & "-" & sb.ToString()
    End Function

    ''' <summary>Starts or continues the trial. Returns the days left (0 or less = expired).</summary>
    Public Shared Function UseTrial() As Integer
        If Data.Demo_Date = Date.MinValue OrElse Data.Demo_Count < 1 Then
            Data.Demo_Date = Date.Today
            Data.Demo_Count = 1
        Else
            Data.Demo_Count += 1
        End If
        Save()
        Dim used = CInt(Math.Floor((Date.Today - Data.Demo_Date.Date).TotalDays))
        If used < 0 Then Return 0
        Return TrialDays - used
    End Function

#End Region

End Class

#Region "Licence server (SOAP)"

Public Enum LicenceStatusType
    Valid = 0
    Invalid = 1
    Activated = 2
    Deactivated = 3
    Expired = 4
End Enum

Public Class ActivationStatus
    Public Property LicenceStatus As LicenceStatusType = LicenceStatusType.Invalid
    Public Property Message As String = ""
    Public Property ActivationToken As Guid
    Public Property Termination As Date
    Public Property NextCheckRequiredDays As Integer
    Public Property LicencedTo As String = ""
    Public Property MarketingMessage As String = ""
End Class

''' <summary>Raised when the licence server cannot be reached (no internet, server down).</summary>
Public Class LicenceServerUnavailableException
    Inherits Exception
    Public Sub New(inner As Exception)
        MyBase.New("Licence server unavailable", inner)
    End Sub
End Class

''' <summary>
''' Calls the Thomson licence server. The previous version used a WCF service reference; .NET 8
''' has no built-in WCF client, so the three SOAP calls are made directly with HttpClient.
''' </summary>
Public NotInheritable Class LicenceApiClient
    Private Sub New()
    End Sub

    Private Shared ReadOnly Endpoints As String() = {
        "https://licensing.thomson-cloud.com/v/1/licenceAPI.svc",
        "http://licensing.thomson-cloud.com/v/1/licenceAPI.svc"}

    Private Shared ReadOnly Http As New HttpClient With {.Timeout = TimeSpan.FromSeconds(20)}
    Private Shared ReadOnly NsSoap As XNamespace = "http://schemas.xmlsoap.org/soap/envelope/"
    Private Shared ReadOnly NsTemp As XNamespace = "http://tempuri.org/"
    Private Shared ReadOnly NsDto As XNamespace = "http://schemas.datacontract.org/2004/07/Thomson.licencing"

    Public Shared Function ActivateAsync(email As String, licenceKey As Guid) As Task(Of ActivationStatus)
        Return CallAsync("ActivateLicence",
                         New XElement(NsTemp + "productToken", LicenceService.ProductToken),
                         New XElement(NsTemp + "email", email),
                         New XElement(NsTemp + "licenceToken", licenceKey.ToString()))
    End Function

    Public Shared Function DeactivateAsync(licenceKey As Guid, activationToken As Guid) As Task(Of ActivationStatus)
        Return CallAsync("DeactivateLicence",
                         New XElement(NsTemp + "productToken", LicenceService.ProductToken),
                         New XElement(NsTemp + "licenceToken", licenceKey.ToString()),
                         New XElement(NsTemp + "activationClientToken", activationToken.ToString()))
    End Function

    Public Shared Function CheckAsync(licenceKey As Guid, activationToken As Guid) As Task(Of ActivationStatus)
        Return CallAsync("LicenceCheck",
                         New XElement(NsTemp + "productToken", LicenceService.ProductToken),
                         New XElement(NsTemp + "licenceToken", licenceKey.ToString()),
                         New XElement(NsTemp + "activationClientToken", activationToken.ToString()))
    End Function

    Private Shared Async Function CallAsync(operation As String, ParamArray args As XElement()) As Task(Of ActivationStatus)
        Dim envelope As New XDocument(
            New XElement(NsSoap + "Envelope",
                New XAttribute(XNamespace.Xmlns + "s", NsSoap),
                New XElement(NsSoap + "Body",
                    New XElement(NsTemp + operation, args))))
        Dim body = envelope.ToString(SaveOptions.DisableFormatting)

        Dim lastError As Exception = Nothing
        For Each url In Endpoints
            Try
                Using req As New HttpRequestMessage(HttpMethod.Post, url)
                    req.Content = New StringContent(body, Text.Encoding.UTF8, "text/xml")
                    req.Headers.Add("SOAPAction", $"""http://tempuri.org/ILicence/{operation}""")
                    Using resp = Await Http.SendAsync(req).ConfigureAwait(False)
                        Dim xml = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                        If Not resp.IsSuccessStatusCode AndAlso Not xml.Contains("Envelope") Then
                            lastError = New HttpRequestException($"HTTP {CInt(resp.StatusCode)}")
                            Continue For
                        End If
                        Return Parse(XDocument.Parse(xml), operation)
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
        Log.Warn($"Licence server {operation}: {lastError?.Message}")
        Throw New LicenceServerUnavailableException(lastError)
    End Function

    Private Shared Function Parse(doc As XDocument, operation As String) As ActivationStatus
        ' A SOAP fault or a reply without a result is a server problem, not an answer about the
        ' licence. Reporting it as "invalid" deactivated every clinic during a server outage even
        ' though they still had grace days; "unavailable" lets the grace days apply.
        Dim fault = doc.Descendants(NsSoap + "Fault").FirstOrDefault()
        If fault IsNot Nothing Then
            Dim msg = If(fault.Element("faultstring")?.Value, "SOAP fault")
            Log.Warn($"Licence server {operation} fault: {msg}")
            Throw New LicenceServerUnavailableException(New InvalidOperationException(msg))
        End If
        Dim result = doc.Descendants(NsTemp + (operation & "Result")).FirstOrDefault()
        If result Is Nothing Then
            Throw New LicenceServerUnavailableException(New InvalidOperationException($"No {operation}Result in the reply"))
        End If
        Dim v = Function(name As String) result.Element(NsDto + name)?.Value
        Dim s As New ActivationStatus With {.Message = If(v("Message"), ""), .LicencedTo = If(v("LicencedTo"), ""),
                                            .MarketingMessage = If(v("MarketingMessage"), "")}
        Dim status As LicenceStatusType
        If [Enum].TryParse(v("LicenceStatus"), True, status) Then s.LicenceStatus = status
        Dim g As Guid
        If Guid.TryParse(v("ActivationToken"), g) Then s.ActivationToken = g
        Dim d As Date
        If Date.TryParse(v("Termination"), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, d) Then s.Termination = d
        Dim n As Integer
        If Integer.TryParse(v("NextCheckRequiredDays"), NumberStyles.Integer, CultureInfo.InvariantCulture, n) Then s.NextCheckRequiredDays = n
        Return s
    End Function
End Class

#End Region

#Region "Previous version's licence structure (field names and types must match exactly)"

<Serializable>
Friend Structure LegacyLicenceStructure
    Public LicenceMode As Boolean
    Public Act_Registered As Boolean
    Public Act_Registration_Date As String
    Public Demo_Count As Integer
    Public Demo_Date As String
    Public Act_ActivationCode As String
    Public Act_FirstUse As Boolean
    Public Act_SerialNo As String
    Public OL_Registered As Boolean
    Public OL_ProductToken As String
    Public OL_LicenceCode As String
    Public OL_ActivationClientToken As String
    Public OL_Email As String
    Public OL_ExpiryDate As Date
    Public OL_PhoneHomeDays As Integer
    Public OL_LastCheckDate As Date
    Public OL_LicencedTo As String
    Public Hostname As String
End Structure

#End Region
