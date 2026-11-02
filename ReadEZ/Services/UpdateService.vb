Option Strict On

Imports System.IO
Imports System.Net.Http
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Threading

' ====================================================================================================
'   UpdateService.vb - checks the web for a new version and installs it
'
'   The server holds two files in one folder (written by Admin > Publish update):
'     update.json                      what the newest version is
'     ReadEZ-Setup-<version>.msi   its installer (signed)
'
'   update.json:
'     { "product": "ReadEZ", "version": "3.0.1", "released": "2026-10-27",
'       "file": "ReadEZ-Setup-3.0.1.msi", "size": 123456789,
'       "sha256": "<hex>", "notes": "What has changed" }
'   "file" is relative to update.json (an absolute "url" may be given instead).
'
'   Installing: the installer is downloaded to %TEMP%, its size, SHA-256 and signature are checked
'   (it must be intact and signed with the same certificate as the running program, when that is
'   signed - Thomson's certificate is self-signed, so Windows' own trust is not required), then
'   a small script waits for ReadEZ to close, runs the installer (Windows asks for permission)
'   and starts ReadEZ again. The installer is a Windows Installer upgrade: it replaces only
'   the program folder. Settings (%AppData%\ReadEZ), the licence, the databases and the
'   stimuli in the data folder are never touched; new starter stimuli are copied in only where a
'   file is missing (StandardFiles).
' ====================================================================================================

Public NotInheritable Class UpdateInfo
    Public Property Version As Version
    Public Property Released As String = ""
    Public Property FileName As String = ""
    Public Property DownloadUrl As String = ""
    Public Property Size As Long
    Public Property Sha256 As String = ""
    Public Property Notes As String = ""

    Public ReadOnly Property VersionText As String
        Get
            Return UpdateService.VersionText(Version)
        End Get
    End Property
End Class

Public NotInheritable Class UpdateService
    Private Sub New()
    End Sub

    ''' <summary>
    ''' The address every installed copy checks. ADJUST before the first release: the folder on the
    ''' web server that update.json and the installers are uploaded to. Admin > Publish update can
    ''' override it on one computer (for testing).
    ''' </summary>
    Public Const DefaultUpdateUrl As String = "https://www.thomson-software-solutions.com/updates/readez/update.json"

    Public Const ManifestFileName As String = "update.json"

    Private Shared ReadOnly Http As New HttpClient With {.Timeout = TimeSpan.FromMinutes(30)}

    Public Shared ReadOnly Property UpdateUrl As String
        Get
            Dim u = AppHost.Admin?.UpdateUrl
            Return If(String.IsNullOrWhiteSpace(u), DefaultUpdateUrl, u.Trim())
        End Get
    End Property

    ''' <summary>Version of the running program (from &lt;Version&gt; in ReadEZ.vbproj).</summary>
    Public Shared ReadOnly Property CurrentVersion As Version
        Get
            Dim a = If(Assembly.GetEntryAssembly(), GetType(UpdateService).Assembly)
            Return Normalise(a.GetName().Version)
        End Get
    End Property

    ''' <summary>Version with all four parts set (3.0.1 and 3.0.1.0 compare equal).</summary>
    Public Shared Function Normalise(v As Version) As Version
        If v Is Nothing Then Return New Version(0, 0, 0, 0)
        Return New Version(Math.Max(0, v.Major), Math.Max(0, v.Minor), Math.Max(0, v.Build), Math.Max(0, v.Revision))
    End Function

    ''' <summary>"3.0.1" (the fourth part only if it isn't 0).</summary>
    Public Shared Function VersionText(v As Version) As String
        Dim n = Normalise(v)
        Return If(n.Revision = 0, $"{n.Major}.{n.Minor}.{n.Build}", n.ToString())
    End Function

    Public Shared Function ParseVersion(text As String) As Version
        Dim v As Version = Nothing
        If Version.TryParse(If(text, "").Trim().TrimStart("v"c, "V"c), v) Then Return Normalise(v)
        Return Nothing
    End Function

#Region "Check"

    ''' <summary>Reads update.json. Throws if it can't be read.</summary>
    Public Shared Async Function ReadManifestAsync(Optional url As String = Nothing, Optional ct As CancellationToken = Nothing) As Task(Of UpdateInfo)
        url = If(url, UpdateUrl)
        ' No cached copy, so a new upload is seen straight away
        Dim requestUrl = url & If(url.Contains("?"), "&", "?") & "t=" & Date.UtcNow.Ticks
        Dim json As String
        Using cts = CancellationTokenSource.CreateLinkedTokenSource(ct)
            cts.CancelAfter(TimeSpan.FromSeconds(20))
            json = Await Http.GetStringAsync(requestUrl, cts.Token)
        End Using
        Return ParseManifest(json, url)
    End Function

    Public Shared Function ParseManifest(json As String, manifestUrl As String) As UpdateInfo
        Using doc = JsonDocument.Parse(json)
            Dim r = doc.RootElement
            Dim info As New UpdateInfo With {
                .Version = ParseVersion(JStr(r, "version")),
                .Released = JStr(r, "released"),
                .FileName = JStr(r, "file"),
                .Sha256 = JStr(r, "sha256").Replace(" ", "").ToUpperInvariant(),
                .Notes = JStr(r, "notes")}
            Dim size As JsonElement
            If r.TryGetProperty("size", size) AndAlso size.ValueKind = JsonValueKind.Number Then info.Size = size.GetInt64()
            If info.Version Is Nothing Then Throw New InvalidDataException("update.json has no valid ""version""")
            Dim direct = JStr(r, "url")
            If direct <> "" Then
                info.DownloadUrl = direct
            ElseIf info.FileName <> "" Then
                info.DownloadUrl = New Uri(New Uri(manifestUrl), Uri.EscapeDataString(info.FileName)).ToString()
            Else
                Throw New InvalidDataException("update.json has no ""file""")
            End If
            If info.FileName = "" Then info.FileName = Path.GetFileName(New Uri(info.DownloadUrl).LocalPath)
            Return info
        End Using
    End Function

    Private Shared Function JStr(e As JsonElement, name As String) As String
        Dim v As JsonElement
        If e.TryGetProperty(name, v) AndAlso v.ValueKind = JsonValueKind.String Then Return If(v.GetString(), "")
        Return ""
    End Function

    ''' <summary>The newest version on the server if it is newer than this one, otherwise Nothing. Throws if the server can't be read.</summary>
    Public Shared Async Function CheckAsync(Optional ct As CancellationToken = Nothing) As Task(Of UpdateInfo)
        Dim info = Await ReadManifestAsync(Nothing, ct)
        Log.Info($"Update check: server has {info.VersionText}, this is {VersionText(CurrentVersion)}")
        Return If(info.Version > CurrentVersion, info, Nothing)
    End Function

#End Region

#Region "Download and install"

    Public Shared ReadOnly Property DownloadFolder As String
        Get
            Return Path.Combine(Path.GetTempPath(), "ReadEZ update")
        End Get
    End Property

    ''' <summary>
    ''' Downloads the installer and checks it. Returns its path. progress receives 0-100.
    ''' Throws with a message for the user if the file is not the one described in update.json.
    ''' </summary>
    Public Shared Async Function DownloadAsync(info As UpdateInfo, progress As IProgress(Of Integer), ct As CancellationToken) As Task(Of String)
        Directory.CreateDirectory(DownloadFolder)
        Dim target = Path.Combine(DownloadFolder, Path.GetFileName(info.FileName))
        Dim partFile = target & ".part"
        Using response = Await Http.GetAsync(info.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, ct)
            response.EnsureSuccessStatusCode()
            Dim total = If(response.Content.Headers.ContentLength, info.Size)
            Using src = Await response.Content.ReadAsStreamAsync(ct), dst = File.Create(partFile)
                Dim buffer(81919) As Byte
                Dim done As Long = 0
                Dim lastPercent = -1
                Do
                    Dim n = Await src.ReadAsync(buffer.AsMemory(0, buffer.Length), ct)
                    If n = 0 Then Exit Do
                    Await dst.WriteAsync(buffer.AsMemory(0, n), ct)
                    done += n
                    If total > 0 Then
                        Dim pc = CInt(Math.Min(100, done * 100 \ total))
                        If pc <> lastPercent Then
                            lastPercent = pc
                            progress?.Report(pc)
                        End If
                    End If
                Loop
            End Using
        End Using
        Verify(partFile, info)
        If File.Exists(target) Then File.Delete(target)
        File.Move(partFile, target)
        Return target
    End Function

    ''' <summary>Size, SHA-256 and signature of a downloaded installer.</summary>
    Private Shared Sub Verify(file As String, info As UpdateInfo)
        Dim length = New FileInfo(file).Length
        If info.Size > 0 AndAlso length <> info.Size Then
            Throw New InvalidDataException(Loc.F("update.badSize", length, info.Size))
        End If
        If info.Sha256 <> "" Then
            Dim hash = FileSha256(file)
            If hash <> info.Sha256 Then Throw New InvalidDataException(Loc.T("update.badHash"))
        End If
        Dim program = CodeSigning.Program
        Dim installer = CodeSigning.Check(file)
        Log.Info("Update installer: " & installer.Describe())
        If program.IsSigned Then
            ' A released copy only runs an installer that is intact and signed with the same
            ' certificate as itself (Thomson's self-signed certificate is not trusted by Windows,
            ' so "valid" in Windows' sense is not required)
            If Not installer.IsIntact OrElse Not CodeSigning.SameCertificate(program, installer) Then
                Throw New InvalidDataException(Loc.T("update.badSignature"))
            End If
        ElseIf Not installer.IsIntact Then
            Log.Warn("Update installer is not signed (accepted because this build is not signed either)")
        End If
    End Sub

    Public Shared Function FileSha256(file As String) As String
        Using s = IO.File.OpenRead(file)
            Return Convert.ToHexString(SHA256.HashData(s))
        End Using
    End Function

    ''' <summary>
    ''' Starts the script that waits for ReadEZ to close, runs the installer and starts the
    ''' program again. The caller then closes the main window.
    ''' </summary>
    Public Shared Sub StartInstall(msiFile As String)
        Dim pid = Environment.ProcessId
        Dim exe = CodeSigning.ProgramFile
        Dim logFile = Path.Combine(DownloadFolder, "install.log")
        Dim script = Path.Combine(DownloadFolder, "install-update.cmd")
        Dim sb As New StringBuilder
        sb.AppendLine("@echo off")
        sb.AppendLine("rem ReadEZ update: wait for the program to close, install, start it again")
        sb.AppendLine(":wait")
        sb.AppendLine($"tasklist /FI ""PID eq {pid}"" 2>NUL | find ""{pid}"" >NUL")
        sb.AppendLine("if not errorlevel 1 (")
        sb.AppendLine("  ping -n 2 127.0.0.1 >NUL")
        sb.AppendLine("  goto wait")
        sb.AppendLine(")")
        sb.AppendLine($"start """" /wait msiexec.exe /i ""{msiFile}"" /passive /norestart /l*v ""{logFile}""")
        ' Starts the installed program (the installer keeps the install folder; a Visual Studio
        ' build starts itself again).
        ' Start the program again
        sb.AppendLine($"if exist ""{exe}"" start """" ""{exe}""")
        File.WriteAllText(script, sb.ToString(), Encoding.Default)
        Log.Info($"Installing update {msiFile} (script {script})")
        Process.Start(New ProcessStartInfo("cmd.exe", $"/c ""{script}""") With {
            .UseShellExecute = False, .CreateNoWindow = True, .WindowStyle = ProcessWindowStyle.Hidden,
            .WorkingDirectory = DownloadFolder})
    End Sub

#End Region

End Class
