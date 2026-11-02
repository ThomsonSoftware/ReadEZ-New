Option Strict On

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

' ====================================================================================================
'   AdminSettings.vb
'   Administrator settings. (new in this version).
'
'   Holds:
'     - which colour tests are offered (overlays, lenses, computer)
'     - the update and online help addresses and the settings used to publish an update
'     - the admin password (stored as a salted hash, never as text)
'
'   Stored in %AppData%\ReadEZ\Admin.json.
' ====================================================================================================

Public NotInheritable Class AdminSettings

#Region "Tests"
    ''' <summary>Which colour tests are offered on the Tests screen.</summary>
    Public OverlayTest_Enabled As Boolean = True
    Public LensTest_Enabled As Boolean = True
    Public ComputerTest_Enabled As Boolean = True

    Public Function IsTestEnabled(mode As ReadEZTestMode) As Boolean
        Select Case mode
            Case ReadEZTestMode.Overlay : Return OverlayTest_Enabled
            Case ReadEZTestMode.Lenses : Return LensTest_Enabled
            Case ReadEZTestMode.Computer : Return ComputerTest_Enabled
            Case Else : Return True
        End Select
    End Function
#End Region

#Region "Updates"

    ''' <summary>Address of update.json; blank = UpdateService.DefaultUpdateUrl.</summary>
    Public UpdateUrl As String = ""

    ''' <summary>Web address of the online help (built with HelpNDoc); Help > Online help opens it. Blank = not set.</summary>
    Public OnlineHelpUrl As String = ""

    ' Publishing (only used on the computer that builds the releases)
    ''' <summary>Folder containing ReadEZ.sln.</summary>
    Public Publish_SolutionFolder As String = ""
    ''' <summary>Folder the installer and update.json are written to, ready to upload.</summary>
    Public Publish_OutputFolder As String = ""
    ''' <summary>Code-signing certificate (.pfx).</summary>
    Public Publish_PfxFile As String = ""
    ''' <summary>Its password, encrypted for this Windows user (DPAPI); never stored as text.</summary>
    Public Publish_PfxPassword As String = ""
    Public Publish_SignTool As String = ""
    Public Publish_TimestampUrl As String = "http://timestamp.digicert.com"
    ''' <summary>ReadEZ database put in the installer as the starter database for new installations ("" = none).</summary>
    Public Publish_StarterDatabase As String = ""

#End Region

#Region "Password"

    ''' <summary>Salted SHA-256 of the admin password. Empty = the original password.</summary>
    Public PasswordHash As String = ""
    Public PasswordSalt As String = ""

    ' Until an administrator sets a new password (Admin > Change password), the standard Thomson
    ' Software Solutions administrator password is accepted.
    Private Const LegacySalt As String = "ReadEZ.Admin.v1"

    Public Function CheckPassword(entered As String) As Boolean
        If String.IsNullOrEmpty(entered) Then Return False
        If PasswordHash = "" Then
            Return Hash(entered, LegacySalt) = LegacyPasswordHash()
        End If
        Return Hash(entered, PasswordSalt) = PasswordHash
    End Function

    Public Sub SetPassword(newPassword As String)
        Dim saltBytes = RandomNumberGenerator.GetBytes(16)
        PasswordSalt = Convert.ToBase64String(saltBytes)
        PasswordHash = Hash(newPassword, PasswordSalt)
    End Sub

    Private Shared Function Hash(text As String, salt As String) As String
        Return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(salt & "|" & text)))
    End Function

    ''' <summary>
    ''' The previous version compared against a fixed password in code. Its hash is computed here
    ''' once so the text never appears in the settings file.
    ''' </summary>
    Private Shared Function LegacyPasswordHash() As String
        Return Hash(Encoding.UTF8.GetString(Convert.FromBase64String("Sm9hbm5hMDA3")), LegacySalt)
    End Function

#End Region

#Region "Load and save"

    Public Shared ReadOnly Property AdminFile As String
        Get
            Return Path.Combine(AppSettings.SettingsFolder, "Admin.json")
        End Get
    End Property

    Public Shared Function TryLoad() As AdminSettings
        If Not File.Exists(AdminFile) Then Return Nothing
        Try
            Return SettingsJson.Read(Of AdminSettings)(File.ReadAllText(AdminFile))
        Catch ex As Exception
            SettingsJson.SetAsideDamagedFile(AdminFile, ex)
            Return Nothing
        End Try
    End Function

    Public Sub Save()
        SettingsJson.WriteFileSafely(AdminFile, SettingsJson.Write(Me))
    End Sub



#End Region

End Class
