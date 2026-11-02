Option Strict On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography.X509Certificates

' ====================================================================================================
'   CodeSigning.vb - Authenticode signature of a file (the program itself, or a downloaded update)
'
'   Used for:
'     - the release builds: the program and installer are signed with the .pfx certificate,
'       an unsigned build the Developer licence (AdminSettings.EffectiveLicenceMode)
'     - updates: a downloaded installer must carry a valid signature from the same publisher as
'       the running program before it is run (UpdateService)
'     - the Admin "Publish update" window, to show whether the build was signed
'
'   The signature is checked with WinVerifyTrust; revocation is not checked online so the check
'   also works without internet. Thomson's certificate (Thomson.pfx) is self-signed, so Windows
'   reports "untrusted root" (0x800B0109) even though the file is intact. That is normal: the
'   update check relies on the file being intact and signed with that
'   same certificate (IsIntact + same thumbprint), not on Windows trusting the certificate.
' ====================================================================================================

Public NotInheritable Class SignatureInfo
    ''' <summary>The file has an Authenticode signature.</summary>
    Public Property IsSigned As Boolean
    ''' <summary>The signature is valid: the file is unchanged and the certificate is trusted.</summary>
    Public Property IsValid As Boolean
    ''' <summary>Signer's certificate subject, e.g. "CN=Thomson Software Solutions, ...".</summary>
    Public Property Subject As String = ""
    Public Property Thumbprint As String = ""
    ''' <summary>The certificate was issued by itself (Thomson.pfx), so Windows doesn't trust it.</summary>
    Public Property IsSelfSigned As Boolean

    ''' <summary>
    ''' Signed and unchanged since signing: valid, or valid except that Windows doesn't trust the
    ''' certificate's root (self-signed certificate).
    ''' </summary>
    Public ReadOnly Property IsIntact As Boolean
        Get
            Return IsSigned AndAlso (IsValid OrElse ErrorCode = CodeSigning.CERT_E_UNTRUSTEDROOT OrElse
                                     ErrorCode = CodeSigning.CERT_E_CHAINING OrElse ErrorCode = CodeSigning.CERT_E_UNTRUSTEDTESTROOT)
        End Get
    End Property

    ''' <summary>Short description for logs, e.g. "signed by Thomson (self-signed certificate, intact)".</summary>
    Public Function Describe() As String
        If Not IsSigned Then Return "not signed"
        If IsValid Then Return $"signed by {SignerName} (valid)"
        If IsIntact Then Return $"signed by {SignerName} ({If(IsSelfSigned, "self-signed certificate", "certificate not trusted by Windows")}, file intact)"
        Return $"signed by {SignerName} (NOT valid - file changed or bad signature, error 0x{ErrorCode:X8})"
    End Function
    ''' <summary>WinVerifyTrust result when the signature is not valid (0 = valid).</summary>
    Public Property ErrorCode As Integer

    ''' <summary>The CN part of the subject (the publisher's name).</summary>
    Public ReadOnly Property SignerName As String
        Get
            For Each part In Subject.Split(","c)
                Dim p = part.Trim()
                If p.StartsWith("CN=", StringComparison.OrdinalIgnoreCase) Then Return p.Substring(3).Trim(""""c)
            Next
            Return Subject
        End Get
    End Property
End Class

Public NotInheritable Class CodeSigning
    Private Sub New()
    End Sub

    Private Shared _self As SignatureInfo

    Friend Const CERT_E_UNTRUSTEDROOT As Integer = &H800B0109
    Friend Const CERT_E_CHAINING As Integer = &H800B010A
    Friend Const CERT_E_UNTRUSTEDTESTROOT As Integer = &H800B010D

    ''' <summary>Path of the running program (ReadEZ.exe).</summary>
    Public Shared ReadOnly Property ProgramFile As String
        Get
            Dim p = Environment.ProcessPath
            If String.IsNullOrEmpty(p) Then p = Path.Combine(AppContext.BaseDirectory, "ReadEZ.exe")
            Return p
        End Get
    End Property

    ''' <summary>Signature of the running program (checked once).</summary>
    Public Shared ReadOnly Property Program As SignatureInfo
        Get
            If _self Is Nothing Then _self = Check(ProgramFile)
            Return _self
        End Get
    End Property

    ''' <summary>Reads and verifies the Authenticode signature of a file.</summary>
    Public Shared Function Check(file As String) As SignatureInfo
        Dim info As New SignatureInfo
        If String.IsNullOrEmpty(file) OrElse Not IO.File.Exists(file) Then Return info
        Try
#Disable Warning SYSLIB0057
            Using cert As New X509Certificate2(X509Certificate.CreateFromSignedFile(file))
#Enable Warning SYSLIB0057
                info.IsSigned = True
                info.Subject = cert.Subject
                info.Thumbprint = cert.Thumbprint
                info.IsSelfSigned = String.Equals(cert.Subject, cert.Issuer, StringComparison.OrdinalIgnoreCase)
            End Using
        Catch ex As Exception
            Return info          ' no signature
        End Try
        Try
            info.ErrorCode = VerifyTrust(file)
            info.IsValid = info.ErrorCode = 0
        Catch ex As Exception
            Log.Warn("WinVerifyTrust failed: " & ex.Message)
        End Try
        Return info
    End Function

    ''' <summary>
    ''' True if two files are signed with the same certificate (same thumbprint). Stricter than the
    ''' same name: anyone can make a self-signed certificate called "Thomson".
    ''' </summary>
    Public Shared Function SameCertificate(a As SignatureInfo, b As SignatureInfo) As Boolean
        Return a IsNot Nothing AndAlso b IsNot Nothing AndAlso a.IsSigned AndAlso b.IsSigned AndAlso
               a.Thumbprint <> "" AndAlso String.Equals(a.Thumbprint, b.Thumbprint, StringComparison.OrdinalIgnoreCase)
    End Function

#Region "WinVerifyTrust"

    Private Shared ReadOnly WINTRUST_ACTION_GENERIC_VERIFY_V2 As New Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE")

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Private Structure WINTRUST_FILE_INFO
        Public cbStruct As UInteger
        Public pcwszFilePath As String
        Public hFile As IntPtr
        Public pgKnownSubject As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure WINTRUST_DATA
        Public cbStruct As UInteger
        Public pPolicyCallbackData As IntPtr
        Public pSIPClientData As IntPtr
        Public dwUIChoice As UInteger
        Public fdwRevocationChecks As UInteger
        Public dwUnionChoice As UInteger
        Public pFile As IntPtr
        Public dwStateAction As UInteger
        Public hWVTStateData As IntPtr
        Public pwszURLReference As IntPtr
        Public dwProvFlags As UInteger
        Public dwUIContext As UInteger
        Public pSignatureSettings As IntPtr
    End Structure

    Private Const WTD_UI_NONE As UInteger = 2
    Private Const WTD_REVOKE_NONE As UInteger = 0
    Private Const WTD_CHOICE_FILE As UInteger = 1
    Private Const WTD_STATEACTION_VERIFY As UInteger = 1
    Private Const WTD_STATEACTION_CLOSE As UInteger = 2
    Private Const WTD_CACHE_ONLY_URL_RETRIEVAL As UInteger = &H1000

    <DllImport("wintrust.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function WinVerifyTrust(hwnd As IntPtr, <MarshalAs(UnmanagedType.LPStruct)> pgActionID As Guid, ByRef pWVTData As WINTRUST_DATA) As Integer
    End Function

    Private Shared Function VerifyTrust(file As String) As Integer
        Dim fileInfo As New WINTRUST_FILE_INFO With {
            .cbStruct = CUInt(Marshal.SizeOf(GetType(WINTRUST_FILE_INFO))), .pcwszFilePath = file}
        Dim pFile = Marshal.AllocHGlobal(Marshal.SizeOf(fileInfo))
        Try
            Marshal.StructureToPtr(fileInfo, pFile, False)
            Dim data As New WINTRUST_DATA With {
                .cbStruct = CUInt(Marshal.SizeOf(GetType(WINTRUST_DATA))), .dwUIChoice = WTD_UI_NONE,
                .fdwRevocationChecks = WTD_REVOKE_NONE, .dwUnionChoice = WTD_CHOICE_FILE, .pFile = pFile,
                .dwStateAction = WTD_STATEACTION_VERIFY, .dwProvFlags = WTD_CACHE_ONLY_URL_RETRIEVAL}
            Dim result = WinVerifyTrust(IntPtr.Zero, WINTRUST_ACTION_GENERIC_VERIFY_V2, data)
            data.dwStateAction = WTD_STATEACTION_CLOSE
            WinVerifyTrust(IntPtr.Zero, WINTRUST_ACTION_GENERIC_VERIFY_V2, data)
            Return result
        Finally
            Marshal.DestroyStructure(Of WINTRUST_FILE_INFO)(pFile)
            Marshal.FreeHGlobal(pFile)
        End Try
    End Function

#End Region

End Class
