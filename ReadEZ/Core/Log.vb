Option Strict On

Imports System.IO

' ====================================================================================================
'   Log.vb
'   A small daily log file for support and debugging:
'       %AppData%\ReadEZ\Logs\2026-10-26.log
'   Logs older than 30 days are deleted at start-up. Messages also go to the Visual Studio
'   Output window (Debug.WriteLine).
' ====================================================================================================

Public NotInheritable Class Log
    Private Sub New()
    End Sub

    Private Shared ReadOnly _lock As New Object
    Private Shared _cleaned As Boolean

    Public Shared ReadOnly Property Folder As String
        Get
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ReadEZ", "Logs")
        End Get
    End Property

    Public Shared Sub Info(message As String)
        Write("INFO ", message)
    End Sub

    Public Shared Sub Warn(message As String)
        Write("WARN ", message)
    End Sub

    Public Shared Sub [Error](where As String, ex As Exception)
        Write("ERROR", where & ": " & If(ex Is Nothing, "(no details)", ex.ToString()))
    End Sub

    Public Shared Sub [Error](message As String)
        Write("ERROR", message)
    End Sub

    Private Shared Sub Write(level As String, message As String)
        Dim line = $"{Date.Now:HH:mm:ss.fff} {level} [{Threading.Thread.CurrentThread.ManagedThreadId}] {message}"
        Debug.WriteLine(line)
        SyncLock _lock
            Try
                Directory.CreateDirectory(Folder)
                If Not _cleaned Then
                    _cleaned = True
                    For Each f In Directory.GetFiles(Folder, "*.log")
                        If File.GetLastWriteTime(f) < Date.Now.AddDays(-30) Then File.Delete(f)
                    Next
                End If
                File.AppendAllText(Path.Combine(Folder, Date.Now.ToString("yyyy-MM-dd") & ".log"), line & Environment.NewLine)
            Catch
                ' Logging must never stop the program.
            End Try
        End SyncLock
    End Sub
End Class
