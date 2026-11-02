Option Strict On

Imports System.IO

' ====================================================================================================
'   StandardFiles.vb - reading texts, report letters and pictures in the data folder
'
'   ReadEZ reads these from the data folder (Settings.DataDirectory):
'     Reading texts      extra texts for the symptoms and rate-of-reading tests (*.txt, one row per line)
'     ReadEZ Templates   the opening and closing paragraphs of the letters to parents and teachers
'     Report images      the logo printed on reports
'   The starter set is installed with the program in "<program>\Starter" (from ReadEZ\Deploy\Starter
'   when the program is built). It is kept apart from the data folder because the program folder
'   (Program Files) cannot be written to and is replaced by every update. At start-up:
'     - any of the standard files missing from the data folder are copied in (existing files are
'       never replaced, so texts and letters the user has changed are kept), and
'     - Starter\Database\ReadEZ.db (if the installer has one) is copied to the database path in
'       Settings, but only the very first time ReadEZ runs on the computer and only at the default
'       path, so an existing database is never replaced.
'   Settings > Data > "Copy from another folder" does the same from a folder the user picks.
' ====================================================================================================

Public NotInheritable Class StandardFiles
    Private Sub New()
    End Sub

    Public Shared ReadOnly Folders As String() = {"Reading texts", "ReadEZ Templates", "Report images"}

    ''' <summary>The starter set shipped with the program.</summary>
    Public Shared ReadOnly Property InstalledFolder As String
        Get
            Return Path.Combine(AppHost.AppFolder, "Starter")
        End Get
    End Property

    ''' <summary>Copies any missing standard files and starter databases into place (quietly; called at start-up).</summary>
    Public Shared Sub EnsureInstalled()
        Try
            Dim target = AppHost.Settings.DataDirectory
            If String.IsNullOrWhiteSpace(target) Then Return
            If Not Directory.Exists(InstalledFolder) Then Return
            If String.Equals(Path.GetFullPath(target).TrimEnd("\"c), Path.GetFullPath(InstalledFolder).TrimEnd("\"c), StringComparison.OrdinalIgnoreCase) Then Return
            Dim n = CopyMissing(InstalledFolder, target)
            If n > 0 Then Log.Info($"Installed {n} standard files into {target}")
            InstallStarterDatabase("ReadEZ.db", AppHost.Settings.Local_Database)
        Catch ex As Exception
            Log.Error("StandardFiles.EnsureInstalled", ex)
        End Try
    End Sub

    ''' <summary>Only the starter databases (needed before the database is opened at start-up; the rest
    ''' of EnsureInstalled runs in the background once the main window is showing).</summary>
    Public Shared Sub EnsureStarterDatabases()
        Try
            If Not Directory.Exists(InstalledFolder) Then Return
            InstallStarterDatabase("ReadEZ.db", AppHost.Settings.Local_Database)
        Catch ex As Exception
            Log.Error("StandardFiles.EnsureStarterDatabases", ex)
        End Try
    End Sub

    ''' <summary>Copies a starter database to its path in Settings when no database exists there yet.</summary>
    Private Shared Sub InstallStarterDatabase(fileName As String, target As String)
        If String.IsNullOrWhiteSpace(target) OrElse File.Exists(target) Then Return
        If Not IsNewInstallation(target) Then Return
        Dim source = {Path.Combine(InstalledFolder, "Database", fileName), Path.Combine(InstalledFolder, fileName)}.FirstOrDefault(Function(f) File.Exists(f))
        If source Is Nothing Then Return
        Directory.CreateDirectory(Path.GetDirectoryName(target))
        File.Copy(source, target, overwrite:=False)
        Log.Info($"Installed starter database {source} as {target}")
    End Sub

    ''' <summary>
    ''' The starter database is only for a computer ReadEZ has never run on, and only at
    ''' the default place (a database moved elsewhere, or on a drive that isn't connected, is not
    ''' replaced by example patients).
    ''' </summary>
    Private Shared Function IsNewInstallation(target As String) As Boolean
        If Not AppHost.FirstRun Then Return False
        Try
            Dim standard = Path.Combine(AppHost.DefaultDataFolder, "Database", "ReadEZ.db")
            Return String.Equals(Path.GetFullPath(target), Path.GetFullPath(standard), StringComparison.OrdinalIgnoreCase)
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>Copies the standard folders from source into target, skipping files that exist. Returns the number copied.</summary>
    Public Shared Function CopyMissing(source As String, target As String) As Integer
        Dim count = 0
        For Each f In Folders
            Dim src = Path.Combine(source, f)
            If Not Directory.Exists(src) Then Continue For
            For Each file In Directory.EnumerateFiles(src, "*", SearchOption.AllDirectories)
                Dim rel = Path.GetRelativePath(source, file)
                Dim dest = Path.Combine(target, rel)
                If IO.File.Exists(dest) Then Continue For
                Directory.CreateDirectory(Path.GetDirectoryName(dest))
                IO.File.Copy(file, dest)
                count += 1
            Next
        Next
        Return count
    End Function

    ''' <summary>True if a folder looks like a data folder (has at least one of the standard folders).</summary>
    Public Shared Function LooksLikeDataFolder(folder As String) As Boolean
        Return Folders.Any(Function(f) Directory.Exists(Path.Combine(folder, f)))
    End Function
End Class
