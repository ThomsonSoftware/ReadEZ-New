Option Strict On

Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Text.Json

' ====================================================================================================
'   BackupService.vb - backups of the ReadEZ database, and moving it
'
'   A backup is one zip file in the chosen backup folder:
'       ReadEZ backup 2026-10-29 14-03-22.zip
'           backup.json          when, why (manual / on close / before a restore), version, counts
'           Database\ReadEZ.db   patients and results (copied with SQLite's backup function, so a
'                                consistent copy even if the database is in use)
'           Data\...             reading texts, letters, report pictures (optional)
'   Only the newest five are kept.
'
'   Restore copies the database from a backup over the one in use (and the data files over the
'   data folder, if wanted). A backup of the current state is made first, so a restore can
'   itself be undone. Settings are not part of a backup (they belong to the computer).
' ====================================================================================================

Public Enum BackupKind
    Manual
    OnClose
    BeforeRestore
End Enum

''' <summary>A backup file and what its backup.json says.</summary>
Public NotInheritable Class BackupInfo
    Public Property FilePath As String = ""
    Public Property Created As Date
    Public Property Kind As BackupKind
    Public Property Version As String = ""
    Public Property Computer As String = ""
    Public Property DatabaseName As String = ""
    Public Property Patients As Integer
    Public Property Recordings As Integer
    Public Property HasQuestionnaires As Boolean
    Public Property HasStimuli As Boolean
    Public Property SizeBytes As Long
End Class

Public NotInheritable Class BackupService

    ''' <summary>ReadEZ has no questionnaire database (kept so backups of other programs' format still read).</summary>
    Private Const NoQuestionnaireDb As String = ""
    Private Sub New()
    End Sub

    Public Const KeepCount As Integer = 5
    Private Const Prefix As String = "ReadEZ backup "
    Private Const ManifestName As String = "backup.json"
    Private Const DbFolder As String = "Database/"
    Private Const StimuliFolder As String = "Data/"

    Private Shared ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture
    Private Shared ReadOnly StoredExtensions As String() = {".jpg", ".jpeg", ".png", ".gif", ".mp4", ".avi", ".wmv", ".mov", ".mkv", ".mp3", ".wav", ".zip"}

    ''' <summary>Raised on the UI thread's side by callers after a restore or a move (screens reload).</summary>
    Public Shared Event DatabaseReplaced As EventHandler

    Public Shared Sub NotifyDatabaseReplaced()
        RaiseEvent DatabaseReplaced(Nothing, EventArgs.Empty)
    End Sub

#Region "Backup"

    ''' <summary>
    ''' Writes a backup to the folder and deletes all but the newest five (never the one in keepFile).
    ''' progress(text) reports what is being done. Throws if the backup could not be written.
    ''' </summary>
    Public Shared Function Create(folder As String, kind As BackupKind, includeStimuli As Boolean,
                                  Optional progress As Action(Of String) = Nothing, Optional keepFile As String = Nothing) As BackupInfo
        If String.IsNullOrWhiteSpace(folder) Then Throw New InvalidOperationException("No backup folder")
        Directory.CreateDirectory(folder)
        Dim s = AppHost.Settings
        Dim now = Date.Now
        Dim target = Path.Combine(folder, Prefix & now.ToString("yyyy-MM-dd HH-mm-ss", Inv) & ".zip")
        Dim n = 1
        While File.Exists(target)
            n += 1
            target = Path.Combine(folder, Prefix & now.ToString("yyyy-MM-dd HH-mm-ss", Inv) & $" ({n}).zip")
        End While
        Dim partialFile = target & ".partial"
        Dim temp = Path.Combine(Path.GetTempPath(), "ReadEZBackup_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(temp)
        Try
            ' Consistent copies of the databases
            progress?.Invoke(Loc.T("backup.progress.database"))
            Dim mainCopy = Path.Combine(temp, Path.GetFileName(s.Local_Database))
            CopyDatabase(s.Local_Database, mainCopy)
            Dim counts = (Patients:=0, Recordings:=0)
            Try
                counts = New ReadEZDatabase(mainCopy).Counts()
            Catch ex As Exception
                Log.Warn("Backup counts: " & ex.Message)
            End Try
            Dim questCopy As String = Nothing
            If Not String.IsNullOrWhiteSpace(NoQuestionnaireDb) AndAlso File.Exists(NoQuestionnaireDb) AndAlso
               Not String.Equals(Path.GetFullPath(NoQuestionnaireDb), Path.GetFullPath(s.Local_Database), StringComparison.OrdinalIgnoreCase) Then
                questCopy = Path.Combine(temp, "Quest_" & Path.GetFileName(NoQuestionnaireDb))
                CopyDatabase(NoQuestionnaireDb, questCopy)
            End If

            Dim stimuliRoot = If(s.DataDirectory, "")
            Dim withStimuli = includeStimuli AndAlso s.DataDirectory <> "" AndAlso Directory.Exists(stimuliRoot)
            Dim info As New BackupInfo With {
                .FilePath = target, .Created = now, .Kind = kind, .Version = AppHost.VersionNumber,
                .Computer = Environment.MachineName, .DatabaseName = Path.GetFileName(s.Local_Database),
                .Patients = counts.Patients, .Recordings = counts.Recordings,
                .HasQuestionnaires = questCopy IsNot Nothing, .HasStimuli = withStimuli}

            Using fs As New FileStream(partialFile, FileMode.Create, FileAccess.Write), zip As New ZipArchive(fs, ZipArchiveMode.Create)
                Dim manifest = zip.CreateEntry(ManifestName)
                Using w As New StreamWriter(manifest.Open())
                    w.Write(ManifestJson(info, Path.GetFileName(NoQuestionnaireDb)))
                End Using
                zip.CreateEntryFromFile(mainCopy, DbFolder & Path.GetFileName(s.Local_Database), CompressionLevel.Optimal)
                If questCopy IsNot Nothing Then
                    zip.CreateEntryFromFile(questCopy, DbFolder & "Questionnaires/" & Path.GetFileName(NoQuestionnaireDb), CompressionLevel.Optimal)
                End If
                If withStimuli Then
                    progress?.Invoke(Loc.T("backup.progress.stimuli"))
                    For Each f In StandardFiles.Folders.Select(Function(x) Path.Combine(stimuliRoot, x)).Where(Function(d) Directory.Exists(d)).SelectMany(Function(d) Directory.EnumerateFiles(d, "*", SearchOption.AllDirectories))
                        Dim rel = Path.GetRelativePath(stimuliRoot, f).Replace("\"c, "/"c)
                        Dim level = If(StoredExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()), CompressionLevel.NoCompression, CompressionLevel.Optimal)
                        Try
                            zip.CreateEntryFromFile(f, StimuliFolder & rel, level)
                        Catch ex As IOException
                            Log.Warn($"Backup: skipped {f}: {ex.Message}")
                        End Try
                    Next
                End If
            End Using
            File.Move(partialFile, target)
            info.SizeBytes = New FileInfo(target).Length
            Log.Info($"Backup written: {target} ({info.SizeBytes:N0} bytes, {kind})")

            s.BackupLastStamp = DatabaseStamp()
            Prune(folder, If(keepFile, target))
            Return info
        Finally
            Try
                SQLiteConnection.ClearAllPools()
                If File.Exists(partialFile) Then File.Delete(partialFile)
                Directory.Delete(temp, True)
            Catch ex As Exception
                Log.Warn("Backup clean-up: " & ex.Message)
            End Try
        End Try
    End Function

    Private Shared Function ManifestJson(info As BackupInfo, questName As String) As String
        Dim d As New Dictionary(Of String, Object) From {
            {"created", info.Created.ToString("yyyy-MM-dd HH:mm:ss", Inv)},
            {"kind", info.Kind.ToString()},
            {"version", info.Version},
            {"computer", info.Computer},
            {"database", info.DatabaseName},
            {"questionnaires", If(info.HasQuestionnaires, questName, "")},
            {"patients", info.Patients},
            {"recordings", info.Recordings},
            {"stimuli", info.HasStimuli}}
        Return JsonSerializer.Serialize(d, New JsonSerializerOptions With {
            .WriteIndented = True, .TypeInfoResolver = New Serialization.Metadata.DefaultJsonTypeInfoResolver()})
    End Function

    ''' <summary>Copies a SQLite database with SQLite's own backup function (safe while it is in use).</summary>
    Private Shared Sub CopyDatabase(source As String, target As String)
        If Not File.Exists(source) Then Throw New FileNotFoundException(Loc.F("backup.dbMissing", source), source)
        Using src As New SQLiteConnection($"Data Source={source};Version=3;Read Only=True;BusyTimeout=10000;"),
              dst As New SQLiteConnection($"Data Source={target};Version=3;")
            src.Open()
            dst.Open()
            src.BackupDatabase(dst, "main", "main", -1, Nothing, 0)
        End Using
        SQLiteConnection.ClearAllPools()
    End Sub

    ''' <summary>Deletes all but the newest KeepCount backups in a folder (never keepFile).</summary>
    Public Shared Sub Prune(folder As String, Optional keepFile As String = Nothing)
        Dim all = List(folder)
        Dim keep = all.Take(KeepCount).Select(Function(b) b.FilePath).ToList()
        If keepFile IsNot Nothing AndAlso Not keep.Contains(keepFile, StringComparer.OrdinalIgnoreCase) Then
            keep = all.Where(Function(b) String.Equals(b.FilePath, keepFile, StringComparison.OrdinalIgnoreCase)).
                       Concat(all.Where(Function(b) Not String.Equals(b.FilePath, keepFile, StringComparison.OrdinalIgnoreCase)).Take(KeepCount - 1)).
                       Select(Function(b) b.FilePath).ToList()
        End If
        For Each b In all
            If keep.Contains(b.FilePath, StringComparer.OrdinalIgnoreCase) Then Continue For
            Try
                File.Delete(b.FilePath)
                Log.Info("Old backup deleted: " & b.FilePath)
            Catch ex As Exception
                Log.Warn($"Delete old backup {b.FilePath}: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>Size and time of the databases: if unchanged since the last backup, closing doesn't back up again.</summary>
    Public Shared Function DatabaseStamp() As String
        Dim s = AppHost.Settings
        Dim parts As New List(Of String)
        For Each f In {s.Local_Database, NoQuestionnaireDb}
            Try
                If Not String.IsNullOrWhiteSpace(f) AndAlso File.Exists(f) Then
                    Dim fi As New FileInfo(f)
                    parts.Add($"{fi.FullName}|{fi.Length}|{fi.LastWriteTimeUtc.Ticks}")
                End If
            Catch
            End Try
        Next
        Return String.Join(";", parts)
    End Function

    ''' <summary>True if closing should make a backup (switched on, folder chosen, and something changed).</summary>
    Public Shared Function ShouldBackUpOnClose() As Boolean
        Dim s = AppHost.Settings
        Return s.BackupOnClose AndAlso Not String.IsNullOrWhiteSpace(s.BackupFolder) AndAlso DatabaseStamp() <> s.BackupLastStamp
    End Function

#End Region

#Region "List"

    ''' <summary>The backups in a folder, newest first.</summary>
    Public Shared Function List(folder As String) As List(Of BackupInfo)
        Dim result As New List(Of BackupInfo)
        If String.IsNullOrWhiteSpace(folder) OrElse Not Directory.Exists(folder) Then Return result
        For Each f In Directory.GetFiles(folder, Prefix & "*.zip")
            Dim info = ReadInfo(f)
            If info IsNot Nothing Then result.Add(info)
        Next
        Return result.OrderByDescending(Function(b) b.Created).ToList()
    End Function

    Private Shared Function ReadInfo(file As String) As BackupInfo
        Dim info As New BackupInfo With {.FilePath = file, .Created = IO.File.GetLastWriteTime(file)}
        Try
            info.SizeBytes = New FileInfo(file).Length
            Using zip = ZipFile.OpenRead(file)
                Dim m = zip.GetEntry(ManifestName)
                If m Is Nothing Then Return Nothing
                Using r As New StreamReader(m.Open()), doc = JsonDocument.Parse(r.ReadToEnd())
                    Dim root = doc.RootElement
                    Dim p As JsonElement
                    If root.TryGetProperty("created", p) Then
                        Dim d As Date
                        If Date.TryParseExact(p.GetString(), "yyyy-MM-dd HH:mm:ss", Inv, DateTimeStyles.None, d) Then info.Created = d
                    End If
                    If root.TryGetProperty("kind", p) Then
                        Dim k As BackupKind
                        If [Enum].TryParse(p.GetString(), k) Then info.Kind = k
                    End If
                    If root.TryGetProperty("version", p) Then info.Version = p.GetString()
                    If root.TryGetProperty("computer", p) Then info.Computer = p.GetString()
                    If root.TryGetProperty("database", p) Then info.DatabaseName = p.GetString()
                    If root.TryGetProperty("questionnaires", p) Then info.HasQuestionnaires = p.GetString() <> ""
                    If root.TryGetProperty("patients", p) Then info.Patients = p.GetInt32()
                    If root.TryGetProperty("recordings", p) Then info.Recordings = p.GetInt32()
                    If root.TryGetProperty("stimuli", p) Then info.HasStimuli = p.GetBoolean()
                End Using
            End Using
            Return info
        Catch ex As Exception
            Log.Warn($"Read backup {file}: {ex.Message}")
            Return Nothing
        End Try
    End Function

#End Region

#Region "Restore"

    ''' <summary>
    ''' Restores a backup: a backup of the current state is made first (in the same folder), then the
    ''' databases (and the stimulus files, if wanted) are copied back. Throws if it fails.
    ''' </summary>
    Public Shared Sub Restore(info As BackupInfo, restoreStimuli As Boolean, Optional progress As Action(Of String) = Nothing)
        Dim s = AppHost.Settings
        Dim temp = Path.Combine(Path.GetTempPath(), "ReadEZRestore_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(temp)
        Try
            ' Unpack and check the databases before touching anything
            progress?.Invoke(Loc.T("backup.progress.checking"))
            Dim mainCopy As String = Nothing, questCopy As String = Nothing
            Using zip = ZipFile.OpenRead(info.FilePath)
                For Each e In zip.Entries
                    If Not e.FullName.StartsWith(DbFolder, StringComparison.OrdinalIgnoreCase) OrElse e.Name = "" Then Continue For
                    Dim isQuest = e.FullName.StartsWith(DbFolder & "Questionnaires/", StringComparison.OrdinalIgnoreCase)
                    Dim target = Path.Combine(temp, If(isQuest, "quest.db", "main.db"))
                    e.ExtractToFile(target, True)
                    If isQuest Then questCopy = target Else mainCopy = target
                Next
            End Using
            If mainCopy Is Nothing Then Throw New InvalidDataException(Loc.T("backup.noDatabase"))
            CheckDatabase(mainCopy)
            If questCopy IsNot Nothing Then CheckDatabase(questCopy)

            ' Undo point
            If Not String.IsNullOrWhiteSpace(s.BackupFolder) Then
                progress?.Invoke(Loc.T("backup.progress.safety"))
                Create(s.BackupFolder, BackupKind.BeforeRestore, s.BackupIncludeStimuli, keepFile:=info.FilePath)
            End If

            ' Copy over the databases in use (in place: same file, so nothing else needs to change)
            progress?.Invoke(Loc.T("backup.progress.restoring"))
            SQLiteConnection.ClearAllPools()
            CopyDatabaseInto(mainCopy, s.Local_Database)
            If questCopy IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(NoQuestionnaireDb) Then
                CopyDatabaseInto(questCopy, NoQuestionnaireDb)
            End If

            If restoreStimuli AndAlso info.HasStimuli AndAlso Not String.IsNullOrWhiteSpace(s.DataDirectory) Then
                progress?.Invoke(Loc.T("backup.progress.stimuli"))
                Dim root = Path.GetFullPath(s.DataDirectory).TrimEnd(Path.DirectorySeparatorChar)
                Using zip = ZipFile.OpenRead(info.FilePath)
                    For Each e In zip.Entries
                        If Not e.FullName.StartsWith(StimuliFolder, StringComparison.OrdinalIgnoreCase) OrElse e.Name = "" Then Continue For
                        Dim target = Path.GetFullPath(Path.Combine(root, e.FullName.Substring(StimuliFolder.Length)))
                        If Not target.StartsWith(root & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then Continue For   ' stay inside the data folder
                        Directory.CreateDirectory(Path.GetDirectoryName(target))
                        Try
                            e.ExtractToFile(target, True)
                        Catch ex As IOException
                            Log.Warn($"Restore: skipped {target}: {ex.Message}")
                        End Try
                    Next
                End Using
            End If
            s.BackupLastStamp = DatabaseStamp()
            Log.Info("Backup restored: " & info.FilePath)
        Finally
            Try
                SQLiteConnection.ClearAllPools()
                Directory.Delete(temp, True)
            Catch ex As Exception
                Log.Warn("Restore clean-up: " & ex.Message)
            End Try
        End Try
    End Sub

    ''' <summary>Throws unless the file is a readable SQLite database that passes a quick check.</summary>
    Private Shared Sub CheckDatabase(file As String)
        Using c As New SQLiteConnection($"Data Source={file};Version=3;Read Only=True;"), cmd = c.CreateCommand()
            c.Open()
            cmd.CommandText = "PRAGMA quick_check;"
            Dim r = Convert.ToString(cmd.ExecuteScalar(), Inv)
            If Not String.Equals(r, "ok", StringComparison.OrdinalIgnoreCase) Then Throw New InvalidDataException(Loc.F("backup.damaged", r))
        End Using
        SQLiteConnection.ClearAllPools()
    End Sub

    ''' <summary>Replaces the contents of a database file with another database (SQLite backup function).</summary>
    Private Shared Sub CopyDatabaseInto(source As String, target As String)
        Directory.CreateDirectory(Path.GetDirectoryName(target))
        Using src As New SQLiteConnection($"Data Source={source};Version=3;Read Only=True;"),
              dst As New SQLiteConnection($"Data Source={target};Version=3;BusyTimeout=10000;")
            src.Open()
            dst.Open()
            src.BackupDatabase(dst, "main", "main", -1, Nothing, 0)
        End Using
        SQLiteConnection.ClearAllPools()
    End Sub

#End Region

#Region "Move"

    ''' <summary>
    ''' Copies the patient database to a folder, checks the copy, and from then on uses the copy.
    ''' The old file is renamed to "... .moved" (kept, but no longer opened). Returns the new path.
    ''' </summary>
    Public Shared Function MoveDatabase(newFolder As String) As String
        Dim s = AppHost.Settings
        Dim oldPath = s.Local_Database
        Dim newPath = Path.Combine(newFolder, Path.GetFileName(oldPath))
        If String.Equals(Path.GetFullPath(newPath), Path.GetFullPath(oldPath), StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException(Loc.T("backup.move.same"))
        End If
        If File.Exists(newPath) Then Throw New IOException(Loc.F("backup.move.exists", newPath))
        Directory.CreateDirectory(newFolder)
        Dim before = AppHost.Db.Counts()
        CopyDatabase(oldPath, newPath)
        Dim moved As ReadEZDatabase
        Try
            moved = New ReadEZDatabase(newPath)
            Dim after = moved.Counts()
            If Not after.Equals(before) Then Throw New InvalidDataException(Loc.T("backup.move.verify"))
        Catch
            Try
                SQLiteConnection.ClearAllPools()
                File.Delete(newPath)
            Catch
            End Try
            Throw
        End Try
        AppHost.Db = moved
        s.Local_Database = newPath
        AppHost.SaveSettings()
        SQLiteConnection.ClearAllPools()
        Try
            Dim retired = oldPath & ".moved"
            If File.Exists(retired) Then File.Delete(retired)
            File.Move(oldPath, retired)
        Catch ex As Exception
            Log.Warn($"Rename old database {oldPath}: {ex.Message}")
        End Try
        s.BackupLastStamp = ""
        Log.Info($"Database moved: {oldPath} -> {newPath}")
        Return newPath
    End Function

#End Region

End Class
