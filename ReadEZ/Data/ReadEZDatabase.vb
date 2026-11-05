Option Strict On

Imports System.Data
Imports System.Data.SQLite

' ====================================================================================================
'   ReadEZDatabase.vb - patients, organisations, screeners and test results (SQLite)
'
'   New, optimised structure (schema version 1). ReadEZ 2020 kept baseline and colour results in two
'   tables with ~45 loosely typed columns each (numbers stored as text, dates in several formats,
'   unused eye-tracking and video columns). This version keeps:
'
'     Organisations  Id, Name, address, Tel
'     Screeners      Id, Title, Firstname, Surname, Active
'     Patients       Id, names, DoB (ISO date), Reference, address, contact, OrganisationId -> Organisations
'     Tests          one row per test - baseline (Mode 0) and colour tests (1 overlay, 2 computer,
'                    3 lenses) together, so a patient's history is one indexed query:
'                      colour      ColourIndex (0 grey, 1-12 = A-L), ColourSet (1 overlays, 2 lenses,
'                                  3 dark lenses), ColourRgb (as shown), GreyPreferred, Method, AutoScore
'                      symptoms    Symptom1-8 (0-5), SymptomsDone, SymptomNotes
'                      reading     Wpm, ReadSeconds, WordsRead, Errors, RowsSkipped, Transcript
'                      how shown   TextIndex, TextName, FontName, FontSize, LineSpacing, RowCount,
'                                  OverlaySize, PresentationMs, ScreenWidth, ScreenHeight
'                      Notes, LegacySource / LegacyId (where an imported test came from)
'     Meta           Key / Value (SchemaVersion, ...)
'
'   Integers are INTEGER, dates are ISO-8601 text (sortable), foreign keys are enforced with
'   ON DELETE CASCADE (deleting a patient deletes their tests), the journal is WAL and the lists
'   use covering indexes. Older databases are brought across with LegacyDatabaseImport.
' ====================================================================================================

Public NotInheritable Class ReadEZDatabase

    Public Const SchemaVersion As Integer = 1

    Private ReadOnly _connectionString As String
    Private Shared ReadOnly Inv As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
    Private Const DateTimeFormat As String = "yyyy-MM-dd HH:mm:ss"

    Public ReadOnly Property FilePath As String

    Public Sub New(databasePath As String)
        FilePath = databasePath
        _connectionString = $"Data Source={databasePath};Version=3;BusyTimeout=5000;Foreign Keys=True;"
        If IsLegacyFile(databasePath) Then
            Throw New InvalidOperationException(Loc.F("db.legacyFile", databasePath))
        End If
        InitializeSchema()
    End Sub

    Friend Function Open() As SQLiteConnection
        Dim conn As New SQLiteConnection(_connectionString)
        conn.Open()
        Using cmd As New SQLiteCommand("PRAGMA foreign_keys = ON;", conn)
            cmd.ExecuteNonQuery()
        End Using
        Return conn
    End Function

#Region "Schema"

    ''' <summary>True if the file is a ReadEZ 2020 database (Clients table, no Patients table): it must be imported, not opened.</summary>
    Public Shared Function IsLegacyFile(path As String) As Boolean
        If Not IO.File.Exists(path) OrElse New IO.FileInfo(path).Length = 0 Then Return False
        Try
            Using conn As New SQLiteConnection($"Data Source={path};Version=3;Read Only=True;")
                conn.Open()
                Dim tables As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Using cmd As New SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", conn), r = cmd.ExecuteReader()
                    While r.Read()
                        tables.Add(Str(r(0)))
                    End While
                End Using
                Return tables.Contains("Clients") AndAlso Not tables.Contains("Patients")
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub InitializeSchema()
        Using conn = Open(), tx = conn.BeginTransaction(), cmd = conn.CreateCommand()
            Dim exec = Sub(sql As String)
                           cmd.CommandText = sql
                           cmd.ExecuteNonQuery()
                       End Sub
            exec("CREATE TABLE IF NOT EXISTS Meta (Key TEXT PRIMARY KEY, Value TEXT) WITHOUT ROWID;")
            exec("CREATE TABLE IF NOT EXISTS Organisations (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL DEFAULT '', Address1 TEXT NOT NULL DEFAULT '', Address2 TEXT NOT NULL DEFAULT '',
                    Town TEXT NOT NULL DEFAULT '', County TEXT NOT NULL DEFAULT '', Postcode TEXT NOT NULL DEFAULT '',
                    Tel TEXT NOT NULL DEFAULT '', LegacyId INTEGER);")
            exec("CREATE TABLE IF NOT EXISTS Screeners (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL DEFAULT '', Firstname TEXT NOT NULL DEFAULT '', Surname TEXT NOT NULL DEFAULT '',
                    Active INTEGER NOT NULL DEFAULT 1, LegacyId INTEGER);")
            exec("CREATE TABLE IF NOT EXISTS Patients (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL DEFAULT '', Firstname TEXT NOT NULL DEFAULT '', Surname TEXT NOT NULL DEFAULT '',
                    DoB TEXT, Reference TEXT NOT NULL DEFAULT '',
                    Address1 TEXT NOT NULL DEFAULT '', Address2 TEXT NOT NULL DEFAULT '', Town TEXT NOT NULL DEFAULT '',
                    County TEXT NOT NULL DEFAULT '', Postcode TEXT NOT NULL DEFAULT '',
                    Tel TEXT NOT NULL DEFAULT '', Mobile TEXT NOT NULL DEFAULT '', Email TEXT NOT NULL DEFAULT '',
                    OrganisationId INTEGER REFERENCES Organisations(Id) ON DELETE SET NULL,
                    Created TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S','now','localtime')),
                    LegacySource TEXT, LegacyId INTEGER);")
            exec("CREATE TABLE IF NOT EXISTS Tests (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    PatientId INTEGER NOT NULL REFERENCES Patients(Id) ON DELETE CASCADE,
                    ScreenerId INTEGER REFERENCES Screeners(Id) ON DELETE SET NULL,
                    TestDate TEXT NOT NULL,
                    Mode INTEGER NOT NULL DEFAULT 0 CHECK (Mode BETWEEN 0 AND 3),
                    ColourIndex INTEGER NOT NULL DEFAULT 0, ColourSet INTEGER NOT NULL DEFAULT 1, ColourRgb INTEGER,
                    GreyPreferred INTEGER NOT NULL DEFAULT 0, Method INTEGER NOT NULL DEFAULT 0, AutoScore TEXT NOT NULL DEFAULT '',
                    Symptom1 INTEGER NOT NULL DEFAULT 0, Symptom2 INTEGER NOT NULL DEFAULT 0, Symptom3 INTEGER NOT NULL DEFAULT 0,
                    Symptom4 INTEGER NOT NULL DEFAULT 0, Symptom5 INTEGER NOT NULL DEFAULT 0, Symptom6 INTEGER NOT NULL DEFAULT 0,
                    Symptom7 INTEGER NOT NULL DEFAULT 0, Symptom8 INTEGER NOT NULL DEFAULT 0,
                    SymptomsDone INTEGER NOT NULL DEFAULT 0, SymptomNotes TEXT NOT NULL DEFAULT '',
                    Wpm INTEGER NOT NULL DEFAULT 0, ReadSeconds INTEGER NOT NULL DEFAULT 0, WordsRead INTEGER NOT NULL DEFAULT 0,
                    Errors INTEGER NOT NULL DEFAULT 0, RowsSkipped INTEGER NOT NULL DEFAULT 0, Transcript TEXT NOT NULL DEFAULT '',
                    TextIndex INTEGER NOT NULL DEFAULT 0, TextName TEXT NOT NULL DEFAULT '',
                    FontName TEXT NOT NULL DEFAULT 'Times New Roman', FontSize INTEGER NOT NULL DEFAULT 5,
                    LineSpacing INTEGER NOT NULL DEFAULT 6, RowCount INTEGER NOT NULL DEFAULT 20,
                    OverlaySize INTEGER NOT NULL DEFAULT 5, PresentationMs INTEGER NOT NULL DEFAULT 0,
                    ScreenWidth INTEGER NOT NULL DEFAULT 0, ScreenHeight INTEGER NOT NULL DEFAULT 0,
                    Notes TEXT NOT NULL DEFAULT '',
                    LegacySource TEXT, LegacyId INTEGER);")
            ' Pattern tests (added in ReadEZ 3.x): the stimulus (0 text, 1 pattern) and the pattern used
            Dim existingColumns As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            cmd.CommandText = "PRAGMA table_info(Tests);"
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    existingColumns.Add(Str(rd("name")))
                End While
            End Using
            If Not existingColumns.Contains("Stimulus") Then exec("ALTER TABLE Tests ADD COLUMN Stimulus INTEGER NOT NULL DEFAULT 0;")
            If Not existingColumns.Contains("PatternCpd") Then exec("ALTER TABLE Tests ADD COLUMN PatternCpd REAL NOT NULL DEFAULT 0;")
            If Not existingColumns.Contains("PatternDistanceCm") Then exec("ALTER TABLE Tests ADD COLUMN PatternDistanceCm INTEGER NOT NULL DEFAULT 0;")
            exec("CREATE INDEX IF NOT EXISTS ix_patients_name ON Patients(Surname COLLATE NOCASE, Firstname COLLATE NOCASE);")
            exec("CREATE INDEX IF NOT EXISTS ix_patients_org ON Patients(OrganisationId);")
            exec("CREATE INDEX IF NOT EXISTS ix_tests_patient ON Tests(PatientId, TestDate DESC);")
            exec("CREATE UNIQUE INDEX IF NOT EXISTS ux_tests_legacy ON Tests(LegacySource, LegacyId) WHERE LegacySource IS NOT NULL;")
            exec("CREATE UNIQUE INDEX IF NOT EXISTS ux_patients_legacy ON Patients(LegacySource, LegacyId) WHERE LegacySource IS NOT NULL;")
            exec($"INSERT OR IGNORE INTO Meta (Key, Value) VALUES ('SchemaVersion', '{SchemaVersion}');")
            exec("INSERT OR IGNORE INTO Meta (Key, Value) VALUES ('Created', strftime('%Y-%m-%d %H:%M:%S','now','localtime'));")
            tx.Commit()
        End Using
        Try
            Using conn = Open(), cmd As New SQLiteCommand("PRAGMA journal_mode=WAL;", conn)
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Log.Warn("WAL journal: " & ex.Message)   ' e.g. a network drive; the default journal is fine
        End Try
    End Sub

#End Region

#Region "Helpers"

    Friend Function Query(sql As String, ParamArray args As Object()) As DataTable
        Using conn = Open(), cmd As New SQLiteCommand(sql, conn)
            AddArgs(cmd, args)
            Dim dt As New DataTable()
            Using r = cmd.ExecuteReader()
                dt.Load(r)
            End Using
            Return dt
        End Using
    End Function

    Friend Function Execute(sql As String, ParamArray args As Object()) As Integer
        Using conn = Open(), cmd As New SQLiteCommand(sql, conn)
            AddArgs(cmd, args)
            Return cmd.ExecuteNonQuery()
        End Using
    End Function

    Private Function Insert(sql As String, ParamArray args As Object()) As Integer
        Using conn = Open()
            Return InsertOn(conn, Nothing, sql, args)
        End Using
    End Function

    ''' <summary>Runs an INSERT on an open connection and returns the new row id.</summary>
    Friend Shared Function InsertOn(conn As SQLiteConnection, tx As SQLiteTransaction, sql As String, ParamArray args As Object()) As Integer
        Using cmd As New SQLiteCommand(sql, conn, tx)
            AddArgs(cmd, args)
            cmd.ExecuteNonQuery()
            cmd.Parameters.Clear()
            cmd.CommandText = "SELECT last_insert_rowid();"
            Return CInt(Convert.ToInt64(cmd.ExecuteScalar()))
        End Using
    End Function

    Friend Shared Sub AddArgs(cmd As SQLiteCommand, args As Object())
        For i = 0 To args.Length - 2 Step 2
            Dim v = args(i + 1)
            If v Is Nothing Then v = DBNull.Value
            If TypeOf v Is Boolean Then v = If(DirectCast(v, Boolean), 1, 0)
            cmd.Parameters.AddWithValue(CStr(args(i)), v)
        Next
    End Sub

    Public Shared Function Str(v As Object) As String
        If v Is Nothing OrElse v Is DBNull.Value Then Return ""
        Return v.ToString()
    End Function

    Public Shared Function Int(v As Object) As Integer
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0
        If TypeOf v Is Boolean Then Return If(DirectCast(v, Boolean), 1, 0)
        Dim s = TryCast(v, String)
        If s IsNot Nothing Then
            Dim n As Double
            Return If(Double.TryParse(s.Trim(), Globalization.NumberStyles.Float, Inv, n), CInt(Math.Round(n)), 0)
        End If
        Try
            Return Convert.ToInt32(v, Inv)
        Catch
            Return 0
        End Try
    End Function

    Public Shared Function Bool(v As Object) As Boolean
        If v Is Nothing OrElse v Is DBNull.Value Then Return False
        If TypeOf v Is Boolean Then Return DirectCast(v, Boolean)
        Dim s = v.ToString().Trim()
        Return s = "1" OrElse s.Equals("true", StringComparison.OrdinalIgnoreCase) OrElse s = "-1"
    End Function

    ''' <summary>Reads a date written by this version (ISO) or by ReadEZ 2020 (several formats, or ticks).</summary>
    Public Shared Function DateOf(v As Object) As Date
        If v Is Nothing OrElse v Is DBNull.Value Then Return Date.MinValue
        If TypeOf v Is Date Then Return DirectCast(v, Date)
        Dim s = v.ToString().Trim()
        If s = "" Then Return Date.MinValue
        Dim d As Date
        If Date.TryParse(s, Inv, Globalization.DateTimeStyles.None, d) Then Return d
        If Date.TryParse(s, d) Then Return d
        Dim ticks As Long
        If Long.TryParse(s, ticks) AndAlso ticks > 0 Then
            Try
                Return New Date(ticks)
            Catch
            End Try
        End If
        Return Date.MinValue
    End Function

    Friend Shared Function DateText(d As Date) As String
        Return d.ToString(DateTimeFormat, Inv)
    End Function

    Private Shared Function Col(r As DataRow, name As String) As Object
        Return If(r.Table.Columns.Contains(name), r(name), Nothing)
    End Function

#End Region

#Region "Meta"

    Public Function GetMeta(key As String) As String
        Dim dt = Query("SELECT Value FROM Meta WHERE Key=@k;", "@k", key)
        Return If(dt.Rows.Count = 0, "", Str(dt.Rows(0)(0)))
    End Function

    Public Sub SetMeta(key As String, value As String)
        Execute("INSERT INTO Meta (Key, Value) VALUES (@k, @v) ON CONFLICT(Key) DO UPDATE SET Value=excluded.Value;", "@k", key, "@v", value)
    End Sub

#End Region

#Region "Patients"

    Friend Shared Function PatientArgs(p As PatientStructure) As Object()
        Return {"@Title", If(p.Title, ""), "@Firstname", If(p.Firstname, ""), "@Surname", If(p.Surname, ""), "@Reference", If(p.Reference, ""),
                "@Address1", If(p.Address1, ""), "@Address2", If(p.Address2, ""), "@Town", If(p.Town, ""), "@County", If(p.County, ""),
                "@Postcode", If(p.Postcode, ""), "@Tel", If(p.Tel, ""), "@Mobile", If(p.Mobile, ""), "@Email", If(p.Email, ""),
                "@DoB", If(Ages.IsKnown(p.DoB), CObj(p.DoB.ToString("yyyy-MM-dd", Inv)), Nothing),
                "@Org", If(p.Organisation_ID > 0, CObj(p.Organisation_ID), Nothing)}
    End Function

    Friend Const PatientInsertSql As String =
        "INSERT INTO Patients (Title, Firstname, Surname, Reference, Address1, Address2, Town, County, Postcode, Tel, Mobile, Email, DoB, OrganisationId, LegacySource, LegacyId)
         VALUES (@Title, @Firstname, @Surname, @Reference, @Address1, @Address2, @Town, @County, @Postcode, @Tel, @Mobile, @Email, @DoB, @Org, @LegacySource, @LegacyId);"

    ''' <summary>Adds a patient and returns the new ID.</summary>
    Public Function SavePatient(p As PatientStructure) As Integer
        Return Insert(PatientInsertSql, PatientArgs(p).Concat({"@LegacySource", Nothing, "@LegacyId", Nothing}).ToArray())
    End Function

    Public Sub UpdatePatient(id As Integer, p As PatientStructure)
        Execute("UPDATE Patients SET Title=@Title, Firstname=@Firstname, Surname=@Surname, Reference=@Reference, Address1=@Address1,
                 Address2=@Address2, Town=@Town, County=@County, Postcode=@Postcode, Tel=@Tel, Mobile=@Mobile, Email=@Email, DoB=@DoB,
                 OrganisationId=@Org WHERE Id=@id;", PatientArgs(p).Concat({"@id", CObj(id)}).ToArray())
    End Sub

    Public Function GetPatientById(id As Integer) As PatientStructure?
        Dim dt = Query("SELECT * FROM Patients WHERE Id=@id;", "@id", id)
        If dt.Rows.Count = 0 Then Return Nothing
        Return PatientFrom(dt.Rows(0))
    End Function

    Private Shared Function PatientFrom(r As DataRow) As PatientStructure
        Dim dob = DateOf(Col(r, "DoB"))
        Return New PatientStructure With {
            .ID = Int(Col(r, "Id")), .Title = Str(Col(r, "Title")), .Firstname = Str(Col(r, "Firstname")), .Surname = Str(Col(r, "Surname")),
            .DoB = If(dob = Date.MinValue, Ages.UnknownDoB, dob), .Reference = Str(Col(r, "Reference")),
            .Address1 = Str(Col(r, "Address1")), .Address2 = Str(Col(r, "Address2")), .Town = Str(Col(r, "Town")),
            .County = Str(Col(r, "County")), .Postcode = Str(Col(r, "Postcode")), .Tel = Str(Col(r, "Tel")),
            .Mobile = Str(Col(r, "Mobile")), .Email = Str(Col(r, "Email")), .Organisation_ID = Int(Col(r, "OrganisationId"))}
    End Function

    ''' <summary>
    ''' Patients whose name, reference or organisation contains the text (organisationId -1 = all),
    ''' with their number of tests and the date of the last one. Columns: ID, Title, Firstname,
    ''' Surname, DoB, Reference, Organisation, ResultCount, LastTest.
    ''' </summary>
    Public Function SearchPatients(text As String, Optional organisationId As Integer = -1) As DataTable
        Dim t = "%" & If(text, "").Trim() & "%"
        Dim sql = "SELECT p.Id AS ID, p.Title, p.Firstname, p.Surname, p.DoB, p.Reference, IFNULL(o.Name,'') AS Organisation,
                     COUNT(t.Id) AS ResultCount, MAX(t.TestDate) AS LastTest
                   FROM Patients p
                   LEFT JOIN Organisations o ON o.Id = p.OrganisationId
                   LEFT JOIN Tests t ON t.PatientId = p.Id
                   WHERE (p.Firstname LIKE @t OR p.Surname LIKE @t OR p.Reference LIKE @t
                          OR (p.Firstname || ' ' || p.Surname) LIKE @t OR IFNULL(o.Name,'') LIKE @t)"
        If organisationId >= 0 Then sql &= " AND IFNULL(p.OrganisationId,0) = @org"
        sql &= " GROUP BY p.Id ORDER BY p.Surname COLLATE NOCASE, p.Firstname COLLATE NOCASE;"
        Return Query(sql, "@t", t, "@org", organisationId)
    End Function

    Public Function PatientCount() As Integer
        Return Int(Query("SELECT COUNT(*) FROM Patients;").Rows(0)(0))
    End Function

    ''' <summary>Deletes a patient and (cascade) all their tests. Returns the number of tests deleted.</summary>
    Public Function DeletePatient(id As Integer) As Integer
        Dim n = Int(Query("SELECT COUNT(*) FROM Tests WHERE PatientId=@id;", "@id", id).Rows(0)(0))
        Execute("DELETE FROM Patients WHERE Id=@id;", "@id", id)
        Return n
    End Function

#End Region

#Region "Organisations and screeners"

    Public Function GetOrganisations() As List(Of OrganisationStructure)
        Dim list As New List(Of OrganisationStructure)
        For Each r As DataRow In Query("SELECT * FROM Organisations ORDER BY Name COLLATE NOCASE;").Rows
            list.Add(New OrganisationStructure With {
                .ID = Int(r("Id")), .Name = Str(r("Name")), .Address1 = Str(r("Address1")), .Address2 = Str(r("Address2")),
                .Town = Str(r("Town")), .County = Str(r("County")), .Postcode = Str(r("Postcode")), .Tel = Str(r("Tel"))})
        Next
        Return list
    End Function

    Public Function GetOrganisation(id As Integer) As OrganisationStructure?
        If id <= 0 Then Return Nothing
        For Each o In GetOrganisations()
            If o.ID = id Then Return o
        Next
        Return Nothing
    End Function

    Public Function SaveOrganisation(o As OrganisationStructure) As Integer
        Dim args As Object() = {"@Name", If(o.Name, ""), "@A1", If(o.Address1, ""), "@A2", If(o.Address2, ""), "@Town", If(o.Town, ""),
                                "@County", If(o.County, ""), "@Pc", If(o.Postcode, ""), "@Tel", If(o.Tel, ""), "@id", o.ID}
        If o.ID > 0 Then
            Execute("UPDATE Organisations SET Name=@Name, Address1=@A1, Address2=@A2, Town=@Town, County=@County, Postcode=@Pc, Tel=@Tel WHERE Id=@id;", args)
            Return o.ID
        End If
        Return Insert("INSERT INTO Organisations (Name, Address1, Address2, Town, County, Postcode, Tel) VALUES (@Name, @A1, @A2, @Town, @County, @Pc, @Tel);", args)
    End Function

    ''' <summary>Deletes an organisation; its patients are kept (without an organisation).</summary>
    Public Sub DeleteOrganisation(id As Integer)
        Execute("DELETE FROM Organisations WHERE Id=@id;", "@id", id)
    End Sub

    Public Function GetScreeners() As List(Of ScreenerStructure)
        Dim list As New List(Of ScreenerStructure)
        For Each r As DataRow In Query("SELECT * FROM Screeners WHERE Active=1 ORDER BY Surname COLLATE NOCASE, Firstname COLLATE NOCASE;").Rows
            list.Add(New ScreenerStructure With {.ID = Int(r("Id")), .Title = Str(r("Title")), .Firstname = Str(r("Firstname")), .Surname = Str(r("Surname"))})
        Next
        Return list
    End Function

    ''' <summary>The name of a screener (also of one removed from the list, so old results still show who tested).</summary>
    Public Function ScreenerName(id As Integer) As String
        If id <= 0 Then Return ""
        Dim dt = Query("SELECT Title, Firstname, Surname FROM Screeners WHERE Id=@id;", "@id", id)
        If dt.Rows.Count = 0 Then Return ""
        Return New ScreenerStructure With {.Title = Str(dt.Rows(0)(0)), .Firstname = Str(dt.Rows(0)(1)), .Surname = Str(dt.Rows(0)(2))}.FullName
    End Function

    Public Function SaveScreener(s As ScreenerStructure) As Integer
        Dim args As Object() = {"@T", If(s.Title, ""), "@F", If(s.Firstname, ""), "@S", If(s.Surname, ""), "@id", s.ID}
        If s.ID > 0 Then
            Execute("UPDATE Screeners SET Title=@T, Firstname=@F, Surname=@S WHERE Id=@id;", args)
            Return s.ID
        End If
        Return Insert("INSERT INTO Screeners (Title, Firstname, Surname) VALUES (@T, @F, @S);", args)
    End Function

    ''' <summary>Removes a screener from the list (kept in the database so old results still name them).</summary>
    Public Sub DeleteScreener(id As Integer)
        Execute("UPDATE Screeners SET Active=0 WHERE Id=@id;", "@id", id)
    End Sub

#End Region

#Region "Tests"

    Friend Const TestColumns As String =
        "PatientId, ScreenerId, TestDate, Mode, ColourIndex, ColourSet, ColourRgb, GreyPreferred, Method, AutoScore, " &
        "Symptom1, Symptom2, Symptom3, Symptom4, Symptom5, Symptom6, Symptom7, Symptom8, SymptomsDone, SymptomNotes, " &
        "Wpm, ReadSeconds, WordsRead, Errors, RowsSkipped, Transcript, TextIndex, TextName, FontName, FontSize, LineSpacing, " &
        "RowCount, OverlaySize, PresentationMs, ScreenWidth, ScreenHeight, Notes, LegacySource, LegacyId, " &
        "Stimulus, PatternCpd, PatternDistanceCm"

    ''' <summary>Parameter list for TestColumns (same order).</summary>
    Friend Shared Function TestArgs(r As ReadEZResult, legacySource As String, legacyId As Integer) As Object()
        Dim rgb As Object = Nothing
        If Not r.IsBaseline AndAlso (r.Colour_R + r.Colour_G + r.Colour_B) > 0 Then rgb = (r.Colour_R << 16) Or (r.Colour_G << 8) Or r.Colour_B
        Dim a As New List(Of Object) From {
            "@PatientId", r.Client_ID, "@ScreenerId", If(r.Screener_ID > 0, CObj(r.Screener_ID), Nothing),
            "@TestDate", DateText(If(r.Test_Date = Date.MinValue, Date.Now, r.Test_Date)), "@Mode", If(r.IsBaseline, 0, CInt(r.Mode)),
            "@ColourIndex", r.Colour1, "@ColourSet", Math.Max(1, Math.Min(3, r.PaletteSet)), "@ColourRgb", rgb,
            "@GreyPreferred", r.GreyPreferred, "@Method", r.Screening_Mode, "@AutoScore", If(r.AutoScore, "")}
        For i = 0 To 7
            a.Add("@Symptom" & (i + 1)) : a.Add(Math.Max(0, Math.Min(5, r.Symptoms(i))))
        Next
        a.AddRange({"@SymptomsDone", r.SymptomsDone, "@SymptomNotes", If(r.Symptoms_Notes, ""),
                    "@Wpm", r.ROR, "@ReadSeconds", r.ROR_RecordTime, "@WordsRead", r.ROR_NoWordsRead, "@Errors", r.ROR_NoErrors,
                    "@RowsSkipped", r.ROR_RowsSkipped, "@Transcript", If(r.ROR_Transcript, ""), "@TextIndex", r.RoR_TextOptions,
                    "@TextName", If(r.ROR_Text, ""), "@FontName", If(String.IsNullOrEmpty(r.RoR_FontName), "Times New Roman", r.RoR_FontName),
                    "@FontSize", r.RoR_FontSize, "@LineSpacing", r.RoR_LineSpacing, "@RowCount", r.RoR_NoRows, "@OverlaySize", r.RoR_OverlaySize,
                    "@PresentationMs", r.RoR_PresentationTime, "@ScreenWidth", r.Test_Width, "@ScreenHeight", r.Test_Height,
                    "@Notes", If(r.Notes, ""), "@LegacySource", legacySource, "@LegacyId", If(legacySource Is Nothing, Nothing, CObj(legacyId))})
        a.AddRange({"@Stimulus", If(r.IsPattern, 1, 0), "@PatternCpd", r.PatternCpd, "@PatternDistanceCm", r.PatternDistanceCm})
        Return a.ToArray()
    End Function

    Friend Shared ReadOnly Property TestInsertSql As String
        Get
            Return $"INSERT INTO Tests ({TestColumns}) VALUES ({String.Join(", ", TestColumns.Split(","c).Select(Function(c) "@" & c.Trim()))});"
        End Get
    End Property

    ''' <summary>Saves a test (insert if ID = 0, otherwise update) and returns its ID.</summary>
    Public Function SaveResult(r As ReadEZResult) As Integer
        Dim args = TestArgs(r, Nothing, 0)
        If r.ID > 0 Then
            Dim sets = String.Join(", ", TestColumns.Split(","c).Select(Function(c) c.Trim()).
                                   Where(Function(c) c <> "LegacySource" AndAlso c <> "LegacyId").Select(Function(c) c & "=@" & c))
            Execute($"UPDATE Tests SET {sets} WHERE Id=@id;", args.Concat({"@id", CObj(r.ID)}).ToArray())
            Return r.ID
        End If
        r.ID = Insert(TestInsertSql, args)
        Return r.ID
    End Function

    Public Function LoadResult(id As Integer) As ReadEZResult
        Dim dt = Query("SELECT * FROM Tests WHERE Id=@id;", "@id", id)
        If dt.Rows.Count = 0 Then Return Nothing
        Return ResultFrom(dt.Rows(0))
    End Function

    Private Shared Function ResultFrom(row As DataRow) As ReadEZResult
        Dim mode = Math.Max(0, Math.Min(3, Int(row("Mode"))))
        Dim r As New ReadEZResult With {
            .ID = Int(row("Id")), .IsBaseline = mode = 0, .Mode = CType(mode, ReadEZTestMode),
            .Client_ID = Int(row("PatientId")), .Screener_ID = Int(row("ScreenerId")), .Test_Date = DateOf(row("TestDate")),
            .Colour1 = Int(row("ColourIndex")), .PaletteSet = Math.Max(1, Math.Min(3, Int(row("ColourSet")))),
            .GreyPreferred = Bool(row("GreyPreferred")), .Screening_Mode = Int(row("Method")), .AutoScore = Str(row("AutoScore")),
            .SymptomsDone = Bool(row("SymptomsDone")), .Symptoms_Notes = Str(row("SymptomNotes")),
            .ROR = Int(row("Wpm")), .ROR_RecordTime = Int(row("ReadSeconds")), .ROR_NoWordsRead = Int(row("WordsRead")),
            .ROR_NoErrors = Int(row("Errors")), .ROR_RowsSkipped = Int(row("RowsSkipped")), .ROR_Transcript = Str(row("Transcript")),
            .RoR_TextOptions = Int(row("TextIndex")), .ROR_Text = Str(row("TextName")), .RoR_FontName = Str(row("FontName")),
            .RoR_FontSize = Int(row("FontSize")), .RoR_LineSpacing = Int(row("LineSpacing")), .RoR_NoRows = Int(row("RowCount")),
            .RoR_OverlaySize = Int(row("OverlaySize")), .RoR_PresentationTime = Int(row("PresentationMs")),
            .Test_Width = Int(row("ScreenWidth")), .Test_Height = Int(row("ScreenHeight")), .Notes = Str(row("Notes")),
            .ImportedFrom = Str(row("LegacySource"))}
        If row("ColourRgb") IsNot DBNull.Value Then
            Dim rgb = Int(row("ColourRgb"))
            r.Colour_R = (rgb >> 16) And 255 : r.Colour_G = (rgb >> 8) And 255 : r.Colour_B = rgb And 255
        End If
        For i = 1 To 8
            r.Symptoms(i - 1) = Int(row("Symptom" & i))
        Next
        r.IsPattern = Int(Col(row, "Stimulus")) = 1
        Dim cpd = Col(row, "PatternCpd")
        r.PatternCpd = If(cpd Is Nothing OrElse cpd Is DBNull.Value, 0.0, Convert.ToDouble(cpd, Inv))
        r.PatternDistanceCm = Int(Col(row, "PatternDistanceCm"))
        If r.RoR_NoRows <= 0 Then r.RoR_NoRows = 20
        Return r
    End Function

    ''' <summary>All tests of a patient, newest first (one indexed query).</summary>
    Public Function GetResults(patientId As Integer) As List(Of ReadEZResult)
        Return Query("SELECT * FROM Tests WHERE PatientId=@id ORDER BY TestDate DESC, Id DESC;", "@id", patientId).
               Rows.Cast(Of DataRow)().Select(Function(r) ResultFrom(r)).ToList()
    End Function

    Public Sub DeleteResult(id As Integer)
        Execute("DELETE FROM Tests WHERE Id=@id;", "@id", id)
    End Sub

    Public Sub MoveResult(id As Integer, toPatientId As Integer)
        Execute("UPDATE Tests SET PatientId=@p WHERE Id=@id;", "@p", toPatientId, "@id", id)
    End Sub

    Public Sub UpdateNotes(id As Integer, notes As String)
        Execute("UPDATE Tests SET Notes=@n WHERE Id=@id;", "@n", If(notes, ""), "@id", id)
    End Sub

    ''' <summary>Patients and tests (to check a copy of the database is complete).</summary>
    Public Function Counts() As (Patients As Integer, Recordings As Integer)
        Return (PatientCount(), ResultCount())
    End Function

    Public Function ResultCount() As Integer
        Return Int(Query("SELECT COUNT(*) FROM Tests;").Rows(0)(0))
    End Function

#End Region

End Class
