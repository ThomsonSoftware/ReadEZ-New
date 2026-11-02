Option Strict On

Imports System.Data
Imports System.Data.SQLite

' ====================================================================================================
'   LegacyDatabaseImport.vb - brings a ReadEZ 2020 database (ReadEZ2020.db) into the new structure
'
'   ReadEZ 2020 tables          ->  new tables
'     Organisations                 Organisations   (matched by name; not duplicated)
'     Tester                        Screeners       (matched by name; not duplicated)
'     Clients                       Patients        (LegacySource = file name, LegacyId = old id)
'     Baseline                      Tests, Mode 0
'     Results                       Tests, Mode = Test_Mode (1 overlay, 2 computer, 3 lenses)
'
'   The old file is opened read-only and never changed. Every imported row remembers where it came
'   from (LegacySource + LegacyId, unique), so importing the same file again only adds what is new -
'   patients and tests are never duplicated. Values ReadEZ 2020 stored loosely (numbers as text,
'   dates in several formats, True/False as text) are converted on the way in; a row that cannot be
'   read is skipped and counted, not allowed to stop the import. Everything is written in one
'   transaction: if the import fails, nothing is changed.
' ====================================================================================================

Public NotInheritable Class LegacyDatabaseImport
    Private Sub New()
    End Sub

    Public NotInheritable Class Summary
        Public Patients As Integer
        Public Tests As Integer
        Public Organisations As Integer
        Public Screeners As Integer
        Public AlreadyThere As Integer
        Public Skipped As Integer
    End Class

    ''' <summary>True if the file looks like a ReadEZ 2020 database.</summary>
    Public Shared Function IsLegacyDatabase(path As String) As Boolean
        Return ReadEZDatabase.IsLegacyFile(path)
    End Function

    ''' <summary>The ReadEZ 2020 database named in an old RZSettings.ini, if there is one.</summary>
    Public Shared Function FindOldDatabase() As String
        For Each ini In LegacyImport.CandidateIniFiles()
            Try
                If Not IO.File.Exists(ini) Then Continue For
                For Each line In IO.File.ReadAllLines(ini)
                    If line.StartsWith("Local_Database=", StringComparison.OrdinalIgnoreCase) Then
                        Dim f = line.Substring(15).Trim()
                        If f <> "" AndAlso IsLegacyDatabase(f) Then Return f
                    End If
                Next
                Dim folder = IO.Path.GetDirectoryName(ini)
                For Each name In {"ReadEZ2020.db", IO.Path.Combine("RZ Database", "ReadEZ.db"), "ReadEZ.db"}
                    Dim f = IO.Path.Combine(folder, name)
                    If IsLegacyDatabase(f) Then Return f
                Next
            Catch ex As Exception
                Log.Warn("Looking for ReadEZ 2020 database: " & ex.Message)
            End Try
        Next
        Return Nothing
    End Function

    ''' <summary>Imports everything from a ReadEZ 2020 database into target.</summary>
    Public Shared Function Import(oldFile As String, target As ReadEZDatabase, Optional progress As Action(Of String) = Nothing) As Summary
        Dim sum As New Summary
        Dim source = IO.Path.GetFileName(oldFile)
        Dim oldTables = ReadAll(oldFile)

        Using conn = target.Open(), tx = conn.BeginTransaction()
            ' ---- Organisations and screeners (by name) ----
            Dim orgMap As New Dictionary(Of Integer, Integer)
            Dim existingOrgs = target.GetOrganisations()
            For Each r As DataRow In Rows(oldTables, "Organisations")
                Dim name = ReadEZDatabase.Str(Value(r, "Name")).Trim()
                If name = "" Then Continue For
                Dim found = existingOrgs.FirstOrDefault(Function(o) String.Equals(o.Name, name, StringComparison.CurrentCultureIgnoreCase))
                If found.ID > 0 Then
                    orgMap(ReadEZDatabase.Int(Value(r, "id"))) = found.ID
                Else
                    Dim id = ReadEZDatabase.InsertOn(conn, tx, "INSERT INTO Organisations (Name, Address1, Address2, Town, County, Postcode, Tel, LegacyId) VALUES (@n,@a1,@a2,@t,@c,@p,@tel,@lid);",
                        "@n", name, "@a1", S(r, "Address1"), "@a2", S(r, "Address2"), "@t", S(r, "Town"), "@c", S(r, "County"),
                        "@p", S(r, "Postcode"), "@tel", S(r, "Tel"), "@lid", ReadEZDatabase.Int(Value(r, "id")))
                    orgMap(ReadEZDatabase.Int(Value(r, "id"))) = id
                    existingOrgs.Add(New OrganisationStructure With {.ID = id, .Name = name})
                    sum.Organisations += 1
                End If
            Next

            Dim screenerMap As New Dictionary(Of Integer, Integer)
            Dim existingScreeners = target.GetScreeners()
            For Each r As DataRow In Rows(oldTables, "Tester")
                Dim sc As New ScreenerStructure With {.Title = S(r, "Title"), .Firstname = S(r, "Firstname"), .Surname = S(r, "Surname")}
                If sc.FullName = "" Then Continue For
                Dim found = existingScreeners.FirstOrDefault(Function(x) String.Equals(x.FullName, sc.FullName, StringComparison.CurrentCultureIgnoreCase))
                If found.ID > 0 Then
                    screenerMap(ReadEZDatabase.Int(Value(r, "id"))) = found.ID
                Else
                    Dim id = ReadEZDatabase.InsertOn(conn, tx, "INSERT INTO Screeners (Title, Firstname, Surname, LegacyId) VALUES (@t,@f,@s,@lid);",
                        "@t", sc.Title, "@f", sc.Firstname, "@s", sc.Surname, "@lid", ReadEZDatabase.Int(Value(r, "id")))
                    screenerMap(ReadEZDatabase.Int(Value(r, "id"))) = id
                    sc.ID = id
                    existingScreeners.Add(sc)
                    sum.Screeners += 1
                End If
            Next

            ' ---- Patients ----
            progress?.Invoke(Loc.T("import.progress.patients"))
            Dim patientMap = ExistingLegacyIds(conn, tx, "Patients", source)
            For Each r As DataRow In Rows(oldTables, "Clients")
                Dim oldId = ReadEZDatabase.Int(Value(r, "id"))
                If patientMap.ContainsKey(oldId) Then sum.AlreadyThere += 1 : Continue For
                Try
                    Dim dob = ReadEZDatabase.DateOf(Value(r, "DoB"))
                    Dim org = 0
                    orgMap.TryGetValue(ReadEZDatabase.Int(Value(r, "Organisation_ID")), org)
                    Dim p As New PatientStructure With {
                        .Title = S(r, "Title"), .Firstname = S(r, "Firstname"), .Surname = S(r, "Surname"),
                        .DoB = If(dob = Date.MinValue, Ages.UnknownDoB, dob), .Reference = S(r, "Reference"),
                        .Address1 = S(r, "Address1"), .Address2 = S(r, "Address2"), .Town = S(r, "Town"), .County = S(r, "County"),
                        .Postcode = S(r, "Postcode"), .Tel = S(r, "Tel"), .Mobile = S(r, "Mobile"), .Email = S(r, "Email"), .Organisation_ID = org}
                    Dim args = ReadEZDatabase.PatientArgs(p).Concat({"@LegacySource", CObj(source), "@LegacyId", CObj(oldId)}).ToArray()
                    patientMap(oldId) = ReadEZDatabase.InsertOn(conn, tx, ReadEZDatabase.PatientInsertSql, args)
                    sum.Patients += 1
                Catch ex As Exception
                    Log.Warn($"Import patient {oldId}: {ex.Message}")
                    sum.Skipped += 1
                End Try
            Next

            ' ---- Tests ----
            progress?.Invoke(Loc.T("import.progress.tests"))
            For Each baseline In {True, False}
                Dim tag = source & If(baseline, "#Baseline", "#Results")
                Dim done = ExistingLegacyIds(conn, tx, "Tests", tag)
                For Each r As DataRow In Rows(oldTables, If(baseline, "Baseline", "Results"))
                    Dim oldId = ReadEZDatabase.Int(Value(r, "id"))
                    If done.ContainsKey(oldId) Then sum.AlreadyThere += 1 : Continue For
                    Dim patientId = 0
                    If Not patientMap.TryGetValue(ReadEZDatabase.Int(Value(r, "Client_ID")), patientId) Then sum.Skipped += 1 : Continue For
                    Try
                        Dim t = TestFrom(r, baseline)
                        t.Client_ID = patientId
                        Dim sid = 0
                        t.Screener_ID = If(screenerMap.TryGetValue(ReadEZDatabase.Int(Value(r, "Screener_ID")), sid), sid, 0)
                        ReadEZDatabase.InsertOn(conn, tx, ReadEZDatabase.TestInsertSql, ReadEZDatabase.TestArgs(t, tag, oldId))
                        sum.Tests += 1
                    Catch ex As Exception
                        Log.Warn($"Import {(If(baseline, "baseline", "result"))} {oldId}: {ex.Message}")
                        sum.Skipped += 1
                    End Try
                Next
            Next
            tx.Commit()
        End Using
        target.SetMeta("ImportedFrom." & source, ReadEZDatabase.DateText(Date.Now))
        Log.Info($"Imported {oldFile}: {sum.Patients} patients, {sum.Tests} tests, {sum.AlreadyThere} already there, {sum.Skipped} skipped")
        Return sum
    End Function

    ''' <summary>A ReadEZ 2020 Baseline or Results row as a test.</summary>
    Private Shared Function TestFrom(r As DataRow, baseline As Boolean) As ReadEZResult
        Dim t As New ReadEZResult With {
            .IsBaseline = baseline, .Test_Date = ReadEZDatabase.DateOf(Value(r, "Test_Date")),
            .Symptoms_Notes = S(r, "Symptoms_Notes"),
            .ROR = I(r, "ROR"), .ROR_RecordTime = I(r, "ROR_RecordTime"), .ROR_NoErrors = I(r, "ROR_NoErrors"),
            .ROR_RowsSkipped = I(r, "ROR_RowsSkipped"), .ROR_NoWordsRead = I(r, "ROR_NoWordsRead"),
            .ROR_Transcript = S(r, "ROR_Transcript"),
            .Test_Width = I(r, "Test_Width"), .Test_Height = I(r, "Test_Height"),
            .RoR_LineSpacing = I(r, "RoR_LineSpacing"), .RoR_FontSize = I(r, "RoR_FontSize"), .RoR_FontName = S(r, "RoR_FontName"),
            .RoR_NoRows = I(r, "RoR_NoRows"), .RoR_TextOptions = I(r, "RoR_TextOptions"), .ImportedFrom = "ReadEZ 2020"}
        If t.Test_Date = Date.MinValue Then t.Test_Date = New Date(2000, 1, 1)
        t.ROR_Text = ReadingTexts.BuiltInName(t.RoR_TextOptions)
        For k = 1 To 8
            t.Symptoms(k - 1) = Math.Max(0, Math.Min(5, I(r, "Symptoms" & k)))
        Next
        t.SymptomsDone = t.Symptoms.Any(Function(x) x > 0) OrElse t.Symptoms_Notes <> ""
        If t.RoR_NoRows <= 0 Then t.RoR_NoRows = 20
        ' The rate was stored, but recompute it if only the words and time were
        If t.ROR = 0 AndAlso t.ROR_RecordTime > 0 Then t.UpdateRate()
        If baseline Then
            t.Mode = ReadEZTestMode.Baseline
        Else
            Dim m = I(r, "Test_Mode")
            t.Mode = If(m >= 1 AndAlso m <= 3, CType(m, ReadEZTestMode), ReadEZTestMode.Overlay)
            t.Colour1 = Math.Max(0, Math.Min(12, I(r, "Colour1")))
            t.PaletteSet = Math.Max(1, Math.Min(3, I(r, "PaletteSet")))
            t.Colour_R = Clamp255(I(r, "Colour_R")) : t.Colour_G = Clamp255(I(r, "Colour_G")) : t.Colour_B = Clamp255(I(r, "Colour_B"))
            t.GreyPreferred = ReadEZDatabase.Bool(Value(r, "GreyPreferred"))
            t.Screening_Mode = I(r, "Screening_Mode")
            t.RoR_OverlaySize = I(r, "RoR_OverlaySize")
            t.RoR_PresentationTime = I(r, "RoR_PresentationTime")
            t.AutoScore = S(r, "Autoscore")
        End If
        Return t
    End Function

    Private Shared Function Clamp255(v As Integer) As Integer
        Return Math.Max(0, Math.Min(255, v))
    End Function

#Region "Reading the old file"

    Private Shared Function ReadAll(file As String) As Dictionary(Of String, DataTable)
        Dim d As New Dictionary(Of String, DataTable)(StringComparer.OrdinalIgnoreCase)
        ' Read-only: the old database is never changed. Dates are read as text and converted here.
        Using conn As New SQLiteConnection($"Data Source={file};Version=3;Read Only=True;DateTimeFormat=Ticks;")
            conn.Open()
            Dim tables As New List(Of String)
            Using cmd As New SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", conn), r = cmd.ExecuteReader()
                While r.Read()
                    tables.Add(ReadEZDatabase.Str(r(0)))
                End While
            End Using
            For Each t In {"Organisations", "Tester", "Clients", "Baseline", "Results"}
                Dim actual = tables.FirstOrDefault(Function(x) String.Equals(x, t, StringComparison.OrdinalIgnoreCase))
                If actual Is Nothing Then Continue For
                ' Every column as text (CAST), so a value of an unexpected type can't stop the read
                Dim cols As New List(Of String)
                Using cmd As New SQLiteCommand($"PRAGMA table_info([{actual}]);", conn), r = cmd.ExecuteReader()
                    While r.Read()
                        cols.Add(ReadEZDatabase.Str(r("name")))
                    End While
                End Using
                Dim sql = $"SELECT {String.Join(", ", cols.Select(Function(c) $"CAST([{c}] AS TEXT) AS [{c}]"))} FROM [{actual}];"
                Using cmd As New SQLiteCommand(sql, conn), r = cmd.ExecuteReader()
                    Dim dt As New DataTable()
                    dt.Load(r)
                    d(t) = dt
                End Using
            Next
        End Using
        Return d
    End Function

    Private Shared Function Rows(tables As Dictionary(Of String, DataTable), name As String) As IEnumerable(Of DataRow)
        Dim dt As DataTable = Nothing
        If tables.TryGetValue(name, dt) Then Return dt.Rows.Cast(Of DataRow)()
        Return Enumerable.Empty(Of DataRow)()
    End Function

    Private Shared Function Value(r As DataRow, col As String) As Object
        For Each c As DataColumn In r.Table.Columns
            If String.Equals(c.ColumnName, col, StringComparison.OrdinalIgnoreCase) Then Return r(c)
        Next
        Return Nothing
    End Function

    Private Shared Function S(r As DataRow, col As String) As String
        Return ReadEZDatabase.Str(Value(r, col)).Trim()
    End Function

    Private Shared Function I(r As DataRow, col As String) As Integer
        Return ReadEZDatabase.Int(Value(r, col))
    End Function

    ''' <summary>Old ids already imported from this source (old id -> new id).</summary>
    Private Shared Function ExistingLegacyIds(conn As SQLiteConnection, tx As SQLiteTransaction, table As String, source As String) As Dictionary(Of Integer, Integer)
        Dim d As New Dictionary(Of Integer, Integer)
        Using cmd As New SQLiteCommand($"SELECT LegacyId, Id FROM {table} WHERE LegacySource=@s;", conn, tx)
            cmd.Parameters.AddWithValue("@s", source)
            Using r = cmd.ExecuteReader()
                While r.Read()
                    d(ReadEZDatabase.Int(r(0))) = ReadEZDatabase.Int(r(1))
                End While
            End Using
        End Using
        Return d
    End Function

#End Region
End Class
