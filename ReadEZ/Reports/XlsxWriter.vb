Option Strict On

Imports System.IO.Compression
Imports System.Security
Imports System.Text

' ====================================================================================================
'   XlsxWriter.vb - writes a simple Excel workbook (.xlsx) without Excel being installed
'
'   The previous export drove Excel through COM interop, one cell at a time: it needed Excel on
'   the PC, was very slow on long recordings and swallowed its errors. This writes the Office
'   Open XML files directly: text cells (inline strings), numbers, an optional bold row per sheet,
'   and frozen header rows.
' ====================================================================================================

Public NotInheritable Class XlsxWriter

    Public NotInheritable Class Sheet
        Public Property Name As String
        Public ReadOnly Property Rows As New List(Of Object())
        ''' <summary>Rows (0-based) shown in bold.</summary>
        Public ReadOnly Property BoldRows As New HashSet(Of Integer)
        ''' <summary>Rows above this are frozen (0 = none).</summary>
        Public Property FreezeRows As Integer

        Public Sub AddRow(ParamArray cells As Object())
            Rows.Add(cells)
        End Sub

        Public Sub AddBoldRow(ParamArray cells As Object())
            BoldRows.Add(Rows.Count)
            Rows.Add(cells)
        End Sub
    End Class

    Public ReadOnly Property Sheets As New List(Of Sheet)

    Public Function AddSheet(name As String) As Sheet
        Dim clean = New String(name.Where(Function(c) Not "[]:*?/\".Contains(c)).ToArray())
        Dim s As New Sheet With {.Name = If(clean.Length > 31, clean.Substring(0, 31), clean)}
        Sheets.Add(s)
        Return s
    End Function

    Public Sub Save(path As String)
        If IO.File.Exists(path) Then IO.File.Delete(path)
        Using zip = ZipFile.Open(path, ZipArchiveMode.Create)
            Put(zip, "[Content_Types].xml", ContentTypes())
            Put(zip, "_rels/.rels",
                "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/></Relationships>")
            Put(zip, "xl/workbook.xml", Workbook())
            Put(zip, "xl/_rels/workbook.xml.rels", WorkbookRels())
            Put(zip, "xl/styles.xml",
                "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><fonts count=""2""><font><sz val=""11""/><name val=""Calibri""/></font><font><b/><sz val=""11""/><name val=""Calibri""/></font></fonts><fills count=""2""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill></fills><borders count=""1""><border/></borders><cellStyleXfs count=""1""><xf/></cellStyleXfs><cellXfs count=""2""><xf fontId=""0""/><xf fontId=""1"" applyFont=""1""/></cellXfs></styleSheet>")
            For i = 0 To Sheets.Count - 1
                Dim entry = zip.CreateEntry($"xl/worksheets/sheet{i + 1}.xml", CompressionLevel.Optimal)
                Using st = entry.Open(), w As New IO.StreamWriter(st, New UTF8Encoding(False))
                    WriteSheet(w, Sheets(i))
                End Using
            Next
        End Using
    End Sub

    Private Shared Sub Put(zip As ZipArchive, name As String, text As String)
        Dim entry = zip.CreateEntry(name, CompressionLevel.Optimal)
        Using st = entry.Open(), w As New IO.StreamWriter(st, New UTF8Encoding(False))
            w.Write(text)
        End Using
    End Sub

    Private Function ContentTypes() As String
        Dim sb As New StringBuilder("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/><Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/><Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>")
        For i = 1 To Sheets.Count
            sb.Append($"<Override PartName=""/xl/worksheets/sheet{i}.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>")
        Next
        sb.Append("</Types>")
        Return sb.ToString()
    End Function

    Private Function Workbook() As String
        Dim sb As New StringBuilder("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets>")
        For i = 0 To Sheets.Count - 1
            sb.Append($"<sheet name=""{SecurityElement.Escape(Sheets(i).Name)}"" sheetId=""{i + 1}"" r:id=""rId{i + 1}""/>")
        Next
        sb.Append("</sheets></workbook>")
        Return sb.ToString()
    End Function

    Private Function WorkbookRels() As String
        Dim sb As New StringBuilder("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">")
        For i = 0 To Sheets.Count - 1
            sb.Append($"<Relationship Id=""rId{i + 1}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet{i + 1}.xml""/>")
        Next
        sb.Append($"<Relationship Id=""rId{Sheets.Count + 1}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/></Relationships>")
        Return sb.ToString()
    End Function

    Private Shared Function ColName(c As Integer) As String
        Dim s = ""
        c += 1
        While c > 0
            Dim m = (c - 1) Mod 26
            s = ChrW(65 + m) & s
            c = (c - m - 1) \ 26
        End While
        Return s
    End Function

    Private Shared Sub WriteSheet(w As IO.StreamWriter, s As Sheet)
        w.Write("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">")
        If s.FreezeRows > 0 Then
            w.Write($"<sheetViews><sheetView workbookViewId=""0""><pane ySplit=""{s.FreezeRows}"" topLeftCell=""A{s.FreezeRows + 1}"" activePane=""bottomLeft"" state=""frozen""/></sheetView></sheetViews>")
        End If
        Dim widest = If(s.Rows.Count = 0, 0, s.Rows.Max(Function(r) r.Length))
        If widest > 0 Then w.Write($"<cols><col min=""1"" max=""{widest}"" width=""16"" customWidth=""1""/></cols>")
        w.Write("<sheetData>")
        Dim inv = Globalization.CultureInfo.InvariantCulture
        For r = 0 To s.Rows.Count - 1
            Dim row = s.Rows(r)
            Dim style = If(s.BoldRows.Contains(r), " s=""1""", "")
            w.Write($"<row r=""{r + 1}"">")
            For c = 0 To row.Length - 1
                Dim v = row(c)
                If v Is Nothing Then Continue For
                Dim refr = ColName(c) & (r + 1)
                Select Case True
                    Case TypeOf v Is Boolean
                        w.Write($"<c r=""{refr}""{style}><v>{If(CBool(v), 1, 0)}</v></c>")
                    Case TypeOf v Is Integer, TypeOf v Is Long, TypeOf v Is Short, TypeOf v Is Single, TypeOf v Is Double, TypeOf v Is Decimal
                        Dim d = Convert.ToDouble(v)
                        If Double.IsNaN(d) OrElse Double.IsInfinity(d) Then Continue For
                        ' A Single is written with its own shortest form (0.1, not 0.10000000149011612)
                        Dim text = If(TypeOf v Is Single, CSng(v).ToString("R", inv), d.ToString("R", inv))
                        w.Write($"<c r=""{refr}""{style}><v>{text}</v></c>")
                    Case Else
                        w.Write($"<c r=""{refr}"" t=""inlineStr""{style}><is><t xml:space=""preserve"">{SecurityElement.Escape(XmlSafe(v.ToString()))}</t></is></c>")
                End Select
            Next
            w.Write("</row>")
        Next
        w.Write("</sheetData></worksheet>")
    End Sub

    ''' <summary>
    ''' Removes characters XML can't contain (control characters, lone surrogates), which make
    ''' Excel report the file as corrupt - e.g. a vertical tab pasted into notes.
    ''' </summary>
    Private Shared Function XmlSafe(s As String) As String
        Dim sb As Text.StringBuilder = Nothing
        For i = 0 To s.Length - 1
            Dim ch = s(i)
            Dim ok As Boolean
            If Char.IsHighSurrogate(ch) AndAlso i + 1 < s.Length AndAlso Char.IsLowSurrogate(s(i + 1)) Then
                ok = True
            ElseIf Char.IsLowSurrogate(ch) AndAlso i > 0 AndAlso Char.IsHighSurrogate(s(i - 1)) Then
                ok = True
            Else
                ok = Xml.XmlConvert.IsXmlChar(ch)
            End If
            If ok Then
                sb?.Append(ch)
            ElseIf sb Is Nothing Then
                sb = New Text.StringBuilder(s, 0, i, s.Length)
            End If
        Next
        Return If(sb Is Nothing, s, sb.ToString())
    End Function
End Class
