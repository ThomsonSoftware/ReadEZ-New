Option Strict On

Imports DevExpress.Drawing
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraPrintingLinks

' ====================================================================================================
'   ReportBuilder.vb - every printed report is built with this (DevExpress XtraPrinting)
'
'   The previous version copied the same header, footer and patient-details code into every
'   results form. Here it is written once:
'     header   report title (right), practice name, logo on the left (Settings.ShowLogo,
'              LogoFilename, LogoScale) - on every page
'     footer   practice name · "ReadEZ", print date, "Page n of m"
'     body     patient / test information boxes, section headings, pictures (charts drawn by ReadEZ)
'              (aspect ratio kept), text boxes and tables, added in order
'   Style: the Thomson Software Solutions colours (Brand.vb) - title and a rule in brand blue in the
'   header, section headings with a blue bar, shaded box and table headers, and key-result panels
'   (AddKeyResult) in green / amber / red / blue for a test's main result. Segoe UI throughout.
'   A4 portrait. The DevExpress preview window provides Print, Export to PDF / Word / Excel /
'   image and Send by e-mail (as in Eye Dynamix).
' ====================================================================================================

Public NotInheritable Class ReportBuilder
    Implements IDisposable

    Private ReadOnly _ps As New PrintingSystem()
    Private ReadOnly _composite As CompositeLink
    Private ReadOnly _title As String
    Private ReadOnly _keepAlive As New List(Of IDisposable)
    Private _logo As Image

    Private Shared ReadOnly FontTitle As New Font("Segoe UI Semibold", 16)
    Private Shared ReadOnly FontSection As New Font("Segoe UI Semibold", 12)
    Private Shared ReadOnly FontHead As New Font("Segoe UI Semibold", 9)
    Private Shared ReadOnly FontBody As New Font("Segoe UI", 9.5F)
    Private Shared ReadOnly FontSmall As New Font("Segoe UI", 8)
    Private Shared ReadOnly FontLabel As New Font("Segoe UI Semibold", 8)
    Private Shared ReadOnly FontKey As New Font("Segoe UI Semibold", 15)
    Private Shared ReadOnly Ink As Color = Brand.Charcoal
    Private Shared ReadOnly Rule As Color = Brand.TintBorder
    Private Shared ReadOnly Shade As Color = Brand.Tint

    Public Sub New(titleKey As String)
        _title = Loc.T(titleKey).ToUpper()
        _composite = New CompositeLink(_ps)
        _composite.PaperKind = DevExpress.Drawing.Printing.DXPaperKind.A4
        _composite.Landscape = False
        _composite.Margins = New System.Drawing.Printing.Margins(50, 50, 110, 60)
        LoadLogo()
        AddHandler _composite.CreateMarginalHeaderArea, AddressOf DrawHeader
        AddHandler _composite.CreateMarginalFooterArea, AddressOf DrawFooter
    End Sub

#Region "Header and footer"

    Private Sub LoadLogo()
        Dim s = AppHost.Settings
        If Not s.ShowLogo Then Return
        Dim file = ResolveLogoFile(s.LogoFilename)
        If file = "" Then
            Log.Warn("Report logo not found: " & s.LogoFilename)
            Return
        End If
        Try
            Using fs = IO.File.OpenRead(file), img = Image.FromStream(fs)
                ' Redrawn at 96 dpi on white: a picture saved at another resolution, or with a
                ' transparent background, then prints at the size and colour expected.
                Dim bmp As New Bitmap(img.Width, img.Height)
                bmp.SetResolution(96, 96)
                Using gr = Graphics.FromImage(bmp)
                    gr.Clear(Color.White)
                    gr.DrawImage(img, 0, 0, img.Width, img.Height)
                End Using
                _logo = bmp
            End Using
        Catch ex As Exception
            Log.Error("Report logo", ex)
        End Try
    End Sub

    ''' <summary>
    ''' The logo file from Settings. If it is not where Settings says (a relative name, or settings
    ''' from the previous version), the "Report images" folders of the program and of the data
    ''' folder are tried as well.
    ''' </summary>
    Private Shared Function ResolveLogoFile(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return ""
        name = name.Trim().Trim(""""c)
        If IO.File.Exists(name) Then Return name
        Dim justName = IO.Path.GetFileName(name)
        For Each candidate In {IO.Path.Combine(AppHost.AppFolder, name),
                               IO.Path.Combine(AppHost.DefaultDataFolder, "Report images", justName),
                               IO.Path.Combine(AppHost.AppFolder, "Report images", justName)}
            If IO.File.Exists(candidate) Then Return candidate
        Next
        Return ""
    End Function

    Private Sub DrawHeader(sender As Object, e As CreateAreaEventArgs)
        Dim g = e.Graph
        Dim w = g.ClientPageSize.Width
        g.BackColor = Color.White

        ' A dark blue band across the top: the logo on a white tile at the left, the title and the
        ' practice name in white at the right; a light blue line under it.
        Const BandH As Single = 78
        g.DrawRect(New RectangleF(0, 0, w, BandH), BorderSide.None, Brand.Navy, Brand.Navy)
        g.DrawRect(New RectangleF(0, BandH, w, 3), BorderSide.None, Brand.LightBlue, Brand.LightBlue)
        Dim textLeft = 16.0F
        If _logo IsNot Nothing Then
            Dim maxH = BandH - 22
            Dim lw = 150.0F * Math.Max(5, AppHost.Settings.LogoScale) / 100.0F
            Dim lh = lw * _logo.Height / Math.Max(1, _logo.Width)
            If lh > maxH Then lw *= maxH / lh : lh = maxH
            If lw > w / 2 Then lh *= (w / 2) / lw : lw = w / 2
            Dim tile As New RectangleF(10, (BandH - lh) / 2 - 5, lw + 10, lh + 10)
            g.DrawRect(tile, BorderSide.None, Color.White, Color.White)
            Dim ib = g.DrawImage(_logo, New RectangleF(tile.X + 5, tile.Y + 5, lw, lh), BorderSide.None, Color.White)
            ' The default (Normal) printed the picture at full size, clipped to this box, so only
            ' its top-left corner - usually white - appeared.
            ib.SizeMode = ImageSizeMode.ZoomImage
            textLeft = tile.Right + 14
        End If

        g.BackColor = Color.Transparent
        g.Font = FontTitle
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Far, DXStringAlignment.Center)
        Dim practice = AppHost.Settings.Practice_Name
        Dim hasPractice = Not String.IsNullOrWhiteSpace(practice)
        g.DrawString(_title, Color.White, New RectangleF(textLeft, If(hasPractice, 10, 0), w - textLeft - 16, If(hasPractice, 32, BandH)), BorderSide.None)
        If hasPractice Then
            g.Font = FontBody
            g.DrawString(practice, Color.FromArgb(205, 222, 242), New RectangleF(textLeft, 44, w - textLeft - 16, 22), BorderSide.None)
        End If
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
        g.BackColor = Color.White
    End Sub

    Private Sub DrawFooter(sender As Object, e As CreateAreaEventArgs)
        Dim g = e.Graph
        Dim w = g.ClientPageSize.Width
        g.Font = FontSmall
        g.BackColor = Color.White
        g.DrawRect(New RectangleF(0, 2, w, 1), BorderSide.None, Brand.TintBorder, Brand.TintBorder)
        Dim practice = AppHost.Settings.Practice_Name
        Dim centre = If(String.IsNullOrWhiteSpace(practice), AppHost.ProductTitle, practice & "  ·  " & AppHost.ProductTitle)
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Center)
        g.DrawString(centre, Brand.Grey, New RectangleF(0, 10, w, 16), BorderSide.None)
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
        g.DrawPageInfo(PageInfo.DateTime, "{0:g}", Brand.Grey, New RectangleF(0, 10, w / 3, 16), BorderSide.None)
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Far)
        g.DrawPageInfo(PageInfo.NumberOfTotal, Loc.T("report.pageOf"), Brand.Grey, New RectangleF(w * 2 / 3, 10, w / 3, 16), BorderSide.None)
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
    End Sub

#End Region

#Region "Body parts"

    ''' <summary>Free text (a letter paragraph), wrapped to the page width.</summary>
    Public Sub AddParagraph(text As String)
        If String.IsNullOrWhiteSpace(text) Then Return
        AddLink(Sub(g, w)
                    g.Font = FontBody
                    g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
                    Dim h = TextHeight(text, FontBody, w)
                    g.DrawString(text, Ink, New RectangleF(0, 8, w, h), BorderSide.None)
                End Sub)
    End Sub

    Private Sub AddLink(draw As Action(Of BrickGraphics, Single))
        Dim l As New Link(_ps)
        AddHandler l.CreateDetailArea, Sub(s, e)
                                           e.Graph.BackColor = Color.White
                                           draw(e.Graph, e.Graph.ClientPageSize.Width)
                                       End Sub
        _composite.Links.Add(l)
    End Sub

    ''' <summary>Two boxes: patient details and test information (notes included).</summary>
    Public Sub AddPatientBlock(patientId As Integer, testLines As IEnumerable(Of String))
        Dim p = AppHost.Db.GetPatientById(patientId)
        Dim left As New List(Of String)
        If p.HasValue Then
            Dim pv = p.Value
            left.Add(Loc.T("report.name") & ": " & String.Join(" ", {pv.Title, pv.Firstname, pv.Surname}.Where(Function(x) Not String.IsNullOrWhiteSpace(x))))
            If Not String.IsNullOrWhiteSpace(pv.Reference) Then left.Add(Loc.T("report.reference") & ": " & pv.Reference)
            If Ages.IsKnown(pv.DoB) Then left.Add(Loc.T("report.dob") & ": " & pv.DoB.ToShortDateString() & "  (" & Loc.F("report.age", Ages.YearsOld(pv.DoB)) & ")")
            Dim org = AppHost.Db.GetOrganisation(pv.Organisation_ID)
            If org.HasValue Then left.Add(Loc.T("report.organisation") & ": " & org.Value.Name)
        End If
        Dim right = testLines.ToList()
        AddLink(Sub(g, w)
                    Dim half = (w - 12) / 2
                    Dim lines = Math.Max(left.Count, right.Count)
                    ' Measured, so long or multi-line notes are printed in full instead of clipped
                    Dim h = Math.Max(14 + lines * 16, Math.Max(BodyHeight(left, half), BodyHeight(right, half)))
                    Box(g, New RectangleF(0, 4, half, 22 + h), Loc.T("report.patientDetails"), left)
                    Box(g, New RectangleF(half + 12, 4, half, 22 + h), Loc.T("report.testInformation"), right)
                End Sub)
    End Sub

    ''' <summary>A titled box: the title in brand blue capitals on a light blue band, the lines below.</summary>
    Private Shared Sub Box(g As BrickGraphics, r As RectangleF, title As String, lines As IEnumerable(Of String))
        g.Font = FontHead
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Near, DXStringAlignment.Center)
        Dim hb = g.DrawString("   " & title.ToUpper(), Brand.Blue, New RectangleF(r.X, r.Y, r.Width, 22), BorderSide.All)
        hb.BackColor = Shade
        hb.BorderColor = Rule
        g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
        g.Font = FontBody
        Dim body = "  " & String.Join(Environment.NewLine & "  ", lines)
        Dim tb = g.DrawString(Environment.NewLine & body, Ink, New RectangleF(r.X, r.Y + 22, r.Width, r.Height - 22), BorderSide.All)
        tb.BorderColor = Rule
    End Sub

    ''' <summary>A section heading across the page.</summary>
    Public Sub AddSection(text As String)
        AddLink(Sub(g, w)
                    ' A blue bar, the heading in brand blue, and a light rule under it
                    g.DrawRect(New RectangleF(0, 16, 4, 22), BorderSide.None, Brand.Blue, Brand.Blue)
                    g.Font = FontSection
                    g.StringFormat = New BrickStringFormat(DXStringAlignment.Near, DXStringAlignment.Center)
                    g.DrawString(text, Brand.Blue, New RectangleF(12, 14, w - 12, 26), BorderSide.None)
                    g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
                    g.DrawRect(New RectangleF(0, 42, w, 1), BorderSide.None, Rule, Rule)
                End Sub)
    End Sub

    ''' <summary>
    ''' The main result of a test as a tinted panel with a coloured edge (green good, amber worth a
    ''' look, red a concern, blue information): label in small capitals, the result large, a note.
    ''' </summary>
    Public Sub AddKeyResult(label As String, value As String, note As String, lvl As Brand.Level)
        AddLink(Sub(g, w)
                    Dim edge = Brand.LevelColour(lvl)
                    Dim textW = w - 34
                    Dim valueH = TextHeight(value, FontKey, textW)
                    Dim noteH = If(String.IsNullOrEmpty(note), 0.0F, TextHeight(note, FontBody, textW))
                    Dim y0 = 12.0F
                    Dim h = 12 + 16 + valueH + If(noteH > 0, 2 + noteH, 0.0F) + 10
                    g.DrawRect(New RectangleF(0, y0, w, h), BorderSide.All, Brand.LevelTint(lvl), edge)
                    g.DrawRect(New RectangleF(0, y0, 6, h), BorderSide.None, edge, edge)
                    g.BackColor = Color.Transparent
                    g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
                    g.Font = FontLabel
                    g.DrawString(label.ToUpper(), edge, New RectangleF(20, y0 + 10, textW, 16), BorderSide.None)
                    g.Font = FontKey
                    g.DrawString(value, Ink, New RectangleF(20, y0 + 26, textW, valueH), BorderSide.None)
                    If noteH > 0 Then
                        g.Font = FontBody
                        g.DrawString(note, Brand.Grey, New RectangleF(20, y0 + 28 + valueH, textW, noteH), BorderSide.None)
                    End If
                    g.BackColor = Color.White
                End Sub)
    End Sub

    ''' <summary>Height (report pixels, 96 dpi) a piece of text needs at this width.</summary>
    Private Shared Function TextHeight(text As String, f As Font, width As Single) As Single
        Using bmp As New Bitmap(1, 1)
            bmp.SetResolution(96, 96)
            Using gr = Graphics.FromImage(bmp)
                Return gr.MeasureString(text, f, Math.Max(1, CInt(width) - 4)).Height + 2
            End Using
        End Using
    End Function

    ''' <summary>A picture at the page width (or maxHeight), aspect ratio kept, centred.</summary>
    Public Sub AddImage(img As Image, Optional maxHeight As Single = 420, Optional border As Boolean = True)
        _keepAlive.Add(img)
        AddLink(Sub(g, w)
                    Dim iw = CSng(w), ih = iw * img.Height / Math.Max(1, img.Width)
                    If ih > maxHeight Then iw *= maxHeight / ih : ih = maxHeight
                    Dim ib = g.DrawImage(img, New RectangleF((w - iw) / 2, 8, iw, ih), If(border, BorderSide.All, BorderSide.None), Color.White)
                    ib.BorderColor = Rule
                    ib.SizeMode = ImageSizeMode.ZoomImage
                End Sub)
    End Sub

    ''' <summary>A titled box of text lines.</summary>
    Public Sub AddTextBox(title As String, lines As IEnumerable(Of String))
        Dim l = lines.ToList()
        AddLink(Sub(g, w) Box(g, New RectangleF(0, 10, w, 22 + Math.Max(8 + l.Count * 16, BodyHeight(l, w))), title, l))
    End Sub

    ''' <summary>Height (report pixels, 96 dpi) the body text of a Box needs at this width, wrapping included.</summary>
    Private Shared Function BodyHeight(lines As IEnumerable(Of String), width As Single) As Single
        Dim text = Environment.NewLine & "  " & String.Join(Environment.NewLine & "  ", lines)
        Using bmp As New Bitmap(1, 1)
            bmp.SetResolution(96, 96)
            Using gr = Graphics.FromImage(bmp)
                Return gr.MeasureString(text, FontBody, Math.Max(1, CInt(width) - 6)).Height + 10
            End Using
        End Using
    End Function

    ''' <summary>A simple table: header row shaded, columns sized by weight.</summary>
    Public Sub AddTable(headers As String(), rows As IEnumerable(Of String()), Optional weights As Single() = Nothing)
        Dim data = rows.ToList()
        If weights Is Nothing Then weights = Enumerable.Repeat(1.0F, headers.Length).ToArray()
        Dim total = weights.Sum()
        AddLink(Sub(g, w)
                    Dim y = 8.0F
                    Dim rowH = 21.0F
                    For r = -1 To data.Count - 1
                        Dim cells = If(r < 0, headers, data(r))
                        Dim x = 0.0F
                        g.Font = If(r < 0, FontHead, FontBody)
                        ' Header: white on brand blue; rows banded in a very light blue
                        Dim fore = If(r < 0, Color.White, Ink)
                        Dim fillColour = If(r < 0, Brand.Blue, If(r Mod 2 = 1, Color.FromArgb(246, 249, 253), Color.White))
                        For c = 0 To headers.Length - 1
                            Dim cw = w * weights(c) / total
                            Dim text = If(c < cells.Length, cells(c), "")
                            g.StringFormat = New BrickStringFormat(If(c = 0, DXStringAlignment.Near, DXStringAlignment.Center), DXStringAlignment.Center)
                            Dim tb = g.DrawString(" " & text, fore, New RectangleF(x, y, cw, rowH), BorderSide.All)
                            tb.BorderColor = Rule
                            tb.BackColor = fillColour
                            x += cw
                        Next
                        y += rowH
                    Next
                    g.StringFormat = New BrickStringFormat(DXStringAlignment.Near)
                End Sub)
    End Sub

#End Region

    ''' <summary>Builds the document and opens the DevExpress preview (print / PDF / e-mail from there).</summary>
    Public Sub ShowPreview(owner As IWin32Window)
        _composite.CreateDocument()
        _composite.ShowPreviewDialog(owner)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        For Each d In _keepAlive
            Try
                d.Dispose()
            Catch
            End Try
        Next
        _logo?.Dispose()
        _composite.Dispose()
        _ps.Dispose()
    End Sub
End Class
