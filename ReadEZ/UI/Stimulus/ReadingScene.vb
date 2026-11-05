Option Strict On

Imports System.Drawing.Drawing2D
Imports System.Drawing.Text

' ====================================================================================================
'   ReadingScene.vb - the reading text on the patient's screen (RZfrmTest.SetupROR / SetColours)
'
'   One scene draws every presentation ReadEZ uses:
'     White         the text on white (baseline)
'     Overlay       white screen, a coloured panel round the text (an overlay on the page)
'     TwoOverlays   two copies of the text side by side, each on its own colour (overlay test)
'     FullScreen    the whole screen in the colour (lenses: in a dark room; computer: room lit)
'     ColourOnly    the colour alone, no text (between presentations)
'
'   Sizes follow ReadEZ 2020: the font is drawn at the font-size setting + 8 points, the rows are
'   (0.8 + spacing / 10) line heights apart, and the overlay reaches (0.1 + size / 10) of the text
'   width beyond it. Everything is laid out in the patient's screen pixels and scaled to whatever
'   it is drawn into, so the clinician's preview is an exact copy.
'
'   For the rate-of-reading test the clinician's copy also shows the transcript marks (misread
'   words, skipped rows and the last word read), and WordAt finds the word under the mouse there.
' ====================================================================================================

Public Enum ReadingPresentation
    White
    Overlay
    TwoOverlays
    FullScreen
    ColourOnly
End Enum

''' <summary>A word's place: row and word number (both from 0).</summary>
Public Structure WordRef
    Public Row As Integer
    Public Word As Integer
    Public Sub New(row As Integer, word As Integer)
        Me.Row = row
        Me.Word = word
    End Sub
End Structure

Public NotInheritable Class ReadingScene
    Implements IStimulusScene

    Public Property Rows As List(Of String) = New List(Of String)
    Public Property FontName As String = "Times New Roman"
    ''' <summary>The font-size setting (drawn at this + 8 points).</summary>
    Public Property FontSize As Integer = 5
    ''' <summary>Line spacing setting: rows are (0.8 + this / 10) line heights apart.</summary>
    Public Property LineSpacing As Integer = 6
    ''' <summary>Overlay size setting 0 - 10.</summary>
    Public Property OverlaySize As Integer = 5
    Public Property Presentation As ReadingPresentation = ReadingPresentation.White
    Public Property Colour1 As Color = Color.White
    Public Property Colour2 As Color = Color.White
    ''' <summary>
    ''' Pattern tests: the height of a stripe (patient's screen pixels) of the horizontal black stripes
    ''' shown instead of the text, in a square block where the text would be (0 = the text). The
    ''' white of the stripes is the background, so an overlay or a coloured screen tints it.
    ''' </summary>
    Public Property PatternStripe As Integer

    ''' <summary>A small label in a corner of the patient's screen ("1", "2" for colours shown one after the other).</summary>
    Public Property Caption As String = ""

    ' ---- Transcript marks (clinician's copy only) ----
    Public Property ShowMarks As Boolean
    Public ReadOnly Property Errors As New HashSet(Of WordRef)
    Public ReadOnly Property SkippedRows As New HashSet(Of Integer)
    Public Property LastWord As WordRef? = Nothing
    ''' <summary>The word under the mouse in the clinician's copy.</summary>
    Public Property HoverWord As WordRef? = Nothing

    ' ---- The patient's own answers (colour tests, "patient chooses") ----
    ''' <summary>Buttons on the patient's screen (e.g. "The first", "The second", "Show again"); Nothing = none.</summary>
    Public Property ChoiceButtons As String() = Nothing
    ''' <summary>A question or hint on the patient's screen ("" = none).</summary>
    Public Property ChoicePrompt As String = ""
    ''' <summary>
    ''' Feedback for the patient's answer, shown for a moment: 0 / 1 = the left / right copy (or the
    ''' first / second button), 2 = the third button; -1 = none.
    ''' </summary>
    Public Property Highlight As Integer = -1

    Private Shared ReadOnly Measure As Graphics = Graphics.FromImage(New Bitmap(1, 1))
    Private Shared ReadOnly Typo As StringFormat = MakeFormat()

    Private Shared Function MakeFormat() As StringFormat
        Dim f = CType(StringFormat.GenericTypographic.Clone(), StringFormat)
        f.FormatFlags = f.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces Or StringFormatFlags.NoWrap
        Return f
    End Function

#Region "Layout (patient's screen pixels)"

    ''' <summary>Where each word is, in the patient's screen pixels, for one copy of the text.</summary>
    Private NotInheritable Class Layout
        Public FontPx As Single
        Public LineHeight As Single
        Public Pitch As Single
        Public TextWidth As Single
        Public Lefts As Single()       ' left edge of each copy (1 or 2)
        Public Top As Single
        Public WordRects As New List(Of List(Of RectangleF))   ' per row: each word, relative to the copy's left
        ''' <summary>Pattern tests: the side of the square block of stripes (0 = text).</summary>
        Public PatternSide As Single
        Public ReadOnly Property TextHeight As Single
            Get
                If PatternSide > 0 Then Return PatternSide
                Return If(WordRects.Count = 0, 0, (WordRects.Count - 1) * Pitch + LineHeight)
            End Get
        End Property
    End Class

    Private _layout As Layout
    Private _layoutKey As String = ""

    Private Function Key() As String
        Dim sz = StimulusArt.ScreenSize
        Return String.Join("|", {FontName, FontSize.ToString(), LineSpacing.ToString(), Presentation.ToString(), sz.Width.ToString(), sz.Height.ToString(),
                                 OverlaySize.ToString(), PatternStripe.ToString(), String.Join(vbLf, Rows)})
    End Function

    Private Function GetLayout() As Layout
        Dim k = Key()
        If _layout IsNot Nothing AndAlso k = _layoutKey Then Return _layout
        Dim sz = StimulusArt.ScreenSize
        Dim lay As New Layout
        Dim two = Presentation = ReadingPresentation.TwoOverlays
        If PatternStripe > 0 Then
            ' The pattern: a square block as large as fits (each in half the screen for two side by side,
            ' with room for the overlay round it), on whole pixels so the stripes are exact
            Dim grow = 1.0F + 0.1F + OverlaySize / 10.0F
            Dim room = If(two, sz.Width / 2.0F, CSng(sz.Width)) * 0.92F / grow
            Dim side = CSng(Math.Floor(Math.Min(Math.Min(sz.Height * 0.46F, room), sz.Height * 0.92F / grow)))
            lay.PatternSide = side
            lay.TextWidth = side
            lay.LineHeight = Math.Max(4.0F, side / 12)
            lay.Pitch = lay.LineHeight
            lay.Top = CSng(Math.Floor((sz.Height - side) / 2))
            If two Then
                lay.Lefts = {CSng(Math.Floor((sz.Width / 2.0F - side) / 2)), CSng(Math.Floor(sz.Width / 2.0F + (sz.Width / 2.0F - side) / 2))}
            Else
                lay.Lefts = {CSng(Math.Floor((sz.Width - side) / 2))}
            End If
            _layout = lay
            _layoutKey = k
            Return lay
        End If
        lay.FontPx = (FontSize + 8) * 96.0F / 72.0F
        ' Two copies side by side must each fit in half the screen (with room for the overlay)
        For attempt = 0 To 30
            Measure2(lay)
            Dim room = If(two, sz.Width / 2.0F, CSng(sz.Width)) * 0.92F / (1.0F + 0.1F + OverlaySize / 10.0F)
            If lay.TextWidth <= room OrElse lay.FontPx < 8 Then Exit For
            lay.FontPx *= Math.Max(0.6F, room / lay.TextWidth)
        Next
        lay.Top = (sz.Height - lay.TextHeight) / 2
        If two Then
            lay.Lefts = {(sz.Width / 2.0F - lay.TextWidth) / 2, sz.Width / 2.0F + (sz.Width / 2.0F - lay.TextWidth) / 2}
        Else
            lay.Lefts = {(sz.Width - lay.TextWidth) / 2}
        End If
        _layout = lay
        _layoutKey = k
        Return lay
    End Function

    Private Sub Measure2(lay As Layout)
        lay.WordRects.Clear()
        Using f = MakeFont(lay.FontPx)
            lay.LineHeight = f.GetHeight(96)
            lay.Pitch = lay.LineHeight * (0.8F + LineSpacing / 10.0F)
            lay.TextWidth = 0
            For r = 0 To Rows.Count - 1
                Dim rects As New List(Of RectangleF)
                Dim row = Rows(r)
                Dim y = r * lay.Pitch
                Dim i = 0
                While i < row.Length
                    While i < row.Length AndAlso row(i) = " "c
                        i += 1
                    End While
                    If i >= row.Length Then Exit While
                    Dim j = i
                    While j < row.Length AndAlso row(j) <> " "c
                        j += 1
                    End While
                    Dim x0 = If(i = 0, 0, Measure.MeasureString(row.Substring(0, i), f, PointF.Empty, Typo).Width)
                    Dim x1 = Measure.MeasureString(row.Substring(0, j), f, PointF.Empty, Typo).Width
                    rects.Add(New RectangleF(x0, y, x1 - x0, lay.LineHeight))
                    i = j
                End While
                lay.WordRects.Add(rects)
                If rects.Count > 0 Then lay.TextWidth = Math.Max(lay.TextWidth, rects.Last().Right)
            Next
        End Using
    End Sub

    Private Function MakeFont(px As Single) As Font
        Try
            Return New Font(If(String.IsNullOrWhiteSpace(FontName), "Times New Roman", FontName), Math.Max(4, px), FontStyle.Regular, GraphicsUnit.Pixel)
        Catch
            Return New Font("Times New Roman", Math.Max(4, px), FontStyle.Regular, GraphicsUnit.Pixel)
        End Try
    End Function

    ''' <summary>The coloured panel round copy n (patient's screen pixels).</summary>
    Private Function OverlayRect(lay As Layout, n As Integer) As RectangleF
        Dim overlap = (0.1F + OverlaySize / 10.0F) * lay.TextWidth
        Return New RectangleF(lay.Lefts(n) - overlap / 2, lay.Top - overlap / 2, lay.TextWidth + overlap, lay.TextHeight + overlap)
    End Function

    ''' <summary>
    ''' The text block of the (first) copy in the patient's screen pixels, with a margin of one line
    ''' round it. Empty if there is no text. Used to draw an enlarged copy for scoring.
    ''' </summary>
    Public Function TextBounds() As RectangleF
        If Rows.Count = 0 AndAlso PatternStripe <= 0 Then Return RectangleF.Empty
        Dim lay = GetLayout()
        If lay.TextWidth <= 0 OrElse lay.TextHeight <= 0 Then Return RectangleF.Empty
        Dim m = lay.LineHeight
        Return New RectangleF(lay.Lefts(0) - m, lay.Top - m, lay.TextWidth + 2 * m, lay.TextHeight + 2 * m)
    End Function

    ''' <summary>A row of the (first) copy, in the patient's screen pixels.</summary>
    Public Function RowBounds(row As Integer) As RectangleF
        Dim lay = GetLayout()
        If row < 0 OrElse row >= lay.WordRects.Count OrElse lay.WordRects(row).Count = 0 Then Return RectangleF.Empty
        Return New RectangleF(lay.Lefts(0), lay.Top + row * lay.Pitch, lay.WordRects(row).Last().Right, lay.LineHeight)
    End Function

    ''' <summary>A word of the (first) copy, in the patient's screen pixels.</summary>
    Public Function WordBounds(row As Integer, word As Integer) As RectangleF
        Return RectOf(GetLayout(), New WordRef(row, word))
    End Function

    ''' <summary>Copy n (0 or 1): its coloured overlay, or the text block itself on white / full screen.</summary>
    Public Function CopyBounds(n As Integer) As RectangleF
        If Rows.Count = 0 AndAlso PatternStripe <= 0 Then Return RectangleF.Empty
        Dim lay = GetLayout()
        If n < 0 OrElse n >= lay.Lefts.Length Then Return RectangleF.Empty
        If Presentation = ReadingPresentation.Overlay OrElse Presentation = ReadingPresentation.TwoOverlays Then Return OverlayRect(lay, n)
        Return New RectangleF(lay.Lefts(n), lay.Top, lay.TextWidth, lay.TextHeight)
    End Function

    ''' <summary>Number of words in each row.</summary>
    Public Function WordCounts() As List(Of Integer)
        Return GetLayout().WordRects.Select(Function(r) r.Count).ToList()
    End Function

#End Region

#Region "Drawing"

    Public Sub Paint(g As Graphics, bounds As Rectangle, preview As Boolean) Implements IStimulusScene.Paint
        Dim sz = StimulusArt.ScreenSize
        Dim k = bounds.Width / CSng(sz.Width)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = If(preview, TextRenderingHint.AntiAliasGridFit, TextRenderingHint.ClearTypeGridFit)
        Dim back = If(Presentation = ReadingPresentation.FullScreen OrElse Presentation = ReadingPresentation.ColourOnly, Colour1, Color.White)
        Using b As New SolidBrush(back)
            g.FillRectangle(b, bounds)
        End Using
        Dim state = g.Save()
        g.TranslateTransform(bounds.X, bounds.Y)
        g.ScaleTransform(k, k)
        Try
            If Presentation <> ReadingPresentation.ColourOnly AndAlso (Rows.Count > 0 OrElse PatternStripe > 0) Then DrawText(g, preview)
            If Caption <> "" Then
                Using f As New Font("Segoe UI Semibold", 40, GraphicsUnit.Pixel), cb As New SolidBrush(Color.FromArgb(130, 60, 60, 60))
                    g.DrawString(Caption, f, cb, 40, 30)
                End Using
            End If
            DrawChoices(g)
        Finally
            g.Restore(state)
        End Try
    End Sub

    Private Sub DrawText(g As Graphics, preview As Boolean)
        Dim lay = GetLayout()
        Dim copies = lay.Lefts.Length
        For n = 0 To copies - 1
            ' The coloured panel (overlay)
            If Presentation = ReadingPresentation.Overlay OrElse Presentation = ReadingPresentation.TwoOverlays Then
                Using b As New SolidBrush(If(n = 0, Colour1, Colour2))
                    g.FillRectangle(b, OverlayRect(lay, n))
                End Using
            End If
            If PatternStripe > 0 Then
                ' The pattern: black stripes; the white ones are the background (white, the overlay or the screen colour)
                Dim oldSmoothing = g.SmoothingMode
                g.SmoothingMode = SmoothingMode.None
                Dim y = lay.Top
                While y < lay.Top + lay.PatternSide
                    g.FillRectangle(Brushes.Black, lay.Lefts(n), y, lay.PatternSide, Math.Min(PatternStripe, lay.Top + lay.PatternSide - y))
                    y += 2 * PatternStripe
                End While
                g.SmoothingMode = oldSmoothing
                Continue For
            End If
            Using f = MakeFont(lay.FontPx)
                For r = 0 To Rows.Count - 1
                    g.DrawString(Rows(r), f, Brushes.Black, lay.Lefts(n), lay.Top + r * lay.Pitch, Typo)
                Next
            End Using
        Next
        If preview AndAlso ShowMarks AndAlso copies = 1 Then DrawMarks(g, lay)
    End Sub

    ''' <summary>Transcript marks on the clinician's copy: red cross over a misread word, a line through a skipped row, a bar after the last word.</summary>
    Private Sub DrawMarks(g As Graphics, lay As Layout)
        Dim x0 = lay.Lefts(0), y0 = lay.Top
        Dim w = Math.Max(2.0F, lay.LineHeight * 0.08F)
        Using red As New Pen(Color.FromArgb(220, 210, 30, 30), w), skip As New Pen(Color.FromArgb(200, 230, 120, 0), w * 1.5F),
              last As New Pen(Color.FromArgb(230, 20, 120, 220), w * 1.6F), hover As New SolidBrush(Color.FromArgb(55, 20, 120, 220))
            If HoverWord.HasValue Then
                Dim h = RectOf(lay, HoverWord.Value)
                If Not h.IsEmpty Then g.FillRectangle(hover, RectangleF.Inflate(h, 3, 2))
            End If
            For Each r In SkippedRows
                If r < 0 OrElse r >= lay.WordRects.Count OrElse lay.WordRects(r).Count = 0 Then Continue For
                Dim y = y0 + r * lay.Pitch + lay.LineHeight * 0.55F
                g.DrawLine(skip, x0 - 6, y, x0 + lay.WordRects(r).Last().Right + 6, y)
            Next
            For Each e In Errors
                Dim rc = RectOf(lay, e)
                If rc.IsEmpty Then Continue For
                g.DrawLine(red, rc.Left, rc.Top + 2, rc.Right, rc.Bottom - 2)
                g.DrawLine(red, rc.Left, rc.Bottom - 2, rc.Right, rc.Top + 2)
            Next
            If LastWord.HasValue Then
                Dim rc = RectOf(lay, LastWord.Value)
                If Not rc.IsEmpty Then
                    g.DrawLine(last, rc.Right + 4, rc.Top - 2, rc.Right + 4, rc.Bottom + 2)
                    g.DrawLine(last, rc.Left, rc.Bottom + 1, rc.Right + 4, rc.Bottom + 1)
                End If
            End If
        End Using
    End Sub

    Private Shared Function RectOf(lay As Layout, w As WordRef) As RectangleF
        If w.Row < 0 OrElse w.Row >= lay.WordRects.Count Then Return RectangleF.Empty
        Dim row = lay.WordRects(w.Row)
        If w.Word < 0 OrElse w.Word >= row.Count Then Return RectangleF.Empty
        Dim r = row(w.Word)
        Return New RectangleF(lay.Lefts(0) + r.X, lay.Top + r.Y, r.Width, r.Height)
    End Function

#Region "The patient's answers"

    ''' <summary>Where the patient's buttons are (patient's screen pixels): a row across the middle.</summary>
    Private Function ChoiceRects() As List(Of RectangleF)
        Dim list As New List(Of RectangleF)
        If ChoiceButtons Is Nothing OrElse ChoiceButtons.Length = 0 Then Return list
        Dim sz = StimulusArt.ScreenSize
        Dim w = sz.Width * 0.22F, h = sz.Height * 0.12F, gap = sz.Width * 0.03F
        Dim total = ChoiceButtons.Length * w + (ChoiceButtons.Length - 1) * gap
        Dim x = (sz.Width - total) / 2, y = sz.Height * 0.5F
        For i = 0 To ChoiceButtons.Length - 1
            list.Add(New RectangleF(x + i * (w + gap), y, w, h))
        Next
        Return list
    End Function

    ''' <summary>The patient's button under a point of the screen drawn into bounds (-1 = none).</summary>
    Public Function ChoiceAt(bounds As Rectangle, p As Point) As Integer
        Dim k = bounds.Width / CSng(StimulusArt.ScreenSize.Width)
        If k <= 0 Then Return -1
        Dim x = (p.X - bounds.X) / k, y = (p.Y - bounds.Y) / k
        Dim rs = ChoiceRects()
        For i = 0 To rs.Count - 1
            If rs(i).Contains(x, y) Then Return i
        Next
        Return -1
    End Function

    ''' <summary>The question and the buttons, large, for the patient to click (in screen pixels; Paint scales them).</summary>
    Private Sub DrawChoices(g As Graphics)
        Dim sz = StimulusArt.ScreenSize
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit
        Using fmt As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            If ChoicePrompt <> "" Then
                Dim hasButtons = ChoiceButtons IsNot Nothing AndAlso ChoiceButtons.Length > 0
                Dim py = If(hasButtons, sz.Height * 0.36F, sz.Height * 0.92F)
                Using f As New Font("Segoe UI Semibold", sz.Height * 0.045F, GraphicsUnit.Pixel), b As New SolidBrush(Color.FromArgb(40, 45, 55))
                    g.DrawString(ChoicePrompt, f, b, New RectangleF(0, py - sz.Height * 0.06F, sz.Width, sz.Height * 0.12F), fmt)
                End Using
            End If
            Dim rs = ChoiceRects()
            For i = 0 To rs.Count - 1
                Dim r = rs(i)
                Dim chosen = i = Highlight
                ' The button tapped turns green with a tick, a little larger
                Dim br = If(chosen, RectangleF.Inflate(r, r.Height * 0.06F, r.Height * 0.06F), r)
                Using path = Draw.RoundedRect(br, br.Height * 0.2F), b As New SolidBrush(If(chosen, TapGreen, Brand.Navy)),
                      f As New Font("Segoe UI Semibold", r.Height * 0.32F, GraphicsUnit.Pixel)
                    g.FillPath(b, path)
                    g.DrawString(If(chosen, ChrW(&H2713) & "  ", "") & ChoiceButtons(i), f, Brushes.White, br, fmt)
                End Using
            Next
        End Using
        ' Overlays: the copy chosen gets a green frame and a tick
        If Highlight >= 0 AndAlso Highlight <= 1 AndAlso (ChoiceButtons Is Nothing OrElse ChoiceButtons.Length = 0) AndAlso
           Presentation = ReadingPresentation.TwoOverlays AndAlso Rows.Count > 0 Then
            Dim lay = GetLayout()
            If Highlight < lay.Lefts.Length Then
                Dim r = OverlayRect(lay, Highlight)
                Dim w = Math.Max(4.0F, sz.Height * 0.012F)
                Using pen As New Pen(TapGreen, w), b As New SolidBrush(TapGreen),
                      f As New Font("Segoe UI Semibold", sz.Height * 0.05F, GraphicsUnit.Pixel),
                      fmt As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
                    g.DrawRectangle(pen, r.X - w, r.Y - w, r.Width + 2 * w, r.Height + 2 * w)
                    Dim d = sz.Height * 0.09F
                    Dim badge As New RectangleF(r.Right - d * 0.6F, r.Y - d * 0.4F, d, d)
                    g.FillEllipse(b, badge)
                    g.DrawString(ChrW(&H2713).ToString(), f, Brushes.White, badge, fmt)
                End Using
            End If
        End If
    End Sub

    Private Shared ReadOnly TapGreen As Color = Color.FromArgb(46, 140, 60)

#End Region

    ''' <summary>The word under a point of a copy drawn into bounds (Nothing if none).</summary>
    Public Function WordAt(bounds As Rectangle, p As Point) As WordRef?
        If Presentation = ReadingPresentation.TwoOverlays OrElse Presentation = ReadingPresentation.ColourOnly Then Return Nothing
        Dim lay = GetLayout()
        Dim k = bounds.Width / CSng(StimulusArt.ScreenSize.Width)
        Dim x = (p.X - bounds.X) / k, y = (p.Y - bounds.Y) / k
        For r = 0 To lay.WordRects.Count - 1
            Dim top = lay.Top + r * lay.Pitch
            If y < top - lay.Pitch / 2 + lay.LineHeight / 2 OrElse y > top + lay.Pitch / 2 + lay.LineHeight / 2 Then Continue For
            Dim row = lay.WordRects(r)
            For i = 0 To row.Count - 1
                Dim wr = row(i)
                Dim gapRight = If(i < row.Count - 1, (row(i + 1).Left - wr.Right) / 2, lay.LineHeight)
                Dim gapLeft = If(i > 0, (wr.Left - row(i - 1).Right) / 2, lay.LineHeight)
                If x >= lay.Lefts(0) + wr.Left - gapLeft AndAlso x <= lay.Lefts(0) + wr.Right + gapRight Then Return New WordRef(r, i)
            Next
        Next
        Return Nothing
    End Function

    ''' <summary>Words up to and including the last word read (rows before it counted in full).</summary>
    Public Function WordsReadTo(last As WordRef) As Integer
        Dim counts = WordCounts()
        Dim n = 0
        For r = 0 To Math.Min(last.Row, counts.Count) - 1
            n += counts(r)
        Next
        Return n + last.Word + 1
    End Function

    ''' <summary>The transcript as ReadEZ 2020 kept it: per row, "X" under a misread word, "-" for a skipped row, "]" after the last word.</summary>
    Public Function TranscriptText() As String
        Dim lines As New List(Of String)
        For r = 0 To Rows.Count - 1
            Dim row = Rows(r)
            Dim marks = New String(" "c, row.Length).ToCharArray()
            If SkippedRows.Contains(r) Then
                For i = 0 To marks.Length - 1
                    marks(i) = "-"c
                Next
            End If
            Dim wordStarts As New List(Of Integer), wordEnds As New List(Of Integer)
            Dim inWord = False
            For i = 0 To row.Length - 1
                If row(i) <> " "c AndAlso Not inWord Then wordStarts.Add(i) : inWord = True
                If row(i) = " "c AndAlso inWord Then wordEnds.Add(i - 1) : inWord = False
            Next
            If inWord Then wordEnds.Add(row.Length - 1)
            Dim rowNo = r
            For Each e In Errors.Where(Function(x) x.Row = rowNo)
                If e.Word < wordStarts.Count Then marks(wordStarts(e.Word)) = "X"c
            Next
            If LastWord.HasValue AndAlso LastWord.Value.Row = r AndAlso LastWord.Value.Word < wordEnds.Count Then marks(wordEnds(LastWord.Value.Word)) = "]"c
            lines.Add(New String(marks).TrimEnd())
        Next
        Return String.Join(vbLf, lines)
    End Function

#End Region

End Class
