Option Strict On

Imports System.Drawing.Drawing2D
Imports System.Drawing.Text

' ====================================================================================================
'   SplashScreen.vb - the start-up label (also the top of the About box)
'
'     SplashArt        draws the label: manufacturer, product and version, UKCA / CE marks,
'                      EU representative. Drawn (not a picture) so it is sharp at any scaling.
'     SplashArtPanel   a control showing the label (About box)
'     SplashForm       shown the moment the program starts, while settings, licence and database
'                      load; it stays up for at least two seconds and fades once the main window
'                      is showing.
'
'   The splash uses nothing that is slow to load (no settings, translations, theme or DevExpress)
'   so it appears straight away. The regulatory wording is in English, as on the device label.
' ====================================================================================================

Public NotInheritable Class SplashArt
    Private Sub New()
    End Sub

    ''' <summary>Size of the label in logical pixels (96 dpi).</summary>
    Public Shared ReadOnly LogicalSize As New Size(510, 381)

    Public Const Manufacturer As String = "131a Dixons Hill Road, Hatfield, Herts AL9 7DW  UK"
    Public Const EuRepresentative As String = "Advena Ltd. Tower Business Centre, 2nd Flr," & vbLf & "Tower Street, Swatar, BKR 4013 Malta"

    Private Shared ReadOnly Frame As Color = Color.FromArgb(185, 205, 228)
    Private Shared ReadOnly Ink As Color = Color.FromArgb(20, 20, 20)
    Private Shared _logo As Image

    ''' <summary>"ReadEZ 3.0.0" (read from the program file directly: no other start-up code needed).</summary>
    Public Shared ReadOnly Property ProductLine As String
        Get
            Dim v = GetType(SplashArt).Assembly.GetName().Version
            Dim text = If(v Is Nothing, "", If(v.Revision > 0, v.ToString(4), v.ToString(3)))
            Return AppHost.ProductTitle & " " & text
        End Get
    End Property

    Private Shared ReadOnly Property Logo As Image
        Get
            If _logo Is Nothing Then
                Try
                    Using s = GetType(SplashArt).Assembly.GetManifestResourceStream("ThomsonLogo.png")
                        If s IsNot Nothing Then _logo = New Bitmap(s)
                    End Using
                Catch
                End Try
            End If
            Return _logo
        End Get
    End Property

    ''' <summary>Draws the label filling the rectangle (keeping its proportions).</summary>
    Public Shared Sub Draw(g As Graphics, bounds As Rectangle)
        Dim k = Math.Min(bounds.Width / CSng(LogicalSize.Width), bounds.Height / CSng(LogicalSize.Height))
        Dim state = g.Save()
        g.TranslateTransform(bounds.X + (bounds.Width - LogicalSize.Width * k) / 2, bounds.Y + (bounds.Height - LogicalSize.Height * k) / 2)
        g.ScaleTransform(k, k)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.PixelOffsetMode = PixelOffsetMode.HighQuality

        ' Frame and the two white panels
        Using b As New SolidBrush(Frame)
            g.FillRectangle(b, 0, 0, LogicalSize.Width, LogicalSize.Height)
        End Using
        g.FillRectangle(Brushes.White, 10, 10, 490, 215)
        g.FillRectangle(Brushes.White, 10, 232, 490, 139)

        Using ink As New SolidBrush(SplashArt.Ink),
              fAddress As New Font("Segoe UI", 12.5F, FontStyle.Regular, GraphicsUnit.Pixel),
              fProduct As New Font("Segoe UI", 16.5F, FontStyle.Regular, GraphicsUnit.Pixel),
              fSmall As New Font("Segoe UI", 11.2F, FontStyle.Regular, GraphicsUnit.Pixel),
              fMark As New Font("Arial", 22.0F, FontStyle.Bold, GraphicsUnit.Pixel),
              fBox As New Font("Arial", 16.5F, FontStyle.Bold, GraphicsUnit.Pixel),
              fMd As New Font("Arial", 15.5F, FontStyle.Bold, GraphicsUnit.Pixel)

            ' Manufacturer
            Dim img = Logo
            If img IsNot Nothing Then
                Dim w = 305.0F, h = w * img.Height / img.Width
                g.DrawImage(img, 18.0F, 20.0F, w, h)
            End If
            g.DrawString(Manufacturer, fAddress, ink, 104, 146)
            DrawFactory(g, 26, 182, ink)
            g.DrawString(ProductLine, fProduct, ink, 114, 171)

            ' UKCA
            DrawUkca(g, 19, 239, fMark, ink)
            Dim wrap As New StringFormat With {.Trimming = StringTrimming.Word}
            g.DrawString("The " & AppHost.ProductTitle & " Software complies with the requirements of the United Kingdom Medical " &
                         "Devices Regulations 2002 (SI 2002 No. 618, as amended) (UK MDR 2002).",
                         fSmall, ink, New RectangleF(66, 239, 420, 48), wrap)

            ' CE, EU representative, medical device
            DrawCe(g, 20, 296, 27, ink)
            DrawBoxedLabel(g, New RectangleF(76, 293, 46, 31), "EC", fBox, ink)
            DrawBoxedLabel(g, New RectangleF(122, 293, 54, 31), "REP", fBox, ink)
            g.DrawString(EuRepresentative, fSmall, ink, 188, 294)
            DrawBoxedLabel(g, New RectangleF(448, 294, 42, 24), "MD", fMd, ink, 1.6F)
            g.DrawString("The " & AppHost.ProductTitle & " Software complies with the requirements of the European Union " &
                         "Medical Devices Regulation 2017/745/EU",
                         fSmall, ink, New RectangleF(18, 331, 474, 36), wrap)
        End Using
        g.Restore(state)
    End Sub

    ''' <summary>The manufacturer symbol (ISO 7000-3082): a factory with a chimney.</summary>
    Private Shared Sub DrawFactory(g As Graphics, x As Single, y As Single, b As Brush)
        Dim pts = {New PointF(0, 27), New PointF(0, 0), New PointF(4.5F, 0), New PointF(4.5F, 15), New PointF(12, 9), New PointF(12, 15),
                   New PointF(19.5F, 9), New PointF(19.5F, 15), New PointF(27, 9), New PointF(27, 27)}
        g.FillPolygon(b, pts.Select(Function(p) New PointF(x + p.X, y + p.Y)).ToArray())
    End Sub

    ''' <summary>The UKCA mark: U K over C A.</summary>
    Private Shared Sub DrawUkca(g As Graphics, x As Single, y As Single, f As Font, b As Brush)
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
        Dim cw = 17.5F, ch = 20.0F
        g.DrawString("U", f, b, New RectangleF(x, y, cw, ch), sf)
        g.DrawString("K", f, b, New RectangleF(x + cw, y, cw, ch), sf)
        g.DrawString("C", f, b, New RectangleF(x, y + ch, cw, ch), sf)
        g.DrawString("A", f, b, New RectangleF(x + cw, y + ch, cw, ch), sf)
    End Sub

    ''' <summary>The CE mark: C and E drawn from circles of the same size.</summary>
    Private Shared Sub DrawCe(g As Graphics, x As Single, y As Single, h As Single, b As Brush)
        Dim w = h * 0.13F                   ' stroke
        Dim r = (h - w) / 2                 ' radius of the stroke's centre line
        Using p As New Pen(b, w) With {.StartCap = LineCap.Flat, .EndCap = LineCap.Flat}
            ' C
            g.DrawArc(p, x + w / 2, y + w / 2, 2 * r, 2 * r, 52, 256)
            ' E: same arc, further right, with a middle bar
            Dim ex = x + h * 0.9F
            g.DrawArc(p, ex + w / 2, y + w / 2, 2 * r, 2 * r, 52, 256)
            Using bar As New Pen(b, w * 0.85F)
                g.DrawLine(bar, ex + w * 0.5F, y + h / 2, ex + h * 0.78F, y + h / 2)
            End Using
        End Using
    End Sub

    Private Shared Sub DrawBoxedLabel(g As Graphics, r As RectangleF, text As String, f As Font, b As Brush, Optional lineWidth As Single = 2.0F)
        Using p As New Pen(b, lineWidth)
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height)
        End Using
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
        g.DrawString(text, f, b, r, sf)
    End Sub
End Class

''' <summary>The label as a control (About box). Keeps its proportions; white around it.</summary>
Public NotInheritable Class SplashArtPanel
    Inherits Control

    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        Size = SplashArt.LogicalSize
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        e.Graphics.Clear(Color.White)
        SplashArt.Draw(e.Graphics, ClientRectangle)
    End Sub
End Class

''' <summary>The start-up window. ShowNow at the very start, Finish when the main window is showing.</summary>
Public NotInheritable Class SplashForm
    Inherits Form

    Private Shared _current As SplashForm
    Private Shared ReadOnly _clock As Stopwatch = New Stopwatch()
    Private Const MinimumMs As Integer = 2000
    Private Const FadeMs As Integer = 250
    Private ReadOnly _timer As New Timer With {.Interval = 15}
    Private _fadeStart As Long = -1

    Private Sub New()
        FormBorderStyle = FormBorderStyle.None
        StartPosition = FormStartPosition.Manual
        ShowInTaskbar = True
        Text = AppHost.ProductTitle
        TopMost = True
        BackColor = Color.White
        AutoScaleMode = AutoScaleMode.None
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.UserPaint, True)
        Try
            Using s = GetType(SplashForm).Assembly.GetManifestResourceStream("ReadEZ.ico")
                If s IsNot Nothing Then Icon = New Icon(s)
            End Using
        Catch
        End Try
        AddHandler _timer.Tick, AddressOf OnTick
        AddHandler Click, Sub() Finish(Nothing, immediately:=True)
    End Sub

    ''' <summary>Shows the splash straight away (centred on the primary monitor) and paints it.</summary>
    Public Shared Sub ShowNow()
        If _current IsNot Nothing Then Return
        Try
            _clock.Restart()
            Dim f As New SplashForm()
            f.CreateHandle()                                    ' DeviceDpi of the monitor it is on
            Dim k = f.DeviceDpi / 96.0
            Dim area = Screen.PrimaryScreen.WorkingArea
            Dim sz As New Size(CInt(SplashArt.LogicalSize.Width * k), CInt(SplashArt.LogicalSize.Height * k))
            f.Bounds = New Rectangle(area.X + (area.Width - sz.Width) \ 2, area.Y + (area.Height - sz.Height) \ 2, sz.Width, sz.Height)
            _current = f
            f.Show()
            f.Update()                                          ' paint now: the start-up work below blocks this thread
        Catch ex As Exception
            Log.Warn("Splash: " & ex.Message)
            _current = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' The main window is showing: the splash stays above it until it has been up for two seconds,
    ''' then fades. It is no longer "always on top", so any message that appears now is above it.
    ''' </summary>
    Public Shared Sub Finish(main As Form, Optional immediately As Boolean = False)
        Dim f = _current
        If f Is Nothing OrElse f.IsDisposed Then Return
        If main IsNot Nothing Then
            f.TopMost = False
            f.Owner = main                                      ' kept above the main window, below its dialogs
        End If
        If immediately Then f._fadeStart = 0
        f._timer.Start()
    End Sub

    ''' <summary>Closes the splash at once (before a message or a dialog that would be hidden behind it).</summary>
    Public Shared Sub CloseNow()
        Dim f = _current
        _current = Nothing
        If f Is Nothing OrElse f.IsDisposed Then Return
        f._timer.Stop()
        f.Close()
        f.Dispose()
    End Sub

    Private Sub OnTick(sender As Object, e As EventArgs)
        If _fadeStart < 0 Then
            If _clock.ElapsedMilliseconds < MinimumMs Then Return
            _fadeStart = _clock.ElapsedMilliseconds
        ElseIf _fadeStart = 0 Then
            _fadeStart = _clock.ElapsedMilliseconds
        End If
        Dim t = (_clock.ElapsedMilliseconds - _fadeStart) / CDbl(FadeMs)
        If t >= 1 Then
            CloseNow()
        Else
            Opacity = 1 - t
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        SplashArt.Draw(e.Graphics, ClientRectangle)
    End Sub

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        ' everything is painted in OnPaint
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _timer.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class
