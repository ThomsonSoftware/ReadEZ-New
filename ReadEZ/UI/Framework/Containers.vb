Option Strict On

Imports System.ComponentModel
Imports System.Drawing.Drawing2D

' ====================================================================================================
'   Containers.vb
'   Card, collapsible card (replaces the accordion groups), segmented control (replaces groups of
'   CheckButtons / XtraTab headers), buffered panel and label helpers.
' ====================================================================================================

#Region "Label helpers"

Public NotInheritable Class Ui
    Private Sub New()
    End Sub

    Public Shared Function Label(key As String, Optional muted As Boolean = False) As Label
        Dim l As New Label With {
            .AutoSize = True,
            .Font = If(muted, Fonts.Small, Fonts.Body),
            .Margin = New Padding(0, 0, 0, 4),
            .BackColor = Color.Transparent
        }
        If key <> "" Then Loc.Bind(l, key)
        If muted Then LabelStyle.SetMuted(l) Else l.ForeColor = ThemeManager.Current.Ink
        Return l
    End Function

    Public Shared Function Section(key As String) As Label
        Dim l As New Label With {
            .AutoSize = True,
            .Font = Fonts.Section,
            .ForeColor = ThemeManager.Current.Ink,
            .Margin = New Padding(0, 6, 0, 10),
            .BackColor = Color.Transparent
        }
        If key <> "" Then Loc.Bind(l, key)
        Return l
    End Function

    Public Shared Function Title(key As String) As Label
        Dim l As New Label With {
            .AutoSize = True,
            .Font = Fonts.Title,
            .ForeColor = ThemeManager.Current.Ink,
            .Margin = New Padding(0, 0, 0, 4),
            .BackColor = Color.Transparent
        }
        If key <> "" Then Loc.Bind(l, key)
        Return l
    End Function

    ''' <summary>Label above a control, stacked in one cell.</summary>
    Public Shared Function Stack(labelKey As String, field As Control, Optional width As Integer = 0) As Control
        Dim p As New TableLayoutPanel With {
            .ColumnCount = 1, .RowCount = 2, .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Margin = New Padding(0, 0, 16, 4), .BackColor = Color.Transparent
        }
        p.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        If width > 0 Then field.Width = width
        p.Controls.Add(Label(labelKey, muted:=True), 0, 0)
        p.Controls.Add(field, 0, 1)
        Return p
    End Function

    ''' <summary>Vertical stack that grows with its contents (for forms built in code).</summary>
    Public Shared Function Column(Optional padding As Integer = 0) As FlowLayoutPanel
        Return New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Padding = New Padding(padding), .Margin = New Padding(0),
            .BackColor = Color.Transparent}
    End Function

    ''' <summary>Horizontal row that grows with its contents.</summary>
    Public Shared Function Row(Optional wrap As Boolean = False) As FlowLayoutPanel
        Return New FlowLayoutPanel With {
            .FlowDirection = FlowDirection.LeftToRight, .WrapContents = wrap, .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Margin = New Padding(0, 0, 0, 8),
            .BackColor = Color.Transparent}
    End Function

    Public Shared Function Separator() As Control
        Dim s As New Panel With {.Width = 1, .Height = 28, .Margin = New Padding(6, 4, 14, 0)}
        AddHandler s.Paint, Sub(o, e) e.Graphics.Clear(ThemeManager.Current.Border)
        Return s
    End Function

    Public Shared Function NumberBox(min As Decimal, max As Decimal, value As Decimal, Optional decimals As Integer = 0, Optional width As Integer = 90) As NumericUpDown
        Dim n As New NumericUpDown With {
            .Minimum = min, .Maximum = max, .DecimalPlaces = decimals, .Width = width,
            .BorderStyle = BorderStyle.FixedSingle, .Font = Fonts.Body, .Margin = New Padding(0, 0, 0, 12),
            .BackColor = ThemeManager.Current.InputBack, .ForeColor = ThemeManager.Current.Ink}
        n.Value = Math.Max(min, Math.Min(max, value))
        Return n
    End Function
End Class

#End Region

#Region "Panels"

''' <summary>Panel with double buffering switched on (for custom painting).</summary>
Public Class BufferedPanel
    Inherits Panel

    Public Sub New()
        SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
    End Sub
End Class

''' <summary>PictureBox with double buffering and high-quality defaults for live drawing.</summary>
Public Class Canvas
    Inherits PictureBox

    Public Sub New()
        SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
    End Sub
End Class

''' <summary>Rounded surface panel with a hairline border and an optional title.</summary>
''' <summary>
''' A lighter rounded panel inside a card that groups related controls (e.g. a test's options).
''' It sizes itself to the one control it holds (Content), with Padding around it.
''' </summary>
Public Class InsetPanel
    Inherits Panel
    Implements IThemed

    Private ReadOnly _content As Control

    Public Sub New(content As Control, Optional pad As Integer = 10)
        SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
        _content = content
        AutoSize = True
        AutoSizeMode = AutoSizeMode.GrowAndShrink
        Padding = New Padding(pad)
        Margin = New Padding(0, 0, 0, 8)
        ThemeManager.SetRole(Me, SurfaceRole.Inset)       ' children are drawn on the inset colour
        content.Margin = New Padding(0)
        content.Location = New Point(pad, pad)            ' a Panel ignores Padding for undocked children
        Controls.Add(content)
    End Sub

    Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
        Dim s = _content.GetPreferredSize(Size.Empty)
        Return New Size(_content.Left + Math.Max(s.Width, _content.Width) + Padding.Right,
                        _content.Top + Math.Max(s.Height, _content.Height) + Padding.Bottom)
    End Function

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        BackColor = p.Inset
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return True
        End Get
    End Property

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        e.Graphics.Clear(Draw.ParentBack(Me))
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Using path = Draw.RoundedRect(New RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F), LogicalToDeviceUnits(6)),
              b As New SolidBrush(BackColor), pen As New Pen(p.Border)
            e.Graphics.FillPath(b, path)
            e.Graphics.DrawPath(pen, path)
        End Using
    End Sub
End Class

''' <summary>
''' A picture drawn centred in the control (zoomed to fit) with rounded corners and a soft shadow.
''' </summary>
Public Class ShadowPicture
    Inherits Control

    Private _image As Image

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor, True)
        BackColor = Color.Transparent
    End Sub

    Public Property Image As Image
        Get
            Return _image
        End Get
        Set(value As Image)
            _image = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        If _image Is Nothing OrElse _image.Width = 0 OrElse _image.Height = 0 Then Return
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        ' Room for the shadow around the picture
        Dim blur = LogicalToDeviceUnits(8)
        Dim off = LogicalToDeviceUnits(3)
        Dim avail = New Rectangle(blur, blur \ 2, Width - blur * 2, Height - blur * 2)
        If avail.Width <= 0 OrElse avail.Height <= 0 Then Return
        Dim scale = Math.Min(avail.Width / CDbl(_image.Width), avail.Height / CDbl(_image.Height))
        Dim w = CInt(_image.Width * scale), h = CInt(_image.Height * scale)
        Dim r As New Rectangle(avail.X + (avail.Width - w) \ 2, avail.Y + (avail.Height - h) \ 2, w, h)
        Dim radius = LogicalToDeviceUnits(6)
        ' Soft shadow: layers of faint black, each a little larger
        For i = blur To 1 Step -1
            Dim a = CInt(28.0 * (1 - i / CDbl(blur + 1)) ^ 2)
            Dim sr = New RectangleF(r.X - i + off \ 2, r.Y - i + off, r.Width + i * 2, r.Height + i * 2)
            Using path = Draw.RoundedRect(sr, radius + i), b As New SolidBrush(Color.FromArgb(a, Color.Black))
                g.FillPath(b, path)
            End Using
        Next
        Using path = Draw.RoundedRect(New RectangleF(r.X, r.Y, r.Width, r.Height), radius)
            Dim old = g.Clip
            g.SetClip(path)
            g.DrawImage(_image, r)
            g.Clip = old
        End Using
    End Sub
End Class

''' <summary>Copies a chart or picture to the clipboard as an image.</summary>
Public NotInheritable Class CopyPicture
    Private Sub New()
    End Sub

    Public Shared Sub ToClipboard(target As Control, owner As Form)
        If target Is Nothing OrElse target.Width <= 0 OrElse target.Height <= 0 Then Return
        Try
            Using bmp = Capture(target)
                Clipboard.SetImage(bmp)
            End Using
            If owner IsNot Nothing Then Toast.Show(owner, Loc.T("common.copied"), ToastKind.Success, 2000)
        Catch ex As Exception
            Log.Error("Copy picture", ex)
            Ask.Error(owner, "common.errorWithDetail", ex.Message)
        End Try
    End Sub

    ''' <summary>The control as it looks on screen (ReadEZ draws its own charts, so DrawToBitmap catches everything).</summary>
    Public Shared Function Capture(target As Control) As Bitmap
        Dim bmp As New Bitmap(target.Width, target.Height)
        target.DrawToBitmap(bmp, New Rectangle(0, 0, bmp.Width, bmp.Height))
        Return bmp
    End Function

    ''' <summary>True if the control is, or contains, a chart (none in ReadEZ).</summary>
    Public Shared Function HasChart(c As Control) As Boolean
        Return False
    End Function
End Class

Public Class Card
    Inherits Panel
    Implements IThemed

    Private _titleKey As String = ""
    Private _title As String = ""
    Protected HeaderHeight As Integer = 40
    Protected TitleFont As Font = Fonts.Section

    Public Sub New()
        SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
        Padding = New Padding(16)
        Margin = New Padding(0, 0, 12, 12)
        ThemeManager.SetRole(Me, SurfaceRole.Surface)
    End Sub

    ''' <summary>Shows a translated title at the top of the card.</summary>
    Private _titleBound As Boolean

    Public Sub SetTitle(key As String)
        _titleKey = key
        _title = Loc.T(_titleKey)
        Invalidate()
        If _titleBound Then Return               ' already follows language changes (the title can change, e.g. per test)
        _titleBound = True
        Loc.OnChange(Me, Sub()
                             _title = Loc.T(_titleKey)
                             Invalidate()
                         End Sub)
        Padding = New Padding(Padding.Left, Dpi.Px(Me, HeaderHeight), Padding.Right, Padding.Bottom)
    End Sub

    Public ReadOnly Property TitleText As String
        Get
            Return _title
        End Get
    End Property

#Region "Header icons (Zoom, Copy)"

    Private ReadOnly _headerIcons As New List(Of ModernButton)
    Private Shared ReadOnly _iconTip As New ToolTip

    ''' <summary>Adds a small icon button at the right of the header (right to left in the order added).</summary>
    Protected Friend Function AddHeaderIcon(glyph As String, tipKey As String, onClick As Action) As ModernButton
        Dim b = ModernButton.CreateIcon(glyph, If(tipKey = "", Nothing, _iconTip), tipKey)
        b.Size = New Size(Dpi.Px(Me, 28), Dpi.Px(Me, 28))
        b.Margin = New Padding(0)
        b.TabStop = False
        b.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        AddHandler b.Click, Sub() onClick()
        _headerIcons.Add(b)
        Controls.Add(b)
        b.BringToFront()
        PlaceHeaderIcons()
        Return b
    End Function

    Private Sub PlaceHeaderIcons()
        Dim x = Width - Dpi.Px(Me, 4)
        Dim head = Dpi.Px(Me, HeaderHeight)
        For Each b In _headerIcons
            x -= b.Width
            b.Location = New Point(x, Math.Max(Dpi.Px(Me, 2), (head - b.Height) \ 2))
            x -= Dpi.Px(Me, 2)
        Next
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        If _headerIcons.Count > 0 Then PlaceHeaderIcons()
    End Sub

    ''' <summary>Adds a Copy button that puts a picture of the control on the clipboard (charts and pictures).</summary>
    Public Sub EnableCopy(target As Control)
        AddHeaderIcon(Glyphs.Copy, "common.copyPicture", Sub() CopyPicture.ToClipboard(target, FindForm()))
    End Sub

#End Region

    Public Overridable Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        BackColor = p.Surface
        ForeColor = p.Ink
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return True
        End Get
    End Property

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        e.Graphics.Clear(Draw.ParentBack(Me))
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Using path = Draw.RoundedRect(New RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F), LogicalToDeviceUnits(8)),
              b As New SolidBrush(BackColor), pen As New Pen(p.Border)
            e.Graphics.FillPath(b, path)
            e.Graphics.DrawPath(pen, path)
        End Using
    End Sub

    ''' <summary>Shows the title in white on a blue band (e.g. the running test's name), so it stands out.</summary>
    Public Property TitleBanner As Boolean
        Get
            Return _banner
        End Get
        Set(value As Boolean)
            _banner = value
            HeaderHeight = If(value, 52, 40)          ' room for the band and a gap under it (set before SetTitle)
            Invalidate()
        End Set
    End Property
    Private _banner As Boolean

    ''' <summary>Title in the brand accent colour with a short accent bar before it (analysis windows).</summary>
    Public Property Accent As Boolean

    ''' <summary>True inside a screen that brands all its cards (IBrandedHost, e.g. Replay and the test screen).</summary>
    Private Function InBrandedHost() As Boolean
        Dim c = Parent
        While c IsNot Nothing
            Dim h = TryCast(c, IBrandedHost)
            If h IsNot Nothing Then Return h.Branded
            c = c.Parent
        End While
        Return False
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        If _title = "" Then Return
        Dim p = ThemeManager.Current
        Dim head = LogicalToDeviceUnits(HeaderHeight)
        If TitleBanner Then
            Dim inset = LogicalToDeviceUnits(8)
            Dim band As New Rectangle(inset, inset, Width - inset * 2, LogicalToDeviceUnits(34))
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Using path = Draw.RoundedRect(New RectangleF(band.X, band.Y, band.Width, band.Height), LogicalToDeviceUnits(6)),
                  b As New SolidBrush(p.Primary)
                e.Graphics.FillPath(b, path)
            End Using
            Dim textR = Rectangle.Inflate(band, -LogicalToDeviceUnits(10), 0)
            TextRenderer.DrawText(e.Graphics, _title, TitleFont, textR, p.OnPrimary,
                                  TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
        ElseIf Accent OrElse InBrandedHost() Then
            Dim acc = Brand.Accent()
            Dim barW = LogicalToDeviceUnits(4), barH = LogicalToDeviceUnits(18)
            Dim bx = Padding.Left, by = (head - barH) \ 2
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Using path = Draw.RoundedRect(New RectangleF(bx, by, barW, barH), barW / 2.0F), b As New SolidBrush(acc)
                e.Graphics.FillPath(b, path)
            End Using
            Dim tx = bx + barW + LogicalToDeviceUnits(8)
            TextRenderer.DrawText(e.Graphics, _title, TitleFont, New Rectangle(tx, 0, Math.Max(1, Width - tx - Padding.Right), head),
                                  acc, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
        Else
            TextRenderer.DrawText(e.Graphics, _title, TitleFont, New Rectangle(Padding.Left, 0, Width - Padding.Horizontal, head),
                                  p.Ink, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
        End If
    End Sub
End Class

''' <summary>
''' Card whose body can be collapsed by clicking the title. Replaces the accordion groups
''' (horizontal, vertical, convergence, reading analysis, gaze metrics, position metrics).
''' </summary>
Public Class CollapsibleCard
    Inherits Card

    Private _expanded As Boolean = True
    Private _expandedHeight As Integer
    Public Event ExpandedChanged As EventHandler

    Public Sub New()
        Cursor = Cursors.Default
    End Sub

    Public Property Expanded As Boolean
        Get
            Return _expanded
        End Get
        Set(value As Boolean)
            If _expanded = value Then Return
            If Not value Then _expandedHeight = Height
            _expanded = value
            For Each c As Control In Controls
                c.Visible = value
            Next
            If Dock <> DockStyle.Fill Then
                Height = If(value, Math.Max(_expandedHeight, LogicalToDeviceUnits(HeaderHeight) + 20), LogicalToDeviceUnits(HeaderHeight))
            End If
            Invalidate()
            RaiseEvent ExpandedChanged(Me, EventArgs.Empty)
        End Set
    End Property

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        If e.Y < LogicalToDeviceUnits(HeaderHeight) Then Expanded = Not Expanded
        MyBase.OnMouseClick(e)
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        Cursor = If(e.Y < LogicalToDeviceUnits(HeaderHeight), Cursors.Hand, Cursors.Default)
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        If Fonts.Icon IsNot Nothing Then
            Dim glyph = If(_expanded, Glyphs.ChevronDown, Glyphs.ChevronRight)
            TextRenderer.DrawText(e.Graphics, glyph, Fonts.Icon,
                                  New Rectangle(Width - LogicalToDeviceUnits(36), 0, LogicalToDeviceUnits(24), LogicalToDeviceUnits(HeaderHeight)),
                                  ThemeManager.Current.Muted, TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter)
        End If
    End Sub
End Class

#End Region

#Region "Segmented control"

''' <summary>A row of mutually exclusive options drawn as one pill. Replaces tab headers and option buttons.</summary>
<DefaultEvent("SelectedIndexChanged")>
Public Class SegmentedControl
    Inherits Control
    Implements IThemed

    Private ReadOnly _keys As New List(Of String)
    Private ReadOnly _texts As New List(Of String)
    Private _selected As Integer = 0
    Private _hover As Integer = -1
    Private ReadOnly _disabled As New HashSet(Of Integer)

    Public Event SelectedIndexChanged As EventHandler

    ''' <summary>Greys out one option so it can't be chosen (e.g. not enough data for it).</summary>
    Public Sub SetOptionEnabled(index As Integer, enabled As Boolean)
        If enabled Then _disabled.Remove(index) Else _disabled.Add(index)
        Invalidate()
    End Sub

    Public Function IsOptionEnabled(index As Integer) As Boolean
        Return Not _disabled.Contains(index)
    End Function

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.Selectable, True)
        Font = Fonts.Strong
        AutoSize = True
        Margin = New Padding(0, 0, 0, 12)
        Cursor = Cursors.Hand
        TabStop = True
        Loc.OnChange(Me, AddressOf RefreshTexts)
    End Sub

    Public Sub SetOptions(ParamArray keys As String())
        _keys.Clear()
        _keys.AddRange(keys)
        RefreshTexts()
    End Sub

    Private Sub RefreshTexts()
        _texts.Clear()
        For Each k In _keys
            _texts.Add(Loc.T(k))
        Next
        If AutoSize Then Size = GetPreferredSize(Size.Empty)
        Invalidate()
    End Sub

    Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
        Dim w = _texts.Sum(Function(t) TextRenderer.MeasureText(t, Font).Width + LogicalToDeviceUnits(28)) + 4
        Return New Size(Math.Max(w, LogicalToDeviceUnits(40)), LogicalToDeviceUnits(34))
    End Function

    Public Property SelectedIndex As Integer
        Get
            Return _selected
        End Get
        Set(value As Integer)
            If value = _selected Then Return
            _selected = value
            Invalidate()
            RaiseEvent SelectedIndexChanged(Me, EventArgs.Empty)
        End Set
    End Property

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Private Function SegmentRects() As List(Of Rectangle)
        Dim list As New List(Of Rectangle)
        If _texts.Count = 0 Then Return list
        ' Natural widths; when the control is set wider than that (AutoSize off), the segments are
        ' stretched in proportion so together they fill the whole pill.
        Dim widths = _texts.Select(Function(t) TextRenderer.MeasureText(t, Font).Width + LogicalToDeviceUnits(28)).ToList()
        Dim natural = widths.Sum()
        Dim avail = Math.Max(natural, Width - 4)
        Dim x = 2.0
        For i = 0 To widths.Count - 1
            Dim w = widths(i) * avail / CDbl(natural)
            Dim left = CInt(Math.Round(x))
            Dim right = If(i = widths.Count - 1, 2 + avail, CInt(Math.Round(x + w)))
            list.Add(New Rectangle(left, 2, right - left, Height - 4))
            x += w
        Next
        Return list
    End Function

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        Dim rects = SegmentRects()
        Dim h = rects.FindIndex(Function(r) r.Contains(e.Location))
        If _disabled.Contains(h) Then h = -1
        If h <> _hover Then
            _hover = h
            Invalidate()
        End If
        MyBase.OnMouseMove(e)
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        _hover = -1 : Invalidate() : MyBase.OnMouseLeave(e)
    End Sub
    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        Dim idx = SegmentRects().FindIndex(Function(r) r.Contains(e.Location))
        If idx >= 0 AndAlso Enabled AndAlso Not _disabled.Contains(idx) Then SelectedIndex = idx
        MyBase.OnMouseClick(e)
    End Sub
    Protected Overrides Function IsInputKey(keyData As Keys) As Boolean
        If keyData = Keys.Left OrElse keyData = Keys.Right Then Return True
        Return MyBase.IsInputKey(keyData)
    End Function
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.KeyCode = Keys.Left AndAlso _selected > 0 AndAlso Not _disabled.Contains(_selected - 1) Then SelectedIndex -= 1
        If e.KeyCode = Keys.Right AndAlso _selected < _texts.Count - 1 AndAlso Not _disabled.Contains(_selected + 1) Then SelectedIndex += 1
        MyBase.OnKeyDown(e)
    End Sub
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e) : Invalidate()
    End Sub
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        MyBase.OnLostFocus(e) : Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim outer = New RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F)
        Using path = Draw.RoundedRect(outer, LogicalToDeviceUnits(8)), b As New SolidBrush(p.SurfaceAlt), pen As New Pen(If(Focused, p.Primary, p.Border))
            g.FillPath(b, path)
            g.DrawPath(pen, path)
        End Using
        Dim rects = SegmentRects()
        For i = 0 To rects.Count - 1
            Dim r = rects(i)
            Dim sel = i = _selected
            If sel OrElse i = _hover Then
                Using path = Draw.RoundedRect(New RectangleF(r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2), LogicalToDeviceUnits(6)),
                      b As New SolidBrush(If(sel, If(Enabled, p.Primary, p.Disabled), Draw.Mix(p.SurfaceAlt, p.Ink, 0.06)))
                    g.FillPath(b, path)
                End Using
            End If
            TextRenderer.DrawText(g, _texts(i), Font, r, If(sel, p.OnPrimary, If(Enabled AndAlso Not _disabled.Contains(i), p.Ink, p.Disabled)),
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        Next
    End Sub
End Class

#End Region
