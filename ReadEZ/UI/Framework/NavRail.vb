Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   NavRail.vb
'   The vertical navigation bar on the left of the main window. Replaces the tab control of the
'   previous version (Select patient / Calibrate / Run module / Play results ...).
'   Items at the top are screens; items added with atBottom:=True sit at the foot (Settings, Help).
'   The rail can be collapsed to icons only.
'   The Thomson Software Solutions logo sits in full colour on a white panel at the foot of the rail
'   (just its "T" mark when the rail is collapsed).
' ====================================================================================================

Public Class NavRail
    Inherits Control
    Implements IThemed

    Public NotInheritable Class NavItem
        Public Property Id As String
        Public Property Key As String
        Public Property Glyph As String
        Public Property AtBottom As Boolean
        Public Property Visible As Boolean = True
        Public Property Enabled As Boolean = True
        Public Property Badge As String = ""
        Friend Bounds As Rectangle
    End Class

    Private ReadOnly _items As New List(Of NavItem)
    Private _selectedId As String = ""
    Private _hover As NavItem
    Private _collapsed As Boolean
    Private _logo As Image

    Public Event ItemClicked(sender As Object, id As String)
    Public Event CollapsedChanged As EventHandler

    Public Const ExpandedWidth As Integer = 216
    Public Const CollapsedWidth As Integer = 60

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Dock = DockStyle.Left
        Width = ExpandedWidth
        Font = Fonts.Strong
        Loc.OnChange(Me, Sub() Invalidate())
        ThemeManager.SetRole(Me, SurfaceRole.Nav)
    End Sub

    Public Property Logo As Image
        Get
            Return _logo
        End Get
        Set(value As Image)
            _logo = value
            Invalidate()
        End Set
    End Property

    Public Function AddItem(id As String, key As String, glyph As String, Optional atBottom As Boolean = False) As NavItem
        Dim item As New NavItem With {.Id = id, .Key = key, .Glyph = glyph, .AtBottom = atBottom}
        _items.Add(item)
        Invalidate()
        Return item
    End Function

    Public Function Item(id As String) As NavItem
        Return _items.FirstOrDefault(Function(i) i.Id = id)
    End Function

    Public Sub SetVisible(id As String, visible As Boolean)
        Dim i = Item(id)
        If i IsNot Nothing Then i.Visible = visible : Invalidate()
    End Sub

    Public Sub SetEnabled(id As String, enabled As Boolean)
        Dim i = Item(id)
        If i IsNot Nothing Then i.Enabled = enabled : Invalidate()
    End Sub

    Public Property SelectedId As String
        Get
            Return _selectedId
        End Get
        Set(value As String)
            _selectedId = value
            Invalidate()
        End Set
    End Property

    Public Property Collapsed As Boolean
        Get
            Return _collapsed
        End Get
        Set(value As Boolean)
            If _collapsed = value Then Return
            _collapsed = value
            Width = Dpi.Px(Me, If(value, CollapsedWidth, ExpandedWidth))
            Invalidate()
            RaiseEvent CollapsedChanged(Me, EventArgs.Empty)
        End Set
    End Property

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        BackColor = p.NavBack
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

#Region "Layout and input"

    Private ReadOnly Property ItemHeight As Integer
        Get
            Return LogicalToDeviceUnits(44)
        End Get
    End Property

    Private ReadOnly Property TopArea As Integer
        Get
            Return LogicalToDeviceUnits(72)
        End Get
    End Property

    Private ReadOnly Property ToggleRect As Rectangle
        Get
            Return New Rectangle(LogicalToDeviceUnits(10), Height - LogicalToDeviceUnits(44), LogicalToDeviceUnits(40), LogicalToDeviceUnits(36))
        End Get
    End Property

    Private Sub LayoutItems()
        Dim pad = LogicalToDeviceUnits(8)
        Dim y = TopArea
        For Each i In _items.Where(Function(x) x.Visible AndAlso Not x.AtBottom)
            i.Bounds = New Rectangle(pad, y, Width - pad * 2, ItemHeight)
            y += ItemHeight + LogicalToDeviceUnits(2)
        Next
        Dim bottomItems = _items.Where(Function(x) x.Visible AndAlso x.AtBottom).ToList()
        Dim yb = Height - LogicalToDeviceUnits(52) - LogoPanelSpace - bottomItems.Count * (ItemHeight + LogicalToDeviceUnits(2))
        For Each i In bottomItems
            i.Bounds = New Rectangle(pad, yb, Width - pad * 2, ItemHeight)
            yb += ItemHeight + LogicalToDeviceUnits(2)
        Next
    End Sub

    ''' <summary>Height of the logo panel (0 without a logo).</summary>
    Private ReadOnly Property LogoPanelHeight As Integer
        Get
            If _logo Is Nothing Then Return 0
            Return LogicalToDeviceUnits(If(_collapsed, 58, 80))
        End Get
    End Property

    ''' <summary>Room kept above the collapse toggle for the logo panel and the gaps around it.</summary>
    Private ReadOnly Property LogoPanelSpace As Integer
        Get
            Return If(_logo Is Nothing, 0, LogoPanelHeight + LogicalToDeviceUnits(16))
        End Get
    End Property

    Private ReadOnly Property LogoPanelRect As Rectangle
        Get
            Dim pad = LogicalToDeviceUnits(8)
            Dim h = LogoPanelHeight
            Return New Rectangle(pad, ToggleRect.Y - LogicalToDeviceUnits(8) - h, Width - pad * 2, h)
        End Get
    End Property

    Private Function HitTest(pt As Point) As NavItem
        Return _items.FirstOrDefault(Function(i) i.Visible AndAlso i.Bounds.Contains(pt))
    End Function

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        Dim h = HitTest(e.Location)
        If h IsNot _hover Then
            _hover = h
            Invalidate()
        End If
        Cursor = If(h IsNot Nothing AndAlso h.Enabled OrElse ToggleRect.Contains(e.Location), Cursors.Hand, Cursors.Default)
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        _hover = Nothing
        Invalidate()
        MyBase.OnMouseLeave(e)
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        If ToggleRect.Contains(e.Location) Then
            Collapsed = Not Collapsed
            Return
        End If
        Dim h = HitTest(e.Location)
        If h IsNot Nothing AndAlso h.Enabled Then RaiseEvent ItemClicked(Me, h.Id)
        MyBase.OnMouseClick(e)
    End Sub

#End Region

#Region "Painting"

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(p.NavBack)
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit
        LayoutItems()

        ' Product mark
        Dim markR = LogicalToDeviceUnits(14)
        Dim cx = LogicalToDeviceUnits(30)
        Dim cy = TopArea \ 2
        DrawEyeMark(g, New PointF(cx, cy), markR, p.NavFore)
        If Not _collapsed Then
            Using f As New Font("Segoe UI Semibold", 13.0F)
                TextRenderer.DrawText(g, AppHost.ShortTitle, f, New Rectangle(cx + markR + LogicalToDeviceUnits(10), 0, Width - cx - markR, TopArea),
                                      p.NavFore, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
            End Using
        End If

        For Each i In _items.Where(Function(x) x.Visible)
            Dim selected = i.Id = _selectedId
            Dim r = i.Bounds
            If selected OrElse (i Is _hover AndAlso i.Enabled) Then
                Using path = Draw.RoundedRect(New RectangleF(r.X, r.Y, r.Width, r.Height), LogicalToDeviceUnits(8)),
                      b As New SolidBrush(If(selected, p.NavSelected, Color.FromArgb(28, p.NavFore)))
                    g.FillPath(b, path)
                End Using
            End If
            If selected Then
                Using b As New SolidBrush(p.NavFore)
                    g.FillRectangle(b, r.X, r.Y + LogicalToDeviceUnits(10), LogicalToDeviceUnits(3), r.Height - LogicalToDeviceUnits(20))
                End Using
            End If
            Dim fore = If(Not i.Enabled, Color.FromArgb(110, p.NavMuted), If(selected, p.NavFore, p.NavMuted))
            Dim iconRect = New Rectangle(r.X + LogicalToDeviceUnits(6), r.Y, LogicalToDeviceUnits(32), r.Height)
            If Fonts.Icon IsNot Nothing Then
                TextRenderer.DrawText(g, i.Glyph, Fonts.Icon, iconRect, If(selected, p.NavFore, fore),
                                      TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
            End If
            If Not _collapsed Then
                TextRenderer.DrawText(g, Loc.T(i.Key), Font, New Rectangle(iconRect.Right + LogicalToDeviceUnits(8), r.Y, r.Right - iconRect.Right - 12, r.Height),
                                      If(selected, p.NavFore, fore), TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
                If i.Badge <> "" Then
                    Dim bs = TextRenderer.MeasureText(i.Badge, Fonts.SmallStrong)
                    Dim br = New Rectangle(r.Right - bs.Width - LogicalToDeviceUnits(14), r.Y + (r.Height - LogicalToDeviceUnits(20)) \ 2, bs.Width + 8, LogicalToDeviceUnits(20))
                    Using path = Draw.RoundedRect(br, br.Height / 2.0F), b As New SolidBrush(Color.FromArgb(50, p.NavFore))
                        g.FillPath(b, path)
                    End Using
                    TextRenderer.DrawText(g, i.Badge, Fonts.SmallStrong, br, p.NavFore, TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
                End If
            End If
        Next

        ' Company logo: full colour on a white panel above the collapse toggle
        If _logo IsNot Nothing Then DrawCompanyLogo(g)

        ' Collapse toggle
        If Fonts.Icon IsNot Nothing Then
            TextRenderer.DrawText(g, If(_collapsed, Glyphs.Forward, Glyphs.Back), Fonts.Icon, ToggleRect, p.NavMuted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        End If
    End Sub

    ''' <summary>
    ''' The logo is dark grey and blue, so it is drawn on its own white panel to keep its real
    ''' colours on the dark rail in every theme. Collapsed, only the "T" mark is shown.
    ''' </summary>
    Private Sub DrawCompanyLogo(g As Graphics)
        Dim panel = LogoPanelRect
        If panel.Width <= 0 OrElse panel.Height <= 0 Then Return
        Using path = Draw.RoundedRect(New RectangleF(panel.X, panel.Y, panel.Width, panel.Height), LogicalToDeviceUnits(8)),
              b As New SolidBrush(Color.White)
            g.FillPath(b, path)
        End Using
        ' The "T" mark is the left 148 of the logo's 720 pixels
        Dim src = If(_collapsed,
                     New Rectangle(0, 0, CInt(_logo.Width * 148 / 720.0), _logo.Height),
                     New Rectangle(0, 0, _logo.Width, _logo.Height))
        Dim inner = Rectangle.Inflate(panel, -LogicalToDeviceUnits(If(_collapsed, 8, 12)), -LogicalToDeviceUnits(10))
        If inner.Width <= 0 OrElse inner.Height <= 0 Then Return
        Dim scale = Math.Min(inner.Width / CDbl(src.Width), inner.Height / CDbl(src.Height))
        Dim w = CInt(src.Width * scale), h = CInt(src.Height * scale)
        Dim r As New Rectangle(inner.X + (inner.Width - w) \ 2, inner.Y + (inner.Height - h) \ 2, w, h)
        Dim oldMode = g.InterpolationMode
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.DrawImage(_logo, r, src, GraphicsUnit.Pixel)
        g.InterpolationMode = oldMode
    End Sub

    ''' <summary>Simple eye symbol used as the product mark.</summary>
    Public Shared Sub DrawEyeMark(g As Graphics, centre As PointF, r As Single, colour As Color)
        Using pen As New Pen(colour, Math.Max(1.5F, r / 6)), b As New SolidBrush(colour)
            Using path As New GraphicsPath()
                path.AddBezier(centre.X - r * 1.3F, centre.Y, centre.X - r * 0.5F, centre.Y - r * 1.05F, centre.X + r * 0.5F, centre.Y - r * 1.05F, centre.X + r * 1.3F, centre.Y)
                path.AddBezier(centre.X + r * 1.3F, centre.Y, centre.X + r * 0.5F, centre.Y + r * 1.05F, centre.X - r * 0.5F, centre.Y + r * 1.05F, centre.X - r * 1.3F, centre.Y)
                g.DrawPath(pen, path)
            End Using
            g.FillEllipse(b, centre.X - r * 0.42F, centre.Y - r * 0.42F, r * 0.84F, r * 0.84F)
        End Using
    End Sub

#End Region

End Class
