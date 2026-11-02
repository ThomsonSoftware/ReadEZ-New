Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   ColourWheel.vb - the twelve ReadEZ colours round a circle, grey in the middle
'
'   Used to pick a colour (manual test), to show the automatic test's result ("spokes": how often
'   each colour was chosen) and on the Results screen. With two rings (lenses and computer) the
'   outer ring is the light lens set and the inner ring the dark set, as in ReadEZ 2020.
' ====================================================================================================

Public NotInheritable Class ColourWheel
    Inherits Control
    Implements IThemed

    ''' <summary>The outer ring's set (overlays or lenses).</summary>
    Public Property OuterSet As ReadEZSet = ReadEZSet.Overlay
    ''' <summary>Show the dark lens set as an inner ring.</summary>
    Public Property TwoRings As Boolean
    Public Property Selectable As Boolean = True
    ''' <summary>The marked colour (Nothing = none).</summary>
    Public Property Selected As (SetOf As ReadEZSet, Index As Integer)? = Nothing
    ''' <summary>A second marked colour (the other side, overlays), drawn dashed.</summary>
    Public Property Second As (SetOf As ReadEZSet, Index As Integer)? = Nothing
    ''' <summary>How often each colour was chosen (spokes; Nothing = none).</summary>
    Public Property Spokes As Integer() = Nothing

    Public Event ColourClicked(s As ReadEZSet, index As Integer, button As MouseButtons)

    Private _hover As (SetOf As ReadEZSet, Index As Integer)? = Nothing

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Size = New Size(280, 280)
        AddHandler ReadEZColours.Changed, AddressOf ColoursChanged
    End Sub

    Private Sub ColoursChanged(sender As Object, e As EventArgs)
        If Not IsDisposed Then Invalidate()
    End Sub

    Private Function Geometry() As (Cx As Single, Cy As Single, R As Single, D As Single)
        Dim size = Math.Min(Width, Height)
        Dim r = size * If(TwoRings, 0.4F, 0.37F)
        Return (Width / 2.0F, Height / 2.0F, r, size * If(TwoRings, 0.15F, 0.18F))
    End Function

    ''' <summary>Where each swatch is: (set, index, rectangle).</summary>
    Private Iterator Function Swatches() As IEnumerable(Of (SetOf As ReadEZSet, Index As Integer, Rect As RectangleF))
        Dim g = Geometry()
        For i = 1 To 12
            Dim a = -Math.PI / 2 + (i - 1) * Math.PI / 6
            Yield (OuterSet, i, New RectangleF(g.Cx + CSng(g.R * Math.Cos(a)) - g.D / 2, g.Cy + CSng(g.R * Math.Sin(a)) - g.D / 2, g.D, g.D))
            If TwoRings Then
                Dim r2 = g.R * 0.58F
                Yield (ReadEZSet.LensStrong, i, New RectangleF(g.Cx + CSng(r2 * Math.Cos(a)) - g.D * 0.42F, g.Cy + CSng(r2 * Math.Sin(a)) - g.D * 0.42F, g.D * 0.84F, g.D * 0.84F))
            End If
        Next
        Dim dg = g.D * If(TwoRings, 0.7F, 0.9F)
        Yield (OuterSet, 0, New RectangleF(g.Cx - dg / 2, g.Cy - dg / 2, dg, dg))
    End Function

    Private Function HitTest(p As Point) As (SetOf As ReadEZSet, Index As Integer)?
        For Each s In Swatches()
            Dim c = New PointF(s.Rect.X + s.Rect.Width / 2, s.Rect.Y + s.Rect.Height / 2)
            If (p.X - c.X) ^ 2 + (p.Y - c.Y) ^ 2 <= (s.Rect.Width / 2) ^ 2 Then Return (s.SetOf, s.Index)
        Next
        Return Nothing
    End Function

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If Not Selectable Then Return
        Dim h = HitTest(e.Location)
        If Not Nullable.Equals(h, _hover) Then
            _hover = h
            Cursor = If(h.HasValue, Cursors.Hand, Cursors.Default)
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _hover = Nothing
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        If Not Selectable Then Return
        Dim h = HitTest(e.Location)
        If h.HasValue Then RaiseEvent ColourClicked(h.Value.SetOf, h.Value.Index, e.Button)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim geo = Geometry()
        ' Spokes first, under the swatches
        If Spokes IsNot Nothing AndAlso Spokes.Length >= 13 Then
            Dim most = Math.Max(1, Spokes.Skip(1).Max())
            Using pen As New Pen(p.Muted, 4) With {.StartCap = LineCap.Round, .EndCap = LineCap.Round}
                For i = 1 To 12
                    If Spokes(i) <= 0 Then Continue For
                    Dim a = -Math.PI / 2 + (i - 1) * Math.PI / 6
                    Dim len = geo.R * Spokes(i) / most
                    g.DrawLine(pen, geo.Cx, geo.Cy, geo.Cx + CSng(len * Math.Cos(a)), geo.Cy + CSng(len * Math.Sin(a)))
                Next
            End Using
        End If
        Using f As New Font("Segoe UI Semibold", Math.Max(7, geo.D * 0.3F), GraphicsUnit.Pixel), edge As New Pen(Color.FromArgb(120, p.Ink), 1),
              sel As New Pen(p.Ink, 3), sel2 As New Pen(p.Ink, 2) With {.DashStyle = DashStyle.Dash}, hov As New Pen(p.Primary, 2)
            Dim c As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            For Each s In Swatches()
                Using b As New SolidBrush(ReadEZColours.Colour(s.SetOf, s.Index))
                    g.FillEllipse(b, s.Rect)
                End Using
                g.DrawEllipse(edge, s.Rect)
                Dim isSel = Selected.HasValue AndAlso Selected.Value.Index = s.Index AndAlso (Selected.Value.SetOf = s.SetOf OrElse s.Index = 0)
                Dim isSecond = Second.HasValue AndAlso Second.Value.Index = s.Index AndAlso (Second.Value.SetOf = s.SetOf OrElse s.Index = 0)
                If isSel Then g.DrawEllipse(sel, RectangleF.Inflate(s.Rect, 3, 3))
                If isSecond Then g.DrawEllipse(sel2, RectangleF.Inflate(s.Rect, 6, 6))
                If _hover.HasValue AndAlso _hover.Value.SetOf = s.SetOf AndAlso _hover.Value.Index = s.Index Then g.DrawEllipse(hov, RectangleF.Inflate(s.Rect, 1.5F, 1.5F))
                If s.SetOf <> ReadEZSet.LensStrong Then
                    g.DrawString(If(s.Index = 0, "–", ReadEZColours.Letter(s.Index)), f, Brushes.Black, s.Rect, c)
                End If
            Next
        End Using
    End Sub

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then RemoveHandler ReadEZColours.Changed, AddressOf ColoursChanged
        MyBase.Dispose(disposing)
    End Sub
End Class
