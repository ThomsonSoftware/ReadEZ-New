Option Strict On

Imports System.ComponentModel
Imports System.Drawing.Drawing2D

' ====================================================================================================
'   Buttons.vb
'   ModernButton replaces DevExpress SimpleButton and CheckButton.
'   Set IsToggle = True to make it behave like a CheckButton (Checked / CheckedChanged).
' ====================================================================================================

Public Enum ButtonKind
    Primary
    Secondary
    Ghost
    Danger
    OnDark
    Success
End Enum

''' <summary>Flat, rounded button with an optional icon glyph. Can act as a toggle.</summary>
<DefaultEvent("Click")>
Public Class ModernButton
    Inherits Button
    Implements IThemed

    Private _kind As ButtonKind = ButtonKind.Secondary
    Private _icon As String = ""
    Private _hover As Boolean
    Private _pressed As Boolean
    Private _checked As Boolean
    Private _isToggle As Boolean
    Private _iconOnly As Boolean

    Public Event CheckedChanged As EventHandler

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or
                 ControlStyles.SupportsTransparentBackColor, True)
        FlatStyle = FlatStyle.Flat
        FlatAppearance.BorderSize = 0
        Cursor = Cursors.Hand
        Font = Fonts.Strong
        Height = 36
        Margin = New Padding(0, 0, 8, 0)
        UseVisualStyleBackColor = False
    End Sub

#Region "Properties"

    <DefaultValue(ButtonKind.Secondary)>
    Public Property Kind As ButtonKind
        Get
            Return _kind
        End Get
        Set(value As ButtonKind)
            _kind = value
            Invalidate()
        End Set
    End Property

    <DefaultValue("")>
    Public Property IconGlyph As String
        Get
            Return _icon
        End Get
        Set(value As String)
            _icon = If(value, "")
            Invalidate()
        End Set
    End Property

    ''' <summary>When True the button keeps a pressed state between clicks (like a check button).</summary>
    <DefaultValue(False)>
    Public Property IsToggle As Boolean
        Get
            Return _isToggle
        End Get
        Set(value As Boolean)
            _isToggle = value
            Invalidate()
        End Set
    End Property

    <DefaultValue(False)>
    Public Property Checked As Boolean
        Get
            Return _checked
        End Get
        Set(value As Boolean)
            If _checked = value Then Return
            _checked = value
            Invalidate()
            RaiseEvent CheckedChanged(Me, EventArgs.Empty)
        End Set
    End Property

    ''' <summary>Square button showing only the icon (text becomes the tooltip).</summary>
    <DefaultValue(False)>
    Public Property IconOnly As Boolean
        Get
            Return _iconOnly
        End Get
        Set(value As Boolean)
            _iconOnly = value
            Invalidate()
        End Set
    End Property

    Public Shared Function Create(textKey As String, kind As ButtonKind, Optional icon As String = "") As ModernButton
        Dim b As New ModernButton With {.Kind = kind, .IconGlyph = icon, .AutoSize = True}
        If textKey <> "" Then Loc.Bind(b, textKey)
        Return b
    End Function

    Public Shared Function CreateToggle(textKey As String, Optional icon As String = "") As ModernButton
        Dim b = Create(textKey, ButtonKind.Secondary, icon)
        b.IsToggle = True
        Return b
    End Function

    Public Shared Function CreateIcon(icon As String, tip As ToolTip, tipKey As String, Optional kind As ButtonKind = ButtonKind.Ghost) As ModernButton
        Dim b As New ModernButton With {.Kind = kind, .IconGlyph = icon, .IconOnly = True, .Size = New Size(34, 34), .Margin = New Padding(2)}
        If tip IsNot Nothing AndAlso tipKey <> "" Then Loc.BindTip(tip, b, tipKey)
        Return b
    End Function

#End Region

#Region "IThemed"

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

#End Region

#Region "Mouse"

    Protected Overrides Sub OnClick(e As EventArgs)
        If _isToggle Then Checked = Not _checked
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        _hover = True : Invalidate() : MyBase.OnMouseEnter(e)
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        _hover = False : _pressed = False : Invalidate() : MyBase.OnMouseLeave(e)
    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        _pressed = True : Invalidate() : MyBase.OnMouseDown(e)
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        _pressed = False : Invalidate() : MyBase.OnMouseUp(e)
    End Sub
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        Invalidate() : MyBase.OnEnabledChanged(e)
    End Sub
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        If AutoSize Then Size = GetPreferredSize(Size.Empty)
        Invalidate()
    End Sub

#End Region

#Region "Painting"

    Private ReadOnly Property HasIcon As Boolean
        Get
            Return _icon <> "" AndAlso Fonts.Icon IsNot Nothing
        End Get
    End Property

    Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
        If _iconOnly Then Return New Size(LogicalToDeviceUnits(34), LogicalToDeviceUnits(34))
        Dim textSize = TextRenderer.MeasureText(Text, Font)
        Dim w = textSize.Width + LogicalToDeviceUnits(28)
        If HasIcon Then w += TextRenderer.MeasureText(_icon, Fonts.Icon).Width + LogicalToDeviceUnits(2)
        Return New Size(w, LogicalToDeviceUnits(36))
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias

        Dim back As Color, fore As Color, edge As Color = Color.Empty
        Dim isOn = _isToggle AndAlso _checked
        Select Case _kind
            Case ButtonKind.Primary
                back = If(_pressed, p.PrimaryPressed, If(_hover, p.PrimaryHover, p.Primary))
                fore = p.OnPrimary
            Case ButtonKind.Success
                back = If(_hover OrElse _pressed, Draw.Mix(p.Success, Color.Black, 0.12), p.Success)
                fore = Color.White
            Case ButtonKind.Danger
                back = If(_hover OrElse _pressed, p.DangerSoft, p.Surface)
                fore = p.Danger
                edge = Draw.Mix(p.Danger, p.Surface, 0.6)
            Case ButtonKind.Ghost
                back = If(_pressed, p.PrimarySoft, If(_hover, Draw.Mix(Draw.ParentBack(Me), p.Ink, 0.08), Color.Transparent))
                fore = p.Ink
            Case ButtonKind.OnDark
                back = If(_pressed, Color.FromArgb(60, 255, 255, 255), If(_hover, Color.FromArgb(35, 255, 255, 255), Color.Transparent))
                fore = p.NavFore
            Case Else
                back = If(_pressed, p.PrimarySoft, If(_hover, p.SurfaceAlt, p.Surface))
                fore = p.Ink
                edge = p.Border
        End Select
        If isOn Then
            back = If(_hover, p.PrimaryHover, p.Primary)
            fore = p.OnPrimary
            edge = Color.Empty
        End If
        If Not Enabled Then
            back = If(_kind = ButtonKind.Primary, Draw.Mix(p.Primary, p.Surface, 0.55),
                   If(_kind = ButtonKind.OnDark OrElse _kind = ButtonKind.Ghost, Color.Transparent, p.SurfaceAlt))
            fore = If(_kind = ButtonKind.Primary, Draw.Mix(p.OnPrimary, p.Surface, 0.3), p.Disabled)
        End If

        Dim rect = New RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F)
        Using path = Draw.RoundedRect(rect, LogicalToDeviceUnits(6))
            If back.A > 0 Then
                Using b As New SolidBrush(back)
                    g.FillPath(b, path)
                End Using
            End If
            If Not edge.IsEmpty AndAlso Enabled Then
                Using pen As New Pen(edge)
                    g.DrawPath(pen, path)
                End Using
            End If
            If Focused AndAlso ShowFocusCues Then
                Using pen As New Pen(If(_kind = ButtonKind.OnDark, p.NavFore, p.Primary), 2)
                    g.DrawPath(pen, path)
                End Using
            End If
        End Using

        If _iconOnly Then
            If HasIcon Then
                TextRenderer.DrawText(g, _icon, Fonts.Icon, ClientRectangle, fore,
                                      TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
            End If
            Return
        End If

        Dim textSize = TextRenderer.MeasureText(g, Text, Font)
        Dim iconWidth = 0
        If HasIcon Then iconWidth = TextRenderer.MeasureText(g, _icon, Fonts.Icon).Width
        Dim total = textSize.Width + iconWidth
        ' TextAlign MiddleLeft gives a list-style button (icon and text from the left edge)
        Dim leftAligned = TextAlign = ContentAlignment.MiddleLeft OrElse TextAlign = ContentAlignment.TopLeft OrElse TextAlign = ContentAlignment.BottomLeft
        Dim x = If(leftAligned, LogicalToDeviceUnits(12), Math.Max(4, (Width - total) \ 2))
        If HasIcon Then
            TextRenderer.DrawText(g, _icon, Fonts.Icon, New Rectangle(x, 0, iconWidth, Height), fore,
                                  TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding Or TextFormatFlags.Left)
            x += iconWidth
        End If
        TextRenderer.DrawText(g, Text, Font, New Rectangle(x, 0, Math.Max(0, Width - x - 4), Height), fore,
                              TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
    End Sub

#End Region

End Class

''' <summary>
''' Large tile button used on the Run tests screen: icon, title and a one-line description.
''' </summary>
Public Class TileButton
    Inherits Control
    Implements IThemed

    Private _hover As Boolean
    Private _pressed As Boolean
    Private _description As String = ""

    Public Property Glyph As String = Glyphs.Eye
    Public Property AccentColor As Color = Color.Empty

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.Selectable, True)
        Cursor = Cursors.Hand
        Size = New Size(230, 104)
        Margin = New Padding(0, 0, 14, 14)
        TabStop = True
    End Sub

    Public Property Description As String
        Get
            Return _description
        End Get
        Set(value As String)
            _description = If(value, "")
            Invalidate()
        End Set
    End Property

    Public Shared Function Create(titleKey As String, descriptionKey As String, glyph As String) As TileButton
        Dim t As New TileButton With {.Glyph = glyph}
        Loc.Bind(t, titleKey)
        Loc.OnChange(t, Sub() t.Description = Loc.T(descriptionKey))
        Return t
    End Function

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e) : Invalidate()
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        _hover = True : Invalidate() : MyBase.OnMouseEnter(e)
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        _hover = False : _pressed = False : Invalidate() : MyBase.OnMouseLeave(e)
    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        _pressed = True : Focus() : Invalidate() : MyBase.OnMouseDown(e)
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        _pressed = False : Invalidate() : MyBase.OnMouseUp(e)
    End Sub
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        Invalidate() : MyBase.OnEnabledChanged(e)
    End Sub
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        Invalidate() : MyBase.OnGotFocus(e)
    End Sub
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        Invalidate() : MyBase.OnLostFocus(e)
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Space Then OnClick(EventArgs.Empty)
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim accent = If(AccentColor.IsEmpty, p.Primary, AccentColor)
        Dim back = If(_pressed, p.PrimarySoft, If(_hover, Draw.Mix(p.Surface, accent, 0.06), p.Surface))
        Dim edge = If(_hover OrElse Focused, accent, p.Border)
        Dim rect = New RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F)
        Using path = Draw.RoundedRect(rect, LogicalToDeviceUnits(10))
            Using b As New SolidBrush(back)
                g.FillPath(b, path)
            End Using
            Using pen As New Pen(edge, If(_hover OrElse Focused, 1.6F, 1.0F))
                g.DrawPath(pen, path)
            End Using
        End Using

        Dim pad = LogicalToDeviceUnits(16)
        Dim circle = LogicalToDeviceUnits(40)
        Dim circleRect = New Rectangle(pad, pad, circle, circle)
        Using b As New SolidBrush(If(Enabled, Draw.Mix(p.Surface, accent, 0.16), p.SurfaceAlt))
            g.FillEllipse(b, circleRect)
        End Using
        If Fonts.IconLarge IsNot Nothing Then
            TextRenderer.DrawText(g, Glyph, Fonts.IconLarge, circleRect, If(Enabled, accent, p.Disabled),
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter Or TextFormatFlags.NoPadding)
        End If
        Dim tx = pad + circle + LogicalToDeviceUnits(12)
        Dim titleRect = New Rectangle(tx, pad, Width - tx - pad \ 2, LogicalToDeviceUnits(22))
        ' A title too wide for the tile (some translations) uses the smaller font before it is cut off
        Dim titleFont = If(TextRenderer.MeasureText(g, Text, Fonts.Section, titleRect.Size, TextFormatFlags.NoPadding).Width > titleRect.Width,
                           Fonts.Strong, Fonts.Section)
        TextRenderer.DrawText(g, Text, titleFont, titleRect, If(Enabled, p.Ink, p.Disabled),
                              TextFormatFlags.Left Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPadding)
        Dim descRect = New Rectangle(tx, titleRect.Bottom + LogicalToDeviceUnits(4), Width - tx - pad \ 2, Height - titleRect.Bottom - pad \ 2)
        TextRenderer.DrawText(g, _description, Fonts.Small, descRect, p.Muted,
                              TextFormatFlags.Left Or TextFormatFlags.WordBreak Or TextFormatFlags.EndEllipsis Or TextFormatFlags.NoPadding)
    End Sub
End Class
