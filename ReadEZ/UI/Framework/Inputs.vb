Option Strict On

Imports System.ComponentModel
Imports System.Drawing.Drawing2D

' ====================================================================================================
'   Inputs.vb
'   Themed input controls. These replace the DevExpress editors used previously:
'     TextEdit          -> TextField
'     ToggleSwitch      -> ToggleSwitch
'     TrackBarControl   -> Slider (supports value labels like TrackBarLabel)
'     CheckEdit         -> ThemedCheckBox
'     ColorPickEdit     -> ColorButton
'     FontEdit          -> FontButton
' ====================================================================================================

#Region "TextField"

''' <summary>Borderless text box inside a padded, outlined panel.</summary>
Public Class TextField
    Inherits Panel
    Implements IThemed

    Public ReadOnly Property Box As New TextBox
    Private _invalid As Boolean

    Public Sub New(Optional multiline As Boolean = False)
        SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
        Padding = New Padding(8, 7, 8, 5)
        Height = If(multiline, 90, 34)
        Margin = New Padding(0, 0, 0, 12)
        Box.BorderStyle = BorderStyle.None
        Box.Font = Fonts.Body
        Box.Dock = DockStyle.Fill
        Box.Multiline = multiline
        If multiline Then Box.ScrollBars = ScrollBars.Vertical
        Controls.Add(Box)
        AddHandler Box.GotFocus, Sub() Invalidate()
        AddHandler Box.LostFocus, Sub() Invalidate()
        ApplyTheme(ThemeManager.Current)
    End Sub

    Public Overrides Property Text As String
        Get
            Return Box.Text
        End Get
        Set(value As String)
            Box.Text = value
        End Set
    End Property

    Public Property Invalid As Boolean
        Get
            Return _invalid
        End Get
        Set(value As Boolean)
            _invalid = value
            Invalidate()
        End Set
    End Property

    ''' <summary>Grey hint text shown when the box is empty (translated).</summary>
    Public Sub SetPlaceholder(key As String)
#If Not COMPILECHECK Then
        Loc.OnChange(Me, Sub() Box.PlaceholderText = Loc.T(key))
#End If
    End Sub

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        BackColor = p.InputBack
        Box.BackColor = p.InputBack
        Box.ForeColor = p.Ink
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub OnClick(e As EventArgs)
        Box.Focus()
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim p = ThemeManager.Current
        Dim colour = If(_invalid, p.Danger, If(Box.Focused, p.Primary, p.Border))
        Dim thick = Box.Focused OrElse _invalid
        Using pen As New Pen(colour, If(thick, 2, 1))
            If thick Then
                e.Graphics.DrawRectangle(pen, 1, 1, Width - 2, Height - 2)
            Else
                e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1)
            End If
        End Using
    End Sub
End Class

#End Region

#Region "ToggleSwitch"

''' <summary>On/off switch with an optional caption to its right.</summary>
<DefaultEvent("Toggled")>
Public Class ToggleSwitch
    Inherits Control
    Implements IThemed

    Private _checked As Boolean
    Private _hover As Boolean
    Private _anim As Single
    Private ReadOnly _timer As New Timer With {.Interval = 15}

    ''' <summary>Raised when the user (or code) changes the state. Same name as DevExpress ToggleSwitch.</summary>
    Public Event Toggled As EventHandler

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.Selectable Or ControlStyles.SupportsTransparentBackColor, True)
        Cursor = Cursors.Hand
        Font = Fonts.Body
        Height = 28
        Width = 180
        TabStop = True
        AddHandler _timer.Tick, AddressOf Animate
    End Sub

    <DefaultValue(False)>
    Public Property IsOn As Boolean
        Get
            Return _checked
        End Get
        Set(value As Boolean)
            If _checked = value Then Return
            _checked = value
            _timer.Start()
            RaiseEvent Toggled(Me, EventArgs.Empty)
        End Set
    End Property

    ''' <summary>Alias so code written for check boxes reads naturally.</summary>
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property Checked As Boolean
        Get
            Return IsOn
        End Get
        Set(value As Boolean)
            IsOn = value
        End Set
    End Property

    ''' <summary>Sets the state without animating or raising Toggled (used when loading settings).</summary>
    Public Sub SetSilently(value As Boolean)
        _checked = value
        _anim = If(value, 1.0F, 0.0F)
        Invalidate()
    End Sub

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Private Sub Animate(sender As Object, e As EventArgs)
        Dim target = If(_checked, 1.0F, 0.0F)
        _anim += If(target > _anim, 0.2F, -0.2F)
        If Math.Abs(_anim - target) < 0.2F Then
            _anim = target
            _timer.Stop()
        End If
        Invalidate()
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        IsOn = Not IsOn
        MyBase.OnClick(e)
    End Sub
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.KeyCode = Keys.Space Then IsOn = Not IsOn
        MyBase.OnKeyDown(e)
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        _hover = True : Invalidate() : MyBase.OnMouseEnter(e)
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        _hover = False : Invalidate() : MyBase.OnMouseLeave(e)
    End Sub
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e) : Invalidate()
    End Sub
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e) : Invalidate()
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
        Dim parentBack = Draw.ParentBack(Me)
        g.Clear(parentBack)
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim h = LogicalToDeviceUnits(20)
        Dim w = LogicalToDeviceUnits(38)
        Dim y = (Height - h) \ 2
        Dim track = New RectangleF(1, y, w, h)
        Dim offColour = If(p.IsDark, Color.FromArgb(80, 86, 94), Draw.Mix(p.Border, p.Muted, 0.35))
        Dim onColour = If(_hover, p.PrimaryHover, p.Primary)
        Dim fill = If(Enabled, Draw.Mix(offColour, onColour, _anim), p.Disabled)
        Using path = Draw.RoundedRect(track, h / 2.0F), b As New SolidBrush(fill)
            g.FillPath(b, path)
        End Using
        Dim knob = h - LogicalToDeviceUnits(6)
        Dim kx = track.X + LogicalToDeviceUnits(3) + (w - knob - LogicalToDeviceUnits(6)) * _anim
        Using b As New SolidBrush(Color.White)
            g.FillEllipse(b, kx, y + LogicalToDeviceUnits(3), knob, knob)
        End Using
        If Focused AndAlso ShowFocusCues Then
            Using pen As New Pen(p.Primary, 1), path = Draw.RoundedRect(New RectangleF(track.X - 2, track.Y - 2, track.Width + 4, track.Height + 4), (h + 4) / 2.0F)
                g.DrawPath(pen, path)
            End Using
        End If
        If Text <> "" Then
            Dim tx = CInt(track.Right) + LogicalToDeviceUnits(10)
            TextRenderer.DrawText(g, Text, Font, New Rectangle(tx, 0, Width - tx, Height), If(Enabled, ForeColorFor(p), p.Disabled),
                                  TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
        End If
    End Sub

    Private Function ForeColorFor(p As Palette) As Color
        Dim role = SurfaceRole.Auto
        Dim c As Control = Parent
        While c IsNot Nothing AndAlso role = SurfaceRole.Auto
            role = ThemeManager.GetRole(c)
            c = c.Parent
        End While
        Return If(role = SurfaceRole.Nav, p.NavFore, p.Ink)
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _timer.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class

#End Region

#Region "Slider"

''' <summary>
''' Horizontal slider with optional labelled ticks and value read-out.
''' Replaces DevExpress TrackBarControl; Labels replaces TrackBarLabel.
''' </summary>
<DefaultEvent("ValueChanged")>
Public Class Slider
    Inherits Control
    Implements IThemed

    Private _min As Integer = 0
    Private _max As Integer = 100
    Private _value As Integer = 0
    Private _dragging As Boolean
    Private _hover As Boolean

    Public Event ValueChanged As EventHandler
    ''' <summary>Raised only for user changes (mouse/keyboard), like TrackBar.Scroll.</summary>
    Public Event Scroll As EventHandler

    ''' <summary>Value/text pairs drawn under the track.</summary>
    Public ReadOnly Property Labels As New List(Of KeyValuePair(Of Integer, String))
    Public Property SmallChange As Integer = 1
    Public Property LargeChange As Integer = 5
    Public Property ShowValue As Boolean = False
    Public Property ValueFormat As Func(Of Integer, String) = Function(v) v.ToString()

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or
                 ControlStyles.ResizeRedraw Or ControlStyles.Selectable, True)
        Height = 40
        Width = 200
        TabStop = True
        Font = Fonts.Small
        Cursor = Cursors.Hand
    End Sub

    Public Property Minimum As Integer
        Get
            Return _min
        End Get
        Set(value As Integer)
            _min = value
            If _max < _min Then _max = _min
            Me.Value = _value
            Invalidate()
        End Set
    End Property

    Public Property Maximum As Integer
        Get
            Return _max
        End Get
        Set(value As Integer)
            _max = Math.Max(value, _min)
            Me.Value = _value
            Invalidate()
        End Set
    End Property

    Public Property Value As Integer
        Get
            Return _value
        End Get
        Set(v As Integer)
            v = Math.Max(_min, Math.Min(_max, v))
            If v = _value Then Return
            _value = v
            Invalidate()
            RaiseEvent ValueChanged(Me, EventArgs.Empty)
        End Set
    End Property

    Public Sub AddLabel(value As Integer, text As String)
        Labels.Add(New KeyValuePair(Of Integer, String)(value, text))
        Invalidate()
    End Sub

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Private ReadOnly Property TrackLeft As Integer
        Get
            Return LogicalToDeviceUnits(9)
        End Get
    End Property

    Private ReadOnly Property TrackRight As Integer
        Get
            Dim right = Width - LogicalToDeviceUnits(9)
            If ShowValue Then right -= LogicalToDeviceUnits(48)
            Return right
        End Get
    End Property

    Private ReadOnly Property TrackY As Integer
        Get
            Return If(Labels.Count > 0, LogicalToDeviceUnits(12), Height \ 2)
        End Get
    End Property

    Private Function XFor(v As Integer) As Single
        If _max = _min Then Return TrackLeft
        Return TrackLeft + CSng((v - _min) / (_max - _min)) * (TrackRight - TrackLeft)
    End Function

    Private Function ValueAt(x As Integer) As Integer
        Dim frac = (x - TrackLeft) / Math.Max(1.0, TrackRight - TrackLeft)
        Return CInt(Math.Round(_min + Math.Max(0, Math.Min(1, frac)) * (_max - _min)))
    End Function

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            _dragging = True
            Focus()
            SetFromUser(ValueAt(e.X))
        End If
        MyBase.OnMouseDown(e)
    End Sub
    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        If _dragging Then SetFromUser(ValueAt(e.X))
        MyBase.OnMouseMove(e)
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        _dragging = False
        MyBase.OnMouseUp(e)
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        _hover = True : Invalidate() : MyBase.OnMouseEnter(e)
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        _hover = False : Invalidate() : MyBase.OnMouseLeave(e)
    End Sub
    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        ' Windows sends the wheel to the control under the pointer, so an unfocused slider would
        ' change value while the user scrolls the page past it. Only a focused slider takes the wheel.
        If Not Focused Then
            MyBase.OnMouseWheel(e)
            Return
        End If
        SetFromUser(_value + If(e.Delta > 0, SmallChange, -SmallChange))
        Dim h = TryCast(e, HandledMouseEventArgs)
        If h IsNot Nothing Then h.Handled = True      ' don't scroll the page as well
        MyBase.OnMouseWheel(e)
    End Sub
    Protected Overrides Function IsInputKey(keyData As Keys) As Boolean
        If keyData = Keys.Left OrElse keyData = Keys.Right OrElse keyData = Keys.Up OrElse keyData = Keys.Down Then Return True
        Return MyBase.IsInputKey(keyData)
    End Function
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Left, Keys.Down : SetFromUser(_value - SmallChange)
            Case Keys.Right, Keys.Up : SetFromUser(_value + SmallChange)
            Case Keys.PageDown : SetFromUser(_value - LargeChange)
            Case Keys.PageUp : SetFromUser(_value + LargeChange)
            Case Keys.Home : SetFromUser(_min)
            Case Keys.End : SetFromUser(_max)
        End Select
        MyBase.OnKeyDown(e)
    End Sub
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e) : Invalidate()
    End Sub
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        MyBase.OnLostFocus(e) : Invalidate()
    End Sub
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e) : Invalidate()
    End Sub

    Private Sub SetFromUser(v As Integer)
        Dim before = _value
        Value = v
        If _value <> before Then RaiseEvent Scroll(Me, EventArgs.Empty)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim y = TrackY
        Dim thick = LogicalToDeviceUnits(4)
        Dim xv = XFor(_value)
        Dim accent = If(Enabled, If(_hover OrElse _dragging, p.PrimaryHover, p.Primary), p.Disabled)
        Using path = Draw.RoundedRect(New RectangleF(TrackLeft, y - thick / 2.0F, TrackRight - TrackLeft, thick), thick / 2.0F),
              b As New SolidBrush(If(p.IsDark, Color.FromArgb(70, 76, 84), p.Border))
            g.FillPath(b, path)
        End Using
        If xv > TrackLeft Then
            Using path = Draw.RoundedRect(New RectangleF(TrackLeft, y - thick / 2.0F, xv - TrackLeft, thick), thick / 2.0F),
                  b As New SolidBrush(accent)
                g.FillPath(b, path)
            End Using
        End If
        ' Labels
        If Labels.Count > 0 Then
            Using tick As New Pen(p.Muted)
                For Each kv In Labels
                    Dim lx = XFor(kv.Key)
                    g.DrawLine(tick, lx, y + thick, lx, y + thick + LogicalToDeviceUnits(4))
                    Dim sz = TextRenderer.MeasureText(kv.Value, Font)
                    Dim rx = CInt(Math.Max(0, Math.Min(Width - sz.Width, lx - sz.Width / 2.0F)))
                    TextRenderer.DrawText(g, kv.Value, Font, New Point(rx, y + thick + LogicalToDeviceUnits(6)), p.Muted)
                Next
            End Using
        End If
        ' Thumb
        Dim r = LogicalToDeviceUnits(8)
        Using b As New SolidBrush(If(p.IsDark, p.Surface, Color.White)), pen As New Pen(accent, LogicalToDeviceUnits(2))
            g.FillEllipse(b, xv - r, y - r, r * 2, r * 2)
            g.DrawEllipse(pen, xv - r, y - r, r * 2, r * 2)
        End Using
        If Focused AndAlso ShowFocusCues Then
            Using pen As New Pen(Color.FromArgb(90, accent), 3)
                g.DrawEllipse(pen, xv - r - 3, y - r - 3, r * 2 + 6, r * 2 + 6)
            End Using
        End If
        If ShowValue Then
            TextRenderer.DrawText(g, ValueFormat(_value), Fonts.SmallStrong,
                                  New Rectangle(TrackRight + LogicalToDeviceUnits(10), y - 10, LogicalToDeviceUnits(48), 20),
                                  p.Ink, TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
        End If
    End Sub
End Class

#End Region

#Region "Check box and radio button"

''' <summary>Check box drawn from the palette so it looks right in dark mode.</summary>
Public Class ThemedCheckBox
    Inherits CheckBox
    Implements IThemed

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Font = Fonts.Body
        AutoSize = True
        Cursor = Cursors.Hand
        Padding = New Padding(24, 2, 0, 2)
        Margin = New Padding(0, 4, 12, 4)
    End Sub

    Public Shared Function Create(key As String) As ThemedCheckBox
        Return Loc.Bind(New ThemedCheckBox(), key)
    End Function

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
        Dim sz = TextRenderer.MeasureText(Text, Font)
        Return New Size(sz.Width + LogicalToDeviceUnits(28), Math.Max(sz.Height + 6, LogicalToDeviceUnits(24)))
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim s = LogicalToDeviceUnits(16)
        Dim box = New RectangleF(1, (Height - s) / 2.0F, s, s)
        Using path = Draw.RoundedRect(box, LogicalToDeviceUnits(3))
            If Checked Then
                Using b As New SolidBrush(If(Enabled, p.Primary, p.Disabled))
                    g.FillPath(b, path)
                End Using
                Using pen As New Pen(p.OnPrimary, LogicalToDeviceUnits(2))
                    g.DrawLines(pen, {New PointF(box.X + s * 0.22F, box.Y + s * 0.52F),
                                      New PointF(box.X + s * 0.43F, box.Y + s * 0.72F),
                                      New PointF(box.X + s * 0.78F, box.Y + s * 0.3F)})
                End Using
            Else
                Using b As New SolidBrush(p.InputBack), pen As New Pen(If(Focused, p.Primary, p.Muted))
                    g.FillPath(b, path)
                    g.DrawPath(pen, path)
                End Using
            End If
        End Using
        TextRenderer.DrawText(g, Text, Font, New Rectangle(s + LogicalToDeviceUnits(8), 0, Width - s - 8, Height),
                              If(Enabled, p.Ink, p.Disabled), TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
    End Sub
End Class

''' <summary>Radio button drawn from the palette.</summary>
Public Class ThemedRadioButton
    Inherits RadioButton
    Implements IThemed

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Font = Fonts.Body
        AutoSize = True
        Cursor = Cursors.Hand
        Margin = New Padding(0, 4, 12, 4)
    End Sub

    Public Shared Function Create(key As String) As ThemedRadioButton
        Return Loc.Bind(New ThemedRadioButton(), key)
    End Function

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
        Dim sz = TextRenderer.MeasureText(Text, Font)
        Return New Size(sz.Width + LogicalToDeviceUnits(28), Math.Max(sz.Height + 6, LogicalToDeviceUnits(24)))
    End Function

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim s = LogicalToDeviceUnits(16)
        Dim box = New RectangleF(1, (Height - s) / 2.0F, s, s)
        Using b As New SolidBrush(p.InputBack), pen As New Pen(If(Checked, If(Enabled, p.Primary, p.Disabled), If(Focused, p.Primary, p.Muted)), If(Checked, 2, 1))
            g.FillEllipse(b, box)
            g.DrawEllipse(pen, box)
        End Using
        If Checked Then
            Dim inner = s * 0.5F
            Using b As New SolidBrush(If(Enabled, p.Primary, p.Disabled))
                g.FillEllipse(b, box.X + (s - inner) / 2, box.Y + (s - inner) / 2, inner, inner)
            End Using
        End If
        TextRenderer.DrawText(g, Text, Font, New Rectangle(s + LogicalToDeviceUnits(8), 0, Width - s - 8, Height),
                              If(Enabled, p.Ink, p.Disabled), TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
    End Sub
End Class

#End Region

#Region "Combo box"

''' <summary>Flat drop-down list that follows the palette.</summary>
Public Class ThemedComboBox
    Inherits ComboBox
    Implements IThemed

    Public Sub New()
        DropDownStyle = ComboBoxStyle.DropDownList
        FlatStyle = FlatStyle.Flat
        DrawMode = DrawMode.OwnerDrawFixed
        Font = Fonts.Body
        ItemHeight = Font.Height + 8
        Margin = New Padding(0, 0, 0, 12)
        ApplyTheme(ThemeManager.Current)
    End Sub

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        BackColor = p.InputBack
        ForeColor = p.Ink
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        Dim p = ThemeManager.Current
        If e.Index < 0 Then Return
        Dim selected = (e.State And DrawItemState.Selected) = DrawItemState.Selected AndAlso (e.State And DrawItemState.ComboBoxEdit) <> DrawItemState.ComboBoxEdit
        Using b As New SolidBrush(If(selected, p.PrimarySoft, p.InputBack))
            e.Graphics.FillRectangle(b, e.Bounds)
        End Using
        TextRenderer.DrawText(e.Graphics, GetItemText(Items(e.Index)), Font, New Rectangle(e.Bounds.X + 6, e.Bounds.Y, e.Bounds.Width - 6, e.Bounds.Height),
                              If(Enabled, p.Ink, p.Disabled), TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.EndEllipsis)
    End Sub
End Class

#End Region

#Region "Colour and font pickers"

''' <summary>Button showing a colour swatch; opens the colour dialog. Replaces ColorPickEdit.</summary>
<DefaultEvent("ColorChanged")>
Public Class ColorButton
    Inherits ModernButton

    Private _colour As Color = Color.White
    Public Event ColorChanged As EventHandler

    Public Sub New()
        IconGlyph = ""
        Width = 120
        AutoSize = False
    End Sub

    Public Property SelectedColor As Color
        Get
            Return _colour
        End Get
        Set(value As Color)
            If _colour = value Then Return
            _colour = value
            Invalidate()
            RaiseEvent ColorChanged(Me, EventArgs.Empty)
        End Set
    End Property

    Protected Overrides Sub OnClick(e As EventArgs)
        Using dlg As New ColorDialog With {.Color = _colour, .FullOpen = True, .AnyColor = True}
            If dlg.ShowDialog(FindForm()) = DialogResult.OK Then SelectedColor = dlg.Color
        End Using
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim p = ThemeManager.Current
        Dim s = Height - LogicalToDeviceUnits(14)
        Dim r = New Rectangle(Width - s - LogicalToDeviceUnits(8), LogicalToDeviceUnits(7), s, s)
        Using b As New SolidBrush(_colour), pen As New Pen(p.Border)
            e.Graphics.FillRectangle(b, r)
            e.Graphics.DrawRectangle(pen, r)
        End Using
    End Sub
End Class

''' <summary>Button showing a font name; opens the font dialog. Replaces FontEdit.</summary>
<DefaultEvent("SelectedFontChanged")>
Public Class FontButton
    Inherits ModernButton

    Private _fontName As String = "Arial"
    Private _fontSize As Single = 12
    Public Event SelectedFontChanged As EventHandler

    Public Sub New()
        IconGlyph = Glyphs.Font
        AutoSize = False
        Width = 180
        UpdateCaption()
    End Sub

    Public Property SelectedFontName As String
        Get
            Return _fontName
        End Get
        Set(value As String)
            _fontName = If(String.IsNullOrWhiteSpace(value), "Arial", value)
            UpdateCaption()
        End Set
    End Property

    Public Property SelectedFontSize As Single
        Get
            Return _fontSize
        End Get
        Set(value As Single)
            _fontSize = If(value <= 0, 12, value)
            UpdateCaption()
        End Set
    End Property

    Private _showSize As Boolean

    Public Property ShowSize As Boolean
        Get
            Return _showSize
        End Get
        Set(value As Boolean)
            _showSize = value
            UpdateCaption()
        End Set
    End Property

    Private Sub UpdateCaption()
        Text = If(_showSize, $"{_fontName}, {_fontSize:0.#}", _fontName)
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        Using dlg As New FontDialog With {.ShowEffects = False}
            Try
                dlg.Font = New Font(_fontName, _fontSize)
            Catch
            End Try
            If dlg.ShowDialog(FindForm()) = DialogResult.OK Then
                _fontName = dlg.Font.Name
                _fontSize = dlg.Font.Size
                UpdateCaption()
                RaiseEvent SelectedFontChanged(Me, EventArgs.Empty)
            End If
        End Using
        MyBase.OnClick(e)
    End Sub
End Class

#End Region
