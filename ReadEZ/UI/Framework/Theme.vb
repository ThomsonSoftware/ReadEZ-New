Option Strict On

Imports System.Drawing.Drawing2D
Imports System.Drawing.Text

' ====================================================================================================
'   Theme.vb
'   Colours, fonts and drawing helpers shared by every window.
'
'   Three modes are available:
'     Light - clean white surfaces with the Thomson corporate blue.
'     Dark  - charcoal surfaces for dim clinical rooms.
'     Blue  - based on the blue used for the panels in the previous version
'             (Windows "ActiveCaption" #99B4D1 and "InactiveCaption" #BFCDDB).
'
'   Every custom control reads ThemeManager.Current when it paints, so switching
'   the mode only needs ThemeManager.SetMode(...) - the manager then re-applies
'   colours to standard WinForms controls in every open form.
' ====================================================================================================

Public Enum ThemeMode
    Light = 0
    Dark = 1
    Blue = 2
End Enum

''' <summary>Role of a container, used when re-colouring standard controls.</summary>
Public Enum SurfaceRole
    Auto = 0
    Canvas
    Surface
    SurfaceAlt
    Panel
    Nav
    Header
    Inset         ' a lighter panel inside a card (InsetPanel)
    Stimulus      ' never touched by the theme (stimulus screens keep their own colours)
End Enum

''' <summary>A complete set of colours for one theme mode.</summary>
Public NotInheritable Class Palette
    Public Property Mode As ThemeMode
    Public Property IsDark As Boolean

    ' Backgrounds
    Public Property Canvas As Color
    Public Property Surface As Color
    Public Property SurfaceAlt As Color
    Public Property Panel As Color

    ''' <summary>A panel set inside a card (InsetPanel): lighter than the card in the Blue theme.</summary>
    Public ReadOnly Property Inset As Color
        Get
            Return If(Mode = ThemeMode.Blue, Draw.Mix(Surface, InputBack, 0.6), SurfaceAlt)
        End Get
    End Property
    Public Property InputBack As Color

    ' Text
    Public Property Ink As Color
    Public Property Muted As Color
    Public Property Disabled As Color

    ' Lines
    Public Property Border As Color
    Public Property GridLine As Color

    ' Accent
    Public Property Primary As Color
    Public Property PrimaryHover As Color
    Public Property PrimaryPressed As Color
    Public Property PrimarySoft As Color
    Public Property OnPrimary As Color

    ' Navigation rail and header
    Public Property NavBack As Color
    Public Property NavFore As Color
    Public Property NavMuted As Color
    Public Property NavSelected As Color
    Public Property HeaderBack As Color
    Public Property HeaderFore As Color

    ' Status
    Public Property Danger As Color
    Public Property DangerSoft As Color
    Public Property Success As Color
    Public Property SuccessSoft As Color
    Public Property Warning As Color
    Public Property WarningSoft As Color

    ' Charts
    Public Property ChartBack As Color
    Public Property ChartFore As Color
    Public Property ChartGrid As Color
    ''' <summary>Right eye series colour (kept red/green convention from the old version).</summary>
    Public Property RightEye As Color
    ''' <summary>Left eye series colour.</summary>
    Public Property LeftEye As Color
    Public Property Stimulus As Color
    Public Property Series As Color()
End Class

''' <summary>Holds the current palette and switches between modes.</summary>
Public NotInheritable Class ThemeManager

    Private Sub New()
    End Sub

    Private Shared _current As Palette = BuildLight()
    Private Shared ReadOnly _roles As New Runtime.CompilerServices.ConditionalWeakTable(Of Control, RoleBox)

    Private NotInheritable Class RoleBox
        Public Role As SurfaceRole
    End Class

    ''' <summary>Raised after the mode changes and every open form has been re-coloured.</summary>
    Public Shared Event ThemeChanged As EventHandler

    Public Shared ReadOnly Property Current As Palette
        Get
            Return _current
        End Get
    End Property

    Public Shared ReadOnly Property Mode As ThemeMode
        Get
            Return _current.Mode
        End Get
    End Property

#Region "Switching"

    Public Shared Sub SetMode(mode As ThemeMode)
        _current = BuildPalette(mode)
        For Each f As Form In Application.OpenForms.Cast(Of Form)().ToList()
            Apply(f)
        Next
        RaiseEvent ThemeChanged(Nothing, EventArgs.Empty)
    End Sub

    Public Shared Function BuildPalette(mode As ThemeMode) As Palette
        Select Case mode
            Case ThemeMode.Dark : Return BuildDark()
            Case ThemeMode.Blue : Return BuildBlue()
            Case Else : Return BuildLight()
        End Select
    End Function

#End Region

#Region "Palettes"

    Private Shared Function BuildLight() As Palette
        Return New Palette With {
            .Mode = ThemeMode.Light, .IsDark = False,
            .Canvas = Color.FromArgb(238, 242, 246),
            .Surface = Color.White,
            .SurfaceAlt = Color.FromArgb(246, 248, 250),
            .Panel = Color.FromArgb(246, 248, 250),
            .InputBack = Color.White,
            .Ink = Color.FromArgb(27, 39, 51),
            .Muted = Color.FromArgb(91, 107, 123),
            .Disabled = Color.FromArgb(160, 170, 180),
            .Border = Color.FromArgb(213, 221, 229),
            .GridLine = Color.FromArgb(232, 237, 242),
            .Primary = Color.FromArgb(9, 85, 149),
            .PrimaryHover = Color.FromArgb(14, 103, 178),
            .PrimaryPressed = Color.FromArgb(7, 70, 124),
            .PrimarySoft = Color.FromArgb(220, 233, 246),
            .OnPrimary = Color.White,
            .NavBack = Color.FromArgb(12, 46, 82),
            .NavFore = Color.White,
            .NavMuted = Color.FromArgb(165, 189, 214),
            .NavSelected = Color.FromArgb(9, 85, 149),
            .HeaderBack = Color.White,
            .HeaderFore = Color.FromArgb(27, 39, 51),
            .Danger = Color.FromArgb(180, 35, 24),
            .DangerSoft = Color.FromArgb(253, 236, 234),
            .Success = Color.FromArgb(22, 128, 61),
            .SuccessSoft = Color.FromArgb(225, 245, 232),
            .Warning = Color.FromArgb(176, 106, 0),
            .WarningSoft = Color.FromArgb(255, 243, 220),
            .ChartBack = Color.White,
            .ChartFore = Color.FromArgb(27, 39, 51),
            .ChartGrid = Color.FromArgb(226, 232, 238),
            .RightEye = Color.FromArgb(214, 39, 40),
            .LeftEye = Color.FromArgb(31, 119, 180),
            .Stimulus = Color.FromArgb(44, 160, 44),
            .Series = {Color.FromArgb(214, 39, 40), Color.FromArgb(31, 119, 180), Color.FromArgb(44, 160, 44),
                       Color.FromArgb(255, 127, 14), Color.FromArgb(148, 103, 189), Color.FromArgb(140, 86, 75)}
        }
    End Function

    Private Shared Function BuildDark() As Palette
        Return New Palette With {
            .Mode = ThemeMode.Dark, .IsDark = True,
            .Canvas = Color.FromArgb(23, 25, 28),
            .Surface = Color.FromArgb(33, 36, 40),
            .SurfaceAlt = Color.FromArgb(42, 46, 51),
            .Panel = Color.FromArgb(38, 42, 47),
            .InputBack = Color.FromArgb(28, 31, 35),
            .Ink = Color.FromArgb(232, 234, 237),
            .Muted = Color.FromArgb(154, 163, 173),
            .Disabled = Color.FromArgb(95, 102, 110),
            .Border = Color.FromArgb(58, 63, 70),
            .GridLine = Color.FromArgb(48, 52, 58),
            .Primary = Color.FromArgb(61, 139, 218),
            .PrimaryHover = Color.FromArgb(86, 158, 230),
            .PrimaryPressed = Color.FromArgb(44, 114, 186),
            .PrimarySoft = Color.FromArgb(30, 58, 86),
            .OnPrimary = Color.White,
            .NavBack = Color.FromArgb(17, 19, 22),
            .NavFore = Color.FromArgb(232, 234, 237),
            .NavMuted = Color.FromArgb(130, 140, 150),
            .NavSelected = Color.FromArgb(40, 70, 104),
            .HeaderBack = Color.FromArgb(28, 31, 35),
            .HeaderFore = Color.FromArgb(232, 234, 237),
            .Danger = Color.FromArgb(239, 97, 85),
            .DangerSoft = Color.FromArgb(70, 32, 30),
            .Success = Color.FromArgb(87, 200, 120),
            .SuccessSoft = Color.FromArgb(28, 60, 40),
            .Warning = Color.FromArgb(240, 180, 70),
            .WarningSoft = Color.FromArgb(70, 55, 25),
            .ChartBack = Color.FromArgb(33, 36, 40),
            .ChartFore = Color.FromArgb(220, 224, 228),
            .ChartGrid = Color.FromArgb(55, 60, 66),
            .RightEye = Color.FromArgb(255, 99, 99),
            .LeftEye = Color.FromArgb(92, 170, 255),
            .Stimulus = Color.FromArgb(110, 220, 120),
            .Series = {Color.FromArgb(255, 99, 99), Color.FromArgb(92, 170, 255), Color.FromArgb(110, 220, 120),
                       Color.FromArgb(255, 170, 70), Color.FromArgb(190, 150, 240), Color.FromArgb(200, 150, 120)}
        }
    End Function

    ''' <summary>
    ''' Based on the panel blue of the previous version: Windows ActiveCaption (153,180,209)
    ''' for panels and InactiveCaption (191,205,219) for the background.
    ''' </summary>
    Private Shared Function BuildBlue() As Palette
        Return New Palette With {
            .Mode = ThemeMode.Blue, .IsDark = False,
            .Canvas = Color.FromArgb(191, 205, 219),
            .Surface = Color.FromArgb(251, 252, 254),          ' near-white panels, so they stand out on the blue
            .SurfaceAlt = Color.FromArgb(236, 242, 248),
            .Panel = Color.FromArgb(153, 180, 209),
            .InputBack = Color.FromArgb(244, 248, 252),
            .Ink = Color.FromArgb(16, 33, 54),
            .Muted = Color.FromArgb(54, 78, 105),
            .Disabled = Color.FromArgb(120, 138, 158),
            .Border = Color.FromArgb(150, 174, 202),
            .GridLine = Color.FromArgb(185, 202, 220),
            .Primary = Color.FromArgb(31, 78, 121),
            .PrimaryHover = Color.FromArgb(42, 97, 148),
            .PrimaryPressed = Color.FromArgb(24, 62, 98),
            .PrimarySoft = Color.FromArgb(176, 199, 224),
            .OnPrimary = Color.White,
            .NavBack = Color.FromArgb(31, 78, 121),
            .NavFore = Color.White,
            .NavMuted = Color.FromArgb(190, 210, 230),
            .NavSelected = Color.FromArgb(91, 128, 168),
            .HeaderBack = Color.FromArgb(153, 180, 209),
            .HeaderFore = Color.FromArgb(16, 33, 54),
            .Danger = Color.FromArgb(170, 30, 20),
            .DangerSoft = Color.FromArgb(246, 222, 220),
            .Success = Color.FromArgb(20, 115, 55),
            .SuccessSoft = Color.FromArgb(214, 238, 222),
            .Warning = Color.FromArgb(160, 95, 0),
            .WarningSoft = Color.FromArgb(246, 233, 205),
            .ChartBack = Color.White,
            .ChartFore = Color.FromArgb(16, 33, 54),
            .ChartGrid = Color.FromArgb(196, 210, 226),
            .RightEye = Color.FromArgb(200, 30, 30),
            .LeftEye = Color.FromArgb(20, 90, 170),
            .Stimulus = Color.FromArgb(30, 140, 50),
            .Series = {Color.FromArgb(200, 30, 30), Color.FromArgb(20, 90, 170), Color.FromArgb(30, 140, 50),
                       Color.FromArgb(220, 110, 10), Color.FromArgb(120, 80, 170), Color.FromArgb(120, 70, 60)}
        }
    End Function

#End Region

#Region "Roles and applying colours to standard controls"

    ''' <summary>Tags a container so Apply() gives it (and its plain children) the right background.</summary>
    Public Shared Sub SetRole(c As Control, role As SurfaceRole)
        _roles.Remove(c)
        _roles.Add(c, New RoleBox With {.Role = role})
        If role <> SurfaceRole.Stimulus Then
            Dim back = BackFor(role)
            If Not back.IsEmpty Then c.BackColor = back
        End If
    End Sub

    Public Shared Function GetRole(c As Control) As SurfaceRole
        Dim box As RoleBox = Nothing
        If _roles.TryGetValue(c, box) Then Return box.Role
        Return SurfaceRole.Auto
    End Function

    Public Shared Function BackFor(role As SurfaceRole) As Color
        Dim p = _current
        Select Case role
            Case SurfaceRole.Canvas : Return p.Canvas
            Case SurfaceRole.Surface : Return p.Surface
            Case SurfaceRole.SurfaceAlt : Return p.SurfaceAlt
            Case SurfaceRole.Panel : Return p.Panel
            Case SurfaceRole.Nav : Return p.NavBack
            Case SurfaceRole.Header : Return p.HeaderBack
            Case SurfaceRole.Inset : Return p.Inset
            Case Else : Return Color.Empty
        End Select
    End Function

    ''' <summary>
    ''' Re-colours a control tree. Custom controls (IThemed) repaint themselves; standard
    ''' WinForms controls get colours based on the nearest role-tagged container.
    ''' </summary>
    Public Shared Sub Apply(root As Control)
        If root Is Nothing Then Return
        root.SuspendLayout()
        Try
            ApplyOne(root, If(TypeOf root Is Form, SurfaceRole.Canvas, SurfaceRole.Surface))
        Finally
            root.ResumeLayout(True)
        End Try
        root.Invalidate(True)
    End Sub

    Private Shared Sub ApplyOne(c As Control, inherited As SurfaceRole)
        Dim role = GetRole(c)
        If role = SurfaceRole.Stimulus Then Return
        If role = SurfaceRole.Auto Then role = inherited
        Dim p = _current
        Dim back = BackFor(role)
        Dim isNav = role = SurfaceRole.Nav
        Dim fore = If(isNav, p.NavFore, If(role = SurfaceRole.Header, p.HeaderFore, p.Ink))

        If TypeOf c Is IThemed Then
            DirectCast(c, IThemed).ApplyTheme(p)
        ElseIf TypeOf c Is Form OrElse TypeOf c Is Panel OrElse TypeOf c Is TableLayoutPanel OrElse
               TypeOf c Is FlowLayoutPanel OrElse TypeOf c Is SplitContainer OrElse TypeOf c Is SplitterPanel OrElse
               TypeOf c Is UserControl OrElse TypeOf c Is GroupBox OrElse TypeOf c Is TabPage Then
            If Not back.IsEmpty Then c.BackColor = back
            c.ForeColor = fore
        ElseIf TypeOf c Is Label Then
            Dim l = DirectCast(c, Label)
            If Not LabelStyle.IsCustom(l) Then
                If l.BackColor <> Color.Transparent AndAlso Not back.IsEmpty Then l.BackColor = back
                l.ForeColor = If(LabelStyle.IsMuted(l), If(isNav, p.NavMuted, p.Muted), fore)
            End If
        ElseIf TypeOf c Is TextBox OrElse TypeOf c Is RichTextBox OrElse TypeOf c Is NumericUpDown OrElse
               TypeOf c Is ListBox OrElse TypeOf c Is ComboBox OrElse TypeOf c Is DateTimePicker OrElse
               TypeOf c Is MaskedTextBox Then
            c.BackColor = p.InputBack
            c.ForeColor = p.Ink
        ElseIf TypeOf c Is CheckBox OrElse TypeOf c Is RadioButton Then
            c.BackColor = If(back.IsEmpty, c.BackColor, back)
            c.ForeColor = fore
        ElseIf TypeOf c Is DataGridView Then
            ColourGrid(DirectCast(c, DataGridView))
        ElseIf TypeOf c Is TreeView Then
            Dim t = DirectCast(c, TreeView)
            t.BackColor = p.Surface
            t.ForeColor = p.Ink
            t.LineColor = p.Border
        ElseIf TypeOf c Is ListView Then
            c.BackColor = p.Surface
            c.ForeColor = p.Ink
        ElseIf TypeOf c Is ToolStrip Then
            Dim ts = DirectCast(c, ToolStrip)
            ts.Renderer = New ThemedToolStripRenderer()
            ts.BackColor = p.Surface
            ts.ForeColor = p.Ink
        ElseIf TypeOf c Is TabControl Then
            c.BackColor = back
        End If

        If TypeOf c Is IThemed AndAlso Not DirectCast(c, IThemed).ThemeChildren Then Return
        For Each child As Control In c.Controls
            ApplyOne(child, role)
        Next
        If c.ContextMenuStrip IsNot Nothing Then
            c.ContextMenuStrip.Renderer = New ThemedToolStripRenderer()
        End If
    End Sub

#End Region

#Region "Grid styling"

    ''' <summary>One-time set-up of a grid's behaviour and look (call once when the grid is created).</summary>
    Public Shared Sub StyleGrid(grid As DataGridView)
        With grid
            .BorderStyle = BorderStyle.None
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            .ColumnHeadersHeight = grid.Font.Height + 18
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .EnableHeadersVisualStyles = False
            .ColumnHeadersDefaultCellStyle.Font = Fonts.Small
            .ColumnHeadersDefaultCellStyle.Padding = New Padding(8, 0, 0, 0)
            .DefaultCellStyle.Font = Fonts.Body
            .DefaultCellStyle.Padding = New Padding(8, 0, 4, 0)
            .RowTemplate.Height = grid.Font.Height + 16
        End With
        ColourGrid(grid)
    End Sub

    ''' <summary>Applies the palette's colours to a grid (called on every theme change).</summary>
    Public Shared Sub ColourGrid(grid As DataGridView)
        Dim p = _current
        With grid
            .BackgroundColor = p.Surface
            .GridColor = p.GridLine
            .ColumnHeadersDefaultCellStyle.BackColor = p.Surface
            .ColumnHeadersDefaultCellStyle.ForeColor = p.Muted
            .ColumnHeadersDefaultCellStyle.SelectionBackColor = p.Surface
            .ColumnHeadersDefaultCellStyle.SelectionForeColor = p.Muted
            .DefaultCellStyle.ForeColor = p.Ink
            .DefaultCellStyle.BackColor = p.Surface
            .DefaultCellStyle.SelectionBackColor = p.PrimarySoft
            .DefaultCellStyle.SelectionForeColor = p.Ink
            .AlternatingRowsDefaultCellStyle.BackColor = p.Surface
        End With
    End Sub

#End Region

End Class

''' <summary>Implemented by custom controls that paint themselves from the palette.</summary>
Public Interface IThemed
    Sub ApplyTheme(p As Palette)
    ''' <summary>False if the control manages the colours of its own children.</summary>
    ReadOnly Property ThemeChildren As Boolean
End Interface

''' <summary>Marks labels as muted (secondary text) or as having a fixed custom colour.</summary>
Public NotInheritable Class LabelStyle
    Private Sub New()
    End Sub

    Private Shared ReadOnly _muted As New Runtime.CompilerServices.ConditionalWeakTable(Of Label, Object)
    Private Shared ReadOnly _custom As New Runtime.CompilerServices.ConditionalWeakTable(Of Label, Object)

    Public Shared Sub SetMuted(l As Label)
        _muted.Remove(l)
        _muted.Add(l, True)
        l.ForeColor = ThemeManager.Current.Muted
    End Sub

    Public Shared Function IsMuted(l As Label) As Boolean
        Dim o As Object = Nothing
        Return _muted.TryGetValue(l, o)
    End Function

    ''' <summary>The theme will not change this label's text colour.</summary>
    Public Shared Sub SetCustom(l As Label)
        _custom.Remove(l)
        _custom.Add(l, True)
    End Sub

    Public Shared Function IsCustom(l As Label) As Boolean
        Dim o As Object = Nothing
        Return _custom.TryGetValue(l, o)
    End Function
End Class

''' <summary>Fonts used throughout the interface.</summary>
Public NotInheritable Class Fonts
    Private Sub New()
    End Sub

    Public Shared ReadOnly Body As New Font("Segoe UI", 9.75F, FontStyle.Regular)
    Public Shared ReadOnly Strong As New Font("Segoe UI Semibold", 9.75F, FontStyle.Regular)
    Public Shared ReadOnly Small As New Font("Segoe UI", 8.75F, FontStyle.Regular)
    Public Shared ReadOnly SmallStrong As New Font("Segoe UI Semibold", 8.75F, FontStyle.Regular)
    Public Shared ReadOnly Section As New Font("Segoe UI Semibold", 11.5F, FontStyle.Regular)
    Public Shared ReadOnly Title As New Font("Segoe UI Semibold", 16.0F, FontStyle.Regular)
    Public Shared ReadOnly Display As New Font("Segoe UI Light", 26.0F, FontStyle.Regular)
    Public Shared ReadOnly Icon As Font = CreateIconFont(10.5F)
    Public Shared ReadOnly IconLarge As Font = CreateIconFont(16.0F)

    Private Shared Function CreateIconFont(size As Single) As Font
        Using fonts As New InstalledFontCollection()
            For Each name In {"Segoe Fluent Icons", "Segoe MDL2 Assets"}
                If fonts.Families.Any(Function(f) f.Name = name) Then Return New Font(name, size, FontStyle.Regular)
            Next
        End Using
        Return Nothing
    End Function
End Class

''' <summary>Segoe MDL2 / Fluent icon glyphs (built into Windows 10 and 11).</summary>
Public NotInheritable Class Glyphs
    Private Sub New()
    End Sub

    Public Const Add As String = ChrW(&HE710)
    Public Const Edit As String = ChrW(&HE70F)
    Public Const Play As String = ChrW(&HE768)
    Public Const Pause As String = ChrW(&HE769)
    Public Const [Stop] As String = ChrW(&HE71A)
    Public Const Record As String = ChrW(&HE7C8)
    Public Const Rewind As String = ChrW(&HEB9E)
    Public Const Delete As String = ChrW(&HE74D)
    Public Const Print As String = ChrW(&HE749)
    Public Const Export As String = ChrW(&HE896)
    Public Const Settings As String = ChrW(&HE713)
    Public Const Help As String = ChrW(&HE897)
    Public Const More As String = ChrW(&HE712)
    Public Const Search As String = ChrW(&HE721)
    Public Const Back As String = ChrW(&HE72B)
    Public Const Forward As String = ChrW(&HE72A)
    Public Const Accept As String = ChrW(&HE73E)
    Public Const Cancel As String = ChrW(&HE711)
    Public Const Save As String = ChrW(&HE74E)
    Public Const Open As String = ChrW(&HE8E5)
    Public Const Copy As String = ChrW(&HE8C8)
    Public Const ZoomIn As String = ChrW(&HE8A3)
    Public Const ZoomOut As String = ChrW(&HE71F)
    Public Const People As String = ChrW(&HE716)
    Public Const Person As String = ChrW(&HE77B)
    Public Const Target As String = ChrW(&HE81D)
    Public Const Eye As String = ChrW(&HE7B3)
    Public Const Chart As String = ChrW(&HE9D2)
    Public Const List As String = ChrW(&HEA37)
    Public Const Home As String = ChrW(&HE80F)
    Public Const Admin As String = ChrW(&HE7EF)
    Public Const Lock As String = ChrW(&HE72E)
    Public Const Globe As String = ChrW(&HE774)
    Public Const Brightness As String = ChrW(&HE706)
    Public Const Report As String = ChrW(&HE9F9)
    Public Const Video As String = ChrW(&HE714)
    Public Const Picture As String = ChrW(&HE8B9)
    Public Const Font As String = ChrW(&HE8D2)
    Public Const Color As String = ChrW(&HE790)
    Public Const Refresh As String = ChrW(&HE72C)
    Public Const Info As String = ChrW(&HE946)
    Public Const Warning As String = ChrW(&HE7BA)
    Public Const Tracker As String = ChrW(&HE7F4)
    Public Const Calibrate As String = ChrW(&HE7B5)
    Public Const Questionnaire As String = ChrW(&HE9D5)
    Public Const Favourite As String = ChrW(&HE734)
    Public Const ChevronDown As String = ChrW(&HE70D)
    Public Const ChevronUp As String = ChrW(&HE70E)
    Public Const ChevronRight As String = ChrW(&HE76C)
    Public Const Folder As String = ChrW(&HE8B7)
    Public Const Restore As String = ChrW(&HE777)
    Public Const Move As String = ChrW(&HE7C2)
End Class

''' <summary>Drawing helpers.</summary>
Public NotInheritable Class Draw
    Private Sub New()
    End Sub

    Public Shared Function RoundedRect(r As RectangleF, radius As Single) As GraphicsPath
        Dim path As New GraphicsPath()
        ' A control can be 0 or negative in size while a layout shrinks; AddArc throws for that,
        ' which escaped OnPaint as an error dialog and a red X.
        If r.Width <= 0 OrElse r.Height <= 0 Then Return path
        Dim d = Math.Min(radius * 2, Math.Min(r.Width, r.Height))
        If radius <= 0 OrElse d <= 0 Then
            path.AddRectangle(r)
            Return path
        End If
        path.AddArc(r.X, r.Y, d, d, 180, 90)
        path.AddArc(r.Right - d, r.Y, d, d, 270, 90)
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90)
        path.AddArc(r.X, r.Bottom - d, d, d, 90, 90)
        path.CloseFigure()
        Return path
    End Function

    ''' <summary>Blend two colours: amount 0 = a, 1 = b.</summary>
    Public Shared Function Mix(a As Color, b As Color, amount As Double) As Color
        amount = Math.Max(0, Math.Min(1, amount))
        Return Color.FromArgb(
            CInt(a.A + (CInt(b.A) - a.A) * amount),
            CInt(a.R + (CInt(b.R) - a.R) * amount),
            CInt(a.G + (CInt(b.G) - a.G) * amount),
            CInt(a.B + (CInt(b.B) - a.B) * amount))
    End Function

    ''' <summary>Background of the nearest parent that has a solid colour.</summary>
    Public Shared Function ParentBack(c As Control) As Color
        Dim p = c.Parent
        While p IsNot Nothing
            If p.BackColor.A = 255 Then Return p.BackColor
            p = p.Parent
        End While
        Return ThemeManager.Current.Surface
    End Function
End Class

''' <summary>Menu and toolbar renderer that follows the palette.</summary>
Public Class ThemedToolStripRenderer
    Inherits ToolStripProfessionalRenderer

    Public Sub New()
        MyBase.New(New ThemedColorTable())
        RoundedEdges = False
    End Sub

    Protected Overrides Sub OnRenderItemText(e As ToolStripItemTextRenderEventArgs)
        Dim p = ThemeManager.Current
        ' Help topics without a video yet are muted (they still open their description)
        Dim topic = TryCast(e.Item.Tag, HelpTopic)
        e.TextColor = If(Not e.Item.Enabled, p.Disabled, If(topic IsNot Nothing AndAlso Not topic.HasVideo, p.Muted, p.Ink))
        MyBase.OnRenderItemText(e)
    End Sub

    Protected Overrides Sub OnRenderArrow(e As ToolStripArrowRenderEventArgs)
        e.ArrowColor = ThemeManager.Current.Muted
        MyBase.OnRenderArrow(e)
    End Sub

    Private NotInheritable Class ThemedColorTable
        Inherits ProfessionalColorTable

        Private ReadOnly Property P As Palette
            Get
                Return ThemeManager.Current
            End Get
        End Property

        Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property MenuBorder As Color
            Get
                Return P.Border
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemBorder As Color
            Get
                Return P.PrimarySoft
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemSelected As Color
            Get
                Return P.PrimarySoft
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemSelectedGradientBegin As Color
            Get
                Return P.PrimarySoft
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemSelectedGradientEnd As Color
            Get
                Return P.PrimarySoft
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemPressedGradientBegin As Color
            Get
                Return P.PrimarySoft
            End Get
        End Property
        Public Overrides ReadOnly Property MenuItemPressedGradientEnd As Color
            Get
                Return P.PrimarySoft
            End Get
        End Property
        Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property SeparatorDark As Color
            Get
                Return P.Border
            End Get
        End Property
        Public Overrides ReadOnly Property SeparatorLight As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property ToolStripGradientBegin As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property ToolStripGradientMiddle As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property ToolStripGradientEnd As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property MenuStripGradientBegin As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property MenuStripGradientEnd As Color
            Get
                Return P.Surface
            End Get
        End Property
        Public Overrides ReadOnly Property ToolStripBorder As Color
            Get
                Return P.Surface
            End Get
        End Property
    End Class
End Class
