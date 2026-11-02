Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   ColoursScreen.vb - the ReadEZ colours and their calibration (replaces RZfrmColourCalibration)
'
'   The screen colour that matches a real ReadEZ overlay or lens depends on the screen, so every
'   colour should be checked on the patient's screen before testing. This screen shows the three
'   sets (overlays, lenses, dark lenses), which colours have been checked on the screen now used
'   for the patient and when, and opens Eye Dynamix's colour check (ReadEZColourCheck) - compare
'   the colour with the overlay or lens, and if it doesn't match, find the match in a few choices.
' ====================================================================================================

Public Class ColoursScreen
    Inherits ScreenBase

    Private ReadOnly _sets As New List(Of ColourSetView)
    Private ReadOnly lblScreen As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(900, 0)}
    Private ReadOnly lblSummary As New Label With {.AutoSize = True, .Font = Fonts.Strong, .MaximumSize = New Size(900, 0)}

    Public Sub New()
        TitleKey = "colours.title"
        SubtitleKey = "colours.subtitle"
        HelpTopic = "screen.colours"
        Dim check = AddHeaderButton("colours.check", ButtonKind.Primary, Glyphs.Calibrate)
        AddHandler check.Click, Sub() RunCheck("")

        Dim stack As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown, .WrapContents = False, .AutoScroll = True, .BackColor = Color.Transparent}
        Dim intro As New Card With {.Width = 980, .Height = 170, .Padding = New Padding(20, 48, 20, 12), .Margin = New Padding(0, 0, 0, 12)}
        intro.SetTitle("colours.about.title")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Fill
        Dim about = Ui.Label("colours.about", muted:=True)
        about.MaximumSize = New Size(920, 0)
        col.Controls.Add(about)
        LabelStyle.SetMuted(lblScreen)
        lblScreen.Margin = New Padding(0, 8, 0, 2)
        col.Controls.Add(lblScreen)
        col.Controls.Add(lblSummary)
        intro.Controls.Add(col)
        stack.Controls.Add(intro)

        For Each s In ReadEZColours.Sets
            Dim card As New Card With {.Width = 980, .Height = 200, .Padding = New Padding(20, 48, 20, 12), .Margin = New Padding(0, 0, 0, 12)}
            card.SetTitle(If(s = ReadEZSet.Overlay, "readez.overlays", If(s = ReadEZSet.Lens, "readez.lightTint", "readez.darkTint")))
            Dim v As New ColourSetView With {.SetOf = s, .Dock = DockStyle.Fill}
            AddHandler v.ColourClicked, Sub(code) RunCheck(code)
            _sets.Add(v)
            card.Controls.Add(v)
            stack.Controls.Add(card)
        Next
        Content.Controls.Add(stack)
        AddHandler ReadEZColours.Changed, Sub() UpdateStatus()
        Loc.OnChange(Me, AddressOf UpdateStatus)
    End Sub

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        ReadEZColours.ForgetScreen()
        UpdateStatus()
    End Sub

    Private Sub RunCheck(code As String)
        ReadEZColourCheck.Run(FindForm(), code)
        Stimulus.ShowLogo()
        UpdateStatus()
    End Sub

    Private Sub UpdateStatus()
        Dim n = ReadEZColours.CheckedCount(), total = 3 * ReadEZColours.PerSet
        lblScreen.Text = Loc.F("readez.check.screen", ReadEZColours.ScreenName)
        lblSummary.Text = If(n = total, Loc.T("readez.allChecked"), Loc.F("readez.someChecked", n, total))
        lblSummary.ForeColor = If(n = total, ThemeManager.Current.Success, ThemeManager.Current.Warning)
        For Each v In _sets
            v.Invalidate()
        Next
    End Sub
End Class

''' <summary>One ReadEZ set: grey and A - L as large swatches, each with a tick (and date) when checked on this screen.</summary>
Public NotInheritable Class ColourSetView
    Inherits Control
    Implements IThemed

    Public Property SetOf As ReadEZSet
    Public Event ColourClicked(code As String)
    Private _hover As Integer = -1

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Cursor = Cursors.Hand
    End Sub

    Private Function Pitch() As Single
        Return Width / CSng(ReadEZColours.PerSet)
    End Function

    ''' <summary>Order shown: A - L, then grey.</summary>
    Private Shared Function IndexAt(slot As Integer) As Integer
        Return If(slot < 12, slot + 1, 0)
    End Function

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim h = Math.Min(ReadEZColours.PerSet - 1, CInt(Math.Floor(e.X / Pitch())))
        If h <> _hover Then _hover = h : Invalidate()
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        _hover = -1
        Invalidate()
    End Sub

    Protected Overrides Sub OnMouseClick(e As MouseEventArgs)
        MyBase.OnMouseClick(e)
        Dim slot = Math.Min(ReadEZColours.PerSet - 1, CInt(Math.Floor(e.X / Pitch())))
        RaiseEvent ColourClicked(ReadEZColours.Code(SetOf, IndexAt(slot)))
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim w = Pitch()
        Dim d = Math.Min(w - 12, Height - 56)
        For slot = 0 To ReadEZColours.PerSet - 1
            Dim i = IndexAt(slot)
            Dim x = slot * w + (w - d) / 2
            Dim r As New RectangleF(x, 4, d, d)
            Using b As New SolidBrush(ReadEZColours.Colour(SetOf, i)), pen As New Pen(If(slot = _hover, p.Primary, p.Border), If(slot = _hover, 2.5F, 1))
                g.FillEllipse(b, r)
                g.DrawEllipse(pen, r)
            End Using
            Dim label = ReadEZColours.Letter(i)
            TextRenderer.DrawText(g, label, Fonts.Strong, New Rectangle(CInt(slot * w), CInt(r.Bottom + 2), CInt(w), 20), p.Ink, TextFormatFlags.HorizontalCenter)
            Dim checkedDate = ReadEZColours.CheckedOn(SetOf, i)
            Dim txt As String, colour As Color
            If checkedDate.HasValue Then
                txt = ChrW(&H2713) & " " & checkedDate.Value.ToString("d", Loc.Culture)
                colour = p.Success
            ElseIf ReadEZColours.CheckedElsewhere(SetOf, i) Then
                txt = Loc.T("colours.otherScreen")
                colour = p.Warning
            Else
                txt = Loc.T("readez.notChecked")
                colour = p.Muted
            End If
            TextRenderer.DrawText(g, txt, Fonts.Small, New Rectangle(CInt(slot * w), CInt(r.Bottom + 22), CInt(w), 18), colour, TextFormatFlags.HorizontalCenter Or TextFormatFlags.EndEllipsis)
        Next
    End Sub

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property
End Class
