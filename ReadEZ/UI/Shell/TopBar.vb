Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   TopBar.vb
'   The bar across the top of the main window:
'     left   - the current patient (click to go to Patients)
'     right  - screener, licence, ReadEZ colour check and screen mode (one / two screens),
'              theme picker (language: Settings). Each chip opens what it describes.
' ====================================================================================================

Public Enum ChipState
    Neutral
    Good
    Warning
    Bad
    Busy
End Enum

''' <summary>Small rounded status indicator with a coloured dot.</summary>
Public Class StatusChip
    Inherits Control
    Implements IThemed

    Private _state As ChipState = ChipState.Neutral

    Public Sub New()
        SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw, True)
        Font = Fonts.SmallStrong
        AutoSize = True
        Margin = New Padding(6, 0, 0, 0)
    End Sub

    Public Overrides Function GetPreferredSize(proposedSize As Size) As Size
        Return New Size(TextRenderer.MeasureText(Text, Font).Width + LogicalToDeviceUnits(36), LogicalToDeviceUnits(30))
    End Function

    Public Property State As ChipState
        Get
            Return _state
        End Get
        Set(value As ChipState)
            _state = value
            Invalidate()
        End Set
    End Property

    Protected Overrides Sub OnTextChanged(e As EventArgs)
        MyBase.OnTextChanged(e)
        If AutoSize Then Size = GetPreferredSize(Size.Empty)
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

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(Draw.ParentBack(Me))
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim dot As Color, back As Color
        Select Case _state
            Case ChipState.Good : dot = p.Success : back = p.SuccessSoft
            Case ChipState.Warning : dot = p.Warning : back = p.WarningSoft
            Case ChipState.Bad : dot = p.Danger : back = p.DangerSoft
            Case ChipState.Busy : dot = p.Primary : back = p.PrimarySoft
            Case Else : dot = p.Muted : back = p.SurfaceAlt
        End Select
        Using path = Draw.RoundedRect(New RectangleF(0.5F, 0.5F, Width - 1.5F, Height - 1.5F), (Height - 1) / 2.0F), b As New SolidBrush(back)
            g.FillPath(b, path)
        End Using
        Dim d = LogicalToDeviceUnits(8)
        Using b As New SolidBrush(dot)
            g.FillEllipse(b, LogicalToDeviceUnits(12), (Height - d) \ 2, d, d)
        End Using
        TextRenderer.DrawText(g, Text, Font, New Rectangle(LogicalToDeviceUnits(26), 0, Width - LogicalToDeviceUnits(30), Height), p.Ink,
                              TextFormatFlags.VerticalCenter Or TextFormatFlags.Left)
    End Sub
End Class

Public Class TopBar
    Inherits Panel
    Implements IThemed

    Private ReadOnly _patientButton As New Control
    Private ReadOnly _licence As New StatusChip
    Private ReadOnly _screener As New StatusChip
    Private ReadOnly _colours As New StatusChip
    Private ReadOnly _screens As New StatusChip
    Private ReadOnly _theme As New ModernButton
    Private ReadOnly _themeMenu As New ContextMenuStrip
    Private ReadOnly _tip As New ToolTip
    Private _patientHover As Boolean

    Public Event PatientClicked As EventHandler
    Public Event ScreenerClicked As EventHandler
    Public Event ColoursClicked As EventHandler
    Public Event ScreensClicked As EventHandler
    Public Event LicenceClicked As EventHandler

    Public Sub New()
        SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.ResizeRedraw, True)
        Dock = DockStyle.Top
        Height = 60
        Padding = New Padding(20, 0, 16, 0)
        ThemeManager.SetRole(Me, SurfaceRole.Header)

        ' ----- Right-hand group -----
        Dim right As New FlowLayoutPanel With {
            .Dock = DockStyle.Right, .AutoSize = True, .WrapContents = False,
            .FlowDirection = FlowDirection.LeftToRight, .Padding = New Padding(0, 15, 0, 0), .BackColor = Color.Transparent}

        _licence.Cursor = Cursors.Hand
        AddHandler _licence.Click, Sub() RaiseEvent LicenceClicked(Me, EventArgs.Empty)
        For Each c In {_screener, _colours, _screens}
            c.Cursor = Cursors.Hand
        Next
        AddHandler _screener.Click, Sub() RaiseEvent ScreenerClicked(Me, EventArgs.Empty)
        AddHandler _colours.Click, Sub() RaiseEvent ColoursClicked(Me, EventArgs.Empty)
        AddHandler _screens.Click, Sub() RaiseEvent ScreensClicked(Me, EventArgs.Empty)

        _theme.Kind = ButtonKind.Ghost
        _theme.IconGlyph = Glyphs.Brightness
        _theme.IconOnly = True
        _theme.Size = New Size(34, 30)
        _theme.Margin = New Padding(14, 0, 0, 0)
        Loc.BindTip(_tip, _theme, "top.theme")
        For Each m In {ThemeMode.Light, ThemeMode.Dark, ThemeMode.Blue}
            Dim mode = m
            Dim item As New ToolStripMenuItem
            Loc.Bind(item, "theme." & mode.ToString().ToLowerInvariant())
            AddHandler item.Click, Sub() ChangeTheme(mode)
            _themeMenu.Items.Add(item)
        Next
        AddHandler _theme.Click, Sub()
                                     _themeMenu.Renderer = New ThemedToolStripRenderer()
                                     _themeMenu.Show(_theme, New Point(0, _theme.Height + 2))
                                 End Sub

        right.Controls.Add(_screener)
        right.Controls.Add(_licence)
        right.Controls.Add(_colours)
        right.Controls.Add(_screens)
        right.Controls.Add(_theme)
        Controls.Add(right)

        ' ----- Patient (left) -----
        _patientButton.Dock = DockStyle.Fill
        _patientButton.Cursor = Cursors.Hand
        AddHandler _patientButton.Paint, AddressOf PaintPatient
        AddHandler _patientButton.MouseEnter, Sub()
                                                  _patientHover = True
                                                  _patientButton.Invalidate()
                                              End Sub
        AddHandler _patientButton.MouseLeave, Sub()
                                                  _patientHover = False
                                                  _patientButton.Invalidate()
                                              End Sub
        AddHandler _patientButton.Click, Sub() RaiseEvent PatientClicked(Me, EventArgs.Empty)
        Controls.Add(_patientButton)
        _patientButton.BringToFront()

        AddHandler AppHost.Session.PatientChanged, Sub() _patientButton.Invalidate()
        AddHandler AppHost.Session.ScreenerChanged, Sub() RefreshStatus()
        AddHandler ReadEZColours.Changed, Sub() RefreshStatus()
        AddHandler Loc.LanguageChanged, Sub()
                                            RefreshStatus()
                                            RefreshLicence()
                                            _patientButton.Invalidate()
                                        End Sub
        RefreshStatus()
        RefreshLicence()
    End Sub

#Region "Theme"

    Private Sub ChangeTheme(mode As ThemeMode)
        AppHost.Settings.Theme = mode
        AppHost.SaveSettings()
        ThemeManager.SetMode(mode)
    End Sub

#End Region

#Region "Status chips"

    ''' <summary>
    ''' Licence chip: green while licensed, amber within a month of expiry (as D-Chart's status bar
    ''' message) or in the trial, red when not licensed. The tooltip gives the details.
    ''' </summary>
    Public Sub RefreshLicence()
        Dim d = LicenceService.Data
        Dim days = LicenceService.DaysToExpiry()
        Dim expiry = LicenceService.ExpiryDate()
        Dim tip As New List(Of String)
        If d Is Nothing OrElse LicenceService.NeedsLicenceWindow() Then
            Dim trial = d IsNot Nothing AndAlso d.Demo_Date <> Date.MinValue AndAlso d.Demo_Count > 0 AndAlso
                        LicenceService.TrialDays - CInt(Math.Floor((Date.Today - d.Demo_Date.Date).TotalDays)) > 0
            _licence.State = If(trial, ChipState.Warning, ChipState.Bad)
            _licence.Text = Loc.T(If(trial, "licence.chip.trial", "licence.chip.none"))
        ElseIf days.HasValue AndAlso days.Value <= LicenceService.WarnDaysBeforeExpiry Then
            _licence.State = ChipState.Warning
            _licence.Text = If(days.Value <= 0, Loc.T("licence.chip.expiresToday"), Loc.F("licence.chip.expiresIn", days.Value))
        Else
            _licence.State = ChipState.Good
            _licence.Text = If(expiry.HasValue, Loc.F("licence.chip.validUntil", expiry.Value.ToShortDateString()), Loc.T("licence.chip.licensed"))
        End If
        tip.Add(_licence.Text)
        If d IsNot Nothing AndAlso Not d.LicenceMode AndAlso d.OL_LicencedTo <> "" Then tip.Add(Loc.F("licence.licensedTo", d.OL_LicencedTo))
        If expiry.HasValue Then tip.Add(Loc.F("licence.expiryDate", expiry.Value.ToShortDateString()))
        If LicenceService.MarketingMessage <> "" Then tip.Add(LicenceService.MarketingMessage)
        tip.Add(Loc.T("licence.chip.tip"))
        _tip.SetToolTip(_licence, String.Join(Environment.NewLine, tip))
    End Sub

    ''' <summary>Screener, ReadEZ colours checked on the patient's screen, and one or two screens.</summary>
    Public Sub RefreshStatus()
        Dim name = ""
        Try
            name = AppHost.Session.ScreenerName
        Catch ex As Exception
            Log.Warn("Screener: " & ex.Message)
        End Try
        _screener.State = If(name = "", ChipState.Warning, ChipState.Neutral)
        _screener.Text = If(name = "", Loc.T("top.noScreener"), name)
        _tip.SetToolTip(_screener, Loc.T("top.screener.tip"))

        Dim n = ReadEZColours.CheckedCount(), total = 3 * ReadEZColours.PerSet
        _colours.State = If(n = total, ChipState.Good, If(n = 0, ChipState.Warning, ChipState.Neutral))
        _colours.Text = If(n = total, Loc.T("top.colours.checked"), Loc.F("top.colours.some", n, total))
        _tip.SetToolTip(_colours, Loc.T("top.colours.tip"))

        Dim two = ScreenLayout.IsTwoScreenMode()
        _screens.State = If(two, ChipState.Good, ChipState.Neutral)
        _screens.Text = Loc.T(If(two, "top.screens.two", "top.screens.one"))
        _tip.SetToolTip(_screens, Loc.T("top.screens.tip"))
    End Sub

#End Region

#Region "Painting"

    Public Sub ApplyTheme(p As Palette) Implements IThemed.ApplyTheme
        BackColor = p.HeaderBack
        _patientButton.BackColor = p.HeaderBack
        Invalidate(True)
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return True
        End Get
    End Property

    Private Sub PaintPatient(sender As Object, e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.Clear(p.HeaderBack)
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim s = AppHost.Session
        Dim h = _patientButton.Height
        Dim avatar = LogicalToDeviceUnits(36)
        Dim ay = (h - avatar) \ 2
        Dim accent = If(p.Mode = ThemeMode.Blue, p.Primary, p.Primary)
        Using b As New SolidBrush(If(s.HasPatient, Draw.Mix(p.HeaderBack, accent, 0.18), Draw.Mix(p.HeaderBack, p.Muted, 0.15)))
            g.FillEllipse(b, 0, ay, avatar, avatar)
        End Using
        Dim ini = If(s.HasPatient, Initials(s.Patient), "")
        If ini <> "" Then
            TextRenderer.DrawText(g, ini, Fonts.Strong, New Rectangle(0, ay, avatar, avatar), accent,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        ElseIf Fonts.Icon IsNot Nothing Then
            TextRenderer.DrawText(g, Glyphs.Person, Fonts.Icon, New Rectangle(0, ay, avatar, avatar), p.Muted,
                                  TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter)
        End If
        Dim x = avatar + LogicalToDeviceUnits(12)
        Dim title = If(s.HasPatient, s.PatientName, Loc.T("top.noPatient"))
        Dim detail = If(s.HasPatient, PatientDetail(s.Patient), Loc.T("top.choosePatient"))
        Dim fore = If(_patientHover, p.Primary, p.HeaderFore)
        TextRenderer.DrawText(g, title, Fonts.Strong, New Rectangle(x, h \ 2 - LogicalToDeviceUnits(19), _patientButton.Width - x, LogicalToDeviceUnits(20)),
                              fore, TextFormatFlags.Left Or TextFormatFlags.Bottom Or TextFormatFlags.EndEllipsis)
        TextRenderer.DrawText(g, detail, Fonts.Small, New Rectangle(x, h \ 2 + 1, _patientButton.Width - x, LogicalToDeviceUnits(18)),
                              If(p.Mode = ThemeMode.Blue, p.Muted, p.Muted), TextFormatFlags.Left Or TextFormatFlags.Top Or TextFormatFlags.EndEllipsis)
    End Sub

    Private Shared Function Initials(p As PatientStructure) As String
        Dim a = If(String.IsNullOrEmpty(p.Firstname), "", p.Firstname.Substring(0, 1))
        Dim b = If(String.IsNullOrEmpty(p.Surname), "", p.Surname.Substring(0, 1))
        Return (a & b).ToUpperInvariant()
    End Function

    Public Shared Function PatientDetail(p As PatientStructure) As String
        Dim parts As New List(Of String)
        If p.DoB > New Date(1901, 1, 1) Then
            parts.Add(Loc.F("patients.bornAge", p.DoB.ToString("d", Loc.Culture), Ages.YearsOld(p.DoB)))
        End If
        If Not String.IsNullOrWhiteSpace(p.Reference) Then parts.Add(Loc.F("patients.refShort", p.Reference))
        Return String.Join("  ·  ", parts)
    End Function

#End Region

End Class

Public NotInheritable Class Ages
    Private Sub New()
    End Sub

    ''' <summary>Stored when the date of birth isn't known (as for the Trash patient).</summary>
    Public Shared ReadOnly UnknownDoB As New Date(1900, 1, 1)

    ''' <summary>False for an empty, unreadable or placeholder (1900) date of birth.</summary>
    Public Shared Function IsKnown(dob As Date) As Boolean
        Return dob > New Date(1901, 1, 1)
    End Function

    Public Shared Function YearsOld(dob As Date) As Integer
        Dim today = Date.Today
        Dim age = today.Year - dob.Year
        If dob.Date > today.AddYears(-age) Then age -= 1
        Return Math.Max(0, age)
    End Function
End Class
