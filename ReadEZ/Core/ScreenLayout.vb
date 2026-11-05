Option Strict On

' ====================================================================================================
'   ScreenLayout.vb
'   Helpers for the screen set-up: the clinician uses the control screen, the patient looks at the
'   patient's (stimulus) screen. Both can be on one monitor (one-screen mode: the test screen shows
'   the patient's view itself) or on two (two-screen mode), as in Eye Dynamix.
' ====================================================================================================

Public NotInheritable Class ScreenLayout
    Private Sub New()
    End Sub

    ''' <summary>All screens with the primary screen first (getPrimaryLeadingScreenArray).</summary>
    Public Shared Function PrimaryLeadingScreens() As List(Of Screen)
        Dim list As New List(Of Screen) From {Screen.PrimaryScreen}
        list.AddRange(Screen.AllScreens.Where(Function(s) Not s.Primary))
        Return list
    End Function

    ''' <summary>Rectangle that encloses every screen (getWiderScreensRect).</summary>
    Public Shared Function AllScreensRect() As Rectangle
        Dim screens = PrimaryLeadingScreens()
        Dim left = screens.Min(Function(s) s.Bounds.Left)
        Dim top = screens.Min(Function(s) s.Bounds.Top)
        Dim right = screens.Max(Function(s) s.Bounds.Right)
        Dim bottom = screens.Max(Function(s) s.Bounds.Bottom)
        Return Rectangle.FromLTRB(Math.Min(0, left), Math.Min(0, top), right, bottom)
    End Function

    ''' <summary>
    ''' Area of screen i for a window: the full bounds (patient screen, covers the taskbar) or the
    ''' working area (control window, leaves the taskbar visible on whichever screen has it).
    ''' </summary>
    Public Shared Function ScreenRect(index As Integer, fullScreen As Boolean) As Rectangle
        Dim screens = PrimaryLeadingScreens()
        If index < 0 OrElse index >= screens.Count Then Return Rectangle.Empty
        Return If(fullScreen, screens(index).Bounds, screens(index).WorkingArea)
    End Function

    ''' <summary>The screen holding the centre of r (or the nearest one).</summary>
    Public Shared Function ScreenOf(r As Rectangle) As Screen
        Return Screen.FromPoint(New Point(r.Left + r.Width \ 2, r.Top + r.Height \ 2))
    End Function

    ''' <summary>True if r is usable: big enough and mostly on an attached screen.</summary>
    Public Shared Function IsOnScreen(r As Rectangle) As Boolean
        If r.Width < 100 OrElse r.Height < 100 Then Return False
        Return Screen.AllScreens.Any(Function(sc) sc.Bounds.Contains(New Point(r.Left + r.Width \ 2, r.Top + r.Height \ 2)))
    End Function

    ''' <summary>Rectangle centred on the control window.</summary>
    Public Shared Function CentredOnControl(size As Size) As Point
        Dim c = AppHost.Settings.ControlBounds
        If c.Width <= 0 Then c = Screen.PrimaryScreen.WorkingArea
        Return New Point(c.Left + (c.Width - size.Width) \ 2, c.Top + (c.Height - size.Height) \ 2)
    End Function

    ''' <summary>
    ''' Two-screen mode: the patient's screen (Settings > Screens) is on a different monitor from
    ''' the clinician's window. Otherwise both share one monitor.
    ''' </summary>
    Public Shared Function IsTwoScreenMode() As Boolean
        Dim s = AppHost.Settings
        If s Is Nothing Then Return False
        If TabletScreen.IsActive Then Return True     ' a tablet is the patient's screen
        If s.ForceOneScreen Then Return False         ' one screen chosen, even with a second monitor
        If s.UseTablet Then Return False              ' tablet chosen but not connected yet: one screen meanwhile
        Dim stim = s.StimulusBounds
        If Not IsOnScreen(stim) Then stim = PrimaryLeadingScreens().Last().Bounds
        Dim ctl = s.ControlBounds
        If Not IsOnScreen(ctl) Then ctl = Screen.PrimaryScreen.WorkingArea
        Return ScreenOf(ctl).DeviceName <> ScreenOf(stim).DeviceName
    End Function

    ''' <summary>
    ''' Where the patient's screen goes: the area set in Settings > Screens - or, when one screen
    ''' is chosen, the whole monitor the control window is on.
    ''' </summary>
    Public Shared Function PatientScreenRect() As Rectangle
        Dim s = AppHost.Settings
        If Not TabletScreen.IsActive Then
            If s.ForceOneScreen Then Return OneScreenMonitor().Bounds
            If s.UseTablet Then                        ' waiting for the tablet: the control screen meanwhile
                Dim ctl = s.ControlBounds
                If Not IsOnScreen(ctl) Then ctl = Screen.PrimaryScreen.WorkingArea
                Return ScreenOf(ctl).Bounds
            End If
        End If
        Dim r = s.StimulusBounds
        If Not IsOnScreen(r) Then r = PrimaryLeadingScreens().Last().Bounds
        Return r
    End Function

    ''' <summary>One screen chosen: the monitor chosen for it (screen 1 if that monitor is no longer attached).</summary>
    Public Shared Function OneScreenMonitor() As Screen
        Dim screens = PrimaryLeadingScreens()
        Dim i = AppHost.Settings.OneScreenMonitor
        Return If(i >= 0 AndAlso i < screens.Count, screens(i), screens(0))
    End Function

    ''' <summary>True if a second screen is available for the patient.</summary>
    Public Shared ReadOnly Property HasStimulusScreen As Boolean
        Get
            Return Screen.AllScreens.Length > 1
        End Get
    End Property
End Class
