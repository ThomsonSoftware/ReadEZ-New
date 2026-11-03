Option Strict On

' ====================================================================================================
'   TabletScreen.vb - a tablet as the patient's screen (Settings > Screens > Tablet)
'
'   Uses the Near Chart's technology (DeviceLink, the Remote Test server): the tablet opens its
'   device link once (Devices screen: QR code) and then shows whatever it is sent.
'
'     - Everything the patient's screen would show (the reading text, colours, instructions, the
'       patient's answer buttons) is drawn at the tablet's own size and sent to it whenever it
'       changes (newest picture wins). The patient's window on the computer is hidden meanwhile.
'     - The tablet then counts as the second screen: two-screen mode, the test screen shows a copy.
'     - While a test is running the tablet is watched for touches (every 200 ms); a touch is passed
'       on as a click on the patient's screen, so it answers exactly as a mouse click would (the
'       side of the text, The first / The second / Show again ...).
'     - Turning the tablet (or resizing its browser) is noticed and the pictures are redrawn.
'
'   Limits of the Remote Test server: each picture takes a moment to arrive (so timed presentations
'   can look a little shorter on the tablet), and two touches on exactly the same pixel count as one.
' ====================================================================================================

Public NotInheritable Class TabletScreen
    Private Sub New()
    End Sub

    Private Shared _link As DeviceLink
    Private Shared ReadOnly _sendTimer As New Timer With {.Interval = 40}
    Private Shared _started As Boolean
    Private Shared _listening As Boolean

    ''' <summary>The tablet was turned on or off, connected, renamed or turned round.</summary>
    Public Shared Event Changed As EventHandler

    ''' <summary>The link to the Remote Test server (created on first use, on the UI thread).</summary>
    Public Shared ReadOnly Property Link As DeviceLink
        Get
            If _link Is Nothing Then _link = New DeviceLink()
            Return _link
        End Get
    End Property

    ''' <summary>True when the tablet is the patient's screen now (turned on, and a device has connected).</summary>
    Public Shared ReadOnly Property IsActive As Boolean
        Get
            Return AppHost.Settings IsNot Nothing AndAlso AppHost.Settings.UseTablet AndAlso _link IsNot Nothing AndAlso _link.CanDraw
        End Get
    End Property

    ''' <summary>The tablet's screen in pixels (as it reported it).</summary>
    Public Shared ReadOnly Property ScreenSize As Size
        Get
            Dim d = _link?.Current
            If d Is Nothing Then Return Size.Empty
            Return New Size(Math.Max(1, d.Width), Math.Max(1, d.Height))
        End Get
    End Property

    ''' <summary>The device's name (for the top bar and Settings).</summary>
    Public Shared ReadOnly Property DeviceName As String
        Get
            Return If(_link?.Current?.Name, "")
        End Get
    End Property

    ''' <summary>Called once the main window is up.</summary>
    Public Shared Sub Start()
        If _started Then Return
        _started = True
        AddHandler _sendTimer.Tick, Sub() SendNow()
        AddHandler Link.CurrentChanged, Sub() Apply()
        AddHandler Link.DevicesChanged, Sub() Apply()
        AddHandler Link.StatusChanged, Sub() RaiseEvent Changed(Nothing, EventArgs.Empty)
        AddHandler Link.Touched, AddressOf OnTouched
        AddHandler Stimulus.Window.Changed, Sub() ScheduleSend()
        If AppHost.Settings.UseTablet AndAlso AppHost.Settings.CompanyGuid <> Guid.Empty Then
            Dim unused = Link.RefreshDevicesAsync()
        End If
        Apply()
    End Sub

    ''' <summary>Turns the tablet on or off as the patient's screen (Settings).</summary>
    Public Shared Sub SetEnabled(useTablet As Boolean)
        AppHost.Settings.UseTablet = useTablet
        AppHost.SaveSettings()
        If useTablet AndAlso Not Link.CanDraw AndAlso AppHost.Settings.CompanyGuid <> Guid.Empty Then
            Dim unused = Link.RefreshDevicesAsync()
        End If
        Apply()
    End Sub

    ''' <summary>A test is running: watch the tablet for the patient's touches (every 200 ms).</summary>
    Public Shared Sub SetListening(listening As Boolean)
        _listening = listening
        If _link IsNot Nothing Then _link.DrawingMode = listening AndAlso IsActive
    End Sub

    ''' <summary>Shows or hides the computer's patient window, starts or stops watching, sends the picture.</summary>
    Public Shared Sub Apply()
        Dim active = IsActive
        Dim w = Stimulus.Window
        If active Then
            If w.Visible Then w.Hide()
        ElseIf Not w.Visible Then
            w.Show()
            w.PlaceOnStimulusScreen()
            If Not w.HasOwnMonitor Then w.SendToBack()
        End If
        w.PlaceOnStimulusScreen()           ' (one or two screens may have changed)
        Link.DrawingMode = active AndAlso _listening
        Link.UpdatePolling()
        If active Then
            w.RequestRedraw()                ' the layout follows the tablet's size
            ScheduleSend()
        End If
        RaiseEvent Changed(Nothing, EventArgs.Empty)
    End Sub

    Private Shared Sub ScheduleSend()
        If Not IsActive Then Return
        _sendTimer.Stop()
        _sendTimer.Start()                   ' a burst of changes is sent once
    End Sub

    ''' <summary>Draws the patient's screen at the tablet's size and sends it.</summary>
    Private Shared Sub SendNow()
        _sendTimer.Stop()
        If Not IsActive Then Return
        Dim sz = ScreenSize
        Try
            Using bmp As New Bitmap(sz.Width, sz.Height)
                Using g = Graphics.FromImage(bmp)
                    Stimulus.Window.RenderPatient(g, New Rectangle(Point.Empty, sz))
                End Using
                Link.SendImage(bmp)          ' (sends a copy)
            End Using
        Catch ex As Exception
            Log.Error("Tablet: draw picture", ex)
        End Try
    End Sub

    ''' <summary>A touch on the tablet: passed on as a click on the patient's screen.</summary>
    Private Shared Sub OnTouched(sender As Object, t As TouchInfo)
        If Not IsActive OrElse t Is Nothing Then Return
        Dim sz = ScreenSize
        Dim x = t.X, y = t.Y
        ' The touch is in the tablet's coordinates at the time; the picture is drawn at its reported size
        If t.Width > 0 AndAlso t.Height > 0 AndAlso (t.Width <> sz.Width OrElse t.Height <> sz.Height) Then
            x = CInt(t.X * sz.Width / CDbl(t.Width))
            y = CInt(t.Y * sz.Height / CDbl(t.Height))
        End If
        Stimulus.Window.SimulateClick(New Rectangle(Point.Empty, sz), New MouseEventArgs(MouseButtons.Left, 1, x, y, 0))
    End Sub
End Class
