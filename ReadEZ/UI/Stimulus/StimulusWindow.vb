Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   StimulusWindow.vb - the patient's screen (replaces RZfrmTest and RZfrmBlank)
'
'   One borderless window covers the patient's screen (Settings > Screens) for the whole session,
'   exactly as in Eye Dynamix. What it shows is made of two layers:
'     1. Picture  - a still image (the ReadEZ home screen when nothing is running).
'     2. Scene    - something drawn at the window's size: the reading text on white or in colour
'                   (ReadingScene). A scene is any IStimulusScene; it also draws itself into the
'                   clinician's preview, so the preview always shows exactly what the patient sees.
'
'   Two screens: the window stays on the patient's monitor, top-most while a test runs.
'   One screen:  it shares the monitor with the clinician's window and is never top-most; the test
'                screen shows the patient's view in its own large panel, and "Patient full screen"
'                (F11) brings this window to the front with a small control bar (ControlStrip).
'
'   The window is never themed (SurfaceRole.Stimulus): the colours are part of the test.
' ====================================================================================================

''' <summary>Something drawn on the patient's screen at any size.</summary>
Public Interface IStimulusScene
    ''' <summary>Draws the scene into bounds. preview = True when drawing the clinician's small copy.</summary>
    Sub Paint(g As Graphics, bounds As Rectangle, preview As Boolean)
End Interface

Public NotInheritable Class StimulusWindow
    Inherits Form

    Private _picture As Image
    Private _ownsPicture As Boolean
    Private _scene As IStimulusScene

    ''' <summary>Raised whenever what the patient sees changes (the previews redraw).</summary>
    Public Event Changed As EventHandler

    Public Sub New()
        FormBorderStyle = FormBorderStyle.None
        StartPosition = FormStartPosition.Manual
        ShowInTaskbar = False
        BackColor = Color.White
        Text = AppHost.ProductTitle & " - " & Loc.T("stimulus.title")
        Icon = AppHost.AppIcon
        KeyPreview = True
        SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw, True)
        ThemeManager.SetRole(Me, SurfaceRole.Stimulus)
        PlaceOnStimulusScreen()
        ShowLogo()
    End Sub

    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property

#Region "Placement"

    Private _placing As Boolean
    Private _ownMonitor As Boolean
    Private ReadOnly _inUse As New HashSet(Of String)
    Private _topMostSuspended As Boolean
    Private _forcedFront As Boolean

    ''' <summary>True when the patient's screen is on a monitor of its own (two-screen mode).</summary>
    Public ReadOnly Property HasOwnMonitor As Boolean
        Get
            Return _ownMonitor
        End Get
    End Property

    ''' <summary>Moves the window to the patient's screen area from Settings > Screens.</summary>
    Public Sub PlaceOnStimulusScreen()
        Dim r = AppHost.Settings.StimulusBounds
        If Not ScreenLayout.IsOnScreen(r) Then r = ScreenLayout.PrimaryLeadingScreens().Last().Bounds
        If WindowState <> FormWindowState.Normal Then WindowState = FormWindowState.Normal
        _ownMonitor = ScreenLayout.IsTwoScreenMode()
        UpdateTopMost()
        _placing = True
        Try
            Bounds = r
        Finally
            _placing = False
        End Try
        Invalidate()
    End Sub

    ''' <summary>
    ''' Marks the patient's screen as in use by a test or colour check (reason = "test" ...). Only
    ''' while it is in use is the window top-most, and only when it has a monitor of its own.
    ''' </summary>
    Public Sub SetInUse(reason As String, inUse As Boolean)
        If inUse Then _inUse.Add(reason) Else _inUse.Remove(reason)
        UpdateTopMost()
    End Sub

    ''' <summary>One screen: brings the window in front of the clinician's window (F11), or sends it back.</summary>
    Public Sub SetFront(front As Boolean)
        _forcedFront = front
        UpdateTopMost()
        If front Then
            BringToFront()
            Activate()
        End If
    End Sub

    Public ReadOnly Property IsInFront As Boolean
        Get
            Return _forcedFront
        End Get
    End Property

    ''' <summary>Emergency (Ctrl+S): never top-most again until Settings are saved or the program restarts.</summary>
    Public Sub SuspendTopMost()
        _topMostSuspended = True
        _forcedFront = False
        UpdateTopMost()
    End Sub

    ''' <summary>Allows top-most again (called when Settings are saved).</summary>
    Public Sub ResumeTopMost()
        _topMostSuspended = False
        UpdateTopMost()
    End Sub

    Private Sub UpdateTopMost()
        Dim want = Not _topMostSuspended AndAlso (_forcedFront OrElse (_ownMonitor AndAlso _inUse.Count > 0))
        If TopMost = want Then Return
        ' Changing TopMost activates this window; give the focus back to the clinician's window.
        Dim active = Form.ActiveForm
        TopMost = want
        If Not _forcedFront AndAlso active IsNot Nothing AndAlso active IsNot Me AndAlso Not active.IsDisposed Then active.Activate()
    End Sub

#If Not COMPILECHECK Then
    ''' <summary>
    ''' Moving onto a monitor with different scaling makes Windows suggest a rescaled size, which
    ''' would leave the patient's screen part-covered. Ignore it and keep the exact screen area.
    ''' </summary>
    Protected Overrides Sub OnDpiChanged(e As DpiChangedEventArgs)
        e.Cancel = True
        MyBase.OnDpiChanged(e)
        If Not _placing Then BeginInvoke(New Action(AddressOf PlaceOnStimulusScreen))
    End Sub
#End If

#End Region

#Region "Layers"

    ''' <summary>Shows the ReadEZ home screen and removes any scene.</summary>
    Public Sub ShowLogo()
        _scene = HomeScene.Instance
        SetPicture(Nothing)
        BackColor = Color.White
        Invalidate()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub

    ''' <summary>Replaces the still picture. When owns = True the window disposes it later.</summary>
    Public Sub SetPicture(img As Image, Optional owns As Boolean = False)
        Dim old = _picture
        Dim oldOwned = _ownsPicture
        _picture = img
        _ownsPicture = owns
        If oldOwned AndAlso old IsNot Nothing AndAlso old IsNot img Then old.Dispose()
        Invalidate()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub

    ''' <summary>Sets (or clears) the scene. The picture is hidden while a scene is shown.</summary>
    Public Property Scene As IStimulusScene
        Get
            Return _scene
        End Get
        Set(value As IStimulusScene)
            _scene = value
            Invalidate()
            RaiseEvent Changed(Me, EventArgs.Empty)
        End Set
    End Property

    ''' <summary>Redraws (after the scene changed itself), from any thread.</summary>
    Public Sub RequestRedraw()
        If IsDisposed Then Return
        If InvokeRequired Then
            BeginInvoke(Sub() RequestRedraw())
            Return
        End If
        Invalidate()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub

#End Region

#Region "Drawing"

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        e.Graphics.Clear(BackColor)
        DrawLayers(e.Graphics, ClientRectangle, preview:=False)
    End Sub

    ''' <summary>Draws the patient's view into another surface (the clinician's preview).</summary>
    Public Sub RenderPreview(g As Graphics, bounds As Rectangle)
        Using b As New SolidBrush(BackColor)
            g.FillRectangle(b, bounds)
        End Using
        DrawLayers(g, bounds, preview:=True)
    End Sub

    Private Sub DrawLayers(g As Graphics, bounds As Rectangle, preview As Boolean)
        If _scene IsNot Nothing Then
            Try
                _scene.Paint(g, bounds, preview)
            Catch ex As Exception
                Log.Error("Stimulus scene", ex)
            End Try
        ElseIf _picture IsNot Nothing Then
            g.InterpolationMode = If(preview, InterpolationMode.HighQualityBilinear, InterpolationMode.Bilinear)
            g.DrawImage(_picture, bounds)
        End If
    End Sub

#End Region

#Region "Keys and lifetime"

    ''' <summary>Keys pressed on the patient's screen go to the running test (one-screen full screen).</summary>
    Public Event KeyPressedOnScreen(keyData As Keys)

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        RaiseEvent KeyPressedOnScreen(keyData)
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        ' "O" opens the options, as in Eye Dynamix (when no test is using the keys).
        If Char.ToUpperInvariant(e.KeyChar) = "O"c AndAlso _inUse.Count = 0 AndAlso Not _forcedFront Then
            AppHost.MainWindow?.ShowScreen(MainForm.ScreenSettings)
            AppHost.MainWindow?.Activate()
        End If
        MyBase.OnKeyPress(e)
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        ' Only closes with the main window.
        If e.CloseReason = CloseReason.UserClosing Then e.Cancel = True
        MyBase.OnFormClosing(e)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso _ownsPicture Then _picture?.Dispose()
        MyBase.Dispose(disposing)
    End Sub

#End Region

End Class

''' <summary>The ReadEZ home screen shown on the patient's screen between tests.</summary>
Public NotInheritable Class HomeScene
    Implements IStimulusScene

    Public Shared ReadOnly Instance As New HomeScene()

    Public Sub Paint(g As Graphics, bounds As Rectangle, preview As Boolean) Implements IStimulusScene.Paint
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        g.Clear(Color.White)
        Dim k = bounds.Width / CSng(Math.Max(1, StimulusArt.ScreenSize.Width))
        ' The twelve overlay colours round a circle, the name in the middle
        Dim cx = bounds.X + bounds.Width / 2.0F, cy = bounds.Y + bounds.Height / 2.0F
        Dim r = Math.Min(bounds.Width, bounds.Height) * 0.28F
        Dim d = r * 0.42F
        For i = 1 To 12
            Dim a = -Math.PI / 2 + (i - 1) * Math.PI / 6
            Using b As New SolidBrush(ReadEZColours.Colour(ReadEZSet.Overlay, i))
                g.FillEllipse(b, cx + CSng(r * Math.Cos(a)) - d / 2, cy + CSng(r * Math.Sin(a)) - d / 2, d, d)
            End Using
        Next
        Using f As New Font("Segoe UI Semibold", Math.Max(6, 54 * k), GraphicsUnit.Pixel),
              fs As New Font("Segoe UI", Math.Max(5, 18 * k), GraphicsUnit.Pixel),
              ink As New SolidBrush(Brand.Navy), grey As New SolidBrush(Brand.Grey)
            Dim c As New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center}
            g.DrawString("ReadEZ", f, ink, New RectangleF(bounds.X, cy - 40 * k, bounds.Width, 60 * k), c)
            g.DrawString("Thomson Software Solutions", fs, grey, New RectangleF(bounds.X, cy + 18 * k, bounds.Width, 30 * k), c)
        End Using
    End Sub
End Class

''' <summary>Access to the single patient window.</summary>
Public NotInheritable Class Stimulus
    Private Sub New()
    End Sub

    Private Shared _window As StimulusWindow

    ''' <summary>The patient window, created and shown on first use.</summary>
    Public Shared ReadOnly Property Window As StimulusWindow
        Get
            If _window Is Nothing OrElse _window.IsDisposed Then
                _window = New StimulusWindow()
                ' Not owned by the main window: with a single screen the clinician's window must be
                ' able to come in front of it.
                _window.Show()
                ' One screen: keep it behind the clinician's window until a test asks for it
                If Not _window.HasOwnMonitor Then _window.SendToBack()
            End If
            Return _window
        End Get
    End Property

    Public Shared ReadOnly Property IsOpen As Boolean
        Get
            Return _window IsNot Nothing AndAlso Not _window.IsDisposed
        End Get
    End Property

    ''' <summary>Back to the home screen.</summary>
    Public Shared Sub ShowLogo()
        Window.ShowLogo()
    End Sub

    Public Shared Sub CloseWindow()
        If IsOpen Then
            _window.Dispose()
            _window = Nothing
        End If
    End Sub
End Class
