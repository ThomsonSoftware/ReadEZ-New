Option Strict On

' ====================================================================================================
'   EmergencyShow.vb - Ctrl+S brings the control window back
'
'   If Settings > Screens is wrong (a monitor removed, a different PC, the patient's screen over the
'   clinician's), the program can end up hidden behind the patient's screen or off-screen.
'   Ctrl+S, pressed in any ReadEZ window (including the patient's screen), then:
'     - stops the patient's screen being top-most until Settings are saved or the program restarts;
'     - moves the control window, and any dialog that is open, to the centre of the primary monitor
'       and brings them to the front.
'   It works as an application-wide message filter, so it does not matter which window has focus.
' ====================================================================================================

Public NotInheritable Class EmergencyShow
    Implements IMessageFilter

    Private Const WM_KEYDOWN As Integer = &H100

    Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
        If m.Msg <> WM_KEYDOWN Then Return False
        If (CInt(m.WParam.ToInt64()) And &HFFFF) <> CInt(Keys.S) Then Return False
        If Control.ModifierKeys <> Keys.Control Then Return False
        If (m.LParam.ToInt64() And &H40000000L) <> 0 Then Return True      ' key held down: act once
        Try
            Run()
        Catch ex As Exception
            Log.Error("Emergency show", ex)
        End Try
        Return True
    End Function

    ''' <summary>Puts the control window (and open dialogs) in front, centred on the primary monitor.</summary>
    Public Shared Sub Run()
        Dim main = AppHost.MainWindow
        If main Is Nothing OrElse main.IsDisposed Then Return
        Log.Info("Ctrl+S: control window moved to the primary monitor")
        If Stimulus.IsOpen Then Stimulus.Window.SuspendTopMost()

        Dim area = Screen.PrimaryScreen.WorkingArea
        BringToPrimary(main, area, fullSizeIfMaximised:=True)
        ' Dialogs after the main window, in the order they were opened, so the newest ends on top
        For Each f In Application.OpenForms.Cast(Of Form)().ToList()
            If f Is main OrElse TypeOf f Is StimulusWindow OrElse TypeOf f Is Toast Then Continue For
            If f.IsDisposed OrElse f.InvokeRequired OrElse Not f.Visible Then Continue For   ' other threads: stimulus overlays
            BringToPrimary(f, area, fullSizeIfMaximised:=False)
        Next
        Toast.Show(main, Loc.T("app.emergency"), ToastKind.Warning, 6000)
    End Sub

    Private Shared Sub BringToPrimary(f As Form, area As Rectangle, fullSizeIfMaximised As Boolean)
        Dim wasMaximised = f.WindowState = FormWindowState.Maximized
        If f.WindowState <> FormWindowState.Normal Then f.WindowState = FormWindowState.Normal
        Dim w = f.Width, h = f.Height
        If wasMaximised AndAlso fullSizeIfMaximised Then
            w = area.Width * 9 \ 10
            h = area.Height * 9 \ 10
        End If
        w = Math.Min(w, area.Width)
        h = Math.Min(h, area.Height)
        f.Bounds = New Rectangle(area.X + (area.Width - w) \ 2, area.Y + (area.Height - h) \ 2, w, h)
        ' Top-most for a moment puts it above everything; it then stays in front of normal windows.
        f.TopMost = True
        f.Activate()
        f.TopMost = False
    End Sub
End Class
