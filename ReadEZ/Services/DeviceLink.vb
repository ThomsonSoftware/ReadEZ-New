Option Strict On

Imports System.Threading

' ====================================================================================================
'   DeviceLink.vb - the phone or tablet used as the patient's screen (from Thomson Near Chart 3)
'
'     - the devices registered for this practice on the Remote Test server, and the one selected
'     - sending pictures: uploads run in the background and only the newest picture is sent, so a
'       quick series of changes never queues up
'     - watching the device: touches (every 200 ms while a test is listening) and the device being
'       rotated or resized (every 2 s)
'
'   As in Near Chart, minus calibration (ReadEZ draws in screen pixels). Events are raised on the UI
'   thread.
' ====================================================================================================

Public Enum LinkStatus
    NotConnected
    Connecting
    Ready
    Sending
    Offline
End Enum

Public NotInheritable Class DeviceLink
    Implements IDisposable

    Private _client As RemoteTestClient
    Private _devices As New List(Of DeviceInfo)
    Private _current As DeviceInfo
    Private _status As LinkStatus = LinkStatus.NotConnected
    Private _lastError As String = ""
    Private ReadOnly _ui As SynchronizationContext

    ' Upload queue (newest wins)
    Private ReadOnly _sendLock As New Object
    Private _pending As Bitmap
    Private _sending As Boolean

    ' Polling
    Private _pollCts As CancellationTokenSource
    Private _drawingMode As Boolean
    Private _lastTouch As TouchInfo
    Private _refreshing As Boolean

    Public Event DevicesChanged As EventHandler
    Public Event CurrentChanged As EventHandler
    Public Event StatusChanged As EventHandler
    ''' <summary>The device reported a new touch (only while touches are being watched).</summary>
    Public Event Touched(sender As Object, t As TouchInfo)

    ''' <summary>Create on the UI thread (events are raised there).</summary>
    Public Sub New()
        _ui = If(SynchronizationContext.Current, New WindowsFormsSynchronizationContext())
        _client = New RemoteTestClient(AppHost.Settings.CompanyGuid)
    End Sub

#Region "State"

    Public ReadOnly Property Client As RemoteTestClient
        Get
            If _client.CompanyId <> AppHost.Settings.CompanyGuid Then _client = New RemoteTestClient(AppHost.Settings.CompanyGuid)
            Return _client
        End Get
    End Property

    Public ReadOnly Property Devices As IReadOnlyList(Of DeviceInfo)
        Get
            Return _devices
        End Get
    End Property

    ''' <summary>The selected device (Nothing if there are none).</summary>
    Public ReadOnly Property Current As DeviceInfo
        Get
            Return _current
        End Get
    End Property

    Public ReadOnly Property CurrentIndex As Integer
        Get
            Return If(_current Is Nothing, -1, _devices.IndexOf(_current))
        End Get
    End Property

    ''' <summary>The current device can show pictures (it has connected and reported its size).</summary>
    Public ReadOnly Property CanDraw As Boolean
        Get
            Return _current IsNot Nothing AndAlso _current.HasConnected
        End Get
    End Property

    Public ReadOnly Property Status As LinkStatus
        Get
            Return _status
        End Get
    End Property

    Public ReadOnly Property LastError As String
        Get
            Return _lastError
        End Get
    End Property

    Private Sub SetStatus(s As LinkStatus, Optional err As String = Nothing)
        If err IsNot Nothing Then _lastError = err
        If s = _status AndAlso err Is Nothing Then Return
        _status = s
        Post(Sub() RaiseEvent StatusChanged(Me, EventArgs.Empty))
    End Sub

    ''' <summary>Queues an action on the UI thread.</summary>
    Private Sub Post(action As Action)
        _ui.Post(Sub(o) action(), Nothing)
    End Sub

    ''' <summary>The URL a phone/tablet opens (or scans as a QR code) to link to a device.</summary>
    Public Shared Function ConnectUrl(d As DeviceInfo) As String
        If d Is Nothing Then Return ""
        Dim baseUrl = AppHost.Settings.ConnectUrl
        If String.IsNullOrWhiteSpace(baseUrl) Then baseUrl = AppSettings.DefaultConnectUrl
        Return baseUrl & d.Id.ToString()
    End Function

#End Region

#Region "Device list"

    ''' <summary>
    ''' Reads all devices from the server and keeps the selected one (by id, then by index).
    ''' Returns False (status Offline) if the server could not be reached.
    ''' </summary>
    Public Async Function RefreshDevicesAsync() As Task(Of Boolean)
        If AppHost.Settings.CompanyGuid = Guid.Empty Then
            SetStatus(LinkStatus.NotConnected, "")
            Return False
        End If
        SetStatus(If(_status = LinkStatus.NotConnected OrElse _status = LinkStatus.Offline, LinkStatus.Connecting, _status))
        Try
            Dim list = Await Client.DeviceGetAllAsync()
            Dim previousId = If(_current?.Id, Guid.Empty)
            Dim wantedId As Guid
            If previousId = Guid.Empty Then Guid.TryParse(AppHost.Settings.SelectedDeviceId, wantedId) Else wantedId = previousId
            _devices = list
            Dim pick = list.FirstOrDefault(Function(d) d.Id = wantedId)
            If pick Is Nothing AndAlso list.Count > 0 Then
                Dim idx = Math.Max(0, Math.Min(list.Count - 1, AppHost.Settings.SelectedDevice))
                pick = list(idx)
            End If
            _current = pick
            RememberSelection()
            SetStatus(LinkStatus.Ready, "")
            Post(Sub()
                     RaiseEvent DevicesChanged(Me, EventArgs.Empty)
                     RaiseEvent CurrentChanged(Me, EventArgs.Empty)
                 End Sub)
            UpdatePolling()
            Return True
        Catch ex As RemoteServerUnavailableException
            SetStatus(LinkStatus.Offline, ex.Message)
            Return False
        Catch ex As RemoteServerFaultException
            SetStatus(LinkStatus.Offline, ex.Message)
            Return False
        End Try
    End Function

    Public Sub SelectDevice(index As Integer)
        If index < 0 OrElse index >= _devices.Count Then Return
        If _devices(index) Is _current Then Return
        _current = _devices(index)
        _lastTouch = Nothing
        RememberSelection()
        AppHost.SaveSettings()
        RaiseEvent CurrentChanged(Me, EventArgs.Empty)
        UpdatePolling()
    End Sub

    Private Sub RememberSelection()
        If _current Is Nothing Then Return
        AppHost.Settings.SelectedDeviceId = _current.Id.ToString()
        AppHost.Settings.SelectedDevice = _devices.IndexOf(_current)
    End Sub

#End Region

#Region "Sending pictures"

    ''' <summary>Sends a picture to the current device. Only the newest picture is kept if one is already being sent.</summary>
    Public Sub SendImage(bmp As Bitmap)
        If bmp Is Nothing OrElse Not CanDraw Then Return
        Dim copy As New Bitmap(bmp)          ' the caller keeps (and disposes) the original
        Dim deviceId = _current.Id
        SyncLock _sendLock
            _pending?.Dispose()
            _pending = copy
            If _sending Then Return
            _sending = True
        End SyncLock
        Task.Run(Function() SendLoopAsync(deviceId))
    End Sub

    Private Async Function SendLoopAsync(deviceId As Guid) As Task
        Do
            Dim bmp As Bitmap
            SyncLock _sendLock
                bmp = _pending
                _pending = Nothing
                If bmp Is Nothing Then
                    _sending = False
                    Exit Do
                End If
            End SyncLock
            SetStatus(LinkStatus.Sending)
            Try
                Dim png As Byte()
                Using bmp
                    png = RemoteTestClient.ToPng(bmp)
                End Using
                Dim imageId = Await Client.UploadImageAsync(png, deviceId).ConfigureAwait(False)
                If imageId <> Guid.Empty Then Await Client.SelectImageAsync(imageId, deviceId).ConfigureAwait(False)
                SetStatus(LinkStatus.Ready, "")
            Catch ex As Exception
                Log.Error("Tablet: send picture", ex)
                SetStatus(LinkStatus.Offline, ex.Message)
            End Try
        Loop
    End Function

#End Region

#Region "Watching the device"

    ''' <summary>
    ''' True: poll every 200 ms and report touches (a test is listening for the patient's answers).
    ''' Otherwise poll every 2 s and only watch for the device being rotated or resized.
    ''' </summary>
    Public Property DrawingMode As Boolean
        Get
            Return _drawingMode
        End Get
        Set(value As Boolean)
            If _drawingMode = value Then Return
            _drawingMode = value
            _lastTouch = Nothing
            UpdatePolling()
        End Set
    End Property

    Public Sub UpdatePolling()
        Dim wanted = CanDraw AndAlso AppHost.Settings.UseTablet
        If wanted Then
            If _pollCts Is Nothing Then
                _pollCts = New CancellationTokenSource()
                Dim token = _pollCts.Token
                Task.Run(Function() PollLoopAsync(token))
            End If
        Else
            StopPolling()
        End If
    End Sub

    Public Sub StopPolling()
        _pollCts?.Cancel()
        _pollCts = Nothing
    End Sub

    Private Async Function PollLoopAsync(token As CancellationToken) As Task
        While Not token.IsCancellationRequested
            Try
                Dim d = _current
                If d IsNot Nothing AndAlso Not _refreshing Then
                    Dim t = Await Client.DeviceTouchedValueAsync(d.Id).ConfigureAwait(False)
                    If t IsNot Nothing AndAlso Not token.IsCancellationRequested Then CheckTouch(d, t)
                End If
            Catch ex As Exception
                ' Polling failures are not reported - the next picture sent shows the status.
            End Try
            Try
                Await Task.Delay(If(_drawingMode, 200, 2000), token).ConfigureAwait(False)
            Catch ex As TaskCanceledException
                Exit While
            End Try
        End While
    End Function

    Private Sub CheckTouch(d As DeviceInfo, t As TouchInfo)
        If _drawingMode Then
            Dim last = _lastTouch
            If last Is Nothing OrElse last.X <> t.X OrElse last.Y <> t.Y OrElse last.FullScreen <> t.FullScreen Then
                ' The first value read is the last touch from before - it is only remembered.
                If last IsNot Nothing Then Post(Sub() RaiseEvent Touched(Me, t))
            End If
        End If
        _lastTouch = t
        ' Rotated or resized: read the device again (the pictures are then drawn at the new size).
        If t.Width > 0 AndAlso t.Height > 0 AndAlso (t.Width <> d.Width OrElse t.Height <> d.Height) Then
            _refreshing = True
            Post(Async Sub()
                     Try
                         Await RefreshDevicesAsync()
                     Finally
                         _refreshing = False
                     End Try
                 End Sub)
        End If
    End Sub

#End Region

#Region "Device management (Devices screen)"

    Public Async Function AddDeviceAsync(name As String) As Task(Of DeviceInfo)
        Dim d = Await Client.DeviceAddAsync(name)
        If d IsNot Nothing Then
            AppHost.Settings.SelectedDeviceId = d.Id.ToString()
            _current = Nothing
        End If
        Await RefreshDevicesAsync()
        Return d
    End Function

    Public Async Function RenameDeviceAsync(d As DeviceInfo, name As String) As Task
        Await Client.DeviceUpdateNameAsync(d.Id, name)
        Await RefreshDevicesAsync()
    End Function

    Public Async Function DeleteDeviceAsync(d As DeviceInfo) As Task
        StopPolling()
        Await Client.DeviceDeleteAsync(d.Id)
        If _current IsNot Nothing AndAlso _current.Id = d.Id Then
            _current = Nothing
            AppHost.Settings.SelectedDeviceId = ""
            AppHost.Settings.SelectedDevice = 0
        End If
        Await RefreshDevicesAsync()
    End Function

    ''' <summary>Marks the current device as linked once a phone/tablet has opened its link.</summary>
    Public Async Function MarkLinkedIfConnectedAsync() As Task
        Dim d = _current
        If d Is Nothing OrElse d.Linked OrElse Not d.ClientConnected.HasValue Then Return
        Try
            Await Client.DeviceUpdateAsync("DeviceUpdateLinked", d.Id, True)
            d.Linked = True
        Catch ex As Exception
            Log.Error("DeviceUpdateLinked", ex)
        End Try
    End Function

#End Region

    Public Sub Dispose() Implements IDisposable.Dispose
        StopPolling()
    End Sub

End Class
