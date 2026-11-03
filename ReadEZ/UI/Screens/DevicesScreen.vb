Option Strict On

Imports System.IO

' ====================================================================================================
'   DevicesScreen.vb - phones and tablets for this practice (as Thomson Near Chart's Devices page)
'
'   Left  : the practice ID (the Remote Test server's serial number for this practice, as used by
'           Near Chart - it can be copied from Near Chart), and the devices registered on the
'           server - add, rename, delete, refresh.
'   Right : how to link the selected device - the QR code to scan, the link to open, Copy - and its
'           status. "Check connection" reads the device again and shows a confirmation on it.
'   Settings > Screens > "Use a tablet as the patient's screen" then sends the tests to it.
' ====================================================================================================

Public Class DevicesScreen
    Inherits ScreenBase

    Public Const DevicesGlyph As String = ChrW(&HE772)
    Private Const RenameGlyph As String = ChrW(&HE8AC)

    Private ReadOnly txtPractice As New TextField With {.Width = 320}
    Private ReadOnly lblPractice As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(330, 0), .Margin = New Padding(0, 4, 0, 8)}
    Private ReadOnly lst As New DeviceListBox With {.Width = 330, .Height = 220, .Margin = New Padding(0, 0, 0, 8)}
    Private ReadOnly btnAdd As ModernButton = ModernButton.Create("dev.add", ButtonKind.Primary, Glyphs.Add)
    Private ReadOnly btnRename As ModernButton = ModernButton.Create("dev.rename", ButtonKind.Secondary, RenameGlyph)
    Private ReadOnly btnDelete As ModernButton = ModernButton.Create("dev.delete", ButtonKind.Danger, Glyphs.Delete)
    Private ReadOnly pctQR As New PictureBox With {.Size = New Size(220, 220), .SizeMode = PictureBoxSizeMode.Zoom, .Margin = New Padding(0, 0, 0, 12)}
    Private ReadOnly txtUrl As New TextField With {.Width = 460}
    Private ReadOnly btnCopy As ModernButton = ModernButton.Create("dev.copy", ButtonKind.Secondary, Glyphs.Copy)
    Private ReadOnly btnCheck As ModernButton = ModernButton.Create("dev.check", ButtonKind.Primary, Glyphs.Refresh)
    Private ReadOnly chip As New StatusChip
    Private ReadOnly lblDetail As New Label With {.AutoSize = True, .Font = Fonts.Body, .MaximumSize = New Size(480, 0), .Margin = New Padding(0, 6, 0, 0)}
    Private _loading As Boolean

    Public Sub New()
        TitleKey = "dev.title"
        SubtitleKey = "dev.subtitle"
        HelpTopic = "screen.devices"
        Dim refresh = AddHeaderButton("dev.refresh", ButtonKind.Secondary, Glyphs.Refresh)
        AddHandler refresh.Click, Async Sub() Await RunAsync(Function() Link.RefreshDevicesAsync())

        Dim row As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoScroll = True, .WrapContents = True, .BackColor = Color.Transparent}

        ' ---- Practice and devices ----
        Dim left As New Card With {.Size = New Size(380, 600), .Padding = New Padding(20, 44, 20, 16)}
        left.SetTitle("dev.list")
        Dim lcol = Ui.Column()
        lcol.Dock = DockStyle.Fill
        txtPractice.Text = AppHost.Settings.CompanyID
        lcol.Controls.Add(Ui.Stack("dev.practiceId", txtPractice))
        Dim practiceButtons = Ui.Row(wrap:=True)
        Dim savePractice = ModernButton.Create("dev.practiceSave", ButtonKind.Secondary, Glyphs.Save)
        Dim fromNear = ModernButton.Create("dev.practiceFromNearChart", ButtonKind.Ghost, Glyphs.Copy)
        For Each b In {savePractice, fromNear}
            b.Margin = New Padding(0, 0, 8, 4)
            practiceButtons.Controls.Add(b)
        Next
        lcol.Controls.Add(practiceButtons)
        LabelStyle.SetMuted(lblPractice)
        lcol.Controls.Add(lblPractice)
        lcol.Controls.Add(Ui.Section("dev.devices"))
        lcol.Controls.Add(lst)
        Dim buttons = Ui.Row(wrap:=True)
        buttons.MaximumSize = New Size(340, 0)
        For Each b In {btnAdd, btnRename, btnDelete}
            b.Margin = New Padding(0, 0, 8, 8)
            buttons.Controls.Add(b)
        Next
        lcol.Controls.Add(buttons)
        left.Controls.Add(lcol)
        row.Controls.Add(left)

        ' ---- Linking ----
        Dim right As New Card With {.Size = New Size(540, 600), .Padding = New Padding(20, 44, 20, 16)}
        right.SetTitle("dev.linkTitle")
        Dim rcol = Ui.Column()
        rcol.Dock = DockStyle.Fill
        Dim help = Ui.Label("dev.linkHelp")
        help.MaximumSize = New Size(480, 0)
        help.Margin = New Padding(0, 0, 0, 12)
        rcol.Controls.Add(help)
        rcol.Controls.Add(pctQR)
        txtUrl.Box.ReadOnly = True
        rcol.Controls.Add(Ui.Stack("dev.link", txtUrl))
        btnCopy.Margin = New Padding(0, 0, 8, 8)
        rcol.Controls.Add(btnCopy)
        rcol.Controls.Add(Ui.Section("dev.status"))
        Dim stateRow = Ui.Row()
        chip.Margin = New Padding(0, 4, 12, 0)
        stateRow.Controls.Add(chip)
        stateRow.Controls.Add(btnCheck)
        rcol.Controls.Add(stateRow)
        LabelStyle.SetMuted(lblDetail)
        rcol.Controls.Add(lblDetail)
        right.Controls.Add(rcol)
        row.Controls.Add(right)

        Content.Controls.Add(row)

        AddHandler savePractice.Click, Async Sub() Await SavePracticeAsync(txtPractice.Text)
        AddHandler fromNear.Click, Async Sub()
                                       Dim id = NearChartPracticeId()
                                       If id = "" Then
                                           Ask.Info(FindForm(), "dev.practiceNoNearChart")
                                           Return
                                       End If
                                       txtPractice.Text = id
                                       Await SavePracticeAsync(id)
                                   End Sub
        AddHandler lst.SelectedIndexChanged, Sub()
                                                 If _loading Then Return
                                                 Link.SelectDevice(lst.SelectedIndex)
                                             End Sub
        AddHandler btnAdd.Click, Async Sub() Await AddAsync()
        AddHandler btnRename.Click, Async Sub() Await RenameAsync()
        AddHandler btnDelete.Click, Async Sub() Await DeleteAsync()
        AddHandler btnCopy.Click, Sub()
                                      If txtUrl.Text <> "" Then
                                          Clipboard.SetText(txtUrl.Text)
                                          Toast.Show(FindForm(), Loc.T("dev.copied"), ToastKind.Success)
                                      End If
                                  End Sub
        AddHandler btnCheck.Click, Async Sub() Await CheckAsync()
        AddHandler Link.DevicesChanged, Sub() RefreshList()
        AddHandler Link.CurrentChanged, Sub() RefreshList()
        AddHandler Link.StatusChanged, Sub() RefreshList()
        Loc.OnChange(Me, AddressOf RefreshList)
    End Sub

    Private Shared ReadOnly Property Link As DeviceLink
        Get
            Return TabletScreen.Link
        End Get
    End Property

    Public Overrides Sub OnActivated()
        MyBase.OnActivated()
        RefreshList()
        If AppHost.Settings.CompanyGuid <> Guid.Empty AndAlso Link.Devices.Count = 0 Then
            Dim unused = RunAsync(Function() Link.RefreshDevicesAsync())
        End If
    End Sub

    Public Overrides Sub OnDeactivated()
        MyBase.OnDeactivated()
        ' As in Near Chart: once a phone/tablet has opened its link, the device is marked as linked.
        Dim unused = Link.MarkLinkedIfConnectedAsync()
    End Sub

    Private Sub RefreshList()
        Dim hasPractice = AppHost.Settings.CompanyGuid <> Guid.Empty
        lblPractice.Text = Loc.T(If(Not hasPractice, "dev.practiceNone",
                                    If(Link.Status = LinkStatus.Offline, "dev.practiceOffline", "dev.practiceOk")))
        _loading = True
        lst.SetItems(Link.Devices.Select(Function(x) x.Name).ToList(), Link.CurrentIndex)
        _loading = False
        btnAdd.Enabled = hasPractice
        Dim d = Link.Current
        btnRename.Enabled = d IsNot Nothing
        btnDelete.Enabled = d IsNot Nothing
        btnCopy.Enabled = d IsNot Nothing
        btnCheck.Enabled = d IsNot Nothing
        If d Is Nothing Then
            txtUrl.Text = ""
            SetQR(Nothing)
            chip.State = ChipState.Neutral
            chip.Text = Loc.T("dev.noDevices")
            lblDetail.Text = Loc.T(If(hasPractice, "dev.noDevicesHelp", "dev.practiceFirst"))
            Return
        End If
        Dim url = DeviceLink.ConnectUrl(d)
        If txtUrl.Text <> url Then
            txtUrl.Text = url
            SetQR(QrCode.Create(url))
        End If
        If d.HasConnected Then
            chip.State = ChipState.Good
            chip.Text = Loc.T("dev.connected")
            lblDetail.Text = Loc.F("dev.detail", d.Width, d.Height, Loc.T(If(d.IsTablet, "dev.tablet", "dev.phone")),
                                   Loc.T(If(d.IsLandscape, "dev.landscape", "dev.portrait"))) &
                             If(AppHost.Settings.UseTablet, "", vbCrLf & Loc.T("dev.turnOnHint"))
        Else
            chip.State = ChipState.Bad
            chip.Text = Loc.T("dev.notConnected")
            lblDetail.Text = Loc.T("dev.notConnectedHelp")
        End If
    End Sub

    Private Sub SetQR(img As Image)
        Dim old = pctQR.Image
        pctQR.Image = img
        old?.Dispose()
    End Sub

#Region "Practice ID"

    Private Async Function SavePracticeAsync(text As String) As Task
        Dim g As Guid
        If Not Guid.TryParse(text.Trim(), g) Then
            Ask.Info(FindForm(), "dev.practiceFormat")
            Return
        End If
        UseWaitCursor = True
        Try
            Dim ok = Await New RemoteTestClient(g).ValidateCompanyAsync()
            If Not ok Then
                Ask.Warn(FindForm(), "dev.practiceUnknown")
                Return
            End If
            AppHost.Settings.CompanyID = g.ToString()
            AppHost.SaveSettings()
            Toast.Show(FindForm(), Loc.T("dev.practiceSaved"), ToastKind.Success)
            Await Link.RefreshDevicesAsync()
        Catch ex As Exception
            Log.Error("Practice ID", ex)
            Ask.Warn(FindForm(), "dev.serverFailed", ex.Message)
        Finally
            UseWaitCursor = False
            RefreshList()
        End Try
    End Function

    ''' <summary>The practice ID saved by Thomson Near Chart on this computer ("" if none).</summary>
    Private Shared Function NearChartPracticeId() As String
        Try
            Dim f = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Thomson Near Chart", "Settings.json")
            If Not File.Exists(f) Then Return ""
            Using doc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(f))
                Dim v As System.Text.Json.JsonElement
                If doc.RootElement.TryGetProperty("CompanyID", v) AndAlso v.ValueKind = System.Text.Json.JsonValueKind.String Then
                    Dim g As Guid
                    If Guid.TryParse(v.GetString(), g) AndAlso g <> Guid.Empty Then Return g.ToString()
                End If
            End Using
        Catch ex As Exception
            Log.Warn("Near Chart practice ID: " & ex.Message)
        End Try
        Return ""
    End Function

#End Region

#Region "Devices"

    Private Async Function AddAsync() As Task
        Dim name = Ask.Text(FindForm(), "dev.addTitle", "dev.addPrompt")
        If String.IsNullOrWhiteSpace(name) Then Return
        Await RunAsync(Function() Link.AddDeviceAsync(name.Trim()))
    End Function

    Private Async Function RenameAsync() As Task
        Dim d = Link.Current
        If d Is Nothing Then Return
        Dim name = Ask.Text(FindForm(), "dev.renameTitle", "dev.renamePrompt", d.Name)
        If String.IsNullOrWhiteSpace(name) OrElse name.Trim() = d.Name Then Return
        Await RunAsync(Function() Link.RenameDeviceAsync(d, name.Trim()))
    End Function

    Private Async Function DeleteAsync() As Task
        Dim d = Link.Current
        If d Is Nothing Then Return
        If Not Ask.YesNoWarning(FindForm(), "dev.confirmDelete", d.Name) Then Return
        Await RunAsync(Function() Link.DeleteDeviceAsync(d))
    End Function

    Private Async Function CheckAsync() As Task
        Await RunAsync(Function() Link.RefreshDevicesAsync())
        Dim d = Link.Current
        If d Is Nothing Then Return
        If Not d.HasConnected Then
            Ask.Info(FindForm(), "dev.notYetConnected")
            Return
        End If
        Using bmp = ConnectedPicture(d)
            Link.SendImage(bmp)
        End Using
        Await Link.MarkLinkedIfConnectedAsync()
        Toast.Show(FindForm(), Loc.T("dev.connectedToast"), ToastKind.Success)
    End Function

    ''' <summary>"ReadEZ is connected" on the device, at its own size.</summary>
    Private Shared Function ConnectedPicture(d As DeviceInfo) As Bitmap
        Dim bmp As New Bitmap(Math.Max(1, d.Width), Math.Max(1, d.Height))
        Using g = Graphics.FromImage(bmp)
            g.Clear(Color.White)
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
            Dim h = CSng(bmp.Height)
            Dim r = h * 0.09F
            Dim cx = bmp.Width / 2.0F, cy = h * 0.38F
            Using b As New SolidBrush(Color.FromArgb(46, 140, 60))
                g.FillEllipse(b, cx - r, cy - r, 2 * r, 2 * r)
            End Using
            Using pen As New Pen(Color.White, r * 0.22F) With {.StartCap = Drawing2D.LineCap.Round, .EndCap = Drawing2D.LineCap.Round}
                g.DrawLines(pen, {New PointF(cx - r * 0.45F, cy), New PointF(cx - r * 0.1F, cy + r * 0.35F), New PointF(cx + r * 0.5F, cy - r * 0.35F)})
            End Using
            Using f As New Font("Segoe UI Semibold", h * 0.05F, GraphicsUnit.Pixel), fmt As New StringFormat With {.Alignment = StringAlignment.Center},
                  ink As New SolidBrush(Color.FromArgb(30, 35, 45))
                g.DrawString(Loc.F("dev.connectedOnDevice", AppHost.ProductTitle), f, ink, New RectangleF(0, cy + r * 1.6F, bmp.Width, h * 0.2F), fmt)
            End Using
        End Using
        Return bmp
    End Function

    ''' <summary>Runs a server call with a wait cursor, reporting failures.</summary>
    Private Async Function RunAsync(work As Func(Of Task)) As Task
        UseWaitCursor = True
        Try
            Await work()
        Catch ex As Exception
            Log.Error("Devices", ex)
            Ask.Warn(FindForm(), "dev.serverFailed", ex.Message)
        Finally
            UseWaitCursor = False
            RefreshList()
        End Try
    End Function

#End Region

End Class

''' <summary>The list of devices (names), in the screen's colours.</summary>
Public NotInheritable Class DeviceListBox
    Inherits ListBox
    Implements IThemed

    Public Sub New()
        DrawMode = DrawMode.OwnerDrawFixed
        ItemHeight = 34
        IntegralHeight = False
        BorderStyle = BorderStyle.FixedSingle
        Font = Fonts.Body
    End Sub

    Public Sub SetItems(names As List(Of String), selected As Integer)
        BeginUpdate()
        Items.Clear()
        For Each n In names
            Items.Add(n)
        Next
        EndUpdate()
        If selected >= 0 AndAlso selected < Items.Count Then SelectedIndex = selected
    End Sub

    Protected Overrides Sub OnDrawItem(e As DrawItemEventArgs)
        If e.Index < 0 OrElse e.Index >= Items.Count Then Return
        Dim p = ThemeManager.Current
        Dim sel = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Using b As New SolidBrush(If(sel, p.PrimarySoft, p.Surface))
            e.Graphics.FillRectangle(b, e.Bounds)
        End Using
        TextRenderer.DrawText(e.Graphics, CStr(Items(e.Index)), Font, New Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 12, e.Bounds.Height), p.Ink,
                              TextFormatFlags.Left Or TextFormatFlags.VerticalCenter Or TextFormatFlags.EndEllipsis)
    End Sub

    Public Sub ApplyTheme(pal As Palette) Implements IThemed.ApplyTheme
        BackColor = pal.Surface
        Invalidate()
    End Sub

    Public ReadOnly Property ThemeChildren As Boolean Implements IThemed.ThemeChildren
        Get
            Return False
        End Get
    End Property
End Class

''' <summary>QR codes for device links (QRCoder, as Near Chart).</summary>
Public NotInheritable Class QrCode
    Private Sub New()
    End Sub

    Public Shared Function Create(text As String) As Image
        If String.IsNullOrEmpty(text) Then Return Nothing
        Try
            Using gen As New QRCoder.QRCodeGenerator(), data = gen.CreateQrCode(text, QRCoder.QRCodeGenerator.ECCLevel.L)
                Dim png = New QRCoder.PngByteQRCode(data).GetGraphic(8)
                Using ms As New MemoryStream(png)
                    Return New Bitmap(Image.FromStream(ms))
                End Using
            End Using
        Catch ex As Exception
            Log.Error("QR", ex)
            Return Nothing
        End Try
    End Function
End Class
