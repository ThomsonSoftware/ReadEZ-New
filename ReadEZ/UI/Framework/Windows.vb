Option Strict On

Imports System.Drawing.Drawing2D

' ====================================================================================================
'   Windows.vb
'   DialogBase - common layout for dialogs (title area, scrolling body, footer buttons).
'   Toast      - small notification in the corner of a window (replaces DevExpress AlertControl).
'   Ask        - translated message boxes.
' ====================================================================================================

#Region "DialogBase"

Public Class DialogBase
    Inherits Form

    Protected ReadOnly Header As New Panel
    Protected ReadOnly Body As New Panel
    Protected ReadOnly Footer As New FlowLayoutPanel
    Protected ReadOnly TitleLabel As New Label
    Protected ReadOnly SubtitleLabel As New Label

    Public Sub New()
        SuspendLayout()
        AutoScaleMode = AutoScaleMode.None   ' scaled once in OnLoad (see Dpi.vb)
        Font = Fonts.Body
        FormBorderStyle = FormBorderStyle.FixedDialog
        MaximizeBox = False
        MinimizeBox = False
        ShowInTaskbar = False
        StartPosition = FormStartPosition.CenterParent
        KeyPreview = True
        Icon = AppHost.AppIcon

        Header.Dock = DockStyle.Top
        Header.Height = 76
        Header.Padding = New Padding(28, 20, 28, 0)
        ThemeManager.SetRole(Header, SurfaceRole.Surface)

        TitleLabel.AutoSize = True
        TitleLabel.Font = Fonts.Title
        TitleLabel.Location = New Point(26, 16)
        SubtitleLabel.AutoSize = True
        SubtitleLabel.Font = Fonts.Small
        SubtitleLabel.Location = New Point(28, 50)
        LabelStyle.SetMuted(SubtitleLabel)
        Header.Controls.Add(TitleLabel)
        Header.Controls.Add(SubtitleLabel)

        Footer.Dock = DockStyle.Bottom
        Footer.Height = 68
        Footer.FlowDirection = FlowDirection.RightToLeft
        Footer.Padding = New Padding(28, 16, 20, 16)
        ThemeManager.SetRole(Footer, SurfaceRole.SurfaceAlt)
        AddHandler Footer.Paint, Sub(s, e)
                                     Using p As New Pen(ThemeManager.Current.Border)
                                         e.Graphics.DrawLine(p, 0, 0, Footer.Width, 0)
                                     End Using
                                 End Sub

        Body.Dock = DockStyle.Fill
        Body.Padding = New Padding(28, 4, 28, 16)
        Body.AutoScroll = True
        ThemeManager.SetRole(Body, SurfaceRole.Surface)

        Controls.Add(Body)
        Controls.Add(Footer)
        Controls.Add(Header)
        ThemeManager.SetRole(Me, SurfaceRole.Surface)
        ResumeLayout(False)
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        Dpi.ScaleOnce(Me)
        ThemeManager.Apply(Me)
        MyBase.OnLoad(e)
    End Sub

    ''' <summary>Sets the translated title (and optional subtitle) of the dialog.</summary>
    Protected Sub SetTitle(titleKey As String, Optional subtitleKey As String = "")
        Loc.OnChange(Me, Sub()
                             Text = Loc.T(titleKey)
                             TitleLabel.Text = Text
                             SubtitleLabel.Text = If(subtitleKey = "", "", Loc.T(subtitleKey))
                         End Sub)
        SubtitleLabel.Visible = subtitleKey <> ""
        If subtitleKey = "" Then Header.Height = Dpi.Px(Header, 64)
    End Sub

    ''' <summary>Sets an untranslated subtitle (e.g. a patient's name).</summary>
    Protected Sub SetSubtitleText(text As String)
        SubtitleLabel.Text = text
        SubtitleLabel.Visible = text <> ""
        Header.Height = Dpi.Px(Header, If(text = "", 64, 76))
    End Sub

    Private _helpTip As ToolTip

    ''' <summary>The help topic of this window: F1 opens it in the help file (the contents if "").</summary>
    Protected Property HelpTopicId As String = ""

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = Keys.F1 Then
            If Not HelpFile.Show(Me, HelpTopicId) AndAlso HelpCatalog.Find(HelpTopicId) IsNot Nothing AndAlso
               TypeOf Me IsNot HelpTopicDialog Then HelpCatalog.ShowTopic(Me, HelpTopicId)
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    ''' <summary>A "?" button at the top right that opens a help topic (HelpCatalog).</summary>
    Protected Sub AddHelpButton(topicId As String)
        HelpTopicId = If(topicId, "")
        If HelpCatalog.Find(topicId) Is Nothing Then Return
        If _helpTip Is Nothing Then _helpTip = New ToolTip
        Dim b = ModernButton.CreateIcon(Glyphs.Help, _helpTip, "help.button.tip")
        b.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        Header.Controls.Add(b)
        Dim place = Sub() b.Location = New Point(Header.ClientSize.Width - b.Width - Dpi.Px(Header, 20), Dpi.Px(Header, 14))
        place()
        AddHandler Header.Resize, Sub() place()
        AddHandler b.Click, Sub() HelpCatalog.ShowTopic(Me, topicId)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _helpTip?.Dispose()
        MyBase.Dispose(disposing)
    End Sub

    ''' <summary>Adds a footer button. Buttons appear right-to-left in the order added.</summary>
    Protected Function AddFooterButton(key As String, kind As ButtonKind, Optional icon As String = "") As ModernButton
        Dim b = ModernButton.Create(key, kind, icon)
        b.Margin = New Padding(8, 0, 0, 0)
        Footer.Controls.Add(b)
        Return b
    End Function

    ''' <summary>A grid that lays out label/field pairs in N columns.</summary>
    Protected Shared Function NewGrid(columns As Integer) As TableLayoutPanel
        Dim t As New TableLayoutPanel With {
            .ColumnCount = columns,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Dock = DockStyle.Top,
            .Margin = New Padding(0),
            .Padding = New Padding(0)
        }
        For i = 0 To columns - 1
            t.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F / columns))
        Next
        Return t
    End Function
End Class

#End Region

#Region "Toast"

Public Enum ToastKind
    Info
    Success
    Warning
    [Error]
End Enum

''' <summary>Non-blocking notification shown in the bottom-right corner of the main window.</summary>
Public NotInheritable Class Toast
    Inherits Form

    Private ReadOnly _kind As ToastKind
    Private ReadOnly _message As String
    Private ReadOnly _timer As New Timer
    Private _opacityStep As Double = 0.1

    Private Sub New(message As String, kind As ToastKind, durationMs As Integer)
        _kind = kind
        _message = message
        FormBorderStyle = FormBorderStyle.None
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        TopMost = True
        Opacity = 0
        Font = Fonts.Body
        Dim sz = TextRenderer.MeasureText(message, Font, New Size(360, 0), TextFormatFlags.WordBreak)
        Size = New Size(Math.Max(LogicalToDeviceUnits(260), sz.Width + LogicalToDeviceUnits(76)), Math.Max(LogicalToDeviceUnits(56), sz.Height + LogicalToDeviceUnits(32)))
        SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
        _timer.Interval = 20
        Dim shownFor = 0
        AddHandler _timer.Tick, Sub()
                                    If Opacity < 1 AndAlso _opacityStep > 0 Then
                                        Opacity = Math.Min(1, Opacity + _opacityStep)
                                    ElseIf _opacityStep > 0 Then
                                        shownFor += _timer.Interval
                                        If shownFor >= durationMs Then _opacityStep = -0.08
                                    Else
                                        Opacity = Math.Max(0, Opacity + _opacityStep)
                                        If Opacity <= 0 Then
                                            _timer.Stop()
                                            Close()
                                        End If
                                    End If
                                End Sub
        AddHandler Click, Sub() _opacityStep = -0.2
    End Sub

    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property

    Public Shared Shadows Sub Show(owner As Form, message As String, Optional kind As ToastKind = ToastKind.Info, Optional durationMs As Integer = 3500)
        If owner Is Nothing OrElse owner.IsDisposed Then Return
        Dim t As New Toast(message, kind, durationMs)
        Dim area = Screen.FromControl(owner).WorkingArea
        Dim ownerRect = owner.RectangleToScreen(owner.ClientRectangle)
        Dim gap = t.LogicalToDeviceUnits(20)
        Dim x = Math.Min(ownerRect.Right, area.Right) - t.Width - gap
        Dim y = Math.Min(ownerRect.Bottom, area.Bottom) - t.Height - gap
        t.Location = New Point(x, y)
        DirectCast(t, Form).Show(owner)
        t._timer.Start()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        Dim p = ThemeManager.Current
        Dim g = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        Dim accent As Color
        Dim glyph As String
        Select Case _kind
            Case ToastKind.Success : accent = p.Success : glyph = Glyphs.Accept
            Case ToastKind.Warning : accent = p.Warning : glyph = Glyphs.Warning
            Case ToastKind.Error : accent = p.Danger : glyph = Glyphs.Warning
            Case Else : accent = p.Primary : glyph = Glyphs.Info
        End Select
        g.Clear(p.Surface)
        Using pen As New Pen(p.Border)
            g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1)
        End Using
        Using b As New SolidBrush(accent)
            g.FillRectangle(b, 0, 0, LogicalToDeviceUnits(4), Height)
        End Using
        If Fonts.IconLarge IsNot Nothing Then
            TextRenderer.DrawText(g, glyph, Fonts.IconLarge, New Rectangle(LogicalToDeviceUnits(14), 0, LogicalToDeviceUnits(32), Height), accent,
                                  TextFormatFlags.VerticalCenter Or TextFormatFlags.HorizontalCenter)
        End If
        TextRenderer.DrawText(g, _message, Font, New Rectangle(LogicalToDeviceUnits(54), 0, Width - LogicalToDeviceUnits(66), Height), p.Ink,
                              TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.WordBreak)
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _timer.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class

#End Region

#Region "Message boxes"

''' <summary>Translated message boxes. Text is looked up by key; placeholders use Loc.F.</summary>
Public NotInheritable Class Ask
    Private Sub New()
    End Sub

    Public Shared Function YesNo(owner As IWin32Window, messageKey As String, ParamArray args As Object()) As Boolean
        Return MessageBox.Show(owner, Loc.F(messageKey, args), AppHost.ProductTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function

    Public Shared Function YesNoWarning(owner As IWin32Window, messageKey As String, ParamArray args As Object()) As Boolean
        Return MessageBox.Show(owner, Loc.F(messageKey, args), AppHost.ProductTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = DialogResult.Yes
    End Function

    Public Shared Function YesNoCancel(owner As IWin32Window, messageKey As String, ParamArray args As Object()) As DialogResult
        Return MessageBox.Show(owner, Loc.F(messageKey, args), AppHost.ProductTitle, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
    End Function

    Public Shared Sub Info(owner As IWin32Window, messageKey As String, ParamArray args As Object())
        MessageBox.Show(owner, Loc.F(messageKey, args), AppHost.ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Shared Sub Warn(owner As IWin32Window, messageKey As String, ParamArray args As Object())
        MessageBox.Show(owner, Loc.F(messageKey, args), AppHost.ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Public Shared Sub [Error](owner As IWin32Window, messageKey As String, ParamArray args As Object())
        MessageBox.Show(owner, Loc.F(messageKey, args), AppHost.ProductTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    ''' <summary>Simple translated text prompt (replaces InputBox).</summary>
    Public Shared Function Text(owner As IWin32Window, titleKey As String, promptKey As String, Optional initial As String = "", Optional password As Boolean = False) As String
        Using f As New TextPromptDialog(titleKey, promptKey, initial, password)
            If f.ShowDialog(owner) = DialogResult.OK Then Return f.Value
        End Using
        Return Nothing
    End Function

    Private NotInheritable Class TextPromptDialog
        Inherits DialogBase

        Private ReadOnly _field As New TextField

        Public ReadOnly Property Value As String
            Get
                Return _field.Text
            End Get
        End Property

        Public Sub New(titleKey As String, promptKey As String, initial As String, password As Boolean)
            SetTitle(titleKey)
            ClientSize = New Size(440, 230)
            Dim col = Ui.Column()
            col.Dock = DockStyle.Fill
            col.Controls.Add(Ui.Label(promptKey))
            _field.Width = 380
            _field.Text = initial
            _field.Box.UseSystemPasswordChar = password
            col.Controls.Add(_field)
            Body.Controls.Add(col)
            Dim ok = AddFooterButton("common.ok", ButtonKind.Primary)
            Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
            ok.DialogResult = DialogResult.OK
            cancel.DialogResult = DialogResult.Cancel
            AcceptButton = ok
            CancelButton = cancel
            AddHandler Shown, Sub() _field.Box.Focus()
        End Sub
    End Class
End Class

#End Region
