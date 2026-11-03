Option Strict On

' ====================================================================================================
'   ScreenBase.vb
'   Every page shown in the main window (Patients, Calibrate, Tests, Results ...) inherits this.
'   The main window calls OnActivated / OnDeactivated as the user moves between screens, so a
'   screen only does work (timers, live tracking) while it is visible.
'   Every screen has the brand look (UseBrandHeader): the title on a dark blue band, white
'   header buttons, and card titles in the brand accent.
' ====================================================================================================

Public Class ScreenBase
    Inherits UserControl
    Implements IBrandedHost

    Private _branded As Boolean

    ''' <summary>True after UseBrandHeader: the cards on the screen take the brand accent.</summary>
    Public ReadOnly Property Branded As Boolean Implements IBrandedHost.Branded
        Get
            Return _branded
        End Get
    End Property

    ''' <summary>
    ''' Brand look (as the analysis windows): the title in white on a dark blue band, the header
    ''' buttons on the band, a gap under it, and brand-accent card titles. The band is coloured here,
    ''' not by the theme.
    ''' </summary>
    Protected Sub UseBrandHeader()
        _branded = True
        ThemeManager.SetRole(HeaderBar, SurfaceRole.Stimulus)
        HeaderBar.BackColor = Brand.Navy
        HeaderBar.Height = 70
        HeaderBar.Padding = New Padding(0, 0, 12, 0)
        _title.Location = New Point(16, 8)
        _title.BackColor = Brand.Navy
        _title.ForeColor = Color.White
        _subtitle.Location = New Point(18, 44)
        _subtitle.BackColor = Brand.Navy
        _subtitle.ForeColor = Color.FromArgb(205, 222, 242)
        HeaderActions.BackColor = Brand.Navy
        HeaderActions.Padding = New Padding(0, 16, 0, 0)
        _help.Kind = ButtonKind.OnDark
        Content.Padding = New Padding(0, 12, 0, 0)
    End Sub

    ''' <summary>Translation key of the title shown at the top of the screen.</summary>
    Public Property TitleKey As String = ""

    ''' <summary>Translation key of the line under the title.</summary>
    Public Property SubtitleKey As String = ""

    ''' <summary>Area for the screen's own content (below the title).</summary>
    Protected ReadOnly Content As New Panel
    Protected ReadOnly HeaderBar As New Panel
    Protected ReadOnly HeaderActions As New FlowLayoutPanel
    Private ReadOnly _title As New Label
    Private ReadOnly _subtitle As New Label
    Private ReadOnly _helpTip As New ToolTip
    Private ReadOnly _help As ModernButton

    ''' <summary>Help topic for the "?" button at the right of the header ("" = no button). See HelpCatalog.</summary>
    Public Property HelpTopic As String = ""

    ''' <summary>The topic the "?" button opens now (the test screen uses the running test's).</summary>
    Protected Overridable Function CurrentHelpTopic() As String
        Return HelpTopic
    End Function

    ''' <summary>The help topic for this screen now (F1 opens it in the help file).</summary>
    Public Function HelpTopicNow() As String
        Return CurrentHelpTopic()
    End Function

    Public Sub New()
        SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
        Dock = DockStyle.Fill
        Font = Fonts.Body
        Padding = New Padding(24, 18, 24, 18)
        ThemeManager.SetRole(Me, SurfaceRole.Canvas)

        HeaderBar.Dock = DockStyle.Top
        HeaderBar.Height = 64
        HeaderBar.BackColor = Color.Transparent
        _title.AutoSize = True
        _title.Font = Fonts.Title
        _title.Location = New Point(0, 2)
        _subtitle.AutoSize = True
        _subtitle.Font = Fonts.Small
        _subtitle.Location = New Point(2, 36)
        LabelStyle.SetMuted(_subtitle)
        HeaderActions.Dock = DockStyle.Right
        HeaderActions.AutoSize = True
        HeaderActions.WrapContents = False
        HeaderActions.FlowDirection = FlowDirection.RightToLeft
        HeaderActions.Padding = New Padding(0, 8, 0, 0)
        HeaderActions.BackColor = Color.Transparent
        ' "?" first, so it is the right-most header button
        _help = ModernButton.CreateIcon(Glyphs.Help, _helpTip, "help.button.tip")
        _help.Margin = New Padding(8, 1, 0, 0)
        _help.Visible = False
        AddHandler _help.Click, Sub() HelpCatalog.ShowTopic(FindForm(), CurrentHelpTopic())
        HeaderActions.Controls.Add(_help)
        HeaderBar.Controls.Add(HeaderActions)
        HeaderBar.Controls.Add(_title)
        HeaderBar.Controls.Add(_subtitle)

        Content.Dock = DockStyle.Fill
        Content.BackColor = Color.Transparent
        Controls.Add(Content)
        Controls.Add(HeaderBar)
        UseBrandHeader()                ' every screen: dark blue title band, brand card titles

        Loc.OnChange(Me, AddressOf RefreshTitle)
    End Sub

    Protected Sub RefreshTitle()
        _title.Text = If(TitleKey = "", "", Loc.T(TitleKey))
        _subtitle.Text = If(SubtitleKey = "", "", Loc.T(SubtitleKey))
        _subtitle.Visible = SubtitleKey <> ""
        _help.Visible = HelpCatalog.Find(CurrentHelpTopic()) IsNot Nothing
    End Sub

    Private _prominent As Boolean
    Private _divider As Panel

    ''' <summary>
    ''' Shows the subtitle (the patient's name) as large as the title, on the same line to the right
    ''' of it after a thin divider, instead of in small text underneath.
    ''' </summary>
    Protected Sub UseProminentSubtitle()
        If _prominent Then Return
        _prominent = True
        _subtitle.Font = Fonts.Title
        _subtitle.ForeColor = If(_branded, Color.White, ThemeManager.Current.Ink)
        LabelStyle.SetCustom(_subtitle)
        _divider = New Panel With {.Width = 2, .BackColor = If(_branded, Color.FromArgb(120, 150, 195), ThemeManager.Current.Border)}
        HeaderBar.Controls.Add(_divider)
        AddHandler _title.SizeChanged, Sub() PlaceSubtitle()
        AddHandler _title.LocationChanged, Sub() PlaceSubtitle()
        AddHandler _subtitle.VisibleChanged, Sub() PlaceSubtitle()
        PlaceSubtitle()
    End Sub

    Private Sub PlaceSubtitle()
        If Not _prominent Then Return
        Dim gap = Dpi.Px(Me, 18)
        _divider.Bounds = New Rectangle(_title.Right + gap, _title.Top + Dpi.Px(Me, 6), _divider.Width, Math.Max(4, _title.Height - Dpi.Px(Me, 10)))
        _divider.Visible = _subtitle.Visible
        _subtitle.Location = New Point(_divider.Right + gap, _title.Top)
    End Sub

    ''' <summary>Replaces the subtitle with untranslated text (e.g. a patient name).</summary>
    Protected Sub SetSubtitleText(text As String)
        _subtitle.Text = text
        _subtitle.Visible = text <> ""
    End Sub

    ''' <summary>Adds a button to the right of the title. Added right-to-left.</summary>
    Protected Function AddHeaderButton(key As String, kind As ButtonKind, Optional icon As String = "") As ModernButton
        If _branded Then
            ' On the dark blue band: plain buttons in white, the main one as a white button
            If kind = ButtonKind.Ghost Then kind = ButtonKind.OnDark
            If kind = ButtonKind.Primary Then kind = ButtonKind.Secondary
        End If
        Dim b = ModernButton.Create(key, kind, icon)
        b.Margin = New Padding(8, 0, 0, 0)
        HeaderActions.Controls.Add(b)
        Return b
    End Function

    ''' <summary>Called each time the screen becomes visible.</summary>
    Public Overridable Sub OnActivated()
        RefreshTitle()
    End Sub

    ''' <summary>Called when the user leaves the screen. Return False to stay (e.g. a test is running).</summary>
    Public Overridable Function CanLeave() As Boolean
        Return True
    End Function

    ''' <summary>Called after the user has left the screen.</summary>
    Public Overridable Sub OnDeactivated()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _helpTip.Dispose()
        MyBase.Dispose(disposing)
    End Sub
End Class

''' <summary>Shown for modules that are scheduled for a later phase of the rewrite.</summary>
Public Class ComingSoonScreen
    Inherits ScreenBase

    Public Sub New(titleKey As String)
        Me.TitleKey = titleKey
        Dim card As New Card With {.Dock = DockStyle.Top, .Height = 120}
        Dim l = Ui.Label("common.comingSoon")
        l.Location = New Point(20, 20)
        l.MaximumSize = New Size(700, 0)
        card.Controls.Add(l)
        Content.Controls.Add(card)
    End Sub
End Class
