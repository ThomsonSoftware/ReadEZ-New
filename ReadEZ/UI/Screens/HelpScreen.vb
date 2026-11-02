Option Strict On

' ====================================================================================================
'   HelpScreen.vb - help videos, support links and About
'
'     +--------------------+  +-----------------------------+  +------------------+
'     | Help topics (tree) |  | Topic: title, description,  |  | Support          |
'     |  Getting started   |  | Watch video, more videos    |  | About            |
'     |  Screens ...       |  |                             |  |                  |
'     +--------------------+  +-----------------------------+  +------------------+
'
'   The topics and their video links come from HelpCatalog (links are set in Admin > Help videos).
'   Videos and pages open in the computer's normal web browser.
' ====================================================================================================

Public Class HelpScreen
    Inherits ScreenBase

    Private Shared ReadOnly Links As (Key As String, Url As String, Glyph As String)() = {
        ("help.link.home", "https://thomson-software-solutions.com/", Glyphs.Home),
        ("help.link.support", "https://thomson-software-solutions.com/support/", Glyphs.Help),
        ("help.link.teamviewer", "https://get.teamviewer.com/qwh3uwf", Glyphs.Export)
    }

    Private ReadOnly lblAbout As New Label
    Private ReadOnly tree As New TreeView With {.BorderStyle = BorderStyle.None, .Font = Fonts.Body, .HideSelection = False,
                                                 .ShowLines = False, .FullRowSelect = True, .Dock = DockStyle.Fill}
    Private ReadOnly lblTopic As New Label With {.AutoSize = True, .Font = Fonts.Section, .Margin = New Padding(0, 0, 0, 2)}
    Private ReadOnly lblGroup As New Label With {.AutoSize = True, .Font = Fonts.Small, .Margin = New Padding(0, 0, 0, 12)}
    Private ReadOnly lblDesc As New Label With {.AutoSize = True, .Font = Fonts.Body, .Margin = New Padding(0, 0, 0, 14)}
    Private ReadOnly lblNoVideo As Label = Ui.Label("help.noVideo", muted:=True)
    Private ReadOnly btnWatch As ModernButton = ModernButton.Create("help.watch", ButtonKind.Primary, Glyphs.Video)
    Private ReadOnly btnRead As ModernButton = ModernButton.Create("help.readTopic", ButtonKind.Secondary, Glyphs.Help)
    Private ReadOnly moreHead As Label = Ui.Label("help.moreVideos")
    Private ReadOnly moreList As FlowLayoutPanel = Ui.Column()
    Private ReadOnly detail As FlowLayoutPanel = Ui.Column()
    Private _topic As HelpTopic

    Public Sub New()
        TitleKey = "help.title"
        SubtitleKey = "help.subtitle"

        ' ----- Topics (left) -----
        Dim topicsCard As New Card With {.Dock = DockStyle.Left, .Width = 380, .Margin = New Padding(0)}
        topicsCard.SetTitle("help.topics")
        tree.ItemHeight = tree.Font.Height + 12
        tree.Indent = 18
        topicsCard.Controls.Add(tree)
        AddHandler tree.AfterSelect, Sub() ShowTopic(TryCast(tree.SelectedNode?.Tag, HelpTopic))
        AddHandler tree.NodeMouseDoubleClick, Sub(s, e)
                                                  Dim t = TryCast(e.Node.Tag, HelpTopic)
                                                  If t IsNot Nothing AndAlso t.HasVideo Then AppHost.OpenUrl(t.Url)
                                              End Sub

        ' ----- Selected topic (middle) -----
        Dim detailCard As New Card With {.Dock = DockStyle.Fill, .Margin = New Padding(0)}
        detailCard.SetTitle("help.topic")
        detail.Dock = DockStyle.Fill
        detail.AutoSize = False
        detail.AutoScroll = True
        LabelStyle.SetMuted(lblGroup)
        lblNoVideo.Margin = New Padding(0, 0, 0, 12)
        btnWatch.Margin = New Padding(0, 0, 0, 8)
        AddHandler btnWatch.Click, Sub() If _topic IsNot Nothing AndAlso _topic.HasVideo Then AppHost.OpenUrl(_topic.Url)
        ' The topic in the help file (shown when the help file is installed)
        btnRead.Margin = New Padding(0, 0, 0, 18)
        AddHandler btnRead.Click, Sub() If _topic IsNot Nothing Then HelpFile.Show(FindForm(), _topic.Id)
        moreHead.Font = Fonts.Strong
        moreHead.Margin = New Padding(0, 0, 0, 4)
        For Each c In New Control() {lblTopic, lblGroup, lblDesc, lblNoVideo, btnWatch, btnRead, moreHead, moreList}
            detail.Controls.Add(c)
        Next
        detailCard.Controls.Add(detail)
        AddHandler detail.Resize, Sub() lblDesc.MaximumSize = New Size(Math.Max(200, detail.ClientSize.Width - 24), 0)

        ' ----- Support and About (right) -----
        Dim side As New Panel With {.Dock = DockStyle.Right, .Width = 360, .Padding = New Padding(12, 0, 0, 0), .BackColor = Color.Transparent}
        Dim support As New Card With {.Dock = DockStyle.Top, .Height = 360}
        support.SetTitle("help.support")
        Dim col = Ui.Column()
        col.Dock = DockStyle.Fill
        ' The whole help manual (when the help file is installed) and the online help
        If HelpFile.Available Then
            Dim manual = ModernButton.Create("menu.helpContents", ButtonKind.Ghost, Glyphs.Help)
            AddHandler manual.Click, Sub() HelpFile.Show(FindForm())
            col.Controls.Add(manual)
        End If
        Dim online = ModernButton.Create("menu.onlineHelp", ButtonKind.Ghost, Glyphs.Globe)
        AddHandler online.Click, Sub() MainMenu.OpenOnlineHelp()
        col.Controls.Add(online)
        For Each l In Links
            Dim url = l.Url
            Dim b = ModernButton.Create(l.Key, ButtonKind.Ghost, l.Glyph)
            AddHandler b.Click, Sub() AppHost.OpenUrl(url)
            col.Controls.Add(b)
        Next
        Dim logs = ModernButton.Create("help.openLogs", ButtonKind.Ghost, Glyphs.Folder)
        AddHandler logs.Click, Sub()
                                   IO.Directory.CreateDirectory(Log.Folder)
                                   AppHost.OpenUrl(Log.Folder)
                               End Sub
        col.Controls.Add(logs)
        Dim licence = ModernButton.Create("help.licence", ButtonKind.Ghost, Glyphs.Lock)
        AddHandler licence.Click, Sub()
                                      Dim main = TryCast(FindForm(), MainForm)
                                      If main IsNot Nothing Then
                                          main.ShowLicence()
                                      Else
                                          Using dlg As New LicenceDialog(LicenceDialogMode.Info)
                                              dlg.ShowDialog(FindForm())
                                          End Using
                                      End If
                                  End Sub
        col.Controls.Add(licence)
        Dim eula = ModernButton.Create("eula.button", ButtonKind.Ghost, Glyphs.Report)
        AddHandler eula.Click, Sub()
                                   Using dlg As New EulaDialog()
                                       dlg.ShowDialog(FindForm())
                                   End Using
                               End Sub
        col.Controls.Add(eula)
        ' Check for updates: in the About card, beside the version number
        Dim updates = ModernButton.Create("help.checkUpdates", ButtonKind.Secondary, Glyphs.Refresh)
        AddHandler updates.Click, Sub()
                                      Dim main = TryCast(FindForm(), MainForm)
                                      If main IsNot Nothing Then
                                          main.ShowUpdate()
                                      Else
                                          Using dlg As New UpdateDialog()
                                              dlg.ShowDialog(FindForm())
                                          End Using
                                      End If
                                  End Sub
        support.Controls.Add(col)

        Dim about As New Card With {.Dock = DockStyle.Top, .Height = 230, .Margin = New Padding(0, 12, 0, 0)}
        about.SetTitle("help.about")
        lblAbout.AutoSize = False
        lblAbout.Dock = DockStyle.Fill
        LabelStyle.SetMuted(lblAbout)
        about.Controls.Add(lblAbout)
        updates.Dock = DockStyle.Bottom
        about.Controls.Add(updates)
        side.Controls.Add(about)
        side.Controls.Add(New Panel With {.Dock = DockStyle.Top, .Height = 12, .BackColor = Color.Transparent})
        side.Controls.Add(support)

        Content.Controls.Add(detailCard)
        Content.Controls.Add(New Panel With {.Dock = DockStyle.Left, .Width = 12, .BackColor = Color.Transparent})
        Content.Controls.Add(topicsCard)
        Content.Controls.Add(side)
        detailCard.BringToFront()               ' the filling card is docked last

        Loc.OnChange(Me, Sub()
                             lblAbout.Text = Loc.F("help.aboutText", AppHost.ProductTitle, AppHost.VersionNumber, AppHost.UdiText,
                                                   GetType(HelpScreen).Assembly.GetName().Version)
                             FillTree()
                         End Sub)
        AddHandler HelpCatalog.TopicsChanged, AddressOf OnLinksChanged
        AddHandler ThemeManager.ThemeChanged, AddressOf OnLinksChanged
    End Sub

    Private Sub OnLinksChanged(sender As Object, e As EventArgs)
        If Not IsDisposed Then FillTree()
    End Sub

    ''' <summary>Groups (bold) with their topics; topics without a video yet are muted.</summary>
    Private Sub FillTree()
        Dim selectedId = _topic?.Id
        tree.BeginUpdate()
        tree.Nodes.Clear()
        Dim p = ThemeManager.Current
        Dim reselect As TreeNode = Nothing
        For Each g In HelpCatalog.Groups
            ' Trailing spaces: a bold node is measured with the tree's normal font and would be clipped
            Dim gn = tree.Nodes.Add(g.Title & "    ")
            gn.NodeFont = Fonts.Strong
            gn.Tag = g
            AddNodes(gn.Nodes, g.Children, p, selectedId, reselect)
        Next
        tree.ExpandAll()
        tree.SelectedNode = If(reselect, If(tree.Nodes.Count > 0, tree.Nodes(0), Nothing))
        tree.SelectedNode?.EnsureVisible()
        If tree.Nodes.Count > 0 Then tree.Nodes(0).EnsureVisible()
        tree.EndUpdate()
        ShowTopic(TryCast(tree.SelectedNode?.Tag, HelpTopic))
    End Sub

    Private Shared Sub AddNodes(nodes As TreeNodeCollection, topics As IEnumerable(Of HelpTopic), p As Palette, selectedId As String, ByRef reselect As TreeNode)
        For Each t In topics
            Dim n = nodes.Add(t.Title)
            n.Tag = t
            If Not t.HasVideo Then n.ForeColor = p.Muted
            If t.Id = selectedId Then reselect = n
            AddNodes(n.Nodes, t.Children, p, selectedId, reselect)
        Next
    End Sub

    Private Sub ShowTopic(t As HelpTopic)
        _topic = t
        detail.SuspendLayout()
        For Each c As Control In moreList.Controls.Cast(Of Control)().ToList()
            c.Dispose()
        Next
        If t Is Nothing Then
            lblTopic.Text = ""
            lblGroup.Text = ""
            lblDesc.Text = Loc.T("help.selectTopic")
            lblNoVideo.Visible = False
            btnWatch.Visible = False
            btnRead.Visible = False
            moreHead.Visible = False
            detail.ResumeLayout(True)
            Return
        End If
        lblTopic.Text = t.Title
        lblGroup.Text = If(t.Parent IsNot Nothing, t.Parent.Title, "")
        If t.IsGroup Then
            lblDesc.Text = Loc.T("help.groupHint")
            lblNoVideo.Visible = False
            btnWatch.Visible = False
        Else
            lblDesc.Text = If(t.Description <> "", t.Description, Loc.T("help.noDescription"))
            lblNoVideo.Visible = Not t.HasVideo
            btnWatch.Visible = t.HasVideo
        End If
        btnRead.Visible = HelpFile.HasPage(t.Id)
        ' Sub-topics and related topics (for a group: its topics)
        Dim more = t.Children.Concat(t.Related.Select(Function(id) HelpCatalog.Find(id)).Where(Function(x) x IsNot Nothing)).ToList()
        moreHead.Visible = more.Count > 0
        For Each m In more
            Dim item = m
            Dim b = ModernButton.Create("", ButtonKind.Ghost, Glyphs.Video)
            b.Text = item.Title & If(item.HasVideo OrElse item.Children.Count > 0, "", "  ·  " & Loc.T("help.comingSoon"))
            b.Margin = New Padding(0, 0, 0, 2)
            AddHandler b.Click, Sub() SelectTopic(item.Id)
            moreList.Controls.Add(b)
            Dpi.ScaleNew(b, moreList)
            ThemeManager.Apply(b)
        Next
        detail.ResumeLayout(True)
    End Sub

    ''' <summary>Selects a topic in the tree (from the "more videos" list).</summary>
    Private Sub SelectTopic(id As String)
        For Each n In AllNodes(tree.Nodes)
            Dim t = TryCast(n.Tag, HelpTopic)
            If t IsNot Nothing AndAlso t.Id = id Then
                tree.SelectedNode = n
                n.EnsureVisible()
                Return
            End If
        Next
    End Sub

    Private Shared Iterator Function AllNodes(nodes As TreeNodeCollection) As IEnumerable(Of TreeNode)
        For Each n As TreeNode In nodes
            Yield n
            For Each c In AllNodes(n.Nodes)
                Yield c
            Next
        Next
    End Function

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveHandler HelpCatalog.TopicsChanged, AddressOf OnLinksChanged
            RemoveHandler ThemeManager.ThemeChanged, AddressOf OnLinksChanged
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
