Option Strict On

' ====================================================================================================
'   HelpDialogs.vb
'     HelpTopicDialog   - the "?" button on each screen, test and analysis: what the screen is for,
'                         its video, and any related videos
'     HelpTopicsDialog  - Admin > Help topics and videos: add, delete and reorder topics and groups,
'                         and set each topic's video link, title and description
'     EulaDialog        - About > Licence agreement (the same text as the installer shows)
' ====================================================================================================

Public NotInheritable Class HelpTopicDialog
    Inherits DialogBase

    Private ReadOnly _topic As HelpTopic

    Public Sub New(topic As HelpTopic)
        _topic = topic
        Dim hasGroup = topic.Parent IsNot Nothing
        SetTitle("help.title", If(hasGroup, "help.title", ""))
        ' Topic titles can be typed in Admin, so they are set as text (after SetTitle's own)
        Loc.OnChange(Me, Sub()
                             Me.Text = topic.Title
                             TitleLabel.Text = topic.Title
                             SubtitleLabel.Text = If(hasGroup, topic.Parent.Title, "")
                         End Sub)
        ClientSize = New Size(580, 440)

        Dim col = Ui.Column()
        Dim text As New Label With {.AutoSize = True, .MaximumSize = New Size(520, 0), .Font = Fonts.Body, .Margin = New Padding(0, 0, 0, 12)}
        Loc.OnChange(text, Sub() text.Text = If(topic.Description <> "", topic.Description, Loc.T("help.noDescription")))
        col.Controls.Add(text)
        If Not topic.HasVideo Then
            Dim none = Ui.Label("help.noVideo", muted:=True)
            none.Margin = New Padding(0, 0, 0, 12)
            col.Controls.Add(none)
        End If

        ' Sub-topics and related topics, each a link to its video
        Dim more = topic.Children.Concat(topic.Related.Select(Function(id) HelpCatalog.Find(id)).Where(Function(x) x IsNot Nothing)).ToList()
        If more.Count > 0 Then
            Dim head = Ui.Label("help.moreVideos")
            head.Font = Fonts.Strong
            head.Margin = New Padding(0, 4, 0, 4)
            col.Controls.Add(head)
            For Each m In more
                Dim item = m
                Dim b = ModernButton.Create("", ButtonKind.Ghost, Glyphs.Video)
                Loc.OnChange(b, Sub() b.Text = item.Title & If(item.HasVideo, "", "  ·  " & Loc.T("help.comingSoon")))
                b.Enabled = item.HasVideo
                b.Margin = New Padding(0, 0, 0, 2)
                AddHandler b.Click, Sub() AppHost.OpenUrl(item.Url)
                col.Controls.Add(b)
            Next
        End If
        ' Docked, so it sits inside the body's padding (an undocked child ignores Padding)
        col.Dock = DockStyle.Top
        Body.Controls.Add(col)

        Dim close = AddFooterButton("common.close", ButtonKind.Secondary)
        close.DialogResult = DialogResult.Cancel
        CancelButton = close
        HelpTopicId = topic.Id
        ' The topic in the help file (F1 does the same)
        If HelpFile.Available Then
            Dim read = AddFooterButton("help.readTopic", If(topic.HasVideo, ButtonKind.Secondary, ButtonKind.Primary), Glyphs.Help)
            AddHandler read.Click, Sub() HelpFile.Show(Me, _topic.Id)
            If Not topic.HasVideo Then AcceptButton = read
        End If
        If topic.HasVideo Then
            Dim watch = AddFooterButton("help.watch", ButtonKind.Primary, Glyphs.Video)
            AddHandler watch.Click, Sub() AppHost.OpenUrl(_topic.Url)
            AcceptButton = watch
        End If
    End Sub
End Class

''' <summary>
''' Admin > Help topics and videos. Works on a copy of the arrangement; Save replaces it (HelpCatalog.Save).
'''   left   the groups and topics as they appear in the Help menu, with New group / New topic /
'''          New sub-topic / Delete / Move up / Move down
'''   right  the selected item's title, video link and description
''' Built-in topics show their translated title and description unless something is typed here.
''' </summary>
Public NotInheritable Class HelpTopicsDialog
    Inherits DialogBase

    Private ReadOnly _root As List(Of HelpTopic) = HelpCatalog.CloneAll()
    Private ReadOnly tree As New TreeView With {.Dock = DockStyle.Fill, .BorderStyle = BorderStyle.FixedSingle, .HideSelection = False,
                                                 .FullRowSelect = True, .ShowLines = False, .Font = Fonts.Body}
    Private ReadOnly txtTitle As New TextField
    Private ReadOnly txtUrl As New TextField
    Private ReadOnly txtDesc As New TextBox With {.Multiline = True, .ScrollBars = ScrollBars.Vertical, .Font = Fonts.Body, .BorderStyle = BorderStyle.FixedSingle}
    Private ReadOnly lblTitleHint As New Label With {.AutoSize = True, .Font = Fonts.Small}
    Private ReadOnly lblDescHint As New Label With {.AutoSize = True, .Font = Fonts.Small}
    Private ReadOnly urlStack As Control
    Private ReadOnly descStack As Control
    Private ReadOnly btnOpen As ModernButton = ModernButton.Create("helpvideos.test", ButtonKind.Secondary, Glyphs.Video)
    Private ReadOnly btnSubTopic, btnDelete, btnUp, btnDown, btnTopic As ModernButton
    Private ReadOnly _tip As New ToolTip
    Private _loading As Boolean
    Private _dirty As Boolean
    Private _toldProject As Boolean

    Public Sub New()
        SetTitle("helptopics.title", "helptopics.subtitle")
        ClientSize = New Size(980, 700)
        LabelStyle.SetMuted(lblTitleHint)
        LabelStyle.SetMuted(lblDescHint)

        ' ----- Left: tree and its buttons -----
        Dim left As New Panel With {.Dock = DockStyle.Left, .Width = 400, .Padding = New Padding(0, 0, 16, 0)}
        Dim bar As New FlowLayoutPanel With {.Dock = DockStyle.Bottom, .AutoSize = True, .WrapContents = True, .Padding = New Padding(0, 8, 0, 0)}
        Dim btnGroup = ModernButton.Create("helptopics.newGroup", ButtonKind.Secondary, Glyphs.Add)
        btnTopic = ModernButton.Create("helptopics.newTopic", ButtonKind.Secondary, Glyphs.Add)
        btnSubTopic = ModernButton.Create("helptopics.newSubTopic", ButtonKind.Ghost, Glyphs.Add)
        btnDelete = ModernButton.CreateIcon(Glyphs.Delete, _tip, "helptopics.delete")
        btnUp = ModernButton.CreateIcon(Glyphs.ChevronUp, _tip, "custom.fav.up")
        btnDown = ModernButton.CreateIcon(Glyphs.ChevronDown, _tip, "custom.fav.down")
        For Each b In {btnGroup, btnTopic, btnSubTopic, btnDelete, btnUp, btnDown}
            b.Margin = New Padding(0, 0, 6, 6)
            bar.Controls.Add(b)
        Next
        AddHandler btnGroup.Click, Sub() AddGroup()
        AddHandler btnTopic.Click, Sub() AddTopic(asChild:=False)
        AddHandler btnSubTopic.Click, Sub() AddTopic(asChild:=True)
        AddHandler btnDelete.Click, Sub() DeleteSelected()
        AddHandler btnUp.Click, Sub() MoveSelected(-1)
        AddHandler btnDown.Click, Sub() MoveSelected(1)
        left.Controls.Add(tree)
        left.Controls.Add(bar)
        tree.ItemHeight = tree.Font.Height + 10
        AddHandler tree.AfterSelect, Sub() ShowSelected()

        ' ----- Right: the selected item -----
        Dim right = Ui.Column()
        right.Dock = DockStyle.Fill
        right.AutoSize = False
        txtTitle.Width = 500
        right.Controls.Add(Ui.Stack("helptopics.titleField", txtTitle))
        lblTitleHint.MaximumSize = New Size(500, 0)
        lblTitleHint.Margin = New Padding(0, 0, 0, 12)
        right.Controls.Add(lblTitleHint)
        Dim urlRow = Ui.Row()
        txtUrl.Width = 370
        urlRow.Controls.Add(Ui.Stack("helpvideos.col.url", txtUrl))
        btnOpen.Margin = New Padding(8, 22, 0, 0)
        urlRow.Controls.Add(btnOpen)
        urlStack = urlRow
        right.Controls.Add(urlRow)
        txtDesc.Size = New Size(500, 170)
        descStack = Ui.Stack("helptopics.descField", txtDesc)
        right.Controls.Add(descStack)
        lblDescHint.MaximumSize = New Size(500, 0)
        right.Controls.Add(lblDescHint)

        Body.Controls.Add(right)
        Body.Controls.Add(left)
        right.BringToFront()

        AddHandler txtTitle.Box.TextChanged, Sub() FieldChanged()
        AddHandler txtUrl.Box.TextChanged, Sub() FieldChanged()
        AddHandler txtDesc.TextChanged, Sub() FieldChanged()
        AddHandler btnOpen.Click, Sub() If txtUrl.Text.Trim() <> "" Then AppHost.OpenUrl(txtUrl.Text.Trim())

        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        ' Save keeps the editor open (the Help menu and Help screen update straight away); Close leaves
        Dim close = AddFooterButton("common.close", ButtonKind.Secondary)
        close.DialogResult = DialogResult.Cancel
        CancelButton = close
        AddHandler save.Click, Sub() SaveAll()

        FillTree()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _tip.Dispose()
        MyBase.Dispose(disposing)
    End Sub

#Region "Tree"

    Private ReadOnly Property Selected As HelpTopic
        Get
            Return TryCast(tree.SelectedNode?.Tag, HelpTopic)
        End Get
    End Property

    Private Sub FillTree(Optional selectTopic As HelpTopic = Nothing)
        tree.BeginUpdate()
        tree.Nodes.Clear()
        Dim found As TreeNode = Nothing
        For Each g In _root
            AddNode(tree.Nodes, g, selectTopic, found)
        Next
        tree.ExpandAll()
        tree.SelectedNode = If(found, If(tree.Nodes.Count > 0, tree.Nodes(0), Nothing))
        tree.SelectedNode?.EnsureVisible()
        tree.EndUpdate()
        ShowSelected()
    End Sub

    Private Sub AddNode(nodes As TreeNodeCollection, t As HelpTopic, selectTopic As HelpTopic, ByRef found As TreeNode)
        Dim n = nodes.Add(NodeText(t))
        n.Tag = t
        If t.IsGroup Then n.NodeFont = Fonts.Strong
        If t Is selectTopic Then found = n
        For Each c In t.Children
            AddNode(n.Nodes, c, selectTopic, found)
        Next
    End Sub

    Private Shared Function NodeText(t As HelpTopic) As String
        ' Groups are bold: trailing spaces stop the wider text being clipped
        Return t.Title & If(t.IsGroup, "    ", If(t.HasVideo, "", "  ·  " & Loc.T("helptopics.noLink")))
    End Function

    Private Function Siblings(t As HelpTopic) As List(Of HelpTopic)
        Return If(t.Parent Is Nothing, _root, t.Parent.Children)
    End Function

#End Region

#Region "Editing"

    Private Sub ShowSelected()
        Dim t = Selected
        _loading = True
        Dim has = t IsNot Nothing
        txtTitle.Enabled = has
        txtTitle.Text = If(has, t.CustomTitle, "")
        Dim builtInTitle = If(has, t.BuiltInTitle, "")
        lblTitleHint.Text = If(builtInTitle <> "", Loc.F("helptopics.builtInTitle", builtInTitle), Loc.T("helptopics.ownTitle"))
        urlStack.Visible = has AndAlso Not t.IsGroup
        descStack.Visible = has AndAlso Not t.IsGroup
        lblDescHint.Visible = has AndAlso Not t.IsGroup
        txtUrl.Text = If(has, t.VideoUrl, "")
        txtDesc.Text = If(has, t.CustomDescription, "")
        Dim builtInDesc = If(has, t.BuiltInDescription, "")
        lblDescHint.Text = If(builtInDesc <> "", Loc.T("helptopics.builtInDesc") & Environment.NewLine & builtInDesc, "")
        _loading = False
        btnTopic.Enabled = has
        btnSubTopic.Enabled = has AndAlso Not t.IsGroup
        btnDelete.Enabled = has
        btnUp.Enabled = has AndAlso Siblings(t).IndexOf(t) > 0
        btnDown.Enabled = has AndAlso Siblings(t).IndexOf(t) < Siblings(t).Count - 1
    End Sub

    Private Sub FieldChanged()
        If _loading Then Return
        Dim t = Selected
        If t Is Nothing Then Return
        t.CustomTitle = txtTitle.Text.Trim()
        If Not t.IsGroup Then
            t.VideoUrl = txtUrl.Text.Trim()
            t.CustomDescription = txtDesc.Text.Trim()
        End If
        tree.SelectedNode.Text = NodeText(t)
        _dirty = True
    End Sub

    ''' <summary>Closing with unsaved changes asks whether to save them.</summary>
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If _dirty AndAlso e.CloseReason = CloseReason.UserClosing Then
            Select Case Ask.YesNoCancel(Me, "helptopics.unsaved")
                Case DialogResult.Yes
                    If Not SaveAll() Then e.Cancel = True
                Case DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
        MyBase.OnFormClosing(e)
    End Sub

    Private Sub AddGroup()
        Dim g As New HelpTopic With {.Id = HelpCatalog.NewId(), .IsGroup = True, .CustomTitle = Loc.T("helptopics.newGroupName")}
        Dim t = Selected
        Dim top = t
        While top IsNot Nothing AndAlso top.Parent IsNot Nothing
            top = top.Parent
        End While
        Dim at = If(top Is Nothing, _root.Count, _root.IndexOf(top) + 1)
        _root.Insert(at, g)
        _dirty = True
        FillTree(g)
        txtTitle.Box.Focus()
        txtTitle.Box.SelectAll()
    End Sub

    ''' <summary>New topic: in the selected group, or after the selected topic; asChild = under the selected topic.</summary>
    Private Sub AddTopic(asChild As Boolean)
        Dim t = Selected
        If t Is Nothing Then Return
        Dim n As New HelpTopic With {.Id = HelpCatalog.NewId(), .CustomTitle = Loc.T("helptopics.newTopicName")}
        If t.IsGroup OrElse asChild Then
            n.Parent = t
            t.Children.Add(n)
        Else
            n.Parent = t.Parent
            Dim list = Siblings(t)
            list.Insert(list.IndexOf(t) + 1, n)
        End If
        _dirty = True
        FillTree(n)
        txtTitle.Box.Focus()
        txtTitle.Box.SelectAll()
    End Sub

    Private Sub DeleteSelected()
        Dim t = Selected
        If t Is Nothing Then Return
        Dim key = If(t.IsGroup OrElse t.Children.Count > 0, "helptopics.deleteGroupConfirm", "helptopics.deleteConfirm")
        If Not Ask.YesNoWarning(Me, key, t.Title) Then Return
        Dim list = Siblings(t)
        Dim i = list.IndexOf(t)
        list.RemoveAt(i)
        _dirty = True
        Dim nextOne = If(list.Count > 0, list(Math.Min(i, list.Count - 1)), t.Parent)
        FillTree(nextOne)
    End Sub

    Private Sub MoveSelected(delta As Integer)
        Dim t = Selected
        If t Is Nothing Then Return
        Dim list = Siblings(t)
        Dim i = list.IndexOf(t)
        Dim j = i + delta
        If j < 0 OrElse j >= list.Count Then Return
        list.RemoveAt(i)
        list.Insert(j, t)
        _dirty = True
        FillTree(t)
    End Sub

    ''' <summary>Checks and saves every topic. Returns False if something needs fixing first.</summary>
    Private Function SaveAll() As Boolean
        ' Check every topic before saving
        For Each t In Flatten(_root)
            If t.Title.Trim() = "" OrElse (Not t.IsBuiltIn AndAlso t.CustomTitle.Trim() = "") Then
                FillTree(t)
                Ask.Warn(Me, "helptopics.needTitle")
                Return False
            End If
            Dim u = t.Url
            If u <> "" AndAlso Not (u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse u.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) Then
                FillTree(t)
                Ask.Warn(Me, "helpvideos.invalid", t.Title)
                Return False
            End If
        Next
        Try
            Dim project = HelpCatalog.Save(_root)
            _dirty = False
            If project <> "" AndAlso Not _toldProject Then
                _toldProject = True
                Ask.Info(Me, "helpvideos.savedProject", project)
            End If
            Toast.Show(Me, Loc.T("helpvideos.saved"), ToastKind.Success, 2500)
            Return True
        Catch ex As Exception
            Log.Error("Save help topics", ex)
            Ask.Error(Me, "common.saveFailed", ex.Message)
            Return False
        End Try
    End Function

    Private Shared Iterator Function Flatten(items As IEnumerable(Of HelpTopic)) As IEnumerable(Of HelpTopic)
        For Each t In items
            Yield t
            For Each c In Flatten(t.Children)
                Yield c
            Next
        Next
    End Function

#End Region

End Class

''' <summary>About > Licence agreement: the end user licence agreement shown by the installer.</summary>
Public NotInheritable Class EulaDialog
    Inherits DialogBase

    Public Const ResourceName As String = "EULA.rtf"

    Public Sub New()
        SetTitle("eula.title")
        ClientSize = New Size(760, 640)
        Dim box As New RichTextBox With {.Dock = DockStyle.Fill, .ReadOnly = True, .BorderStyle = BorderStyle.None,
                                         .BackColor = Color.White, .ForeColor = Color.Black, .DetectUrls = False}
        Try
            Using s = Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
                If s IsNot Nothing Then
                    Using r As New IO.StreamReader(s)
                        box.Rtf = r.ReadToEnd()
                    End Using
                Else
                    box.Text = Loc.T("eula.missing")
                End If
            End Using
        Catch ex As Exception
            Log.Error("EULA", ex)
            box.Text = Loc.T("eula.missing")
        End Try
        ' White page whatever the theme, like the document it is
        ThemeManager.SetRole(box, SurfaceRole.Stimulus)
        Dim page As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(16), .BackColor = Color.White}
        ThemeManager.SetRole(page, SurfaceRole.Stimulus)
        page.Controls.Add(box)
        Body.Controls.Add(page)
        Dim close = AddFooterButton("common.close", ButtonKind.Primary)
        close.DialogResult = DialogResult.Cancel
        CancelButton = close
        AcceptButton = close
    End Sub
End Class
