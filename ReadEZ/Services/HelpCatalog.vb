Option Strict On

Imports System.IO
Imports System.Reflection
Imports System.Text.Json
Imports System.Text.Json.Nodes

' ====================================================================================================
'   HelpCatalog.vb - the help topics, their order and their video links
'
'   Topics are arranged in groups (Help menu, Help screen and the "?" buttons all use this one list).
'   The built-in arrangement is:
'     Getting started   tour, installing, setting up the screens, the test environment
'     Screens           Patients, Tests, Results, Colours, Settings
'     Tests             step-by-step guide, symptoms, rate of reading, colour test, one or two screens
'   Built-in topics take their title and description from Strings.json (help.topic.<id>,
'   help.desc.<id>, help.group.<id>), so they follow the program's language. Topics added in
'   Admin have their own title and description (as typed), and any topic's title or description
'   can be replaced the same way.
'
'   Admin > Help topics and videos can add, delete and reorder topics and set each video link.
'   Where the arrangement comes from:
'     1. Assets\HelpVideos.json, built into the program (so a release carries it), then
'     2. %AppData%\ReadEZ\HelpVideos.json - the arrangement saved on this computer, if it
'        differs from 1.
'   Saving in Admin writes 2, and when the program runs from the Visual Studio solution it also
'   writes the project's Assets\HelpVideos.json, so the next release includes the changes.
'   Older files that only list video links ({"topic id": "url"}) are still read.
'
'   The "?" buttons ask for topics by id (screen.patients, test.fixation, analysis.hess ...). If
'   such a topic is deleted, its button is simply not shown.
' ====================================================================================================

Public NotInheritable Class HelpTopic
    Public Property Id As String = ""
    Public Property IsGroup As Boolean
    Public Property Parent As HelpTopic
    Public ReadOnly Property Children As New List(Of HelpTopic)
    ''' <summary>Other topics offered with this one (e.g. the questionnaire videos on the Questionnaires screen).</summary>
    Public Property Related As String() = {}
    ''' <summary>Title typed in Admin ("" = the built-in, translated title).</summary>
    Public Property CustomTitle As String = ""
    ''' <summary>Description typed in Admin ("" = the built-in, translated description).</summary>
    Public Property CustomDescription As String = ""
    ''' <summary>Web address of the video ("" = no video yet).</summary>
    Public Property VideoUrl As String = ""

    Public ReadOnly Property TitleKey As String
        Get
            Return If(IsGroup, "help.group." & Id, "help.topic." & Id)
        End Get
    End Property

    Public ReadOnly Property DescriptionKey As String
        Get
            Return "help.desc." & Id
        End Get
    End Property

    ''' <summary>True for topics that come with the program (they have translated text).</summary>
    Public ReadOnly Property IsBuiltIn As Boolean
        Get
            Return Loc.Has(TitleKey)
        End Get
    End Property

    Public ReadOnly Property BuiltInTitle As String
        Get
            Return If(Loc.Has(TitleKey), Loc.T(TitleKey), "")
        End Get
    End Property

    Public ReadOnly Property BuiltInDescription As String
        Get
            Return If(Loc.Has(DescriptionKey), Loc.T(DescriptionKey), "")
        End Get
    End Property

    Public ReadOnly Property Title As String
        Get
            If Not String.IsNullOrWhiteSpace(CustomTitle) Then Return CustomTitle
            Dim b = BuiltInTitle
            Return If(b <> "", b, Id)
        End Get
    End Property

    Public ReadOnly Property Description As String
        Get
            Return If(Not String.IsNullOrWhiteSpace(CustomDescription), CustomDescription, BuiltInDescription)
        End Get
    End Property

    Public ReadOnly Property Url As String
        Get
            Return If(VideoUrl, "").Trim()
        End Get
    End Property

    Public ReadOnly Property HasVideo As Boolean
        Get
            Return Not IsGroup AndAlso Url <> ""
        End Get
    End Property

    ''' <summary>Depth below the top level (groups are 0).</summary>
    Public ReadOnly Property Depth As Integer
        Get
            Dim d = 0, p = Parent
            While p IsNot Nothing
                d += 1 : p = p.Parent
            End While
            Return d
        End Get
    End Property

    ''' <summary>A separate copy of this topic and everything under it (the Admin editor works on a copy).</summary>
    Public Function Clone(Optional newParent As HelpTopic = Nothing) As HelpTopic
        Dim c As New HelpTopic With {.Id = Id, .IsGroup = IsGroup, .Parent = newParent, .Related = CType(Related.Clone(), String()),
                                     .CustomTitle = CustomTitle, .CustomDescription = CustomDescription, .VideoUrl = VideoUrl}
        For Each ch In Children
            c.Children.Add(ch.Clone(c))
        Next
        Return c
    End Function
End Class

Public NotInheritable Class HelpCatalog
    Private Sub New()
    End Sub

    Private Shared _root As List(Of HelpTopic)
    Private Shared ReadOnly _byId As New Dictionary(Of String, HelpTopic)(StringComparer.OrdinalIgnoreCase)
    Private Shared _builtInJson As String = ""

    Public Const ResourceName As String = "HelpVideos.json"

    ''' <summary>Raised after Admin saves the topics (menus and the Help screen rebuild).</summary>
    Public Shared Event TopicsChanged As EventHandler

#Region "Built-in arrangement"

    Private Shared Function BuildDefault() As List(Of HelpTopic)
        Dim root As New List(Of HelpTopic)
        Dim group = Function(id As String) As HelpTopic
                        Dim g As New HelpTopic With {.Id = id, .IsGroup = True}
                        root.Add(g)
                        Return g
                    End Function
        Dim topic = Function(parent As HelpTopic, id As String) As HelpTopic
                        Dim t As New HelpTopic With {.Id = id, .Parent = parent}
                        parent.Children.Add(t)
                        Return t
                    End Function

        Dim start = group("start")
        For Each id In {"start.tour", "start.install", "start.setup", "start.environment"}
            topic(start, id)
        Next

        Dim screens = group("screens")
        For Each id In {"screen.patients", "screen.tests", "screen.results", "screen.colours", "screen.settings"}
            topic(screens, id)
        Next

        Dim tests = group("tests")
        For Each id In {"tests.guide", "test.symptoms", "test.reading", "test.colour", "tests.run"}
            topic(tests, id)
        Next
        Return root
    End Function

#End Region

#Region "Topics"

    Private Shared Sub EnsureLoaded()
        If _root IsNot Nothing Then Return
        Dim root = BuildDefault()
        ' 1. Built into the program
        Try
            Using s = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
                If s IsNot Nothing Then root = ReadInto(root, s)
            End Using
        Catch ex As Exception
            Log.Error("Help topics (built in)", ex)
        End Try
        _builtInJson = ToJson(root)
        ' 2. Saved on this computer
        Try
            If File.Exists(OverrideFile) Then
                Using s = File.OpenRead(OverrideFile)
                    root = ReadInto(root, s)
                End Using
            End If
        Catch ex As Exception
            Log.Error("Help topics " & OverrideFile, ex)
        End Try
        SetRoot(root)
    End Sub

    Private Shared Sub SetRoot(root As List(Of HelpTopic))
        _root = root
        _byId.Clear()
        For Each t In Walk(root)
            _byId(t.Id) = t
        Next
    End Sub

    Private Shared Iterator Function Walk(items As IEnumerable(Of HelpTopic)) As IEnumerable(Of HelpTopic)
        For Each t In items
            Yield t
            For Each c In Walk(t.Children)
                Yield c
            Next
        Next
    End Function

    ''' <summary>The groups, in menu order.</summary>
    Public Shared ReadOnly Property Groups As IReadOnlyList(Of HelpTopic)
        Get
            EnsureLoaded()
            Return _root
        End Get
    End Property

    Public Shared Function Find(id As String) As HelpTopic
        If String.IsNullOrEmpty(id) Then Return Nothing
        EnsureLoaded()
        Dim t As HelpTopic = Nothing
        Return If(_byId.TryGetValue(id, t), t, Nothing)
    End Function

    ''' <summary>Every topic (not the groups) in menu order.</summary>
    Public Shared Function AllTopics() As List(Of HelpTopic)
        EnsureLoaded()
        Return Walk(_root).Where(Function(t) Not t.IsGroup).ToList()
    End Function

    ''' <summary>A copy of the whole arrangement for editing (Admin).</summary>
    Public Shared Function CloneAll() As List(Of HelpTopic)
        EnsureLoaded()
        Return _root.Select(Function(g) g.Clone()).ToList()
    End Function

    Public Shared Function Url(id As String) As String
        Dim t = Find(id)
        Return If(t Is Nothing, "", t.Url)
    End Function

    ''' <summary>The topic for an analysis window, from its title key (analysis.cover.titleNear -> analysis.cover).</summary>
    Public Shared Function TopicForAnalysis(titleKey As String) As String
        Dim id = If(titleKey, "")
        Dim cut = id.LastIndexOf(".title", StringComparison.Ordinal)
        If cut > 0 Then id = id.Substring(0, cut)
        Return If(Find(id) IsNot Nothing, id, "")
    End Function

    ''' <summary>A new, unused id for a topic added in Admin.</summary>
    Public Shared Function NewId() As String
        Return "user." & Guid.NewGuid().ToString("N").Substring(0, 10)
    End Function

#End Region

#Region "Reading and writing"

    Private Shared ReadOnly Property OverrideFile As String
        Get
            Return Path.Combine(AppSettings.SettingsFolder, ResourceName)
        End Get
    End Property

    ''' <summary>
    ''' A file with a "topics" list replaces the arrangement; an older file of id -> link pairs only
    ''' sets the links of the topics it names.
    ''' </summary>
    Private Shared Function ReadInto(root As List(Of HelpTopic), s As Stream) As List(Of HelpTopic)
        Dim node = JsonNode.Parse(s, documentOptions:=New JsonDocumentOptions With {.AllowTrailingCommas = True, .CommentHandling = JsonCommentHandling.Skip})
        Dim obj = TryCast(node, JsonObject)
        If obj Is Nothing Then Return root
        Dim topics = TryCast(obj("topics"), JsonArray)
        If topics IsNot Nothing Then
            Dim result As New List(Of HelpTopic)
            For Each n In topics
                Dim t = ReadTopic(TryCast(n, JsonObject), Nothing)
                If t IsNot Nothing Then result.Add(t)
            Next
            Return If(result.Count > 0, result, root)
        End If
        ' Older format: { "topic id": "link", ... }
        Dim byId = Walk(root).ToDictionary(Function(t) t.Id, StringComparer.OrdinalIgnoreCase)
        For Each kv In obj
            Dim t As HelpTopic = Nothing
            If kv.Value IsNot Nothing AndAlso byId.TryGetValue(kv.Key, t) Then t.VideoUrl = kv.Value.ToString()
        Next
        Return root
    End Function

    Private Shared Function ReadTopic(o As JsonObject, parent As HelpTopic) As HelpTopic
        If o Is Nothing Then Return Nothing
        Dim id = Str(o, "id")
        If id = "" Then Return Nothing
        Dim t As New HelpTopic With {.Id = id, .Parent = parent, .IsGroup = parent Is Nothing,
                                     .CustomTitle = Str(o, "title"), .CustomDescription = Str(o, "description"), .VideoUrl = Str(o, "url")}
        Dim rel = TryCast(o("related"), JsonArray)
        If rel IsNot Nothing Then t.Related = rel.Select(Function(x) If(x?.ToString(), "")).Where(Function(x) x <> "").ToArray()
        Dim kids = TryCast(o("children"), JsonArray)
        If kids IsNot Nothing Then
            For Each k In kids
                Dim c = ReadTopic(TryCast(k, JsonObject), t)
                If c IsNot Nothing Then t.Children.Add(c)
            Next
        End If
        Return t
    End Function

    Private Shared Function Str(o As JsonObject, name As String) As String
        Dim v = o(name)
        Return If(v Is Nothing, "", v.ToString())
    End Function

    Private Shared Function ToJson(root As IEnumerable(Of HelpTopic)) As String
        Dim arr As New JsonArray()
        For Each g In root
            arr.Add(TopicNode(g))
        Next
        Dim doc As New JsonObject From {{"version", 2}, {"topics", arr}}
        ' A resolver must be given explicitly: without it .NET refuses the options (InvalidOperationException)
        Return doc.ToJsonString(New JsonSerializerOptions With {.WriteIndented = True,
                                .TypeInfoResolver = New Serialization.Metadata.DefaultJsonTypeInfoResolver(),
                                .Encoder = Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping})
    End Function

    Private Shared Function TopicNode(t As HelpTopic) As JsonObject
        Dim o As New JsonObject From {{"id", t.Id}}
        If t.CustomTitle.Trim() <> "" Then o("title") = t.CustomTitle.Trim()
        If t.CustomDescription.Trim() <> "" Then o("description") = t.CustomDescription.Trim()
        If Not t.IsGroup Then o("url") = t.Url
        If t.Related.Length > 0 Then
            Dim rel As New JsonArray()
            For Each r In t.Related
                rel.Add(r)
            Next
            o("related") = rel
        End If
        If t.Children.Count > 0 Then
            Dim kids As New JsonArray()
            For Each c In t.Children
                kids.Add(TopicNode(c))
            Next
            o("children") = kids
        End If
        Return o
    End Function

    ''' <summary>
    ''' Saves an edited arrangement (from CloneAll). Returns the project file also written (a Visual
    ''' Studio build), or "".
    ''' </summary>
    Public Shared Function Save(root As List(Of HelpTopic)) As String
        EnsureLoaded()
        Dim json = ToJson(root)
        Directory.CreateDirectory(AppSettings.SettingsFolder)
        ' This computer: only kept when it differs from what is built into the program
        If json = _builtInJson Then
            If File.Exists(OverrideFile) Then File.Delete(OverrideFile)
        Else
            File.WriteAllText(OverrideFile, json)
        End If
        ' Use a copy straight away (the editor stays open and goes on changing its own), with the
        ' parents fixed up in case topics were moved
        root = root.Select(Function(g) g.Clone()).ToList()
        For Each g In root
            g.Parent = Nothing
            g.IsGroup = True
            FixParents(g)
        Next
        SetRoot(root)
        ' Visual Studio build: also the project's copy, which is built into the next release
        Dim project = ProjectFile()
        If project <> "" Then
            Try
                File.WriteAllText(project, json)
            Catch ex As Exception
                Log.Error("Help topics " & project, ex)
                project = ""
            End Try
        End If
        Log.Info($"Help topics saved ({AllTopics().Count} topics)")
        RaiseEvent TopicsChanged(Nothing, EventArgs.Empty)
        Return project
    End Function

    Private Shared Sub FixParents(t As HelpTopic)
        For Each c In t.Children
            c.Parent = t
            c.IsGroup = False
            FixParents(c)
        Next
    End Sub

    ''' <summary>ReadEZ\Assets\HelpVideos.json when running from the solution, else "".</summary>
    Private Shared Function ProjectFile() As String
        Dim d = New DirectoryInfo(AppHost.AppFolder)
        While d IsNot Nothing
            If File.Exists(Path.Combine(d.FullName, "ReadEZ.slnx")) OrElse File.Exists(Path.Combine(d.FullName, "ReadEZ.sln")) Then
                Dim assets = Path.Combine(d.FullName, "ReadEZ", "Assets")
                Return If(Directory.Exists(assets), Path.Combine(assets, ResourceName), "")
            End If
            d = d.Parent
        End While
        Return ""
    End Function

#End Region

#Region "Showing help"

    ''' <summary>"?" buttons: the topic's description, its video and related videos.</summary>
    Public Shared Sub ShowTopic(owner As IWin32Window, id As String)
        Dim t = Find(id)
        If t Is Nothing Then Return
        Using dlg As New HelpTopicDialog(t)
            dlg.ShowDialog(owner)
        End Using
    End Sub

    ''' <summary>Help menu: plays the video, or explains there isn't one yet.</summary>
    Public Shared Sub OpenVideo(owner As IWin32Window, id As String)
        Dim t = Find(id)
        If t Is Nothing Then Return
        If t.HasVideo Then AppHost.OpenUrl(t.Url) Else ShowTopic(owner, id)
    End Sub

#End Region

End Class
