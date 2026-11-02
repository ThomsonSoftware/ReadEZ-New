Option Strict On

Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Text.Json

' ====================================================================================================
'   Loc.vb - Localisation
'
'   All on-screen text comes from Localisation\Strings.json (embedded in the program). Each entry
'   has an English, Spanish, German and Italian value:
'
'       "patients.title": { "en": "Patients", "es": "Pacientes", "de": "Patienten", "it": "Pazienti" }
'
'   Code never puts literal text on a control. Instead it binds the control to a key:
'
'       Loc.Bind(lblTitle, "patients.title")
'
'   When the language changes, every bound control in every open window is updated at once.
'   Text that is built from data (e.g. "3 recordings") uses Loc.F("key", value) inside a
'   refresh routine registered with Loc.OnChange(owner, AddressOf RefreshTexts).
'
'   Missing translations fall back to English, then to the key itself, and are listed in
'   %AppData%\ReadEZ\missing-translations.txt in debug builds.
' ====================================================================================================

Public Enum AppLanguage
    English = 0
    Spanish = 1
    German = 2
    Italian = 3
End Enum

Public NotInheritable Class Loc

    Private Sub New()
    End Sub

    Private Shared _strings As Dictionary(Of String, String()) = Nothing
    Private Shared _language As AppLanguage = AppLanguage.English
    Private Shared ReadOnly _bindings As New List(Of Binding)
    Private Shared ReadOnly _missing As New HashSet(Of String)

    ''' <summary>Raised after the language has changed and bound controls have been updated.</summary>
    Public Shared Event LanguageChanged As EventHandler

#Region "Language"

    Public Shared ReadOnly Property Language As AppLanguage
        Get
            Return _language
        End Get
    End Property

    Public Shared ReadOnly Property Culture As CultureInfo
        Get
            Return CultureFor(_language)
        End Get
    End Property

    Public Shared Function CultureFor(lang As AppLanguage) As CultureInfo
        Select Case lang
            Case AppLanguage.Spanish : Return New CultureInfo("es-ES")
            Case AppLanguage.German : Return New CultureInfo("de-DE")
            Case AppLanguage.Italian : Return New CultureInfo("it-IT")
            Case Else : Return New CultureInfo("en-GB")
        End Select
    End Function

    ''' <summary>Name of a language in its own language (for the language picker).</summary>
    Public Shared Function NativeName(lang As AppLanguage) As String
        Select Case lang
            Case AppLanguage.Spanish : Return "Español"
            Case AppLanguage.German : Return "Deutsch"
            Case AppLanguage.Italian : Return "Italiano"
            Case Else : Return "English"
        End Select
    End Function

    ''' <summary>Reads the language names stored by the previous version ("English", "Spanish", ...).</summary>
    Public Shared Function Parse(name As String) As AppLanguage
        Select Case If(name, "").Trim().ToUpperInvariant()
            Case "SPANISH", "ES", "ESPAÑOL", "ES-ES" : Return AppLanguage.Spanish
            Case "GERMAN", "DE", "DEUTSCH", "DE-DE" : Return AppLanguage.German
            Case "ITALIAN", "IT", "ITALIANO", "IT-IT" : Return AppLanguage.Italian
            Case Else : Return AppLanguage.English
        End Select
    End Function

    Public Shared Sub SetLanguage(lang As AppLanguage)
        _language = lang
        Dim ci = CultureFor(lang)
        Threading.Thread.CurrentThread.CurrentUICulture = ci
        Threading.Thread.CurrentThread.CurrentCulture = ci
        CultureInfo.DefaultThreadCurrentUICulture = ci
        CultureInfo.DefaultThreadCurrentCulture = ci
        RefreshAll()
        RaiseEvent LanguageChanged(Nothing, EventArgs.Empty)
    End Sub

#End Region

#Region "Lookup"

    ''' <summary>Translated text for a key.</summary>
    Public Shared Function T(key As String) As String
        EnsureLoaded()
        Dim values As String() = Nothing
        If _strings.TryGetValue(key, values) Then
            Dim v = values(CInt(_language))
            If Not String.IsNullOrEmpty(v) Then Return v
            NoteMissing(key)
            If Not String.IsNullOrEmpty(values(0)) Then Return values(0)
        Else
            NoteMissing(key)
        End If
        Return key
    End Function

    ''' <summary>Translated text with {0}, {1} ... placeholders filled in.</summary>
    Public Shared Function F(key As String, ParamArray args As Object()) As String
        Try
            Return String.Format(Culture, T(key), args)
        Catch ex As FormatException
            Return T(key)
        End Try
    End Function

    Public Shared Function Has(key As String) As Boolean
        EnsureLoaded()
        Return _strings.ContainsKey(key)
    End Function

    ''' <summary>All keys (used by the translation check in Admin).</summary>
    Public Shared Function AllKeys() As IEnumerable(Of String)
        EnsureLoaded()
        Return _strings.Keys
    End Function

    ''' <summary>Keys that have no text for the given language.</summary>
    Public Shared Function MissingFor(lang As AppLanguage) As List(Of String)
        EnsureLoaded()
        Return _strings.Where(Function(kv) String.IsNullOrEmpty(kv.Value(CInt(lang)))).Select(Function(kv) kv.Key).OrderBy(Function(k) k).ToList()
    End Function

    Private Shared Sub NoteMissing(key As String)
#If DEBUG Then
        SyncLock _missing
            If _missing.Add(key & " [" & _language.ToString() & "]") Then
                Try
                    Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ReadEZ")
                    Directory.CreateDirectory(folder)
                    File.AppendAllText(Path.Combine(folder, "missing-translations.txt"), key & " [" & _language.ToString() & "]" & Environment.NewLine)
                Catch
                End Try
            End If
        End SyncLock
#End If
    End Sub

    Private Shared Sub EnsureLoaded()
        If _strings IsNot Nothing Then Return
        Dim table As New Dictionary(Of String, String())(StringComparer.Ordinal)
        Dim asm = Assembly.GetExecutingAssembly()
        For Each resName In asm.GetManifestResourceNames().Where(Function(n) n.EndsWith("Strings.json", StringComparison.OrdinalIgnoreCase))
            Using s = asm.GetManifestResourceStream(resName)
                If s IsNot Nothing Then LoadInto(table, s)
            End Using
        Next
        ' A Strings.json next to the program (if present) overrides the built-in text.
        ' This lets translators correct wording without rebuilding.
        Dim external = Path.Combine(AppContext.BaseDirectory, "Strings.json")
        If File.Exists(external) Then
            Try
                Using s = File.OpenRead(external)
                    LoadInto(table, s)
                End Using
            Catch
            End Try
        End If
        _strings = table
    End Sub

    Private Shared Sub LoadInto(table As Dictionary(Of String, String()), s As Stream)
        Using doc = JsonDocument.Parse(s, New JsonDocumentOptions With {.CommentHandling = JsonCommentHandling.Skip, .AllowTrailingCommas = True})
            For Each entry In doc.RootElement.EnumerateObject()
                If entry.Value.ValueKind <> JsonValueKind.Object Then Continue For
                Dim values(3) As String
                Dim el As JsonElement
                If entry.Value.TryGetProperty("en", el) Then values(0) = el.GetString()
                If entry.Value.TryGetProperty("es", el) Then values(1) = el.GetString()
                If entry.Value.TryGetProperty("de", el) Then values(2) = el.GetString()
                If entry.Value.TryGetProperty("it", el) Then values(3) = el.GetString()
                table(entry.Name) = values
            Next
        End Using
    End Sub

#End Region

#Region "Binding controls to keys"

    Private NotInheritable Class Binding
        Public Target As WeakReference
        Public Owner As WeakReference
        Public Apply As Action(Of Object)
    End Class

    ''' <summary>Sets a control's Text from a key now and whenever the language changes.</summary>
    Public Shared Function Bind(Of TControl As Control)(c As TControl, key As String) As TControl
        Register(c, c, Sub(o) DirectCast(o, Control).Text = T(key))
        Return c
    End Function

    ''' <summary>Same as Bind but with placeholders.</summary>
    Public Shared Function BindF(Of TControl As Control)(c As TControl, key As String, ParamArray args As Object()) As TControl
        Register(c, c, Sub(o) DirectCast(o, Control).Text = F(key, args))
        Return c
    End Function

    Public Shared Function Bind(item As ToolStripItem, key As String) As ToolStripItem
        Register(item, item, Sub(o) DirectCast(o, ToolStripItem).Text = T(key))
        Return item
    End Function

    Public Shared Sub Bind(column As DataGridViewColumn, key As String)
        Register(column, column, Sub(o) DirectCast(o, DataGridViewColumn).HeaderText = T(key))
    End Sub

    Public Shared Sub BindTip(tip As ToolTip, c As Control, key As String)
        Register(c, c, Sub(o) tip.SetToolTip(DirectCast(o, Control), T(key)))
    End Sub

    ''' <summary>Fills a combo box with translated items, keeping the selected index on language change.</summary>
    Public Shared Sub BindItems(combo As ComboBox, ParamArray keys As String())
        Register(combo, combo,
                 Sub(o)
                     Dim cb = DirectCast(o, ComboBox)
                     Dim idx = cb.SelectedIndex
                     cb.BeginUpdate()
                     cb.Items.Clear()
                     For Each k In keys
                         cb.Items.Add(T(k))
                     Next
                     cb.EndUpdate()
                     If idx >= 0 AndAlso idx < cb.Items.Count Then cb.SelectedIndex = idx
                 End Sub)
    End Sub

    ''' <summary>
    ''' Runs an action now and on every language change, for as long as the owner control exists.
    ''' Use this for text that is built from data.
    ''' </summary>
    Public Shared Sub OnChange(owner As Control, refresh As Action)
        Register(owner, owner, Sub(o) refresh())
    End Sub

    Private Shared Sub Register(target As Object, owner As Object, apply As Action(Of Object))
        Dim b As New Binding With {.Target = New WeakReference(target), .Owner = New WeakReference(owner), .Apply = apply}
        SyncLock _bindings
            _bindings.Add(b)
        End SyncLock
        ' Drop the binding as soon as its control is disposed (the Apply closure may hold the
        ' control strongly, so waiting for garbage collection is not enough).
        Dim ctl = TryCast(owner, Control)
        If ctl IsNot Nothing Then
            AddHandler ctl.Disposed, Sub()
                                         SyncLock _bindings
                                             _bindings.Remove(b)
                                         End SyncLock
                                     End Sub
        Else
            Dim item = TryCast(owner, Component)
            If item IsNot Nothing Then
                AddHandler item.Disposed, Sub()
                                              SyncLock _bindings
                                                  _bindings.Remove(b)
                                              End SyncLock
                                          End Sub
            End If
        End If
        Try
            apply(target)
        Catch
        End Try
    End Sub

    ''' <summary>Re-applies every binding. Dead entries (disposed controls) are removed.</summary>
    Public Shared Sub RefreshAll()
        Dim snapshot As List(Of Binding)
        SyncLock _bindings
            _bindings.RemoveAll(Function(b) Not b.Target.IsAlive OrElse IsDisposed(b.Owner.Target))
            snapshot = _bindings.ToList()
        End SyncLock
        For Each b In snapshot
            Dim target = b.Target.Target
            If target Is Nothing Then Continue For
            Try
                b.Apply(target)
            Catch
            End Try
        Next
    End Sub

    Private Shared Function IsDisposed(o As Object) As Boolean
        If o Is Nothing Then Return True
        Dim c = TryCast(o, Control)
        If c IsNot Nothing Then Return c.IsDisposed
        Dim ti = TryCast(o, ToolStripItem)
        If ti IsNot Nothing Then Return ti.IsDisposed
        Return False
    End Function

#End Region

End Class
