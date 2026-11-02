Option Strict On

Imports System.Globalization
Imports System.Reflection
Imports System.Text.Json
Imports System.Text.Json.Nodes

' ====================================================================================================
'   SettingsJson.vb
'   Reads and writes the public fields of a settings class as readable JSON.
'   Supported field types: String, Boolean, Integer, Long, Single, Double, Decimal, Date, enums
'   (stored by name), Color ("#AARRGGBB") and Rectangle ("x,y,width,height").
'   Unknown or invalid values are skipped so an older or hand-edited file never stops start-up.
' ====================================================================================================

Public NotInheritable Class SettingsJson
    Private Sub New()
    End Sub

    Private Shared ReadOnly Inv As CultureInfo = CultureInfo.InvariantCulture

    Public Shared Function Write(obj As Object) As String
        Dim root As New JsonObject()
        For Each f In obj.GetType().GetFields(BindingFlags.Public Or BindingFlags.Instance)
            root(f.Name) = ToNode(f.GetValue(obj))
        Next
        Return root.ToJsonString(New JsonSerializerOptions With {.WriteIndented = True,
                                 .TypeInfoResolver = New Serialization.Metadata.DefaultJsonTypeInfoResolver(),
                                 .Encoder = Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping})
    End Function

    Public Shared Function Read(Of T As New)(json As String) As T
        Dim result As New T()
        Dim root = TryCast(JsonNode.Parse(json, Nothing, New JsonDocumentOptions With {.AllowTrailingCommas = True, .CommentHandling = JsonCommentHandling.Skip}), JsonObject)
        If root Is Nothing Then Return result
        Dim boxed As Object = result
        For Each f In GetType(T).GetFields(BindingFlags.Public Or BindingFlags.Instance)
            Dim node As JsonNode = Nothing
            If Not root.TryGetPropertyValue(f.Name, node) OrElse node Is Nothing Then Continue For
            Try
                Dim v = FromNode(node, f.FieldType)
                If v IsNot Nothing Then f.SetValue(boxed, v)
            Catch ex As Exception
                Log.Warn($"Settings value '{f.Name}' ignored: {ex.Message}")
            End Try
        Next
        Return DirectCast(boxed, T)
    End Function

    ''' <summary>Writes via a temporary file so a crash mid-write never leaves a half-written file.</summary>
    Public Shared Sub WriteFileSafely(path As String, text As String)
        IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(path))
        Dim tmp = path & ".tmp"
        IO.File.WriteAllText(tmp, text)
        If IO.File.Exists(path) Then
            IO.File.Replace(tmp, path, Nothing)
        Else
            IO.File.Move(tmp, path)
        End If
    End Sub

    ''' <summary>Renames an unreadable settings file to .bak so the program can start with defaults.</summary>
    Public Shared Sub SetAsideDamagedFile(path As String, ex As Exception)
        Log.Error("Damaged settings file " & path, ex)
        Try
            Dim bak = path & "." & Date.Now.ToString("yyyyMMdd-HHmmss", Inv) & ".bak"
            IO.File.Move(path, bak)
        Catch moveEx As Exception
            Log.Error("Could not set aside " & path, moveEx)
        End Try
    End Sub

    Private Shared Function ToNode(v As Object) As JsonNode
        If v Is Nothing Then Return Nothing
        Select Case True
            Case TypeOf v Is String : Return JsonValue.Create(DirectCast(v, String))
            Case TypeOf v Is Boolean : Return JsonValue.Create(DirectCast(v, Boolean))
            Case TypeOf v Is Integer : Return JsonValue.Create(DirectCast(v, Integer))
            Case TypeOf v Is Long : Return JsonValue.Create(DirectCast(v, Long))
            Case TypeOf v Is Single : Return JsonValue.Create(DirectCast(v, Single))
            Case TypeOf v Is Double : Return JsonValue.Create(DirectCast(v, Double))
            Case TypeOf v Is Decimal : Return JsonValue.Create(DirectCast(v, Decimal))
            Case TypeOf v Is Date : Return JsonValue.Create(DirectCast(v, Date).ToString("o", Inv))
            Case TypeOf v Is [Enum] : Return JsonValue.Create(v.ToString())
            Case TypeOf v Is Color : Return JsonValue.Create("#" & DirectCast(v, Color).ToArgb().ToString("X8", Inv))
            Case TypeOf v Is Rectangle
                Dim r = DirectCast(v, Rectangle)
                Return JsonValue.Create(String.Join(",", r.X, r.Y, r.Width, r.Height))
            Case Else : Return JsonValue.Create(Convert.ToString(v, Inv))
        End Select
    End Function

    Private Shared Function FromNode(node As JsonNode, type As Type) As Object
        Dim text = If(node.GetValueKind() = JsonValueKind.String, node.GetValue(Of String)(), node.ToJsonString())
        If type Is GetType(String) Then Return text
        If type Is GetType(Boolean) Then Return Boolean.Parse(text)
        If type Is GetType(Integer) Then Return CInt(Double.Parse(text, Inv))
        If type Is GetType(Long) Then Return CLng(Double.Parse(text, Inv))
        If type Is GetType(Single) Then Return Single.Parse(text, Inv)
        If type Is GetType(Double) Then Return Double.Parse(text, Inv)
        If type Is GetType(Decimal) Then Return Decimal.Parse(text, Inv)
        If type Is GetType(Date) Then Return Date.Parse(text, Inv, DateTimeStyles.RoundtripKind)
        If type.IsEnum Then Return [Enum].Parse(type, text, ignoreCase:=True)
        If type Is GetType(Color) Then
            If text.StartsWith("#") Then
                Dim argb = Convert.ToUInt32(text.Substring(1), 16)
                If text.Length = 7 Then argb = argb Or &HFF000000UI
                Return Color.FromArgb(BitConverter.ToInt32(BitConverter.GetBytes(argb), 0))
            End If
            Return Color.FromName(text)
        End If
        If type Is GetType(Rectangle) Then
            Dim p = text.Split(","c).Select(Function(x) Integer.Parse(x.Trim(), Inv)).ToArray()
            Return If(p.Length = 4, New Rectangle(p(0), p(1), p(2), p(3)), Rectangle.Empty)
        End If
        Return Nothing
    End Function
End Class
