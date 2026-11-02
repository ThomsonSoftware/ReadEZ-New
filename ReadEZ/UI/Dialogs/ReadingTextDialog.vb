Option Strict On

' ====================================================================================================
'   ReadingTextDialog.vb - add or edit a reading text (Settings > Test setup), one row per line
' ====================================================================================================

Public Class ReadingTextDialog
    Inherits DialogBase

    Private ReadOnly txtName As New TextField
    Private ReadOnly txtBody As New TextField(multiline:=True)
    Private ReadOnly _editing As Boolean

    Public ReadOnly Property TextName As String
        Get
            Return txtName.Text.Trim()
        End Get
    End Property

    Public ReadOnly Property TextBody As String
        Get
            Return txtBody.Text
        End Get
    End Property

    ''' <param name="name">Nothing for a new text.</param>
    Public Sub New(name As String, bodyText As String)
        _editing = name IsNot Nothing
        SetTitle(If(_editing, "custom.ror.editTitle", "custom.ror.newTitle"), "custom.ror.textHint")
        ClientSize = New Size(640, 560)
        Dim col = Ui.Column()
        col.Dock = DockStyle.Top
        txtName.Width = 400
        txtName.Text = If(name, "")
        txtName.Enabled = Not _editing
        col.Controls.Add(Ui.Stack("custom.ror.name", txtName))
        txtBody.Width = 580
        txtBody.Height = 300
        txtBody.Text = If(bodyText, "")
        txtBody.Box.WordWrap = False
        txtBody.Box.ScrollBars = ScrollBars.Both
        col.Controls.Add(Ui.Stack("custom.ror.lines", txtBody))
        Dim row = Ui.Row()
        Dim paste = ModernButton.Create("licence.paste", ButtonKind.Ghost, Glyphs.Copy)
        Dim clear = ModernButton.Create("custom.ror.clear", ButtonKind.Ghost, Glyphs.Delete)
        AddHandler paste.Click, Sub() If Clipboard.ContainsText() Then txtBody.Box.SelectedText = Clipboard.GetText()
        AddHandler clear.Click, Sub() txtBody.Text = ""
        row.Controls.Add(paste)
        row.Controls.Add(clear)
        col.Controls.Add(row)
        Body.Controls.Add(col)

        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AddHandler save.Click, Sub()
                                   If TextName = "" OrElse TextName.IndexOfAny(IO.Path.GetInvalidFileNameChars()) >= 0 Then
                                       txtName.Invalid = True
                                       Ask.Warn(Me, "custom.ror.badName")
                                       Return
                                   End If
                                   If Not _editing AndAlso IO.File.Exists(IO.Path.Combine(ReadingTexts.UserFolder, TextName & ".txt")) Then
                                       If Not Ask.YesNo(Me, "custom.ror.overwrite", TextName) Then Return
                                   End If
                                   DialogResult = DialogResult.OK
                               End Sub
    End Sub
End Class
