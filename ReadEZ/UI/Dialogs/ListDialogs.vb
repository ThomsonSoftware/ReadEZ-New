Option Strict On

' ====================================================================================================
'   ListDialogs.vb - organisations and screeners (replace RZfrmOrganisations and RZfrmTesters)
' ====================================================================================================

''' <summary>A list with Add / Edit / Delete buttons (shared by the two dialogs below).</summary>
Public MustInherit Class EditableListDialog
    Inherits DialogBase

    Protected ReadOnly List As New ListBox With {.Dock = DockStyle.Fill, .IntegralHeight = False, .Font = Fonts.Body, .ItemHeight = 22, .BorderStyle = BorderStyle.FixedSingle}

    Protected Sub New(titleKey As String, subtitleKey As String)
        SetTitle(titleKey, subtitleKey)
        ClientSize = New Size(520, 520)
        Dim layout As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .BackColor = Color.Transparent}
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.Controls.Add(List, 0, 0)
        Dim buttons = Ui.Column()
        buttons.Margin = New Padding(10, 0, 0, 0)
        Dim add = ModernButton.Create("common.add", ButtonKind.Secondary, Glyphs.Add)
        Dim edit = ModernButton.Create("common.edit", ButtonKind.Secondary, Glyphs.Edit)
        Dim del = ModernButton.Create("common.delete", ButtonKind.Ghost, Glyphs.Delete)
        For Each b In {add, edit, del}
            b.Margin = New Padding(0, 0, 0, 8)
            buttons.Controls.Add(b)
        Next
        layout.Controls.Add(buttons, 1, 0)
        Body.Controls.Add(layout)
        AddHandler add.Click, Sub() EditItem(-1)
        AddHandler edit.Click, Sub() If List.SelectedIndex >= 0 Then EditItem(List.SelectedIndex)
        AddHandler List.DoubleClick, Sub() If List.SelectedIndex >= 0 Then EditItem(List.SelectedIndex)
        AddHandler del.Click, Sub() If List.SelectedIndex >= 0 Then DeleteItem(List.SelectedIndex)
    End Sub

    Protected MustOverride Sub Fill()
    Protected MustOverride Sub EditItem(index As Integer)
    Protected MustOverride Sub DeleteItem(index As Integer)
End Class

Public NotInheritable Class OrganisationsDialog
    Inherits EditableListDialog

    Private _items As New List(Of OrganisationStructure)

    Public Sub New()
        MyBase.New("orgs.title", "orgs.subtitle")
        Dim close = AddFooterButton("common.close", ButtonKind.Primary)
        close.DialogResult = DialogResult.OK
        AcceptButton = close
        CancelButton = close
        Fill()
    End Sub

    Protected Overrides Sub Fill()
        _items = AppHost.Db.GetOrganisations()
        List.Items.Clear()
        For Each o In _items
            List.Items.Add(o.Name & If(String.IsNullOrWhiteSpace(o.Town), "", "  (" & o.Town & ")"))
        Next
    End Sub

    Protected Overrides Sub EditItem(index As Integer)
        Dim o = If(index >= 0, _items(index), New OrganisationStructure)
        Using d As New OrganisationDialog(o)
            If d.ShowDialog(Me) <> DialogResult.OK Then Return
            AppHost.Db.SaveOrganisation(d.Value)
        End Using
        Fill()
    End Sub

    Protected Overrides Sub DeleteItem(index As Integer)
        If Not Ask.YesNoWarning(Me, "orgs.delete.confirm", _items(index).Name) Then Return
        AppHost.Db.DeleteOrganisation(_items(index).ID)
        Fill()
    End Sub
End Class

Public NotInheritable Class OrganisationDialog
    Inherits DialogBase

    Private ReadOnly _fields As New Dictionary(Of String, TextField)
    Private ReadOnly _id As Integer

    Public Property Value As OrganisationStructure

    Public Sub New(o As OrganisationStructure)
        _id = o.ID
        SetTitle(If(o.ID > 0, "orgs.edit", "orgs.add"))
        ClientSize = New Size(560, 470)
        Dim grid = NewGrid(2)
        Dim keys = {"Name", "Tel", "Address1", "Address2", "Town", "County", "Postcode"}
        Dim values = {o.Name, o.Tel, o.Address1, o.Address2, o.Town, o.County, o.Postcode}
        For i = 0 To keys.Length - 1
            Dim f As New TextField With {.Text = If(values(i), "")}
            _fields(keys(i)) = f
            grid.Controls.Add(Ui.Stack("orgs.field." & keys(i).ToLowerInvariant(), f, 240), i Mod 2, i \ 2)
        Next
        Body.Controls.Add(grid)
        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AcceptButton = save
        AddHandler save.Click, Sub()
                                   If _fields("Name").Text.Trim() = "" Then
                                       _fields("Name").Invalid = True
                                       Return
                                   End If
                                   Value = New OrganisationStructure With {.ID = _id, .Name = _fields("Name").Text.Trim(), .Tel = _fields("Tel").Text.Trim(),
                                       .Address1 = _fields("Address1").Text.Trim(), .Address2 = _fields("Address2").Text.Trim(),
                                       .Town = _fields("Town").Text.Trim(), .County = _fields("County").Text.Trim(), .Postcode = _fields("Postcode").Text.Trim()}
                                   DialogResult = DialogResult.OK
                               End Sub
    End Sub
End Class

''' <summary>The screeners (the people running the tests). picking = choose the one running tests now.</summary>
Public NotInheritable Class ScreenersDialog
    Inherits EditableListDialog

    Private _items As New List(Of ScreenerStructure)
    Private ReadOnly _picking As Boolean

    Public Property SelectedScreenerId As Integer

    Public Sub New(Optional picking As Boolean = False)
        MyBase.New("screeners.title", If(picking, "screeners.pickSubtitle", "screeners.subtitle"))
        _picking = picking
        If picking Then
            Dim ok = AddFooterButton("screeners.use", ButtonKind.Primary, Glyphs.Accept)
            Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
            cancel.DialogResult = DialogResult.Cancel
            CancelButton = cancel
            AcceptButton = ok
            AddHandler ok.Click, Sub()
                                     If List.SelectedIndex < 0 Then Return
                                     SelectedScreenerId = _items(List.SelectedIndex).ID
                                     DialogResult = DialogResult.OK
                                 End Sub
        Else
            Dim close = AddFooterButton("common.close", ButtonKind.Primary)
            close.DialogResult = DialogResult.Cancel
            CancelButton = close
        End If
        Fill()
    End Sub

    Protected Overrides Sub Fill()
        _items = AppHost.Db.GetScreeners()
        List.Items.Clear()
        For Each s In _items
            List.Items.Add(s.FullName)
        Next
        Dim i = _items.FindIndex(Function(x) x.ID = AppHost.Session.ScreenerId)
        If i >= 0 Then List.SelectedIndex = i
    End Sub

    Protected Overrides Sub EditItem(index As Integer)
        Dim s = If(index >= 0, _items(index), New ScreenerStructure)
        Dim t = Ask.Text(Me, "screeners.title", "screeners.titlePrompt", If(s.Title, ""))
        If t Is Nothing Then Return
        Dim f = Ask.Text(Me, "screeners.title", "screeners.firstPrompt", If(s.Firstname, ""))
        If f Is Nothing Then Return
        Dim n = Ask.Text(Me, "screeners.title", "screeners.surnamePrompt", If(s.Surname, ""))
        If n Is Nothing Then Return
        s.Title = t.Trim() : s.Firstname = f.Trim() : s.Surname = n.Trim()
        If s.FullName = "" Then Return
        Dim id = AppHost.Db.SaveScreener(s)
        Fill()
        Dim i = _items.FindIndex(Function(x) x.ID = id)
        If i >= 0 Then List.SelectedIndex = i
    End Sub

    Protected Overrides Sub DeleteItem(index As Integer)
        If Not Ask.YesNo(Me, "screeners.delete.confirm", _items(index).FullName) Then Return
        AppHost.Db.DeleteScreener(_items(index).ID)
        If AppHost.Session.ScreenerId = _items(index).ID Then AppHost.Session.SelectScreener(0)
        Fill()
    End Sub
End Class
