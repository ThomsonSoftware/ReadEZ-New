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
        ' Title, first name and surname in one window
        Using d As New ScreenerDialog(s)
            If d.ShowDialog(Me) <> DialogResult.OK Then Return
            s = d.Value
        End Using
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

''' <summary>A screener's title, first name and surname, side by side in one window.</summary>
Public NotInheritable Class ScreenerDialog
    Inherits DialogBase

    Private ReadOnly txtTitle As New TextField With {.Width = 90}
    Private ReadOnly txtFirst As New TextField With {.Width = 190}
    Private ReadOnly txtSurname As New TextField With {.Width = 220}

    Public Property Value As ScreenerStructure

    Public Sub New(s As ScreenerStructure)
        SetTitle(If(s.ID > 0, "screeners.edit", "screeners.add"))
        ClientSize = New Size(640, 250)
        Value = s
        txtTitle.Text = If(s.Title, "")
        txtFirst.Text = If(s.Firstname, "")
        txtSurname.Text = If(s.Surname, "")
        ' Three columns of fixed widths, inside the window's margins (no scrolling)
        Body.AutoScroll = False
        Dim grid = NewGrid(3)
        grid.ColumnStyles.Clear()
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110))
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 210))
        grid.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 240))
        grid.Padding = New Padding(0, 6, 0, 0)
        grid.Controls.Add(Ui.Stack("screeners.field.title", txtTitle), 0, 0)
        grid.Controls.Add(Ui.Stack("screeners.field.first", txtFirst), 1, 0)
        grid.Controls.Add(Ui.Stack("screeners.field.surname", txtSurname), 2, 0)
        Body.Controls.Add(grid)
        Dim save = AddFooterButton("common.save", ButtonKind.Primary, Glyphs.Save)
        Dim cancel = AddFooterButton("common.cancel", ButtonKind.Secondary)
        cancel.DialogResult = DialogResult.Cancel
        CancelButton = cancel
        AcceptButton = save
        AddHandler save.Click, Sub()
                                   ' A name is needed (a first name or a surname)
                                   If txtFirst.Text.Trim() = "" AndAlso txtSurname.Text.Trim() = "" Then
                                       txtSurname.Invalid = True
                                       Return
                                   End If
                                   Dim v = s
                                   v.Title = txtTitle.Text.Trim()
                                   v.Firstname = txtFirst.Text.Trim()
                                   v.Surname = txtSurname.Text.Trim()
                                   Value = v
                                   DialogResult = DialogResult.OK
                               End Sub
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        txtTitle.Focus()
    End Sub
End Class
