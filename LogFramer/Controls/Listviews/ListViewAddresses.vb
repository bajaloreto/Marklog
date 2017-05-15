Public Class ListViewAddresses
    Inherits ListViewSortable

    Private SearchType As Integer
    Private objAddresses As Addresses
    Private CopyRowIndex As Integer = -1

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Type, 100, HorizontalAlignment.Left)
            .Columns.Add(LANG_Address, 400, HorizontalAlignment.Left)
            .AllowDrop = True
        End With
    End Sub

    Public Property Addresses() As Addresses
        Get
            Return objAddresses
        End Get
        Set(ByVal value As Addresses)
            objAddresses = value
            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedAddress() As Address
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selAddress As Address = Me.Addresses.GetAddressByGuid(Me.SelectedGuid)
                Return selAddress
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedAddresses() As Address()
        Get
            Dim objAddresses(SelectedItems.Count - 1) As Address

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objAddresses(i) = Me.Addresses.GetAddressByGuid(Me.SelectedGuids(i))
                Next

                Return objAddresses
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub LoadItems()
        Items.Clear()
        For Each selAddress As Address In Me.Addresses
            Dim newItem As New ListViewItem(selAddress.TypeName)
            newItem.Name = selAddress.Guid.ToString
            newItem.SubItems.Add(selAddress.FullAddress)
            Items.Add(newItem)
        Next
        If Items.Count > 0 Then
            Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        PopUpAddressDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.Addresses.Count > 0 AndAlso Me.SelectedAddress IsNot Nothing Then
            PopUpAddressDetails(Me.SelectedAddress)
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.Addresses.Count > 0 AndAlso Me.SelectedAddress IsNot Nothing Then
            If MsgBox(LANG_RemoveAddress, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selAddress As Address = Me.SelectedAddresses(0)
                UndoRedo.ItemRemoved(selAddress, Me.Addresses)

                Me.Addresses.Remove(selAddress)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpAddressDetails(ByVal selAddress As Address)
        Dim boolNew As Boolean

        If selAddress Is Nothing Then
            boolNew = True
            selAddress = New Address
        End If

        Dim dialogAddress As New DialogAddress(selAddress)

        dialogAddress.ShowDialog()
        If dialogAddress.DialogResult = vbOK Then
            If boolNew = True Then
                Me.Addresses.Add(selAddress)
                UndoRedo.ItemInserted(selAddress, Me.Addresses)
            End If

            Me.LoadItems()
        End If
        dialogAddress.Dispose()
        dialogAddress = Nothing
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selAddress As Address In SelectedAddresses
            UndoRedo.ItemCut(selAddress, Me.Addresses)

            Addresses.Remove(selAddress)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selAddress As Address In SelectedAddresses
            Dim NewItem As New ClipboardItem(CopyGroup, selAddress, Address.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selAddress As Address

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Address)
                    selAddress = CType(selItem.Item, Address)
                    Dim NewAddress As New Address

                    Using copier As New ObjectCopy
                        NewAddress = copier.CopyObject(selAddress)
                    End Using
                    Me.Addresses.Add(NewAddress)

                    UndoRedo.ItemPasted(NewAddress, Me.Addresses)
            End Select
        Next

        Me.LoadItems()
    End Sub

    Public Sub SendLetter(ByVal selContact As Contact, ByVal selOrganisation As Organisation)
        Dim selAddress As Address

        If SelectedAddress IsNot Nothing Then
            selAddress = SelectedAddress
        Else
            selAddress = Addresses.GetMainAddress
        End If

        If selAddress IsNot Nothing Then
            Using objWordIO As New WordIO
                objWordIO.SendLetterToAddress(selContact, selOrganisation, selAddress)
            End Using
        End If
    End Sub
End Class

