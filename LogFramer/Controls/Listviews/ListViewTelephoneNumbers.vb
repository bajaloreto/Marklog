Public Class ListViewTelephoneNumbers
    Inherits ListViewSortable

    Private objTelephoneNumbers As TelephoneNumbers

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Type, 100, HorizontalAlignment.Left)
            .Columns.Add(LANG_Number, 200, HorizontalAlignment.Left)
        End With
    End Sub

    Public Property TelephoneNumbers() As TelephoneNumbers
        Get
            Return objTelephoneNumbers
        End Get
        Set(ByVal value As TelephoneNumbers)
            objTelephoneNumbers = value
            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedTelephoneNumber() As TelephoneNumber
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selTelephoneNumber As TelephoneNumber = Me.TelephoneNumbers.GetTelephoneNumberByGuid(Me.SelectedGuid)
                Return selTelephoneNumber
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedTelephoneNumbers() As TelephoneNumber()
        Get
            Dim objTelephoneNumbers(SelectedItems.Count - 1) As TelephoneNumber

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objTelephoneNumbers(i) = Me.TelephoneNumbers.GetTelephoneNumberByGuid(Me.SelectedGuids(i))
                Next

                Return objTelephoneNumbers
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub LoadItems()
        Items.Clear()
        For Each selNumber As TelephoneNumber In Me.TelephoneNumbers
            Dim newItem As New ListViewItem(selNumber.TypeName)
            newItem.Name = selNumber.Guid.ToString
            newItem.SubItems.Add(selNumber.Number)
            Items.Add(newItem)
        Next
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
    End Sub

    Public Overrides Sub NewItem()
        PopUpTelephoneNumberDetails(Nothing)
    End Sub

    Public Overrides Sub EditItem()
        If Me.TelephoneNumbers.Count > 0 AndAlso Me.SelectedTelephoneNumber IsNot Nothing Then
            PopUpTelephoneNumberDetails(Me.SelectedTelephoneNumber)
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.TelephoneNumbers.Count > 0 AndAlso Me.SelectedTelephoneNumber IsNot Nothing Then
            If MsgBox(LANG_RemoveTelephoneNumber, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selTelephoneNumber As TelephoneNumber = Me.SelectedTelephoneNumbers(0)
                UndoRedo.ItemRemoved(selTelephoneNumber, Me.TelephoneNumbers)

                Me.TelephoneNumbers.Remove(selTelephoneNumber)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpTelephoneNumberDetails(ByVal selTelephoneNumber As TelephoneNumber)
        Dim boolNew As Boolean

        If selTelephoneNumber Is Nothing Then
            boolNew = True
            selTelephoneNumber = New TelephoneNumber
        End If

        Dim dialogTelephoneNumber As New DialogTelephoneNumber(selTelephoneNumber)

        dialogTelephoneNumber.ShowDialog()
        If dialogTelephoneNumber.DialogResult = vbOK Then
            If boolNew = True Then
                Me.TelephoneNumbers.Add(selTelephoneNumber)
                UndoRedo.ItemInserted(selTelephoneNumber, Me.TelephoneNumbers)
            End If

            Me.LoadItems()
        End If
        dialogTelephoneNumber.Dispose()
        dialogTelephoneNumber = Nothing
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selTelephoneNumber As TelephoneNumber In SelectedTelephoneNumbers
            UndoRedo.ItemCut(selTelephoneNumber, Me.TelephoneNumbers)

            TelephoneNumbers.Remove(selTelephoneNumber)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selTelephoneNumber As TelephoneNumber In SelectedTelephoneNumbers
            Dim NewItem As New ClipboardItem(CopyGroup, selTelephoneNumber, TelephoneNumber.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selTelephoneNumber As TelephoneNumber

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(TelephoneNumber)
                    selTelephoneNumber = CType(selItem.Item, TelephoneNumber)
                    Dim NewTelephoneNumber As New TelephoneNumber

                    Using copier As New ObjectCopy
                        NewTelephoneNumber = copier.CopyObject(selTelephoneNumber)
                    End Using
                    Me.TelephoneNumbers.Add(NewTelephoneNumber)

                    UndoRedo.ItemPasted(NewTelephoneNumber, Me.TelephoneNumbers)
            End Select
        Next

        Me.LoadItems()
    End Sub
End Class
