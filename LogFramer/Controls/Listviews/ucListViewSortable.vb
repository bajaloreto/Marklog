Public MustInherit Class ListViewSortable
    Inherits ListView

    Private intSortColumnIndex As Integer = -1

    Public MustOverride Sub NewItem()
    Public MustOverride Sub EditItem()
    Public MustOverride Sub RemoveItem()
    Public MustOverride Sub CutItems()
    Public MustOverride Sub CopyItems()
    Public MustOverride Sub PasteItems(ByVal PasteItems As ClipboardItems)

    Public Property SortColumnIndex() As Integer
        Get
            Return intSortColumnIndex
        End Get
        Set(ByVal value As Integer)
            intSortColumnIndex = value
        End Set
    End Property

    Public ReadOnly Property SelectedGuid() As Guid
        Get
            If SelectedItems.Count > 0 Then
                Dim selItem As ListViewItem = SelectedItems(0)
                Try
                    Dim objGuid As New Guid(selItem.Name)
                    Return objGuid
                Catch ex As Exception
                    Return Nothing
                End Try
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedGuids() As Guid()
        Get
            If SelectedItems.Count > 0 Then
                Dim intCount As Integer = SelectedItems.Count - 1
                Dim objGuids(intCount) As Guid
                Dim selItem As ListViewItem

                For i = 0 To intCount
                    selItem = SelectedItems(i)
                    Try
                        Dim objGuid As New Guid(selItem.Name)

                        objGuids(i) = objGuid
                    Catch ex As Exception

                    End Try
                Next
                Return objGuids
            Else
                Return Nothing
            End If
        End Get
    End Property

    Protected Overrides Sub OnColumnClick(ByVal e As System.Windows.Forms.ColumnClickEventArgs)
        ' Determine whether the column is the same as the last column clicked.
        If e.Column <> intSortColumnIndex Then
            ' Set the sort column to the new column.
            intSortColumnIndex = e.Column
            ' Set the sort order to ascending by default.
            Me.Sorting = SortOrder.Ascending
        Else
            ' Determine what the last sort order was and change it.
            If Me.Sorting = SortOrder.Ascending Then
                Me.Sorting = SortOrder.Descending
            Else
                Me.Sorting = SortOrder.Ascending
            End If
        End If
        ' Call the sort method to manually sort.
        Me.Sort()
        ' Set the ListViewItemSorter property to a new ListViewItemComparer
        ' object.
        Me.ListViewItemSorter = New ListViewItemComparer(e.Column, Me.Sorting)

        MyBase.OnColumnClick(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)

        CurrentControl = Me
    End Sub
End Class

Public Class ListViewItemComparer
    Implements IComparer
    Private col As Integer
    Private order As SortOrder

    Public Sub New()
        col = 0
        order = SortOrder.Ascending
    End Sub

    Public Sub New(ByVal column As Integer, ByVal order As SortOrder)
        col = column
        Me.order = order
    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
                        Implements System.Collections.IComparer.Compare
        Dim returnVal As Integer = -1
        returnVal = [String].Compare(CType(x,  _
                        ListViewItem).SubItems(col).Text, _
                        CType(y, ListViewItem).SubItems(col).Text)
        ' Determine whether the sort order is descending.
        If order = SortOrder.Descending Then
            ' Invert the value returned by String.Compare.
            returnVal *= -1
        End If

        Return returnVal
    End Function
End Class

