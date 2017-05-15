Public Class classUndoList
    Inherits System.Collections.CollectionBase

    Public Event UndoItemAdded()
    Public Event UndoItemRemoved()

    Public Sub Add(ByVal listitem As UndoListItem)
        List.Add(listitem)

        RaiseEvent UndoItemAdded()
    End Sub

    Public Function Contains(ByVal listitem As UndoListItem) As Boolean
        Return List.Contains(listitem)
    End Function

    Public Function IndexOf(ByVal listitem As UndoListItem) As Integer
        Return List.IndexOf(listitem)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal listitem As UndoListItem)
        If index > Count Or index < 0 Then
            Dim strMsg As String = String.Format(ERR_IndexNotValidCannotInsert, {LANG_UndoListItem})
            System.Windows.Forms.MessageBox.Show(strMsg)
        ElseIf index = Count Then
            List.Add(listitem)
            RaiseEvent UndoItemAdded()
        Else
            List.Insert(index, listitem)
            RaiseEvent UndoItemAdded()
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As UndoListItem
        Get
            Return CType(List.Item(index), UndoListItem)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
            RaiseEvent UndoItemRemoved()
        End If
    End Sub

    Public Overloads Sub RemoveAt(ByVal index As Integer)
        MyBase.RemoveAt(index)
        RaiseEvent UndoItemRemoved()
    End Sub

    Public Sub Remove(ByVal listitem As UndoListItem)
        If List.Contains(listitem) Then List.Remove(listitem)
        RaiseEvent UndoItemRemoved()
    End Sub
End Class
