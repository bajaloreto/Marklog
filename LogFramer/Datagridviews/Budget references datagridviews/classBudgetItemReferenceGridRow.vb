Public Class BudgetItemReferenceGridRow
    Private objBudgetItemReference As BudgetItemReference
    Private strSortNumber As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal sortnumber As String, ByVal budgetitemreference As BudgetItemReference)
        Me.SortNumber = sortnumber
        Me.BudgetItemReference = budgetitemreference
    End Sub

#Region "Properties"
    Public Property BudgetItemReference As BudgetItemReference
        Get
            Return objBudgetItemReference
        End Get
        Set(value As BudgetItemReference)
            objBudgetItemReference = value
        End Set
    End Property

    Public Property SortNumber As String
        Get
            Return strSortNumber
        End Get
        Set(value As String)
            strSortNumber = value
        End Set
    End Property

    Public Property ReferenceBudgetItemGuid As Guid
        Get
            If objBudgetItemReference IsNot Nothing Then
                Return objBudgetItemReference.ReferenceBudgetItemGuid
            Else
                Return Guid.Empty
            End If
        End Get
        Set(ByVal value As Guid)
            If objBudgetItemReference IsNot Nothing Then
                objBudgetItemReference.ReferenceBudgetItemGuid = value
            End If
        End Set
    End Property

    Public Property Percentage As Single
        Get
            If objBudgetItemReference IsNot Nothing Then
                Return objBudgetItemReference.Percentage
            Else
                Return 0
            End If
        End Get
        Set(value As Single)
            If objBudgetItemReference IsNot Nothing Then
                objBudgetItemReference.Percentage = value
            End If
        End Set
    End Property

    Public Property TotalCost() As Currency
        Get
            If objBudgetItemReference IsNot Nothing Then
                Return objBudgetItemReference.TotalCost
            Else
                Return New Currency
            End If

        End Get
        Set(ByVal value As Currency)
            If objBudgetItemReference IsNot Nothing Then
                objBudgetItemReference.TotalCost = value
            End If
        End Set
    End Property
#End Region
End Class

Public Class BudgetItemReferenceGridRows
    Inherits System.Collections.CollectionBase

    Public Event BudgetItemReferenceGridRowAdded(ByVal sender As Object, ByVal e As BudgetItemReferenceGridRowAddedEventArgs)

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As BudgetItemReferenceGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), BudgetItemReferenceGridRow)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal budgetitemreference As BudgetItemReferenceGridRow)
        List.Add(budgetitemreference)

        RaiseEvent BudgetItemReferenceGridRowAdded(Me, New BudgetItemReferenceGridRowAddedEventArgs(budgetitemreference))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal budgetitemreference As BudgetItemReferenceGridRow)
        If intIndex > Count Or intIndex < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf intIndex = Count Then
            List.Add(budgetitemreference)

            RaiseEvent BudgetItemReferenceGridRowAdded(Me, New BudgetItemReferenceGridRowAddedEventArgs(budgetitemreference))
        Else
            List.Insert(intIndex, budgetitemreference)

            RaiseEvent BudgetItemReferenceGridRowAdded(Me, New BudgetItemReferenceGridRowAddedEventArgs(budgetitemreference))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal budgetitemreference As BudgetItemReferenceGridRow)
        If List.Contains(budgetitemreference) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(budgetitemreference)
        End If
    End Sub

    Public Function IndexOf(ByVal budgetitemreference As BudgetItemReferenceGridRow) As Integer
        Return List.IndexOf(budgetitemreference)
    End Function

    Public Function Contains(ByVal budgetitemreference As BudgetItemReferenceGridRow) As Boolean
        Return List.Contains(budgetitemreference)
    End Function
#End Region
End Class

Public Class BudgetItemReferenceGridRowAddedEventArgs
    Inherits EventArgs

    Public Property BudgetItemReferenceGridRow As BudgetItemReferenceGridRow

    Public Sub New(ByVal objBudgetItemReferenceGridRow As BudgetItemReferenceGridRow)
        MyBase.New()

        Me.BudgetItemReferenceGridRow = objBudgetItemReferenceGridRow
    End Sub
End Class
