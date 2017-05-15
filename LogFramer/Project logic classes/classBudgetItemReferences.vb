Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class BudgetItemReference
    Inherits LogframeObject

    Private intIdBudgetItemReference As Integer = -1
    Private intIdParentResource As Integer
    Private sngPercentage As Single = 1
    Private curTotalCost As New Currency
    Private objParentReferenceGuid, objReferenceBudgetItemGuid As Guid

    Public Sub New()

    End Sub

    Public Sub New(ByVal referenceguid As Guid)
        Me.ReferenceBudgetItemGuid = referenceguid
    End Sub

#Region "Properties"
    Public Property idBudgetItemReference As Integer
        Get
            Return intIdBudgetItemReference
        End Get
        Set(ByVal value As Integer)
            intIdBudgetItemReference = value
        End Set
    End Property

    Public Property idParentResource As Integer
        Get
            Return intIdParentResource
        End Get
        Set(ByVal value As Integer)
            intIdParentResource = value
        End Set
    End Property

    Public Property ParentResourceGuid() As Guid
        Get
            Return objParentReferenceGuid
        End Get
        Set(ByVal value As Guid)
            objParentReferenceGuid = value
        End Set
    End Property

    Public Property ReferenceBudgetItemGuid As Guid
        Get
            Return objReferenceBudgetItemGuid
        End Get
        Set(ByVal value As Guid)
            objReferenceBudgetItemGuid = value
        End Set
    End Property

    Public Property Percentage As Single
        Get
            Return sngPercentage
        End Get
        Set(value As Single)
            sngPercentage = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property TotalCost() As Currency
        Get
            Return curTotalCost
        End Get
        Set(ByVal value As Currency)
            curTotalCost = value
        End Set
    End Property

    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_BudgetItemReference
        End Get
    End Property

    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_BudgetItemReferences
        End Get
    End Property
#End Region

#Region "Events"
    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        'no child collections
    End Sub
#End Region

End Class

Public Class BudgetItemReferences
    Inherits System.Collections.CollectionBase

    Public Event BudgetItemReferenceAdded(ByVal sender As Object, ByVal e As BudgetItemReferenceAddedEventArgs)

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As BudgetItemReference
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), BudgetItemReference)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal budgetitemreference As BudgetItemReference)
        List.Add(budgetitemreference)

        RaiseEvent BudgetItemReferenceAdded(Me, New BudgetItemReferenceAddedEventArgs(budgetitemreference))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal budgetitemreference As BudgetItemReference)
        If intIndex > Count Or intIndex < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, budgetitemreference.ItemName))
        ElseIf intIndex = Count Then
            List.Add(budgetitemreference)

            RaiseEvent BudgetItemReferenceAdded(Me, New BudgetItemReferenceAddedEventArgs(budgetitemreference))
        Else
            List.Insert(intIndex, budgetitemreference)

            RaiseEvent BudgetItemReferenceAdded(Me, New BudgetItemReferenceAddedEventArgs(budgetitemreference))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, BudgetItemReference.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal budgetitemreference As BudgetItemReference)
        If List.Contains(budgetitemreference) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, budgetitemreference.ItemName))
        Else
            List.Remove(budgetitemreference)
        End If
    End Sub

    Public Function IndexOf(ByVal budgetitemreference As BudgetItemReference) As Integer
        Return List.IndexOf(budgetitemreference)
    End Function

    Public Function Contains(ByVal budgetitemreference As BudgetItemReference) As Boolean
        Return List.Contains(budgetitemreference)
    End Function

    Public Function GetBudgetItemReferenceByGuid(ByVal objGuid As Guid) As BudgetItemReference
        Dim selBudgetItemReference As BudgetItemReference = Nothing
        For Each objBudgetItemReference As BudgetItemReference In Me.List
            If objBudgetItemReference.Guid = objGuid Then
                selBudgetItemReference = objBudgetItemReference
                Exit For
            End If
        Next
        Return selBudgetItemReference
    End Function

    Public Function GetTotalCost() As Currency
        Dim curTotalCost As New Currency()
        For Each selItem As BudgetItemReference In Me.List
            If selItem.TotalCost IsNot Nothing Then
                If curTotalCost.CurrencyCode <> selItem.TotalCost.CurrencyCode Then _
                    curTotalCost.CurrencyCode = selItem.TotalCost.CurrencyCode
                curTotalCost.Amount += selItem.TotalCost.Amount
            End If
        Next

        Return curTotalCost
    End Function
#End Region

End Class

Public Class BudgetItemReferenceAddedEventArgs
    Inherits EventArgs

    Public Property BudgetItemReference As BudgetItemReference

    Public Sub New(ByVal objBudgetItemReference As BudgetItemReference)
        MyBase.New()

        Me.BudgetItemReference = objBudgetItemReference
    End Sub
End Class
