Public Class BudgetResourceReferenceGridRow
    Private strOutputSortNumber, strActivitySortNumber, strResourceSortNumber As String
    Private strOutputText, strActivityText, strResourceText As String
    Private objOutputGuid, objActivityGuid, objResourceGuid As New Guid
    Private objBudgetItemReference As BudgetItemReference


    Public Sub New()

    End Sub

    Public Sub New(ByVal budgetitemreference As BudgetItemReference, _
                   ByVal resourceguid As Guid, ByVal resourcenumber As String, ByVal resourcetext As String, _
                   ByVal activityguid As Guid, ByVal activitynumber As String, ByVal activitytext As String, _
                   ByVal outputguid As Guid, ByVal outputnumber As String, ByVal outputtext As String)
        Me.OutputGuid = outputguid
        Me.OutputSortNumber = outputnumber
        Me.OutputText = outputtext
        Me.ActivityGuid = activityguid
        Me.ActivitySortNumber = activitynumber
        Me.ActivityText = activitytext
        Me.ResourceGuid = resourceguid
        Me.ResourceSortNumber = resourcenumber
        Me.ResourceText = resourcetext
    End Sub

    Public Property OutputGuid As Guid
        Get
            Return objOutputGuid
        End Get
        Set(value As Guid)
            objOutputGuid = value
        End Set
    End Property

    Public Property OutputSortNumber As String
        Get
            Return strOutputSortNumber
        End Get
        Set(value As String)
            strOutputSortNumber = value
        End Set
    End Property

    Public Property OutputText As String
        Get
            Return strOutputText
        End Get
        Set(value As String)
            strOutputText = value
        End Set
    End Property

    Public Property ActivityGuid As Guid
        Get
            Return objActivityGuid
        End Get
        Set(value As Guid)
            objActivityGuid = value
        End Set
    End Property

    Public Property ActivitySortNumber As String
        Get
            Return strActivitySortNumber
        End Get
        Set(value As String)
            strActivitySortNumber = value
        End Set
    End Property

    Public Property ActivityText As String
        Get
            Return strActivityText
        End Get
        Set(value As String)
            strActivityText = value
        End Set
    End Property

    Public Property ResourceGuid As Guid
        Get
            Return objResourceGuid
        End Get
        Set(value As Guid)
            objResourceGuid = value
        End Set
    End Property

    Public Property ResourceSortNumber As String
        Get
            Return strResourceSortNumber
        End Get
        Set(value As String)
            strResourceSortNumber = value
        End Set
    End Property

    Public Property ResourceText As String
        Get
            Return strResourceText
        End Get
        Set(value As String)
            strResourceText = value
        End Set
    End Property

    Public Property BudgetItemReference As BudgetItemReference
        Get
            Return objBudgetItemReference
        End Get
        Set(value As BudgetItemReference)
            objBudgetItemReference = value
        End Set
    End Property

    Public Property Percentage As Single
        Get
            Return objBudgetItemReference.Percentage
        End Get
        Set(value As Single)
            objBudgetItemReference.Percentage = value
        End Set
    End Property

    Public Property TotalCost As Currency
        Get
            Return objBudgetItemReference.TotalCost
        End Get
        Set(value As Currency)
            objBudgetItemReference.TotalCost = value
        End Set
    End Property
End Class

Public Class BudgetResourceReferenceGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As BudgetResourceReferenceGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As BudgetResourceReferenceGridRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(GridRow)
        Else
            List.Insert(index, GridRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal gridrow As BudgetResourceReferenceGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As BudgetResourceReferenceGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), BudgetResourceReferenceGridRow)
            End If
        End Get
        Set(value As BudgetResourceReferenceGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As BudgetResourceReferenceGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As BudgetResourceReferenceGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
End Class
