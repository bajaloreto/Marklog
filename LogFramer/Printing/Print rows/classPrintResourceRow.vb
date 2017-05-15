Public Class PrintResourceRow
    Private intType As Integer
    Private intRowHeight As Integer
    Private strStructSortNumber As String
    Private objStruct As Struct
    Private intStructHeight As Integer
    Private intStructIndent As Integer
    Private strResourceSortNumber As String
    Private objResource As Resource
    Private intResourceHeight As Integer
    Private strBudgetItemSortNumber As String
    Private objBudgetItem As BudgetItem
    Private intBudgetItemHeight As Integer
    Private sngPercentage As Single
    Private curTotalCost As New Currency

    Public Enum Types
        Purpose = 0
        Output = 1
        Activity = 2
        Resource = 3
        BudgetItem = 4
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal inttype As Integer, ByVal structsortnumber As String, ByVal struct As Struct)
        Me.Type = inttype
        Me.StructSort = structsortnumber
        Me.Struct = struct
    End Sub

    Public Sub New(ByVal resourcesortnumber As String, ByVal resource As Resource)
        Me.Type = Types.Resource
        Me.strResourceSortNumber = resourcesortnumber
        Me.Resource = resource
    End Sub

    Public Sub New(ByVal budgetitemsortnumber As String, ByVal budgetitem As BudgetItem, ByVal percentage As Single, ByVal totalcost As Currency)
        Me.Type = Types.BudgetItem
        Me.BudgetItemSortNumber = budgetitemsortnumber
        Me.BudgetItem = budgetitem
        Me.Percentage = percentage
        Me.TotalCost = totalcost
    End Sub

    Public Property RowHeight As Integer
        Get
            Return intRowHeight
        End Get
        Set(ByVal value As Integer)
            intRowHeight = value
        End Set
    End Property

    Public Property Type As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    Public Property StructSort() As String
        Get
            Return strStructSortNumber
        End Get
        Set(ByVal value As String)
            strStructSortNumber = value
        End Set
    End Property

    Public Property Struct() As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
        End Set
    End Property

    Public Property StructHeight As Integer
        Get
            Return intStructHeight
        End Get
        Set(ByVal value As Integer)
            intStructHeight = value
        End Set
    End Property

    Public Property StructIndent As Integer
        Get
            Return intStructIndent
        End Get
        Set(ByVal value As Integer)
            intStructIndent = value
        End Set
    End Property

    Public Property ResourceSort() As String
        Get
            Return strResourceSortNumber
        End Get
        Set(ByVal value As String)
            strResourceSortNumber = value
        End Set
    End Property

    Public Property Resource() As Resource
        Get
            Return objResource
        End Get
        Set(ByVal value As Resource)
            objResource = value
        End Set
    End Property

    Public Property ResourceHeight As Integer
        Get
            Return intResourceHeight
        End Get
        Set(ByVal value As Integer)
            intResourceHeight = value
        End Set
    End Property

    Public Property BudgetItemSortNumber() As String
        Get
            Return strBudgetItemSortNumber
        End Get
        Set(ByVal value As String)
            strBudgetItemSortNumber = value
        End Set
    End Property

    Public Property BudgetItem As BudgetItem
        Get
            Return objBudgetItem
        End Get
        Set(ByVal value As BudgetItem)
            objBudgetItem = value
        End Set
    End Property

    Public Property BudgetItemHeight As Integer
        Get
            Return intBudgetItemHeight
        End Get
        Set(ByVal value As Integer)
            intBudgetItemHeight = value
        End Set
    End Property

    Public Property Percentage As Single
        Get
            Return sngPercentage
        End Get
        Set(ByVal value As Single)
            sngPercentage = value
        End Set
    End Property

    Public Property TotalCost() As Currency
        Get
            Return curTotalCost
        End Get
        Set(ByVal value As Currency)
            curTotalCost = value
        End Set
    End Property
End Class

Public Class PrintResourceRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintResourceRow)
        List.Add(row)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintResourceRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(row)
        Else
            List.Insert(index, row)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal row As PrintResourceRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("PrintResourceRow not in list!")
        Else
            List.Remove(row)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintResourceRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintResourceRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintResourceRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintResourceRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
