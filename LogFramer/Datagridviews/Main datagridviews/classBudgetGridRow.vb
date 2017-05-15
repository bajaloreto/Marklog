Public Class BudgetGridRow
    Private objBudgetItem As BudgetItem
    Private strSortNumber As String
    Private intIndent As Integer
    Private bmBudgetItemImage As Bitmap
    Private boolBudgetItemImageDirty As Boolean = True
    Private intBudgetItemHeight As Integer

    Public Sub New()

    End Sub

    Public Sub New(ByVal budgetitem As BudgetItem)
        Me.BudgetItem = budgetitem
    End Sub

    Public Property SortNumber() As String
        Get
            Return strSortNumber
        End Get
        Set(ByVal value As String)
            strSortNumber = value
        End Set
    End Property

    Public Property Indent As Integer
        Get
            Return intIndent
        End Get
        Set(value As Integer)
            intIndent = value
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

    Public Property BudgetItemImage As Bitmap
        Get
            Return bmBudgetItemImage
        End Get
        Set(ByVal value As Bitmap)
            bmBudgetItemImage = value
        End Set
    End Property

    Public Property BudgetItemImageDirty As Boolean
        Get
            Return boolBudgetItemImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolBudgetItemImageDirty = value
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

    Public Property RTF() As String
        Get
            Return objBudgetItem.RTF
        End Get
        Set(ByVal value As String)
            objBudgetItem.RTF = value
        End Set
    End Property

    Public Property Number() As Single
        Get
            Return objBudgetItem.Number
        End Get
        Set(ByVal value As Single)
            objBudgetItem.Number = value
        End Set
    End Property

    Public Property NumberUnit() As String
        Get
            Return objBudgetItem.NumberUnit
        End Get
        Set(ByVal value As String)
            objBudgetItem.NumberUnit = value
        End Set
    End Property

    Public ReadOnly Property NumberSet As Boolean
        Get
            If Number = 0 And String.IsNullOrEmpty(NumberUnit) Then Return False Else Return True
        End Get
    End Property

    Public Property Duration() As Single
        Get
            Return objBudgetItem.Duration
        End Get
        Set(ByVal value As Single)
            objBudgetItem.Duration = value
        End Set
    End Property

    Public Property DurationUnit() As String
        Get
            Return objBudgetItem.DurationUnit
        End Get
        Set(ByVal value As String)
            objBudgetItem.DurationUnit = value
        End Set
    End Property

    Public ReadOnly Property DurationSet As Boolean
        Get
            If Duration = 0 And String.IsNullOrEmpty(DurationUnit) Then Return False Else Return True
        End Get
    End Property

    Public Property UnitCost() As Currency
        Get
            Return objBudgetItem.UnitCost
        End Get
        Set(ByVal value As Currency)
            objBudgetItem.UnitCost = value
            'SetTotalCost()
        End Set
    End Property

    Public Property TotalLocalCost() As Currency
        Get
            Return objBudgetItem.TotalLocalCost
        End Get
        Set(ByVal value As Currency)
            objBudgetItem.TotalLocalCost = value
        End Set
    End Property

    Public Property TotalCost() As Currency
        Get
            Return objBudgetItem.TotalCost
        End Get
        Set(ByVal value As Currency)
            objBudgetItem.TotalCost = value
        End Set
    End Property

    Public Property Type() As Integer
        Get
            If objBudgetItem IsNot Nothing Then
                Return objBudgetItem.Type
            Else
                Return BudgetItem.BudgetItemTypes.Normal
            End If
        End Get
        Set(ByVal value As Integer)
            objBudgetItem.Type = value
        End Set
    End Property

    Public Property Formula() As String
        Get
            Return objBudgetItem.Formula
        End Get
        Set(ByVal value As String)
            objBudgetItem.Formula = value
        End Set
    End Property

    Public ReadOnly Property IsSubTotal As Boolean
        Get
            If Me.BudgetItem IsNot Nothing Then
                If Me.BudgetItem.BudgetItems.Count > 0 Then
                    Return True
                End If
            End If

            Return False
        End Get
    End Property

    'Public Sub SetTotalCost()
    '    Dim sngNumberValue As Single = 1
    '    Dim sngDurationValue As Single = 1
    '    Dim sngAmount As Single
    '    If Me.UnitCost IsNot Nothing AndAlso Me.UnitCost.Amount > 0 Then
    '        If NumberSet = True Then
    '            sngNumberValue = Me.Number
    '        End If
    '        If DurationSet = True Then
    '            sngDurationValue = Me.Duration
    '        End If
    '        sngAmount = Me.UnitCost.Amount * sngNumberValue * sngDurationValue
    '        objBudgetItem.TotalLocalCost = New Currency(sngAmount, UnitCost.CurrencyCode)

    '        If UnitCost.ExchangeRate <> 0 Then sngAmount /= UnitCost.ExchangeRate
    '        objBudgetItem.TotalCost = New Currency(sngAmount, CurrentLogFrame.CurrencyCode)
    '    End If
    'End Sub
End Class

Public Class BudgetGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As BudgetGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As BudgetGridRow)
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

    Public Sub Remove(ByVal gridrow As BudgetGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As BudgetGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), BudgetGridRow)
            End If
        End Get
        Set(value As BudgetGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As BudgetGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As BudgetGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function

    Public Function GetPreviousBudgetItem(ByVal intGridRowIndex As Integer) As BudgetItem
        intGridRowIndex -= 1
        Dim selBudgetItem As BudgetItem = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As BudgetGridRow = Me(i)
            If selGridRow.BudgetItem IsNot Nothing Then
                selBudgetItem = selGridRow.BudgetItem
                Exit For
            End If
        Next

        Return selBudgetItem
    End Function

    Public Function GetNextBudgetItem(ByVal intGridRowIndex As Integer) As BudgetItem
        intGridRowIndex += 1
        Dim selBudgetItem As BudgetItem = Nothing

        For i = intGridRowIndex To Me.Count - 1
            Dim selGridRow As BudgetGridRow = Me(i)
            If selGridRow.BudgetItem IsNot Nothing Then
                selBudgetItem = selGridRow.BudgetItem
                Exit For
            End If
        Next

        Return selBudgetItem
    End Function
End Class