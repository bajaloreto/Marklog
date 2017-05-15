Public Class PrintBudgetRow
    Inherits BudgetGridRow

    Private intRowHeight As Integer
    Private strDurationText As String
    Private strNumberText As String
    Private strUnitCostText As String
    Private strTotalLocalCostText As String
    Private strTotalCostText As String
    Private intRowType As Integer
    Private objExchangeRate As ExchangeRate
    Private strCurrencyText As String
    Private strExchangeRateText As String
    Private strConversion1, strConversion2 As String
    Private strDefaultCurrencyCode As String

    Public Enum RowTypes
        Normal = 0
        TitleRow = 1
        TotalBudget = 2
        TotalBudgetRow = 3
        ExchangeRate = 4
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal budgetitem As BudgetItem)
        Me.BudgetItem = budgetitem

        SetText()
    End Sub

    Public Sub New(ByVal exchangerate As ExchangeRate, defaultcurrencycode As String)
        Me.ExchangeRate = exchangerate
        Me.DefaultCurrencyCode = defaultcurrencycode
        Me.RowType = RowTypes.ExchangeRate

        SetConversions()
    End Sub

    Private Sub SetText()
        If Me.BudgetItem IsNot Nothing Then
            If Me.DurationSet Then
                Me.DurationText = String.Format("{0} {1}", Me.Duration, Me.DurationUnit)
            End If
            If Me.IsSubTotal Then
                If Me.Number > 0 Then
                    If BudgetItem.BudgetItems.UniformNumberUnit = True Then
                        Me.NumberText = String.Format("{0} {1}", Me.Number, Me.NumberUnit)
                    End If
                End If
            Else
                If Me.NumberSet Then
                    Me.NumberText = String.Format("{0} {1}", Me.Number, Me.NumberUnit)
                End If
            End If

            Me.UnitCostText = Me.UnitCost.ToString

            If Me.IsSubTotal Then
                If Me.TotalLocalCost.Amount > 0 Then
                    If BudgetItem.BudgetItems.UniformCurrencyCode = True Then _
                        Me.TotalLocalCostText = Me.TotalLocalCost.ToString
                End If
            Else
                Me.TotalLocalCostText = Me.TotalLocalCost.ToString
            End If

            Me.TotalCostText = Me.TotalCost.ToString
        End If
    End Sub

    Private Sub SetConversions()
        Dim sngRate As Single
        Dim strConversion As String

        CurrencyText = ExchangeRate.CurrencyCode
        ExchangeRateText = ExchangeRate.ExchangeRate.ToString("N4")

        If ExchangeRate.ExchangeRate <> 0 Then
            sngRate = 1 / ExchangeRate.ExchangeRate
            strConversion = String.Format("1 {0} = {1} {2}", ExchangeRate.CurrencyCode, sngRate.ToString("N4"), DefaultCurrencyCode)
            Conversion1 = strConversion

            sngRate = ExchangeRate.ExchangeRate
            strConversion = String.Format("1 {0} = {1} {2}", DefaultCurrencyCode, sngRate.ToString("N4"), ExchangeRate.CurrencyCode)
            Conversion2 = strConversion
        End If
    End Sub

    Public Property RowHeight As Integer
        Get
            Return intRowHeight
        End Get
        Set(ByVal value As Integer)
            intRowHeight = value
        End Set
    End Property

    Public Property RowType As Integer
        Get
            Return intRowType
        End Get
        Set(ByVal value As Integer)
            intRowType = value
        End Set
    End Property

    Public Property DurationText As String
        Get
            Return strDurationText
        End Get
        Set(value As String)
            strDurationText = value
        End Set
    End Property

    Public Property NumberText As String
        Get
            Return strNumberText
        End Get
        Set(value As String)
            strNumberText = value
        End Set
    End Property

    Public Property UnitCostText As String
        Get
            Return strUnitCostText
        End Get
        Set(value As String)
            strUnitCostText = value
        End Set
    End Property

    Public Property TotalLocalCostText As String
        Get
            Return strTotalLocalCostText
        End Get
        Set(value As String)
            strTotalLocalCostText = value
        End Set
    End Property

    Public Property TotalCostText As String
        Get
            Return strTotalCostText
        End Get
        Set(value As String)
            strTotalCostText = value
        End Set
    End Property

    Public Property ExchangeRate As ExchangeRate
        Get
            Return objExchangeRate
        End Get
        Set(value As ExchangeRate)
            objExchangeRate = value
        End Set
    End Property

    Public Property CurrencyText As String
        Get
            Return strCurrencyText
        End Get
        Set(value As String)
            strCurrencyText = value
        End Set
    End Property

    Public Property ExchangeRateText As String
        Get
            Return strExchangeRateText
        End Get
        Set(value As String)
            strExchangeRateText = value
        End Set
    End Property

    Public Property Conversion1 As String
        Get
            Return strConversion1
        End Get
        Set(value As String)
            strConversion1 = value
        End Set
    End Property

    Public Property Conversion2 As String
        Get
            Return strConversion2
        End Get
        Set(value As String)
            strConversion2 = value
        End Set
    End Property

    Public Property DefaultCurrencyCode As String
        Get
            Return strDefaultCurrencyCode
        End Get
        Set(value As String)
            strDefaultCurrencyCode = value
        End Set
    End Property
End Class

Public Class PrintBudgetRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal printplanningrow As PrintBudgetRow)
        List.Add(printplanningrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal printplanningrow As PrintBudgetRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(printplanningrow)
        Else
            List.Insert(index, printplanningrow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal printplanningrow As PrintBudgetRow)
        If List.Contains(printplanningrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(printplanningrow)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintBudgetRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintBudgetRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal printplanningrow As PrintBudgetRow) As Integer
        Return List.IndexOf(printplanningrow)
    End Function

    Public Function Contains(ByVal printplanningrow As PrintBudgetRow) As Boolean
        Return List.Contains(printplanningrow)
    End Function
End Class
