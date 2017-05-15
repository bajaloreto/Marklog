Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Budget
    Private intIdBudget, intIdLogframe As Integer
    Private strDefaultCurrencyCode As String = My.Settings.setDefaultCurrency
    Private boolMultiYearBudget As Boolean
    Private objGuid, objParentLogframe As Guid

    <ScriptIgnore()> _
    Public WithEvents BudgetYears As New BudgetYears

    <ScriptIgnore()> _
    Public WithEvents ExchangeRates As New ExchangeRates

#Region "Properties"
    Public Property idBudget As Integer
        Get
            Return intIdBudget
        End Get
        Set(ByVal value As Integer)
            intIdBudget = value
        End Set
    End Property

    Public Property idLogframe As Integer
        Get
            Return intIdLogframe
        End Get
        Set(ByVal value As Integer)
            intIdLogframe = value
        End Set
    End Property

    Public Property ParentLogframeGuid() As Guid
        Get
            Return objParentLogframe
        End Get
        Set(ByVal value As Guid)
            objParentLogframe = value
        End Set
    End Property

    Public Property DefaultCurrencyCode() As String
        Get
            Return strDefaultCurrencyCode
        End Get
        Set(ByVal value As String)
            strDefaultCurrencyCode = Left(value, 3).ToUpper
        End Set
    End Property

    Public Property MultiYearBudget As Boolean
        Get
            Return boolMultiYearBudget
        End Get
        Set(value As Boolean)
            boolMultiYearBudget = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Budget
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Budgets
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Function LoadUsedCurrencyCodesList() As List(Of IdValuePair)
        Dim lstCurrencyCodes As List(Of IdValuePair) = LoadCurrencyCodesList()
        Dim lstUsedCodes As New List(Of IdValuePair)
        Dim selValuePair As IdValuePair = FindCurrency(Me.DefaultCurrencyCode, lstCurrencyCodes)

        If selValuePair IsNot Nothing Then lstUsedCodes.Add(selValuePair)

        For Each selExchangeRate As ExchangeRate In Me.ExchangeRates
            selValuePair = FindCurrency(selExchangeRate.CurrencyCode, lstCurrencyCodes)
            If selValuePair IsNot Nothing Then
                lstUsedCodes.Add(selValuePair)
            End If
        Next

        Return lstUsedCodes
    End Function

    Private Function FindCurrency(ByVal strCurrencyCode As String, ByVal lstCurrencyCodes As List(Of IdValuePair)) As IdValuePair
        For Each selValuePair As IdValuePair In lstCurrencyCodes
            If selValuePair.Id = strCurrencyCode Then
                Return selValuePair
            End If
        Next

        Return Nothing
    End Function

    Public Function GetBudgetYearByGuid(ByVal objGuid As Guid) As BudgetYear
        Dim selBudgetYear As BudgetYear = Nothing

        For Each objBudgetYear As BudgetYear In Me.BudgetYears
            If objBudgetYear.Guid = objGuid Then
                selBudgetYear = objBudgetYear
                Exit For
            End If
        Next
        Return selBudgetYear
    End Function

    Public Function GetBudgetItemByGuid(ByVal objGuid As Guid) As BudgetItem
        Dim selBudgetItem As BudgetItem = Nothing

        For Each selBudgetYear As BudgetYear In Me.BudgetYears
            selBudgetItem = selBudgetYear.BudgetItems.GetBudgetItemByGuid(objGuid)
            If selBudgetItem IsNot Nothing Then Exit For
        Next

        Return selBudgetItem
    End Function

    Public Function GetBudgetItemsList(Optional ByVal ParentBudgetYearGuid As Guid = Nothing) As Dictionary(Of Guid, String)
        Dim objBudgetItemsList As New Dictionary(Of Guid, String)
        Dim selBudgetYear As BudgetYear

        objBudgetItemsList.Add(Guid.Empty, LANG_NotSelected)
        If ParentBudgetYearGuid = Nothing Or ParentBudgetYearGuid = Guid.Empty Then
            selBudgetYear = Me.BudgetYears(0)

            objBudgetItemsList = GetBudgetItemsList_BudgetItems(selBudgetYear.BudgetItems, objBudgetItemsList)
        Else
            selBudgetYear = GetBudgetYearByGuid(ParentBudgetYearGuid)

            objBudgetItemsList = GetBudgetItemsList_BudgetItems(selBudgetYear.BudgetItems, objBudgetItemsList)
        End If

        Return objBudgetItemsList
    End Function

    Public Function GetBudgetItemsList_BudgetItems(ByVal objBudgetItems As BudgetItems, ByVal objBudgetItemsList As Dictionary(Of Guid, String), Optional ByVal strParentSortNumber As String = "") As Dictionary(Of Guid, String)
        Dim strText, strSortNumber As String
        Dim intIndex As Integer

        For Each selBudgetItem As BudgetItem In objBudgetItems
            intIndex = objBudgetItems.IndexOf(selBudgetItem)
            strSortNumber = CurrentLogFrame.CreateSortNumber(intIndex, strParentSortNumber)

            strText = String.Format("{0}  {1}   ----> {2}", strSortNumber, selBudgetItem.Text, selBudgetItem.TotalCost.ToString)
            objBudgetItemsList.Add(selBudgetItem.Guid, strText)

            If selBudgetItem.BudgetItems.Count > 0 Then
                objBudgetItemsList = GetBudgetItemsList_BudgetItems(selBudgetItem.BudgetItems, objBudgetItemsList, strSortNumber)
            End If
        Next

        Return objBudgetItemsList
    End Function
#End Region

#Region "Update parent totals (vertical logic)"
    Public Sub UpdateParentTotals(ByVal ChildItem As BudgetItem)
        If ChildItem.ParentBudgetItemGuid <> Guid.Empty Then
            Dim ParentBudgetItem As BudgetItem = Me.GetBudgetItemByGuid(ChildItem.ParentBudgetItemGuid)

            If ParentBudgetItem Is Nothing Then Exit Sub
            ParentBudgetItem.SetParentTotals()

            UpdateParentTotals(ParentBudgetItem)
        ElseIf ChildItem.ParentBudgetYearGuid <> Guid.Empty Then
            Dim ParentBudgetYear As BudgetYear = CurrentLogFrame.GetBudgetYearByGuid(ChildItem.ParentBudgetYearGuid)

            If ParentBudgetYear Is Nothing Then Exit Sub

            ParentBudgetYear.TotalCost = ParentBudgetYear.GetTotalCost()
        End If
    End Sub

    Public Sub UpdateChildTotals(ByVal Parent As Object)
        If Parent Is Nothing Then Exit Sub

        Select Case Parent.GetType
            Case GetType(BudgetYear)
                Dim objBudgetYear As BudgetYear = CType(Parent, BudgetYear)

                objBudgetYear.TotalCost = objBudgetYear.GetTotalCost()
            Case GetType(BudgetItem)
                Dim ParentBudgetItem As BudgetItem = CType(Parent, BudgetItem)
                ParentBudgetItem.SetParentTotals()

                UpdateParentTotals(ParentBudgetItem)
        End Select
    End Sub
#End Region

#Region "Methods for multi-year budget (horizontal logic)"
    Public Sub ChangeRatioBudgetHeader(ByVal selBudgetItem As BudgetItem)
        If MultiYearBudget = True And BudgetYears.Count > 0 Then
            For i = 1 To BudgetYears.Count - 1
                Dim selBudgetYear As BudgetYear = BudgetYears(i)

                selBudgetYear.ChangeRatioBudgetHeader(selBudgetItem)
                'UpdateParentTotals(selBudgetItem)
            Next
        End If
    End Sub

    Public Sub InsertBudgetHeader(ByVal selBudgetItem As BudgetItem)
        If MultiYearBudget = True And BudgetYears.Count > 0 Then
            For i = 1 To BudgetYears.Count - 1
                Dim selBudgetYear As BudgetYear = BudgetYears(i)

                selBudgetYear.InsertBudgetHeader(selBudgetItem)
                UpdateParentTotals(selBudgetItem)
            Next
        End If
    End Sub

    Public Sub RemoveBudgetHeader(ByVal objGuid As Guid, ByVal objParentGuid As Guid)
        If MultiYearBudget = True And BudgetYears.Count > 0 Then
            For i = 1 To BudgetYears.Count - 1
                Dim selBudgetYear As BudgetYear = BudgetYears(i)

                selBudgetYear.RemoveBudgetHeader(objGuid)

                Dim ParentItem As BudgetItem = selBudgetYear.BudgetItems.GetBudgetItemByReferenceGuid(objParentGuid)
                UpdateChildTotals(ParentItem)
            Next
        End If
    End Sub

    Public Sub MoveBudgetHeader(ByVal selBudgetItem As BudgetItem)
        For i = 1 To BudgetYears.Count - 1
            Dim selBudgetYear As BudgetYear = BudgetYears(i)

            selBudgetYear.MoveBudgetHeader(selBudgetItem)
            UpdateParentTotals(selBudgetItem)
        Next
    End Sub

    Public Sub InsertParentHeader(ByVal selBudgetItem As BudgetItem)
        For i = 1 To BudgetYears.Count - 1
            Dim selBudgetYear As BudgetYear = BudgetYears(i)

            selBudgetYear.InsertParentHeader(selBudgetItem)
            UpdateParentTotals(selBudgetItem)
        Next
    End Sub

    Public Sub InsertChildHeader(ByVal selBudgetItem As BudgetItem)
        For i = 1 To BudgetYears.Count - 1
            Dim selBudgetYear As BudgetYear = BudgetYears(i)

            selBudgetYear.InsertChildHeader(selBudgetItem)
            UpdateChildTotals(selBudgetItem)
        Next
    End Sub

    Public Sub LevelDownHeader(ByVal selBudgetItem As BudgetItem)
        For i = 1 To BudgetYears.Count - 1
            Dim selBudgetYear As BudgetYear = BudgetYears(i)

            selBudgetYear.LevelDownHeader(selBudgetItem)
            UpdateParentTotals(selBudgetItem)
        Next
    End Sub

    Public Sub LevelUpHeader(ByVal selBudgetItem As BudgetItem)
        For i = 1 To BudgetYears.Count - 1
            Dim selBudgetYear As BudgetYear = BudgetYears(i)

            selBudgetYear.LevelUpHeader(selBudgetItem)
            UpdateParentTotals(selBudgetItem)
        Next
    End Sub

    Public Sub UpdateBudgetYears(ByVal datStartDate As Date, ByVal datEndDate As Date)
        If MultiYearBudget = False Then
            If BudgetYears.Count = 0 Then
                BudgetYears.Add(New BudgetYear(datStartDate))
            ElseIf BudgetYears.Count = 1 Then
                BudgetYears(0).BudgetYear = datStartDate
            Else
                For i = 1 To BudgetYears.Count - 1
                    BudgetYears.RemoveAt(1)
                Next
                BudgetYears(0).BudgetYear = datStartDate
            End If
        Else
            Dim intStartYear As Integer
            Dim intEndYear As Integer

            intStartYear = datStartDate.Year
            intEndYear = datEndDate.Year

            If BudgetYears.Count = 0 Then
                'Totals page
                BudgetYears.Add(New BudgetYear())

                'Budget years
                For i = intStartYear To intEndYear
                    BudgetYears.Add(New BudgetYear(DateSerial(i, datStartDate.Month, datStartDate.Day)))
                Next
            ElseIf BudgetYears.Count = 1 Then
                'Totals page exists, only add pages for budget years
                For i = intStartYear To intEndYear
                    BudgetYears.Add(New BudgetYear(DateSerial(i, datStartDate.Month, datStartDate.Day)))
                Next
            Else
                'Add additional pages for budget years that are missing
                For i = intStartYear To intEndYear
                    If i < BudgetYears(0).BudgetYear.Year Then
                        BudgetYears.Insert(1, New BudgetYear(DateSerial(i, datStartDate.Month, datStartDate.Day)))
                    ElseIf i > BudgetYears(BudgetYears.Count - 1).BudgetYear.Year Then
                        BudgetYears.Add(New BudgetYear(DateSerial(i, datStartDate.Month, datStartDate.Day)))
                    End If
                Next
            End If

            UpdateBudgetYears_CopyHeaders()
        End If
    End Sub

    Private Sub UpdateBudgetYears_CopyHeaders()
        If MultiYearBudget = True Then
            Dim objTotalBudget As BudgetYear = BudgetYears(0)

            UpdateBudgetYears_CopyHeaders_Copy(objTotalBudget.BudgetItems)
        End If
    End Sub

    Private Sub UpdateBudgetYears_CopyHeaders_Copy(ByVal objBudgetItems As BudgetItems)
        For Each selBudgetItem As BudgetItem In objBudgetItems
            InsertBudgetHeader(selBudgetItem)

            If selBudgetItem.BudgetItems.Count > 0 Then
                UpdateBudgetYears_CopyHeaders_Copy(selBudgetItem.BudgetItems)
            End If
        Next
    End Sub

    Public Sub UpdateMultiYearBudget()
        If Me.MultiYearBudget = True Then
            UpdateMultiYearBudget_ReferencedAmounts()
            UpdateMultiYearBudget_CalculatedAmounts()
        End If
    End Sub

    Public Sub UpdateMultiYearBudget_ReferencedAmounts()
        Dim TotalBudget As BudgetYear = Me.BudgetYears(0)
        Dim selBudgetYear As BudgetYear
        Dim ReferenceItems As BudgetItems = TotalBudget.GetReferenceItems
        Dim selBudgetItem As BudgetItem
        Dim sngTotalAmount As Single

        For Each ReferenceItem As BudgetItem In ReferenceItems
            For i = 1 To Me.BudgetYears.Count - 1
                selBudgetYear = Me.BudgetYears(i)
                selBudgetItem = selBudgetYear.BudgetItems.GetBudgetItemByReferenceGuid(ReferenceItem.Guid)
                If selBudgetItem IsNot Nothing Then sngTotalAmount += selBudgetItem.TotalCostAmount
            Next

            ReferenceItem.TotalCostAmount = sngTotalAmount
            UpdateParentTotals(ReferenceItem)
            sngTotalAmount = 0
        Next
    End Sub

    Public Sub UpdateMultiYearBudget_CalculatedAmounts()
        Dim selBudgetYear As BudgetYear = Nothing

        If Me.MultiYearBudget = True Then
            For i = 1 To BudgetYears.Count - 1
                selBudgetYear = BudgetYears(i)

                UpdateMultiYearBudget_CalculatedAmounts_Update(selBudgetYear.BudgetItems)
            Next
        End If

        selBudgetYear = BudgetYears(0)
        UpdateMultiYearBudget_CalculatedAmounts_Update(selBudgetYear.BudgetItems)
    End Sub

    Private Sub UpdateMultiYearBudget_CalculatedAmounts_Update(ByVal objBudgetItems As BudgetItems)
        For Each selBudgetItem As BudgetItem In objBudgetItems
            Select Case selBudgetItem.Type
                Case BudgetItem.BudgetItemTypes.Ratio
                    selBudgetItem.CalculateRatio()
                    UpdateParentTotals(selBudgetItem)
                Case BudgetItem.BudgetItemTypes.Formula

                Case Else
                    If selBudgetItem.BudgetItems.Count > 0 Then UpdateMultiYearBudget_CalculatedAmounts_Update(selBudgetItem.BudgetItems)
            End Select
        Next
    End Sub


#End Region

#Region "Events"
    Private Sub BudgetYears_BudgetYearAdded(ByVal sender As Object, ByVal e As BudgetYearAddedEventArgs) Handles BudgetYears.BudgetYearAdded
        Dim selBudgetYear As BudgetYear = e.BudgetYear

        selBudgetYear.idBudget = Me.idBudget
        selBudgetYear.ParentBudgetGuid = Me.Guid
    End Sub
#End Region
End Class

Public Class UpdateParentTotalsEventArgs
    Inherits EventArgs

    Public Property ChildItem As BudgetItem

    Public Sub New(ByVal objChildItem As BudgetItem)
        MyBase.New()

        Me.ChildItem = objChildItem
    End Sub
End Class

Public Class UpdateChildTotalsEventArgs
    Inherits EventArgs

    Public Property Parent As Object

    Public Sub New(ByVal objParent As Object)
        MyBase.New()

        If objParent IsNot Nothing Then
            If objParent.GetType = GetType(BudgetYear) Or objParent.GetType = GetType(BudgetItem) Then
                Me.Parent = objParent
            Else
                Me.Parent = Nothing
            End If
        End If
    End Sub
End Class
