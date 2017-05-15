Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class BudgetItem
    Inherits LogframeObject

    Private intIdBudgetItem As Integer = -1
    Private intIdBudgetYear, intIdResource As Integer
    Private sngDuration As Single
    Private strDurationUnit As String
    Private sngNumber As Single
    Private strNumberUnit As String
    Private curUnitCost As New Currency
    Private curTotalLocalCost As New Currency
    Private curTotalCost As New Currency
    Private intType As Integer
    Private dblRatio As Double
    Private strFormula As String = String.Empty
    Private boolSetValue As Boolean
    Private objParentBudgetYearGuid, objParentBudgetItemGuid, objReferenceBudgetItemGuid As Guid
    Private lstCalculationGuids As New List(Of Guid)

    Public Event TotalCostChanged()

    <ScriptIgnore()> _
    Public WithEvents BudgetItems As New BudgetItems

    Public Enum BudgetItemTypes
        Normal
        Ratio
        Formula
        Category
        Reference
        ReferenceSource
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String, Optional ByVal boolTextToRichText As Boolean = False)
        If boolTextToRichText = False Then
            Me.RTF = RTF
        Else
            Me.RTF = TextToRichText(RTF)
        End If
    End Sub

    Public Sub New(ByVal rtf As String, ByVal number As Single, ByVal numberunit As String, _
                   ByVal duration As Single, ByVal durationunit As String, ByVal unitcost As Currency, ByVal totalcost As Currency)
        Me.RTF = rtf
        Me.Number = number
        Me.NumberUnit = numberunit
        Me.Duration = duration
        Me.DurationUnit = durationunit
        Me.UnitCost = unitcost
        Me.TotalCost = totalcost
    End Sub

    Public Sub New(ByVal rtf As String, ByVal referencebudgetitemguid As Guid)
        Me.RTF = rtf
        Me.ReferenceBudgetItemGuid = referencebudgetitemguid
        Me.Type = BudgetItemTypes.Reference
    End Sub

#Region "Properties"
    Public Property idBudgetItem As Integer
        Get
            Return intIdBudgetItem
        End Get
        Set(ByVal value As Integer)
            intIdBudgetItem = value
        End Set
    End Property

    Public Property idBudgetYear As Integer
        Get
            Return intIdBudgetYear
        End Get
        Set(ByVal value As Integer)
            intIdBudgetYear = value
        End Set
    End Property

    Public Property ParentBudgetYearGuid() As Guid
        Get
            Return objParentBudgetYearGuid
        End Get
        Set(ByVal value As Guid)
            objParentBudgetYearGuid = value
            objParentBudgetItemGuid = Guid.Empty
        End Set
    End Property

    Public Property ParentBudgetItemGuid As Guid
        Get
            Return objParentBudgetItemGuid
        End Get
        Set(ByVal value As Guid)
            objParentBudgetItemGuid = value
            objParentBudgetYearGuid = Guid.Empty
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

    Public Property CalculationGuids As List(Of Guid)
        Get
            Return lstCalculationGuids
        End Get
        Set(value As List(Of Guid))
            lstCalculationGuids = value
        End Set
    End Property

    Public Property CalculationGuidRatio() As Guid
        Get
            If CalculationGuids.Count > 0 And Me.Type = BudgetItemTypes.Ratio Then
                Return CalculationGuids(0)
            Else
                Return Guid.Empty
            End If
        End Get
        Set(value As Guid)
            If Me.Type <> BudgetItemTypes.Ratio Then Me.Type = BudgetItemTypes.Ratio
            CalculationGuids.Clear()
            CalculationGuids.Add(value)
        End Set
    End Property

    Public Property idResource As Integer
        Get
            Return intIdResource
        End Get
        Set(ByVal value As Integer)
            intIdResource = value
        End Set
    End Property

    Public Property Duration() As Single
        Get
            Return sngDuration
        End Get
        Set(ByVal value As Single)
            sngDuration = value
        End Set
    End Property

    Public Property DurationUnit() As String
        Get
            Return strDurationUnit
        End Get
        Set(ByVal value As String)
            strDurationUnit = value
        End Set
    End Property

    Public ReadOnly Property DurationSet As Boolean
        Get
            If Duration = 0 And String.IsNullOrEmpty(DurationUnit) Then Return False Else Return True
        End Get
    End Property

    Public Property Number() As Single
        Get
            Return sngNumber
        End Get
        Set(ByVal value As Single)
            sngNumber = value
        End Set
    End Property

    Public Property NumberUnit() As String
        Get
            Return strNumberUnit
        End Get
        Set(ByVal value As String)
            strNumberUnit = value
        End Set
    End Property

    Public ReadOnly Property NumberSet As Boolean
        Get
            If Number = 0 And String.IsNullOrEmpty(NumberUnit) Then Return False Else Return True
        End Get
    End Property

    <ScriptIgnore()> _
    Public Property UnitCost() As Currency
        Get
            Return curUnitCost
        End Get
        Set(ByVal value As Currency)
            curUnitCost = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property UnitCostAmount() As Single
        Get
            Return curUnitCost.Amount
        End Get
        Set(ByVal value As Single)
            curUnitCost.Amount = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property UnitCostCurrencyCode() As String
        Get
            Return curUnitCost.CurrencyCode
        End Get
        Set(ByVal value As String)
            curUnitCost.CurrencyCode = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property UnitCostExchangeRate() As Single
        Get
            Return curUnitCost.ExchangeRate
        End Get
        Set(ByVal value As Single)
            curUnitCost.ExchangeRate = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property TotalLocalCost() As Currency
        Get
            Return curTotalLocalCost
        End Get
        Set(ByVal value As Currency)
            curTotalLocalCost = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property TotalLocalCostAmount() As Single
        Get
            Return curTotalLocalCost.Amount
        End Get
        Set(ByVal value As Single)
            curTotalLocalCost.Amount = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property TotalLocalCostCurrencyCode() As String
        Get
            Return curTotalLocalCost.CurrencyCode
        End Get
        Set(ByVal value As String)
            curTotalLocalCost.CurrencyCode = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property TotalLocalCostExchangeRate() As Single
        Get
            Return curTotalLocalCost.ExchangeRate
        End Get
        Set(ByVal value As Single)
            curTotalLocalCost.ExchangeRate = value
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

    <XmlIgnore()> _
    Public Property TotalCostAmount() As Single
        Get
            Return curTotalCost.Amount
        End Get
        Set(ByVal value As Single)
            curTotalCost.Amount = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property TotalCostCurrencyCode() As String
        Get
            Return curTotalCost.CurrencyCode
        End Get
        Set(ByVal value As String)
            curTotalCost.CurrencyCode = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Property TotalCostExchangeRate() As Single
        Get
            Return curTotalCost.ExchangeRate
        End Get
        Set(ByVal value As Single)
            curTotalCost.ExchangeRate = value
        End Set
    End Property

    Public Property Type() As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    Public Property Ratio() As Double
        Get
            Return dblRatio
        End Get
        Set(ByVal value As Double)
            dblRatio = value
        End Set
    End Property

    Public Property Formula As String
        Get
            Return strFormula
        End Get
        Set(value As String)
            strFormula = value
        End Set
    End Property

    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_BudgetItem
        End Get
    End Property

    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_BudgetItems
        End Get
    End Property
#End Region

#Region "Events"
    Private Sub BudgetItems_BudgetItemAdded(sender As Object, e As BudgetItemAddedEventArgs) Handles BudgetItems.BudgetItemAdded
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        selBudgetItem.idBudgetItem = Me.idBudgetItem
        selBudgetItem.ParentBudgetItemGuid = Me.Guid

        If Me.Type <> BudgetItemTypes.Reference And Me.Type <> BudgetItemTypes.ReferenceSource Then
            Me.Type = BudgetItemTypes.Category
        End If
    End Sub

    Private Sub BudgetItems_BudgetItemRemoved() Handles BudgetItems.BudgetItemRemoved
        If Me.BudgetItems.Count = 0 Then
            If Me.ReferenceBudgetItemGuid <> Guid.Empty Then
                Me.Type = BudgetItemTypes.Reference
            Else
                Me.Type = BudgetItemTypes.Normal
            End If
        End If
    End Sub

    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selBudgetItem As BudgetItem In Me.BudgetItems
            selBudgetItem.ParentBudgetItemGuid = Me.Guid
        Next
    End Sub
#End Region

#Region "Methods"
    Public Sub SetTotalCost()
        Dim sngNumberValue As Single = 1
        Dim sngDurationValue As Single = 1
        Dim sngAmount As Single
        If Me.UnitCost IsNot Nothing AndAlso Me.UnitCost.Amount > 0 Then
            If NumberSet = True Then
                sngNumberValue = Me.Number
            End If
            If DurationSet = True Then
                sngDurationValue = Me.Duration
            End If
            sngAmount = Me.UnitCost.Amount * sngNumberValue * sngDurationValue
            TotalLocalCost = New Currency(sngAmount, UnitCost.CurrencyCode)

            If UnitCost.ExchangeRate <> 0 Then sngAmount /= UnitCost.ExchangeRate
            TotalCost = New Currency(sngAmount, CurrentLogFrame.CurrencyCode)
        End If
    End Sub

    Public Sub SetParentTotals()
        If Me.BudgetItems.Count > 0 Then
            Me.TotalLocalCost = BudgetItems.GetTotalLocalCost
            Me.TotalCost = BudgetItems.GetTotalCost
            Me.Number = BudgetItems.GetTotalNumber
            If BudgetItems.UniformNumberUnit = True Then
                Me.NumberUnit = BudgetItems(0).NumberUnit
            End If
        Else
            Me.TotalLocalCost = New Currency()
            Me.TotalCost = New Currency()
            Me.Number = 0
            Me.NumberUnit = String.Empty
        End If
    End Sub

    Public Function GetCalculationGuid() As Guid
        If CalculationGuids.Count > 0 Then
            Return CalculationGuids(0)
        Else
            Return Guid.Empty
        End If
    End Function

    Public Sub SetCalculationGuid(ByVal objGuid As Guid)
        CalculationGuids.Clear()
        CalculationGuids.Add(objGuid)
    End Sub

    Public Sub CalculateRatio()
        If Me.CalculationGuidRatio <> Guid.Empty Then
            Dim CalculationReference As BudgetItem = CurrentLogFrame.GetBudgetItemByGuid(CalculationGuidRatio)

            If CalculationReference IsNot Nothing Then
                Dim sngAmount As Single = CalculationReference.TotalCost.Amount * Me.Ratio
                Me.TotalCost = New Currency(sngAmount, CalculationReference.TotalCost.CurrencyCode)
            End If
        End If
    End Sub

    Public Function GetLevelsOfChildren() As Integer
        Dim intLevels As Integer

        If BudgetItems.Count > 0 Then
            intLevels = BudgetItems.GetLevelsOfChildren(0)
        End If
        Return intLevels
    End Function
#End Region

End Class

Public Class BudgetItems
    Inherits System.Collections.CollectionBase

    Public Event TotalCostChanged()
    Public Event BudgetItemAdded(ByVal sender As Object, ByVal e As BudgetItemAddedEventArgs)
    Public Event BudgetItemRemoved()

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As BudgetItem
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), BudgetItem)
            End If
        End Get
    End Property
#End Region

#Region "Totals"
    Public Function GetTotalCost() As Currency
        Dim curTotalCost As New Currency()
        For Each selItem As BudgetItem In Me.List
            If selItem.TotalCost IsNot Nothing Then
                If curTotalCost.CurrencyCode <> selItem.TotalCost.CurrencyCode Then _
                    curTotalCost.CurrencyCode = selItem.TotalCost.CurrencyCode

                curTotalCost.Amount += selItem.TotalCost.Amount
            End If
        Next

        Return curTotalCost
    End Function

    Public Function GetTotalLocalCost() As Currency
        Dim curTotalLocalCost As New Currency()
        If UniformCurrencyCode() = True Then
            For Each selItem As BudgetItem In Me.List
                If selItem.TotalLocalCost IsNot Nothing Then
                    If curTotalLocalCost.CurrencyCode <> selItem.TotalLocalCost.CurrencyCode Then
                        curTotalLocalCost.CurrencyCode = selItem.TotalLocalCost.CurrencyCode
                        curTotalLocalCost.ExchangeRate = selItem.TotalLocalCost.ExchangeRate
                    End If
                    curTotalLocalCost.Amount += selItem.TotalLocalCost.Amount
                End If
            Next
        End If
        Return curTotalLocalCost
    End Function

    Public Function UniformCurrencyCode() As Boolean
        Dim strCurrencyCode As String = String.Empty

        For Each selItem As BudgetItem In Me.List
            If IndexOf(selItem) = 0 Then
                strCurrencyCode = selItem.TotalLocalCostCurrencyCode
            End If
            If selItem.TotalLocalCostCurrencyCode <> strCurrencyCode Then
                Return False
            End If
        Next

        Return True
    End Function

    Public Function GetTotalNumber() As Double
        Dim dblNumber As Double

        If UniformNumberUnit() = True Then
            For Each selItem As BudgetItem In Me.List
                dblNumber += selItem.Number
            Next
        End If
        Return dblNumber
    End Function

    Public Function UniformNumberUnit() As Boolean
        Dim strUnit As String = String.Empty

        For Each selItem As BudgetItem In Me.List
            If IndexOf(selItem) = 0 Then
                strUnit = selItem.NumberUnit
            End If
            If selItem.NumberUnit <> strUnit Then
                Return False
            End If
        Next

        Return True
    End Function

    Public Function GetLevelsOfChildren(ByVal intLevels As Integer) As Integer
        Dim intLevelsChild As Integer

        If Me.Count > 0 Then
            intLevels += 1

            For Each selBudgetItem As BudgetItem In Me.List
                intLevelsChild = selBudgetItem.BudgetItems.GetLevelsOfChildren(intLevels)
                If intLevelsChild > intLevels Then intLevels = intLevelsChild
            Next
        End If

        Return intLevels
    End Function

    Public Sub UpdateCalculatedAmounts()
        For Each selBudgetItem As BudgetItem In Me.List
            Select Case selBudgetItem.Type
                Case BudgetItem.BudgetItemTypes.Ratio
                    selBudgetItem.CalculateRatio()
                Case BudgetItem.BudgetItemTypes.Formula

                Case Else
                    If selBudgetItem.BudgetItems.Count > 0 Then selBudgetItem.BudgetItems.UpdateCalculatedAmounts()
            End Select
        Next
    End Sub
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal budgetitem As BudgetItem)
        List.Add(budgetitem)

        RaiseEvent BudgetItemAdded(Me, New BudgetItemAddedEventArgs(budgetitem))
        RaiseEvent TotalCostChanged()
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal budgetitem As BudgetItem)
        If intIndex > Count Or intIndex < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, budgetitem.ItemName))
        ElseIf intIndex = Count Then
            List.Add(budgetitem)

            RaiseEvent BudgetItemAdded(Me, New BudgetItemAddedEventArgs(budgetitem))
            RaiseEvent TotalCostChanged()
        Else
            List.Insert(intIndex, budgetitem)

            RaiseEvent BudgetItemAdded(Me, New BudgetItemAddedEventArgs(budgetitem))
            RaiseEvent TotalCostChanged()
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, BudgetItem.ItemName))
        Else
            List.RemoveAt(index)

            RaiseEvent BudgetItemRemoved()
            RaiseEvent TotalCostChanged()
        End If
    End Sub

    Public Sub Remove(ByVal budgetitem As BudgetItem)
        If List.Contains(budgetitem) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, budgetitem.ItemName))
        Else
            List.Remove(budgetitem)

            RaiseEvent BudgetItemRemoved()
            RaiseEvent TotalCostChanged()
        End If
    End Sub

    Public Function IndexOf(ByVal budgetitem As BudgetItem) As Integer
        Return List.IndexOf(budgetitem)
    End Function

    Public Function Contains(ByVal budgetitem As BudgetItem) As Boolean
        Return List.Contains(budgetitem)
    End Function

    Public Function GetBudgetItemByGuid(ByVal objGuid As Guid) As BudgetItem
        Dim selBudgetItem As BudgetItem = Nothing
        For Each objBudgetItem As BudgetItem In Me.List
            If objBudgetItem.Guid = objGuid Then
                selBudgetItem = objBudgetItem
                Exit For
            ElseIf objBudgetItem.BudgetItems.Count > 0 Then
                selBudgetItem = objBudgetItem.BudgetItems.GetBudgetItemByGuid(objGuid)
                If selBudgetItem IsNot Nothing Then Exit For
            End If
        Next
        Return selBudgetItem
    End Function

    Public Function GetBudgetItemByReferenceGuid(ByVal objReferenceGuid As Guid) As BudgetItem
        Dim selBudgetItem As BudgetItem = Nothing
        For Each objBudgetItem As BudgetItem In Me.List
            If objBudgetItem.ReferenceBudgetItemGuid = objReferenceGuid Then
                selBudgetItem = objBudgetItem
                Exit For
            ElseIf objBudgetItem.BudgetItems.Count > 0 Then
                selBudgetItem = objBudgetItem.BudgetItems.GetBudgetItemByReferenceGuid(objReferenceGuid)
                If selBudgetItem IsNot Nothing Then Exit For
            End If
        Next
        Return selBudgetItem
    End Function
#End Region

End Class

Public Class BudgetItemAddedEventArgs
    Inherits EventArgs

    Public Property BudgetItem As BudgetItem

    Public Sub New(ByVal objBudgetItem As BudgetItem)
        MyBase.New()

        Me.BudgetItem = objBudgetItem
    End Sub
End Class

Public Class BudgetItemRemovedEventArgs
    Inherits EventArgs

    Public Property Guid As Guid
    Public Property ParentGuid As Guid

    Public Sub New(ByVal objGuid As Guid, ByVal objParentGuid As Guid)
        MyBase.New()

        Me.Guid = objGuid
        Me.ParentGuid = objParentGuid
    End Sub
End Class
