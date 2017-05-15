Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class BudgetYear
    Private intIdBudgetYear, intIdBudget As Integer
    Private datBudgetYear As Date
    Private strDescription As String
    Private curTotalCost As New Currency
    Private objGuid, objParentBudgetGuid As Guid
    Private ReferenceItems As New BudgetItems

    <ScriptIgnore()> _
    Public WithEvents BudgetItems As New BudgetItems

#Region "Properties"
    Public Property idBudgetYear As Integer
        Get
            Return intIdBudgetYear
        End Get
        Set(ByVal value As Integer)
            intIdBudgetYear = value
        End Set
    End Property

    Public Property idBudget As Integer
        Get
            Return intIdBudget
        End Get
        Set(ByVal value As Integer)
            intIdBudget = value
        End Set
    End Property

    Public Property ParentBudgetGuid() As Guid
        Get
            Return objParentBudgetGuid
        End Get
        Set(ByVal value As Guid)
            objParentBudgetGuid = value
        End Set
    End Property

    Public Property BudgetYear As Date
        Get
            Return datBudgetYear
        End Get
        Set(value As Date)
            datBudgetYear = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return strDescription
        End Get
        Set(ByVal value As String)
            strDescription = value
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

    Public Property Guid() As Guid
        Get
            If objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
                UpdateParentGuid()
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
            UpdateParentGuid()
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_BudgetYear
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_BudgetYears
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal budgetyear As Date)
        Me.BudgetYear = budgetyear
    End Sub

    Public Function GetTotalCost() As Currency
        Me.BudgetItems.UpdateCalculatedAmounts()
        Return Me.BudgetItems.GetTotalCost
    End Function

    Public Function GetReferenceItems() As BudgetItems
        ReferenceItems.Clear()

        GetReferenceItems_Get(Me.BudgetItems)

        Return ReferenceItems
    End Function

    Private Sub GetReferenceItems_Get(ByVal objBudgetItems As BudgetItems)
        For Each selBudgetItem As BudgetItem In objBudgetItems
            If selBudgetItem.BudgetItems.Count = 0 Then
                ReferenceItems.Add(selBudgetItem)
            Else
                GetReferenceItems_Get(selBudgetItem.BudgetItems)
            End If
        Next
    End Sub

    Private Sub UpdateParentGuid()
        For Each selBudgetItem As BudgetItem In Me.BudgetItems
            selBudgetItem.ParentBudgetYearGuid = Me.Guid
        Next
    End Sub

    Public Sub UpdateParentTotals(ByVal ChildItem As BudgetItem)
        If ChildItem.ParentBudgetItemGuid <> Guid.Empty Then
            Dim ParentBudgetItem As BudgetItem = Me.BudgetItems.GetBudgetItemByGuid(ChildItem.ParentBudgetItemGuid)

            If ParentBudgetItem Is Nothing Then Exit Sub
            ParentBudgetItem.SetParentTotals()

            UpdateParentTotals(ParentBudgetItem)
        ElseIf ChildItem.ParentBudgetYearGuid = Me.Guid Then
            Me.TotalCost = GetTotalCost()
        End If
    End Sub

    Public Sub ChangeRatioBudgetHeader(ByVal objBudgetHeader As BudgetItem)
        Dim selItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objBudgetHeader.Guid)
        Dim CalculationReference As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objBudgetHeader.CalculationGuidRatio)

        If selItem Is Nothing Or CalculationReference Is Nothing Then Exit Sub

        With selItem
            .Type = BudgetItem.BudgetItemTypes.Ratio
            .RTF = objBudgetHeader.RTF
            .CalculationGuidRatio = CalculationReference.Guid
            .Ratio = objBudgetHeader.Ratio
            .CalculateRatio()
        End With
    End Sub

    Public Sub InsertBudgetHeader(ByVal objBudgetHeader As BudgetItem)
        Dim intIndex As Integer

        If objBudgetHeader.ParentBudgetYearGuid <> Guid.Empty Then
            Dim ParentBudgetYear As BudgetYear = CurrentLogFrame.GetBudgetYearByGuid(objBudgetHeader.ParentBudgetYearGuid)

            If ParentBudgetYear IsNot Nothing Then
                intIndex = ParentBudgetYear.BudgetItems.IndexOf(objBudgetHeader)

                If intIndex >= 0 And intIndex <= Me.BudgetItems.Count - 1 Then
                    Dim CurrentBudgetItem As BudgetItem = Me.BudgetItems(intIndex)

                    If CurrentBudgetItem IsNot Nothing AndAlso CurrentBudgetItem.ReferenceBudgetItemGuid = objBudgetHeader.Guid Then
                        CurrentBudgetItem.RTF = objBudgetHeader.RTF
                        CurrentBudgetItem.Type = BudgetItem.BudgetItemTypes.Category
                    Else
                        Dim NewBudgetHeader As New BudgetItem(objBudgetHeader.RTF, objBudgetHeader.Guid)
                        Me.BudgetItems.Insert(intIndex, NewBudgetHeader)
                    End If
                Else
                    Dim NewBudgetHeader As New BudgetItem(objBudgetHeader.RTF, objBudgetHeader.Guid)
                    Me.BudgetItems.Insert(intIndex, NewBudgetHeader)
                End If
            End If
        ElseIf objBudgetHeader.ParentBudgetItemGuid <> Guid.Empty Then
            Dim ParentBudgetItem As BudgetItem = CurrentLogFrame.GetBudgetItemByGuid(objBudgetHeader.ParentBudgetItemGuid)

            If ParentBudgetItem IsNot Nothing Then
                Dim ReferenceParent As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(ParentBudgetItem.Guid)

                If ReferenceParent IsNot Nothing Then
                    intIndex = ParentBudgetItem.BudgetItems.IndexOf(objBudgetHeader)

                    If intIndex >= 0 And intIndex <= ReferenceParent.BudgetItems.Count - 1 Then
                        Dim CurrentBudgetItem As BudgetItem = ReferenceParent.BudgetItems(intIndex)

                        If CurrentBudgetItem IsNot Nothing AndAlso CurrentBudgetItem.ReferenceBudgetItemGuid = objBudgetHeader.Guid Then
                            CurrentBudgetItem.RTF = objBudgetHeader.RTF
                            CurrentBudgetItem.Type = BudgetItem.BudgetItemTypes.Category
                        Else
                            Dim NewBudgetHeader As New BudgetItem(objBudgetHeader.RTF, objBudgetHeader.Guid)
                            ReferenceParent.BudgetItems.Insert(intIndex, NewBudgetHeader)
                        End If
                    Else
                        Dim NewBudgetHeader As New BudgetItem(objBudgetHeader.RTF, objBudgetHeader.Guid)
                        ReferenceParent.BudgetItems.Insert(intIndex, NewBudgetHeader)
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub RemoveBudgetHeader(ByVal objReferenceGuid As Guid)
        Dim ReferringBudgetItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objReferenceGuid)

        If ReferringBudgetItem IsNot Nothing Then
            Dim objParent As BudgetItems = CurrentLogFrame.GetParentCollection_LogframeObject(ReferringBudgetItem)

            If objParent IsNot Nothing Then
                objParent.Remove(ReferringBudgetItem)

                UpdateParentTotals(ReferringBudgetItem)
            End If
        End If
    End Sub

    Public Sub MoveBudgetHeader(ByVal objBudgetHeader As BudgetItem)
        Dim ReferenceItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objBudgetHeader.Guid)
        Dim intIndex As Integer

        If ReferenceItem Is Nothing Then Exit Sub

        Dim ParentCollection As BudgetItems = CurrentLogFrame.GetParentCollection(ReferenceItem)
        ParentCollection.Remove(ReferenceItem)

        UpdateParentTotals(ReferenceItem)

        If objBudgetHeader.ParentBudgetYearGuid <> Guid.Empty Then
            Dim ParentBudgetYear As BudgetYear = CurrentLogFrame.GetBudgetYearByGuid(objBudgetHeader.ParentBudgetYearGuid)

            If ParentBudgetYear IsNot Nothing Then
                intIndex = ParentBudgetYear.BudgetItems.IndexOf(objBudgetHeader)
                Me.BudgetItems.Insert(intIndex, ReferenceItem)

                UpdateParentTotals(ReferenceItem)
            End If
        ElseIf objBudgetHeader.ParentBudgetItemGuid <> Guid.Empty Then
            Dim ParentReferenceBudgetItem As BudgetItem = CurrentLogFrame.GetBudgetItemByGuid(objBudgetHeader.ParentBudgetItemGuid)
            Dim ParentBudgetItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(ParentReferenceBudgetItem.Guid)

            If ParentReferenceBudgetItem IsNot Nothing Then
                intIndex = ParentReferenceBudgetItem.BudgetItems.IndexOf(objBudgetHeader)
                ParentBudgetItem.BudgetItems.Insert(intIndex, ReferenceItem)

                UpdateParentTotals(ReferenceItem)
            End If
        End If
    End Sub

    Public Sub InsertParentHeader(ByVal objChildItem As BudgetItem)
        Dim ReferenceItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objChildItem.Guid)
        Dim intIndex As Integer

        If ReferenceItem Is Nothing Then Exit Sub

        Dim ParentCollection As BudgetItems = CurrentLogFrame.GetParentCollection(ReferenceItem)
        intIndex = ParentCollection.IndexOf(ReferenceItem)
        ParentCollection.Remove(ReferenceItem)

        Dim NewParent As New BudgetItem
        NewParent.BudgetItems.Add(ReferenceItem)
        NewParent.ReferenceBudgetItemGuid = objChildItem.ParentBudgetItemGuid
        NewParent.SetParentTotals()

        ParentCollection.Insert(intIndex, NewParent)
    End Sub

    Public Sub InsertChildHeader(ByVal objChildItem As BudgetItem)
        Dim ParentItem As BudgetItem = CurrentLogFrame.GetParent(objChildItem)
        Dim ParentReferenceItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(ParentItem.Guid)

        Dim NewChild As New BudgetItem

        With NewChild
            .RTF = objChildItem.RTF
            .ReferenceBudgetItemGuid = objChildItem.Guid
        End With
        ParentReferenceItem.BudgetItems.Insert(0, NewChild)
    End Sub

    Public Sub LevelDownHeader(ByVal objChildItem As BudgetItem)
        Dim ReferenceItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objChildItem.Guid)

        If ReferenceItem Is Nothing Then Exit Sub

        Dim ParentItem As BudgetItem = CurrentLogFrame.GetParent(objChildItem)
        Dim ParentReferenceItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(ParentItem.Guid)
        Dim ParentCollection As BudgetItems = CurrentLogFrame.GetParentCollection(ReferenceItem)
        Dim intIndex As Integer = ParentItem.BudgetItems.IndexOf(objChildItem)

        ParentCollection.Remove(ReferenceItem)
        ParentReferenceItem.BudgetItems.Insert(intIndex, ReferenceItem)
        ParentReferenceItem.SetParentTotals()
    End Sub

    Public Sub LevelUpHeader(ByVal objBudgetHeader As BudgetItem)
        Dim ReferenceItem As BudgetItem = Me.BudgetItems.GetBudgetItemByReferenceGuid(objBudgetHeader.Guid)

        If ReferenceItem Is Nothing Then Exit Sub

        Dim ParentItem As BudgetItem = CurrentLogFrame.GetParent(ReferenceItem)
        Dim ParentCollection As BudgetItems = CurrentLogFrame.GetParentCollection(ReferenceItem)
        Dim TargetItem As BudgetItem = CurrentLogFrame.GetParent(ParentItem)
        Dim intIndex As Integer = TargetItem.BudgetItems.IndexOf(ParentItem)

        ParentItem.BudgetItems.Remove(ReferenceItem)
        ParentItem.SetParentTotals()

        intIndex += 1

        If intIndex > TargetItem.BudgetItems.Count - 1 Then
            TargetItem.BudgetItems.Add(ReferenceItem)
        Else
            TargetItem.BudgetItems.Insert(intIndex, ReferenceItem)
        End If

        TargetItem.SetParentTotals()
    End Sub
#End Region

#Region "Events"
    Private Sub BudgetItems_BudgetItemAdded(ByVal sender As Object, ByVal e As BudgetItemAddedEventArgs) Handles BudgetItems.BudgetItemAdded
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        selBudgetItem.idBudgetYear = Me.idBudgetYear
        selBudgetItem.ParentBudgetYearGuid = Me.Guid
    End Sub
#End Region


End Class

Public Class BudgetYears
    Inherits System.Collections.CollectionBase

    Public Event BudgetYearAdded(ByVal sender As Object, ByVal e As BudgetYearAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal budgetyear As BudgetYear)
        List.Add(budgetyear)

        RaiseEvent BudgetYearAdded(Me, New BudgetYearAddedEventArgs(budgetyear))
    End Sub

    Public Sub AddRange(ByVal budgetyears As BudgetYears)
        InnerList.AddRange(budgetyears)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal budgetyear As BudgetYear)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, budgetyear.ItemName))
        ElseIf index = Count Then
            List.Add(budgetyear)

            RaiseEvent BudgetYearAdded(Me, New BudgetYearAddedEventArgs(budgetyear))
        Else
            List.Insert(index, budgetyear)

            RaiseEvent BudgetYearAdded(Me, New BudgetYearAddedEventArgs(budgetyear))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, BudgetYear.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal budgetyear As BudgetYear)
        If List.Contains(budgetyear) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, budgetyear.ItemName))
        Else
            List.Remove(budgetyear)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As BudgetYear
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), BudgetYear)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal budgetyear As BudgetYear) As Integer
        Return List.IndexOf(budgetyear)
    End Function

    Public Function Contains(ByVal budgetyear As BudgetYear) As Boolean
        Return List.Contains(budgetyear)
    End Function

    Public Function GetBudgetYearByGuid(ByVal objGuid As Guid) As BudgetYear
        Dim selBudgetYear As BudgetYear = Nothing
        For Each objBudgetYear As BudgetYear In Me.List
            If objBudgetYear.Guid = objGuid Then
                selBudgetYear = objBudgetYear
                Exit For
            End If
        Next
        Return selBudgetYear
    End Function
End Class

Public Class BudgetYearAddedEventArgs
    Inherits EventArgs

    Public Property BudgetYear As BudgetYear

    Public Sub New(ByVal objBudgetYear As BudgetYear)
        MyBase.New()

        Me.BudgetYear = objBudgetYear
    End Sub
End Class