Partial Public Class frmParent
    Private Sub RibbonTabLayOutBudget_ActiveChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonTabLayOutBudget.ActiveChanged
        RibbonButtonBudgetShowDurationColumn.Checked = My.Settings.setShowDurationColumns
        RibbonButtonBudgetShowLocalCurrencyColumn.Checked = My.Settings.setShowLocalCurrencyColumn
    End Sub

#Region "Columns"
    Private Sub RibbonButtonBudgetShowDurationColumn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetShowDurationColumn.Click
        ShowDuration()
    End Sub

    Private Sub ShowDuration()
        Dim boolShow As Boolean

        If My.Settings.setShowDatesColumns = True Then boolShow = False Else boolShow = True

        My.Settings.setShowDatesColumns = boolShow
        CurrentProjectForm.dgvBudgetYear_SetVisibilityDurationColumns(boolShow)
        RibbonButtonBudgetShowDurationColumn.Checked = boolShow
    End Sub


    Private Sub RibbonButtonBudgetShowLocalCurrencyColumn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetShowLocalCurrencyColumn.Click
        ShowLocalCurrency()
    End Sub

    Private Sub ShowLocalCurrency()
        Dim boolShow As Boolean

        If My.Settings.setShowLocalCurrencyColumn = True Then boolShow = False Else boolShow = True

        My.Settings.setShowLocalCurrencyColumn = boolShow
        CurrentProjectForm.dgvBudgetYear_SetVisibilityLocalCurrencyColumn(boolShow)
        RibbonButtonBudgetShowLocalCurrencyColumn.Checked = boolShow
    End Sub
#End Region

#Region "View"
    Private Sub RibbonButtonBudgetExchangeRates_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetExchangeRates.Click
        ExchangeRatesShowHide()
    End Sub

    Private Sub RibbonButtonBudgetDetailsWindow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetDetailsWindow.Click
        RibbonButtonBudgetDetailsWindow.Text = CurrentProjectForm.ShowDetailsBudget()
    End Sub

    Private Sub RibbonButtonBudgetEmptyCells_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetEmptyCells.Click
        HideEmptyCellsBudgetYear()
    End Sub

    Private Sub HideEmptyCellsBudgetYear()
        If My.Settings.setBudgetHideEmptyCells = False Then
            RibbonButtonBudgetEmptyCells.Text = LANG_EmptyCellsShow
            My.Settings.setBudgetHideEmptyCells = True

        Else
            RibbonButtonBudgetEmptyCells.Text = LANG_EmptyCellsHide
            My.Settings.setBudgetHideEmptyCells = False
        End If

        CurrentProjectForm.dgvBudgetYear_SetVisibilityEmptyCells(My.Settings.setBudgetHideEmptyCells)
    End Sub

    Public Sub ExchangeRatesShowHide()
        For Each selForm As frmProject In Me.MdiChildren
            selForm.ShowExchangeRates()
        Next
        If My.Settings.setShowExchangeRates = True Then
            My.Settings.setShowExchangeRates = False
        Else
            My.Settings.setShowExchangeRates = True
        End If
    End Sub
#End Region

#Region "Templates - events and methods"
    Private Sub SetBudgetButtons()
        RibbonButtonBudgetShowDurationColumn.Checked = My.Settings.setShowDurationColumns
        RibbonButtonBudgetShowLocalCurrencyColumn.Checked = My.Settings.setShowLocalCurrencyColumn

        If CurrentLogFrame.Budget.MultiYearBudget = True Then
            RibbonButtonMultiYearBudget.Checked = True
            RibbonButtonSimpleBudget.Checked = False
        Else
            RibbonButtonMultiYearBudget.Checked = False
            RibbonButtonSimpleBudget.Checked = True
        End If

        For Each selForm As frmProject In Me.MdiChildren
            selForm.ShowExchangeRates()
        Next

    End Sub

    Private Sub RibbonButtonSimpleBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSimpleBudget.Click
        If RibbonButtonSimpleBudget.Checked = True Then
            CurrentLogFrame.Budget.MultiYearBudget = False
            RibbonButtonMultiYearBudget.Checked = False
        Else
            CurrentLogFrame.Budget.MultiYearBudget = True
            RibbonButtonMultiYearBudget.Checked = True
        End If
        CurrentLogFrame.Budget.UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonMultiYearBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonMultiYearBudget.Click
        If RibbonButtonMultiYearBudget.Checked = True Then
            CurrentLogFrame.Budget.MultiYearBudget = True
            RibbonButtonSimpleBudget.Checked = False
        Else
            CurrentLogFrame.Budget.MultiYearBudget = False
            RibbonButtonSimpleBudget.Checked = True
        End If
        CurrentLogFrame.Budget.UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub BudgetTemplates_AddStaff(ByVal objParentItem As BudgetItem, ByRef intStaffIndex As Integer)
        Dim strItem As String = String.Format("[{0} {1}]", LANG_Function, intStaffIndex)
        Dim objBudgetItem As New BudgetItem(strItem, True)

        objParentItem.BudgetItems.Add(objBudgetItem)

        intStaffIndex += 1
    End Sub

    Private Sub BudgetTemplates_AddItem(ByVal objParentItem As BudgetItem, ByRef intItemIndex As Integer)
        Dim strItem As String = String.Format("[{0} {1}]", LANG_Item, intItemIndex)
        Dim objBudgetItem As New BudgetItem(strItem, True)

        objParentItem.BudgetItems.Add(objBudgetItem)

        intItemIndex += 1
    End Sub

    Private Sub RibbonButtonBudgetTemplateDcAc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetTemplateDcAc.Click
        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetStructureDirectAdminCosts_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonBudgetStructureDirectAdminCosts.Click
        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetTemplateHumanitarian1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetTemplateHumanitarian1.Click
        BudgetTemplate_HumanitarianECHO()

        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetTemplateHumanitarian2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetTemplateHumanitarian2.Click
        BudgetTemplate_HumanitarianDFID()

        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetTemplateHumanitarian3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetTemplateHumanitarian3.Click
        BudgetTemplate_HumanitarianCIDA()

        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetStructureHumanitarian1_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonBudgetStructureHumanitarian1.Click
        BudgetTemplate_HumanitarianECHO()

        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetStructureHumanitarian2_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonBudgetStructureHumanitarian2.Click
        BudgetTemplate_HumanitarianDFID()

        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetStructureHumanitarian3_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonBudgetStructureHumanitarian3.Click
        BudgetTemplate_HumanitarianCIDA()

        BudgetTemplate_AdministrativeDirectCost()

        CurrentProjectForm.ProjectInitialise_Budget()
    End Sub

    Private Sub RibbonButtonBudgetTemplateDevelopment1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonBudgetTemplateDevelopment1.Click
        CurrentBudget.MultiYearBudget = True
        BudgetTemplate_DevelopmentEuropeAid()

        CurrentBudget.UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)
        BudgetTemplate_AdministrativeDirectCost()
        BudgetTemplate_DevelopmentEuropeAid_Finalise()

        CurrentProjectForm.ProjectInitialise_Budget()
        SetBudgetButtons()
    End Sub

    Private Sub RibbonButtonBudgetStructureDevelopment1_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonBudgetStructureDevelopment1.Click
        CurrentBudget.MultiYearBudget = True
        BudgetTemplate_DevelopmentEuropeAid()

        CurrentBudget.UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)
        BudgetTemplate_AdministrativeDirectCost()
        BudgetTemplate_DevelopmentEuropeAid_Finalise()

        CurrentProjectForm.ProjectInitialise_Budget()
        SetBudgetButtons()
    End Sub
#End Region

#Region "Templates - admin/direct cost"
    Private Sub BudgetTemplate_AdministrativeDirectCost()
        Dim strAdminCosts As String = String.Format("{0} (7%)", LANG_AdministrativeCosts)
        Dim TotalBudget As BudgetYear = CurrentBudget.BudgetYears(0)
        Dim objDirectCost As New BudgetItem(LANG_DirectCosts, True)
        Dim objIndirectCost As New BudgetItem(strAdminCosts, True)
        Dim objIndirectCostHeader As New BudgetItem(LANG_AdministrativeCosts, True)


        'Total budget

        If TotalBudget.BudgetItems.Count > 0 Then
            For Each selBudgetItem As BudgetItem In TotalBudget.BudgetItems
                objDirectCost.BudgetItems.Add(selBudgetItem)
            Next
            TotalBudget.BudgetItems.Clear()
        Else
            Dim objChildItem As New BudgetItem(LANG_AddProjectCostsHere, True)

            objDirectCost.BudgetItems.Add(objChildItem)
        End If
        objDirectCost.SetParentTotals()
        TotalBudget.BudgetItems.Insert(0, objDirectCost)

        With objIndirectCost
            .Type = BudgetItem.BudgetItemTypes.Ratio
            .CalculationGuidRatio = objDirectCost.Guid
            .Ratio = 0.07
            .CalculateRatio()
        End With
        objIndirectCostHeader.BudgetItems.Add(objIndirectCost)
        objIndirectCostHeader.SetParentTotals()

        TotalBudget.BudgetItems.Add(objIndirectCostHeader)
        TotalBudget.TotalCost = TotalBudget.GetTotalCost

        'Budget years if budget is multi-year

        If CurrentBudget.MultiYearBudget = True Then
            For i = 1 To CurrentBudget.BudgetYears.Count - 1
                Dim selBudgetYear As BudgetYear = CurrentBudget.BudgetYears(i)
                Dim NewBudgetHeader As New BudgetItem(objDirectCost.RTF, objDirectCost.Guid)

                If selBudgetYear.BudgetItems.Count > 0 Then
                    For Each selBudgetItem As BudgetItem In selBudgetYear.BudgetItems
                        NewBudgetHeader.BudgetItems.Add(selBudgetItem)
                    Next
                    selBudgetYear.BudgetItems.Clear()
                End If

                NewBudgetHeader.SetParentTotals()
                selBudgetYear.BudgetItems.Insert(0, NewBudgetHeader)

                objIndirectCostHeader = New BudgetItem(LANG_AdministrativeCosts, True)
                objIndirectCost = New BudgetItem(strAdminCosts, True)
                With objIndirectCost
                    .Type = BudgetItem.BudgetItemTypes.Ratio
                    .CalculationGuidRatio = NewBudgetHeader.Guid
                    .Ratio = 0.07
                    .CalculateRatio()
                End With
                objIndirectCostHeader.BudgetItems.Add(objIndirectCost)
                objIndirectCostHeader.SetParentTotals()
                selBudgetYear.BudgetItems.Add(objIndirectCostHeader)

                selBudgetYear.TotalCost = selBudgetYear.GetTotalCost
            Next
        End If
    End Sub
#End Region

#Region "Templates - humanitarian"
    Private Sub BudgetTemplate_HumanitarianECHO()
        Dim TotalBudget As BudgetYear = CurrentBudget.BudgetYears(0)
        Dim objMainItem As BudgetItem
        Dim objBudgetItem As BudgetItem
        Dim objIndirectCost As New BudgetItem()
        Dim intItemIndex As Integer = 1
        Dim intStaffIndex As Integer = 1

        objMainItem = New BudgetItem(LANG_Equipment, True)
        BudgetTemplates_AddItem(objMainItem, intItemIndex)
        TotalBudget.BudgetItems.Insert(0, objMainItem)

        objMainItem = New BudgetItem(LANG_HumanResources, True)
        TotalBudget.BudgetItems.Insert(1, objMainItem)

        objBudgetItem = New BudgetItem(LANG_StaffLocal, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffExpatriate, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_RunningCosts, True)
        TotalBudget.BudgetItems.Insert(2, objMainItem)

        objBudgetItem = New BudgetItem(LANG_VehicleCosts, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Travel, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_CommunicationVisibility, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Buildings, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_SuppliesMaterials, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_ExternalServices, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_OtherOperationalCosts, True)
        BudgetTemplates_AddItem(objMainItem, intItemIndex)
        TotalBudget.BudgetItems.Insert(3, objMainItem)
    End Sub

    Private Sub BudgetTemplate_HumanitarianDFID()
        Dim TotalBudget As BudgetYear = CurrentBudget.BudgetYears(0)
        Dim objMainItem As BudgetItem
        Dim objBudgetItem, objSubBudgetItem As BudgetItem
        Dim objIndirectCost As New BudgetItem()
        Dim intItemIndex As Integer = 1
        Dim intStaffIndex As Integer = 1

        objMainItem = New BudgetItem(LANG_Inputs, True)
        TotalBudget.BudgetItems.Insert(0, objMainItem)

        objBudgetItem = New BudgetItem(LANG_Health, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_WASH, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Food, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Livelihoods, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Shelter, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_NonFoodItems, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Education, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_CashTransfers, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Other, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_Transport, True)
        TotalBudget.BudgetItems.Insert(1, objMainItem)

        objBudgetItem = New BudgetItem(LANG_TransportMaterials, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_InternationalShipping, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_PortHandling, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_TruckRental, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffTravel, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_TravelInternational, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_CarRental, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objMainItem = New BudgetItem(LANG_Security, True)
        TotalBudget.BudgetItems.Insert(2, objMainItem)

        objBudgetItem = New BudgetItem(LANG_Staff, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_Guards, True)
        BudgetTemplates_AddStaff(objSubBudgetItem, intStaffIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Equipment, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_LogisticsOverheads, True)
        TotalBudget.BudgetItems.Insert(3, objMainItem)

        objBudgetItem = New BudgetItem(LANG_CountryOffice, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_OfficeSpace, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_Communications, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_Storage, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_BankCharges, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_SmallEquipment, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_Insurance, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objBudgetItem = New BudgetItem(LANG_FieldOffice, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_OfficeSpace, True)
        BudgetTemplates_AddItem(objSubBudgetItem, intItemIndex)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objMainItem = New BudgetItem(LANG_StaffingSupport, True)
        TotalBudget.BudgetItems.Insert(4, objMainItem)

        objBudgetItem = New BudgetItem(LANG_StaffNationalTechnical, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffNationalSupport, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffInternationalTechnical, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffInternationalSupport, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffOffshoreTechnical, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_StaffOffshoreSupport, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Training, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_NonSalaryBenefits, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_MonitoringEvaluation, True)
        TotalBudget.BudgetItems.Insert(5, objMainItem)

        objBudgetItem = New BudgetItem(LANG_Monitoring, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Evaluation, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_CapitalItems, True)
        TotalBudget.BudgetItems.Insert(6, objMainItem)

        objBudgetItem = New BudgetItem(LANG_Communications, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)
    End Sub

    Private Sub BudgetTemplate_HumanitarianCIDA()
        Dim TotalBudget As BudgetYear = CurrentBudget.BudgetYears(0)
        Dim objMainItem As BudgetItem
        Dim objBudgetItem As BudgetItem
        Dim objIndirectCost As New BudgetItem()
        Dim intItemIndex As Integer = 1
        Dim intStaffIndex As Integer = 1

        objMainItem = New BudgetItem(LANG_PersonnelCosts, True)
        TotalBudget.BudgetItems.Insert(0, objMainItem)

        objBudgetItem = New BudgetItem(LANG_HeadOffice, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_FieldOffice, True)
        BudgetTemplates_AddStaff(objBudgetItem, intStaffIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_SuppliesMaterials, True)
        BudgetTemplates_AddItem(objMainItem, intItemIndex)
        TotalBudget.BudgetItems.Insert(1, objMainItem)

        objMainItem = New BudgetItem(LANG_Logistics, True)
        TotalBudget.BudgetItems.Insert(2, objMainItem)

        objBudgetItem = New BudgetItem(LANG_Transport, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Storage, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_DistributionCosts, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_AdministrativeCostsLocal, True)
        BudgetTemplates_AddItem(objMainItem, intItemIndex)
        TotalBudget.BudgetItems.Insert(3, objMainItem)

        objMainItem = New BudgetItem(LANG_TrainingCapacity, True)
        BudgetTemplates_AddItem(objMainItem, intItemIndex)
        TotalBudget.BudgetItems.Insert(4, objMainItem)

        objMainItem = New BudgetItem(LANG_AssessmentMonitoringEvaluation, True)
        BudgetTemplates_AddItem(objMainItem, intItemIndex)
        TotalBudget.BudgetItems.Insert(5, objMainItem)

        objMainItem = New BudgetItem(LANG_SafetySecurity, True)
        TotalBudget.BudgetItems.Insert(6, objMainItem)

        objBudgetItem = New BudgetItem(LANG_MaterialResources, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_HumanResources, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Training, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_SiteEnhancements, True)
        BudgetTemplates_AddItem(objBudgetItem, intItemIndex)
        objMainItem.BudgetItems.Add(objBudgetItem)
    End Sub
#End Region

#Region "Templates - development"
    Private Sub BudgetTemplate_DevelopmentEuropeAid()
        Dim TotalBudget As BudgetYear = CurrentBudget.BudgetYears(0)
        Dim objMainItem As BudgetItem
        Dim objBudgetItem, objSubBudgetItem As BudgetItem
        Dim intItemIndex As Integer = 1
        Dim intStaffIndex As Integer = 1

        objMainItem = New BudgetItem(LANG_HumanResources, True)
        TotalBudget.BudgetItems.Insert(0, objMainItem)

        objBudgetItem = New BudgetItem(LANG_SalariesLocalStaff, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_TechnicalStaff, True)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_SupportStaff, True)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objBudgetItem = New BudgetItem(LANG_SalariesExpats, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_PerDiems, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_Abroad, True)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_Local, True)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objSubBudgetItem = New BudgetItem(LANG_SeminarParticipants, True)
        objBudgetItem.BudgetItems.Add(objSubBudgetItem)

        objMainItem = New BudgetItem(LANG_Travel, True)
        TotalBudget.BudgetItems.Insert(1, objMainItem)

        objBudgetItem = New BudgetItem(LANG_TravelInternational, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_TravelLocal, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_EquipmentSupplies, True)
        TotalBudget.BudgetItems.Insert(2, objMainItem)

        objBudgetItem = New BudgetItem(LANG_PurchaseRentVehicles, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_FurnitureComputers, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_MachinesTools, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_SparePartsTools, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Other, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_LocalOffice, True)
        TotalBudget.BudgetItems.Insert(3, objMainItem)

        objBudgetItem = New BudgetItem(LANG_VehicleCosts, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_OfficeRent, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Consumables, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_OtherServices, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_OtherCostsServices, True)
        TotalBudget.BudgetItems.Insert(4, objMainItem)

        objBudgetItem = New BudgetItem(LANG_Publications, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Studies, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Audit, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_EvaluationCosts, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_Translation, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_FinancialServices, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_ConferencesSeminars, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objBudgetItem = New BudgetItem(LANG_VisibilityActions, True)
        objMainItem.BudgetItems.Add(objBudgetItem)

        objMainItem = New BudgetItem(LANG_Other, True)
        TotalBudget.BudgetItems.Insert(5, objMainItem)
    End Sub

    Private Sub BudgetTemplate_DevelopmentEuropeAid_Finalise()
        Dim TotalBudget As BudgetYear = CurrentBudget.BudgetYears(0)
        Dim objMainItem As BudgetItem

        objMainItem = New BudgetItem(LANG_ContingencyReserve, True)
        TotalBudget.BudgetItems.Add(objMainItem)

        CurrentBudget.InsertBudgetHeader(objMainItem)

        objMainItem = New BudgetItem(LANG_TaxesContributions, True)
        TotalBudget.BudgetItems.Add(objMainItem)

        CurrentBudget.InsertBudgetHeader(objMainItem)
    End Sub
#End Region
End Class
