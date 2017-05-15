Partial Public Class frmParent
#Region "Print preview buttons"
    Private Sub RibbonButtonPreviewLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewLogframe.Click
        Preview_Logframe()
    End Sub

    Private Sub RibbonButtonPreviewPlanning_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewPlanning.Click
        Preview_Planning()
    End Sub

    Private Sub RibbonButtonPreviewBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewBudget.Click
        Preview_Budget()
    End Sub

    Private Sub RibbonButtonPreviewPMF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewPMF.Click
        Preview_PMF()
    End Sub

    Private Sub RibbonButtonPreviewIndicators_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewIndicators.Click
        Preview_Indicators()
    End Sub

    Private Sub RibbonButtonPreviewRiskRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewRiskRegister.Click
        Preview_RiskRegister()
    End Sub

    Private Sub RibbonButtonPreviewIdForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPreviewIdForm.Click
        Preview_TargetGroupIdForm()
    End Sub
#End Region

#Region "Drop-down buttons"
    Private Sub DropDownButtonPreviewLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewLogframe.Click
        Preview_Logframe()
    End Sub

    Private Sub DropDownButtonPreviewPlanning_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewPlanning.Click
        Preview_Planning()
    End Sub

    Private Sub DropDownButtonPreviewBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewBudget.Click
        Preview_Budget()
    End Sub

    Private Sub DropDownButtonPreviewPMF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewPMF.Click
        Preview_PMF()
    End Sub

    Private Sub DropDownButtonPreviewIndicators_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewIndicators.Click
        Preview_Indicators()
    End Sub

    Private Sub DropDownButtonPreviewResources_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewResources.Click
        Preview_Resources()
    End Sub

    Private Sub DropDownButtonPreviewRiskRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewRiskRegister.Click
        Preview_RiskRegister()
    End Sub

    Private Sub DropDownButtonPreviewAssumptionsTable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewAssumptionsTable.Click
        Preview_Assumptions()
    End Sub

    Private Sub DropDownButtonPreviewDependenciesTable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewDependenciesTable.Click
        Preview_Dependencies()
    End Sub

    Private Sub DropDownButtonPreviewPartnerList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewPartnerList.Click
        Preview_PartnerList()
    End Sub

    Private Sub DropDownButtonPreviewTargetGroupId_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownButtonPreviewTargetGroupId.Click
        Preview_TargetGroupIdForm()
    End Sub
#End Region

#Region "Methods"
    Private Sub Preview_Logframe()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageLogframe")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_Planning()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPagePlanning")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_Budget()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageBudget")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_PMF()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPagePmf")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_Indicators()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageIndicators")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_Resources()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageResources")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_RiskRegister()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageRiskRegister")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_Assumptions()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageAssumptions")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_Dependencies()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageDependencies")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_PartnerList()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPagePartnerList")

        dialogPrintPreview.ShowDialog(Me)
    End Sub

    Private Sub Preview_TargetGroupIdForm()
        Dim dialogPrintPreview As New DialogPrintPreview

        dialogPrintPreview.TabControlReports.SelectTab("TabPageTargetGroupIdForm")

        dialogPrintPreview.ShowDialog(Me)
    End Sub
#End Region

    
    
End Class
