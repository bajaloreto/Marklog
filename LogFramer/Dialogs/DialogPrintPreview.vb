Imports System.Windows.Forms

Public Class DialogPrintPreview
    Friend WithEvents PrintDetailLogframe As New ucPrintDetail
    Friend WithEvents PrintDetailPlanning As New ucPrintDetail
    Friend WithEvents PrintDetailBudget As New ucPrintDetail
    Friend WithEvents PrintDetailPmf As New ucPrintDetail
    Friend WithEvents PrintDetailIndicators As New ucPrintDetail
    Friend WithEvents PrintDetailResources As New ucPrintDetail
    Friend WithEvents PrintDetailRiskRegister As New ucPrintDetail
    Friend WithEvents PrintDetailAssumptions As New ucPrintDetail
    Friend WithEvents PrintDetailDependencies As New ucPrintDetail
    Friend WithEvents PrintDetailPartnerList As New ucPrintDetail
    Friend WithEvents PrintDetailTargetGroupIdForm As New ucPrintDetail

    Friend WithEvents PrintSettingsLogframe As New FaciliDev.LogFramer.PrintSettingsLogframe
    Friend WithEvents PrintSettingsPlanning As New FaciliDev.LogFramer.PrintSettingsPlanning
    Friend WithEvents PrintSettingsBudget As New FaciliDev.LogFramer.PrintSettingsBudget
    Friend WithEvents PrintSettingsPmf As New FaciliDev.LogFramer.PrintSettingsPmf
    Friend WithEvents PrintSettingsIndicators As New FaciliDev.LogFramer.PrintSettingsIndicators
    Friend WithEvents PrintSettingsResources As New FaciliDev.LogFramer.ucPrintSettingsResources
    Friend WithEvents PrintSettingsRiskRegister As New FaciliDev.LogFramer.ucPrintSettingsRiskRegister
    Friend WithEvents PrintSettingsAssumptions As New FaciliDev.LogFramer.ucPrintSettingsAssumptions
    Friend WithEvents PrintSettingsDependencies As New FaciliDev.LogFramer.ucPrintSettingsDependencies
    Friend WithEvents PrintSettingsPartnerList As New FaciliDev.LogFramer.ucPrintSettingsPartnerList
    Friend WithEvents PrintSettingsTargetGroupIdForm As New FaciliDev.LogFramer.ucPrintSettingsTargetGroupIdForm

    Private boolBusy As Boolean

    Private Property Busy As Boolean
        Get
            Return boolBusy
        End Get
        Set(ByVal value As Boolean)
            boolBusy = value
            If boolBusy = True Then btnClose.Enabled = False Else btnClose.Enabled = True
        End Set
    End Property

    Public Sub New()

        InitializeComponent()
        Me.SuspendLayout()

        If CurrentLogFrame IsNot Nothing Then
            Dim ptPrintSettings As New Point(3, 415)
            'Logframe
            With PrintDetailLogframe
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            With PrintSettingsLogframe
                .Location = ptPrintSettings
                .ShowIndicatorColumn = CurrentProjectForm.dgvLogframe.ShowIndicatorColumn
                .ShowVerificationSourceColumn = CurrentProjectForm.dgvLogframe.ShowVerificationSourceColumn
                .ShowAssumptionColumn = CurrentProjectForm.dgvLogframe.ShowAssumptionColumn
                .ShowGoals = CurrentProjectForm.dgvLogframe.ShowGoals
                .ShowPurposes = CurrentProjectForm.dgvLogframe.ShowPurposes
                .ShowOutputs = CurrentProjectForm.dgvLogframe.ShowOutputs
                .ShowActivities = CurrentProjectForm.dgvLogframe.ShowActivities
                .ShowResourcesBudget = CurrentProjectForm.dgvLogframe.ShowResourcesBudget
                .Invalidate()
            End With
            With PrintDetailLogframe.PrintSettingsBar
                .Height += PrintSettingsLogframe.Height
                .Controls.Add(PrintSettingsLogframe)
            End With

            'Planning
            With PrintDetailPlanning
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsPlanning.Location = ptPrintSettings

            With PrintDetailPlanning.PrintSettingsBar
                .Height += PrintSettingsPlanning.Height
                .Controls.Add(PrintSettingsPlanning)
            End With

            'Budget
            With PrintDetailBudget
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsBudget.Location = ptPrintSettings

            With PrintDetailBudget.PrintSettingsBar
                .Height += PrintSettingsBudget.Height
                .Controls.Add(PrintSettingsBudget)
            End With

            'PMF
            With PrintDetailPmf
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsPmf.Location = ptPrintSettings

            With PrintDetailPmf.PrintSettingsBar
                .Height += PrintSettingsPmf.Height
                .Controls.Add(PrintSettingsPmf)
            End With

            'Indicators / questionnaires
            With PrintDetailIndicators
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsIndicators.Location = ptPrintSettings

            With PrintDetailIndicators.PrintSettingsBar
                .Height += PrintSettingsIndicators.Height
                .Controls.Add(PrintSettingsIndicators)
            End With

            'Resources
            With PrintDetailResources
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsResources.Location = ptPrintSettings

            With PrintDetailResources.PrintSettingsBar
                .Height += PrintSettingsResources.Height
                .Controls.Add(PrintSettingsResources)
            End With

            'Risk register
            With PrintDetailRiskRegister
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsRiskRegister.Location = ptPrintSettings

            With PrintDetailRiskRegister.PrintSettingsBar
                .Height += PrintSettingsRiskRegister.Height
                .Controls.Add(PrintSettingsRiskRegister)
            End With

            'Assumptions
            With PrintDetailAssumptions
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsAssumptions.Location = ptPrintSettings

            With PrintDetailAssumptions.PrintSettingsBar
                .Height += PrintSettingsAssumptions.Height
                .Controls.Add(PrintSettingsAssumptions)
            End With

            'Dependencies
            With PrintDetailDependencies
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsDependencies.Location = ptPrintSettings

            With PrintDetailDependencies.PrintSettingsBar
                .Height += PrintSettingsDependencies.Height
                .Controls.Add(PrintSettingsDependencies)
            End With

            'Partner list
            With PrintDetailPartnerList
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsPartnerList.Location = ptPrintSettings

            With PrintDetailPartnerList.PrintSettingsBar
                .Height += PrintSettingsPartnerList.Height
                .Controls.Add(PrintSettingsPartnerList)
            End With

            'Target group identification form
            With PrintDetailTargetGroupIdForm
                .Location = New Point(3, 3)
                .Dock = DockStyle.Fill
                .PrintSettingsBar.ProgressBarDocument.Visible = False
            End With
            PrintSettingsTargetGroupIdForm.Location = ptPrintSettings

            With PrintDetailTargetGroupIdForm.PrintSettingsBar
                .Height += PrintSettingsTargetGroupIdForm.Height
                .Controls.Add(PrintSettingsTargetGroupIdForm)
            End With

            TabPageLogFrame.Controls.Add(PrintDetailLogframe)
            TabPagePlanning.Controls.Add(PrintDetailPlanning)
            TabPageBudget.Controls.Add(PrintDetailBudget)
            TabPagePmf.Controls.Add(PrintDetailPmf)
            TabPageIndicators.Controls.Add(PrintDetailIndicators)
            TabPageResources.Controls.Add(PrintDetailResources)
            TabPageRiskRegister.Controls.Add(PrintDetailRiskRegister)
            TabPageAssumptions.Controls.Add(PrintDetailAssumptions)
            TabPageDependencies.Controls.Add(PrintDetailDependencies)
            TabPagePartnerList.Controls.Add(PrintDetailPartnerList)
            TabPageTargetGroupIdForm.Controls.Add(PrintDetailTargetGroupIdForm)

            Me.ResumeLayout()
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub DialogPrintPreview_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Me.Refresh()

        ShowPrintDetail()
    End Sub

    Private Sub DialogPrintPreview_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        PrintDetailLogframe.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupLogframe
        PrintDetailPlanning.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupPlanning
        PrintDetailBudget.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupBudget
        PrintDetailPmf.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupPmf
        PrintDetailIndicators.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupIndicators
        PrintDetailResources.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupResources
        PrintDetailRiskRegister.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupRiskRegister
        PrintDetailAssumptions.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupAssumptions
        PrintDetailDependencies.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupDependencies
        PrintDetailPartnerList.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupPartnerList
        PrintDetailTargetGroupIdForm.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupTargetGroupIdForm
    End Sub

    Private Sub TabControlReports_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlReports.SelectedIndexChanged
        ShowPrintDetail()
    End Sub

    Private Sub ShowPrintDetail()
        With My.Settings
            Select Case TabControlReports.SelectedTab.Name
                Case TabPageLogFrame.Name
                    If PrintDetailLogframe.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailLogframe.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupLogframe

                        With CurrentProjectForm.dgvLogframe
                            LogframePreview(.StructRtfColumnWidth, .IndRtfColumnWidth, .VerRtfColumnWidth, .AsmRtfColumnWidth)
                        End With
                    End If
                Case TabPagePlanning.Name
                    If PrintDetailPlanning.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailPlanning.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupPlanning

                        If .setPrintPlanningPeriodView = PrintPlanning.PeriodViews.Day Then .setPrintPlanningPeriodView = PrintPlanning.PeriodViews.Month

                        PlanningPreview(Guid.Empty, .setPrintPlanningPeriodView, .setPrintPlanningElementsView, CurrentLogFrame.StartDate, CurrentLogFrame.EndDate, _
                                        .setPrintPlanningShowActivityLinks, .setPrintPlanningShowKeyMomentLinks, .setPrintPlanningRepeatRowHeaders, .setPrintPlanningShowDatesColumns)
                    End If
                Case TabPageBudget.Name
                    If PrintDetailBudget.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailBudget.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupBudget

                        BudgetPreview(0, .setPrintBudgetShowDurationColumns, .setPrintBudgetShowLocalCurrencyColumns, .setPrintBudgetShowExchangeRates, .setPrintBudgetPagesWide)

                    End If
                Case TabPagePmf.Name
                    If PrintDetailPmf.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailPmf.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupPmf

                        PmfPreview(.setPrintPmfSection, .setPrintPmfPagesWide, .setPrintPmfShowTargetRowTitles)
                    End If
                Case TabPageIndicators.Name
                    If PrintDetailIndicators.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailIndicators.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupIndicators
                        IndicatorsPreview(.setPrintIndicatorsSection, .setPrintIndicatorsTargetgroupGuid, .setPrintIndicatorsMeasurement, _
                                          .setPrintIndicatorsPrintPurposes, .setPrintIndicatorsPrintOutputs, .setPrintIndicatorsPrintActivities, _
                                          .setPrintIndicatorsPrintOptionValues, .setPrintIndicatorsPrintValueRanges, .setPrintIndicatorsPrintTargets)
                    End If
                Case TabPageResources.Name
                    If PrintDetailResources.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailResources.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupResources
                        ResourcesPreview(.setPrintResourcesPagesWide)
                    End If
                Case TabPageRiskRegister.Name
                    If PrintDetailRiskRegister.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailRiskRegister.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupRiskRegister
                        RiskRegisterPreview(.setPrintRiskRegisterRiskCategory, .setPrintRiskRegisterPagesWide)
                    End If
                Case TabPageAssumptions.Name
                    If PrintDetailAssumptions.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailAssumptions.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupAssumptions
                        AssumptionsPreview(.setPrintAssumptionsSection, .setPrintAssumptionsPagesWide)
                    End If
                Case TabPageDependencies.Name
                    If PrintDetailDependencies.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailDependencies.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupDependencies
                        DependenciesPreview(.setPrintDependenciesSection, .setPrintDependenciesShowDeliverables, .setPrintDependenciesPagesWide)
                    End If
                Case TabPagePartnerList.Name
                    If PrintDetailPartnerList.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailPartnerList.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupPartnerList
                        PartnerListPreview(My.Settings.setPrintPartnersIndex, My.Settings.setPrintPartnersBorders, My.Settings.setPrintPartnersFill)
                    End If
                Case TabPageTargetGroupIdForm.Name
                    If PrintDetailTargetGroupIdForm.CustomPrintPreview.Document Is Nothing Then
                        PrintDetailTargetGroupIdForm.PrintSettingsBar.ReportSetup = CurrentLogFrame.ReportSetupTargetGroupIdForm
                        TargetGroupIdFormPreview(My.Settings.setPrintTargetgroupName, My.Settings.setPrintTargetgroupIdBorders, My.Settings.setPrintTargetgroupIdFill)
                    End If

            End Select
        End With
    End Sub

#Region "Print logframe"
    Private Sub LogframePreview(ByVal intStructColumnwidth As Integer, ByVal intIndColumnWidth As Integer, ByVal intVerColumnWidth As Integer, ByVal intAsmColumnWidth As Integer)
        Dim boolShowIndColumn As Boolean = PrintSettingsLogframe.ShowIndicatorColumn
        Dim boolShowVerColumn As Boolean = PrintSettingsLogframe.ShowVerificationSourceColumn
        Dim boolShowAsmColumn As Boolean = PrintSettingsLogframe.ShowAssumptionColumn
        Dim boolShowGoals As Boolean = PrintSettingsLogframe.ShowGoals
        Dim boolShowPurposes As Boolean = PrintSettingsLogframe.ShowPurposes
        Dim boolShowOutputs As Boolean = PrintSettingsLogframe.ShowOutputs
        Dim boolShowActivities As Boolean = PrintSettingsLogframe.ShowActivities
        Dim boolShowResourcesBudget As Boolean = PrintSettingsLogframe.ShowResourcesBudget

        Dim prntLogFrame As New PrintLogFrame(CurrentLogFrame, intStructColumnwidth, intIndColumnWidth, intVerColumnWidth, intAsmColumnWidth, _
                                              boolShowIndColumn, boolShowVerColumn, boolShowAsmColumn, boolShowGoals, boolShowPurposes, boolShowOutputs, boolShowActivities, _
                                              boolShowResourcesBudget)
        AddHandler prntLogFrame.LinePrinted, AddressOf UpdateProgressBarLogframe
        PrintDetailLogframe.CustomPrintPreview.Document = prntLogFrame
    End Sub

    Private Sub UpdateProgressBarLogframe(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Busy = True
        With PrintDetailLogframe.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = e.RowCount - 1
            .Value = e.LineIndex
            If e.LineIndex = e.RowCount - 1 Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub ColumnSelection_ColumnSelectionChanged() Handles PrintSettingsLogframe.ColumnSelectionChanged
        InitialiseLogframePreview()
    End Sub

    Private Sub PrintSettingsLogframe_SectionSelectionChanged() Handles PrintSettingsLogframe.SectionSelectionChanged
        InitialiseLogframePreview()
    End Sub

    Private Sub PrintSettingsLogframe_ResourceBudgetChanged() Handles PrintSettingsLogframe.ResourceBudgetChanged
        InitialiseLogframePreview()
    End Sub

    Private Sub InitialiseLogframePreview()
        Dim intStructColumnwidth, intIndColumnWidth, intVerColumnWidth, intAsmColumnWidth As Integer
        With PrintSettingsLogframe
            .Refresh()
            intStructColumnwidth = 100
            If .ShowIndicatorColumn = True Then intIndColumnWidth = 100
            If .ShowVerificationSourceColumn = True Then intVerColumnWidth = 100
            If .ShowAssumptionColumn = True Then intAsmColumnWidth = 100
        End With
        LogframePreview(intStructColumnwidth, intIndColumnWidth, intVerColumnWidth, intAsmColumnWidth)
    End Sub
#End Region

#Region "Print Planning"
    Private Sub PlanningPreview(ByVal OutputGuid As Guid, ByVal PeriodView As Integer, ByVal PlanningElements As Integer, ByVal PeriodFrom As Date, ByVal PeriodUntil As Date, _
                                ByVal ShowActivityLinks As Boolean, ByVal ShowKeyMomentLinks As Boolean, ByVal RepeatRowHeaders As Boolean, ByVal ShowDateColumns As Boolean)
        Dim prntPlanning As New PrintPlanning(CurrentLogFrame, OutputGuid, PeriodView, PlanningElements, PeriodFrom, PeriodUntil, _
                                              ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDateColumns)

        AddHandler prntPlanning.PagePrinted, AddressOf UpdateProgressBarPlanning
        PrintDetailPlanning.CustomPrintPreview.Document = prntPlanning
    End Sub

    Private Sub UpdateProgressBarPlanning(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        If e.RowCount > 0 Then
            With PrintDetailPlanning.PrintSettingsBar.ProgressBarDocument
                .Visible = True
                .Maximum = intMax
                .Value = e.LineIndex
                If e.LineIndex = intMax Then
                    .Visible = False
                    .Value = 0
                    Busy = False
                End If
            End With
        End If
    End Sub

    Private Sub PrintSettingsPlanning_PrintPlanningSetupChanged(ByVal sender As Object, ByVal e As PrintPlanningSetupChangedEventArgs) Handles PrintSettingsPlanning.PrintPlanningSetupChanged
        PrintSettingsPlanning.Refresh()
        PlanningPreview(e.OutputGuid, e.PeriodView, e.PlanningElements, e.PeriodFrom, e.PeriodUntil, e.ShowActivityLinks, e.ShowKeyMomentLinks, e.RepeatRowHeaders, e.ShowDatesColumns)
    End Sub
#End Region

#Region "Print Budget"
    Private Sub BudgetPreview(ByVal BudgetYearIndex As Integer, ByVal ShowDurationColumns As Boolean, ByVal ShowLocalCurrencyColumns As Boolean, ByVal ShowExchangeRates As Boolean, _
                              ByVal PagesWide As Integer)
        Dim prntBudget As New PrintBudget(CurrentLogFrame, BudgetYearIndex, ShowDurationColumns, ShowLocalCurrencyColumns, ShowExchangeRates, PagesWide)

        AddHandler prntBudget.PagePrinted, AddressOf Budget_UpdateProgressBar
        AddHandler prntBudget.MinimumPageWidthChanged, AddressOf Budget_MinimumPageWidthChanged
        PrintDetailBudget.CustomPrintPreview.Document = prntBudget
    End Sub

    Private Sub Budget_UpdateProgressBar(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        If e.RowCount > 0 Then
            With PrintDetailBudget.PrintSettingsBar.ProgressBarDocument
                .Visible = True
                .Maximum = intMax
                .Value = e.LineIndex
                If e.LineIndex = intMax Then
                    .Visible = False
                    .Value = 0
                    Busy = False
                End If
            End With
        End If
    End Sub

    Private Sub Budget_MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)
        With PrintSettingsBudget
            .boolLoad = True
            .nudPagesWide.Value = e.PageWidth
            .boolLoad = False
        End With
    End Sub

    Private Sub PrintSettingsBudget_PrintBudgetSetupChanged(ByVal sender As Object, ByVal e As PrintBudgetSetupChangedEventArgs) Handles PrintSettingsBudget.PrintBudgetSetupChanged
        PrintSettingsBudget.Refresh()
        BudgetPreview(e.BudgetYearIndex, e.ShowDurationColumns, e.ShowLocalCurrencyColumns, e.ShowExchangeRates, e.PagesWide)
    End Sub
#End Region

#Region "Print PMF"
    Private Sub PmfPreview(ByVal intPrintSection As Integer, ByVal intPagesWide As Integer, ByVal boolShowTargetRowTitles As Boolean)
        Dim prntPmf As New PrintPMF(CurrentLogFrame, intPrintSection, intPagesWide, boolShowTargetRowTitles)

        AddHandler prntPmf.PagePrinted, AddressOf Pmf_UpdateProgressBar
        AddHandler prntPmf.MinimumPageWidthChanged, AddressOf Pmf_MinimumPageWidthChanged

        PrintDetailPmf.CustomPrintPreview.Document = prntPmf
    End Sub

    Private Sub Pmf_UpdateProgressBar(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailPmf.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub Pmf_MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)
        With PrintSettingsPmf
            .boolLoad = True
            .nudPagesWide.Value = e.PageWidth
            .boolLoad = False
        End With
    End Sub

    Private Sub PrintSettingsPmf_PrintPmfSetupChanged(ByVal sender As Object, ByVal e As PrintPmfSetupChangedEventArgs) Handles PrintSettingsPmf.PrintPmfSetupChanged
        PrintSettingsPmf.Refresh()
        PmfPreview(e.PrintSection, e.PagesWide, e.ShowTargetRowTitles)
    End Sub
#End Region

#Region "Print indicators"
    Private Sub IndicatorsPreview(ByVal intPrintSection As Integer, ByVal objTargetGroupGuid As Guid, ByVal intMeasurement As Integer, _
                                  ByVal boolPrintPurposes As Boolean, ByVal boolPrintOutputs As Boolean, _
                                  ByVal boolPrintActivities As Boolean, ByVal boolPrintOptionValues As Boolean, _
                                  ByVal boolPrintValueRanges As Boolean, ByVal boolPrintTargets As Boolean)

        Dim prntIndicators As New PrintIndicators(CurrentLogFrame, intPrintSection, objTargetGroupGuid, intMeasurement, _
                                                              boolPrintPurposes, boolPrintOutputs, boolPrintActivities, _
                                                              boolPrintOptionValues, boolPrintValueRanges, boolPrintTargets)
        AddHandler prntIndicators.LinePrinted, AddressOf UpdateProgressBarIndicators
        PrintDetailIndicators.CustomPrintPreview.Document = prntIndicators
    End Sub

    Private Sub UpdateProgressBarIndicators(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailIndicators.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub PrintSettingsIndicators_PrintIndicatorsSetupChanged(ByVal sender As Object, ByVal e As PrintIndicatorsSetupChangedEventArgs) Handles PrintSettingsIndicators.PrintIndicatorsSetupChanged
        PrintSettingsIndicators.Refresh()
        IndicatorsPreview(e.PrintSection, e.TargetGroupGuid, e.Measurement, e.PrintPurposes, e.PrintOutputs, e.PrintActivities, e.PrintOptionValues, e.PrintValueRanges, e.PrintTargets)
    End Sub
#End Region

#Region "Print Resources"
    Private Sub ResourcesPreview(ByVal intPagesWide As Integer)
        Dim prntResources As New PrintResources(CurrentLogFrame, intPagesWide)

        AddHandler prntResources.PagePrinted, AddressOf Resources_UpdateProgressBar
        AddHandler prntResources.MinimumPageWidthChanged, AddressOf Resources_MinimumPageWidthChanged

        PrintDetailResources.CustomPrintPreview.Document = prntResources
    End Sub

    Private Sub Resources_UpdateProgressBar(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailResources.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub Resources_MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)
        With PrintSettingsResources
            .boolLoad = True
            .nudPagesWide.Value = e.PageWidth
            .boolLoad = False
        End With
    End Sub

    Private Sub PrintSettingsResources_PrintResourcesSetupChanged(ByVal sender As Object, ByVal e As PrintResourcesSetupChangedEventArgs) Handles PrintSettingsResources.PrintResourcesSetupChanged
        PrintSettingsResources.Refresh()
        ResourcesPreview(e.PagesWide)
    End Sub
#End Region

#Region "Print Risk Register"
    Private Sub RiskRegisterPreview(ByVal intRiskCategory As Integer, ByVal intPagesWide As Integer)

        Dim prntRiskRegister As New PrintRiskRegister(CurrentLogFrame, intRiskCategory, intPagesWide)
        AddHandler prntRiskRegister.PagePrinted, AddressOf RiskRegister_UpdateProgressBar
        AddHandler prntRiskRegister.MinimumPageWidthChanged, AddressOf RiskRegister_MinimumPageWidthChanged

        PrintDetailRiskRegister.CustomPrintPreview.Document = prntRiskRegister
    End Sub

    Private Sub RiskRegister_UpdateProgressBar(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailRiskRegister.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub RiskRegister_MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)
        With PrintSettingsRiskRegister
            .boolLoad = True
            .nudPagesWide.Value = e.PageWidth
            .boolLoad = False
        End With
    End Sub

    Private Sub PrintSettingsRiskRegister_PrintRiskRegisterSetupChanged(ByVal sender As Object, ByVal e As PrintRiskRegisterSetupChangedEventArgs) Handles PrintSettingsRiskRegister.PrintRiskRegisterSetupChanged
        PrintSettingsRiskRegister.Refresh()
        RiskRegisterPreview(e.RiskCategory, e.PagesWide)
    End Sub
#End Region

#Region "Print Assumptions table"
    Private Sub AssumptionsPreview(ByVal intPrintSection As Integer, ByVal intPagesWide As Integer)

        Dim prntAssumptions As New PrintAssumptions(CurrentLogFrame, intPrintSection, intPagesWide)
        AddHandler prntAssumptions.PagePrinted, AddressOf Assumptions_UpdateProgressBar
        AddHandler prntAssumptions.MinimumPageWidthChanged, AddressOf Assumptions_MinimumPageWidthChanged

        PrintDetailAssumptions.CustomPrintPreview.Document = prntAssumptions
    End Sub

    Private Sub Assumptions_UpdateProgressBar(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailAssumptions.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub Assumptions_MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)
        With PrintSettingsAssumptions
            .boolLoad = True
            .nudPagesWide.Value = e.PageWidth
            .boolLoad = False
        End With
    End Sub

    Private Sub PrintSettingsAssumptions_PrintAssumptionsSetupChanged(ByVal sender As Object, ByVal e As PrintAssumptionsSetupChangedEventArgs) Handles PrintSettingsAssumptions.PrintAssumptionsSetupChanged
        PrintSettingsAssumptions.Refresh()
        AssumptionsPreview(e.PrintSection, e.PagesWide)
    End Sub
#End Region

#Region "Print Dependencies table"
    Private Sub DependenciesPreview(ByVal intPrintSection As Integer, ByVal boolShowDeliverables As Boolean, ByVal intPagesWide As Integer)

        Dim prntDependencies As New PrintDependencies(CurrentLogFrame, intPrintSection, boolShowDeliverables, intPagesWide)
        AddHandler prntDependencies.PagePrinted, AddressOf Dependencies_UpdateProgressBar
        AddHandler prntDependencies.MinimumPageWidthChanged, AddressOf Dependencies_MinimumPageWidthChanged

        PrintDetailDependencies.CustomPrintPreview.Document = prntDependencies
    End Sub

    Private Sub Dependencies_UpdateProgressBar(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailDependencies.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub Dependencies_MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)
        With PrintSettingsDependencies
            .boolLoad = True
            .nudPagesWide.Value = e.PageWidth
            .boolLoad = False
        End With
    End Sub

    Private Sub PrintSettingsDependencies_PrintDependenciesSetupChanged(ByVal sender As Object, ByVal e As PrintDependenciesSetupChangedEventArgs) Handles PrintSettingsDependencies.PrintDependenciesSetupChanged
        PrintSettingsDependencies.Refresh()
        DependenciesPreview(e.PrintSection, e.ShowDeliverables, e.PagesWide)
    End Sub
#End Region

#Region "Print PartnerList"
    Private Sub PartnerListPreview(ByVal intPartnerIndex As Integer, Optional ByVal PrintBorders As Boolean = False, Optional ByVal FillCells As Boolean = False)
        Dim prntPartnerList As PrintPartnerList
        If intPartnerIndex < CurrentLogFrame.ProjectPartners.Count Then
            prntPartnerList = New PrintPartnerList(CurrentLogFrame.ProjectPartners(intPartnerIndex))
        Else
            prntPartnerList = New PrintPartnerList(CurrentLogFrame.ProjectPartners)
        End If

        prntPartnerList.PrintBorders = PrintBorders
        prntPartnerList.FillCells = FillCells

        AddHandler prntPartnerList.LinePrinted, AddressOf UpdateProgressBarPartnerList
        PrintDetailPartnerList.CustomPrintPreview.Document = prntPartnerList
    End Sub

    Private Sub UpdateProgressBarPartnerList(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailPartnerList.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub PrintSettingsPartnerList_PrintPartnerListSetupChanged(ByVal sender As Object, ByVal e As PrintPartnerListSetupChangedEventArgs) Handles PrintSettingsPartnerList.PrintPartnerListSetupChanged
        PrintSettingsPartnerList.Refresh()
        PartnerListPreview(e.PartnerIndex, e.PrintBorders, e.FillCells)
    End Sub
#End Region

#Region "Print TargetGroupIdForm"
    Private Sub TargetGroupIdFormPreview(ByVal strTargetGroupName As String, Optional ByVal PrintBorders As Boolean = False, Optional ByVal FillCells As Boolean = False)
        Dim prntTargetGroupIdForm As PrintTargetGroupIdForm
        Dim objTargetGroups As TargetGroups = CurrentLogFrame.GetTargetGroups(strTargetGroupName)

        prntTargetGroupIdForm = New PrintTargetGroupIdForm(objTargetGroups)
        prntTargetGroupIdForm.PrintBorders = PrintBorders
        prntTargetGroupIdForm.FillCells = FillCells

        AddHandler prntTargetGroupIdForm.LinePrinted, AddressOf UpdateProgressBarTargetGroupIdForm
        PrintDetailTargetGroupIdForm.CustomPrintPreview.Document = prntTargetGroupIdForm
    End Sub

    Private Sub UpdateProgressBarTargetGroupIdForm(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
        Dim intMax As Integer = e.RowCount - 1
        If intMax < 0 Then intMax = 0

        Busy = True
        With PrintDetailTargetGroupIdForm.PrintSettingsBar.ProgressBarDocument
            .Visible = True
            .Maximum = intMax
            .Value = e.LineIndex
            If e.LineIndex = intMax Then
                .Visible = False
                .Value = 0
                Busy = False
            End If
        End With
    End Sub

    Private Sub PrintSettingsTargetGroupIdForm_PrintTargetGroupIdFormSetupChanged(ByVal sender As Object, ByVal e As PrintTargetGroupIdFormSetupChangedEventArgs) Handles PrintSettingsTargetGroupIdForm.PrintTargetGroupIdFormSetupChanged
        PrintSettingsTargetGroupIdForm.Refresh()
        TargetGroupIdFormPreview(e.TargetGroupName, e.PrintBorders, e.FillCells)
    End Sub
#End Region


End Class
