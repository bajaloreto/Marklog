Partial Public Class frmParent
#Region "Import"
    Private Sub RibbonButtonImportWordLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonImportWordLogframe.Click
        ImportLogframe(True)
    End Sub

    Private Sub RibbonButtonImportExcelLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonImportExcelLogframe.Click
        ImportLogframe(False)
    End Sub

    Private Sub ImportLogframe(ByVal boolFromWordDocument As Boolean)
        Dim dialogImport As New DialogImportLogframe(boolFromWordDocument)
        Dim boolClose As Boolean
        Dim selLogframe As LogFrame = New LogFrame

        If dialogImport.ShowDialog = Windows.Forms.DialogResult.OK Then
            If MdiChildren.Count = 1 And CurrentProjectForm.IsDefaultProjectName And CurrentProjectForm.UndoList.Count = 0 Then boolClose = True

            selLogframe = dialogImport.ImportLogFrame

            OpenProject_CreateProjectForm(selLogframe, dialogImport.DocName, dialogImport.DocPath)

            AddNewProjectButton(CurrentProjectForm.Name)

            Initialize_Panels()
            Initialize_Clipboard()
            ReloadSplitUndoRedoButtons()

            If boolClose = True Then
                MdiChildren(0).Close()
                RibbonPanelProjects.Items.RemoveAt(0)
                RibbonLF.ResumeUpdating(True)
            End If

            StatusLabelGeneral.Text = LANG_Ready
        End If
    End Sub

    Private Sub RibbonButtonImportExcelBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonImportExcelBudget.Click
        Dim dialogImport As New DialogImportBudget()

        If dialogImport.ShowDialog = Windows.Forms.DialogResult.OK Then
            CurrentLogFrame.Budget = dialogImport.ImportBudget

            With CurrentLogFrame.Budget
                .idLogframe = CurrentLogFrame.idLogframe
                .ParentLogframeGuid = CurrentLogFrame.Guid
            End With

            With CurrentProjectForm
                .ProjectInitialise_Budget()
                .TabControlProject.SelectedTab = .TabPageBudget
            End With

            StatusLabelGeneral.Text = LANG_Ready
        End If
    End Sub
#End Region

#Region "Export to Word"
    Private Sub RibbonButtonExportWordLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordLogframe.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.Logframe)
        dialogExport.ShowDialog()
    End Sub

    Private Sub RibbonButtonExportWordPlanning_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordPlanning.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.Planning)
        dialogExport.ShowDialog()
    End Sub

    Private Sub RibbonButtonExportWordPmf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordPmf.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.Pmf)
        dialogExport.ShowDialog()
    End Sub

    Private Sub RibbonButtonExportWordRisks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordRisks.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.RisksTable)
        dialogExport.ShowDialog()
    End Sub

    Private Sub RibbonButtonExportWordAssumptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordAssumptions.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.AssumptionsTable)
        dialogExport.ShowDialog()
    End Sub

    Private Sub RibbonButtonExportWordPartners_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordPartners.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.PartnerList)
        dialogExport.ShowDialog()
    End Sub

    Private Sub RibbonButtonExportWordIdForm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportWordIdForm.Click
        Dim dialogExport As New DialogExportToWord(DialogExportToWord.Reports.TargetGroupIdForm)
        dialogExport.ShowDialog()
    End Sub
#End Region

#Region "Export to Excel"
    Private Sub RibbonButtonExportExcelLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportExcelLogframe.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.Logframe)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            With dialogExport
                Using objExcel As New ExcelIO
                    Dim objExportSettingsLogframe As PrintSettingsLogframe = TryCast(dialogExport.CurrentExportSettings, PrintSettingsLogframe)

                    If objExportSettingsLogframe IsNot Nothing Then
                        With objExportSettingsLogframe
                            objExcel.ExportLogFrameToExcel(CurrentLogFrame, .ShowIndicatorColumn, .ShowVerificationSourceColumn, .ShowAssumptionColumn, _
                                                           .ShowGoals, .ShowPurposes, .ShowOutputs, .ShowActivities, .ShowResourcesBudget)
                        End With
                    End If

                End Using
            End With
        End If
    End Sub

    Private Sub RibbonButtonExportExcelBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportExcelBudget.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.Budget)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using objExcel As New ExcelIO
                objExcel.ExportBudgetToExcel(CurrentLogFrame)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonExportExcelMonitoring_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportExcelMonitoring.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.MonitoringTool)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using objExcel As New ExcelIO
                objExcel.ExportMonitoringToolToExcel(CurrentLogFrame)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonExportExcelPmf_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportExcelPmf.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.Pmf)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using objExcel As New ExcelIO
                objExcel.ExportPmfToExcel(CurrentLogFrame)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonExportExcelRiskRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportExcelRiskRegister.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.RiskRegister)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using objExcel As New ExcelIO
                objExcel.ExportRiskRegisterToExcel(CurrentLogFrame)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonExportMonitoringTool_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportMonitoringTool.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.MonitoringTool)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using objExcel As New ExcelIO
                objExcel.ExportMonitoringToolToExcel(CurrentLogFrame)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonExportRiskRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonExportRiskRegister.Click
        Dim dialogExport As New DialogExportToExcel(DialogExportToExcel.Reports.RiskRegister)

        If dialogExport.ShowDialog = Windows.Forms.DialogResult.OK Then
            Using objExcel As New ExcelIO
                objExcel.ExportRiskRegisterToExcel(CurrentLogFrame)
            End Using
        End If
    End Sub
#End Region
End Class
