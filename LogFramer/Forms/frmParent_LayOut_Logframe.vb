Partial Public Class frmParent
    Public Sub SetTextSelectionToNothing()
        RibbonComboBoxTextSelection.SelectedItem = RibbonButtonSelectNothing
    End Sub

    Private Sub RibbonTabLayOutLogframe_ActiveChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonTabLayOutLogframe.ActiveChanged
        With CurrentProjectForm.dgvLogframe
            RibbonButtonShowIndicators.Checked = .ShowIndicatorColumn
            RibbonButtonShowVerificationSources.Checked = .ShowVerificationSourceColumn
            RibbonButtonShowAssumptions.Checked = .ShowAssumptionColumn
            RibbonButtonShowGoals.Checked = .ShowGoals
            RibbonButtonShowPurposes.Checked = .ShowPurposes
            RibbonButtonShowOutputs.Checked = .ShowOutputs
            RibbonButtonShowActivities.Checked = .ShowActivities

            If .ShowResourcesBudget = True Then
                RibbonButtonShowResourcesBudget.Checked = True
                RibbonButtonShowProcessIndicators.Checked = False
            Else
                RibbonButtonShowResourcesBudget.Checked = False
                RibbonButtonShowProcessIndicators.Checked = True
            End If
        End With
    End Sub

    Private Sub RibbonButtonShowProjectLogic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowProjectLogic.Click
        ShowProjectLogic()
    End Sub

    Private Sub ShowProjectLogic()
        Dim boolShow As Boolean

        With CurrentProjectForm.dgvLogframe
            If .ShowIndicatorColumn = False Or .ShowVerificationSourceColumn = False Or .ShowAssumptionColumn = False Then
                boolShow = True
            Else
                boolShow = False
            End If

            .ShowIndicatorColumn = boolShow
            .ShowVerificationSourceColumn = boolShow
            .ShowAssumptionColumn = boolShow

            My.Settings.setShowIndicatorColumn = boolShow
            My.Settings.setShowVerificationSourceColumn = boolShow
            My.Settings.setShowAssumptionColumn = boolShow

            RibbonButtonShowIndicators.Checked = boolShow
            RibbonButtonShowVerificationSources.Checked = boolShow
            RibbonButtonShowAssumptions.Checked = boolShow

            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonButtonShowIndicators_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowIndicators.Click
        With CurrentProjectForm.dgvLogframe
            If .ShowIndicatorColumn = True Then .ShowIndicatorColumn = False Else .ShowIndicatorColumn = True
        End With
        ShowIndicators()
    End Sub

    Public Sub ShowIndicators()
        With CurrentProjectForm.dgvLogframe
            My.Settings.setShowIndicatorColumn = .ShowIndicatorColumn
            RibbonButtonShowIndicators.Checked = .ShowIndicatorColumn

            If .ShowIndicatorColumn = False Then
                .ShowVerificationSourceColumn = False
                My.Settings.setShowVerificationSourceColumn = .ShowVerificationSourceColumn
                RibbonButtonShowVerificationSources.Checked = False
            End If
            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonButtonShowVerificationSources_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowVerificationSources.Click
        With CurrentProjectForm.dgvLogframe
            If .ShowVerificationSourceColumn = True Then .ShowVerificationSourceColumn = False Else .ShowVerificationSourceColumn = True
        End With
        ShowVerificationSources()
    End Sub

    Public Sub ShowVerificationSources()
        With CurrentProjectForm.dgvLogframe
            My.Settings.setShowVerificationSourceColumn = .ShowVerificationSourceColumn
            RibbonButtonShowVerificationSources.Checked = .ShowVerificationSourceColumn

            If .ShowVerificationSourceColumn = True Then
                .ShowIndicatorColumn = True
                My.Settings.setShowIndicatorColumn = .ShowIndicatorColumn
                RibbonButtonShowIndicators.Checked = True
            End If
            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonButtonShowAssumptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowAssumptions.Click
        With CurrentProjectForm.dgvLogframe
            If .ShowAssumptionColumn = True Then .ShowAssumptionColumn = False Else .ShowAssumptionColumn = True
        End With
        ShowAssumptions()
    End Sub

    Public Sub ShowAssumptions()
        With CurrentProjectForm.dgvLogframe
            My.Settings.setShowAssumptionColumn = .ShowAssumptionColumn
            RibbonButtonShowAssumptions.Checked = .ShowAssumptionColumn
            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonPanelColumns_ButtonMoreClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonPanelColumns.ButtonMoreClick
        Dim dialogSettings As New DialogSettings
        With dialogSettings
            .TabControlSettings.SelectedIndex = 0
            If .ShowDialog() = vbOK Then
                SetProjectLogicScheme()
            End If
        End With
    End Sub

    Private Sub SetProjectLogicScheme()
        With CurrentProjectForm.dgvLogframe
            .Columns(0).DefaultCellStyle.Font = New Font(My.Settings.setDefaultFont, FontStyle.Regular)
            .Columns(2).DefaultCellStyle.Font = New Font(My.Settings.setDefaultFont, FontStyle.Regular)
            .Columns(4).DefaultCellStyle.Font = New Font(My.Settings.setDefaultFont, FontStyle.Regular)
            .Columns(6).DefaultCellStyle.Font = New Font(My.Settings.setDefaultFont, FontStyle.Regular)
            .Reload()
            '.AutoResizeColumn(0)
            '.AutoResizeColumn(2)
            '.AutoResizeColumn(4)
            '.AutoResizeColumn(6)
        End With
        CurrentProjectForm.SetTypeOfDetailWindowLogframe()
    End Sub

    Private Sub RibbonButtonShowGoals_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowGoals.Click
        ShowGoals()
    End Sub

    Private Sub ShowGoals()
        Dim boolShow As Boolean

        With CurrentProjectForm.dgvLogframe
            If .ShowGoals = True Then boolShow = False Else boolShow = True

            .ShowGoals = boolShow
            RibbonButtonShowGoals.Checked = boolShow

            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonShowPurposes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowPurposes.Click
        ShowPurposes()
    End Sub

    Private Sub ShowPurposes()
        Dim boolShow As Boolean

        With CurrentProjectForm.dgvLogframe
            If .ShowPurposes = True Then boolShow = False Else boolShow = True

            .ShowPurposes = boolShow
            RibbonButtonShowPurposes.Checked = boolShow

            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonShowOutputs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowOutputs.Click
        ShowOutputs()
    End Sub

    Private Sub ShowOutputs()
        Dim boolShow As Boolean

        With CurrentProjectForm.dgvLogframe
            If .ShowOutputs = True Then boolShow = False Else boolShow = True

            .ShowOutputs = boolShow
            RibbonButtonShowOutputs.Checked = boolShow

            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonShowActivities_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowActivities.Click
        ShowActivities()
    End Sub

    Private Sub ShowActivities()
        Dim boolShow As Boolean

        With CurrentProjectForm.dgvLogframe
            If .ShowActivities = True Then boolShow = False Else boolShow = True

            .ShowActivities = boolShow
            RibbonButtonShowActivities.Checked = boolShow

            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonShowLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowLogframe.Click
        RibbonButtonShowIndicators.Checked = True
        RibbonButtonShowVerificationSources.Checked = True
        RibbonButtonShowAssumptions.Checked = True
        RibbonButtonShowGoals.Checked = True
        RibbonButtonShowPurposes.Checked = True
        RibbonButtonShowOutputs.Checked = True
        RibbonButtonShowActivities.Checked = True

        With CurrentProjectForm.dgvLogframe
            .ShowIndicatorColumn = True
            .ShowVerificationSourceColumn = True
            .ShowAssumptionColumn = True
            .ShowGoals = True
            .ShowPurposes = True
            .ShowOutputs = True
            .ShowActivities = True
            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonButtonShowPurposesOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowPurposesOnly.Click
        RibbonButtonShowIndicators.Checked = False
        RibbonButtonShowVerificationSources.Checked = False
        RibbonButtonShowAssumptions.Checked = False
        RibbonButtonShowGoals.Checked = False
        RibbonButtonShowPurposes.Checked = True
        RibbonButtonShowOutputs.Checked = False
        RibbonButtonShowActivities.Checked = False

        With CurrentProjectForm.dgvLogframe
            .ShowIndicatorColumn = False
            .ShowVerificationSourceColumn = False
            .ShowAssumptionColumn = False
            .ShowGoals = False
            .ShowPurposes = True
            .ShowOutputs = False
            .ShowActivities = False
            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonButtonShowResourcesBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowResourcesBudget.Click
        CurrentProjectForm.dgvLogframe.ShowResourcesBudget = True
        ShowResourcesBudget()
    End Sub

    Public Sub ShowResourcesBudget()
        With CurrentProjectForm.dgvLogframe
            If .ShowResourcesBudget = True Then
                My.Settings.setShowResourcesBudget = True
                RibbonButtonShowResourcesBudget.Checked = True
                RibbonButtonShowProcessIndicators.Checked = False
            Else
                My.Settings.setShowResourcesBudget = False
                RibbonButtonShowResourcesBudget.Checked = False
                RibbonButtonShowProcessIndicators.Checked = True
            End If
            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonShowProcessIndicators_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonShowProcessIndicators.Click
        CurrentProjectForm.dgvLogframe.ShowResourcesBudget = False
        ShowResourcesBudget()
    End Sub

    Private Sub RibbonButtonDetailsWindow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonDetailsWindow.Click
        RibbonButtonDetailsWindow.Text = CurrentProjectForm.ShowDetailsLogframe()
    End Sub

    Private Sub RibbonButtonEmptyCells_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonEmptyCells.Click
        HideEmptyCellsLogframe()
    End Sub

    Private Sub HideEmptyCellsLogframe()
        With CurrentProjectForm.dgvLogframe
            If .HideEmptyCells = False Then
                .HideEmptyCells = True
                RibbonButtonEmptyCells.Text = LANG_EmptyCellsShow
                My.Settings.setLogframeHideEmptyCells = True

            Else
                .HideEmptyCells = False
                RibbonButtonEmptyCells.Text = LANG_EmptyCellsHide
                My.Settings.setLogframeHideEmptyCells = False
            End If
            .Reload()
        End With
    End Sub

    Private Sub RibbonPanelStructure_ButtonMoreClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonPanelStructure.ButtonMoreClick
        Dim dialogSettings As New DialogSettings
        With dialogSettings
            .TabControlSettings.SelectedIndex = 1
            If .ShowDialog() = vbOK Then
                SetProjectLogicScheme()
            End If
        End With
    End Sub
End Class
