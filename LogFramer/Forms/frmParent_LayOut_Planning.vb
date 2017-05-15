Partial Public Class frmParent
    Private Sub RibbonTabLayOutPlanning_ActiveChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonTabLayOutPlanning.ActiveChanged
        With CurrentProjectForm.dgvPlanning
            RibbonButtonPlanningShowDatesColumn.Checked = .ShowDatesColumns
            SetPlanningPeriodView(.PeriodView)
        End With
    End Sub

    Private Sub RibbonButtonPlanningShowDatesColumn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningShowDatesColumn.Click
        ShowPlanningDates()
    End Sub

    Private Sub ShowPlanningDates()
        Dim boolShow As Boolean

        With CurrentProjectForm.dgvPlanning
            If .ShowDatesColumns = True Then boolShow = False Else boolShow = True

            .ShowDatesColumns = boolShow
            My.Settings.setShowDatesColumns = boolShow
            RibbonButtonPlanningShowDatesColumn.Checked = boolShow

            .ShowColumns()
        End With
    End Sub

    Private Sub RibbonButtonPlanningKeyMoments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningKeyMoments.Click
        With CurrentProjectForm.dgvPlanning
            If RibbonButtonPlanningKeyMoments.Checked = True Then
                If .ElementsView = DataGridViewPlanning.ShowElements.ShowActivities Then
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowBoth
                Else
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowKeyMoments
                End If
            Else
                If .ElementsView = DataGridViewPlanning.ShowElements.ShowBoth Then
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowActivities
                Else
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowActivities
                End If
            End If
            My.Settings.setPlanningElementsView = .ElementsView
            .Reload()
        End With

        SetPlanningElementsViewButtons()
    End Sub

    Private Sub RibbonButtonPlanningActivities_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningActivities.Click
        With CurrentProjectForm.dgvPlanning
            If RibbonButtonPlanningActivities.Checked = True Then
                If .ElementsView = DataGridViewPlanning.ShowElements.ShowKeyMoments Then
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowBoth
                Else
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowActivities
                End If
            Else
                If .ElementsView = DataGridViewPlanning.ShowElements.ShowBoth Then
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowKeyMoments
                Else
                    .ElementsView = DataGridViewPlanning.ShowElements.ShowKeyMoments
                End If
            End If
            My.Settings.setPlanningElementsView = .ElementsView
            .Reload()
        End With

        SetPlanningElementsViewButtons()
    End Sub

    Private Sub SetPlanningElementsViewButtons()
        With CurrentProjectForm.dgvPlanning
            If .ElementsView = DataGridViewPlanning.ShowElements.ShowKeyMoments Or .ElementsView = DataGridViewPlanning.ShowElements.ShowBoth Then
                RibbonButtonPlanningKeyMoments.Checked = True
            Else
                RibbonButtonPlanningKeyMoments.Checked = False
            End If
            If .ElementsView = DataGridViewPlanning.ShowElements.ShowActivities Or .ElementsView = DataGridViewPlanning.ShowElements.ShowBoth Then
                RibbonButtonPlanningActivities.Checked = True
            Else
                RibbonButtonPlanningActivities.Checked = False
            End If
        End With
    End Sub

    Public Sub SetPlanningPeriodView(ByVal intPeriodView As Integer)
        With CurrentProjectForm.dgvPlanning
            If intPeriodView <> .PeriodView Then
                .PeriodView = intPeriodView
                .Reload()

                My.Settings.setPlanningPeriodView = .PeriodView
            End If

            SetPlanningPeriodViewButtons(intPeriodView)
        End With
    End Sub

    Public Sub SetPlanningPeriodViewButtons(ByVal intPeriodView As Integer)
        RibbonButtonPlanningDays.Checked = False
        RibbonButtonPlanningWeeks.Checked = False
        RibbonButtonPlanningMonths.Checked = False
        RibbonButtonPlanningYears.Checked = False

        Select Case intPeriodView
            Case DataGridViewPlanning.PeriodViews.Day
                RibbonButtonPlanningDays.Checked = True
            Case DataGridViewPlanning.PeriodViews.Week
                RibbonButtonPlanningWeeks.Checked = True
            Case DataGridViewPlanning.PeriodViews.Month
                RibbonButtonPlanningMonths.Checked = True
            Case DataGridViewPlanning.PeriodViews.Year
                RibbonButtonPlanningYears.Checked = True
        End Select
    End Sub

    Private Sub RibbonButtonPlanningDays_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningDays.Click
        SetPlanningPeriodView(DataGridViewPlanning.PeriodViews.Day)
    End Sub

    Private Sub RibbonButtonPlanningWeeks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningWeeks.Click
        SetPlanningPeriodView(DataGridViewPlanning.PeriodViews.Week)
    End Sub

    Private Sub RibbonButtonPlanningMonths_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningMonths.Click
        SetPlanningPeriodView(DataGridViewPlanning.PeriodViews.Month)
    End Sub

    Private Sub RibbonButtonPlanningYears_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningYears.Click
        SetPlanningPeriodView(DataGridViewPlanning.PeriodViews.Year)
    End Sub

    Private Sub RibbonButtonPlanningShowKeyMomentLinks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningShowKeyMomentLinks.Click
        With CurrentProjectForm.dgvPlanning
            .ShowKeyMomentLinks = RibbonButtonPlanningShowKeyMomentLinks.Checked
            My.Settings.setPlanningKeyMomentLinks = .ShowKeyMomentLinks
            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonPlanningShowActivityLinks_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningShowActivityLinks.Click
        With CurrentProjectForm.dgvPlanning
            .ShowActivityLinks = RibbonButtonPlanningShowActivityLinks.Checked
            My.Settings.setPlanningActivityLinks = .ShowActivityLinks
            .Reload()
        End With
    End Sub

    Private Sub RibbonButtonPlanningDetailsWindow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningDetailsWindow.Click
        RibbonButtonPlanningDetailsWindow.Text = CurrentProjectForm.ShowDetailsPlanning()
    End Sub

    Private Sub RibbonButtonPlanningEmptyCells_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPlanningEmptyCells.Click
        HideEmptyCellsPlanning()
    End Sub

    Private Sub HideEmptyCellsPlanning()
        With CurrentProjectForm.dgvPlanning
            If .HideEmptyCells = False Then
                .HideEmptyCells = True
                RibbonButtonPlanningEmptyCells.Text = LANG_EmptyCellsShow
                My.Settings.setPlanningHideEmptyCells = True

            Else
                .HideEmptyCells = False
                RibbonButtonPlanningEmptyCells.Text = LANG_EmptyCellsHide
                My.Settings.setPlanningHideEmptyCells = False
            End If
            .Reload()
        End With
    End Sub
End Class
