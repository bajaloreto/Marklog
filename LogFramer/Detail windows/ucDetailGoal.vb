Public Class DetailGoal
    Friend WithEvents ProjectBindingSource As New BindingSource

    Public Sub New()
        InitializeComponent()

        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Sub LoadItems()
        If CurrentLogFrame IsNot Nothing Then
            ProjectBindingSource.DataSource = CurrentLogFrame

            tbProjectTitle.DataBindings.Add(New Binding("Text", ProjectBindingSource, "ProjectTitle"))
            tbShortTitle.DataBindings.Add(New Binding("Text", ProjectBindingSource, "ShortTitle"))
            tbCode.DataBindings.Add(New Binding("Text", ProjectBindingSource, "Code"))
            ntbDuration.DataBindings.Add(New Binding("Text", ProjectBindingSource, "Duration"))

            With cmbDurationUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationRange)
                .DataBindings.Add(New Binding("SelectedIndex", ProjectBindingSource, "DurationUnit"))
            End With

            If CurrentLogFrame.StartDate <> Date.MinValue Then Me.dtpStartDate.Value = CurrentLogFrame.StartDate
            If CurrentLogFrame.EndDate <> Date.MinValue Then Me.dtpEndDate.Value = CurrentLogFrame.EndDate

            lvPartners.ProjectPartners = CurrentLogFrame.ProjectPartners
        End If
    End Sub

    Private Sub dtpStartDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.Enter
        UndoRedo.UndoBuffer_Initialise(CurrentLogFrame, "StartDate")
    End Sub

    Private Sub dtpStartDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.Validated
        If dtpStartDate.Value > Date.MinValue And dtpStartDate.Value <> CurrentLogFrame.StartDate Then
            CurrentLogFrame.StartDate = dtpStartDate.Value
            With CurrentProjectForm.dgvPlanning
                .StartDate = CurrentLogFrame.StartDate
                .Reload()
            End With
        End If
        ModifyEndDate()
    End Sub

    Private Sub dtpEndDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Enter
        UndoRedo.UndoBuffer_Initialise(CurrentLogFrame, "EndDate")
    End Sub

    Private Sub dtpEndDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Validated
        If dtpEndDate.Value > Date.MinValue And dtpEndDate.Value <> CurrentLogFrame.EndDate Then
            CurrentLogFrame.EndDate = dtpEndDate.Value
            With CurrentProjectForm.dgvPlanning
                .EndDate = CurrentLogFrame.EndDate
                .Reload()
            End With
        End If
        ModifyDuration()
    End Sub

    Private Sub ntbDuration_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbDuration.Validated
        ModifyEndDate()
    End Sub

    Private Sub cmbDurationUnit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDurationUnit.Validated
        ModifyEndDate()
    End Sub

    Public Sub ModifyEndDate()
        Dim OldDate As Date = CurrentLogFrame.EndDate

        If CurrentLogFrame.StartDate > Date.MinValue Then
            Select Case CurrentLogFrame.DurationUnit
                Case LogFrame.DurationUnits.Days
                    CurrentLogFrame.EndDate = CurrentLogFrame.StartDate.AddDays(CurrentLogFrame.Duration)
                Case LogFrame.DurationUnits.Weeks
                    Dim intWeeks As Integer = CurrentLogFrame.Duration * 7
                    CurrentLogFrame.EndDate = CurrentLogFrame.StartDate.AddDays(intWeeks).AddDays(-1)
                Case LogFrame.DurationUnits.Months
                    CurrentLogFrame.EndDate = CurrentLogFrame.StartDate.AddMonths(CurrentLogFrame.Duration).AddDays(-1)
                Case LogFrame.DurationUnits.Years
                    CurrentLogFrame.EndDate = CurrentLogFrame.StartDate.AddYears(CurrentLogFrame.Duration).AddDays(-1)

            End Select
        End If
        If CurrentLogFrame.StartDate <> Date.MinValue Then Me.dtpStartDate.Value = CurrentLogFrame.StartDate
        If CurrentLogFrame.EndDate <> Date.MinValue Then Me.dtpEndDate.Value = CurrentLogFrame.EndDate

        If OldDate <> CurrentLogFrame.EndDate Then
            With CurrentProjectForm.dgvPlanning
                .EndDate = CurrentLogFrame.EndDate
                .Reload()
            End With

        End If

        UpdateTargetDeadlines()
        UpdateBudgetYears()
    End Sub

    Private Sub ModifyDuration()
        Dim intDay, intMonth, intMonthOnly, intYear As Integer

        With CurrentLogFrame
            intYear = .EndDate.Year - .StartDate.Year
            intMonth = .EndDate.Month - .StartDate.Month + (intYear * 12)
            intMonthOnly = intMonth - (intYear * 12)
            If intMonthOnly < 0 Then intMonthOnly = intMonth

            If .EndDate.Day > .StartDate.Day Then
                intDay = .EndDate.Day - .StartDate.Day
            Else
                If .EndDate.Day <> .StartDate.Day - 1 Then
                    intDay = .EndDate.Day + (Date.DaysInMonth(.StartDate.Year, .StartDate.Month) - .StartDate.Day)
                    intMonth -= 1
                End If
            End If

            If intYear > 0 And intMonthOnly = 0 And intDay = 0 Then
                .Duration = intYear
                .DurationUnit = LogFrame.DurationUnits.Years
            ElseIf intMonth > 0 And intDay = 0 Then
                .Duration = intMonth
                .DurationUnit = LogFrame.DurationUnits.Months
            Else
                Dim tsDuration As TimeSpan = .EndDate.Subtract(.StartDate)
                intDay = tsDuration.Days + 1
                If Int(intDay / 7) = (intDay / 7) Then
                    .Duration = intDay / 7
                    .DurationUnit = LogFrame.DurationUnits.Weeks
                Else
                    .Duration = intDay
                    .DurationUnit = LogFrame.DurationUnits.Days
                End If
            End If
        End With

        ntbDuration.DataBindings(0).ReadValue()
        cmbDurationUnit.DataBindings(0).ReadValue()

        UpdateTargetDeadlines()
        UpdateBudgetYears()
    End Sub

    Private Sub UpdateTargetDeadlines()
        With CurrentLogFrame
            .TargetDeadlinesGoals.SetTargetDeadlines()
            .TargetDeadlinesPurposes.SetTargetDeadlines()
            .TargetDeadlinesOutputs.SetTargetDeadlines()
            .TargetDeadlinesActivities.SetTargetDeadlines()
        End With
    End Sub

    Private Sub UpdateBudgetYears()
        CurrentLogFrame.Budget.UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)
    End Sub
End Class
