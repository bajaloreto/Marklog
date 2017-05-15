Public Class DetailProjectInfo
    Friend WithEvents ProjectBindingSource As New BindingSource

#Region "Initialise"
    Public Sub LoadItems()
        If CurrentLogFrame IsNot Nothing Then
            ProjectBindingSource.DataSource = CurrentLogFrame

            tbProjectTitle.DataBindings.Add(New Binding("Text", ProjectBindingSource, "ProjectTitle"))
            tbShortTitle.DataBindings.Add(New Binding("Text", ProjectBindingSource, "ShortTitle"))
            tbCode.DataBindings.Add(New Binding("Text", ProjectBindingSource, "Code"))
            ntbDuration.DataBindings.Add(New Binding("Text", ProjectBindingSource, "Duration", True))
            ntbDuration.DataBindings(0).FormatString = "N0"

            With cmbDurationUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationRange)
                .DataBindings.Add(New Binding("SelectedIndex", ProjectBindingSource, "DurationUnit"))
            End With

            If CurrentLogFrame.StartDate <> Date.MinValue Then Me.dtpStartDate.Value = CurrentLogFrame.StartDate
            If CurrentLogFrame.EndDate <> Date.MinValue Then Me.dtpEndDate.Value = CurrentLogFrame.EndDate

            LoadItems_TargetGroups()
            With SplitContainerTargetGroups
                .SplitterDistance = .Size.Height - 120
            End With

            lvPartners.ProjectPartners = CurrentLogFrame.ProjectPartners
            LoadItems_Contacts()
            LoadItems_TargetDeadlines()
            LoadItems_RiskMonitoringDeadlines()
        End If
    End Sub

    Public Sub LoadItems_TargetGroups()
        lvTargetGroups.LoadTargetGroups(True)
        lvTargetGroupStatistics.TargetGroups = lvTargetGroups.TargetGroups
    End Sub

    Private Sub LoadItems_Contacts()
        Dim objContacts As New Contacts
        Dim objParentOrganisationGuid As Guid
        For Each selPartner As ProjectPartner In CurrentLogFrame.ProjectPartners
            For Each selContact As Contact In selPartner.Organisation.Contacts
                objContacts.Add(selContact)
            Next
        Next

        If CurrentLogFrame.ProjectPartners.Count > 0 Then
            objParentOrganisationGuid = CurrentLogFrame.ProjectPartners(0).Organisation.Guid
            lvContacts.Reload(objContacts, objParentOrganisationGuid)
        End If


    End Sub

    Private Sub LoadItems_TargetDeadlines()
        With CurrentLogFrame
            If .TargetDeadlinesGoals.TargetDeadlines.Count = 0 Then
                With .TargetDeadlinesGoals
                    .Repetition = TargetDeadlinesSection.RepetitionOptions.SingleTarget
                    .PeriodEnd = 1
                    .PeriodUnitEnd = DurationUnits.Day
                    .SetTargetDeadlines()
                End With
            End If
            If .TargetDeadlinesPurposes.TargetDeadlines.Count = 0 Then
                With .TargetDeadlinesPurposes
                    .Repetition = TargetDeadlinesSection.RepetitionOptions.SingleTarget
                    .PeriodEnd = 1
                    .PeriodUnitEnd = DurationUnits.Day
                    .SetTargetDeadlines()
                End With
            End If
            If .TargetDeadlinesOutputs.TargetDeadlines.Count = 0 Then _
                .TargetDeadlinesOutputs.SetTargetDeadlines()
            If .TargetDeadlinesActivities.TargetDeadlines.Count = 0 Then _
                .TargetDeadlinesActivities.SetTargetDeadlines()
        End With

        With ucTargetDeadlinesGoals
            .TargetDeadlinesSection = CurrentLogFrame.TargetDeadlinesGoals
            .LoadItems()
        End With
        With ucTargetDeadlinesPurposes
            .TargetDeadlinesSection = CurrentLogFrame.TargetDeadlinesPurposes
            .LoadItems()
        End With
        With ucTargetDeadlinesOutputs
            .TargetDeadlinesSection = CurrentLogFrame.TargetDeadlinesOutputs
            .LoadItems()
        End With
        With ucTargetDeadlinesActivities
            .TargetDeadlinesSection = CurrentLogFrame.TargetDeadlinesActivities
            .LoadItems()
        End With
    End Sub

    Private Sub LoadItems_RiskMonitoringDeadlines()
        With CurrentLogFrame
            If .RiskMonitoring.RiskMonitoringDeadlines.Count = 0 Then _
                .RiskMonitoring.SetRiskMonitoringDeadlines()
        End With

        With ucRiskMonitoring
            .RiskMonitoring = CurrentLogFrame.RiskMonitoring
            .LoadItems()
        End With
    End Sub
#End Region

#Region "General project information"
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
        If dtpEndDate.Value > Date.MinValue AndAlso dtpEndDate.Value <> CurrentLogFrame.EndDate Then
            CurrentLogFrame.EndDate = dtpEndDate.Value
            With CurrentProjectForm.dgvPlanning
                .EndDate = CurrentLogFrame.EndDate
                .Reload()
            End With
        End If
        ModifyDuration()
    End Sub

    Private Sub ntbDuration_Validated(ByVal sender As Object, ByVal e As System.EventArgs)
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
#End Region

#Region "Methods & events"
    Private Sub UpdateTargetDeadlines()
        With CurrentLogFrame
            .TargetDeadlinesGoals.SetTargetDeadlines()
            .TargetDeadlinesPurposes.SetTargetDeadlines()
            .TargetDeadlinesOutputs.SetTargetDeadlines()
            .TargetDeadlinesActivities.SetTargetDeadlines()
            .RiskMonitoring.SetRiskMonitoringDeadlines()
        End With

        ucTargetDeadlinesGoals.Reload()
        ucTargetDeadlinesPurposes.Reload()
        ucTargetDeadlinesOutputs.Reload()
        ucTargetDeadlinesActivities.Reload()
        ucRiskMonitoring.Reload()
    End Sub

    Private Sub UpdateBudgetYears()
        CurrentLogFrame.Budget.UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)
    End Sub

    Private Sub TabControlProjectInfo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlProjectInfo.SelectedIndexChanged
        Select Case TabControlProjectInfo.SelectedTab.Name
            Case TabPageTargetGroups.Name
                lvTargetGroups.Focus()
                CurrentControl = lvTargetGroups
            Case TabPagePartners.Name
                lvPartners.Focus()
                CurrentControl = lvPartners
        End Select

        With frmParent
            If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
        End With
    End Sub
#End Region

#Region "Target groups"
    Private Sub lvTargetGroups_ColumnWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnWidthChangedEventArgs) Handles lvTargetGroups.ColumnWidthChanged
        lvTargetGroupStatistics.Columns(0).Width = lvTargetGroups.Columns(0).Width + lvTargetGroups.Columns(1).Width

        For i = 2 To 5
            lvTargetGroupStatistics.Columns(i - 1).Width = lvTargetGroups.Columns(i).Width
        Next
    End Sub

    Private Sub lvTargetGroups_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvTargetGroups.MouseUp
        With frmParent
            .RibbonButtonNewItem.Enabled = True

            If lvTargetGroups.SelectedIndices.Count = 0 Then '.TargetGroups.Count = 0 Then
                .RibbonButtonEditItem.Enabled = False
                .RibbonButtonRemoveItem.Enabled = False
            Else
                .RibbonButtonEditItem.Enabled = True
                .RibbonButtonRemoveItem.Enabled = True
            End If

            If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
        End With
    End Sub

    Private Sub lvTargetGroups_Updated() Handles lvTargetGroups.Updated
        lvTargetGroupStatistics.TargetGroups = lvTargetGroups.TargetGroups
        lvTargetGroupStatistics.LoadItems()
    End Sub
#End Region

#Region "Partner organisations & contacts"

    Private Sub lvPartners_ListUpdated() Handles lvPartners.ListUpdated
        LoadItems_Contacts()
    End Sub

    Private Sub lvPartners_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvPartners.MouseUp
        With frmParent
            .RibbonButtonNewItem.Enabled = True

            If lvPartners.SelectedIndices.Count = 0 Then '.ProjectPartners.Count = 0 Then
                .RibbonButtonEditItem.Enabled = False
                .RibbonButtonRemoveItem.Enabled = False
            Else
                .RibbonButtonEditItem.Enabled = True
                .RibbonButtonRemoveItem.Enabled = True
            End If

            If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
        End With
    End Sub

    Private Sub lvPartners_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvPartners.SelectedIndexChanged
        With lvPartners
            If .SelectedIndices.Count = 0 Then
                LoadItems_Contacts()
            Else
                Dim selGuid As Guid = New Guid(.SelectedItems(0).Name)
                Dim selProjectPartner As ProjectPartner = CurrentLogFrame.GetProjectPartnerByGuid(selGuid)
                Dim objContacts As New Contacts

                objContacts.AddRange(selProjectPartner.Organisation.Contacts)

                lvContacts.Reload(objContacts, selProjectPartner.Organisation.Guid)
            End If
        End With
    End Sub

    Private Sub lvContacts_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvContacts.MouseUp
        With frmParent
            If CurrentLogFrame.ProjectPartners.Count = 0 Then
                .RibbonButtonNewItem.Enabled = False
            Else
                .RibbonButtonNewItem.Enabled = True
            End If

            If lvContacts.SelectedIndices.Count = 0 Then '.Contacts.Count = 0 Then
                .RibbonButtonEditItem.Enabled = False
                .RibbonButtonRemoveItem.Enabled = False
            Else
                .RibbonButtonEditItem.Enabled = True
                .RibbonButtonRemoveItem.Enabled = True
            End If

            If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
        End With
    End Sub
#End Region
End Class
