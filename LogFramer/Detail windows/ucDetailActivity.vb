Public Class DetailActivity
    Private objCurrentActivity As Activity
    Private boolTypeChanged As Boolean
    Private boolTextSelected As Boolean
    Private boolHasFocus As Boolean

    Friend WithEvents ActivityBindingSource As New BindingSource
    Friend WithEvents ActivityDetailBindingSource As New BindingSource

    Public Property CurrentActivity() As Activity
        Get
            Return objCurrentActivity
        End Get
        Set(ByVal value As Activity)
            objCurrentActivity = value
        End Set
    End Property

#Region "Initialise"
    Public Sub New(ByVal activity As Activity)
        InitializeComponent()

        Me.CurrentActivity = activity
        Me.Dock = DockStyle.Fill
        ProcessOrActivity()
        TabControlActivity.SelectedIndex = CurrentActivitiesTab

        LoadItems()
    End Sub

    Public Sub LoadItems()

        If CurrentActivity IsNot Nothing AndAlso CurrentActivity.ActivityDetail IsNot Nothing Then
            Dim lstActivities As List(Of IdValuePair) = CurrentLogFrame.GetReferenceMomentsList(CurrentActivity)
            ActivityBindingSource.DataSource = CurrentActivity
            ActivityDetailBindingSource.DataSource = CurrentActivity.ActivityDetail


            tbActivity.DataBindings.Add("Text", ActivityBindingSource, "Text")
            dtbExactStartDate.DataBindings.Add("DateValue", ActivityBindingSource, "ExactStartDate", True)
            dtbExactStartDate.DataBindings(0).FormatString = "d"
            dtbExactEndDate.DataBindings.Add("DateValue", ActivityBindingSource, "ExactEndDate", True)
            dtbExactEndDate.DataBindings(0).FormatString = "d"

            ntbPeriod.DataBindings.Add("Text", ActivityDetailBindingSource, "Period")
            With cmbPeriodUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
                .DataBindings.Add("SelectedIndex", ActivityDetailBindingSource, "PeriodUnit")
            End With
            With cmbPeriodDirection
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DirectionValues)
                .DataBindings.Add("SelectedIndex", ActivityDetailBindingSource, "PeriodDirection")
            End With
            With cmbGuidReferenceMoment
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = lstActivities
                .ValueMember = "Id"
                .DisplayMember = "Value"
                .DataBindings.Add("SelectedValue", ActivityDetailBindingSource, "GuidReferenceMoment")
            End With

            ntbDuration.DataBindings.Add("Text", ActivityDetailBindingSource, "Duration")
            With cmbDurationUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
                .DataBindings.Add("SelectedIndex", ActivityDetailBindingSource, "DurationUnit")
            End With
            chkDurationUntilEnd.DataBindings.Add("Checked", ActivityDetailBindingSource, "DurationUntilEnd")
            chkDurationUntilEnd.Type = CheckBoxLF.Types.DurationUntilEndOfProject
            dtbStartDateMainActivity.DataBindings.Add("DateValue", ActivityDetailBindingSource, "StartDateMainActivity")
            dtbEndDateMainActivity.DataBindings.Add("DateValue", ActivityDetailBindingSource, "EndDateMainActivity")


            With cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_ActivityTypes
                .ValueMember = "Id"
                .DisplayMember = "Value"
                .DataBindings.Add("SelectedValue", ActivityDetailBindingSource, "Type")
            End With
            With cmbOrganisation
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = CurrentLogFrame.PartnerNamesList
                .DataBindings.Add("Text", ActivityDetailBindingSource, "Organisation")
            End With
            With cmbLocation
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = CurrentLogFrame.ActivitiesLocationsList.ToArray
                .DataBindings.Add("Text", ActivityDetailBindingSource, "Location")
            End With


            ntbPreparation.DataBindings.Add("Text", ActivityDetailBindingSource, "Preparation")
            With cmbPreparationUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
                .DataBindings.Add("SelectedIndex", ActivityDetailBindingSource, "PreparationUnit")
            End With
            chkPreparationRepeat.DataBindings.Add("Checked", ActivityDetailBindingSource, "PreparationRepeat")
            chkPreparationFromStart.DataBindings.Add("Checked", ActivityDetailBindingSource, "PreparationFromStart")
            chkPreparationFromStart.Type = CheckBoxLF.Types.DurationFromStartOfProject
            dtbStartDatePreparation.DataBindings.Add("DateValue", ActivityDetailBindingSource, "StartDatePreparation")
            dtbEndDatePreparation.DataBindings.Add("DateValue", ActivityDetailBindingSource, "StartDateMainActivity")

            ntbFollowUp.DataBindings.Add("Text", ActivityDetailBindingSource, "FollowUp")
            With cmbFollowUpUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
                .DataBindings.Add("SelectedIndex", ActivityDetailBindingSource, "FollowUpUnit")
            End With
            chkFollowUpRepeat.DataBindings.Add("Checked", ActivityDetailBindingSource, "FollowUpRepeat")
            chkFollowUpUntilEnd.DataBindings.Add("Checked", ActivityDetailBindingSource, "FollowUpUntilEnd")
            chkFollowUpUntilEnd.Type = CheckBoxLF.Types.DurationUntilEndOfProject
            dtbStartDateFollowUp.DataBindings.Add("DateValue", ActivityDetailBindingSource, "EndDateMainActivity")
            dtbEndDateFollowUp.DataBindings.Add("DateValue", ActivityDetailBindingSource, "EndDateFollowUp")


            ntbRepeatEvery.DataBindings.Add("Text", ActivityDetailBindingSource, "RepeatEvery")
            With cmbRepeatUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
                .DataBindings.Add("SelectedIndex", ActivityDetailBindingSource, "RepeatUnit")
            End With
            ntbRepeatTimes.DataBindings.Add("Text", ActivityDetailBindingSource, "RepeatTimes")
            chkRepeatUntilEnd.DataBindings.Add("Checked", ActivityDetailBindingSource, "RepeatUntilEnd")
            chkRepeatUntilEnd.Type = CheckBoxLF.Types.DurationUntilEndOfProject

            With lvRepeatDates
                .View = View.Details
                .FullRowSelect = True
                .Columns.Add(LANG_StartDate, 100, HorizontalAlignment.Right)
                .Columns.Add(LANG_EndDate, 100, HorizontalAlignment.Right)
                .Columns.Add(LANG_PreparationStarts, 100, HorizontalAlignment.Right)
                .Columns.Add(LANG_FollowUpEnds, 100, HorizontalAlignment.Right)
            End With

            lvDetailProcesses.Activities = CurrentActivity.Activities

            SelectDateSystem()

        End If
    End Sub
#End Region

#Region "General methods"
    Private Sub ProcessOrActivity()
        Dim strActivityLabel As String

        If CurrentActivity.IsProcess = True Then
            strActivityLabel = LANG_Process
            TabControlActivity.TabPages.Remove(TabPageDuration)
            TabControlActivity.TabPages.Remove(TabPagePreparation)
            TabControlActivity.TabPages.Remove(TabPageRepeat)
        Else
            strActivityLabel = LANG_Activity
        End If
        Me.lblActivity.Text = String.Format("{0}:", strActivityLabel)
    End Sub

    Private Sub SelectDateSystem()
        If CurrentActivity.ActivityDetail.Relative = True Then
            rbtnReferenceDate.Checked = True
            rbtnExactDate.Checked = False
        Else
            rbtnReferenceDate.Checked = False
            rbtnExactDate.Checked = True
        End If
        ShowStartDate()
    End Sub

    Private Sub ShowStartDate()
        If CurrentActivity.ActivityDetail.Relative = False Then
            If CurrentActivity.ActivityDetail.StartDate > Date.MinValue And CurrentActivity.ActivityDetail.StartDate < Date.MaxValue Then
                'Dim datTmp As Date
                dtbStartDate.Value = CurrentActivity.ActivityDetail.StartDate
            Else
                dtbStartDate.Value = Now
            End If
        Else
            dtbStartDate.Value = Now
        End If
    End Sub

    Public Sub LoadRepeatDates()
        lvRepeatDates.Items.Clear()

        With CurrentActivity.ActivityDetail
            If .RepeatStartDates.Count > 0 Then
                Dim tsPreparation, tsFollowUp As TimeSpan
                Dim datPreparationStart, datFollowUpEnd As Date

                If .PreparationRepeat = True Then
                    tsPreparation = CurrentActivity.ActivityDetail.PreparationPeriod

                End If
                If .FollowUpRepeat = True Then
                    tsFollowUp = CurrentActivity.ActivityDetail.FollowUpPeriod
                End If
                For i = 0 To .RepeatStartDates.Count - 1
                    Dim newItem As New ListViewItem(.RepeatStartDates(i))
                    newItem.SubItems.Add(.RepeatEndDates(i))
                    If .PreparationRepeat = True Then
                        datPreparationStart = .RepeatStartDates(i).Subtract(tsPreparation)
                        newItem.SubItems.Add(datPreparationStart)
                    End If
                    If .FollowUpRepeat = True Then
                        datFollowUpEnd = .RepeatEndDates(i).Add(tsFollowUp)
                        newItem.SubItems.Add(datFollowUpEnd)
                    End If
                    lvRepeatDates.Items.Add(newItem)
                Next
                UpdatePeriods()
            End If
        End With
    End Sub

    Public Sub UpdatePeriods()
        ActivityDetailBindingSource.ResetBindings(False)
        ActivityBindingSource.ResetBindings(False)
        SelectDateSystem()
    End Sub

    Private Sub TabControlActivity_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlActivity.SelectedIndexChanged
        CurrentActivitiesTab = TabControlActivity.SelectedIndex
    End Sub
#End Region

#Region "Tabpage Duration"
    Private Sub rbtnExactDate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDate.MouseUp
        SetRelativeUndoBuffer()
        If rbtnExactDate.Checked = True Then CurrentActivity.ActivityDetail.Relative = False Else CurrentActivity.ActivityDetail.Relative = True
    End Sub

    Private Sub rbtnReferenceDate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDate.MouseUp
        SetRelativeUndoBuffer()
        CurrentActivity.ActivityDetail.Relative = rbtnReferenceDate.Checked

    End Sub

    Private Sub ntbPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriod.Enter
        If CurrentActivity.ActivityDetail.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentActivity.ActivityDetail.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub ntbPeriod_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriod.Validated
        UpdatePeriods()
    End Sub

    Private Sub cmbPeriodUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnit.Enter
        If CurrentActivity.ActivityDetail.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentActivity.ActivityDetail.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodUnit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnit.Validated
        UpdatePeriods()
    End Sub

    Private Sub cmbPeriodDirection_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodDirection.Enter
        If CurrentActivity.ActivityDetail.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentActivity.ActivityDetail.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodDirection_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodDirection.Validated
        UpdatePeriods()
    End Sub

    Private Sub cmbGuidReferenceMoment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGuidReferenceMoment.Enter
        If CurrentActivity.ActivityDetail.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentActivity.ActivityDetail.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub cmbGuidReferenceMoment_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGuidReferenceMoment.Validated
        UpdatePeriods()
    End Sub

    Private Sub dtbStartDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtbStartDate.Enter
        If CurrentActivity.ActivityDetail.Relative = True Then
            SetRelativeUndoBuffer()
            CurrentActivity.ActivityDetail.Relative = False
        End If

        UndoRedo.UndoBuffer_Initialise(CurrentActivity.ActivityDetail, "StartDate")

        boolHasFocus = True

        SelectDateSystem()
    End Sub

    Private Sub dtbStartDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtbStartDate.Validated
        If boolHasFocus = True Then
            Dim datTmp As Date
            If Date.TryParse(dtbStartDate.Value, datTmp) Then
                With CurrentActivity.ActivityDetail
                    .Relative = False
                    .Period = 0
                    .PeriodDirection = ActivityDetail.DirectionValues.NotDefined
                    .PeriodUnit = ActivityDetail.DurationUnits.NotDefined
                    .GuidReferenceMoment = Guid.Empty
                    .StartDate = datTmp
                End With
            End If
            UpdatePeriods()

            boolHasFocus = False
        End If
    End Sub

    Private Sub SetRelativeUndoBuffer()
        Dim objNewValue, objOldValue As Object

        If CurrentActivity.ActivityDetail.Relative = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        UndoRedo.DateRelativeChanged(CurrentActivity.ActivityDetail, "Relative", objOldValue, objNewValue)
    End Sub

    Private Sub ntbDuration_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbDuration.Validated
        UpdatePeriods()
    End Sub

    Private Sub cmbDurationUnit_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDurationUnit.SelectedValueChanged
        If chkDurationUntilEnd.Checked = True Then ntbDuration.DataBindings(0).ReadValue()
    End Sub

    Private Sub cmbDurationUnit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDurationUnit.Validated
        UpdatePeriods()
    End Sub

    Private Sub chkDurationUntilEnd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDurationUntilEnd.CheckedChanged
        ntbDuration.DataBindings(0).ReadValue()
        cmbDurationUnit.DataBindings(0).ReadValue()
    End Sub

    Private Sub chkDurationUntilEnd_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDurationUntilEnd.Validated
        UpdatePeriods()
    End Sub
#End Region

#Region "Tabpage Preparation & Follow-up"
    Private Sub ntbPreparation_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPreparation.Validated
        UpdatePeriods()
    End Sub

    Private Sub cmbPreparationUnit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPreparationUnit.Validated
        If chkPreparationFromStart.Checked = True Then ntbPreparation.DataBindings(0).ReadValue()
        UpdatePeriods()
    End Sub

    Private Sub chkPreparationFromStart_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPreparationFromStart.CheckedChanged
        ntbPreparation.DataBindings(0).ReadValue()
        cmbPreparationUnit.DataBindings(0).ReadValue()
    End Sub

    Private Sub chkPreparationFromStart_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPreparationFromStart.Validated
        UpdatePeriods()
    End Sub

    Private Sub ntbFollowUp_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbFollowUp.Validated
        UpdatePeriods()
    End Sub

    Private Sub cmbFollowUpUnit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFollowUpUnit.Validated
        If chkFollowUpUntilEnd.Checked = True Then ntbFollowUp.DataBindings(0).ReadValue()
        UpdatePeriods()
    End Sub

    Private Sub chkFollowUpUntilEnd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFollowUpUntilEnd.CheckedChanged
        ntbFollowUp.DataBindings(0).ReadValue()
        cmbFollowUpUnit.DataBindings(0).ReadValue()
    End Sub

    Private Sub chkFollowUpUntilEnd_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFollowUpUntilEnd.Validated
        UpdatePeriods()
    End Sub
#End Region

#Region "Tabpage Repeat"
    Private Sub ntbRepeatEvery_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbRepeatEvery.Validated
        If CurrentActivity IsNot Nothing Then
            CurrentActivity.ActivityDetail.CalculateRepeatDates()
            LoadRepeatDates()
        End If
    End Sub

    Private Sub cmbRepeatUnit_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbRepeatUnit.SelectedValueChanged
        If chkRepeatUntilEnd.Checked = True Then ntbRepeatEvery.DataBindings(0).ReadValue()
        If CurrentActivity IsNot Nothing Then
            CurrentActivity.ActivityDetail.CalculateRepeatDates()
            LoadRepeatDates()
        End If
    End Sub

    Private Sub ntbRepeatTimes_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbRepeatTimes.Validated
        If CurrentActivity IsNot Nothing Then
            CurrentActivity.ActivityDetail.CalculateRepeatDates()
            LoadRepeatDates()
        End If
    End Sub

    Private Sub chkRepeatUntilEnd_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRepeatUntilEnd.Validated
        If CurrentActivity IsNot Nothing Then
            CurrentActivity.ActivityDetail.CalculateRepeatDates()
            LoadRepeatDates()
        End If
    End Sub
#End Region

#Region "Tabpage Activities"
    Private Sub lvDetailProcesses_ActivityModified() Handles lvDetailProcesses.ActivityModified
        CurrentProjectForm.dgvLogframe.Reload()
    End Sub

    Private Sub lvDetailProcesses_Updated() Handles lvDetailProcesses.Updated
        ProcessOrActivity()
        LoadRepeatDates()
    End Sub
#End Region

End Class
