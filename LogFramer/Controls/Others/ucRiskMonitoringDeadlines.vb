Public Class ucRiskMonitoringDeadlines
    Private objRiskMonitoring As RiskMonitoring
    Private boolHasFocus As Boolean

    Public Property RiskMonitoring As RiskMonitoring
        Get
            Return objRiskMonitoring
        End Get
        Set(ByVal value As RiskMonitoring)
            objRiskMonitoring = value
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub LoadItems()
        If Me.RiskMonitoring IsNot Nothing Then

            With cmbRepetition
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_RiskMonitoringRepetitionOptions)
            End With

            With cmbPeriodUnitStart
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
            End With

            With cmbPeriodUnitEnd
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_DurationUnits)
            End With

            Reload()
        End If
    End Sub

    Public Sub Reload()
        If Me.RiskMonitoring IsNot Nothing Then

            cmbRepetition.SelectedIndex = RiskMonitoring.Repetition
            ntbPeriodStart.Text = RiskMonitoring.PeriodStart
            cmbPeriodUnitStart.SelectedIndex = RiskMonitoring.PeriodUnitStart
            ntbPeriodEnd.Text = RiskMonitoring.PeriodEnd
            cmbPeriodUnitEnd.SelectedIndex = RiskMonitoring.PeriodUnitEnd

            Select Case RiskMonitoring.Repetition
                Case RiskMonitoring.RepetitionOptions.SingleMoment
                    gbStarting.Enabled = False
                Case Else
                    gbStarting.Enabled = True
            End Select

            SelectDateSystem()
            lvRiskMonitoringDeadlines.RiskMonitoring = Me.RiskMonitoring
            lvRiskMonitoringDeadlines.LoadItems()
        End If
    End Sub

    Private Sub SelectDateSystem()
        If RiskMonitoring.RelativeStart = True Then
            rbtnReferenceDateStart.Checked = True
            rbtnExactDateStart.Checked = False
        Else
            rbtnReferenceDateStart.Checked = False
            rbtnExactDateStart.Checked = True
        End If
        ShowStartDate()

        If RiskMonitoring.RelativeEnd = True Then
            rbtnReferenceDateEnd.Checked = True
            rbtnExactDateEnd.Checked = False
        Else
            rbtnReferenceDateEnd.Checked = False
            rbtnExactDateEnd.Checked = True
        End If
        ShowEndDate()
    End Sub

    Private Sub cmbRepetition_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbRepetition.SelectedIndexChanged
        Dim intRepetition As Integer = cmbRepetition.SelectedIndex

        If intRepetition <> RiskMonitoring.Repetition Then
            RiskMonitoring.Repetition = intRepetition

            RiskMonitoring.SetRiskMonitoringDeadlines()
            Reload()
        End If
    End Sub

    Private Sub ShowStartDate()
        If RiskMonitoring.RelativeStart = False Then
            If RiskMonitoring.StartDate > Date.MinValue And RiskMonitoring.StartDate < Date.MaxValue Then
                Dim datTmp As Date
                If Date.TryParse(RiskMonitoring.StartDate, datTmp) = True Then _
                    dtpStartDate.Value = datTmp
            Else
                dtpStartDate.Value = Now
            End If
        Else
            dtpStartDate.Value = Now
        End If
    End Sub

    Private Sub ntbPeriodStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodStart.Enter
        If RiskMonitoring.RelativeStart = False Then
            SetRelativeUndoBufferStart()
            RiskMonitoring.RelativeStart = True
        End If

        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodUnitStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitStart.Enter
        If RiskMonitoring.RelativeStart = False Then
            SetRelativeUndoBufferStart()
            RiskMonitoring.RelativeStart = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub SetRelativeUndoBufferStart()
        Dim objNewValue, objOldValue As Object

        If RiskMonitoring.RelativeStart = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        'CurrentUndoList.ChangeValueOperation(RiskMonitoring, "RelativeStart", LANG_DateRelativeExact, objNewValue, objOldValue, True)
    End Sub

    Private Sub dtpStartDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.Enter
        If RiskMonitoring.RelativeStart = True Then
            SetRelativeUndoBufferStart()
            RiskMonitoring.RelativeStart = False
        End If

        'CurrentUndoList.ChangeDateOperation_BeforeChange(RiskMonitoring, "StartDate", LANG_Date.ToLower, dtpStartDate.Value)
        boolHasFocus = True

        SelectDateSystem()
    End Sub

    Private Sub dtpStartDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.ValueChanged
        'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeDate
    End Sub

    Private Sub dtpStartDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.Validated
        If boolHasFocus = True Then
            Dim datTmp As Date
            If Date.TryParse(dtpStartDate.Value, datTmp) = True Then
                RiskMonitoring.StartDate = datTmp
                RiskMonitoring.SetRiskMonitoringDeadlines()
                Reload()
            End If
        End If
        'CurrentUndoList.ChangeDateOperation_AfterChange(dtpStartDate.Value, True)

        boolHasFocus = False
    End Sub

    Private Sub cmbPeriodUnitStart_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitStart.SelectedIndexChanged
        Dim intPeriodUnit As Integer = cmbPeriodUnitStart.SelectedIndex

        If intPeriodUnit <> RiskMonitoring.PeriodUnitStart Then
            RiskMonitoring.PeriodUnitStart = intPeriodUnit
            RiskMonitoring.SetRiskMonitoringDeadlines()
            Reload()
        End If
    End Sub

    Private Sub rbtnExactDateStart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDateStart.MouseUp
        If rbtnExactDateStart.Checked = True Then RiskMonitoring.RelativeStart = False Else RiskMonitoring.RelativeStart = True
        RiskMonitoring.SetRiskMonitoringDeadlines()
        Reload()
    End Sub

    Private Sub rbtnReferenceDateStart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDateStart.MouseUp
        RiskMonitoring.RelativeStart = rbtnReferenceDateStart.Checked
        RiskMonitoring.SetRiskMonitoringDeadlines()
        Reload()
    End Sub

    Private Sub ntbPeriodStart_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodStart.Validated
        Dim intPeriod As Integer = ntbPeriodStart.IntegerValue
        If RiskMonitoring.PeriodStart <> intPeriod Then
            RiskMonitoring.PeriodStart = intPeriod
            RiskMonitoring.SetRiskMonitoringDeadlines()
            Reload()
        End If
    End Sub

    Private Sub ShowEndDate()
        If RiskMonitoring.RelativeEnd = False Then
            If RiskMonitoring.EndDate > Date.MinValue And RiskMonitoring.EndDate < Date.MaxValue Then
                Dim datTmp As Date
                If Date.TryParse(RiskMonitoring.EndDate, datTmp) = True Then _
                    dtpEndDate.Value = datTmp
            Else
                dtpEndDate.Value = Now
            End If
        Else
            dtpEndDate.Value = Now
        End If
    End Sub

    Private Sub ntbPeriodEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodEnd.Enter
        If RiskMonitoring.RelativeEnd = False Then
            SetRelativeUndoBufferEnd()
            RiskMonitoring.RelativeEnd = True
        End If

        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodUnitEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitEnd.Enter
        If RiskMonitoring.RelativeEnd = False Then
            SetRelativeUndoBufferEnd()
            RiskMonitoring.RelativeEnd = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub SetRelativeUndoBufferEnd()
        Dim objNewValue, objOldValue As Object

        If RiskMonitoring.RelativeEnd = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        'CurrentUndoList.ChangeValueOperation(RiskMonitoring, "RelativeEnd", LANG_DateRelativeExact, objNewValue, objOldValue, True)
    End Sub

    Private Sub dtpEndDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Enter
        If RiskMonitoring.RelativeEnd = True Then
            SetRelativeUndoBufferEnd()
            RiskMonitoring.RelativeEnd = False
        End If

        'CurrentUndoList.ChangeDateOperation_BeforeChange(RiskMonitoring, "EndDate", LANG_Date.ToLower, dtpEndDate.Value)
        boolHasFocus = True

        SelectDateSystem()
    End Sub

    Private Sub dtpEndDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.ValueChanged
        'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeDate
    End Sub

    Private Sub dtpEndDate_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Validated
        If boolHasFocus = True Then
            Dim datTmp As Date
            If Date.TryParse(dtpEndDate.Value, datTmp) = True Then
                RiskMonitoring.EndDate = datTmp
                RiskMonitoring.SetRiskMonitoringDeadlines()
                Reload()
            End If
        End If
        'CurrentUndoList.ChangeDateOperation_AfterChange(dtpEndDate.Value, True)

        boolHasFocus = False
    End Sub

    Private Sub cmbPeriodUnitEnd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitEnd.SelectedIndexChanged
        Dim intPeriodUnit As Integer = cmbPeriodUnitEnd.SelectedIndex

        If intPeriodUnit <> RiskMonitoring.PeriodUnitEnd Then
            RiskMonitoring.PeriodUnitEnd = intPeriodUnit
            RiskMonitoring.SetRiskMonitoringDeadlines()
            Reload()
        End If
    End Sub

    Private Sub rbtnExactDateEnd_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDateEnd.MouseUp
        If rbtnExactDateEnd.Checked = True Then RiskMonitoring.RelativeEnd = False Else RiskMonitoring.RelativeEnd = True
        RiskMonitoring.SetRiskMonitoringDeadlines()
        Reload()
    End Sub

    Private Sub rbtnReferenceDateEnd_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDateEnd.MouseUp
        RiskMonitoring.RelativeEnd = rbtnReferenceDateEnd.Checked
        RiskMonitoring.SetRiskMonitoringDeadlines()
        Reload()
    End Sub

    Private Sub ntbPeriodEnd_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodEnd.Validated
        Dim intPeriod As Integer = ntbPeriodEnd.IntegerValue
        If RiskMonitoring.PeriodEnd <> intPeriod Then
            RiskMonitoring.PeriodEnd = intPeriod
            RiskMonitoring.SetRiskMonitoringDeadlines()
            Reload()
        End If
    End Sub

    Private Sub rbtnExactDateEnd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnExactDateEnd.CheckedChanged

    End Sub
End Class
