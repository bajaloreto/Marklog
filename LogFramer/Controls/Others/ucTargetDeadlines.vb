Public Class ucTargetDeadlines
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private boolHasFocus As Boolean

    Public Property TargetDeadlinesSection As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesSection
        End Get
        Set(value As TargetDeadlinesSection)
            objTargetDeadlinesSection = value
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub LoadItems()
        If Me.TargetDeadlinesSection IsNot Nothing Then

            With cmbTargetSystem
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_TargetRepetitionOptions)
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
        If Me.TargetDeadlinesSection IsNot Nothing Then

            cmbTargetSystem.SelectedIndex = TargetDeadlinesSection.Repetition
            ntbPeriodStart.Text = TargetDeadlinesSection.PeriodStart
            cmbPeriodUnitStart.SelectedIndex = TargetDeadlinesSection.PeriodUnitStart
            ntbPeriodEnd.Text = TargetDeadlinesSection.PeriodEnd
            cmbPeriodUnitEnd.SelectedIndex = TargetDeadlinesSection.PeriodUnitEnd

            Select Case TargetDeadlinesSection.Repetition
                Case TargetDeadlinesSection.RepetitionOptions.SingleTarget
                    gbStarting.Enabled = False
                Case Else
                    gbStarting.Enabled = True
            End Select

            SelectDateSystem()
            lvTargetDeadlines.TargetDeadlinesSection = Me.TargetDeadlinesSection
            lvTargetDeadlines.LoadItems()
        End If
    End Sub

    Private Sub SelectDateSystem()
        If TargetDeadlinesSection.RelativeStart = True Then
            rbtnReferenceDateStart.Checked = True
            rbtnExactDateStart.Checked = False
        Else
            rbtnReferenceDateStart.Checked = False
            rbtnExactDateStart.Checked = True
        End If
        ShowStartDate()

        If TargetDeadlinesSection.RelativeEnd = True Then
            rbtnReferenceDateEnd.Checked = True
            rbtnExactDateEnd.Checked = False
        Else
            rbtnReferenceDateEnd.Checked = False
            rbtnExactDateEnd.Checked = True
        End If
        ShowEndDate()
    End Sub

    Private Sub cmbTargetSystem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTargetSystem.SelectedIndexChanged
        Dim intTargetSystem As Integer = cmbTargetSystem.SelectedIndex

        If intTargetSystem <> TargetDeadlinesSection.Repetition Then
            TargetDeadlinesSection.Repetition = intTargetSystem

            TargetDeadlinesSection.SetTargetDeadlines()
            Reload()
        End If
    End Sub

    Private Sub ShowStartDate()
        If TargetDeadlinesSection.RelativeStart = False Then
            If TargetDeadlinesSection.StartDate > Date.MinValue And TargetDeadlinesSection.StartDate < Date.MaxValue Then
                Dim datTmp As Date
                If Date.TryParse(TargetDeadlinesSection.StartDate, datTmp) = True Then _
                    dtpStartDate.Value = datTmp
            Else
                dtpStartDate.Value = Now
            End If
        Else
            dtpStartDate.Value = Now
        End If
    End Sub

    Private Sub ntbPeriodStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodStart.Enter
        If TargetDeadlinesSection.RelativeStart = False Then
            SetRelativeUndoBufferStart()
            TargetDeadlinesSection.RelativeStart = True
        End If

        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodUnitStart_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitStart.Enter
        If TargetDeadlinesSection.RelativeStart = False Then
            SetRelativeUndoBufferStart()
            TargetDeadlinesSection.RelativeStart = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub SetRelativeUndoBufferStart()
        Dim objNewValue, objOldValue As Object

        If TargetDeadlinesSection.RelativeStart = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        'CurrentUndoList.ChangeValueOperation(TargetDeadlinesSection, "RelativeStart", LANG_DateRelativeExact, objNewValue, objOldValue, True)
    End Sub

    Private Sub dtpStartDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpStartDate.Enter
        If TargetDeadlinesSection.RelativeStart = True Then
            SetRelativeUndoBufferStart()
            TargetDeadlinesSection.RelativeStart = False
        End If

        'CurrentUndoList.ChangeDateOperation_BeforeChange(TargetDeadlinesSection, "StartDate", LANG_Date.ToLower, dtpStartDate.Value)
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
                TargetDeadlinesSection.StartDate = datTmp
                TargetDeadlinesSection.SetTargetDeadlines()
                Reload()
            End If
        End If
        'CurrentUndoList.ChangeDateOperation_AfterChange(dtpStartDate.Value, True)

        boolHasFocus = False
    End Sub

    Private Sub cmbPeriodUnitStart_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitStart.SelectedIndexChanged
        Dim intPeriodUnit As Integer = cmbPeriodUnitStart.SelectedIndex

        If intPeriodUnit <> TargetDeadlinesSection.PeriodUnitStart Then
            TargetDeadlinesSection.PeriodUnitStart = intPeriodUnit
            TargetDeadlinesSection.SetTargetDeadlines()
            Reload()
        End If
    End Sub

    Private Sub rbtnExactDateStart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDateStart.MouseUp
        If rbtnExactDateStart.Checked = True Then TargetDeadlinesSection.RelativeStart = False Else TargetDeadlinesSection.RelativeStart = True
        TargetDeadlinesSection.SetTargetDeadlines()
        Reload()
    End Sub

    Private Sub rbtnReferenceDateStart_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDateStart.MouseUp
        TargetDeadlinesSection.RelativeStart = rbtnReferenceDateStart.Checked
        TargetDeadlinesSection.SetTargetDeadlines()
        Reload()
    End Sub

    Private Sub ntbPeriodStart_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodStart.Validated
        Dim intPeriod As Integer = ntbPeriodStart.IntegerValue
        If TargetDeadlinesSection.PeriodStart <> intPeriod Then
            TargetDeadlinesSection.PeriodStart = intPeriod
            TargetDeadlinesSection.SetTargetDeadlines()
            Reload()
        End If
    End Sub

    Private Sub ShowEndDate()
        If TargetDeadlinesSection.RelativeEnd = False Then
            If TargetDeadlinesSection.EndDate > Date.MinValue And TargetDeadlinesSection.EndDate < Date.MaxValue Then
                Dim datTmp As Date
                If Date.TryParse(TargetDeadlinesSection.EndDate, datTmp) = True Then _
                    dtpEndDate.Value = datTmp
            Else
                dtpEndDate.Value = Now
            End If
        Else
            dtpEndDate.Value = Now
        End If
    End Sub

    Private Sub ntbPeriodEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriodEnd.Enter
        If TargetDeadlinesSection.RelativeEnd = False Then
            SetRelativeUndoBufferEnd()
            TargetDeadlinesSection.RelativeEnd = True
        End If

        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodUnitEnd_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnitEnd.Enter
        If TargetDeadlinesSection.RelativeEnd = False Then
            SetRelativeUndoBufferEnd()
            TargetDeadlinesSection.RelativeEnd = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub SetRelativeUndoBufferEnd()
        Dim objNewValue, objOldValue As Object

        If TargetDeadlinesSection.RelativeEnd = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        'CurrentUndoList.ChangeValueOperation(TargetDeadlinesSection, "RelativeEnd", LANG_DateRelativeExact, objNewValue, objOldValue, True)
    End Sub

    Private Sub dtpEndDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEndDate.Enter
        If TargetDeadlinesSection.RelativeEnd = True Then
            SetRelativeUndoBufferEnd()
            TargetDeadlinesSection.RelativeEnd = False
        End If

        'CurrentUndoList.ChangeDateOperation_BeforeChange(TargetDeadlinesSection, "EndDate", LANG_Date.ToLower, dtpEndDate.Value)
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
                TargetDeadlinesSection.EndDate = datTmp
                TargetDeadlinesSection.SetTargetDeadlines()
                Reload()
            End If
        End If
        'CurrentUndoList.ChangeDateOperation_AfterChange(dtpEndDate.Value, True)

        boolHasFocus = False
    End Sub

    Private Sub cmbPeriodUnitEnd_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbPeriodUnitEnd.SelectedIndexChanged
        Dim intPeriodUnit As Integer = cmbPeriodUnitEnd.SelectedIndex

        If intPeriodUnit <> TargetDeadlinesSection.PeriodUnitEnd Then
            TargetDeadlinesSection.PeriodUnitEnd = intPeriodUnit
            TargetDeadlinesSection.SetTargetDeadlines()
            Reload()
        End If
    End Sub

    Private Sub rbtnExactDateEnd_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDateEnd.MouseUp
        If rbtnExactDateEnd.Checked = True Then TargetDeadlinesSection.RelativeEnd = False Else TargetDeadlinesSection.RelativeEnd = True
        TargetDeadlinesSection.SetTargetDeadlines()
        Reload()
    End Sub

    Private Sub rbtnReferenceDateEnd_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDateEnd.MouseUp
        TargetDeadlinesSection.RelativeEnd = rbtnReferenceDateEnd.Checked
        TargetDeadlinesSection.SetTargetDeadlines()
        Reload()
    End Sub

    Private Sub ntbPeriodEnd_Validated(sender As Object, e As System.EventArgs) Handles ntbPeriodEnd.Validated
        Dim intPeriod As Integer = ntbPeriodEnd.IntegerValue
        If TargetDeadlinesSection.PeriodEnd <> intPeriod Then
            TargetDeadlinesSection.PeriodEnd = intPeriod
            TargetDeadlinesSection.SetTargetDeadlines()
            Reload()
        End If
    End Sub
End Class
