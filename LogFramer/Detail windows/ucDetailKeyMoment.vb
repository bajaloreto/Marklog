Public Class DetailKeyMoment
    Private KeyMomentBindingSource As New BindingSource
    Private objCurrentKeyMoment As KeyMoment
    Private boolHasFocus As Boolean

    Public Property CurrentKeyMoment() As KeyMoment
        Get
            Return objCurrentKeyMoment
        End Get
        Set(ByVal value As KeyMoment)
            objCurrentKeyMoment = value
        End Set
    End Property

#Region "Initialise"
    Public Sub New(ByVal keymoment As KeyMoment)
        InitializeComponent()
        CurrentKeyMoment = keymoment

        LoadItems()
    End Sub

    Public Sub LoadItems()
        If CurrentKeyMoment IsNot Nothing Then
            KeyMomentBindingSource.DataSource = CurrentKeyMoment

            tbDescription.DataBindings.Add("Text", KeyMomentBindingSource, "Description")
            ntbPeriod.DataBindings.Add("Text", KeyMomentBindingSource, "Period")
            With cmbPeriodUnit
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_DurationUnits
                .DataBindings.Add("SelectedIndex", KeyMomentBindingSource, "PeriodUnit")
            End With
            With cmbPeriodDirection
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_DirectionValues
                .DataBindings.Add("SelectedIndex", KeyMomentBindingSource, "PeriodDirection")
            End With
            With cmbReferenceMoment
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = CurrentLogFrame.GetReferenceMomentsList
                .ValueMember = "Id"
                .DisplayMember = "Value"
                .DataBindings.Add("SelectedValue", KeyMomentBindingSource, "GuidReferenceMoment")
            End With
            dtbExactDateKeyMoment.DataBindings.Add("DateValue", KeyMomentBindingSource, "ExactDateKeyMoment", True)
            dtbExactDateKeyMoment.DataBindings(0).FormatString = "d"

            SelectDateSystem()
        End If
    End Sub

    Private Sub SelectDateSystem()
        If CurrentKeyMoment.Relative = True Then
            rbtnReferenceDate.Checked = True
            rbtnExactDate.Checked = False
        Else
            rbtnReferenceDate.Checked = False
            rbtnExactDate.Checked = True
        End If
        ShowStartDate()
    End Sub

    Private Sub ShowStartDate()
        If CurrentKeyMoment.Relative = False Then
            If CurrentKeyMoment.KeyMoment > Date.MinValue And CurrentKeyMoment.KeyMoment < Date.MaxValue Then
                Dim datTmp As Date
                If Date.TryParse(CurrentKeyMoment.KeyMoment, datTmp) = True Then _
                    dtpKeyMoment.Value = datTmp
            Else
                dtpKeyMoment.Value = Now
            End If
        Else
            dtpKeyMoment.Value = Now
        End If
    End Sub
#End Region

#Region "Methods & events"
    Private Sub dtpKeyMoment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpKeyMoment.Enter
        If CurrentKeyMoment.Relative = True Then
            SetRelativeUndoBuffer()
            CurrentKeyMoment.Relative = False
        End If

        'CurrentUndoList.ChangeDateOperation_BeforeChange(CurrentKeyMoment, "KeyMoment", LANG_Date.ToLower, dtpKeyMoment.Value)
        boolHasFocus = True

        SelectDateSystem()
    End Sub

    Private Sub dtpKeyMoment_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpKeyMoment.ValueChanged
        'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeDate
    End Sub

    Private Sub dtpKeyMoment_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpKeyMoment.Validated
        If boolHasFocus = True Then
            Dim datTmp As Date
            If Date.TryParse(dtpKeyMoment.Value, datTmp) = True Then _
                CurrentKeyMoment.KeyMoment = datTmp
        End If
        'CurrentUndoList.ChangeDateOperation_AfterChange(dtpKeyMoment.Value, True)

        boolHasFocus = False
    End Sub

    Private Sub ntbPeriod_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbPeriod.Enter
        If CurrentKeyMoment.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentKeyMoment.Relative = True
        End If

        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodUnit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodUnit.Enter
        If CurrentKeyMoment.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentKeyMoment.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub cmbPeriodDirection_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriodDirection.Enter
        If CurrentKeyMoment.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentKeyMoment.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub cmbReferenceMoment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbReferenceMoment.Enter
        If CurrentKeyMoment.Relative = False Then
            SetRelativeUndoBuffer()
            CurrentKeyMoment.Relative = True
        End If
        SelectDateSystem()
    End Sub

    Private Sub SetRelativeUndoBuffer()
        Dim objNewValue, objOldValue As Object

        If CurrentKeyMoment.Relative = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        'CurrentUndoList.ChangeValueOperation(CurrentKeyMoment, "Relative", LANG_DateRelativeExact, objNewValue, objOldValue, True)
    End Sub

    Private Sub rbtnExactDate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDate.MouseUp
        If rbtnExactDate.Checked = True Then CurrentKeyMoment.Relative = False Else CurrentKeyMoment.Relative = True
    End Sub

    Private Sub rbtnReferenceDate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDate.MouseUp
        CurrentKeyMoment.Relative = rbtnReferenceDate.Checked
    End Sub
#End Region
End Class
