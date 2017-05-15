Imports System.Windows.Forms

Public Class DialogKeyMoment
    Private KeyMomentBindingSource As New BindingSource
    Private objCurrentKeyMoment As KeyMoment
    Private boolHasFocus As Boolean
    Private UserSelectedGuid As Boolean

    Public Property CurrentKeyMoment() As KeyMoment
        Get
            Return objCurrentKeyMoment
        End Get
        Set(ByVal value As KeyMoment)
            objCurrentKeyMoment = value
        End Set
    End Property

    Public Sub New(ByVal keymoment As KeyMoment)
        InitializeComponent()
        CurrentKeyMoment = keymoment

        LoadItems()
    End Sub

    Public Sub LoadItems()
        If CurrentKeyMoment IsNot Nothing Then
            KeyMomentBindingSource.DataSource = CurrentKeyMoment

            Dim ListOutputs As New Dictionary(Of Guid, String)
            Dim strOutput As String
            Dim intIndex As Integer = 1

            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    strOutput = intIndex.ToString & ".  " & selOutput.Text
                    ListOutputs.Add(selOutput.Guid, strOutput)
                    intIndex += 1
                Next
            Next
            If ListOutputs.Count = 0 Then
                ListOutputs.Add(Guid.NewGuid, "{" & Output.ItemName & "}")
            End If

            With Me.cmbOutput
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = New BindingSource(ListOutputs, Nothing)
                .ValueMember = "Key"
                .DisplayMember = "Value"
                .DataBindings.Add("SelectedValue", KeyMomentBindingSource, "ParentOutputGuid")
            End With

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

    Private Sub SetRelativeUndoBuffer()
        Dim objNewValue, objOldValue As Object

        If CurrentKeyMoment.Relative = True Then
            objOldValue = True
            objNewValue = False
        Else
            objOldValue = False
            objNewValue = True
        End If

        UndoRedo.DateRelativeChanged(CurrentKeyMoment, "Relative", objOldValue, objNewValue)
    End Sub

    Private Sub cmbOutput_Enter(sender As Object, e As System.EventArgs) Handles cmbOutput.Enter
        CurrentControl = cmbOutput
    End Sub

    Private Sub cmbOutput_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles cmbOutput.MouseDown
        Me.UserSelectedGuid = True

        CurrentControl = cmbOutput
    End Sub

    Private Sub cmbOutput_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbOutput.SelectedValueChanged
        If CurrentControl Is cmbOutput And Me.UserSelectedGuid = True Then
            If cmbOutput.SelectedValue.GetType = GetType(Guid) Then
                Dim selGuid As Guid = cmbOutput.SelectedValue
                Dim ParentOutput As Output = CurrentLogFrame.GetOutputByGuid(selGuid)

                UndoRedo.ItemParentChanged(CurrentKeyMoment, ParentOutput.KeyMoments)

                Me.UserSelectedGuid = False
            End If
        End If
    End Sub

    Private Sub cmbOutput_Validated(sender As Object, e As System.EventArgs) Handles cmbOutput.Validated
        Me.UserSelectedGuid = False
    End Sub

    Private Sub rbtnExactDate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnExactDate.MouseUp
        SetRelativeUndoBuffer()
        If rbtnExactDate.Checked = True Then CurrentKeyMoment.Relative = False Else CurrentKeyMoment.Relative = True
    End Sub

    Private Sub rbtnReferenceDate_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnReferenceDate.MouseUp
        SetRelativeUndoBuffer()
        CurrentKeyMoment.Relative = rbtnReferenceDate.Checked

    End Sub

    Private Sub dtpKeyMoment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpKeyMoment.Enter
        If CurrentKeyMoment.Relative = True Then
            SetRelativeUndoBuffer()
            CurrentKeyMoment.Relative = False
        End If

        UndoRedo.UndoBuffer_Initialise(CurrentKeyMoment, "KeyMoment")

        boolHasFocus = True

        SelectDateSystem()
    End Sub

    Private Sub dtpKeyMoment_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpKeyMoment.Validated
        If dtpKeyMoment.Value > Date.MinValue AndAlso dtpKeyMoment.Value <> CurrentKeyMoment.KeyMoment Then
            CurrentKeyMoment.KeyMoment = dtpKeyMoment.Value
        End If

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

    Private Sub btnReady_Click(sender As System.Object, e As System.EventArgs) Handles btnReady.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class
