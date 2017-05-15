Imports System.Windows.Forms

Public Class DialogTargetGroup
    Private objTargetGroup As TargetGroup
    Private bindTargetGroup As New BindingSource
    Private bindProperties As New BindingSource
    Private boolUserEdit As Boolean
    Private UserSelectedGuid As Boolean

    Friend WithEvents lvProperties As New ListViewTargetGroupInformations

    Public Property TargetGroup() As TargetGroup
        Get
            Return objTargetGroup
        End Get
        Set(ByVal value As TargetGroup)
            objTargetGroup = value
        End Set
    End Property

    Public Sub New(ByVal targetgroup As TargetGroup)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.TargetGroup = targetgroup
        If targetgroup IsNot Nothing Then
            bindTargetGroup.DataSource = Me.TargetGroup
            bindProperties.DataSource = Me.TargetGroup.TargetGroupInformations

            Dim ListPurposes As New Dictionary(Of Guid, String)
            Dim strPurpose As String
            Dim intIndex As Integer = 1

            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                strPurpose = String.Format("{0} {1}", CurrentLogFrame.GetStructSortNumber(selPurpose), selPurpose.Text)
                ListPurposes.Add(selPurpose.Guid, strPurpose)
                intIndex += 1
            Next

            Me.tbName.DataBindings.Add("Text", bindTargetGroup, "Name")

            With Me.cmbPurpose
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = New BindingSource(ListPurposes, Nothing)
                .ValueMember = "Key"
                .DisplayMember = "Value"
                .DataBindings.Add("SelectedValue", bindTargetGroup, "ParentPurposeGuid")
            End With

            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_TargetGroupTypes)
                .DataBindings.Add("SelectedIndex", bindTargetGroup, "Type")
            End With

            With Me.ntbNumber
                .DataBindings.Add("Text", bindTargetGroup, "Number", True)
                .DataBindings(0).FormatString = "#,##0"
            End With

            With Me.ntbNumberOfFemales
                .DataBindings.Add("Text", bindTargetGroup, "NumberOfFemales", True)
                .DataBindings(0).FormatString = "#,##0"
                .AllowSpace = True

            End With

            With Me.ntbNumberOfMales
                .DataBindings.Add("Text", bindTargetGroup, "NumberOfMales", True)
                .DataBindings(0).FormatString = "#,##0"
            End With

            With Me.ntbNumberOfPeople
                .DataBindings.Add("Text", bindTargetGroup, "NumberOfPeople", True)
                .DataBindings(0).FormatString = "#,##0"
            End With

            Me.tbLocation.DataBindings.Add("Text", bindTargetGroup, "Location")

            lvProperties.TargetGroupInformations = Me.TargetGroup.TargetGroupInformations
            lvProperties.Dock = DockStyle.Fill
            With Me.PanelProperties.Controls
                .Add(lvProperties)
                .SetChildIndex(lvProperties, 0)
            End With
        End If

    End Sub

    Private Sub btnReady_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReady.Click
        If TargetGroup Is Nothing Then
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        ElseIf TargetGroup IsNot Nothing AndAlso TargetGroup.ParentPurposeGuid = Guid.Empty Then
            If CurrentLogFrame.Purposes.Count > 0 Then
                TargetGroup.ParentPurposeGuid = CurrentLogFrame.Purposes(0).Guid
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Else
                Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            End If

            Me.Close()
        End If
        
    End Sub

    Private Sub cmbPurpose_Enter(sender As Object, e As System.EventArgs) Handles cmbPurpose.Enter
        CurrentControl = cmbPurpose
    End Sub

    Private Sub cmbPurpose_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles cmbPurpose.MouseDown
        Me.UserSelectedGuid = True

        CurrentControl = cmbPurpose
    End Sub

    Private Sub cmbPurpose_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbPurpose.SelectedValueChanged
        If CurrentControl Is cmbPurpose And Me.UserSelectedGuid = True Then
            If cmbPurpose.SelectedValue.GetType = GetType(Guid) Then
                Dim selGuid As Guid = cmbPurpose.SelectedValue
                Dim ParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selGuid)

                UndoRedo.ItemParentChanged(Me.TargetGroup, ParentPurpose.TargetGroups)

                Me.UserSelectedGuid = False
            End If
        End If
    End Sub

    Private Sub cmbPurpose_Validated(sender As Object, e As System.EventArgs) Handles cmbPurpose.Validated
        Me.UserSelectedGuid = False
    End Sub

    Private Sub cmbType_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbType.Enter
        boolUserEdit = True
    End Sub

    Private Sub cmbType_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbType.Leave
        boolUserEdit = False
    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        Select Case cmbType.SelectedIndex
            Case TargetGroup.TargetGroupTypes.Individual, TargetGroup.TargetGroupTypes.Other
                lblNumber.Text = ToLabel(LANG_NumberBeneficiaries)
            Case TargetGroup.TargetGroupTypes.Family, TargetGroup.TargetGroupTypes.ExtendedFamily
                lblNumber.Text = ToLabel(LANG_NumberFamilies)
            Case TargetGroup.TargetGroupTypes.Community
                lblNumber.Text = ToLabel(LANG_NumberCommunities)
            Case TargetGroup.TargetGroupTypes.Association
                lblNumber.Text = ToLabel(LANG_NumberAssociations)
            Case TargetGroup.TargetGroupTypes.Enterprise
                lblNumber.Text = ToLabel(LANG_NumberEnterprises)
            Case TargetGroup.TargetGroupTypes.Authority, TargetGroup.TargetGroupTypes.LocalAuthority
                lblNumber.Text = ToLabel(LANG_NumberAuthorities)
        End Select

        If boolUserEdit = True Then
            Dim intIndex As Integer = cmbType.SelectedIndex
            If intIndex > 0 Then
                Me.TargetGroup.TargetGroupInformations.SetDefaultInformations(cmbType.SelectedIndex)
            End If
        End If
        lvProperties.LoadItems()
    End Sub

    Private Sub ntbTotalNumber_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbNumber.Validated
        Select Case cmbType.SelectedIndex
            Case TargetGroup.TargetGroupTypes.Individual
                ntbNumberOfPeople.Text = ntbNumber.Text
        End Select
    End Sub

    Private Sub ntbNumberOfFemales_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbNumberOfFemales.Validating
        Dim strText As String = ntbNumberOfFemales.Text
        Dim intNrFemales As Integer

        If strText.Contains("%") Then
            strText = strText.Replace("%", "")
            strText = Trim(strText)
            Dim sngValue As Single
            If Single.TryParse(strText, sngValue) = True Then
                If Me.TargetGroup.NumberOfPeople > 0 Then
                    intNrFemales = Int(Me.TargetGroup.NumberOfPeople / 100 * sngValue)
                    ntbNumberOfFemales.Text = (intNrFemales).ToString
                    Me.TargetGroup.NumberOfMales = Me.TargetGroup.NumberOfPeople - intNrFemales
                Else
                    ntbNumberOfFemales.Text = "0"
                End If
            End If

        Else
            Me.TargetGroup.NumberOfPeople = ntbNumberOfFemales.SingleValue + ntbNumberOfMales.SingleValue
            If TargetGroup.Type = TargetGroup.TargetGroupTypes.Individual Then TargetGroup.Number = TargetGroup.NumberOfPeople
        End If
    End Sub

    Private Sub ntbNumberOfMales_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbNumberOfMales.Validating
        Dim strText As String = ntbNumberOfMales.Text
        Dim intNrMales As Integer

        If strText.Contains("%") Then
            strText = strText.Replace("%", "")
            strText = Trim(strText)
            Dim sngValue As Single
            If Single.TryParse(strText, sngValue) = True Then
                If Me.TargetGroup.NumberOfPeople > 0 Then
                    intNrMales = Int(Me.TargetGroup.NumberOfPeople / 100 * sngValue)
                    ntbNumberOfMales.Text = (intNrMales).ToString
                    Me.TargetGroup.NumberOfFemales = Me.TargetGroup.NumberOfPeople - intNrMales
                Else
                    ntbNumberOfMales.Text = "0"
                End If
            End If

        Else
            Me.TargetGroup.NumberOfPeople = ntbNumberOfMales.SingleValue + ntbNumberOfFemales.SingleValue
            If TargetGroup.Type = TargetGroup.TargetGroupTypes.Individual Then TargetGroup.Number = TargetGroup.NumberOfPeople
        End If
    End Sub
    
    Private Sub ToolStripButtonNewProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewProperty.Click
        lvProperties.NewItem()
    End Sub

    Private Sub ToolStripButtonEditProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditProperty.Click
        lvProperties.EditItem()
    End Sub

    Private Sub ToolStripButtonDeleteProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonDeleteProperty.Click
        lvProperties.RemoveItem()
    End Sub

    Private Sub ToolStripButtonPrintIdForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonPrintIdForm.Click
        Dim prntIdForm As New PrintTargetGroupIdForm(Me.TargetGroup)

        Dim dlgPrintPreview As New PrintPreviewDialog
        With dlgPrintPreview
            .ClientSize = New System.Drawing.Size(600, 800)
            .Document = prntIdForm
            .WindowState = FormWindowState.Maximized
            .ShowDialog()
        End With
    End Sub
End Class
