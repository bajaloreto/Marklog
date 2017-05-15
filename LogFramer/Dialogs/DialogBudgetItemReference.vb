Imports System.Windows.Forms

Public Class DialogBudgetItemReference
    Private BudgetItemReferenceBindingSource As New BindingSource
    Private objCurrentBudgetItemReference As BudgetItemReference
    Private objCurrentBudgetItem As BudgetItem
    Private boolHasFocus As Boolean

    Public Property CurrentBudgetItemReference() As BudgetItemReference
        Get
            Return objCurrentBudgetItemReference
        End Get
        Set(ByVal value As BudgetItemReference)
            objCurrentBudgetItemReference = value
        End Set
    End Property

    Public Property CurrentBudgetItem As BudgetItem
        Get
            Return objCurrentBudgetItem
        End Get
        Set(value As BudgetItem)
            objCurrentBudgetItem = value
        End Set
    End Property

    Public Sub New(ByVal budgetitemreference As BudgetItemReference)
        InitializeComponent()
        CurrentBudgetItemReference = BudgetItemReference
        CurrentBudgetItem = CurrentLogFrame.GetBudgetItemByGuid(CurrentBudgetItemReference.ReferenceBudgetItemGuid)
        LoadItems()
    End Sub

    Public Sub LoadItems()
        If CurrentBudgetItemReference IsNot Nothing Then
            BudgetItemReferenceBindingSource.DataSource = CurrentBudgetItemReference

            Dim ListActivities As Dictionary(Of Guid, String) = LoadActivitiesList()
            Dim ParentResource As Resource = CurrentLogFrame.GetResourceByGuid(CurrentBudgetItemReference.ParentResourceGuid)
            Dim ParentActivity As Activity
            Dim selGuid As Guid

            If ParentResource IsNot Nothing Then
                ParentActivity = CurrentLogFrame.GetActivityByGuid(ParentResource.ParentStructGuid)
                selGuid = ParentActivity.Guid
            End If

            With cmbActivity
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = New BindingSource(ListActivities, Nothing)
                .ValueMember = "Key"
                .DisplayMember = "Value"
                If selGuid <> Guid.Empty Then .SelectedValue = selGuid
            End With

            With ntbPercentage
                .ValueType = NumericTextBox.ValueTypes.DoubleValue
                .IsPercentage = True
                .DataBindings.Add("DoubleValue", BudgetItemReferenceBindingSource, "Percentage")
            End With

            With ntbTotalCost
                .IsCurrency = True
                .Unit = CurrentBudgetItemReference.TotalCost.CurrencyCode
                .DataBindings.Add("SingleValue", BudgetItemReferenceBindingSource, "TotalCost.Amount")
            End With
            

        End If
    End Sub

    Private Function LoadActivitiesList() As Dictionary(Of Guid, String)
        Dim ListActivities As New Dictionary(Of Guid, String)
        Dim intIndexPurpose, intIndexOutput As Integer
        Dim strParentSortPurpose, strParentSortOutput As String

        For Each selPurpose As Purpose In CurrentLogFrame.Purposes
            strParentSortPurpose = CurrentLogFrame.CreateSortNumber(intIndexPurpose)
            For Each selOutput As Output In selPurpose.Outputs
                strParentSortOutput = CurrentLogFrame.CreateSortNumber(intIndexOutput, strParentSortPurpose)

                ListActivities = LoadActivitiesList_Activities(ListActivities, selOutput.Activities, strParentSortOutput)

                intIndexOutput += 1
            Next
            intIndexPurpose += 1
        Next
        If ListActivities.Count = 0 Then
            ListActivities.Add(Guid.NewGuid, LANG_NotSelected)
        End If

        Return ListActivities
    End Function

    Private Function LoadActivitiesList_Activities(ByVal ListActivities As Dictionary(Of Guid, String), ByVal selActivities As Activities, ByVal strParentSort As String) As Dictionary(Of Guid, String)
        Dim strActivity, strSort As String
        Dim intIndex As Integer

        For Each selActivity As Activity In selActivities
            strSort = CurrentLogFrame.CreateSortNumber(intIndex, strParentSort)
            strActivity = String.Format("{0}    {1}", strSort, selActivity.Text)
            ListActivities.Add(selActivity.Guid, strActivity)

            If selActivity.Activities.Count > 0 Then
                ListActivities = LoadActivitiesList_Activities(ListActivities, selActivity.Activities, strSort)
            End If
            intIndex += 1
        Next

        Return ListActivities
    End Function

    Private Function LoadResourcesList(ByVal selResources As Resources) As Dictionary(Of Guid, String)
        Dim ListResources As New Dictionary(Of Guid, String)

        For Each selResource As Resource In selResources
            ListResources.Add(selResource.Guid, selResource.Text)
        Next

        If ListResources.Count = 0 Then
            ListResources.Add(Guid.NewGuid, LANG_NotSelected)
        End If


        Return ListResources
    End Function

    Private Sub cmbActivity_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbActivity.SelectedValueChanged
        Dim objGuid As Guid

        If cmbActivity.SelectedValue.GetType = GetType(Guid) Then
            objGuid = cmbActivity.SelectedValue
        ElseIf cmbActivity.SelectedValue.GetType = GetType(KeyValuePair(Of Guid, String)) Then
            Dim selValue As KeyValuePair(Of Guid, String) = cmbActivity.SelectedValue
            objGuid = selValue.Key
        Else
            Exit Sub
        End If

        If objGuid = Guid.Empty Then Exit Sub

        Dim selActivity As Activity = CurrentLogFrame.GetActivityByGuid(objGuid)
        Dim ListResources As Dictionary(Of Guid, String) = LoadResourcesList(selActivity.Resources)

        With Me.cmbResource
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataSource = New BindingSource(ListResources, Nothing)
            .ValueMember = "Key"
            .DisplayMember = "Value"
            .SelectedValue = CurrentBudgetItemReference.ParentResourceGuid
        End With
    End Sub

    Private Sub cmbResource_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbResource.SelectedValueChanged
        If cmbResource.SelectedValue Is Nothing Then Exit Sub
        Dim objGuid As Guid

        If cmbResource.SelectedValue.GetType = GetType(Guid) Then
            objGuid = cmbResource.SelectedValue
        ElseIf cmbResource.SelectedValue.GetType = GetType(KeyValuePair(Of Guid, String)) Then
            Dim selValue As KeyValuePair(Of Guid, String) = cmbResource.SelectedValue
            objGuid = selValue.Key
        Else
            Exit Sub
        End If

        CurrentBudgetItemReference.ParentResourceGuid = objGuid
    End Sub

    Private Sub ntbPercentage_Validated(sender As Object, e As System.EventArgs) Handles ntbPercentage.Validated
        SetTotalCost()
    End Sub

    Private Sub SetTotalCost()
        If CurrentBudgetItem IsNot Nothing Then
            Dim sngAmount As Single = CurrentBudgetItem.TotalCost.Amount
            sngAmount *= CurrentBudgetItemReference.Percentage

            CurrentBudgetItemReference.TotalCost.Amount = sngAmount

            If ntbTotalCost.DataBindings.Count > 0 Then _
                ntbTotalCost.DataBindings(0).ReadValue()
        End If
    End Sub

    Private Sub ntbTotalCost_Validated(sender As Object, e As System.EventArgs) Handles ntbTotalCost.Validated
        If CurrentBudgetItem IsNot Nothing Then
            Dim sngPercentage As Single

            If CurrentBudgetItem.TotalCost.Amount > 0 Then
                sngPercentage = CurrentBudgetItemReference.TotalCost.Amount / CurrentBudgetItem.TotalCost.Amount
            End If

            CurrentBudgetItemReference.Percentage = sngPercentage

            BudgetItemReferenceBindingSource.ResetBindings(False)

            If ntbPercentage.DataBindings.Count > 0 Then _
                ntbPercentage.DataBindings(0).ReadValue()
            'ntbPercentage.GetText()
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class
