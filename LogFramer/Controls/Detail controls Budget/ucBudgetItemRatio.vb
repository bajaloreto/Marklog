Public Class BudgetItemRatio
    Public Event RatioChanged()

    Private objCurrentBudgetItem As BudgetItem
    Private objBudgetItemReferences As New Dictionary(Of Guid, String)

    Private BudgetItemBindingSource As New BindingSource

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal budgetitem As BudgetItem, ByVal budgetitemreferences As Dictionary(Of Guid, String))
        InitializeComponent()

        Me.CurrentBudgetItem = budgetitem
        Me.BudgetItemReferences = budgetitemreferences
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Property CurrentBudgetItem() As BudgetItem
        Get
            Return objCurrentBudgetItem
        End Get
        Set(ByVal value As BudgetItem)
            objCurrentBudgetItem = value
        End Set
    End Property

    Public Property BudgetItemReferences As Dictionary(Of Guid, String)
        Get
            Return objBudgetItemReferences
        End Get
        Set(ByVal value As Dictionary(Of Guid, String))
            objBudgetItemReferences = value
        End Set
    End Property

    Private ReadOnly Property ReferenceGuid As Guid
        Get
            Dim intIndex As Integer = CurrentBudgetItem.Formula.IndexOf("*")
            Dim RefGuid As Guid = Nothing

            If intIndex > 0 Then
                Dim strGuid As String = CurrentBudgetItem.Formula.Substring(0, intIndex).Trim
                If Guid.TryParse(strGuid, RefGuid) Then
                    Return RefGuid
                Else
                    Return Nothing
                End If
            End If
        End Get
    End Property

    Private ReadOnly Property Ratio As Double
        Get
            Dim intIndex As Integer = CurrentBudgetItem.Formula.IndexOf("*")
            Dim dblRatio As Double

            If intIndex > 0 Then
                Dim strRatio As String = CurrentBudgetItem.Formula.Substring(intIndex).Trim
                dblRatio = ParseDouble(strRatio)
            End If

            Return dblRatio
        End Get
    End Property

    Public Sub LoadItems()
        If CurrentBudgetItem IsNot Nothing Then
            BudgetItemBindingSource.DataSource = CurrentBudgetItem

            With cmbReferenceRow
                .DisplayMember = "Value"
                .ValueMember = "Key"
                .DataSource = New BindingSource(BudgetItemReferences, Nothing)
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataBindings.Add("SelectedValue", BudgetItemBindingSource, "CalculationGuidRatio")

            End With

            ntbReferenceAmount_GetValue()

            With ntbRatio
                .DataBindings.Add("Text", BudgetItemBindingSource, "Ratio", True)
                .DataBindings(0).FormatString = "P2"
            End With

            With ntbTotalCost
                .DataBindings.Add("Text", BudgetItemBindingSource, "TotalCost.Amount", True)
                .DataBindings(0).FormatString = String.Format("#,##0.00 {0}", CurrentBudgetItem.TotalCost.CurrencyCode)
            End With
        End If
    End Sub

    Private Sub tbRatio_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbRatio.Validating
        Dim sngRatio As Single = ParseSingle(ntbRatio.Text)

        sngRatio /= 100
        ntbRatio.Text = sngRatio
    End Sub

    Private Sub tbRatio_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbRatio.Validated
        CalculateAmount()
    End Sub

    Private Sub ntbReferenceAmount_GetValue()
        Dim CalculationReference As BudgetItem = CurrentLogFrame.GetBudgetItemByGuid(CurrentBudgetItem.CalculationGuidRatio)

        If CalculationReference IsNot Nothing Then
            With ntbReferenceAmount
                .Text = CalculationReference.TotalCost.ToString
            End With
        End If
    End Sub

    Private Sub cmbReferenceRow_Validated(sender As Object, e As System.EventArgs) Handles cmbReferenceRow.Validated
        CalculateAmount()
    End Sub

    Private Sub CalculateAmount()
        CurrentBudgetItem.CalculateRatio()

        BudgetItemBindingSource.ResetBindings(False)
        ntbReferenceAmount_GetValue()

        RaiseEvent RatioChanged()
    End Sub
End Class
