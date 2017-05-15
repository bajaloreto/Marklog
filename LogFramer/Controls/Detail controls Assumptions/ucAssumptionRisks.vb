Public Class AssumptionRisks
    Public AssumptionBindingSource As New BindingSource
    Private objCurrentAssumption As Assumption
    Private boolTypeChanged As Boolean
    Private boolTextSelected As Boolean

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal assumption As Assumption)
        InitializeComponent()

        Me.CurrentAssumption = Assumption
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Property CurrentAssumption() As Assumption
        Get
            Return objCurrentAssumption
        End Get
        Set(ByVal value As Assumption)
            objCurrentAssumption = value
        End Set
    End Property

    Public Property TextSelected() As Boolean
        Get
            Return boolTextSelected
        End Get
        Set(ByVal value As Boolean)
            boolTextSelected = value
            Me.Refresh()
        End Set
    End Property

    Public Sub LoadItems()
        If CurrentAssumption IsNot Nothing Then
            AssumptionBindingSource.DataSource = CurrentAssumption

            With cmbRiskCategory
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_RiskCategories
                .DataBindings.Add("SelectedIndex", AssumptionBindingSource, "RiskDetail.RiskCategory")
            End With

            tbRiskLevel.DataBindings.Add("Text", AssumptionBindingSource, "RiskDetail.RiskLevel", True)
            tbRiskLevel.DataBindings(0).FormatString = "P0"

            With cmbLikelihood
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_Likelihoods
                .DataBindings.Add("SelectedIndex", AssumptionBindingSource, "RiskDetail.Likelihood")
            End With
            With cmbRiskImpact
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_RiskImpacts
                .DataBindings.Add("SelectedIndex", AssumptionBindingSource, "RiskDetail.RiskImpact")
            End With
            tbImpact.DataBindings.Add("Text", AssumptionBindingSource, "Impact")

            With cmbRiskResponse
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_RiskResponses
                .DataBindings.Add("SelectedIndex", AssumptionBindingSource, "RiskDetail.RiskResponse")
            End With
            tbResponseStrategy.DataBindings.Add("Text", AssumptionBindingSource, "ResponseStrategy")
            With cmbOwner
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = CurrentLogframe.PartnerNamesList
                .DataBindings.Add("SelectedItem", AssumptionBindingSource, "Owner")
            End With
        End If
        cmbRiskCategory.Focus()
    End Sub

    Private Sub cmbOwner_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOwner.SelectedValueChanged
        cmbOwner.Select(0, 0)
    End Sub
End Class
