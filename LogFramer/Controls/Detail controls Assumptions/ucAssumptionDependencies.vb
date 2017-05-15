Public Class AssumptionDependencies
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

            Dim DependencyTypes As List(Of StructuredComboBoxItem) = LoadDependencyTypesComboList()
            With cmbDependencyType
                .DataSource = DependencyTypes
                .ValueMember = "Type"
                .DisplayMember = "Description"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataBindings.Add("SelectedValue", AssumptionBindingSource, "DependencyDetail.DependencyType")
            End With

            Dim InputTypes As List(Of StructuredComboBoxItem) = LoadInputTypesComboList()
            With cmbInputType
                .DataSource = InputTypes
                .ValueMember = "Type"
                .DisplayMember = "Description"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataBindings.Add("SelectedValue", AssumptionBindingSource, "DependencyDetail.InputType")
            End With

            With cmbImportanceLevel
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = LIST_ImportanceLevels
                .DataBindings.Add("SelectedIndex", AssumptionBindingSource, "DependencyDetail.ImportanceLevel")
            End With
            tbImpact.DataBindings.Add("Text", AssumptionBindingSource, "Impact")

            With cmbSupplier
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = CurrentLogFrame.PartnerNamesList
                .DataBindings.Add("SelectedItem", AssumptionBindingSource, "DependencyDetail.Supplier")
            End With
            With cmbDeliverableType
                .DataSource = LIST_DeliverableTypes
                .ValueMember = "Id"
                .DisplayMember = "Value"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataBindings.Add("SelectedValue", AssumptionBindingSource, "DependencyDetail.DeliverableType")
            End With
            dtbDateExpected.DataBindings.Add("DateValue", AssumptionBindingSource, "DependencyDetail.DateExpected", True)
            dtbDateExpected.DataBindings(0).FormatString = "d"
            dtbDateDelivered.DataBindings.Add("DateValue", AssumptionBindingSource, "DependencyDetail.DateDelivered", True)
            dtbDateDelivered.DataBindings(0).FormatString = "d"
            tbDeliverables.DataBindings.Add("Text", AssumptionBindingSource, "DependencyDetail.Deliverables")

            tbResponseStrategy.DataBindings.Add("Text", AssumptionBindingSource, "ResponseStrategy")
            With cmbOwner
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = CurrentLogframe.PartnerNamesList
                .DataBindings.Add("SelectedItem", AssumptionBindingSource, "Owner")
            End With
        End If
    End Sub

    Private Sub cmbOwner_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOwner.SelectedValueChanged
        cmbOwner.Select(0, 0)
    End Sub
End Class
