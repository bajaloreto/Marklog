Public Class AssumptionAssumptions
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

            tbReason.DataBindings.Add("Text", AssumptionBindingSource, "AssumptionDetail.Reason")
            tbHowToValidate.DataBindings.Add("Text", AssumptionBindingSource, "AssumptionDetail.HowToValidate")
            chkValidated.DataBindings.Add("Checked", AssumptionBindingSource, "AssumptionDetail.Validated")
            tbImpact.DataBindings.Add("Text", AssumptionBindingSource, "Impact")
            
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
