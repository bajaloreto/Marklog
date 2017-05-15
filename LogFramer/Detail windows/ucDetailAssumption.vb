Public Class DetailAssumption
    Public AssumptionBindingSource As New BindingSource
    Private objCurrentAssumption As Assumption
    Private intCurrentRaidType As Integer
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

#Region "Properties"
    Public Property CurrentAssumption() As Assumption
        Get
            Return objCurrentAssumption
        End Get
        Set(ByVal value As Assumption)
            objCurrentAssumption = value
        End Set
    End Property

    Private Property CurrentRaidType As Integer
        Get
            Return intCurrentRaidType
        End Get
        Set(ByVal value As Integer)
            intCurrentRaidType = value
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
#End Region

#Region "Initialise"
    Public Sub LoadItems()
        If CurrentAssumption IsNot Nothing Then
            AssumptionBindingSource.DataSource = CurrentAssumption

            tbAssumption.DataBindings.Add("Text", AssumptionBindingSource, "Text")
            With cmbRaidType
                .DataSource = LIST_RaidTypes
                .ValueMember = "Id"
                .DisplayMember = "Value"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataBindings.Add("SelectedValue", AssumptionBindingSource, "RaidType")
            End With

        End If
        cmbRaidType.Focus()

        'set sub forms
        ChangeUserControls(CurrentAssumption.RaidType)
    End Sub
#End Region

#Region "Events"
    Private Sub cmbRaidType_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbRaidType.SelectedValueChanged
        If cmbRaidType.DataBindings.Count < 1 Then Exit Sub

        If cmbRaidType.SelectedItem Is Nothing Then Exit Sub

        Dim intRaidType As Integer = CType(cmbRaidType.SelectedItem, IdValuePair).Id

        With CurrentAssumption
            If intRaidType <> CurrentRaidType Then
                cmbRaidType.DataBindings(0).WriteValue()

                'set value detail
                If .AssumptionDetail Is Nothing Then .AssumptionDetail = New AssumptionDetail
                If .RiskDetail Is Nothing Then .RiskDetail = New RiskDetail
                If .DependencyDetail Is Nothing Then .DependencyDetail = New DependencyDetail

                ChangeUserControls(intRaidType)
                CurrentRaidType = intRaidType
            End If
        End With
    End Sub
#End Region

#Region "Add custom controls"
    Public Sub ChangeUserControls(ByVal selQuestionType As Integer)
        'Change detail windows of tab pages
        PanelRaid.Controls.Clear()

        Select Case selQuestionType
            Case Assumption.RaidTypes.Risk
                Dim ctlRisk As New AssumptionRisks(CurrentAssumption)

                ctlRisk.Dock = DockStyle.Fill
                PanelRaid.Controls.Add(ctlRisk)

            Case Assumption.RaidTypes.Assumption
                Dim ctlAssumption As New AssumptionAssumptions(CurrentAssumption)

                ctlAssumption.Dock = DockStyle.Fill
                PanelRaid.Controls.Add(ctlAssumption)

            Case Assumption.RaidTypes.Dependency
                Dim ctlDependency As New AssumptionDependencies(CurrentAssumption)

                ctlDependency.Dock = DockStyle.Fill
                PanelRaid.Controls.Add(ctlDependency)
        End Select

    End Sub
#End Region
End Class
