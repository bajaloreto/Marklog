Public Class IndicatorFormula
    Friend WithEvents dgvStatementsFormula As New DataGridViewStatementsFormula
    Friend WithEvents ValuesDetailBindingSource As New BindingSource

    Public Event TargetSystemUpdated()
    Public Event ScoreSystemUpdated()
    Public Event StatementsUpdated()
    Public Event CurrentDataGridViewChanged(ByVal sender As Object, ByVal e As CurrentDataGridViewChangedEventArgs)

    Private MeasureUnits As List(Of StructuredComboBoxItem)
    Private indCurrentIndicator As Indicator
    Private FromChildIndicator As Boolean

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()

        Me.CurrentIndicator = indicator
        If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then FromChildIndicator = True

        LoadItems()
    End Sub

    Protected Property CurrentIndicator() As Indicator
        Get
            Return indCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            indCurrentIndicator = value
        End Set
    End Property

    Public Sub LoadItems()
        gbScoring.Visible = My.Settings.setShowAdvancedIndicatorOptions
        gbTargetSystem.Visible = My.Settings.setShowAdvancedIndicatorOptions

        If CurrentIndicator IsNot Nothing Then
            ValuesDetailBindingSource.DataSource = CurrentIndicator.ValuesDetail
            MeasureUnits = LoadMeasureUnits()

            dgvStatementsFormula = New DataGridViewStatementsFormula(CurrentIndicator)
            With dgvStatementsFormula
                .Name = "dgvStatementsFormula"
                .Dock = DockStyle.Fill
            End With
            gbQuestions.Controls.Add(dgvStatementsFormula)

            tbFormula.DataBindings.Add("Text", ValuesDetailBindingSource, "Formula")

            ntbNrDecimals.DataBindings.Add("Text", ValuesDetailBindingSource, "NrDecimals", True)
            ntbNrDecimals.DataBindings(0).FormatString = "N0"

            With cmbUnit
                .DataSource = MeasureUnits
                .DisplayMember = "Description"
                .ValueMember = "Unit"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataBindings.Add("SelectedValue", ValuesDetailBindingSource, "Unit")
                If FromChildIndicator = True Then .Enabled = False
            End With

            SetScoringSystem()
            SetTargetSystem()

            If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
                gbScoring.Enabled = False
                Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(CurrentIndicator.ParentIndicatorGuid)
                If ParentIndicator IsNot Nothing AndAlso ParentIndicator.QuestionType <> Indicator.QuestionTypes.MixedSubIndicators Then _
                    gbTargetSystem.Enabled = False
            End If
        End If
    End Sub

    Public Sub SetScoringSystem()
        Select Case CurrentIndicator.ScoringSystem
            Case Indicator.ScoringSystems.Value
                rbtnScoringValue.Checked = True
                rbtnScoringPercentage.Checked = False
                rbtnScoringScore.Checked = False
            Case Indicator.ScoringSystems.Percentage
                rbtnScoringValue.Checked = False
                rbtnScoringPercentage.Checked = True
                rbtnScoringScore.Checked = False
            Case Indicator.ScoringSystems.Score
                rbtnScoringValue.Checked = False
                rbtnScoringPercentage.Checked = False
                rbtnScoringScore.Checked = True
        End Select
    End Sub

    Public Sub SetTargetSystem()
        Select Case CurrentIndicator.TargetSystem
            Case Indicator.TargetSystems.Simple
                rbtnTargetSystemSimple.Checked = True
            Case Indicator.TargetSystems.ValueRange
                rbtnTargetSystemSimple.Checked = False
            Case Indicator.TargetSystems.Formula
                rbtnTargetSystemSimple.Checked = False
        End Select
    End Sub

    Private Sub rbtnScoringValue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnScoringValue.CheckedChanged
        If rbtnScoringValue.Checked = True Then
            Dim intOldValue As Integer = CurrentIndicator.ScoringSystem

            If intOldValue <> Indicator.ScoringSystems.Value Then
                CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Value

                UndoRedo.OptionChanged(CurrentIndicator, "ScoringSystem", intOldValue, CurrentIndicator.ScoringSystem)
                RaiseEvent ScoreSystemUpdated()
            End If
        End If
    End Sub

    Private Sub rbtnScoringPercentage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnScoringPercentage.CheckedChanged
        If rbtnScoringPercentage.Checked = True Then
            Dim intOldValue As Integer = CurrentIndicator.ScoringSystem

            If intOldValue <> Indicator.ScoringSystems.Percentage Then
                CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Percentage

                UndoRedo.OptionChanged(CurrentIndicator, "ScoringSystem", intOldValue, CurrentIndicator.ScoringSystem)
                RaiseEvent ScoreSystemUpdated()
            End If
        End If
    End Sub

    Private Sub rbtnScoringScore_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnScoringScore.CheckedChanged
        If rbtnScoringScore.Checked = True Then
            Dim intOldValue As Integer = CurrentIndicator.ScoringSystem

            If intOldValue <> Indicator.ScoringSystems.Score Then
                CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Score

                UndoRedo.OptionChanged(CurrentIndicator, "ScoringSystem", intOldValue, CurrentIndicator.ScoringSystem)
                RaiseEvent ScoreSystemUpdated()
            End If
        End If
    End Sub

    Private Sub dgvStatementsFormula_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvStatementsFormula.EditingControlShowing
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
        End With
    End Sub

    Private Sub dgvStatementsFormula_Enter(sender As Object, e As System.EventArgs) Handles dgvStatementsFormula.Enter
        RaiseEvent CurrentDataGridViewChanged(Me, New CurrentDataGridViewChangedEventArgs(dgvStatementsFormula))
    End Sub

    Private Sub dgvStatementsFormula_FormulaUpdated() Handles dgvStatementsFormula.FormulaUpdated
        tbFormula.DataBindings(0).ReadValue()

        UpdateIndicatorUnit()
    End Sub

    Private Sub dgvStatementsFormula_UnitUpdated() Handles dgvStatementsFormula.UnitUpdated
        UpdateIndicatorUnit()
    End Sub

    Private Sub UpdateIndicatorUnit()
        CurrentIndicator.SetUnitWithFormula()

        CurrentLogFrame.GetCustomMeasureUnits()
        MeasureUnits = LoadMeasureUnits()
        cmbUnit.DataSource = MeasureUnits
        cmbUnit.DataBindings(0).ReadValue()

        RaiseEvent StatementsUpdated()
    End Sub

    Private Sub tbFormula_Validated(sender As Object, e As System.EventArgs) Handles tbFormula.Validated
        UpdateIndicatorUnit()
    End Sub

    Private Sub dgvStatementsFormula_UserAddedRow(sender As Object, e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvStatementsFormula.UserAddedRow
        RaiseEvent StatementsUpdated()
    End Sub

    Private Sub dgvStatementsFormula_UserDeletedRow(sender As Object, e As System.Windows.Forms.DataGridViewRowEventArgs) Handles dgvStatementsFormula.UserDeletedRow
        UpdateIndicatorUnit()
    End Sub
End Class
