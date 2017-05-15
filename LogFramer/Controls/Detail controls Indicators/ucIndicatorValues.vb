Public Class IndicatorValues
    Friend WithEvents ValuesDetailBindingSource As New BindingSource

    Public Event TargetSystemUpdated()
    Public Event ScoreSystemUpdated()

    Private indCurrentIndicator As Indicator
    Private MeasureUnits As List(Of StructuredComboBoxItem)
    Private strFind As String
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

#Region "Properties"
    Protected Property CurrentIndicator() As Indicator
        Get
            Return indCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            indCurrentIndicator = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub LoadItems()
        gbValueRange.Visible = My.Settings.setShowAdvancedIndicatorOptions
        gbScoring.Visible = My.Settings.setShowAdvancedIndicatorOptions
        gbTargetSystem.Visible = My.Settings.setShowAdvancedIndicatorOptions

        If CurrentIndicator IsNot Nothing Then
            If CurrentIndicator.ValuesDetail Is Nothing Then CurrentIndicator.ValuesDetail = New ValuesDetail
            ValuesDetailBindingSource.DataSource = CurrentIndicator.ValuesDetail
            MeasureUnits = LoadMeasureUnits()

            With tbNrDecimals
                .DataBindings.Add("Text", ValuesDetailBindingSource, "NrDecimals") ', True)
                '.DataBindings(0).FormatString = "N0"
            End With

            With cmbUnit
                .ValueMember = "Unit"
                .DisplayMember = "Description"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = MeasureUnits
                .DataBindings.Add("SelectedValue", ValuesDetailBindingSource, "Unit")
                If FromChildIndicator = True Then .Enabled = False
            End With

            If CurrentIndicator.Indicators.Count = 0 Then
                With cmbOpMin
                    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .Items.AddRange({String.Empty, CONST_LargerThan, CONST_LargerThanOrEqual})
                    .DataBindings.Add("SelectedItem", ValuesDetailBindingSource, "ValueRange.OpMin")
                End With
                tbMinValue.DataBindings.Add("Text", ValuesDetailBindingSource, "ValueRange.MinValue", True)
                tbMinValue.DataBindings(0).FormatString = CurrentIndicator.ValuesDetail.FormatString
                With cmbOpMax
                    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .Items.AddRange({String.Empty, CONST_SmallerThan, CONST_SmallerThanOrEqual})
                    .DataBindings.Add("SelectedItem", ValuesDetailBindingSource, "ValueRange.OpMax")
                End With
                tbMaxValue.DataBindings.Add("Text", ValuesDetailBindingSource, "ValueRange.MaxValue", True)
                tbMaxValue.DataBindings(0).FormatString = CurrentIndicator.ValuesDetail.FormatString
            Else
                gbValueRange.Visible = False
                gbNumber.Height = gbScoring.Height
            End If

            SetScoringSystem()
            SetTargetSystem()

            If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
                gbScoring.Enabled = False
                Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(CurrentIndicator.ParentIndicatorGuid)
                If ParentIndicator IsNot Nothing AndAlso ParentIndicator.QuestionType <> Indicator.QuestionTypes.MixedSubIndicators Then _
                    gbTargetSystem.Enabled = False
            End If
            'If CurrentIndicator.Indicators.Count > 0 Then gbTargetSystem.Visible = False

            Select Case CurrentIndicator.QuestionType
                Case Indicator.QuestionTypes.PercentageValue
                    If CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Percentage Then CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Value
                    cmbUnit.Enabled = False
                    rbtnScoringPercentage.Enabled = False
            End Select
        End If
    End Sub
#End Region

#Region "Methods"
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
                rbtnTargetSystemValueRange.Checked = False
                rbtnTargetSystemFormula.Checked = False
            Case Indicator.TargetSystems.ValueRange
                rbtnTargetSystemSimple.Checked = False
                rbtnTargetSystemValueRange.Checked = True
                rbtnTargetSystemFormula.Checked = False
            Case Indicator.TargetSystems.Formula
                rbtnTargetSystemSimple.Checked = False
                rbtnTargetSystemValueRange.Checked = False
                rbtnTargetSystemFormula.Checked = True
        End Select
    End Sub

    Private Sub SetValueUnits()
        With CurrentIndicator.ValuesDetail
            If tbMinValue.DataBindings.Count > 0 Then tbMinValue.DataBindings(0).FormatString = .FormatString
            If tbMaxValue.DataBindings.Count > 0 Then tbMaxValue.DataBindings(0).FormatString = .FormatString
        End With
    End Sub

    Private Function FindResponsetypeName(ByVal selItem As StructuredComboBoxItem) As Boolean
        If selItem.Description = strFind Then Return True Else Return False
    End Function
#End Region

#Region "Events"
    Private Sub tbNrDecimals_Validated(sender As Object, e As System.EventArgs) Handles tbNrDecimals.Validated
        SetValueUnits()
    End Sub

    Private Sub cmbUnit_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnit.Validated
        SetValueUnits()
    End Sub

    Private Sub cmbUnit_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmbUnit.Validating
        Dim strText As String = cmbUnit.Text
        Dim objComboBoxItem As Object = Nothing

        strFind = strText
        objComboBoxItem = MeasureUnits.Find(AddressOf FindResponsetypeName)

        If objComboBoxItem Is Nothing Then
            Dim NewItem As New StructuredComboBoxItem(strText, False, False, , strText)
            MeasureUnitsUser.Add(NewItem)
            MeasureUnits = LoadMeasureUnits()

            cmbUnit.DataSource = MeasureUnits
            cmbUnit.Text = strText
        End If
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

    Private Sub rbtnTargetSystemSimple_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnTargetSystemSimple.CheckedChanged
        If rbtnTargetSystemSimple.Checked = True Then
            Dim intOldValue As Integer = CurrentIndicator.TargetSystem

            If intOldValue <> Indicator.TargetSystems.Simple Then
                CurrentIndicator.TargetSystem = Indicator.TargetSystems.Simple

                For Each selTarget As Target In CurrentIndicator.Targets
                    selTarget.OpMin = CONST_LargerThanOrEqual
                Next

                UndoRedo.OptionChanged(CurrentIndicator, "TargetSystem", intOldValue, CurrentIndicator.TargetSystem)
                RaiseEvent TargetSystemUpdated()
            End If
        End If
    End Sub

    Private Sub rbtnTargetSystemValueRange_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnTargetSystemValueRange.CheckedChanged
        If rbtnTargetSystemValueRange.Checked = True Then
            Dim intOldValue As Integer = CurrentIndicator.TargetSystem

            If intOldValue <> Indicator.TargetSystems.ValueRange Then
                CurrentIndicator.TargetSystem = Indicator.TargetSystems.ValueRange

                UndoRedo.OptionChanged(CurrentIndicator, "TargetSystem", intOldValue, CurrentIndicator.TargetSystem)
                RaiseEvent TargetSystemUpdated()
            End If
        End If
    End Sub

    Private Sub rbtnTargetSystemFormula_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnTargetSystemFormula.CheckedChanged
        If rbtnTargetSystemFormula.Checked = True Then
            Dim intOldValue As Integer = CurrentIndicator.TargetSystem

            If intOldValue <> Indicator.TargetSystems.Formula Then
                CurrentIndicator.TargetSystem = Indicator.TargetSystems.Formula

                UndoRedo.OptionChanged(CurrentIndicator, "TargetSystem", intOldValue, CurrentIndicator.TargetSystem)
                RaiseEvent TargetSystemUpdated()
            End If
        End If
    End Sub
#End Region
End Class
