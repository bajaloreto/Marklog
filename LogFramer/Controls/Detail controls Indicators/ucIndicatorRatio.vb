Public Class IndicatorRatio
    Friend WithEvents FirstStatementBindingSource As New BindingSource
    Friend WithEvents SecondStatementBindingSource As New BindingSource

    Public Event TargetSystemUpdated()
    Public Event ScoreSystemUpdated()

    Private MeasureUnits1 As List(Of StructuredComboBoxItem)
    Private MeasureUnits2 As List(Of StructuredComboBoxItem)
    Private indCurrentIndicator As Indicator
    Private strFind As String
    Private FromChildIndicator As Boolean

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()

        Me.CurrentIndicator = indicator
        Me.CurrentIndicator.ValuesDetail.NrDecimals = 2
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

    Public ReadOnly Property FirstStatement As Statement
        Get
            If CurrentIndicator.Statements.Count >= 1 Then
                Return CurrentIndicator.Statements(0)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SecondStatement As Statement
        Get
            If CurrentIndicator.Statements.Count >= 2 Then
                Return CurrentIndicator.Statements(1)
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub LoadItems()
        gbValueRangeFirstStatement.Visible = My.Settings.setShowAdvancedIndicatorOptions
        gbValueRangeSecondStatement.Visible = My.Settings.setShowAdvancedIndicatorOptions
        gbScoring.Visible = My.Settings.setShowAdvancedIndicatorOptions
        gbTargetSystem.Visible = My.Settings.setShowAdvancedIndicatorOptions

        If CurrentIndicator IsNot Nothing AndAlso CurrentIndicator.Statements.Count >= 2 Then
            If FirstStatement.ValuesDetail Is Nothing Then FirstStatement.ValuesDetail = New ValuesDetail
            If SecondStatement.ValuesDetail Is Nothing Then SecondStatement.ValuesDetail = New ValuesDetail
            FirstStatementBindingSource.DataSource = FirstStatement
            SecondStatementBindingSource.DataSource = SecondStatement

            MeasureUnits1 = LoadMeasureUnits()
            MeasureUnits2 = LoadMeasureUnits()

            rtbFirstStatement.DataBindings.Add("Rtf", FirstStatementBindingSource, "FirstLabel")
            With ntbNrDecimals1
                .DataBindings.Add("Text", FirstStatementBindingSource, "ValuesDetail.NrDecimals", True)
                .DataBindings(0).FormatString = "N0"
                If FromChildIndicator = True Then .Enabled = False
            End With

            With cmbUnit1
                .DataSource = MeasureUnits1
                .DisplayMember = "Description"
                .ValueMember = "Unit"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataBindings.Add("SelectedValue", FirstStatementBindingSource, "ValuesDetail.Unit")
                If FromChildIndicator = True Then .Enabled = False
            End With
            With cmbOpMin1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange({String.Empty, CONST_LargerThan, CONST_LargerThanOrEqual})
                .DataBindings.Add("SelectedItem", FirstStatementBindingSource, "ValuesDetail.ValueRange.OpMin")
            End With
            ntbMinValue1.DataBindings.Add("Text", FirstStatementBindingSource, "ValuesDetail.ValueRange.MinValue", True)
            ntbMinValue1.DataBindings(0).FormatString = FirstStatement.ValuesDetail.FormatString
            With cmbOpMax1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange({String.Empty, CONST_SmallerThan, CONST_SmallerThanOrEqual})
                .DataBindings.Add("SelectedItem", FirstStatementBindingSource, "ValuesDetail.ValueRange.OpMax")
            End With
            ntbMaxValue1.DataBindings.Add("Text", FirstStatementBindingSource, "ValuesDetail.ValueRange.MaxValue", True)
            ntbMaxValue1.DataBindings(0).FormatString = FirstStatement.ValuesDetail.FormatString

            rtbSecondStatement.DataBindings.Add("Rtf", SecondStatementBindingSource, "FirstLabel")
            With ntbNrDecimals2
                .DataBindings.Add("Text", SecondStatementBindingSource, "ValuesDetail.NrDecimals", True)
                .DataBindings(0).FormatString = "N0"
                If FromChildIndicator = True Then .Enabled = False
            End With
            With cmbUnit2
                .DataSource = MeasureUnits2
                .DisplayMember = "Description"
                .ValueMember = "Unit"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataBindings.Add("SelectedValue", SecondStatementBindingSource, "ValuesDetail.Unit")
                If FromChildIndicator = True Then .Enabled = False
            End With
            With cmbOpMin2
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange({String.Empty, CONST_LargerThan, CONST_LargerThanOrEqual})
                .DataBindings.Add("SelectedItem", SecondStatementBindingSource, "ValuesDetail.ValueRange.OpMin")
            End With
            ntbMinValue2.DataBindings.Add("Text", SecondStatementBindingSource, "ValuesDetail.ValueRange.MinValue", True)
            ntbMinValue2.DataBindings(0).FormatString = SecondStatement.ValuesDetail.FormatString
            With cmbOpMax2
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange({String.Empty, CONST_SmallerThan, CONST_SmallerThanOrEqual})
                .DataBindings.Add("SelectedItem", SecondStatementBindingSource, "ValuesDetail.ValueRange.OpMax")
            End With
            ntbMaxValue2.DataBindings.Add("Text", SecondStatementBindingSource, "ValuesDetail.ValueRange.MaxValue", True)
            ntbMaxValue2.DataBindings(0).FormatString = SecondStatement.ValuesDetail.FormatString

            SetScoringSystem()
            SetTargetSystem()

            If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
                gbScoring.Enabled = False
                Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(CurrentIndicator.ParentIndicatorGuid)
                If ParentIndicator IsNot Nothing AndAlso ParentIndicator.QuestionType <> Indicator.QuestionTypes.MixedSubIndicators Then _
                    gbTargetSystem.Enabled = False
            End If

            Select Case CurrentIndicator.QuestionType
                Case Indicator.QuestionTypes.PercentageValue
                    If CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Percentage Then CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Value
                    cmbUnit1.Enabled = False
                    rbtnScoringPercentage.Enabled = False
            End Select
        End If
    End Sub
#End Region

#Region "Methods"
    Private Sub SetValueUnitsFirstStatement()
        With FirstStatement.ValuesDetail
            ntbMinValue1.DataBindings(0).FormatString = .FormatString
            ntbMaxValue1.DataBindings(0).FormatString = .FormatString
        End With
    End Sub

    Private Sub SetValueUnitsSecondStatement()
        With SecondStatement.ValuesDetail
            ntbMinValue2.DataBindings(0).FormatString = .FormatString
            ntbMaxValue2.DataBindings(0).FormatString = .FormatString
        End With
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

    Private Sub UpdateIndicatorUnit()
        CurrentIndicator.SetUnitWithFormula()

        CurrentLogFrame.GetCustomMeasureUnits()
        MeasureUnits1 = LoadMeasureUnits()
        MeasureUnits2 = LoadMeasureUnits()
        cmbUnit1.DataSource = MeasureUnits1
        cmbUnit1.DataBindings(0).ReadValue()
        cmbUnit2.DataSource = MeasureUnits2
        cmbUnit2.DataBindings(0).ReadValue()
    End Sub

    Private Sub SetNrDecimals()
        Dim intNrDecimals(1) As Integer

        With CurrentIndicator
            intNrDecimals(0) = .Statements(0).ValuesDetail.NrDecimals
            intNrDecimals(1) = .Statements(1).ValuesDetail.NrDecimals
            If intNrDecimals(0) > intNrDecimals(1) Then
                .ValuesDetail.NrDecimals = intNrDecimals(0)
            Else
                .ValuesDetail.NrDecimals = intNrDecimals(1)
            End If
        End With
    End Sub

    Private Function FindStatementtypeName(ByVal selItem As StructuredComboBoxItem) As Boolean
        If selItem.Description = strFind Then Return True Else Return False
    End Function
#End Region

#Region "Events"
    Private Sub rtbFirstStatement_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbFirstStatement.Enter
        CurrentControl = rtbFirstStatement
    End Sub

    Private Sub rtbFirstStatement_TextChanged(sender As Object, e As System.EventArgs) Handles rtbFirstStatement.TextChanged
        If CurrentControl Is rtbFirstStatement Then
            UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.TextChanged)
        End If
    End Sub

    Private Sub rtbFirstStatement_Validated(sender As Object, e As System.EventArgs) Handles rtbFirstStatement.Validated
        UndoRedo.CurrentControlValidated(rtbFirstStatement)
    End Sub

    Private Sub cmbUnit1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmbUnit1.Validating
        Dim strText As String = cmbUnit1.Text
        Dim objComboBoxItem As Object = Nothing

        strFind = strText
        objComboBoxItem = MeasureUnits1.Find(AddressOf FindStatementtypeName)

        If objComboBoxItem Is Nothing Then
            Dim NewItem As New StructuredComboBoxItem(strText, False, False, , strText)
            MeasureUnitsUser.Add(NewItem)
            MeasureUnits1 = LoadMeasureUnits()

            cmbUnit1.DataSource = MeasureUnits1
            cmbUnit1.Text = strText
        End If
    End Sub

    Private Sub cmbUnit1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnit1.Validated
        SetValueUnitsFirstStatement()

        UpdateIndicatorUnit()
    End Sub

    Private Sub ntbNrDecimals1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbNrDecimals1.Validated
        SetValueUnitsFirstStatement()

        SetNrDecimals()
    End Sub

    Private Sub rtbSecondStatement_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbSecondStatement.Enter
        CurrentControl = rtbSecondStatement
    End Sub

    Private Sub rtbSecondStatement_TextChanged(sender As Object, e As System.EventArgs) Handles rtbSecondStatement.TextChanged
        If CurrentControl Is rtbSecondStatement Then
            UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.TextChanged)
        End If
    End Sub

    Private Sub rtbSecondStatement_Validated(sender As Object, e As System.EventArgs) Handles rtbSecondStatement.Validated
        UndoRedo.CurrentControlValidated(rtbSecondStatement)
    End Sub

    Private Sub cmbUnit2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnit2.Validated
        SetValueUnitsSecondStatement()

        UpdateIndicatorUnit()
    End Sub

    Private Sub cmbUnit2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmbUnit2.Validating
        Dim strText As String = cmbUnit2.Text
        Dim objComboBoxItem As Object = Nothing

        strFind = strText
        objComboBoxItem = MeasureUnits2.Find(AddressOf FindStatementtypeName)

        If objComboBoxItem Is Nothing Then
            Dim NewItem As New StructuredComboBoxItem(strText, False, False, , strText)
            MeasureUnitsUser.Add(NewItem)
            MeasureUnits2 = LoadMeasureUnits()

            cmbUnit2.DataSource = MeasureUnits2
            cmbUnit2.Text = strText
        End If
    End Sub

    Private Sub ntbNrDecimals2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbNrDecimals2.Validated
        SetValueUnitsSecondStatement()

        SetNrDecimals()
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
