Public Class IndicatorFrequencyLikert
    Friend WithEvents ResponseClassesBindingSource As New BindingSource
    Friend WithEvents IndicatorBindingSource As New BindingSource
    Friend WithEvents dgvResponseClasses As New DataGridViewResponseClasses

    Public Event ScoreSystemUpdated()

    Private indCurrentIndicator As Indicator
    Private MeasureUnits As List(Of StructuredComboBoxItem)
    Private FromChildIndicator As Boolean
    Private strFind As String
    Private colClassName As New DataGridViewTextBoxColumn
    Private colValue As New DataGridViewTextBoxColumn

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
    Public Property CurrentIndicator() As Indicator
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

        If CurrentIndicator IsNot Nothing Then
            ResponseClassesBindingSource.DataSource = CurrentIndicator.ResponseClasses
            IndicatorBindingSource.DataSource = CurrentIndicator
            MeasureUnits = LoadMeasureUnits()

            dgvResponseClasses = New DataGridViewResponseClasses(CurrentIndicator.ResponseClasses, True)

            With dgvResponseClasses
                .Name = "dgvResponseClasses"
                .DataSource = ResponseClassesBindingSource
                .Dock = DockStyle.Fill
            End With

            gbScoreValues.Controls.Add(dgvResponseClasses)

            SetScoringSystem()

            ntbNrDecimals.DataBindings.Add("Text", IndicatorBindingSource, "ValuesDetail.NrDecimals", True)
            ntbNrDecimals.DataBindings(0).FormatString = "N0"

            With cmbUnit
                .ValueMember = "Unit"
                .DisplayMember = "Description"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = MeasureUnits
                .DataBindings.Add("SelectedValue", IndicatorBindingSource, "ValuesDetail.Unit")
            End With

            If FromChildIndicator = True Then
                gbScoring.Enabled = False
            End If
        End If
    End Sub
#End Region

#Region "Methods"
    Public Sub SetScoringSystem()
        Select Case CurrentIndicator.ScoringSystem
            Case Indicator.ScoringSystems.Value
                CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Score
                rbtnScoringPercentage.Checked = False
                rbtnScoringScore.Checked = True
            Case Indicator.ScoringSystems.Percentage
                rbtnScoringPercentage.Checked = True
                rbtnScoringScore.Checked = False
            Case Indicator.ScoringSystems.Score
                rbtnScoringPercentage.Checked = False
                rbtnScoringScore.Checked = True
        End Select
    End Sub
#End Region

#Region "Events"
    Private Sub dgvResponseClasses_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvResponseClasses.CellValidated
        If e.ColumnIndex = 1 Then
            CurrentIndicator.CalculateScores()
        End If
    End Sub

    Private Sub dgvResponseClasses_DataSourceUpdated() Handles dgvResponseClasses.DataSourceUpdated
        With dgvResponseClasses
            .DataSource = Nothing
            .DataSource = ResponseClassesBindingSource
        End With
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

    Private Function FindResponsetypeName(ByVal selItem As StructuredComboBoxItem) As Boolean
        If selItem.Description = strFind Then Return True Else Return False
    End Function
#End Region
End Class
