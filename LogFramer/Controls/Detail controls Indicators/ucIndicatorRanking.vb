Public Class IndicatorRanking
    Friend WithEvents ResponseClassesBindingSource As New BindingSource
    Friend WithEvents IndicatorBindingSource As New BindingSource
    Friend WithEvents dgvResponseClasses As New DataGridViewResponseClasses

    Public Event ScoreSystemUpdated()

    Private indCurrentIndicator As Indicator


    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()

        Me.CurrentIndicator = indicator

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

            dgvResponseClasses = New DataGridViewResponseClasses(CurrentIndicator.ResponseClasses, False)

            With dgvResponseClasses
                .Name = "dgvResponseClasses"
                .DataSource = ResponseClassesBindingSource
                .Dock = DockStyle.Fill
            End With

            gbScoreValues.Controls.Add(dgvResponseClasses)

            SetScoringSystem()

            If CurrentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                ntbNrDecimals.DataBindings.Add("Text", IndicatorBindingSource, "ValuesDetail.NrDecimals", True)
                ntbNrDecimals.DataBindings(0).FormatString = "N0"
            Else
                gbNumber.Visible = False
            End If

            If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
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
#End Region
End Class
