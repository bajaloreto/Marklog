Public Class IndicatorMixedSubIndicators
    Friend WithEvents ResponseClassesBindingSource As New BindingSource

    Public Event ScoreSystemUpdated()

    Private indCurrentIndicator As Indicator

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()

        Me.CurrentIndicator = indicator

        If Me.CurrentIndicator.Statements.Count > 1 Then
            CurrentIndicator.Statements.Clear()
        End If
        If CurrentIndicator.Statements.Count = 0 Then
            CurrentIndicator.Statements.Add(New Statement)
        End If

        LoadItems()
    End Sub

    Public Property CurrentIndicator() As Indicator
        Get
            Return indCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            indCurrentIndicator = value
        End Set
    End Property

    Public Sub LoadItems()

        If CurrentIndicator IsNot Nothing Then
            ResponseClassesBindingSource.DataSource = CurrentIndicator.ResponseClasses

            Select Case CurrentIndicator.ScoringSystem
                Case Indicator.ScoringSystems.Value
                    CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Score
                    rbtnScoringScore.Checked = True
                Case Indicator.ScoringSystems.Percentage
                    rbtnScoringPercentage.Checked = True
                Case Indicator.ScoringSystems.Score
                    rbtnScoringScore.Checked = True
            End Select

            If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
                gbScoring.Enabled = False
            End If
        End If
    End Sub

    Private Sub dgvResponseClasses_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        If e.ColumnIndex = 1 Then
            CurrentIndicator.CalculateScores()
        End If
    End Sub

    Private Sub rbtnScoringPercentage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnScoringPercentage.CheckedChanged
        If rbtnScoringPercentage.Checked = True Then
            CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Percentage

            RaiseEvent ScoreSystemUpdated()
        End If
    End Sub

    Private Sub rbtnScoringScore_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnScoringScore.CheckedChanged
        If rbtnScoringScore.Checked = True Then
            CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Score

            RaiseEvent ScoreSystemUpdated()
        End If
    End Sub
End Class
