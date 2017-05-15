Public Class IndicatorScales
    Friend WithEvents dgvStatementsScales As New DataGridViewStatementsScales
    Friend WithEvents IndicatorBindingSource As New BindingSource

    Public Event ScoreSystemUpdated()
    Public Event CurrentDataGridViewChanged(ByVal sender As Object, ByVal e As CurrentDataGridViewChangedEventArgs)

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
            IndicatorBindingSource.DataSource = CurrentIndicator

            dgvStatementsScales = New DataGridViewStatementsScales(CurrentIndicator)
            With dgvStatementsScales
                .Name = "dgvStatementsScales"
                .Dock = DockStyle.Fill
            End With
            gbScoreValues.Controls.Add(dgvStatementsScales)

            SetScoringSystem()

            With ntbNrDecimals
                .DataBindings.Add("Text", IndicatorBindingSource, "ValuesDetail.NrDecimals", True)
                .DataBindings(0).FormatString = "N0"
            End With
            tbAgreeText.DataBindings.Add("Text", IndicatorBindingSource, "ScalesDetail.AgreeText")
            tbDisagreeText.DataBindings.Add("Text", IndicatorBindingSource, "ScalesDetail.DisagreeText")

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
    Private Sub dgvStatementsScales_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvStatementsScales.Enter
        RaiseEvent CurrentDataGridViewChanged(Me, New CurrentDataGridViewChangedEventArgs(dgvStatementsScales))
    End Sub

    Private Sub dgvStatementsScales_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvStatementsScales.CellValidated
        If e.ColumnIndex = 1 Then
            CurrentIndicator.CalculateScores()
        End If
    End Sub

    Private Sub dgvStatementsScales_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvStatementsScales.EditingControlShowing
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
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

Public Class CurrentDataGridViewChangedEventArgs
    Inherits EventArgs

    Public Property DataGridView As DataGridViewBaseClassRichText

    Public Sub New(ByVal objDataGridView As Object)
        MyBase.New()

        Me.DataGridView = objDataGridView
    End Sub
End Class
