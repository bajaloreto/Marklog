Public Class IndicatorMaxDiff
    Friend WithEvents dgvStatementsMaxDiff As New DataGridViewStatementsMaxDiff
    Friend WithEvents IndicatorBindingSource As New BindingSource
    Friend WithEvents ResponsesBindingSource As New BindingSource
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
            IndicatorBindingSource.DataSource = CurrentIndicator
            ResponsesBindingSource.DataSource = CurrentIndicator.ResponseClasses(0)

            dgvStatementsMaxDiff = New DataGridViewStatementsMaxDiff(CurrentIndicator)
            With dgvStatementsMaxDiff
                .Name = "dgvStatementsMaxDiff"
                .Dock = DockStyle.Fill
            End With
            gbStatements.Controls.Add(dgvStatementsMaxDiff)

            ntbNotSelectedScore.Text = "0"
            tbAgreeText.DataBindings.Add("Text", IndicatorBindingSource, "ScalesDetail.AgreeText")
            tbDisagreeText.DataBindings.Add("Text", IndicatorBindingSource, "ScalesDetail.DisagreeText")

            ReadScoreValues()
        End If
    End Sub

    Public Sub ReadScoreValues()
        If CurrentIndicator.ResponseClasses.Count > 0 Then
            Dim dblBestValue As Integer = CurrentIndicator.ResponseClasses(0).Value
            Dim dblWorstValue As Integer = dblBestValue * -1

            ntbBestOptionScore.Text = dblBestValue.ToString("N0")
            ntbWorstOptionScore.Text = dblWorstValue.ToString("N0")
        End If
    End Sub

    Private Sub dgvResponsesScales_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvStatementsMaxDiff.Enter
        RaiseEvent CurrentDataGridViewChanged(Me, New CurrentDataGridViewChangedEventArgs(dgvStatementsMaxDiff))
    End Sub

    Private Sub ntbBestOptionScore_Enter(sender As Object, e As System.EventArgs) Handles ntbBestOptionScore.Enter
        CurrentControl = ntbBestOptionScore
        UndoRedo.UndoBuffer_Initialise(CurrentIndicator.ResponseClasses(0), "Value", ntbBestOptionScore.IntegerValue)
    End Sub

    Private Sub ntbBestOptionScore_TextChanged(sender As Object, e As System.EventArgs) Handles ntbBestOptionScore.TextChanged
        If CurrentControl Is ntbBestOptionScore Then UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.ValueChanged)
    End Sub

    Private Sub ntbBestOptionScore_Validated(sender As Object, e As System.EventArgs) Handles ntbBestOptionScore.Validated
        Dim intOldValue As Integer = CurrentIndicator.ResponseClasses(0).Value
        CurrentIndicator.ResponseClasses(0).Value = ntbBestOptionScore.IntegerValue
        ReadScoreValues()

        UndoRedo.ValueChanged(intOldValue, CurrentIndicator.ResponseClasses(0).Value)
    End Sub
End Class
