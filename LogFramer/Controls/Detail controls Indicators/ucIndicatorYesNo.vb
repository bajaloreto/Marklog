Public Class IndicatorYesNo
    Friend WithEvents ResponseClassesBindingSource As New BindingSource
    Friend WithEvents IndicatorBindingSource As New BindingSource

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

    Protected Property CurrentIndicator() As Indicator
        Get
            Return indCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            indCurrentIndicator = value
        End Set
    End Property

    Public Property YesValue As Double
        Get
            If CurrentIndicator.ResponseClasses.Count >= 1 Then
                Return CurrentIndicator.ResponseClasses(0).Value
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Double)
            If CurrentIndicator.ResponseClasses.Count >= 1 Then
                CurrentIndicator.ResponseClasses(0).Value = value
            End If
        End Set
    End Property

    Public Property NoValue As Double
        Get
            If CurrentIndicator.ResponseClasses.Count >= 2 Then
                Return CurrentIndicator.ResponseClasses(1).Value
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Double)
            If CurrentIndicator.ResponseClasses.Count >= 2 Then
                CurrentIndicator.ResponseClasses(1).Value = value
            End If
        End Set
    End Property

    Public Sub LoadItems()

        If CurrentIndicator IsNot Nothing Then
            IndicatorBindingSource.DataSource = CurrentIndicator

            If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
                gbScoring.Enabled = False
            End If

            With ntbYes
                .IsCurrency = False
                .IsPercentage = False
                .SetDecimals = True
                .NrDecimals = 0
            End With

            With ntbNo
                .IsCurrency = False
                .IsPercentage = False
                .SetDecimals = True
                .NrDecimals = 0
            End With

            ReadScoreValues()
            SetScoringSystem()

            If CurrentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                ntbNrDecimals.DataBindings.Add("Text", IndicatorBindingSource, "ValuesDetail.NrDecimals", True)
                ntbNrDecimals.DataBindings(0).FormatString = "N0"
            Else
                gbNumber.Visible = False
            End If
        End If
    End Sub

    Public Sub ReadScoreValues()
        ntbYes.IntegerValue = YesValue
        ntbNo.IntegerValue = NoValue
    End Sub

    Public Sub SetScoringSystem()
        Select Case CurrentIndicator.ScoringSystem
            Case Indicator.ScoringSystems.Value
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

    Private Sub ntbYes_Enter(sender As Object, e As System.EventArgs) Handles ntbYes.Enter
        CurrentControl = ntbYes
        UndoRedo.UndoBuffer_Initialise(CurrentIndicator.ResponseClasses(0), "Value")
    End Sub

    Private Sub ntbYes_TextChanged(sender As Object, e As System.EventArgs) Handles ntbYes.TextChanged
        UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.ValueChanged)
    End Sub

    Private Sub ntbYes_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbYes.Validated
        Dim intOldValue As Integer = YesValue

        If ntbYes.IntegerValue <> intOldValue Then
            YesValue = ntbYes.IntegerValue
            CurrentIndicator.CalculateScores()

            UndoRedo.ValueChanged(intOldValue, YesValue)
        End If
    End Sub

    Private Sub ntbNo_Enter(sender As Object, e As System.EventArgs) Handles ntbNo.Enter
        CurrentControl = ntbNo
        UndoRedo.UndoBuffer_Initialise(CurrentIndicator.ResponseClasses(1), "Value")
    End Sub

    Private Sub ntbNo_TextChanged(sender As Object, e As System.EventArgs) Handles ntbNo.TextChanged
        UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.ValueChanged)
    End Sub

    Private Sub ntbNo_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbNo.Validated
        Dim intOldValue As Integer = NoValue

        If ntbNo.IntegerValue <> intOldValue Then
            NoValue = ntbNo.IntegerValue
            CurrentIndicator.CalculateScores()

            UndoRedo.ValueChanged(intOldValue, NoValue)
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
End Class
