Imports System.Windows.Forms

Public Class DialogTargetFormula
    Private datDeadline As Date
    Private objTarget As Target
    Private dblScoringValue As Double
    Private ScoreSystem As Integer

#Region "Properties"
    Public Property Deadline As Date
        Get
            Return datDeadline
        End Get
        Set(value As Date)
            datDeadline = value
        End Set
    End Property

    Public Property Target As Target
        Get
            Return objTarget
        End Get
        Set(value As Target)
            objTarget = value
        End Set
    End Property
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal deadline As Date, ByVal target As Target, ByVal scoresystem As Integer)
        InitializeComponent()

        Me.Deadline = deadline
        Me.Target = target
        Me.ScoreSystem = scoresystem

        Text = LANG_Target
        lblFormulaDescription.Text = LANG_FormulaDescription
        dtbDeadline.DateValue = Me.Deadline

        With cmbOpMin
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.Add(CONST_LargerThan)
            .Items.Add(CONST_LargerThanOrEqual)
            .Items.Add(CONST_Equals)
            .Items.Add(CONST_SmallerThan)
            .Items.Add(CONST_SmallerThanOrEqual)
            If String.IsNullOrEmpty(target.OpMin) = False Then .SelectedItem = target.OpMin
        End With

        tbFormula.Text = target.Formula

        If scoresystem = Indicator.ScoringSystems.Score Then
            lblScoringValue.Visible = True
            lblScoringValueDescription.Visible = True
            ntbScoringValue.Visible = True
            ntbScoringValue.Text = target.Score
        Else
            lblScoringValue.Visible = False
            lblScoringValueDescription.Visible = False
            ntbScoringValue.Visible = False
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cmbOpMin_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOpMin.Validated
        If String.IsNullOrEmpty(cmbOpMin.Text) = False Then Target.OpMin = cmbOpMin.Text
    End Sub

    Private Sub tbFormula_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbFormula.Validated
        Target.Formula = tbFormula.Text
    End Sub

    Private Sub ntbScoringValue_Validated(sender As Object, e As System.EventArgs) Handles ntbScoringValue.Validated
        Target.Score = ntbScoringValue.DoubleValue
    End Sub
End Class
