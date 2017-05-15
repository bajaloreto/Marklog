Imports System.Windows.Forms

Public Class DialogTarget
    Friend WithEvents TargetBindingSource As New BindingSource

    Private datDeadline As Date
    Private objTarget As Target
    Private dblScoringValue As Double
    Private ScoreSystem, TargetSystem As Integer
    Private intNrDecimals As Integer
    Private strUnit As String

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

    Public Property NrDecimals() As Integer
        Get
            Return intNrDecimals
        End Get
        Set(ByVal value As Integer)
            intNrDecimals = value
        End Set
    End Property

    Public Property Unit() As String
        Get
            Return strUnit
        End Get
        Set(ByVal value As String)
            strUnit = value
        End Set
    End Property

    Public ReadOnly Property FormatString As String
        Get
            Dim strPrecision As String = "#,##0."
            If Me.NrDecimals > 0 Then
                For i = 1 To Me.NrDecimals
                    strPrecision &= "0"
                Next
            End If
            If String.IsNullOrEmpty(Me.Unit) = False Then strPrecision = String.Format("{0} '{1}'", strPrecision, Me.Unit)

            Return strPrecision
        End Get
    End Property
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal deadline As Date, ByVal target As Target, ByVal scoresystem As Integer, ByVal targetsystem As Integer, ByVal strUnit As String, ByVal intNrDecimals As Integer)
        InitializeComponent()

        Me.Deadline = deadline
        Me.Target = target
        Me.ScoreSystem = scoresystem
        Me.TargetSystem = targetsystem
        Me.Unit = strUnit
        Me.NrDecimals = intNrDecimals

        lblMinValue.Text = LANG_MinValueForSuccess
        lblMinValueDescription.Text = LANG_MinValueDescription
        lblMaxValue.Text = LANG_MaxValueForSuccess
        dtbDeadline.DateValue = Me.Deadline

        If Me.Target IsNot Nothing Then
            TargetBindingSource.DataSource = Me.Target

            With cmbOpMin
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange({String.Empty, CONST_LargerThan, CONST_LargerThanOrEqual})
                .DataBindings.Add("SelectedItem", TargetBindingSource, "OpMin")
            End With
            ntbMinValue.DataBindings.Add("Text", TargetBindingSource, "MinValue", True)
            ntbMinValue.DataBindings(0).FormatString = Me.FormatString

            With cmbOpMax
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange({String.Empty, CONST_SmallerThan, CONST_SmallerThanOrEqual})
                .DataBindings.Add("SelectedItem", TargetBindingSource, "OpMax")
            End With
            ntbMaxValue.DataBindings.Add("Text", TargetBindingSource, "MaxValue", True)
            ntbMaxValue.DataBindings(0).FormatString = Me.FormatString

            If scoresystem = Indicator.ScoringSystems.Score Then
                lblScoringValue.Visible = True
                lblScoringValueDescription.Visible = True
                ntbScoringValue.Visible = True
                ntbScoringValue.DataBindings.Add("Text", TargetBindingSource, "Score")
            Else
                lblScoringValue.Visible = False
                lblScoringValueDescription.Visible = False
                ntbScoringValue.Visible = False
            End If

            If targetsystem = Indicator.TargetSystems.Simple Then
                cmbOpMin.Enabled = False
                lblMaxValue.Visible = False
                lblMaxValueDescription.Visible = False
                cmbOpMax.Visible = False
                ntbMaxValue.Visible = False
            End If
        End If
    End Sub

    Private Function CheckWithinRange(ByVal dblValue As Double) As Boolean
        Dim boolWithinRange As Boolean = True
        Dim strCriteria = String.Empty
        Dim strMsgTitle As String = ERRTITLE_ValueOutOfRange
        Dim strMsg As String

        If String.IsNullOrEmpty(Target.OpMin) = False Then
            Select Case Target.OpMin
                Case CONST_LargerThan
                    If dblValue <= Target.MinValue Then
                        boolWithinRange = False
                        strCriteria = LANG_Larger & Target.MinValue.ToString & " "
                    End If
                Case CONST_LargerThanOrEqual
                    If dblValue < Target.MinValue Then
                        boolWithinRange = False
                        strCriteria = LANG_LargerOrEqual & Target.MinValue.ToString & " "
                    End If

            End Select
        End If
        If String.IsNullOrEmpty(Target.OpMax) = False Then
            Select Case Target.OpMax
                Case CONST_SmallerThan
                    If dblValue >= Target.MaxValue Or boolWithinRange = False Then
                        boolWithinRange = False
                        If String.IsNullOrEmpty(strCriteria) = False Then strCriteria &= LANG_And
                        strCriteria &= LANG_Smaller & Target.MaxValue.ToString
                    End If
                Case CONST_SmallerThanOrEqual
                    If dblValue > Target.MaxValue Or boolWithinRange = False Then
                        boolWithinRange = False
                        If String.IsNullOrEmpty(strCriteria) Then strCriteria &= LANG_And
                        strCriteria &= LANG_Smaller & Target.MaxValue.ToString
                    End If

            End Select
        End If
        If boolWithinRange = False Then
            strMsg = String.Format(ERR_ValueOutOfRange, strCriteria)

            MsgBox(strMsg, MsgBoxStyle.Information, strMsgTitle)
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If String.IsNullOrEmpty(Target.OpMin) = False And String.IsNullOrEmpty(Target.OpMax) = False Then
            If Target.MinValue > Target.MaxValue Then
                MsgBox(ERR_MinValueLargerThanMaxValue)
                Return
            Else
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
            End If
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
        End If

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cmbOpMin_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOpMin.SelectedIndexChanged
        If cmbOpMin.SelectedValue = CONST_Equals Then
            Me.lblMaxValue.Visible = False
            Me.lblMaxValueDescription.Visible = False
            Me.cmbOpMax.Visible = False
            Me.ntbMaxValue.Visible = False
        ElseIf TargetSystem <> Indicator.TargetSystems.Simple Then
            Me.lblMaxValue.Visible = True
            Me.lblMaxValueDescription.Visible = True
            Me.cmbOpMax.Visible = True
            Me.ntbMaxValue.Visible = True
        End If
    End Sub

    Private Sub ntbMinValue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbMinValue.Validating
        If String.IsNullOrEmpty(ntbMinValue.Text) = False Then
            Dim dblValue As Double = ParseDouble(ntbMinValue.Text)
            Dim boolWithinRange As Boolean = True
            Dim strCriteria = String.Empty
            Dim strMsgTitle As String = ERRTITLE_ValueOutOfRange
            Dim strMsg As String

            If String.IsNullOrEmpty(Target.OpMax) = False Then
                Select Case Target.OpMax
                    Case CONST_SmallerThan
                        If dblValue >= Target.MaxValue Or boolWithinRange = False Then
                            boolWithinRange = False
                            If String.IsNullOrEmpty(strCriteria) = False Then strCriteria &= LANG_And
                            strCriteria &= LANG_Smaller & Target.MaxValue.ToString
                        End If
                    Case CONST_SmallerThanOrEqual
                        If dblValue > Target.MaxValue Or boolWithinRange = False Then
                            boolWithinRange = False
                            If String.IsNullOrEmpty(strCriteria) Then strCriteria &= LANG_And
                            strCriteria &= LANG_Smaller & Target.MaxValue.ToString
                        End If

                End Select
            End If
            If boolWithinRange = False Then
                strMsg = String.Format(ERR_ValueOutOfRange, strCriteria)

                MsgBox(strMsg, MsgBoxStyle.Information, strMsgTitle)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub tbMaxValue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbMaxValue.Validating
        If String.IsNullOrEmpty(ntbMaxValue.Text) = False Then
            Dim dblValue As Double = ParseDouble(ntbMaxValue.Text)
            Dim boolWithinRange As Boolean = True
            Dim strCriteria = String.Empty
            Dim strMsgTitle As String = ERRTITLE_ValueOutOfRange
            Dim strMsg As String


            If String.IsNullOrEmpty(Target.OpMin) = False Then
                Select Case Target.OpMin
                    Case CONST_LargerThan
                        If dblValue <= Target.MinValue Then
                            boolWithinRange = False
                            strCriteria = LANG_Larger & Target.MinValue.ToString & " "
                        End If
                    Case CONST_LargerThanOrEqual
                        If dblValue < Target.MinValue Then
                            boolWithinRange = False
                            strCriteria = LANG_LargerOrEqual & Target.MinValue.ToString & " "
                        End If

                End Select
            End If

            If boolWithinRange = False Then
                strMsg = String.Format(ERR_ValueOutOfRange, strCriteria)

                MsgBox(strMsg, MsgBoxStyle.Information, strMsgTitle)
                e.Cancel = True
            End If
        End If
    End Sub
End Class
