Imports System.Windows.Forms

Public Class DialogValueRange
    Private sngRangeMinValue As Single
    Private strRangeOpMin As String
    Private sngRangeMaxValue As Single
    Private strRangeOpMax As String
    Private datDeadline As Date

#Region "Properties"
    Public Property RangeMinValue() As Single
        Get
            Return sngRangeMinValue
        End Get
        Set(ByVal value As Single)
            sngRangeMinValue = value
        End Set
    End Property

    Public Property RangeOpMin() As String
        Get
            Return strRangeOpMin
        End Get
        Set(ByVal value As String)
            strRangeOpMin = value
        End Set
    End Property

    Public Property RangeMaxValue() As Single
        Get
            Return sngRangeMaxValue
        End Get
        Set(ByVal value As Single)
            sngRangeMaxValue = value
        End Set
    End Property

    Public Property RangeOpMax() As String
        Get
            Return strRangeOpMax
        End Get
        Set(ByVal value As String)
            strRangeOpMax = value
        End Set
    End Property

    Public Property Deadline As Date
        Get
            Return datDeadline
        End Get
        Set(value As Date)
            datDeadline = value
        End Set
    End Property
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal sngMinValue As Single, ByVal strOpMin As String, ByVal sngMaxValue As Single, ByVal strOpMax As String, ByVal strUnit As String, ByVal intNrDecimals As Integer)
        InitializeComponent()

        Me.RangeMinValue = sngMinValue
        Me.RangeOpMin = strOpMin
        Me.RangeMaxValue = sngMaxValue
        Me.RangeOpMax = strOpMax
        Me.Deadline = Deadline

        With cmbOpMin
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.Add(CONST_LargerThan)
            .Items.Add(CONST_LargerThanOrEqual)
            If String.IsNullOrEmpty(RangeOpMin) = False Then .SelectedItem = RangeOpMin
        End With
        With tbMinValue
            .Unit = strUnit
            .NrDecimals = intNrDecimals
            If String.IsNullOrEmpty(RangeOpMin) = False Then
                .Text = RangeMinValue
                '.DisplayAsUnit()
            End If

        End With
        With cmbOpMax
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.Add(CONST_SmallerThan)
            .Items.Add(CONST_SmallerThanOrEqual)
            If String.IsNullOrEmpty(RangeOpMax) = False Then .SelectedItem = RangeOpMax
        End With
        With tbMaxValue
            .Unit = strUnit
            .NrDecimals = intNrDecimals
            If String.IsNullOrEmpty(RangeOpMax) = False Then
                .Text = RangeMaxValue
                '.DisplayAsUnit()
            End If
        End With
    End Sub

    Private Function CheckWithinRange(ByVal sngValue As Single) As Boolean
        Dim boolWithinRange As Boolean = True
        Dim strCriteria = String.Empty
        Dim strMsgTitle As String = ERRTITLE_ValueOutOfRange
        Dim strMsg As String

        If String.IsNullOrEmpty(RangeOpMin) = False Then
            Select Case RangeOpMin
                Case CONST_LargerThan
                    If sngValue <= sngRangeMinValue Then
                        boolWithinRange = False
                        strCriteria = LANG_Larger & sngRangeMinValue.ToString & " "
                    End If
                Case CONST_LargerThanOrEqual
                    If sngValue < sngRangeMinValue Then
                        boolWithinRange = False
                        strCriteria = LANG_LargerOrEqual & sngRangeMinValue.ToString & " "
                    End If

            End Select
        End If
        If String.IsNullOrEmpty(RangeOpMax) = False Then
            Select Case RangeOpMax
                Case CONST_SmallerThan
                    If sngValue >= sngRangeMaxValue Or boolWithinRange = False Then
                        boolWithinRange = False
                        If String.IsNullOrEmpty(strCriteria) = False Then strCriteria &= LANG_And
                        strCriteria &= LANG_Smaller & sngRangeMaxValue.ToString
                    End If
                Case CONST_SmallerThanOrEqual
                    If sngValue > sngRangeMaxValue Or boolWithinRange = False Then
                        boolWithinRange = False
                        If String.IsNullOrEmpty(strCriteria) Then strCriteria &= LANG_And
                        strCriteria &= LANG_Smaller & sngRangeMaxValue.ToString
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
        If String.IsNullOrEmpty(RangeOpMin) = False And String.IsNullOrEmpty(RangeOpMax) = False Then
            If RangeMinValue > RangeMaxValue Then
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

    Private Sub cmbOpMin_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOpMin.Validated
        If String.IsNullOrEmpty(cmbOpMin.Text) = False Then Me.RangeOpMin = cmbOpMin.Text
    End Sub

    Private Sub tbMinValue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles tbMinValue.Validating
        If String.IsNullOrEmpty(tbMinValue.Text) = False Then
            If CheckWithinRange(tbMinValue.DoubleValue) = False Then e.Cancel = True
        End If
    End Sub

    Private Sub tbMinValue_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMinValue.Validated
        Single.TryParse(tbMinValue.Text, RangeMinValue)
    End Sub

    Private Sub cmbOpMax_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOpMax.Validated
        If String.IsNullOrEmpty(cmbOpMax.Text) = False Then Me.RangeOpMax = cmbOpMax.Text
    End Sub

    Private Sub tbMaxValue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles tbMaxValue.Validating
        If String.IsNullOrEmpty(tbMaxValue.Text) = False Then
            If CheckWithinRange(tbMaxValue.DoubleValue) = False Then e.Cancel = True
        End If
    End Sub

    Private Sub tbMaxValue_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMaxValue.Validated
        Single.TryParse(tbMaxValue.Text, RangeMaxValue)
    End Sub
End Class
