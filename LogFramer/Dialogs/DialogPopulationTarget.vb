Imports System.Windows.Forms

Public Class DialogPopulationTarget
    Private objPopulationTarget As PopulationTarget
    Private objValueRange As ValueRange
    Private PopulationTargetBindingSource As New BindingSource

#Region "Properties"
    Public Property PopulationTarget() As PopulationTarget
        Get
            Return objPopulationTarget
        End Get
        Set(ByVal value As PopulationTarget)
            objPopulationTarget = value
        End Set
    End Property

    Public Property ValueRange As ValueRange
        Get
            Return objValueRange
        End Get
        Set(ByVal value As ValueRange)
            objValueRange = value
        End Set
    End Property
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal populationtarget As PopulationTarget, ByVal valuerange As ValueRange, ByVal intNrDecimals As Integer)
        InitializeComponent()

        Me.PopulationTarget = populationtarget
        Me.ValueRange = valuerange

        If Me.PopulationTarget IsNot Nothing Then
            PopulationTargetBindingSource.DataSource = Me.PopulationTarget
            With ntbValue
                .NrDecimals = intNrDecimals
                .Unit = "%"
                .DataBindings.Add("Text", PopulationTargetBindingSource, "Percentage")
                .Focus()
                .SelectAll()
            End With
        End If
    End Sub

    Private Function CheckWithinRange(ByVal sngValue As Single) As Boolean
        Dim boolWithinRange As Boolean = True
        Dim strCriteria = String.Empty
        Dim strMsgTitle As String = ERRTITLE_ValueOutOfRange
        Dim strMsg As String

        If Me.ValueRange IsNot Nothing Then
            If String.IsNullOrEmpty(ValueRange.OpMin) = False Then
                Select Case ValueRange.OpMin
                    Case CONST_LargerThan
                        If sngValue <= ValueRange.MinValue Then
                            boolWithinRange = False
                            strCriteria = LANG_Larger & ValueRange.MinValue.ToString & " "
                        End If
                    Case CONST_LargerThanOrEqual
                        If sngValue < ValueRange.MinValue Then
                            boolWithinRange = False
                            strCriteria = LANG_LargerOrEqual & ValueRange.MinValue.ToString & " "
                        End If

                End Select
            End If
            If String.IsNullOrEmpty(ValueRange.OpMax) = False Then
                Select Case ValueRange.OpMax
                    Case CONST_SmallerThan
                        If sngValue >= ValueRange.MaxValue Or boolWithinRange = False Then
                            boolWithinRange = False
                            If String.IsNullOrEmpty(strCriteria) = False Then strCriteria &= LANG_And
                            strCriteria &= LANG_Smaller & ValueRange.MaxValue.ToString
                        End If
                    Case CONST_SmallerThanOrEqual
                        If sngValue > ValueRange.MaxValue Or boolWithinRange = False Then
                            boolWithinRange = False
                            If String.IsNullOrEmpty(strCriteria) Then strCriteria &= LANG_And
                            strCriteria &= LANG_Smaller & ValueRange.MaxValue.ToString
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
        Else
            Return True
        End If
    End Function

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ntbValue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbValue.Validating
        If String.IsNullOrEmpty(ntbValue.Text) = False Then
            If CheckWithinRange(ntbValue.DoubleValue) = False Then e.Cancel = True
        End If
    End Sub
End Class
