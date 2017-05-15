Imports System.Globalization

Public Class NumericTextBoxLF
    Inherits TextBox

    Private sngValue As Single
    Private boolSpaceOK As Boolean = True
    Private intNrDecimals As Integer
    Private strUnit As String
    Private boolIsCurrency As Boolean
    Private boolIsPercentage As Boolean

    Public Property IsCurrency() As Boolean
        Get
            Return boolIsCurrency
        End Get
        Set(ByVal value As Boolean)
            boolIsCurrency = value
            If value = True Then
                boolSpaceOK = True
                boolIsPercentage = False
            End If

        End Set
    End Property

    Public Property IsPercentage() As Boolean
        Get
            Return boolIsPercentage
        End Get
        Set(ByVal value As Boolean)
            boolIsPercentage = value
            If value = True Then
                boolSpaceOK = True
                boolIsCurrency = False
            End If
        End Set
    End Property

    Public ReadOnly Property IntegerValue() As Integer
        Get
            Dim strText As String = Me.Text
            Dim intValue As Integer
            If strText <> "" Then
                If Integer.TryParse(strText, intValue) = True Then
                    Return intValue
                Else
                    Return 0
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property DecimalValue() As Decimal
        Get
            Dim strText As String = Me.Text
            Dim decValue As Decimal
            If strText <> "" Then
                If Decimal.TryParse(strText, decValue) = True Then
                    Return decValue
                Else
                    Return 0
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property SingleValue() As Single
        Get
            Dim strText As String = Me.Text
            Dim sngValue As Single
            If strText <> "" Then
                If Single.TryParse(strText, sngValue) = True Then
                    Return sngValue
                Else
                    Return 0
                End If
            End If
        End Get
    End Property

    Public Property AllowSpace() As Boolean

        Get
            Return Me.boolSpaceOK
        End Get
        Set(ByVal value As Boolean)
            Me.boolSpaceOK = value
        End Set
    End Property

    ' Restricts the entry of characters to digits (including hex),
    ' the negative sign, the e decimal point, and editing keystrokes (backspace).
    Protected Overrides Sub OnKeyPress(ByVal e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)

        Dim numberFormatInfo As NumberFormatInfo = System.Globalization.CultureInfo.CurrentCulture.NumberFormat
        Dim decimalSeparator As String = numberFormatInfo.NumberDecimalSeparator
        Dim groupSeparator As String = numberFormatInfo.NumberGroupSeparator
        Dim negativeSign As String = numberFormatInfo.NegativeSign

        Dim keyInput As String = e.KeyChar.ToString()

        If [Char].IsDigit(e.KeyChar) Then
            ' Digits are OK
        ElseIf keyInput.Equals(decimalSeparator) OrElse keyInput.Equals(groupSeparator) OrElse keyInput.Equals(negativeSign) Then
            ' Decimal separator is OK
        ElseIf e.KeyChar = vbBack Or e.KeyChar = vbTab Then
            ' Backspace and tab keys are OK
        ElseIf e.KeyChar = "%" Then
            ' Percentage sign is ok
        ElseIf Me.boolSpaceOK AndAlso e.KeyChar = " "c Then
            ' Space is ok
        Else
            If Me.IsCurrency = True AndAlso Me.Text.Contains(" ") = True Then

                If Me.Text.IndexOf(e.KeyChar) > Me.Text.IndexOf(" ") And [Char].IsLetter(e.KeyChar) = True Then
                    Dim strSplit() As String = Me.Text.Split(" ")
                    Dim strCurrency As String = strSplit(strSplit.Length - 1)
                    If strCurrency.Length >= 3 Then
                        e.Handled = True
                    End If
                End If
            Else
                e.Handled = True
            End If
        End If

    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        MyBase.OnClick(e)
        If Me.DecimalValue = 0 Then Me.SelectAll()

        'Me.SelectAll()
    End Sub

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        MyBase.OnTextChanged(e)
        If CurrentControl Is Me Then UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.ValueChanged)
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)
        UndoRedo.CurrentControlValidated(Me)
    End Sub
End Class
