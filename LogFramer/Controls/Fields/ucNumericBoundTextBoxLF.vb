Imports System.Globalization

Public Class NumericBoundTextBoxLF
    Inherits TextBox

    Protected Overrides Sub OnKeyPress(ByVal e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)
        ' Restricts the entry of characters to digits (including hex),
        ' the negative sign, the e decimal point, and editing keystrokes (backspace).

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
            'ElseIf e.KeyChar = "%" Then
            '    ' Percentage sign is ok
            'ElseIf Me.boolSpaceOK AndAlso e.KeyChar = " "c Then
            '    ' Space is ok
        Else
            e.Handled = True
        End If

    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        MyBase.OnTextChanged(e)
        UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.ValueChanged)
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)
        UndoRedo.CurrentControlValidated(Me)
    End Sub
End Class
