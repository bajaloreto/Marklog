
Public Class NumericTextBoxCell
    Inherits DataGridViewTextBoxCell

    Public Overrides ReadOnly Property EditType() As Type
        Get
            ' Return the type of the editing contol that 
            ' the DropDownListCell uses.
            Return GetType(NumericTextBoxEditingControl)
        End Get
    End Property
End Class

Public Class NumericTextBoxColumn
    Inherits DataGridViewTextBoxColumn

    Public Sub New()
        'Set the type used in the DataGridView
        Me.CellTemplate = New NumericTextBoxCell
    End Sub

End Class


Public Class NumericTextBoxEditingControl
    Inherits DataGridViewTextBoxEditingControl

    Private SpaceOK As Boolean = False

    Public Sub New()
        MyBase.New()
        
    End Sub

    ' Restricts the entry of characters to digits (including hex),
    ' the negative sign, the e decimal point, and editing keystrokes (backspace).
    Protected Overrides Sub OnKeyPress(ByVal e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)

        Dim numberFormatInfo As System.Globalization.NumberFormatInfo = System.Globalization.CultureInfo.CurrentCulture.NumberFormat
        Dim decimalSeparator As String = numberFormatInfo.NumberDecimalSeparator
        Dim groupSeparator As String = numberFormatInfo.NumberGroupSeparator
        Dim negativeSign As String = numberFormatInfo.NegativeSign

        Dim keyInput As String = e.KeyChar.ToString()

        If [Char].IsDigit(e.KeyChar) Then
            ' Digits are OK
        ElseIf keyInput.Equals(decimalSeparator) OrElse keyInput.Equals(groupSeparator) OrElse keyInput.Equals(negativeSign) Then
            ' Decimal separator is OK
        ElseIf e.KeyChar = vbBack Or e.KeyChar = vbTab Then
            ' Backspace key is OK

        ElseIf Me.SpaceOK AndAlso e.KeyChar = " "c Then

        Else
            ' Consume this invalid key and beep.
            e.Handled = True
        End If

    End Sub

    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        MyBase.OnClick(e)
        Me.SelectAll()
    End Sub

    Public ReadOnly Property IntValue() As Integer

        Get
            If Me.Text <> "" Then
                Return Int32.Parse(Me.Text)
            End If
        End Get

    End Property


    Public ReadOnly Property DecimalValue() As Decimal
        Get
            If Me.Text <> "" Then
                Return [Decimal].Parse(Me.Text)
            End If
        End Get
    End Property

    Public Property AllowSpace() As Boolean

        Get
            Return Me.SpaceOK
        End Get
        Set(ByVal value As Boolean)
            Me.SpaceOK = value
        End Set
    End Property

End Class
