Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class NumericTextBox
    Inherits TextBox

    Private intValueType As Integer
    Private boolSpaceOK As Boolean
    Private boolSetDecimals As Boolean
    Private intNrDecimals As Integer
    Private strUnit As String
    Private boolIsCurrency As Boolean
    Private boolIsPercentage As Boolean
    Private intIntegerValue As Integer
    Private sngSingleValue As Single
    Private dblDoubleValue As Double

    Public Sub New()

    End Sub

    Public Enum ValueTypes
        IntegerValue
        SingleValue
        DoubleValue
    End Enum

#Region "Settings"
    Public Property ValueType As Integer
        Get
            Return intValueType
        End Get
        Set(value As Integer)
            intValueType = value

            If intValueType = ValueTypes.IntegerValue Then intNrDecimals = 0
        End Set
    End Property

    Public Property IsCurrency() As Boolean
        Get
            Return boolIsCurrency
        End Get
        Set(ByVal value As Boolean)
            boolIsCurrency = value
            If value = True Then
                boolSpaceOK = True
                boolIsPercentage = False
                intValueType = ValueTypes.SingleValue
                intNrDecimals = 2
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

    Public Property AllowSpace() As Boolean

        Get
            Return boolSpaceOK
        End Get
        Set(ByVal value As Boolean)
            boolSpaceOK = value
        End Set
    End Property

    Public Property SetDecimals() As Boolean

        Get
            Return boolSetDecimals
        End Get
        Set(ByVal value As Boolean)
            boolSetDecimals = value
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
#End Region

#Region "Values"
    Public Property IntegerValue As Integer
        Get
            Return intIntegerValue
        End Get
        Set(ByVal value As Integer)
            Dim boolSetText As Boolean
            If value <> intIntegerValue Then boolSetText = True
            intIntegerValue = value
            If boolSetText = True Then Me.Text = GetText()
        End Set
    End Property

    Public Property SingleValue As Single
        Get
            Return sngSingleValue
        End Get
        Set(ByVal value As Single)
            Dim boolSetText As Boolean
            If value <> sngSingleValue Then boolSetText = True
            sngSingleValue = value
            If boolSetText = True Then Me.Text = GetText()
        End Set
    End Property

    Public Property DoubleValue As Double
        Get
            Return dblDoubleValue
        End Get
        Set(ByVal value As Double)
            Dim boolSetText As Boolean
            If value <> dblDoubleValue Then boolSetText = True
            dblDoubleValue = value
            If boolSetText = True Then Me.Text = GetText()
        End Set
    End Property
#End Region
    
#Region "Methods"
    Private Function GetValueString(ByVal strText As String) As String
        Dim strPattern As String = "[^0-9.,-]+"
        Dim strTextSections As String() = Regex.Split(strText, strPattern)
        Dim strValue As String = String.Empty

        For i = 0 To strTextSections.Count - 1
            If Val(strTextSections(i)) > 0 Then strValue = strTextSections(i)
        Next

        Return strValue
    End Function

    Private Function GetNumberFormatString() As String
        Dim strFormat As String
        If Me.IsPercentage Then
            If Me.SetDecimals Then
                strFormat = String.Format("P{0}", Me.NrDecimals)
            Else
                strFormat = "P"
            End If
        Else
            If Me.SetDecimals Then
                strFormat = String.Format("N{0}", Me.NrDecimals)
            Else
                strFormat = "N"
            End If
        End If

        Return strFormat
    End Function
#End Region

#Region "Events"
    Protected Overrides Sub OnEnter(e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

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

    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        MyBase.OnClick(e)
    End Sub

    Protected Overrides Sub OnValidating(ByVal e As System.ComponentModel.CancelEventArgs)
        Dim strValue As String = GetValueString(Me.Text)

        MyBase.OnValidating(e)

        If String.IsNullOrEmpty(strValue) = False Then
            Select Case ValueType
                Case ValueTypes.IntegerValue
                    Me.IntegerValue = ParseInteger(strValue)
                    If Me.IsPercentage = True Then IntegerValue /= 100
                Case ValueTypes.SingleValue
                    Me.SingleValue = ParseSingle(strValue)
                    If Me.IsPercentage = True Then SingleValue /= 100
                Case ValueTypes.DoubleValue
                    Me.DoubleValue = ParseDouble(strValue)
                    If Me.IsPercentage = True Then DoubleValue /= 100
            End Select
        End If


    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)

        Me.Text = GetText()
    End Sub

    Public Function GetText() As String
        Dim strText As String = String.Empty

        Select Case ValueType
            Case ValueTypes.IntegerValue
                strText = IntegerValue.ToString(GetNumberFormatString)
            Case ValueTypes.SingleValue
                strText = SingleValue.ToString(GetNumberFormatString)
            Case ValueTypes.DoubleValue
                strText = DoubleValue.ToString(GetNumberFormatString)
        End Select

        If Me.IsPercentage = False Then
            If String.IsNullOrEmpty(Me.Unit) = False Then strText = String.Format("{0} {1}", strText, Me.Unit)
        End If

        Return strText
    End Function

    Public Overrides Function ToString() As String
        Dim strFormat As String = GetNumberFormatString()

        Select Case ValueType
            Case ValueTypes.IntegerValue
                Return IntegerValue.ToString(strFormat)
            Case ValueTypes.SingleValue
                Return SingleValue.ToString(strFormat)
            Case ValueTypes.DoubleValue
                Return DoubleValue.ToString(strFormat)
        End Select

        Return String.Empty
    End Function
#End Region

End Class
