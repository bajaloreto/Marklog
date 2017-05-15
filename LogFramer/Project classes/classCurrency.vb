Public Class Currency
    Private sngAmount As Single
    Private strCurrencyCode As String = My.Settings.setDefaultCurrency
    Private sngExchangeRate As Single = 1

    Public Sub New()

    End Sub

    Public Overrides Function ToString() As String
        Dim strFormat As String
        strFormat = String.Format("#,##0.00 {0}", Me.CurrencyCode)
        Return Amount.ToString(strFormat)
    End Function

    Public Sub New(ByVal amount As Single, ByVal currency As String, Optional ByVal exchangerate As Single = 1)
        Me.Amount = amount
        Me.CurrencyCode = currency
        Me.ExchangeRate = exchangerate
    End Sub

    Public Property Amount() As Single
        Get
            Return sngAmount
        End Get
        Set(ByVal value As Single)
            sngAmount = value
        End Set
    End Property

    Public Property CurrencyCode() As String
        Get
            Return strCurrencyCode
        End Get
        Set(ByVal value As String)
            strCurrencyCode = Left(value, 3).ToUpper
        End Set
    End Property

    Public Property ExchangeRate() As Single
        Get
            Return sngExchangeRate
        End Get
        Set(ByVal value As Single)
            sngExchangeRate = value
        End Set
    End Property
End Class

