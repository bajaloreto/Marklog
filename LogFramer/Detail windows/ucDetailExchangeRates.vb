Public Class DetailExchangeRates
    Public Event CloseButtonClicked()
    Public Event ExchangeRateUpdated()

    Private objBudget As Budget
    Private bindExchangeRates As New BindingSource

    Public Property Budget As Budget
        Get
            Return objBudget
        End Get
        Set(ByVal value As Budget)
            objBudget = value
            'LoadColumns()
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal budget As Budget)
        InitializeComponent()

        Dim lstCurrencyCodes As List(Of IdValuePair) = LoadCurrencyCodesList()

        Me.Budget = budget
        If Me.Budget IsNot Nothing Then
            bindExchangeRates.DataSource = Me.Budget

            With cmbDefaultCurrencyCode
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = lstCurrencyCodes
                .DisplayMember = "Value"
                .ValueMember = "Id"
                .DataBindings.Add("SelectedValue", bindExchangeRates, "DefaultCurrencyCode")
            End With
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        RaiseEvent CloseButtonClicked()
    End Sub

    Private Sub DetailExchangeRate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Budget IsNot Nothing Then
            If String.IsNullOrEmpty(Budget.DefaultCurrencyCode) Then Budget.DefaultCurrencyCode = My.Settings.setDefaultCurrency
            With dgvExchangeRates
                .ExchangeRates = Budget.ExchangeRates
                .DefaultCurrencyCode = Budget.DefaultCurrencyCode
                .Reload()
            End With
        End If
    End Sub

    Private Sub cmbDefaultCurrencyCode_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbDefaultCurrencyCode.SelectedValueChanged
        If Me.Budget IsNot Nothing Then
            With dgvExchangeRates
                .DefaultCurrencyCode = Budget.DefaultCurrencyCode
                '.Reload()
            End With
        End If
    End Sub

    Private Sub dgvExchangeRates_ExchangeRateUpdated() Handles dgvExchangeRates.ExchangeRateUpdated
        RaiseEvent ExchangeRateUpdated()
    End Sub
End Class
