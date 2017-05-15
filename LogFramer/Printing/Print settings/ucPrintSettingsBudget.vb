Imports System.Drawing.Printing

Public Class PrintSettingsBudget
    Private intBudgetYearIndex As Integer
    Private boolShowDurationColumns, boolShowLocalCurrencyColumns, boolShowExchangeRates As Boolean
    Private intPagesWide As Integer
    Public boolLoad As Boolean

    Public Event PrintBudgetSetupChanged(ByVal sender As Object, ByVal e As PrintBudgetSetupChangedEventArgs)

    Public Property BudgetYearIndex() As Integer
        Get
            Return intBudgetYearIndex
        End Get
        Set(ByVal value As Integer)
            intBudgetYearIndex = value
        End Set
    End Property

    Public Property ShowDurationColumns As Boolean
        Get
            Return boolShowDurationColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowDurationColumns = value
        End Set
    End Property

    Public Property ShowLocalCurrencyColumns As Boolean
        Get
            Return boolShowLocalCurrencyColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowLocalCurrencyColumns = value
        End Set
    End Property

    Public Property ShowExchangeRates As Boolean
        Get
            Return boolShowExchangeRates
        End Get
        Set(ByVal value As Boolean)
            boolShowExchangeRates = value
        End Set
    End Property

    Public Property PagesWide As Integer
        Get
            Return intPagesWide
        End Get
        Set(value As Integer)
            intPagesWide = value
        End Set
    End Property

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        With cmbBudgetYear
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.AddRange(ListBudgetYears)
            .SelectedIndex = 0
        End With

        chkShowDurationColumns.Checked = My.Settings.setPrintBudgetShowDurationColumns
        chkShowLocalCurrencyColumns.Checked = My.Settings.setPrintBudgetShowLocalCurrencyColumns
        chkShowExchangeRates.Checked = My.Settings.setPrintBudgetShowExchangeRates
        nudPagesWide.Value = My.Settings.setPrintBudgetPagesWide

        boolLoad = False
    End Sub

    Private Function ListBudgetYears() As String()
        Dim intbudgetIndex As Integer
        Dim strBudgetTitle As String
        Dim lstBudgetYears As New ArrayList

        For Each selBudgetYear As BudgetYear In CurrentLogFrame.Budget.BudgetYears
            If intBudgetIndex = 0 Then
                strBudgetTitle = LANG_TotalBudget

            Else
                strBudgetTitle = selBudgetYear.BudgetYear.ToString("yyyy")
            End If
            lstBudgetYears.Add(strBudgetTitle)

            intbudgetIndex += 1
        Next
        lstBudgetYears.Add(LANG_AllBudgetYears)
        Return lstBudgetYears.ToArray(GetType(String))
    End Function

    Private Sub cmbBudgetYear_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbBudgetYear.SelectedIndexChanged
        If cmbBudgetYear.SelectedIndex = cmbBudgetYear.Items.Count - 1 Then
            Me.BudgetYearIndex = -1
        Else
            Me.BudgetYearIndex = cmbBudgetYear.SelectedIndex
        End If

        If boolLoad = False Then
            RaiseEvent PrintBudgetSetupChanged(Me, New PrintBudgetSetupChangedEventArgs(BudgetYearIndex, ShowDurationColumns, ShowLocalCurrencyColumns, ShowExchangeRates, PagesWide))
        End If
    End Sub

    Private Sub chkShowDurationColumns_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkShowDurationColumns.CheckedChanged
        Me.ShowDurationColumns = chkShowDurationColumns.Checked

        If boolLoad = False Then
            My.Settings.setPrintBudgetShowDurationColumns = Me.ShowDurationColumns
            RaiseEvent PrintBudgetSetupChanged(Me, New PrintBudgetSetupChangedEventArgs(BudgetYearIndex, ShowDurationColumns, ShowLocalCurrencyColumns, ShowExchangeRates, PagesWide))
        End If
    End Sub

    Private Sub chkShowLocalCurrencyColumns_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowLocalCurrencyColumns.CheckedChanged
        Me.ShowLocalCurrencyColumns = chkShowLocalCurrencyColumns.Checked

        If boolLoad = False Then
            My.Settings.setPrintBudgetShowLocalCurrencyColumns = Me.ShowLocalCurrencyColumns
            RaiseEvent PrintBudgetSetupChanged(Me, New PrintBudgetSetupChangedEventArgs(BudgetYearIndex, ShowDurationColumns, ShowLocalCurrencyColumns, ShowExchangeRates, PagesWide))
        End If
    End Sub

    Private Sub chkShowExchangeRates_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowExchangeRates.CheckedChanged
        Me.ShowExchangeRates = chkShowExchangeRates.Checked

        If boolLoad = False Then
            My.Settings.setPrintBudgetShowExchangeRates = Me.ShowExchangeRates
            RaiseEvent PrintBudgetSetupChanged(Me, New PrintBudgetSetupChangedEventArgs(BudgetYearIndex, ShowDurationColumns, ShowLocalCurrencyColumns, ShowExchangeRates, PagesWide))
        End If
    End Sub

    Private Sub nudPagesWide_ValueChanged(sender As System.Object, e As System.EventArgs) Handles nudPagesWide.ValueChanged
        Me.PagesWide = nudPagesWide.Value

        If boolLoad = False Then
            My.Settings.setPrintBudgetPagesWide = Me.PagesWide
            RaiseEvent PrintBudgetSetupChanged(Me, New PrintBudgetSetupChangedEventArgs(BudgetYearIndex, ShowDurationColumns, ShowLocalCurrencyColumns, ShowExchangeRates, PagesWide))
        End If
    End Sub
End Class

Public Class PrintBudgetSetupChangedEventArgs
    Inherits EventArgs

    Public Property BudgetYearIndex As Integer
    Public Property ShowDurationColumns As Boolean
    Public Property ShowLocalCurrencyColumns As Boolean
    Public Property ShowExchangeRates As Boolean
    Public Property PagesWide As Integer

    Public Sub New(ByVal intBudgetYearIndex As Integer, ByVal boolShowDurationColumns As Boolean, ByVal boolShowLocalCurrencyColumns As Boolean, ByVal boolShowExchangeRates As Boolean, _
                   ByVal intPagesWide As Integer)
        MyBase.New()

        BudgetYearIndex = intBudgetYearIndex
        ShowDurationColumns = boolShowDurationColumns
        ShowLocalCurrencyColumns = boolShowLocalCurrencyColumns
        ShowExchangeRates = boolShowExchangeRates
        PagesWide = intPagesWide
    End Sub
End Class
