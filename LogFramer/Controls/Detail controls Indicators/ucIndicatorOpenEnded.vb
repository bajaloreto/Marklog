Public Class IndicatorOpenEnded
    Friend WithEvents OpenEndedDetailBindingSource As New BindingSource

    Private indCurrentIndicator As Indicator

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()

        Me.CurrentIndicator = indicator

        If CurrentIndicator.Indicators.Count = 0 Then
            LoadItems()
        Else
            gbWhiteSpace.Visible = False
        End If
    End Sub

    Protected Property CurrentIndicator() As Indicator
        Get
            Return indCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            indCurrentIndicator = value
        End Set
    End Property

    Public Sub LoadItems()

        If CurrentIndicator IsNot Nothing Then
            OpenEndedDetailBindingSource.DataSource = CurrentIndicator.OpenEndedDetail

            With cmbWhiteSpace
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_WhiteSpaceValues)
                .DataBindings.Add("SelectedIndex", OpenEndedDetailBindingSource, "WhiteSpace")
            End With
        End If
    End Sub


End Class
