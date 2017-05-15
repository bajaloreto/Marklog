Imports System.Windows.Forms

Public Class DialogWebsite
    Private objWebsite As Website
    Private bindWebsite As New BindingSource

    Public Property Website() As Website
        Get
            Return objWebsite
        End Get
        Set(ByVal value As Website)
            objWebsite = value
        End Set
    End Property

    Public Sub New(ByVal address As Website)
        InitializeComponent()

        Me.Website = address
        If Me.Website IsNot Nothing Then
            bindWebsite.DataSource = Me.Website
            tbWebsiteName.DataBindings.Add(New Binding("Text", bindWebsite, "WebsiteName"))
            tbWebsitePath.DataBindings.Add(New Binding("Text", bindWebsite, "WebsiteUrl"))
            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_WebsiteTypes)
                .DataBindings.Add(New Binding("SelectedIndex", bindWebsite, "Type"))
            End With
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class
