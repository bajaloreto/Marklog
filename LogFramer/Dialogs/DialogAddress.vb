Imports System.Windows.Forms

Public Class DialogAddress
    Private objAddress As Address
    Private bindAddress As New BindingSource

    Public Property Address() As Address
        Get
            Return objAddress
        End Get
        Set(ByVal value As Address)
            objAddress = value
        End Set
    End Property

    Public Sub New(ByVal address As Address)
        InitializeComponent()

        Me.Address = address
        If Me.Address IsNot Nothing Then
            bindAddress.DataSource = Me.Address
            tbStreet.DataBindings.Add(New Binding("Text", bindAddress, "Street"))
            tbNumber.DataBindings.Add(New Binding("Text", bindAddress, "Number"))
            tbPostNumber.DataBindings.Add(New Binding("Text", bindAddress, "PostNumber"))
            tbDistrict.DataBindings.Add(New Binding("Text", bindAddress, "District"))
            tbTown.DataBindings.Add(New Binding("Text", bindAddress, "Town"))
            tbCountry.DataBindings.Add(New Binding("Text", bindAddress, "Country"))

            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_AddressTypes)
                .DataBindings.Add(New Binding("SelectedIndex", bindAddress, "Type"))
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
