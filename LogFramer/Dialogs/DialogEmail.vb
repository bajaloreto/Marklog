Imports System.Windows.Forms

Public Class DialogEmail
    Private objEmail As Email
    Private bindEmail As New BindingSource

    Public Property Email() As Email
        Get
            Return objEmail
        End Get
        Set(ByVal value As Email)
            objEmail = value
        End Set
    End Property

    Public Sub New(ByVal address As Email)
        InitializeComponent()

        Me.Email = address
        If Me.Email IsNot Nothing Then
            bindEmail.DataSource = Me.Email
            tbEmail.DataBindings.Add(New Binding("Text", bindEmail, "Email"))

            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_EmailTypes)
                .DataBindings.Add(New Binding("SelectedIndex", bindEmail, "Type"))
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
