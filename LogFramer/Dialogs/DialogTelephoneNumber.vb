Imports System.Windows.Forms

Public Class DialogTelephoneNumber
    Private objTelephoneNumber As TelephoneNumber
    Private bindTelephoneNumber As New BindingSource

    Public Property TelephoneNumber() As TelephoneNumber
        Get
            Return objTelephoneNumber
        End Get
        Set(ByVal value As TelephoneNumber)
            objTelephoneNumber = value
        End Set
    End Property

    Public Sub New(ByVal address As TelephoneNumber)
        InitializeComponent()

        Me.TelephoneNumber = address
        If Me.TelephoneNumber IsNot Nothing Then
            bindTelephoneNumber.DataSource = Me.TelephoneNumber
            tbNumber.DataBindings.Add(New Binding("Text", bindTelephoneNumber, "Number"))

            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_TelephoneNumberTypes)
                .DataBindings.Add(New Binding("SelectedIndex", bindTelephoneNumber, "Type"))
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
