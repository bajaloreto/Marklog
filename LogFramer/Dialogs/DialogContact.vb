Imports System.Windows.Forms

Public Class DialogContact
    Private objContact As Contact
    Private bindContact As New BindingSource

    Friend WithEvents lvAddresses As New ListViewAddresses
    Friend WithEvents lvTelephoneNumbers As New ListViewTelephoneNumbers
    Friend WithEvents lvEmails As New ListViewEmails
    Friend WithEvents lvWebsites As New ListViewWebsites

    Public Property Contact() As Contact
        Get
            Return objContact
        End Get
        Set(ByVal value As Contact)
            objContact = value
        End Set
    End Property

    Public Sub New(ByVal contact As Contact)
        Me.InitializeComponent()
        Me.Contact = contact
        Me.KeyPreview = True

        bindContact.DataSource = Me.Contact

        lvAddresses.Addresses = Me.Contact.Addresses
        lvTelephoneNumbers.TelephoneNumbers = Me.Contact.TelephoneNumbers
        lvEmails.Emails = Me.Contact.Emails
        lvWebsites.Websites = Me.Contact.Websites

        lvAddresses.Name = "lvAddresses"
        lvTelephoneNumbers.Name = "lvTelephoneNumbers"
        lvEmails.Name = "lvEmails"
        lvWebsites.Name = "lvWebsites"

        lvAddresses.Dock = DockStyle.Fill
        With Me.TabPageAddresses.Controls
            .Add(lvAddresses)
            .SetChildIndex(lvAddresses, 0)
            .SetChildIndex(ToolStripAddresses, 1)
            ToolStripAddresses.Visible = True
        End With

        lvTelephoneNumbers.Dock = DockStyle.Fill
        With Me.TabPageTelephoneNumbers.Controls
            .Add(lvTelephoneNumbers)
            .SetChildIndex(lvTelephoneNumbers, 0)
            .SetChildIndex(ToolStripTelephoneNumbers, 1)
            ToolStripTelephoneNumbers.Visible = True
        End With

        lvEmails.Dock = DockStyle.Fill
        With Me.TabPageEmails.Controls
            .Add(lvEmails)
            .SetChildIndex(lvEmails, 0)
            .SetChildIndex(ToolStripEmails, 1)
            ToolStripEmails.Visible = True
        End With

        lvWebsites.Dock = DockStyle.Fill
        With Me.TabPageWebsites.Controls
            .Add(lvWebsites)
            .SetChildIndex(lvWebsites, 0)
            .SetChildIndex(ToolStripWebsites, 1)
            ToolStripWebsites.Visible = True
        End With

        Me.tbFirstName.DataBindings.Add(New Binding("Text", bindContact, "FirstName"))
        Me.tbLastName.DataBindings.Add(New Binding("Text", bindContact, "LastName"))
        Me.tbJobTitle.DataBindings.Add(New Binding("Text", bindContact, "JobTitle"))
        With Me.cmbTitle
            .AutoCompleteMode = AutoCompleteMode.Suggest
            .DropDownStyle = ComboBoxStyle.DropDown
            .Items.AddRange(LIST_ContactTitles)
            .DataBindings.Add(New Binding("Text", bindContact, "Title"))
        End With
        With Me.cmbGender
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.AddRange(LIST_Gender)
            .DataBindings.Add(New Binding("SelectedIndex", bindContact, "Gender"))
        End With
        Me.tbRole.DataBindings.Add(New Binding("Text", bindContact, "Role"))
        Me.tbSkypeAccount.DataBindings.Add(New Binding("Text", bindContact, "SkypeAccount"))
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ToolStripButtonNewAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewAddress.Click
        lvAddresses.NewItem()
    End Sub

    Private Sub ToolStripButtonEditAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditAddress.Click
        lvAddresses.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveAddress.Click
        lvAddresses.RemoveItem()
    End Sub

    Private Sub ToolStripButtonSendLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonSendLetter.Click
        Dim selOrganisation As Organisation = CurrentLogFrame.GetOrganisationByGuid(Me.Contact.ParentOrganisationGuid)
        lvAddresses.SendLetter(Me.Contact, selOrganisation)
    End Sub

    Private Sub ToolStripButtonNewTelephoneNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewTelephoneNumber.Click
        lvTelephoneNumbers.NewItem()
    End Sub

    Private Sub ToolStripButtonEditTelephoneNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditTelephoneNumber.Click
        lvTelephoneNumbers.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveTelephoneNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveTelephoneNumber.Click
        lvTelephoneNumbers.RemoveItem()
    End Sub

    Private Sub ToolStripButtonNewEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewEmail.Click
        lvEmails.NewItem()
    End Sub

    Private Sub ToolStripButtonEditEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditEmail.Click
        lvEmails.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveEmail.Click
        lvEmails.RemoveItem()
    End Sub

    Private Sub ToolStripButtonSendEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonSendEmail.Click
        lvEmails.SendEmail()
    End Sub

    Private Sub ToolStripButtonNewWebsite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewWebsite.Click
        lvWebsites.NewItem()
    End Sub

    Private Sub ToolStripButtonEditWebsite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditWebsite.Click
        lvWebsites.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveWebsite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveWebsite.Click
        lvWebsites.RemoveItem()
    End Sub

    Private Sub ToolStripButtonVisitWebsite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonVisitWebsite.Click
        lvWebsites.OpenWebsite()
    End Sub

    Private Sub btnSkypeCall_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSkypeCall.Click
        Dim strSkypeAccount As String = Me.Contact.SkypeAccount

        If String.IsNullOrEmpty(strSkypeAccount) = False Then
            Using objSkype As New SkypeIO
                objSkype.MakeCall(strSkypeAccount)
            End Using
        End If
    End Sub

    Private Sub DialogContact_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        With frmParent
            If e.Control And e.KeyCode = Keys.C Then
                '<Ctrl><C> Copy item or text
                If CurrentClipboardType = DataGridViewClipboard.ContentTypes.Text Then
                    .CopyText()
                Else
                    .CopyItem()
                End If
            ElseIf e.Control And e.KeyCode = Keys.V Then
                '<Ctrl><V> Paste item or text
                If CurrentClipboardType = DataGridViewClipboard.ContentTypes.Text Then
                    e.SuppressKeyPress = True
                    .PasteText(0)
                Else
                    .PasteItem()
                End If
            ElseIf e.Control And e.KeyCode = Keys.X Then
                '<Ctrl><X> Cut item or text
                If CurrentClipboardType = DataGridViewClipboard.ContentTypes.Text Then
                    .CutText()
                Else
                    .CutItem()
                End If
            ElseIf e.KeyCode = Keys.Delete Then
                e.SuppressKeyPress = True
                .RemoveItem()
            End If
        End With
    End Sub
End Class
