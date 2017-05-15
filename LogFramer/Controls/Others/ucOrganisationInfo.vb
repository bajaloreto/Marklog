Public Class ucOrganisationInfo
    Private objOrganisation As Organisation
    Private bindOrganisation As New BindingSource

    Friend WithEvents lvAddresses As New ListViewAddresses
    Friend WithEvents lvTelephoneNumbers As New ListViewTelephoneNumbers
    Friend WithEvents lvEmails As New ListViewEmails
    Friend WithEvents lvWebsites As New ListViewWebsites
    Friend WithEvents lvContacts As New ListViewContacts

    Public Property Organisation() As Organisation
        Get
            Return objOrganisation
        End Get
        Set(ByVal value As Organisation)
            objOrganisation = value
            LoadOrganisationInfo()
        End Set
    End Property

    Private Sub LoadOrganisationInfo()
        'bindOrganisation.Clear()
        If Me.Organisation IsNot Nothing Then
            bindOrganisation.DataSource = Me.Organisation

            lvAddresses.Addresses = Me.Organisation.Addresses
            lvTelephoneNumbers.TelephoneNumbers = Me.Organisation.TelephoneNumbers
            lvEmails.Emails = Me.Organisation.Emails
            lvWebsites.Websites = Me.Organisation.Websites
            lvContacts.Reload(Me.Organisation.Contacts, Me.Organisation.Guid)

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

            lvContacts.Dock = DockStyle.Fill
            With Me.TabPageContacts.Controls
                .Add(lvContacts)
                .SetChildIndex(lvContacts, 0)
                .SetChildIndex(ToolStripContacts, 1)
                ToolStripContacts.Visible = True
            End With

            Me.tbName.DataBindings.Add(New Binding("Text", bindOrganisation, "Name"))
            Me.tbAbreviation.DataBindings.Add(New Binding("Text", bindOrganisation, "Abreviation"))
            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_OrganisationTypes)
                .DataBindings.Add(New Binding("SelectedIndex", bindOrganisation, "Type"))
            End With

            Me.tbBankAccount.DataBindings.Add(New Binding("Text", bindOrganisation, "BankAccount"))
            Me.tbSwift.DataBindings.Add(New Binding("Text", bindOrganisation, "Swift"))
        End If
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

    Private Sub ToolStripButtonNewTelephoneNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewTelephoneNumber.Click
        lvTelephoneNumbers.NewItem()
    End Sub

    Private Sub ToolStripButtonEditTelephoneNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditTelephoneNumber.Click
        lvTelephoneNumbers.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveTelephoneNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveTelephoneNumber.Click
        lvTelephoneNumbers.NewItem()
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

    Private Sub ToolStripButtonNewContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewContact.Click
        lvContacts.NewItem()
    End Sub

    Private Sub ToolStripButtonEditContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditContact.Click
        lvContacts.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveContact.Click
        lvContacts.RemoveItem()
    End Sub
End Class
