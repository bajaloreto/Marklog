Imports System.Windows.Forms

Public Class DialogProjectPartner
    Friend WithEvents lvAddresses As New ListViewAddresses
    Friend WithEvents lvTelephoneNumbers As New ListViewTelephoneNumbers
    Friend WithEvents lvEmails As New ListViewEmails
    Friend WithEvents lvWebsites As New ListViewWebsites
    Friend WithEvents lvContacts As New ListViewContacts

    Private objProjectPartner As ProjectPartner
    Private bindProjectPartner As New BindingSource

    Public Property ProjectPartner As ProjectPartner
        Get
            Return objProjectPartner
        End Get
        Set(ByVal value As ProjectPartner)
            objProjectPartner = value
        End Set
    End Property

    Public Sub New(ByVal projectpartner As ProjectPartner)
        InitializeComponent()
        Me.ProjectPartner = projectpartner
        Me.KeyPreview = True

        If Me.ProjectPartner IsNot Nothing Then
            bindProjectPartner.DataSource = Me.ProjectPartner

            lvAddresses.Addresses = Me.ProjectPartner.Organisation.Addresses
            lvTelephoneNumbers.TelephoneNumbers = Me.ProjectPartner.Organisation.TelephoneNumbers
            lvEmails.Emails = Me.ProjectPartner.Organisation.Emails
            lvWebsites.Websites = Me.ProjectPartner.Organisation.Websites
            lvContacts.Reload(Me.ProjectPartner.Organisation.Contacts, Me.ProjectPartner.Organisation.Guid)

            lvAddresses.Name = "lvAddresses"
            lvTelephoneNumbers.Name = "lvTelephoneNumbers"
            lvEmails.Name = "lvEmails"
            lvWebsites.Name = "lvWebsites"
            lvContacts.Name = "lvContacts"

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

            Me.tbName.DataBindings.Add(New Binding("Text", bindProjectPartner, "Organisation.Name"))
            Me.tbAcronym.DataBindings.Add(New Binding("Text", bindProjectPartner, "Organisation.Acronym"))

            With Me.cmbType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_OrganisationTypes)
                .DataBindings.Add(New Binding("SelectedIndex", bindProjectPartner, "Organisation.Type"))
            End With

            With Me.cmbRole
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_ProjectPartnerRoleNames)
                .DataBindings.Add(New Binding("SelectedIndex", bindProjectPartner, "Role"))
            End With

            Me.tbBankAccount.DataBindings.Add(New Binding("Text", bindProjectPartner, "Organisation.BankAccount"))
            Me.tbSwift.DataBindings.Add(New Binding("Text", bindProjectPartner, "Organisation.Swift"))
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

    Private Sub ToolStripButtonSendLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonSendLetter.Click
        lvAddresses.SendLetter(Nothing, Me.ProjectPartner.Organisation)
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

    Private Sub ToolStripButtonNewContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNewContact.Click
        lvContacts.NewItem()
    End Sub

    Private Sub ToolStripButtonEditContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonEditContact.Click
        lvContacts.EditItem()
    End Sub

    Private Sub ToolStripButtonRemoveContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRemoveContact.Click
        lvContacts.RemoveItem()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DialogProjectPartner_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
