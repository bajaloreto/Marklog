<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucOrganisationInfo
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucOrganisationInfo))
        Me.TabPageContacts = New System.Windows.Forms.TabPage()
        Me.ToolStripContacts = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewContact = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditContact = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveContact = New System.Windows.Forms.ToolStripButton()
        Me.TabPageWebsites = New System.Windows.Forms.TabPage()
        Me.ToolStripWebsites = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewWebsite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditWebsite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveWebsite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonVisitWebsite = New System.Windows.Forms.ToolStripButton()
        Me.cmbType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblSwift = New System.Windows.Forms.Label()
        Me.lblBankAccount = New System.Windows.Forms.Label()
        Me.tbSwift = New FaciliDev.LogFramer.TextBoxLF()
        Me.tbAbreviation = New FaciliDev.LogFramer.TextBoxLF()
        Me.TabPageBank = New System.Windows.Forms.TabPage()
        Me.tbBankAccount = New FaciliDev.LogFramer.TextBoxLF()
        Me.TabControlDetails = New System.Windows.Forms.TabControl()
        Me.TabPageAddresses = New System.Windows.Forms.TabPage()
        Me.ToolStripAddresses = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewAddress = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditAddress = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveAddress = New System.Windows.Forms.ToolStripButton()
        Me.TabPageTelephoneNumbers = New System.Windows.Forms.TabPage()
        Me.ToolStripTelephoneNumbers = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewTelephoneNumber = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditTelephoneNumber = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveTelephoneNumber = New System.Windows.Forms.ToolStripButton()
        Me.TabPageEmails = New System.Windows.Forms.TabPage()
        Me.ToolStripEmails = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewEmail = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditEmail = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveEmail = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonSendEmail = New System.Windows.Forms.ToolStripButton()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblAbreviation = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.tbName = New FaciliDev.LogFramer.TextBoxLF()
        Me.TabPageContacts.SuspendLayout()
        Me.ToolStripContacts.SuspendLayout()
        Me.TabPageWebsites.SuspendLayout()
        Me.ToolStripWebsites.SuspendLayout()
        Me.TabPageBank.SuspendLayout()
        Me.TabControlDetails.SuspendLayout()
        Me.TabPageAddresses.SuspendLayout()
        Me.ToolStripAddresses.SuspendLayout()
        Me.TabPageTelephoneNumbers.SuspendLayout()
        Me.ToolStripTelephoneNumbers.SuspendLayout()
        Me.TabPageEmails.SuspendLayout()
        Me.ToolStripEmails.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPageContacts
        '
        resources.ApplyResources(Me.TabPageContacts, "TabPageContacts")
        Me.TabPageContacts.Controls.Add(Me.ToolStripContacts)
        Me.TabPageContacts.Name = "TabPageContacts"
        Me.TabPageContacts.UseVisualStyleBackColor = True
        '
        'ToolStripContacts
        '
        resources.ApplyResources(Me.ToolStripContacts, "ToolStripContacts")
        Me.ToolStripContacts.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        Me.ToolStripContacts.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewContact, Me.ToolStripButtonEditContact, Me.ToolStripButtonRemoveContact})
        Me.ToolStripContacts.Name = "ToolStripContacts"
        '
        'ToolStripButtonNewContact
        '
        resources.ApplyResources(Me.ToolStripButtonNewContact, "ToolStripButtonNewContact")
        Me.ToolStripButtonNewContact.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonNewContact.Name = "ToolStripButtonNewContact"
        '
        'ToolStripButtonEditContact
        '
        resources.ApplyResources(Me.ToolStripButtonEditContact, "ToolStripButtonEditContact")
        Me.ToolStripButtonEditContact.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonEditContact.Name = "ToolStripButtonEditContact"
        '
        'ToolStripButtonRemoveContact
        '
        resources.ApplyResources(Me.ToolStripButtonRemoveContact, "ToolStripButtonRemoveContact")
        Me.ToolStripButtonRemoveContact.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonRemoveContact.Name = "ToolStripButtonRemoveContact"
        '
        'TabPageWebsites
        '
        resources.ApplyResources(Me.TabPageWebsites, "TabPageWebsites")
        Me.TabPageWebsites.Controls.Add(Me.ToolStripWebsites)
        Me.TabPageWebsites.Name = "TabPageWebsites"
        Me.TabPageWebsites.UseVisualStyleBackColor = True
        '
        'ToolStripWebsites
        '
        resources.ApplyResources(Me.ToolStripWebsites, "ToolStripWebsites")
        Me.ToolStripWebsites.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        Me.ToolStripWebsites.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewWebsite, Me.ToolStripButtonEditWebsite, Me.ToolStripButtonRemoveWebsite, Me.ToolStripSeparator1, Me.ToolStripButtonVisitWebsite})
        Me.ToolStripWebsites.Name = "ToolStripWebsites"
        '
        'ToolStripButtonNewWebsite
        '
        resources.ApplyResources(Me.ToolStripButtonNewWebsite, "ToolStripButtonNewWebsite")
        Me.ToolStripButtonNewWebsite.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonNewWebsite.Name = "ToolStripButtonNewWebsite"
        '
        'ToolStripButtonEditWebsite
        '
        resources.ApplyResources(Me.ToolStripButtonEditWebsite, "ToolStripButtonEditWebsite")
        Me.ToolStripButtonEditWebsite.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonEditWebsite.Name = "ToolStripButtonEditWebsite"
        '
        'ToolStripButtonRemoveWebsite
        '
        resources.ApplyResources(Me.ToolStripButtonRemoveWebsite, "ToolStripButtonRemoveWebsite")
        Me.ToolStripButtonRemoveWebsite.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonRemoveWebsite.Name = "ToolStripButtonRemoveWebsite"
        '
        'ToolStripSeparator1
        '
        resources.ApplyResources(Me.ToolStripSeparator1, "ToolStripSeparator1")
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        '
        'ToolStripButtonVisitWebsite
        '
        resources.ApplyResources(Me.ToolStripButtonVisitWebsite, "ToolStripButtonVisitWebsite")
        Me.ToolStripButtonVisitWebsite.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonVisitWebsite.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Detailweb
        Me.ToolStripButtonVisitWebsite.Name = "ToolStripButtonVisitWebsite"
        '
        'cmbType
        '
        resources.ApplyResources(Me.cmbType, "cmbType")
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Name = "cmbType"
        '
        'lblSwift
        '
        resources.ApplyResources(Me.lblSwift, "lblSwift")
        Me.lblSwift.Name = "lblSwift"
        '
        'lblBankAccount
        '
        resources.ApplyResources(Me.lblBankAccount, "lblBankAccount")
        Me.lblBankAccount.Name = "lblBankAccount"
        '
        'tbSwift
        '
        resources.ApplyResources(Me.tbSwift, "tbSwift")
        Me.tbSwift.Name = "tbSwift"
        '
        'tbAbreviation
        '
        resources.ApplyResources(Me.tbAbreviation, "tbAbreviation")
        Me.tbAbreviation.Name = "tbAbreviation"
        '
        'TabPageBank
        '
        resources.ApplyResources(Me.TabPageBank, "TabPageBank")
        Me.TabPageBank.Controls.Add(Me.lblSwift)
        Me.TabPageBank.Controls.Add(Me.lblBankAccount)
        Me.TabPageBank.Controls.Add(Me.tbSwift)
        Me.TabPageBank.Controls.Add(Me.tbBankAccount)
        Me.TabPageBank.Name = "TabPageBank"
        Me.TabPageBank.UseVisualStyleBackColor = True
        '
        'tbBankAccount
        '
        resources.ApplyResources(Me.tbBankAccount, "tbBankAccount")
        Me.tbBankAccount.Name = "tbBankAccount"
        '
        'TabControlDetails
        '
        resources.ApplyResources(Me.TabControlDetails, "TabControlDetails")
        Me.TabControlDetails.Controls.Add(Me.TabPageAddresses)
        Me.TabControlDetails.Controls.Add(Me.TabPageTelephoneNumbers)
        Me.TabControlDetails.Controls.Add(Me.TabPageEmails)
        Me.TabControlDetails.Controls.Add(Me.TabPageWebsites)
        Me.TabControlDetails.Controls.Add(Me.TabPageContacts)
        Me.TabControlDetails.Controls.Add(Me.TabPageBank)
        Me.TabControlDetails.Name = "TabControlDetails"
        Me.TabControlDetails.SelectedIndex = 0
        '
        'TabPageAddresses
        '
        resources.ApplyResources(Me.TabPageAddresses, "TabPageAddresses")
        Me.TabPageAddresses.Controls.Add(Me.ToolStripAddresses)
        Me.TabPageAddresses.Name = "TabPageAddresses"
        Me.TabPageAddresses.UseVisualStyleBackColor = True
        '
        'ToolStripAddresses
        '
        resources.ApplyResources(Me.ToolStripAddresses, "ToolStripAddresses")
        Me.ToolStripAddresses.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        Me.ToolStripAddresses.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewAddress, Me.ToolStripButtonEditAddress, Me.ToolStripButtonRemoveAddress})
        Me.ToolStripAddresses.Name = "ToolStripAddresses"
        '
        'ToolStripButtonNewAddress
        '
        resources.ApplyResources(Me.ToolStripButtonNewAddress, "ToolStripButtonNewAddress")
        Me.ToolStripButtonNewAddress.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonNewAddress.Name = "ToolStripButtonNewAddress"
        '
        'ToolStripButtonEditAddress
        '
        resources.ApplyResources(Me.ToolStripButtonEditAddress, "ToolStripButtonEditAddress")
        Me.ToolStripButtonEditAddress.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonEditAddress.Name = "ToolStripButtonEditAddress"
        '
        'ToolStripButtonRemoveAddress
        '
        resources.ApplyResources(Me.ToolStripButtonRemoveAddress, "ToolStripButtonRemoveAddress")
        Me.ToolStripButtonRemoveAddress.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonRemoveAddress.Name = "ToolStripButtonRemoveAddress"
        '
        'TabPageTelephoneNumbers
        '
        resources.ApplyResources(Me.TabPageTelephoneNumbers, "TabPageTelephoneNumbers")
        Me.TabPageTelephoneNumbers.Controls.Add(Me.ToolStripTelephoneNumbers)
        Me.TabPageTelephoneNumbers.Name = "TabPageTelephoneNumbers"
        Me.TabPageTelephoneNumbers.UseVisualStyleBackColor = True
        '
        'ToolStripTelephoneNumbers
        '
        resources.ApplyResources(Me.ToolStripTelephoneNumbers, "ToolStripTelephoneNumbers")
        Me.ToolStripTelephoneNumbers.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        Me.ToolStripTelephoneNumbers.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewTelephoneNumber, Me.ToolStripButtonEditTelephoneNumber, Me.ToolStripButtonRemoveTelephoneNumber})
        Me.ToolStripTelephoneNumbers.Name = "ToolStripTelephoneNumbers"
        '
        'ToolStripButtonNewTelephoneNumber
        '
        resources.ApplyResources(Me.ToolStripButtonNewTelephoneNumber, "ToolStripButtonNewTelephoneNumber")
        Me.ToolStripButtonNewTelephoneNumber.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonNewTelephoneNumber.Name = "ToolStripButtonNewTelephoneNumber"
        '
        'ToolStripButtonEditTelephoneNumber
        '
        resources.ApplyResources(Me.ToolStripButtonEditTelephoneNumber, "ToolStripButtonEditTelephoneNumber")
        Me.ToolStripButtonEditTelephoneNumber.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonEditTelephoneNumber.Name = "ToolStripButtonEditTelephoneNumber"
        '
        'ToolStripButtonRemoveTelephoneNumber
        '
        resources.ApplyResources(Me.ToolStripButtonRemoveTelephoneNumber, "ToolStripButtonRemoveTelephoneNumber")
        Me.ToolStripButtonRemoveTelephoneNumber.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonRemoveTelephoneNumber.Name = "ToolStripButtonRemoveTelephoneNumber"
        '
        'TabPageEmails
        '
        resources.ApplyResources(Me.TabPageEmails, "TabPageEmails")
        Me.TabPageEmails.Controls.Add(Me.ToolStripEmails)
        Me.TabPageEmails.Name = "TabPageEmails"
        Me.TabPageEmails.UseVisualStyleBackColor = True
        '
        'ToolStripEmails
        '
        resources.ApplyResources(Me.ToolStripEmails, "ToolStripEmails")
        Me.ToolStripEmails.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        Me.ToolStripEmails.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewEmail, Me.ToolStripButtonEditEmail, Me.ToolStripButtonRemoveEmail, Me.ToolStripSeparator2, Me.ToolStripButtonSendEmail})
        Me.ToolStripEmails.Name = "ToolStripEmails"
        '
        'ToolStripButtonNewEmail
        '
        resources.ApplyResources(Me.ToolStripButtonNewEmail, "ToolStripButtonNewEmail")
        Me.ToolStripButtonNewEmail.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonNewEmail.Name = "ToolStripButtonNewEmail"
        '
        'ToolStripButtonEditEmail
        '
        resources.ApplyResources(Me.ToolStripButtonEditEmail, "ToolStripButtonEditEmail")
        Me.ToolStripButtonEditEmail.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonEditEmail.Name = "ToolStripButtonEditEmail"
        '
        'ToolStripButtonRemoveEmail
        '
        resources.ApplyResources(Me.ToolStripButtonRemoveEmail, "ToolStripButtonRemoveEmail")
        Me.ToolStripButtonRemoveEmail.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonRemoveEmail.Name = "ToolStripButtonRemoveEmail"
        '
        'ToolStripSeparator2
        '
        resources.ApplyResources(Me.ToolStripSeparator2, "ToolStripSeparator2")
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        '
        'ToolStripButtonSendEmail
        '
        resources.ApplyResources(Me.ToolStripButtonSendEmail, "ToolStripButtonSendEmail")
        Me.ToolStripButtonSendEmail.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonSendEmail.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.DetailMail
        Me.ToolStripButtonSendEmail.Name = "ToolStripButtonSendEmail"
        '
        'lblType
        '
        resources.ApplyResources(Me.lblType, "lblType")
        Me.lblType.Name = "lblType"
        '
        'lblAbreviation
        '
        resources.ApplyResources(Me.lblAbreviation, "lblAbreviation")
        Me.lblAbreviation.Name = "lblAbreviation"
        '
        'lblName
        '
        resources.ApplyResources(Me.lblName, "lblName")
        Me.lblName.Name = "lblName"
        '
        'tbName
        '
        resources.ApplyResources(Me.tbName, "tbName")
        Me.tbName.Name = "tbName"
        '
        'ucOrganisationInfo
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.tbAbreviation)
        Me.Controls.Add(Me.TabControlDetails)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.lblAbreviation)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.tbName)
        Me.Name = "ucOrganisationInfo"
        Me.TabPageContacts.ResumeLayout(False)
        Me.TabPageContacts.PerformLayout()
        Me.ToolStripContacts.ResumeLayout(False)
        Me.ToolStripContacts.PerformLayout()
        Me.TabPageWebsites.ResumeLayout(False)
        Me.TabPageWebsites.PerformLayout()
        Me.ToolStripWebsites.ResumeLayout(False)
        Me.ToolStripWebsites.PerformLayout()
        Me.TabPageBank.ResumeLayout(False)
        Me.TabPageBank.PerformLayout()
        Me.TabControlDetails.ResumeLayout(False)
        Me.TabPageAddresses.ResumeLayout(False)
        Me.TabPageAddresses.PerformLayout()
        Me.ToolStripAddresses.ResumeLayout(False)
        Me.ToolStripAddresses.PerformLayout()
        Me.TabPageTelephoneNumbers.ResumeLayout(False)
        Me.TabPageTelephoneNumbers.PerformLayout()
        Me.ToolStripTelephoneNumbers.ResumeLayout(False)
        Me.ToolStripTelephoneNumbers.PerformLayout()
        Me.TabPageEmails.ResumeLayout(False)
        Me.TabPageEmails.PerformLayout()
        Me.ToolStripEmails.ResumeLayout(False)
        Me.ToolStripEmails.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStripButtonVisitWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents TabPageContacts As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripContacts As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewContact As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditContact As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveContact As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButtonRemoveWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents TabPageWebsites As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripWebsites As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents cmbType As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblSwift As System.Windows.Forms.Label
    Friend WithEvents lblBankAccount As System.Windows.Forms.Label
    Friend WithEvents tbSwift As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents tbAbreviation As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents TabPageBank As System.Windows.Forms.TabPage
    Friend WithEvents tbBankAccount As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents ToolStripButtonSendEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TabControlDetails As System.Windows.Forms.TabControl
    Friend WithEvents TabPageAddresses As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripAddresses As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewAddress As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditAddress As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveAddress As System.Windows.Forms.ToolStripButton
    Friend WithEvents TabPageTelephoneNumbers As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripTelephoneNumbers As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewTelephoneNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditTelephoneNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveTelephoneNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents TabPageEmails As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripEmails As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblAbreviation As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents tbName As FaciliDev.LogFramer.TextBoxLF

End Class
