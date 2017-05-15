<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogContact
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogContact))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.tbFirstName = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblFirstName = New System.Windows.Forms.Label()
        Me.lblLastName = New System.Windows.Forms.Label()
        Me.tbLastName = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblRole = New System.Windows.Forms.Label()
        Me.TabControlDetails = New System.Windows.Forms.TabControl()
        Me.TabPageAddresses = New System.Windows.Forms.TabPage()
        Me.ToolStripAddresses = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewAddress = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditAddress = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveAddress = New System.Windows.Forms.ToolStripButton()
        Me.TabPageSkype = New System.Windows.Forms.TabPage()
        Me.btnSkypeCall = New System.Windows.Forms.Button()
        Me.lblSkypeAccount = New System.Windows.Forms.Label()
        Me.tbSkypeAccount = New FaciliDev.LogFramer.TextBoxLF()
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
        Me.TabPageWebsites = New System.Windows.Forms.TabPage()
        Me.ToolStripWebsites = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonNewWebsite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditWebsite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonRemoveWebsite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonVisitWebsite = New System.Windows.Forms.ToolStripButton()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblJobTitle = New System.Windows.Forms.Label()
        Me.tbJobTitle = New FaciliDev.LogFramer.TextBoxLF()
        Me.cmbTitle = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbGender = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblGender = New System.Windows.Forms.Label()
        Me.tbRole = New FaciliDev.LogFramer.TextBoxLF()
        Me.ToolStripContact = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonPrint = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonSendLetter = New System.Windows.Forms.ToolStripButton()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControlDetails.SuspendLayout()
        Me.TabPageAddresses.SuspendLayout()
        Me.ToolStripAddresses.SuspendLayout()
        Me.TabPageSkype.SuspendLayout()
        Me.TabPageTelephoneNumbers.SuspendLayout()
        Me.ToolStripTelephoneNumbers.SuspendLayout()
        Me.TabPageEmails.SuspendLayout()
        Me.ToolStripEmails.SuspendLayout()
        Me.TabPageWebsites.SuspendLayout()
        Me.ToolStripWebsites.SuspendLayout()
        Me.ToolStripContact.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'OK_Button
        '
        resources.ApplyResources(Me.OK_Button, "OK_Button")
        Me.OK_Button.Name = "OK_Button"
        '
        'Cancel_Button
        '
        resources.ApplyResources(Me.Cancel_Button, "Cancel_Button")
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Name = "Cancel_Button"
        '
        'tbFirstName
        '
        resources.ApplyResources(Me.tbFirstName, "tbFirstName")
        Me.tbFirstName.Name = "tbFirstName"
        '
        'lblFirstName
        '
        resources.ApplyResources(Me.lblFirstName, "lblFirstName")
        Me.lblFirstName.Name = "lblFirstName"
        '
        'lblLastName
        '
        resources.ApplyResources(Me.lblLastName, "lblLastName")
        Me.lblLastName.Name = "lblLastName"
        '
        'tbLastName
        '
        resources.ApplyResources(Me.tbLastName, "tbLastName")
        Me.tbLastName.Name = "tbLastName"
        '
        'lblRole
        '
        resources.ApplyResources(Me.lblRole, "lblRole")
        Me.lblRole.Name = "lblRole"
        '
        'TabControlDetails
        '
        resources.ApplyResources(Me.TabControlDetails, "TabControlDetails")
        Me.TabControlDetails.Controls.Add(Me.TabPageAddresses)
        Me.TabControlDetails.Controls.Add(Me.TabPageSkype)
        Me.TabControlDetails.Controls.Add(Me.TabPageTelephoneNumbers)
        Me.TabControlDetails.Controls.Add(Me.TabPageEmails)
        Me.TabControlDetails.Controls.Add(Me.TabPageWebsites)
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
        Me.ToolStripAddresses.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewAddress, Me.ToolStripButtonEditAddress, Me.ToolStripButtonRemoveAddress, Me.ToolStripSeparator3, Me.ToolStripButtonSendLetter})
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
        'TabPageSkype
        '
        resources.ApplyResources(Me.TabPageSkype, "TabPageSkype")
        Me.TabPageSkype.Controls.Add(Me.btnSkypeCall)
        Me.TabPageSkype.Controls.Add(Me.lblSkypeAccount)
        Me.TabPageSkype.Controls.Add(Me.tbSkypeAccount)
        Me.TabPageSkype.Name = "TabPageSkype"
        Me.TabPageSkype.UseVisualStyleBackColor = True
        '
        'btnSkypeCall
        '
        resources.ApplyResources(Me.btnSkypeCall, "btnSkypeCall")
        Me.btnSkypeCall.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.CollaborateSkype_64px
        Me.btnSkypeCall.Name = "btnSkypeCall"
        Me.btnSkypeCall.UseVisualStyleBackColor = True
        '
        'lblSkypeAccount
        '
        resources.ApplyResources(Me.lblSkypeAccount, "lblSkypeAccount")
        Me.lblSkypeAccount.Name = "lblSkypeAccount"
        '
        'tbSkypeAccount
        '
        resources.ApplyResources(Me.tbSkypeAccount, "tbSkypeAccount")
        Me.tbSkypeAccount.Name = "tbSkypeAccount"
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
        'lblTitle
        '
        resources.ApplyResources(Me.lblTitle, "lblTitle")
        Me.lblTitle.Name = "lblTitle"
        '
        'lblJobTitle
        '
        resources.ApplyResources(Me.lblJobTitle, "lblJobTitle")
        Me.lblJobTitle.Name = "lblJobTitle"
        '
        'tbJobTitle
        '
        resources.ApplyResources(Me.tbJobTitle, "tbJobTitle")
        Me.tbJobTitle.Name = "tbJobTitle"
        '
        'cmbTitle
        '
        resources.ApplyResources(Me.cmbTitle, "cmbTitle")
        Me.cmbTitle.FormattingEnabled = True
        Me.cmbTitle.Name = "cmbTitle"
        '
        'cmbGender
        '
        resources.ApplyResources(Me.cmbGender, "cmbGender")
        Me.cmbGender.FormattingEnabled = True
        Me.cmbGender.Name = "cmbGender"
        '
        'lblGender
        '
        resources.ApplyResources(Me.lblGender, "lblGender")
        Me.lblGender.Name = "lblGender"
        '
        'tbRole
        '
        resources.ApplyResources(Me.tbRole, "tbRole")
        Me.tbRole.Name = "tbRole"
        '
        'ToolStripContact
        '
        resources.ApplyResources(Me.ToolStripContact, "ToolStripContact")
        Me.ToolStripContact.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        Me.ToolStripContact.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonPrint})
        Me.ToolStripContact.Name = "ToolStripContact"
        '
        'ToolStripButtonPrint
        '
        resources.ApplyResources(Me.ToolStripButtonPrint, "ToolStripButtonPrint")
        Me.ToolStripButtonPrint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonPrint.Name = "ToolStripButtonPrint"
        '
        'ToolStripSeparator3
        '
        resources.ApplyResources(Me.ToolStripSeparator3, "ToolStripSeparator3")
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        '
        'ToolStripButtonSendLetter
        '
        resources.ApplyResources(Me.ToolStripButtonSendLetter, "ToolStripButtonSendLetter")
        Me.ToolStripButtonSendLetter.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonSendLetter.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.CollaborateLetter
        Me.ToolStripButtonSendLetter.Margin = New System.Windows.Forms.Padding(0)
        Me.ToolStripButtonSendLetter.Name = "ToolStripButtonSendLetter"
        '
        'DialogContact
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.ToolStripContact)
        Me.Controls.Add(Me.tbRole)
        Me.Controls.Add(Me.cmbGender)
        Me.Controls.Add(Me.lblGender)
        Me.Controls.Add(Me.cmbTitle)
        Me.Controls.Add(Me.lblJobTitle)
        Me.Controls.Add(Me.tbJobTitle)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.TabControlDetails)
        Me.Controls.Add(Me.lblRole)
        Me.Controls.Add(Me.lblLastName)
        Me.Controls.Add(Me.tbLastName)
        Me.Controls.Add(Me.lblFirstName)
        Me.Controls.Add(Me.tbFirstName)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogContact"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControlDetails.ResumeLayout(False)
        Me.TabPageAddresses.ResumeLayout(False)
        Me.TabPageAddresses.PerformLayout()
        Me.ToolStripAddresses.ResumeLayout(False)
        Me.ToolStripAddresses.PerformLayout()
        Me.TabPageSkype.ResumeLayout(False)
        Me.TabPageSkype.PerformLayout()
        Me.TabPageTelephoneNumbers.ResumeLayout(False)
        Me.TabPageTelephoneNumbers.PerformLayout()
        Me.ToolStripTelephoneNumbers.ResumeLayout(False)
        Me.ToolStripTelephoneNumbers.PerformLayout()
        Me.TabPageEmails.ResumeLayout(False)
        Me.TabPageEmails.PerformLayout()
        Me.ToolStripEmails.ResumeLayout(False)
        Me.ToolStripEmails.PerformLayout()
        Me.TabPageWebsites.ResumeLayout(False)
        Me.TabPageWebsites.PerformLayout()
        Me.ToolStripWebsites.ResumeLayout(False)
        Me.ToolStripWebsites.PerformLayout()
        Me.ToolStripContact.ResumeLayout(False)
        Me.ToolStripContact.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents tbFirstName As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents tbLastName As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblRole As System.Windows.Forms.Label
    Friend WithEvents TabControlDetails As System.Windows.Forms.TabControl
    Friend WithEvents TabPageAddresses As System.Windows.Forms.TabPage
    Friend WithEvents TabPageTelephoneNumbers As System.Windows.Forms.TabPage
    Friend WithEvents TabPageEmails As System.Windows.Forms.TabPage
    Friend WithEvents TabPageWebsites As System.Windows.Forms.TabPage
    Friend WithEvents ToolStripAddresses As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewAddress As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditAddress As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveAddress As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripTelephoneNumbers As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewTelephoneNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditTelephoneNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveTelephoneNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripEmails As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButtonSendEmail As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripWebsites As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonNewWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonRemoveWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButtonVisitWebsite As System.Windows.Forms.ToolStripButton
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents lblJobTitle As System.Windows.Forms.Label
    Friend WithEvents tbJobTitle As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents cmbTitle As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents cmbGender As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblGender As System.Windows.Forms.Label
    Friend WithEvents tbRole As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents ToolStripContact As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonPrint As System.Windows.Forms.ToolStripButton
    Friend WithEvents TabPageSkype As System.Windows.Forms.TabPage
    Friend WithEvents btnSkypeCall As System.Windows.Forms.Button
    Friend WithEvents lblSkypeAccount As System.Windows.Forms.Label
    Friend WithEvents tbSkypeAccount As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButtonSendLetter As System.Windows.Forms.ToolStripButton

End Class
