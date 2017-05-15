Imports System.Windows.Forms
Imports System.Reflection

Public Class DialogFindContact
    Private StartIndex, StartIndexContacts As Integer
    Private StartIndexAddresses, StartIndexTelephoneNumbers, StartIndexEmails, StartIndexWebsites As Integer
    Private PropInfoIndex, PropInfoIndexContacts As Integer
    Private PropInfoIndexAddresses, PropInfoIndexTelephoneNumbers, PropInfoIndexEmails, PropInfoIndexWebsites As Integer
    Private ProjectPartnerIndex As Integer, ContactIndex As Integer
    Private AddressIndex As Integer, TelephoneNumberIndex, EmailIndex, WebsiteIndex As Integer
    Private boolOrganisationChecked As Boolean, boolContactChecked As Boolean
    Private boolOrganisationChildrenChecked As Boolean
    Private boolAddressChecked As Boolean, boolTelephoneNumberChecked, boolEmailChecked, boolWebsiteChecked As Boolean

    Private Enum SearchOptions
        SearchAll = 0
        SearchOrganisations = 1
        SearchContacts = 2
    End Enum

    Private ReadOnly Property SearchOption As Integer
        Get
            Return My.Settings.setSearchContacts
        End Get
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        Me.TopMost = True
        lblMessage.Text = ""

        Select Case My.Settings.setSearchContacts
            Case SearchOptions.SearchAll
                RadioButtonAll.Checked = True
            Case SearchOptions.SearchOrganisations
                RadioButtonOrganisations.Checked = True
            Case SearchOptions.SearchContacts
                RadioButtonContacts.Checked = True
        End Select

        Me.tbFind.Focus()
    End Sub

#Region "Events"
    Private Sub btnFindNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindNext.Click
        Dim strFind As String = Trim(tbFind.Text)

        If strFind = "" Then
            lblMessage.Text = "Please indicate which words to find"
            Exit Sub
        End If

        FindText(strFind)
    End Sub

    Private Sub btnFindFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindFirst.Click
        Dim strFind As String = Trim(tbFind.Text)

        If strFind = "" Then
            lblMessage.Text = "Please indicate which words to find"
            Exit Sub
        End If

        StartIndex = 0
        StartIndexContacts = 0
        StartIndexAddresses = 0
        StartIndexTelephoneNumbers = 0
        StartIndexEmails = 0
        StartIndexWebsites = 0
        PropInfoIndex = 0
        PropInfoIndexContacts = 0
        PropInfoIndexAddresses = 0
        PropInfoIndexTelephoneNumbers = 0
        PropInfoIndexEmails = 0
        PropInfoIndexWebsites = 0
        ProjectPartnerIndex = 0
        ContactIndex = 0
        AddressIndex = 0
        TelephoneNumberIndex = 0
        EmailIndex = 0
        WebsiteIndex = 0
        boolOrganisationChecked = False
        boolContactChecked = False
        boolOrganisationChildrenChecked = False
        boolAddressChecked = False
        boolTelephoneNumberChecked = 0
        boolEmailChecked = 0
        boolWebsiteChecked = False

        FindText(strFind)
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub RadioButtonAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonAll.CheckedChanged
        If RadioButtonAll.Checked = True Then My.Settings.setSearchContacts = SearchOptions.SearchAll
    End Sub

    Private Sub RadioButtonOrganisations_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonOrganisations.CheckedChanged
        If RadioButtonOrganisations.Checked = True Then My.Settings.setSearchContacts = SearchOptions.SearchOrganisations
    End Sub

    Private Sub RadioButtonContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonContacts.CheckedChanged
        If RadioButtonContacts.Checked = True Then My.Settings.setSearchContacts = SearchOptions.SearchContacts
    End Sub
#End Region

#Region "General methods"
    Private Sub FindText(ByVal strFind As String, Optional ByVal strReplace As String = "")
        Dim strCheckedAll

        Select Case UserLanguage
            Case "fr"
                strCheckedAll = "Toutes les organisations et contacts ont été cherché"
            Case Else
                strCheckedAll = "Searched all organisations and contacts"
        End Select

        lblMessage.Text = ""

        If strFind = "" Then Exit Sub

        If Me.chkMatchCase.Checked = False Then
            strFind = strFind.ToLower
        End If
        If Me.chkMatchWholeWord.Checked = True Then
            strFind &= " "
        End If

        FindTextInOrganisations(strFind)

        If ProjectPartnerIndex >= CurrentLogFrame.ProjectPartners.Count - 1 Then
            ProjectPartnerIndex = 0

            lblMessage.Text = strCheckedAll

        End If
    End Sub

    Private Function GetControl(ByVal ParentWindow As Object, ByVal strName As String, ByVal strText As String) As Control
        Dim selControl As Control = Nothing

        For Each selControl In ParentWindow.Controls
            If selControl.GetType = GetType(TabControl) Then
                selControl.Focus()
                For Each selPage As TabPage In CType(selControl, TabControl).TabPages
                    selPage.Focus()

                    selControl = GetControl(selPage, strName, strText)
                    If selControl IsNot Nothing Then Return selControl
                Next
            ElseIf selControl.GetType = GetType(SplitContainer) Then
                Dim objSplitContainer As SplitContainer = CType(selControl, SplitContainer)
                objSplitContainer.Focus()
                selControl = GetControl(objSplitContainer.Panel1, strName, strText)
                If selControl IsNot Nothing Then Return selControl
                selControl = GetControl(objSplitContainer.Panel2, strName, strText)
                If selControl IsNot Nothing Then Return selControl
            ElseIf selControl.GetType = GetType(TableLayoutPanel) Then
                selControl.Focus()
                selControl = GetControl(selControl, strName, strText)
                If selControl IsNot Nothing Then Return selControl
            ElseIf selControl.GetType = GetType(GroupBox) Then
                selControl.Focus()
                selControl = GetControl(selControl, strName, strText)
                If selControl IsNot Nothing Then Return selControl
            Else
                If selControl.GetType IsNot GetType(Label) And selControl.Name.Contains(strName) Then
                    Dim ControlText As String
                    If Me.chkMatchCase.Checked = False Then ControlText = selControl.Text.ToLower Else ControlText = selControl.Text
                    If ControlText = strText Then
                        Return selControl
                    End If

                End If
            End If
        Next
        Return Nothing
    End Function

    Private Sub SelectTextInControl(ByVal selControl As Control, ByVal strFind As String)
        selControl.Focus()
        If selControl.GetType = GetType(TextBoxLF) Then
            Dim objTextBoxLF As TextBoxLF = CType(selControl, TextBoxLF)
            objTextBoxLF.Select(StartIndex, strFind.Length)

        ElseIf selControl.GetType = GetType(NumericTextBoxLF) Then
            selControl.Select()
        Else
            If selControl.CanSelect Then selControl.Select()

        End If
    End Sub

    Private Function IsWritableProperty(ByVal selPropInfo As PropertyInfo) As Boolean

        If selPropInfo.MemberType = MemberTypes.Property And selPropInfo.CanWrite = True And _
                    (selPropInfo.PropertyType = GetType(String) Or _
                    selPropInfo.PropertyType = GetType(Integer) Or _
                    selPropInfo.PropertyType = GetType(Single) Or _
                    selPropInfo.PropertyType = GetType(Date)) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub ShowPartnerDetailsWindow()

        'CurrentProjectForm.dgvLogframe.CurrentCell = CurrentProjectForm.dgvLogframe(1, 1)
        'If My.Settings.setShowDetailsLogframe = False Then
        '    frmLogFrame.ShowDetails(False)
        'ElseIf frmLogFrame.CurrentDetail.GetType IsNot GetType(DetailGoal) Then
        '    frmLogFrame.SetTypeOfDetailWindow()
        'End If

        'Dim selDetail As DetailGoal = CType(frmLogFrame.CurrentDetail, DetailGoal)
        'selDetail.TabControlProjectInfo.SelectTab(1)
        'With selDetail.lvPartners
        '    .Focus()
        '    .SelectedItems.Clear()
        '    .Items(ProjectPartnerIndex).Selected = True
        'End With
    End Sub
#End Region

#Region "Search organisations"
    Public Function FindTextInOrganisations(ByVal strFind As String) As Boolean
        Dim boolFound As Boolean
        Dim objValue As Object
        Dim strText As String, strName As String

        Do
            Dim selProjectPartner As ProjectPartner = CurrentLogFrame.ProjectPartners(ProjectPartnerIndex)
            Dim selOrganisation As Organisation = selProjectPartner.Organisation
            If selOrganisation Is Nothing Then Return False
            Dim propInfo() As PropertyInfo = GetType(Organisation).GetProperties

            If SearchOption = SearchOptions.SearchAll Or SearchOption = SearchOptions.SearchOrganisations Then
                If boolOrganisationChecked = False Then
                    Do
                        Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndex)

                        If IsWritableProperty(selPropInfo) = True Then

                            strName = selPropInfo.Name
                            objValue = selPropInfo.GetValue(selOrganisation, Nothing)
                            If objValue IsNot Nothing Then
                                strText = objValue.ToString
                                If Me.chkMatchCase.Checked = False Then strText = strText.ToLower

                                If strText.Contains(strFind) Then
                                    If StartIndex < 0 Then StartIndex = 0
                                    StartIndex = strText.IndexOf(strFind, StartIndex)

                                    ShowPartnerDetailsWindow()

                                    If PopUpOrganisationDetails(selProjectPartner, strName, strText, strFind) = True Then
                                        PropInfoIndex += 1

                                        Return True
                                    End If
                                End If
                            End If
                        End If
                        PropInfoIndex += 1
                        If PropInfoIndex > propInfo.Count - 1 Then
                            PropInfoIndex = 0

                            boolOrganisationChecked = True
                            boolFound = False

                            Exit Do
                        End If
                    Loop
                End If
            Else
                boolOrganisationChecked = True
            End If

            If boolOrganisationChecked = True Then
                If boolOrganisationChildrenChecked = False And SearchOption = SearchOptions.SearchAll Or SearchOption = SearchOptions.SearchOrganisations Then
                    If boolAddressChecked = False And selOrganisation.Addresses.Count > 0 Then
                        boolFound = FindTextInAddresses(selOrganisation.Addresses, selOrganisation.FullName, strFind)
                        If boolFound = True Then
                            PropInfoIndexAddresses += 1
                            Return boolFound
                        End If
                    End If
                    If boolTelephoneNumberChecked = False And selOrganisation.TelephoneNumbers.Count > 0 Then
                        boolFound = FindTextInTelephoneNumbers(selOrganisation.TelephoneNumbers, selOrganisation.FullName, strFind)
                        If boolFound = True Then
                            PropInfoIndexTelephoneNumbers += 1
                            Return boolFound
                        End If
                    End If
                    If boolEmailChecked = False And selOrganisation.Emails.Count > 0 Then
                        boolFound = FindTextInEmails(selOrganisation.Emails, selOrganisation.FullName, strFind)
                        If boolFound = True Then
                            PropInfoIndexEmails += 1
                            Return boolFound
                        End If
                    End If
                    If boolWebsiteChecked = False And selOrganisation.WebSites.Count > 0 Then
                        boolFound = FindTextInWebsites(selOrganisation.WebSites, selOrganisation.FullName, strFind)
                        If boolFound = True Then
                            PropInfoIndexWebsites += 1
                            Return boolFound
                        End If
                    End If
                    boolOrganisationChildrenChecked = True
                Else
                    boolOrganisationChildrenChecked = True
                End If
                If boolOrganisationChildrenChecked = True Then
                    If SearchOption = SearchOptions.SearchAll Or SearchOption = SearchOptions.SearchContacts Then
                        If boolContactChecked = False And selOrganisation.Contacts.Count > 0 Then
                            boolFound = FindTextInContacts(selOrganisation.Contacts, strFind)
                            If boolFound = True Then
                                PropInfoIndexContacts += 1
                                Return boolFound
                            End If
                        End If
                    End If
                End If
            End If

            ProjectPartnerIndex += 1
            boolOrganisationChecked = False
            boolOrganisationChildrenChecked = False
            boolAddressChecked = False
            boolTelephoneNumberChecked = False
            boolEmailChecked = False
            boolWebsiteChecked = False
            boolContactChecked = False
            PropInfoIndexAddresses = 0
            PropInfoIndexTelephoneNumbers = 0
            PropInfoIndexEmails = 0
            PropInfoIndexWebsites = 0
            PropInfoIndexContacts = 0

            If ProjectPartnerIndex > CurrentLogFrame.ProjectPartners.Count - 1 Then
                StartIndex = 0
                StartIndexAddresses = 0
                StartIndexTelephoneNumbers = 0
                StartIndexEmails = 0
                StartIndexWebsites = 0
                StartIndexContacts = 0
                boolFound = False

                Exit Do
            End If
        Loop

        Return boolFound
    End Function

    Private Function PopUpOrganisationDetails(ByVal selProjectPartner As ProjectPartner, ByVal strName As String, ByVal strText As String, ByVal strFind As String) As Boolean
        Dim dialogProjectPartner As New DialogProjectPartner(selProjectPartner)
        Dim boolFound As Boolean
        Dim objParent As Object

        With dialogProjectPartner
            .Show()

            Select Case strName
                Case "Swift", "BankAccount"
                    .TabControlDetails.SelectTab(.TabPageBank)
                    objParent = .TabPageBank
                Case Else
                    objParent = dialogProjectPartner
            End Select

            Dim selControl As Control = GetControl(objParent, strName, strText)
            If selControl IsNot Nothing Then
                SelectTextInControl(selControl, strFind)
                PropInfoIndex += 1
                boolFound = True
            End If

            .TopMost = True
            If .DialogResult = vbOK Then
                selProjectPartner = .ProjectPartner
            End If
        End With

        Return boolFound
    End Function
#End Region

#Region "Search contacts"
    Public Function FindTextInContacts(ByVal selContacts As Contacts, ByVal strFind As String) As Boolean
        Dim boolFound As Boolean
        Dim objValue As Object
        Dim strText As String, strName As String

        Do
            Dim selContact As Contact = selContacts(ContactIndex)
            If selContact Is Nothing Then Return False
            Dim propInfo() As PropertyInfo = GetType(Contact).GetProperties
            If boolContactChecked = False Then
                Do
                    Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndex)

                    If IsWritableProperty(selPropInfo) = True Then

                        strName = selPropInfo.Name
                        objValue = selPropInfo.GetValue(selContact, Nothing)
                        If objValue IsNot Nothing Then
                            strText = objValue.ToString
                            If Me.chkMatchCase.Checked = False Then strText = strText.ToLower

                            If strText.Contains(strFind) Then
                                If StartIndex < 0 Then StartIndex = 0
                                StartIndex = strText.IndexOf(strFind, StartIndex)

                                ShowPartnerDetailsWindow()

                                If PopUpContactDetails(selContact, strName, strText, strFind) = True Then
                                    PropInfoIndex += 1

                                    Return True
                                End If
                            End If
                        End If
                    End If
                    PropInfoIndex += 1
                    If PropInfoIndex > propInfo.Count - 1 Then
                        PropInfoIndex = 0

                        boolContactChecked = True
                        boolAddressChecked = False
                        boolTelephoneNumberChecked = False
                        boolEmailChecked = False
                        boolWebsiteChecked = False
                        boolFound = False

                        Exit Do
                    End If
                Loop
            End If
            If boolContactChecked = True Then
                If boolAddressChecked = False And selContact.Addresses.Count > 0 Then
                    boolFound = FindTextInAddresses(selContact.Addresses, selContact.FullName, strFind)
                    If boolFound = True Then
                        PropInfoIndexAddresses += 1
                        Return boolFound
                    End If
                End If
                If boolTelephoneNumberChecked = False And selContact.TelephoneNumbers.Count > 0 Then
                    boolFound = FindTextInTelephoneNumbers(selContact.TelephoneNumbers, selContact.FullName, strFind)
                    If boolFound = True Then
                        PropInfoIndexTelephoneNumbers += 1
                        Return boolFound
                    End If
                End If
                If boolEmailChecked = False And selContact.Emails.Count > 0 Then
                    boolFound = FindTextInEmails(selContact.Emails, selContact.FullName, strFind)
                    If boolFound = True Then
                        PropInfoIndexEmails += 1
                        Return boolFound
                    End If
                End If
                If boolWebsiteChecked = False And selContact.Websites.Count > 0 Then
                    boolFound = FindTextInWebsites(selContact.Websites, selContact.FullName, strFind)
                    If boolFound = True Then
                        PropInfoIndexWebsites += 1
                        Return boolFound
                    End If
                End If
            End If

            ContactIndex += 1
            boolContactChecked = False
            boolAddressChecked = False
            boolTelephoneNumberChecked = False
            boolEmailChecked = False
            boolWebsiteChecked = False
            PropInfoIndexAddresses = 0
            PropInfoIndexTelephoneNumbers = 0
            PropInfoIndexEmails = 0
            PropInfoIndexWebsites = 0
            PropInfoIndexContacts = 0

            If ContactIndex > selContacts.Count - 1 Then
                ContactIndex = 0
                StartIndex = 0
                StartIndexAddresses = 0
                StartIndexTelephoneNumbers = 0
                StartIndexEmails = 0
                StartIndexWebsites = 0

                boolFound = False

                Exit Do
            End If
        Loop

        Return boolFound
    End Function

    Private Function PopUpContactDetails(ByVal selContact As Contact, ByVal strName As String, ByVal strText As String, ByVal strFind As String) As Boolean
        Dim dialogContact As New DialogContact(selContact)
        Dim boolFound As Boolean

        dialogContact.Show()

        Dim selControl As Control = GetControl(dialogContact, strName, strText)
        If selControl IsNot Nothing Then
            SelectTextInControl(selControl, strFind)
            PropInfoIndex += 1
            boolFound = True
        End If

        dialogContact.TopMost = True
        If dialogContact.DialogResult = vbOK Then
            selContact = dialogContact.Contact
        End If

        Return boolFound
    End Function
#End Region

#Region "Search adresses, telephone numbers, e-mails and websites"
    Public Function FindTextInAddresses(ByVal selAddresses As Addresses, ByVal strOrganisationName As String, ByVal strFind As String) As Boolean
        Dim boolFound As Boolean
        Dim objValue As Object
        Dim strText As String, strName As String

        Do
            Dim selAddress As Address = selAddresses(AddressIndex)
            If selAddress Is Nothing Then Return False
            Dim propInfo() As PropertyInfo = GetType(Address).GetProperties

            Do
                Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndexAddresses)

                If IsWritableProperty(selPropInfo) = True Then

                    strName = selPropInfo.Name
                    objValue = selPropInfo.GetValue(selAddress, Nothing)
                    If objValue IsNot Nothing Then
                        strText = objValue.ToString
                        If Me.chkMatchCase.Checked = False Then strText = strText.ToLower

                        If strText.Contains(strFind) Then
                            StartIndexAddresses = strText.IndexOf(strFind, StartIndexAddresses)

                            ShowPartnerDetailsWindow()

                            If PopUpAddressDetails(selAddress, strOrganisationName, strName, strText, strFind) = True Then
                                PropInfoIndexAddresses += 1

                                Return True
                            End If
                        End If
                    End If
                End If
                PropInfoIndexAddresses += 1
                If PropInfoIndexAddresses > propInfo.Count - 1 Then
                    PropInfoIndexAddresses = 0
                    StartIndexAddresses = 0
                    boolFound = False

                    Exit Do
                End If
            Loop
            AddressIndex += 1
            If AddressIndex > selAddresses.Count - 1 Then
                AddressIndex = 0
                StartIndexAddresses = 0
                boolAddressChecked = True
                boolFound = False

                Exit Do
            End If
        Loop
        Return boolFound
    End Function

    Private Function PopUpAddressDetails(ByVal selAddress As Address, ByVal strOrganisationName As String, ByVal strName As String, ByVal strText As String, ByVal strFind As String) As Boolean
        Dim dialogAddress As New DialogAddress(selAddress)
        dialogAddress.Text &= " of " & strOrganisationName
        Dim boolFound As Boolean

        dialogAddress.Show()
        Dim selControl As Control = GetControl(dialogAddress, strName, strText)
        If selControl IsNot Nothing Then
            SelectTextInControl(selControl, strFind)
            PropInfoIndexAddresses += 1
            boolFound = True
        End If

        dialogAddress.TopMost = True
        If dialogAddress.DialogResult = vbOK Then

            selAddress = dialogAddress.Address
        End If

        Return boolFound
    End Function

    Public Function FindTextInTelephoneNumbers(ByVal selTelephoneNumbers As TelephoneNumbers, ByVal strOrganisationName As String, ByVal strFind As String) As Boolean
        Dim boolFound As Boolean
        Dim objValue As Object
        Dim strText As String, strName As String

        Do
            Dim selTelephoneNumber As TelephoneNumber = selTelephoneNumbers(TelephoneNumberIndex)
            If selTelephoneNumber Is Nothing Then Return False
            Dim propInfo() As PropertyInfo = GetType(TelephoneNumber).GetProperties

            Do
                Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndexTelephoneNumbers)

                If IsWritableProperty(selPropInfo) = True Then

                    strName = selPropInfo.Name
                    objValue = selPropInfo.GetValue(selTelephoneNumber, Nothing)
                    If objValue IsNot Nothing Then
                        strText = objValue.ToString
                        If Me.chkMatchCase.Checked = False Then strText = strText.ToLower

                        If strText.Contains(strFind) Then
                            StartIndexTelephoneNumbers = strText.IndexOf(strFind, StartIndexTelephoneNumbers)

                            ShowPartnerDetailsWindow()

                            If PopUpTelephoneNumberDetails(selTelephoneNumber, strOrganisationName, strName, strText, strFind) = True Then
                                PropInfoIndexTelephoneNumbers += 1

                                Return True
                            End If
                        End If
                    End If
                End If
                PropInfoIndexTelephoneNumbers += 1
                If PropInfoIndexTelephoneNumbers > propInfo.Count - 1 Then
                    PropInfoIndexTelephoneNumbers = 0
                    StartIndexTelephoneNumbers = 0
                    boolFound = False

                    Exit Do
                End If
            Loop
            TelephoneNumberIndex += 1
            If TelephoneNumberIndex > selTelephoneNumbers.Count - 1 Then
                TelephoneNumberIndex = 0
                StartIndexTelephoneNumbers = 0
                boolTelephoneNumberChecked = True
                boolFound = False

                Exit Do
            End If
        Loop
        Return boolFound
    End Function

    Private Function PopUpTelephoneNumberDetails(ByVal selTelephoneNumber As TelephoneNumber, ByVal strOrganisationName As String, ByVal strName As String, ByVal strText As String, ByVal strFind As String) As Boolean
        Dim dialogTelephoneNumber As New DialogTelephoneNumber(selTelephoneNumber)
        dialogTelephoneNumber.Text &= " of " & strOrganisationName
        Dim boolFound As Boolean

        dialogTelephoneNumber.Show()
        Dim selControl As Control = GetControl(dialogTelephoneNumber, strName, strText)
        If selControl IsNot Nothing Then
            SelectTextInControl(selControl, strFind)
            PropInfoIndexTelephoneNumbers += 1
            boolFound = True
        End If

        dialogTelephoneNumber.TopMost = True
        If dialogTelephoneNumber.DialogResult = vbOK Then

            selTelephoneNumber = dialogTelephoneNumber.TelephoneNumber
        End If

        Return boolFound
    End Function

    Public Function FindTextInEmails(ByVal selEmails As Emails, ByVal strOrganisationName As String, ByVal strFind As String) As Boolean
        Dim boolFound As Boolean
        Dim objValue As Object
        Dim strText As String, strName As String

        Do
            Dim selEmail As Email = selEmails(EmailIndex)
            If selEmail Is Nothing Then Return False
            Dim propInfo() As PropertyInfo = GetType(Email).GetProperties

            Do
                Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndexEmails)

                If IsWritableProperty(selPropInfo) = True Then

                    strName = selPropInfo.Name
                    objValue = selPropInfo.GetValue(selEmail, Nothing)
                    If objValue IsNot Nothing Then
                        strText = objValue.ToString
                        If Me.chkMatchCase.Checked = False Then strText = strText.ToLower

                        If strText.Contains(strFind) Then
                            StartIndexEmails = strText.IndexOf(strFind, StartIndexEmails)

                            ShowPartnerDetailsWindow()

                            If PopUpEmailDetails(selEmail, strOrganisationName, strName, strText, strFind) = True Then
                                PropInfoIndexEmails += 1

                                Return True
                            End If
                        End If
                    End If
                End If
                PropInfoIndexEmails += 1
                If PropInfoIndexEmails > propInfo.Count - 1 Then
                    PropInfoIndexEmails = 0
                    StartIndexEmails = 0
                    boolFound = False

                    Exit Do
                End If
            Loop
            EmailIndex += 1
            If EmailIndex > selEmails.Count - 1 Then
                EmailIndex = 0
                StartIndexEmails = 0
                boolEmailChecked = True
                boolFound = False

                Exit Do
            End If
        Loop
        Return boolFound
    End Function

    Private Function PopUpEmailDetails(ByVal selEmail As Email, ByVal strOrganisationName As String, ByVal strName As String, ByVal strText As String, ByVal strFind As String) As Boolean
        Dim dialogEmail As New DialogEmail(selEmail)
        dialogEmail.Text &= " of " & strOrganisationName
        Dim boolFound As Boolean

        dialogEmail.Show()
        Dim selControl As Control = GetControl(dialogEmail, strName, strText)
        If selControl IsNot Nothing Then
            SelectTextInControl(selControl, strFind)
            PropInfoIndexEmails += 1
            boolFound = True
        End If

        dialogEmail.TopMost = True
        If dialogEmail.DialogResult = vbOK Then

            selEmail = dialogEmail.Email
        End If

        Return boolFound
    End Function

    Public Function FindTextInWebsites(ByVal selWebsites As Websites, ByVal strOrganisationName As String, ByVal strFind As String) As Boolean
        Dim boolFound As Boolean
        Dim objValue As Object
        Dim strText As String, strName As String

        Do
            Dim selWebsite As Website = selWebsites(WebsiteIndex)
            If selWebsite Is Nothing Then Return False
            Dim propInfo() As PropertyInfo = GetType(Website).GetProperties

            Do
                Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndexWebsites)

                If IsWritableProperty(selPropInfo) = True Then

                    strName = selPropInfo.Name
                    objValue = selPropInfo.GetValue(selWebsite, Nothing)
                    If objValue IsNot Nothing Then
                        strText = objValue.ToString
                        If Me.chkMatchCase.Checked = False Then strText = strText.ToLower

                        If strText.Contains(strFind) Then
                            StartIndexWebsites = strText.IndexOf(strFind, StartIndexWebsites)

                            ShowPartnerDetailsWindow()

                            If PopUpWebsiteDetails(selWebsite, strOrganisationName, strName, strText, strFind) = True Then
                                PropInfoIndexWebsites += 1

                                Return True
                            End If
                        End If
                    End If
                End If
                PropInfoIndexWebsites += 1
                If PropInfoIndexWebsites > propInfo.Count - 1 Then
                    PropInfoIndexWebsites = 0
                    StartIndexWebsites = 0
                    boolFound = False

                    Exit Do
                End If
            Loop
            WebsiteIndex += 1
            If WebsiteIndex > selWebsites.Count - 1 Then
                WebsiteIndex = 0
                StartIndexWebsites = 0
                boolWebsiteChecked = True
                boolFound = False

                Exit Do
            End If
        Loop
        Return boolFound
    End Function

    Private Function PopUpWebsiteDetails(ByVal selWebsite As Website, ByVal strOrganisationName As String, ByVal strName As String, ByVal strText As String, ByVal strFind As String) As Boolean
        Dim dialogWebsite As New DialogWebsite(selWebsite)
        dialogWebsite.Text &= " of " & strOrganisationName
        Dim boolFound As Boolean

        dialogWebsite.Show()
        Dim selControl As Control = GetControl(dialogWebsite, strName, strText)
        If selControl IsNot Nothing Then
            SelectTextInControl(selControl, strFind)
            PropInfoIndexWebsites += 1
            boolFound = True
        End If

        dialogWebsite.TopMost = True
        If dialogWebsite.DialogResult = vbOK Then

            selWebsite = dialogWebsite.Website
        End If

        Return boolFound
    End Function
#End Region


End Class
