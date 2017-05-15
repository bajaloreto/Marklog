Partial Public Class frmParent

    Private Sub RibbonButtonSendEmail_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonSendEmail.Click
        Select Case CurrentControl.GetType
            Case GetType(ListViewPartners)
                Dim lvPartner As ListViewPartners = CType(CurrentControl, ListViewPartners)
                Dim selPartner As ProjectPartner = lvPartner.SelectedProjectPartner

                If selPartner IsNot Nothing Then
                    If selPartner.Organisation.Emails.Count > 0 Then
                        SendMainEmail(selPartner.Organisation.Emails)
                    End If

                End If
            Case GetType(ListViewContacts)
                Dim lvContact As ListViewContacts = CType(CurrentControl, ListViewContacts)
                Dim selContact As Contact = lvContact.SelectedContact

                If selContact IsNot Nothing Then
                    If selContact.Emails.Count > 0 Then
                        SendMainEmail(selContact.Emails)
                    End If
                End If
        End Select
    End Sub

    Private Sub SendMainEmail(ByVal objEmails As Emails)
        Dim selEmail As Email = objEmails.GetMainEmail()

        If selEmail IsNot Nothing Then
            Using objEmailIO As New EmailIO
                objEmailIO.StartMailClient(selEmail)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonSendLetter_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonSendLetter.Click
        Select Case CurrentControl.GetType
            Case GetType(ListViewPartners)
                Dim lvPartner As ListViewPartners = CType(CurrentControl, ListViewPartners)
                Dim selPartner As ProjectPartner = lvPartner.SelectedProjectPartner

                If selPartner IsNot Nothing Then
                    If selPartner.Organisation.Addresses.Count > 0 Then
                        SendLetterToMainAddress(Nothing, selPartner.Organisation, selPartner.Organisation.Addresses)
                    End If

                End If
            Case GetType(ListViewContacts)
                Dim lvContact As ListViewContacts = CType(CurrentControl, ListViewContacts)
                Dim selContact As Contact = lvContact.SelectedContact

                If selContact IsNot Nothing Then
                    If selContact.Addresses.Count > 0 Then
                        Dim selOrganisation As Organisation = CurrentLogFrame.GetParent(selContact)
                        SendLetterToMainAddress(selContact, selOrganisation, selContact.Addresses)
                    End If
                End If
        End Select
    End Sub

    Private Sub SendLetterToMainAddress(ByVal selContact As Contact, ByVal selOrganisation As Organisation, ByVal objAddresses As Addresses)
        Dim selAddress As Address = objAddresses.GetMainAddress()

        If selAddress IsNot Nothing Then
            Using objWordIO As New WordIO
                objWordIO.SendLetterToAddress(selContact, selOrganisation, selAddress)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonOpenWebsite_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonOpenWebsite.Click
        Select Case CurrentControl.GetType
            Case GetType(ListViewPartners)
                Dim lvPartner As ListViewPartners = CType(CurrentControl, ListViewPartners)
                Dim selPartner As ProjectPartner = lvPartner.SelectedProjectPartner

                If selPartner IsNot Nothing Then
                    If selPartner.Organisation.Websites.Count > 0 Then
                        OpenFirstWebsite(selPartner.Organisation.Websites)
                    End If

                End If
            Case GetType(ListViewContacts)
                Dim lvContact As ListViewContacts = CType(CurrentControl, ListViewContacts)
                Dim selContact As Contact = lvContact.SelectedContact

                If selContact IsNot Nothing Then
                    If selContact.Websites.Count > 0 Then
                        OpenFirstWebsite(selContact.Websites)
                    End If
                End If
        End Select
    End Sub

    Private Sub OpenFirstWebsite(ByVal objWebsites As Websites)
        Dim selWebsite As Website = objWebsites(0)

        If selWebsite IsNot Nothing Then
            Using objInternet As New classInternet
                objInternet.ExecuteFile(selWebsite.WebsiteUrl, ERR_CannotOpenWebsite)
            End Using
        End If
    End Sub

    Private Sub RibbonButtonSkype_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSkype.Click
        Select Case CurrentControl.GetType
            Case GetType(ListViewContacts)
                Dim lvContact As ListViewContacts = CType(CurrentControl, ListViewContacts)
                Dim selContact As Contact = lvContact.SelectedContact

                If selContact IsNot Nothing Then
                    Dim strSkypeAccount As String = selContact.SkypeAccount

                    If String.IsNullOrEmpty(strSkypeAccount) = False Then
                        Using objSkype As New SkypeIO
                            objSkype.MakeCall(strSkypeAccount)
                        End Using
                    End If
                End If
        End Select
    End Sub
End Class
