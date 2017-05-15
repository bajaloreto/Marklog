Public Class ListViewContacts
    Inherits ListViewSortable

    Private SearchType As Integer
    Private objContacts As Contacts
    Private objParentOrganisationGuid As Guid

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Name, 150, HorizontalAlignment.Left)
            .Columns.Add(LANG_JobTitle, 150, HorizontalAlignment.Left)
            .Columns.Add(LANG_Mobile, 80, HorizontalAlignment.Left)
            .Columns.Add(LANG_WorkPhone, 80, HorizontalAlignment.Left)
            .Columns.Add(LANG_Email, 150, HorizontalAlignment.Left)
        End With
    End Sub

    Public Property Contacts() As Contacts
        Get
            Return objContacts
        End Get
        Set(ByVal value As Contacts)
            objContacts = value

        End Set
    End Property

    Public Property ParentOrganisationGuid() As Guid
        Get
            Return objParentOrganisationGuid
        End Get
        Set(ByVal value As Guid)
            objParentOrganisationGuid = value
        End Set
    End Property

    Public ReadOnly Property SelectedContact() As Contact
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selContact As Contact = Me.Contacts.GetContactByGuid(Me.SelectedGuid)
                Return selContact
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedContacts() As Contact()
        Get
            Dim objContacts(SelectedItems.Count - 1) As Contact

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objContacts(i) = Me.Contacts.GetContactByGuid(Me.SelectedGuids(i))
                Next

                Return objContacts
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub Reload(ByVal contacts As Contacts, ByVal parentorganisationguid As Guid)
        Me.Contacts = contacts
        Me.ParentOrganisationGuid = parentorganisationguid

        LoadItems()
    End Sub

    Public Sub LoadItems()
        Dim strText As String
        Items.Clear()

        If objContacts IsNot Nothing Then
            For Each selContact As Contact In Me.Contacts
                Dim newItem As New ListViewItem(selContact.FullName)
                newItem.Name = selContact.Guid.ToString
                newItem.SubItems.Add(selContact.JobTitle)

                If selContact.TelephoneNumbers.Count > 0 Then
                    strText = selContact.TelephoneNumbers.GetMobileNumber
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)

                If selContact.TelephoneNumbers.Count > 0 Then
                    strText = selContact.TelephoneNumbers.GetWorkNumber
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)

                If selContact.Emails.GetMainEmail IsNot Nothing Then
                    strText = selContact.Emails.GetMainEmail.Email
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)
                Items.Add(newItem)
            Next
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
    End Sub

    Public Overrides Sub NewItem()
        PopUpContactDetails(Nothing)
    End Sub

    Public Overrides Sub EditItem()
        If Me.Contacts.Count > 0 AndAlso Me.SelectedContact IsNot Nothing Then
            PopUpContactDetails(Me.SelectedContact)
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.Contacts.Count > 0 AndAlso Me.SelectedContact IsNot Nothing Then
            If MsgBox(LANG_RemoveContact, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selContact As Contact = Me.SelectedContacts(0)
                UndoRedo.ItemRemoved(selContact, Me.Contacts)

                Me.Contacts.Remove(selContact)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpContactDetails(ByVal selContact As Contact)
        Dim boolNew As Boolean

        If selContact Is Nothing Then
            If Me.ParentOrganisationGuid = Guid.Empty Then Exit Sub
            boolNew = True
            selContact = New Contact
            selContact.ParentOrganisationGuid = Me.ParentOrganisationGuid
        End If

        Dim dialogContact As New DialogContact(selContact)

        dialogContact.ShowDialog()
        If dialogContact.DialogResult = vbOK Then
            If boolNew = True Then
                Dim objOrganisation As Organisation = CurrentLogFrame.GetOrganisationByGuid(Me.ParentOrganisationGuid)

                objOrganisation.Contacts.Add(selContact)
                Me.Reload(objOrganisation.Contacts, objOrganisation.Guid)
                'Me.Contacts.Add(selContact)
                UndoRedo.ItemInserted(selContact, objOrganisation.Contacts)
            End If

            Me.LoadItems()
        End If
        dialogContact.Dispose()
        dialogContact = Nothing
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selContact As Contact In SelectedContacts
            UndoRedo.ItemCut(selContact, Me.Contacts)
            Contacts.Remove(selContact)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selContact As Contact In SelectedContacts
            Dim NewItem As New ClipboardItem(CopyGroup, selContact, Contact.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selContact As Contact
        Dim intCountName As Integer

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Contact)
                    selContact = CType(selItem.Item, Contact)
                    Dim NewContact As New Contact

                    Using copier As New ObjectCopy
                        NewContact = copier.CopyObject(selContact)
                    End Using

                    intCountName = Contacts.VerifyIfFullNameExists(NewContact.FullName)
                    If intCountName > 0 Then
                        NewContact.LastName &= String.Format(" ({0})", {intCountName})
                    End If
                    Me.Contacts.Add(NewContact)

                    UndoRedo.ItemPasted(NewContact, Me.Contacts)
            End Select
        Next

        Me.LoadItems()
    End Sub
End Class
