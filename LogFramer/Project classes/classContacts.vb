Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Contact
    Public Event GuidChanged(ByVal sender As Object)

    Private intIdContact, intIdOrganisation As Integer
    Private strFirstName As String
    Private strLastName As String
    Private strTitle As String
    Private strJobTitle As String
    Private intGender As Integer
    Private strRole As String
    Private strSkypeAccount As String
    Private strMemo As String
    Private objGuid, objParentOrganisationGuid As Guid

    <ScriptIgnore()> _
    Public WithEvents Addresses As New Addresses

    <ScriptIgnore()> _
    Public WithEvents TelephoneNumbers As New TelephoneNumbers

    <ScriptIgnore()> _
    Public WithEvents Emails As New Emails

    <ScriptIgnore()> _
    Public WithEvents Websites As New Websites

#Region "Enumerations"
    Public Enum GenderValues As Integer
        Male = 0
        Female = 1
    End Enum
#End Region 'Enumerations

#Region "Properties"
    Public Property idContact As Integer
        Get
            Return intIdContact
        End Get
        Set(value As Integer)
            intIdContact = value
        End Set
    End Property

    Public Property idOrganisation As Integer
        Get
            Return intIdOrganisation
        End Get
        Set(value As Integer)
            intIdOrganisation = value
        End Set
    End Property

    Public Property FirstName() As String
        Get
            Return strFirstName
        End Get
        Set(ByVal value As String)
            strFirstName = value
        End Set
    End Property

    Public Property LastName() As String
        Get
            Return strLastName
        End Get
        Set(ByVal value As String)
            strLastName = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property FullName() As String
        Get
            Dim strName As String = String.Empty
            If String.IsNullOrEmpty(Me.LastName) = False Then strName = strLastName.ToUpper
            If String.IsNullOrEmpty(Me.FirstName) = False Then strName &= If(strName = String.Empty, String.Empty, ", ") & Me.FirstName
            Return strName
        End Get
    End Property

    Public Property Title() As String
        Get
            Return strTitle
        End Get
        Set(ByVal value As String)
            strTitle = value
        End Set
    End Property

    Public Property JobTitle() As String
        Get
            Return strJobTitle
        End Get
        Set(ByVal value As String)
            strJobTitle = value
        End Set
    End Property

    Public Property Gender() As Integer
        Get
            Return intGender
        End Get
        Set(ByVal value As Integer)
            intGender = value
        End Set
    End Property

    Public Property Role() As String
        Get
            Return strRole
        End Get
        Set(ByVal value As String)
            strRole = value
        End Set
    End Property

    Public Property SkypeAccount() As String
        Get
            Return strSkypeAccount
        End Get
        Set(ByVal value As String)
            strSkypeAccount = value
        End Set
    End Property

    Public Property Memo() As String
        Get
            Return strMemo
        End Get
        Set(ByVal value As String)
            strMemo = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
                RaiseEvent GuidChanged(Me)
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
            RaiseEvent GuidChanged(Me)
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

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Contact
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Contacts
        End Get
    End Property
#End Region 'Properties

#Region "Methods and events"
    Private Sub Addresses_AddressAdded(ByVal sender As Object, ByVal e As AddressAddedEventArgs) Handles Addresses.AddressAdded
        Dim selAddress As Address = e.Address

        selAddress.idOrganisation = Me.idOrganisation
        selAddress.ParentGuid = Me.Guid
    End Sub

    Private Sub TelephoneNumbers_TelephoneNumberAdded(ByVal sender As Object, ByVal e As TelephoneNumberAddedEventArgs) Handles TelephoneNumbers.TelephoneNumberAdded
        Dim selTelephoneNumber As TelephoneNumber = e.TelephoneNumber

        selTelephoneNumber.idOrganisation = Me.idOrganisation
        selTelephoneNumber.ParentGuid = Me.Guid
    End Sub

    Private Sub Emails_EmailAdded(ByVal sender As Object, ByVal e As EmailAddedEventArgs) Handles Emails.EmailAdded
        Dim selEmail As Email = e.Email

        selEmail.idOrganisation = Me.idOrganisation
        selEmail.ParentGuid = Me.Guid
    End Sub

    Private Sub Websites_WebsiteAdded(ByVal sender As Object, ByVal e As WebsiteAddedEventArgs) Handles Websites.WebsiteAdded
        Dim selWebsite As Website = e.Website

        selWebsite.idOrganisation = Me.idOrganisation
        selWebsite.ParentGuid = Me.Guid
    End Sub

    Private Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selAddress As Address In Me.Addresses
            selAddress.ParentGuid = Me.Guid
        Next
        For Each selTelephoneNumber As TelephoneNumber In Me.TelephoneNumbers
            selTelephoneNumber.ParentGuid = Me.Guid
        Next
        For Each selEmail As Email In Me.Emails
            selEmail.ParentGuid = Me.Guid
        Next
        For Each selWebsite As Website In Me.Websites
            selWebsite.ParentGuid = Me.Guid
        Next
    End Sub
#End Region
End Class

Public Class Contacts
    Inherits System.Collections.CollectionBase

    Public Event ContactAdded(ByVal sender As Object, ByVal e As ContactAddedEventArgs)

    Public Sub Add(ByVal contact As Contact)
        List.Add(contact)
        RaiseEvent ContactAdded(Me, New ContactAddedEventArgs(contact))
    End Sub

    Public Sub AddRange(ByVal contacts As Contacts)
        InnerList.AddRange(contacts)
    End Sub

    Public Function Contains(ByVal contact As Contact) As Boolean
        Return List.Contains(contact)
    End Function

    Public Function IndexOf(ByVal contact As Contact) As Integer
        Return List.IndexOf(contact)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal contact As Contact)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, contact.ItemName))
        ElseIf index = Count Then
            List.Add(contact)
            RaiseEvent ContactAdded(Me, New ContactAddedEventArgs(contact))
        Else
            List.Insert(index, contact)
            RaiseEvent ContactAdded(Me, New ContactAddedEventArgs(contact))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Contact
        Get
            Return CType(List.Item(index), Contact)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, Contact.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal contact As Contact)
        If List.Contains(contact) Then List.Remove(contact)
    End Sub

    Public Function VerifyIfFullNameExists(ByVal strFullName As String) As Integer
        Dim intCount As Integer
        For Each selContact As Contact In Me.List
            If selContact.FullName.StartsWith(strFullName) Then intCount += 1
        Next

        Return intCount
    End Function

    Public Function GetContactByGuid(ByVal objGuid As Guid) As Contact
        Dim selContact As Contact = Nothing
        For Each objContact As Contact In Me.List
            If objContact.Guid = objGuid Then
                selContact = objContact
                Exit For
            End If
        Next
        Return selContact
    End Function
End Class

Public Class ContactAddedEventArgs
    Inherits EventArgs

    Public Property Contact As Contact

    Public Sub New(ByVal objContact As Contact)
        MyBase.New()

        Me.Contact = objContact
    End Sub
End Class
