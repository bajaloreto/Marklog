Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Organisation
    Public Event GuidChanged(ByVal sender As Object)

    Private intIdOrganisation, intIdParentProjectPartner As Integer
    Private strName As String
    Private strAcronym As String
    Private intType As Integer
    Private intRelation As Integer
    Private strMemo As String
    Private strSwift As String
    Private strBankAccount As String
    Private objGuid, objParentProjectPartnerGuid As Guid

    <ScriptIgnore()> _
    Public WithEvents Addresses As New Addresses

    <ScriptIgnore()> _
    Public WithEvents TelephoneNumbers As New TelephoneNumbers

    <ScriptIgnore()> _
    Public WithEvents Emails As New Emails

    <ScriptIgnore()> _
    Public WithEvents Websites As New Websites

    <ScriptIgnore()> _
    Public WithEvents Contacts As New Contacts

#Region "Enumerations"
    Public Enum Types As Integer
        NonProfit = 0
        NonGovernmental = 1
        Governmental = 2
        Multilateral = 3
        Company = 4
        Bank = 5
        School = 6
        Medical = 7
        Religious = 8
        Other = 9
    End Enum

    Public Enum Roles As Integer
        Lead = 0
        Partner = 1
        ImplementingPartner = 2
        Donor = 3
        Financial = 4
        Supplier = 5
        ServiceProvider = 6
        Auditor = 7
        Network = 8
        Other = 9
    End Enum
#End Region 'Enumerations

#Region "Properties"
    Public Property idOrganisation As Integer
        Get
            Return intIdOrganisation
        End Get
        Set(ByVal value As Integer)
            intIdOrganisation = value
        End Set
    End Property

    Public Property idParentProjectPartner As Integer
        Get
            Return intIdParentProjectPartner
        End Get
        Set(ByVal value As Integer)
            intIdParentProjectPartner = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
        End Set
    End Property

    Public Property Acronym() As String
        Get
            Return strAcronym
        End Get
        Set(ByVal value As String)
            strAcronym = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property FullName() As String
        Get
            Dim strName As String = String.Empty
            If String.IsNullOrEmpty(Me.Acronym) = False Then strName = Me.Acronym.ToUpper
            If String.IsNullOrEmpty(Me.Name) = False Then
                If String.IsNullOrEmpty(strName) = False Then strName &= " - "
                strName &= Me.Name
            End If
            Return strName
        End Get
    End Property

    Public Property Type() As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property TypeName() As String
        Get
            If intType >= 0 Then
                Return LIST_OrganisationTypes(intType)
            Else
                Return "  -"
            End If
        End Get
    End Property

    Public Property Memo() As String
        Get
            Return strMemo
        End Get
        Set(ByVal value As String)
            strMemo = value
        End Set
    End Property

    Public Property Swift() As String
        Get
            Return strSwift
        End Get
        Set(ByVal value As String)
            strSwift = value
        End Set
    End Property

    Public Property BankAccount() As String
        Get
            Return strBankAccount
        End Get
        Set(ByVal value As String)
            strBankAccount = value
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

    Public Property ParentProjectPartnerGuid() As Guid
        Get
            Return objParentProjectPartnerGuid
        End Get
        Set(ByVal value As Guid)
            objParentProjectPartnerGuid = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property MainAddress() As String
        Get
            Dim strAddress As String = String.Empty
            Dim selAddress As Address = Me.Addresses.GetMainAddress()
            If selAddress IsNot Nothing Then
                strAddress = selAddress.FullAddress
            End If
            Return strAddress
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property MainTown() As String
        Get
            Dim strTown As String = String.Empty
            Dim selAddress As Address = Me.Addresses.GetMainAddress()
            If selAddress IsNot Nothing Then
                strTown = selAddress.Town
            End If
            Return strTown
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property MainTelephoneNumber() As String
        Get
            Dim strTelephoneNumber As String = String.Empty
            Dim selTelephoneNumber As TelephoneNumber = Me.TelephoneNumbers.GetMainTelephoneNumber()
            If selTelephoneNumber IsNot Nothing Then
                strTelephoneNumber = Trim(selTelephoneNumber.Number)
            End If
            Return strTelephoneNumber
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property MainEmail() As String
        Get
            Dim strEmail As String = String.Empty
            Dim selEmail As Email = Me.Emails.GetMainEmail()
            If selEmail IsNot Nothing Then
                strEmail = Trim(selEmail.Email)
            End If
            Return strEmail
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Organisation
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Organisations
        End Get
    End Property
#End Region 'Properties

#Region "Methods"
    Public Sub New()

    End Sub

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

    Private Sub Contacts_ContactAdded(ByVal sender As Object, ByVal e As ContactAddedEventArgs) Handles Contacts.ContactAdded
        Dim selContact As Contact = e.Contact

        selContact.idOrganisation = Me.idOrganisation
        selContact.ParentOrganisationGuid = Me.Guid
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
        For Each selContact As Contact In Me.Contacts
            selContact.ParentOrganisationGuid = Me.Guid
        Next
    End Sub

    Public Overrides Function ToString() As String
        Return Me.Name
    End Function
#End Region 'Methods
End Class

Public Class Organisations
    Inherits System.Collections.CollectionBase


    Public Sub Add(ByVal organisation As Organisation)
        List.Add(organisation)
    End Sub

    Public Sub AddRange(ByVal organisations As List(Of Organisation))
        InnerList.AddRange(organisations)
    End Sub

    Public Function Contains(ByVal organisation As Organisation) As Boolean
        Return List.Contains(organisation)
    End Function

    Public Function IndexOf(ByVal organisation As Organisation) As Integer
        Return List.IndexOf(organisation)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal organisation As Organisation)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, organisation.ItemName))
        ElseIf index = Count Then
            List.Add(organisation)
        Else
            List.Insert(index, organisation)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Organisation
        Get
            Return CType(List.Item(index), Organisation)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, Organisation.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal organisation As Organisation)
        If List.Contains(organisation) Then List.Remove(organisation)
    End Sub

    Public Function VerifyIfNameExists(ByVal strName As String) As Integer
        Dim intCount As Integer
        For Each selOrganisation As Organisation In Me.List
            If selOrganisation.Name.StartsWith(strName) Then intCount += 1
        Next

        Return intCount
    End Function

    Public Function GetOrganisationByGuid(ByVal objGuid As Guid) As Organisation
        Dim selOrganisation As Organisation = Nothing
        For Each objOrganisation As Organisation In Me.List
            If objOrganisation.Guid = objGuid Then
                selOrganisation = objOrganisation
                Exit For
            End If
        Next
        Return selOrganisation
    End Function
End Class
