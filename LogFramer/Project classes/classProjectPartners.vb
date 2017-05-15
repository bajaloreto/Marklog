Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ProjectPartner
    Public Event GuidChanged(ByVal sender As Object)

    Private intIdProjectPartner As Integer
    Private strName As String
    Private strAbreviation As String
    Private intType As Integer
    Private intRelation As Integer
    Private strMemo As String
    Private strSwift As String
    Private strBankAccount As String
    Private objGuid As Guid

    Private objOrganisation As New Organisation

#Region "Enumerations"
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
    <ScriptIgnore()> _
    Public Property Organisation As Organisation
        Get
            Return objOrganisation
        End Get
        Set(ByVal value As Organisation)
            objOrganisation = value
        End Set
    End Property

    Public Property idProjectPartner As Integer
        Get
            Return intIdProjectPartner
        End Get
        Set(ByVal value As Integer)
            intIdProjectPartner = value
        End Set
    End Property

    Public Property Role() As Integer
        Get
            Return intRelation
        End Get
        Set(ByVal value As Integer)
            intRelation = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RoleName() As String
        Get
            If intRelation >= 0 Then
                Return LIST_ProjectPartnerRoleNames(intRelation)
            Else
                Return "  -"
            End If
        End Get
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

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_ProjectPartner
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_ProjectPartners
        End Get
    End Property
#End Region 'Properties

#Region "Methods"
    Public Sub New()

    End Sub

    Private Sub OnGuidChanged() Handles Me.GuidChanged
        Me.Organisation.ParentProjectPartnerGuid = Me.Guid
    End Sub

    Public Overrides Function ToString() As String
        If Me.Organisation IsNot Nothing Then
            Return Me.Organisation.Name
        Else
            Return String.Empty
        End If
    End Function
#End Region 'Methods
End Class

Public Class ProjectPartners
    Inherits System.Collections.CollectionBase


    Public Sub Add(ByVal projectpartner As ProjectPartner)
        List.Add(projectpartner)
    End Sub

    Public Function Contains(ByVal projectpartner As ProjectPartner) As Boolean
        Return List.Contains(projectpartner)
    End Function

    Public Function IndexOf(ByVal projectpartner As ProjectPartner) As Integer
        Return List.IndexOf(projectpartner)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal projectpartner As ProjectPartner)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, projectpartner.ItemName))
        ElseIf index = Count Then
            List.Add(projectpartner)
        Else
            List.Insert(index, projectpartner)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ProjectPartner
        Get
            Return CType(List.Item(index), ProjectPartner)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, ProjectPartner.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal projectpartner As ProjectPartner)
        If List.Contains(projectpartner) Then List.Remove(projectpartner)
    End Sub

    Public Function GetProjectPartnerByGuid(ByVal objGuid As Guid) As ProjectPartner
        Dim selProjectPartner As ProjectPartner = Nothing
        For Each objProjectPartner As ProjectPartner In Me.List
            If objProjectPartner.Guid = objGuid Then
                selProjectPartner = objProjectPartner
                Exit For
            End If
        Next
        Return selProjectPartner
    End Function

    Public Function GetProjectPartnerByOrganisationGuid(ByVal objOrganisationGuid As Guid) As ProjectPartner
        Dim selProjectPartner As ProjectPartner = Nothing

        For Each objProjectPartner As ProjectPartner In Me.List
            If objProjectPartner.Organisation.Guid = objOrganisationGuid Then
                selProjectPartner = objProjectPartner
                Exit For
            End If
        Next
        Return selProjectPartner
    End Function

    Public Function GetOrganisationByGuid(ByVal objGuid As Guid) As Organisation
        Dim selOrganisation As Organisation = Nothing
        For Each objProjectPartner As ProjectPartner In Me.List
            If objProjectPartner.Organisation.Guid = objGuid Then
                selOrganisation = objProjectPartner.Organisation
                Exit For
            End If
        Next
        Return selOrganisation
    End Function
End Class
