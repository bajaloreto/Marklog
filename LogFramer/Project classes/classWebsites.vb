Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Website
    Private intIdWebsite, intIdOrganisation As Integer
    Private strWebsiteName As String
    Private strWebsitePath As String
    Private intType As Integer
    Private objGuid, objParentGuid As Guid

    Public Enum WebsiteTypes
        WorkGeneral = 0
        Work = 1
        Home = 2
    End Enum

    Public Property idWebsite As Integer
        Get
            Return intIdWebsite
        End Get
        Set(value As Integer)
            intIdWebsite = value
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

    Public Property WebsiteName() As String
        Get
            Return strWebsiteName
        End Get
        Set(ByVal value As String)
            strWebsiteName = value
        End Set
    End Property

    Public Property WebsiteUrl() As String
        Get
            Return strWebsitePath
        End Get
        Set(ByVal value As String)
            strWebsitePath = value
        End Set
    End Property

    Public Property Type() As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    Public ReadOnly Property TypeName() As String
        Get
            Return LIST_WebsiteTypes(intType)
        End Get
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ParentGuid() As Guid
        Get
            Return objParentGuid
        End Get
        Set(ByVal value As Guid)
            objParentGuid = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Website
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Websites
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Me.WebsiteName
    End Function
End Class

Public Class Websites
    Inherits System.Collections.CollectionBase

    Public Event WebsiteAdded(ByVal sender As Object, ByVal e As WebsiteAddedEventArgs)

    Public Sub Add(ByVal website As Website)
        List.Add(website)
        RaiseEvent WebsiteAdded(Me, New WebsiteAddedEventArgs(website))
    End Sub

    Public Function Contains(ByVal website As Website) As Boolean
        Return List.Contains(website)
    End Function

    Public Function IndexOf(ByVal website As Website) As Integer
        Return List.IndexOf(website)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal website As Website)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of website is not valid, cannot be inserted!")
        ElseIf index = Count Then
            List.Add(website)
            RaiseEvent WebsiteAdded(Me, New WebsiteAddedEventArgs(website))
        Else
            List.Insert(index, website)
            RaiseEvent WebsiteAdded(Me, New WebsiteAddedEventArgs(website))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Website
        Get
            Return CType(List.Item(index), Website)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal website As Website)
        If List.Contains(website) Then List.Remove(website)
    End Sub

    Public Function GetWebsiteByGuid(ByVal objGuid As Guid) As Website
        Dim selWebsite As Website = Nothing
        For Each objWebsite As Website In Me.List
            If objWebsite.Guid = objGuid Then
                selWebsite = objWebsite
                Exit For
            End If
        Next
        Return selWebsite
    End Function
End Class

Public Class WebsiteAddedEventArgs
    Inherits EventArgs

    Public Property Website As Website

    Public Sub New(ByVal objWebsite As Website)
        MyBase.New()

        Me.Website = objWebsite
    End Sub
End Class
