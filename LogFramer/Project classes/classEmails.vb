Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Email
    Private intIdEmail, intIdOrganisation As Integer
    Private strEmail As String
    Private intType As Integer
    Private intState As Integer
    Private objGuid, objParentGuid As Guid

    Public Enum EmailTypes
        WorkGeneral = 0
        Work = 1
        Home = 2
        Temporary = 3
    End Enum

    Public Property idEmail As Integer
        Get
            Return intIdEmail
        End Get
        Set(value As Integer)
            intIdEmail = value
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

    Public Property Email() As String
        Get
            Return strEmail
        End Get
        Set(ByVal value As String)
            strEmail = value
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
            Return LIST_EmailTypes(intType)
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
            Return LANG_Email
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Emails
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Me.Email
    End Function
End Class

Public Class Emails
    Inherits System.Collections.CollectionBase

    Public Event EmailAdded(ByVal sender As Object, ByVal e As EmailAddedEventArgs)

    Public Sub Add(ByVal email As Email)
        List.Add(email)
        RaiseEvent EmailAdded(Me, New EmailAddedEventArgs(email))
    End Sub

    Public Function Contains(ByVal email As Email) As Boolean
        Return List.Contains(email)
    End Function

    Public Function IndexOf(ByVal email As Email) As Integer
        Return List.IndexOf(email)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal email As Email)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of email is not valid, cannot be inserted!")
        ElseIf index = Count Then
            List.Add(email)
            RaiseEvent EmailAdded(Me, New EmailAddedEventArgs(email))
        Else
            List.Insert(index, email)
            RaiseEvent EmailAdded(Me, New EmailAddedEventArgs(email))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Email
        Get
            Return CType(List.Item(index), Email)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal email As Email)
        If List.Contains(email) Then List.Remove(email)
    End Sub

    Public Function GetMainEmail() As Email
        Dim selItem As Email = Nothing
        If List.Count > 0 Then
            If List.Count = 1 Then
                selItem = List.Item(0)
            Else
                For Each selItem In List
                    If selItem.Type = Email.EmailTypes.WorkGeneral Then Exit For
                Next
                If selItem Is Nothing Then selItem = List.Item(0)
            End If

        Else
            selItem = Nothing
        End If
        Return selItem
    End Function

    Public Function GetEmailByGuid(ByVal objGuid As Guid) As Email
        Dim selEmail As Email = Nothing
        For Each objEmail As Email In Me.List
            If objEmail.Guid = objGuid Then
                selEmail = objEmail
                Exit For
            End If
        Next
        Return selEmail
    End Function
End Class

Public Class EmailAddedEventArgs
    Inherits EventArgs

    Public Property Email As Email

    Public Sub New(ByVal objEmail As Email)
        MyBase.New()

        Me.Email = objEmail
    End Sub
End Class
