Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class TelephoneNumber
    Private intIdTelephoneNumber, intIdOrganisation As Integer
    Private strNumber As String
    Private intType As Integer
    Private objGuid, objParentGuid As Guid

    Public Enum TelephoneNumberTypes
        WorkGeneral = 0
        Work = 1
        WorkFax = 2
        Home = 3
        HomeFax = 4
        Mobile = 5
        Internal = 6
        Temporary = 7
    End Enum

    Public Property idTelephoneNumber As Integer
        Get
            Return intIdTelephoneNumber
        End Get
        Set(value As Integer)
            intIdTelephoneNumber = value
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

    Public Property Number() As String
        Get
            Return strNumber
        End Get
        Set(ByVal value As String)
            strNumber = value
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
            Return LIST_TelephoneNumberTypes(intType)
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
            Return LANG_TelephoneNumber
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_TelephoneNumbers
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Me.Number
    End Function
End Class

Public Class TelephoneNumbers
    Inherits System.Collections.CollectionBase

    Public Event TelephoneNumberAdded(ByVal sender As Object, ByVal e As TelephoneNumberAddedEventArgs)

    Public Sub Add(ByVal telephonenumber As TelephoneNumber)
        List.Add(telephonenumber)
        RaiseEvent TelephoneNumberAdded(Me, New TelephoneNumberAddedEventArgs(telephonenumber))
    End Sub

    Public Function Contains(ByVal telephonenumber As TelephoneNumber) As Boolean
        Return List.Contains(telephonenumber)
    End Function

    Public Function IndexOf(ByVal telephonenumber As TelephoneNumber) As Integer
        Return List.IndexOf(telephonenumber)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal telephonenumber As TelephoneNumber)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of telephone number is not valid, cannot be inserted!")
        ElseIf index = Count Then
            List.Add(telephonenumber)
            RaiseEvent TelephoneNumberAdded(Me, New TelephoneNumberAddedEventArgs(telephonenumber))
        Else
            List.Insert(index, telephonenumber)
            RaiseEvent TelephoneNumberAdded(Me, New TelephoneNumberAddedEventArgs(telephonenumber))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As TelephoneNumber
        Get
            Return CType(List.Item(index), TelephoneNumber)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal telephonenumber As TelephoneNumber)
        If List.Contains(telephonenumber) Then List.Remove(telephonenumber)
    End Sub

    Public Function GetMainTelephoneNumber() As TelephoneNumber
        Dim selItem As TelephoneNumber = Nothing
        If List.Count > 0 Then
            If List.Count = 1 Then
                selItem = List.Item(0)
            Else
                For Each selItem In List
                    If selItem.Type = TelephoneNumber.TelephoneNumberTypes.WorkGeneral Then Exit For
                Next
                If selItem Is Nothing Then selItem = List.Item(0)
            End If

        Else
            selItem = Nothing
        End If
        Return selItem
    End Function

    Public Function GetMobileNumber() As String
        Dim strNumber As String = ""

        For Each selItem As TelephoneNumber In List
            If selItem.Type = TelephoneNumber.TelephoneNumberTypes.Mobile Then
                strNumber = selItem.Number
                Exit For
            End If

        Next

        Return strNumber
    End Function

    Public Function GetWorkNumber() As String
        Dim strNumber As String = ""

        For Each selItem As TelephoneNumber In List
            If selItem.Type = TelephoneNumber.TelephoneNumberTypes.Work Then
                strNumber = selItem.Number
                Exit For
            End If

        Next

        Return strNumber
    End Function

    Public Function GetTelephoneNumberByGuid(ByVal objGuid As Guid) As TelephoneNumber
        Dim selTelephoneNumber As TelephoneNumber = Nothing
        For Each objTelephoneNumber As TelephoneNumber In Me.List
            If objTelephoneNumber.Guid = objGuid Then
                selTelephoneNumber = objTelephoneNumber
                Exit For
            End If
        Next
        Return selTelephoneNumber
    End Function
End Class

Public Class TelephoneNumberAddedEventArgs
    Inherits EventArgs

    Public Property TelephoneNumber As TelephoneNumber

    Public Sub New(ByVal objTelephoneNumber As TelephoneNumber)
        MyBase.New()

        Me.TelephoneNumber = objTelephoneNumber
    End Sub
End Class
