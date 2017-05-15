Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Address
    Private intIdAddress, intIdOrganisation As Integer
    Private strStreet As String
    Private strNumber As String
    Private strPostNumber As String
    Private strDistrict As String
    Private strTown As String
    Private strCountry As String
    Private intCountryIndex As Integer
    Private intType As Integer
    Private objGuid, objParentGuid As Guid

    Public Enum AddressTypes
        Work = 0
        Home = 1
        Temporary = 2
        Main = 3
        Correspondence = 4
        Other = 5
    End Enum

    Public Property idAddress As Integer
        Get
            Return intIdAddress
        End Get
        Set(value As Integer)
            intIdAddress = value
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

    Public Property Street() As String
        Get
            Return strStreet
        End Get
        Set(ByVal value As String)
            strStreet = value
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

    Public Property PostNumber() As String
        Get
            Return strPostNumber
        End Get
        Set(ByVal value As String)
            strPostNumber = value
        End Set
    End Property

    Public Property District() As String
        Get
            Return strDistrict
        End Get
        Set(ByVal value As String)
            strDistrict = value
        End Set
    End Property

    Public Property Town() As String
        Get
            Return strTown
        End Get
        Set(ByVal value As String)
            strTown = value
        End Set
    End Property

    Public ReadOnly Property FullStreet() As String
        Get
            Dim strAddress As String = Me.Street & _
                If(String.IsNullOrEmpty(Me.Number), String.Empty, " ") & Me.Number

            Return strAddress
        End Get
    End Property

    Public ReadOnly Property FullTown() As String
        Get
            Dim strAddress As String = Me.PostNumber & _
                If(String.IsNullOrEmpty(Me.PostNumber), String.Empty, " ") & Me.District & _
                If(String.IsNullOrEmpty(Me.District), " ", " - ") & Me.Town

            Return strAddress
        End Get
    End Property

    Public ReadOnly Property FullAddress() As String
        Get
            Dim strAddress As String = Me.FullStreet
            If String.IsNullOrEmpty(strAddress) = False And String.IsNullOrEmpty(Me.FullTown) = False Then strAddress &= " - "
            strAddress &= Me.FullTown
            If String.IsNullOrEmpty(strAddress) = False And String.IsNullOrEmpty(Me.Country) = False Then strAddress &= " - "
            strAddress &= Me.Country.ToUpper

            Return strAddress
        End Get
    End Property

    Public Property Country() As String
        Get
            Return strCountry
        End Get
        Set(ByVal value As String)
            strCountry = value
        End Set
    End Property

    Public Property CountryIndex() As Integer
        Get
            Return intCountryIndex
        End Get
        Set(ByVal value As Integer)
            intCountryIndex = value
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
            Return LIST_AddressTypes(intType)
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
            Return LANG_Address
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Addresses
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Me.FullAddress
    End Function
End Class

Public Class Addresses
    Inherits System.Collections.CollectionBase

    Public Event AddressAdded(ByVal sender As Object, ByVal e As AddressAddedEventArgs)

    Public Sub Add(ByVal address As Address)
        List.Add(address)
        RaiseEvent AddressAdded(Me, New AddressAddedEventArgs(address))
    End Sub

    Public Function Contains(ByVal address As Address) As Boolean
        Return List.Contains(address)
    End Function

    Public Function IndexOf(ByVal address As Address) As Integer
        Return List.IndexOf(address)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal address As Address)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of address is not valid, cannot be inserted!")
        ElseIf index = Count Then
            List.Add(address)
            RaiseEvent AddressAdded(Me, New AddressAddedEventArgs(address))
        Else
            List.Insert(index, address)
            RaiseEvent AddressAdded(Me, New AddressAddedEventArgs(address))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Address
        Get
            Return CType(List.Item(index), Address)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal address As Address)
        If List.Contains(address) Then List.Remove(address)
    End Sub

    Public Function GetMainAddress() As Address
        Dim selItem As Address = Nothing
        If List.Count > 0 Then
            If List.Count = 1 Then
                selItem = List.Item(0)
            Else
                For Each selItem In List
                    If selItem.Type = Address.AddressTypes.Main Then Exit For
                Next
                If selItem Is Nothing Then selItem = List.Item(0)
            End If

        Else
            selItem = Nothing
        End If
        Return selItem
    End Function

    Public Function GetMainOrWorkAddress() As Address
        Dim selItem As Address = Nothing
        If List.Count > 0 Then
            If List.Count = 1 Then
                selItem = List.Item(0)
            Else
                For Each selItem In List
                    If selItem.Type = Address.AddressTypes.Main Or selItem.Type = Address.AddressTypes.Work Then Exit For
                Next
                If selItem Is Nothing Then selItem = List.Item(0)
            End If

        Else
            selItem = Nothing
        End If
        Return selItem
    End Function

    Public Function GetAddressByGuid(ByVal objGuid As Guid) As Address
        Dim selAddress As Address = Nothing
        For Each objAddress As Address In Me.List
            If objAddress.Guid = objGuid Then
                selAddress = objAddress
                Exit For
            End If
        Next
        Return selAddress
    End Function
End Class

Public Class AddressAddedEventArgs
    Inherits EventArgs

    Public Property Address As Address

    Public Sub New(ByVal objAddress As Address)
        MyBase.New()

        Me.Address = objAddress
    End Sub
End Class
