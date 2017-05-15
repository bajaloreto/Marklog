Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Country
    Private intIdCountry As Integer
    Private intIdGeoRegion As Integer?
    Private strNameEnglish As String
    Private strNameFrench As String
    Private strNameDutch As String
    Private strNameOfficial As String
    Private strAbreviation As String
    Private strCurrency As String
    Private strCurrencyCode As String
    Private strEmailCode As String
    Private strTelCountryCode As String

    Public Property idCountry As Integer
        Get
            Return intIdCountry
        End Get
        Set(ByVal value As Integer)
            intIdCountry = value
        End Set
    End Property

    Public Property idGeoRegion As System.Nullable(Of Integer)
        Get
            Return intIdGeoRegion
        End Get
        Set(ByVal value As System.Nullable(Of Integer))
            intIdGeoRegion = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property CountryName As String
        Get
            Dim strName As String = String.Empty
            Select Case UserLanguage
                Case "fr"
                    strName = NameFrench
                Case Else
                    strName = NameEnglish
            End Select
            Return strName
        End Get
    End Property

    Public Property NameEnglish() As String
        Get
            Return strNameEnglish
        End Get
        Set(ByVal value As String)
            strNameEnglish = value
        End Set
    End Property

    Public Property NameFrench() As String
        Get
            Return strNameFrench
        End Get
        Set(ByVal value As String)
            strNameFrench = value
        End Set
    End Property

    'Public Property NameDutch() As String
    '    Get
    '        Return strNameDutch
    '    End Get
    '    Set(ByVal value As String)
    '        strNameDutch = value
    '    End Set
    'End Property

    Public Property NameOfficial() As String
        Get
            Return strNameOfficial
        End Get
        Set(ByVal value As String)
            strNameOfficial = value
        End Set
    End Property

    Public Property Abreviation() As String
        Get
            Return strAbreviation
        End Get
        Set(ByVal value As String)
            strAbreviation = value
        End Set
    End Property

    Public Property Currency() As String
        Get
            Return strCurrency
        End Get
        Set(ByVal value As String)
            strCurrency = value
        End Set
    End Property

    Public Property CurrencyCode() As String
        Get
            Return strCurrencyCode
        End Get
        Set(ByVal value As String)
            strCurrencyCode = value
        End Set
    End Property

    Public Property EmailCode() As String
        Get
            Return strEmailCode
        End Get
        Set(ByVal value As String)
            strEmailCode = value
        End Set
    End Property

    Public Property TelCountryCode() As String
        Get
            Return strTelCountryCode
        End Get
        Set(ByVal value As String)
            strTelCountryCode = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Country
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Countries
        End Get
    End Property
End Class

Public Class Countries
    Inherits System.Collections.CollectionBase

    Public Sub Add(ByVal country As Country)
        List.Add(country)
    End Sub

    Public Function Contains(ByVal country As Country) As Boolean
        Return List.Contains(country)
    End Function

    Public Function IndexOf(ByVal country As Country) As Integer
        Return List.IndexOf(country)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal country As Country)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, country.ItemName))
        ElseIf index = Count Then
            List.Add(country)
        Else
            List.Insert(index, country)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Country
        Get
            Return CType(List.Item(index), Country)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, Country.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal country As Country)
        If List.Contains(country) Then List.Remove(country)
    End Sub

    Public Sub SortByCountryName()
        Dim sorter As System.Collections.IComparer = New CountryComparer()
        InnerList.Sort(sorter)
    End Sub

    Public Function GetCountryByIndex(ByVal intIndex As Integer) As Country
        Dim objCountry As Country = Nothing
        For Each selCountry As Country In Me.List
            If selCountry.idCountry = intIndex Then
                objCountry = selCountry
                Exit For
            End If
        Next
        Return objCountry
    End Function
End Class

Public Class CountryComparer
    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        Return x.CountryName.CompareTo(y.CountryName)
    End Function
End Class