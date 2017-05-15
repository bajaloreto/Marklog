Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class AudioVisualDetail
    Private strUrl, strDescription As String

    Public Property URL As String
        Get
            Return strUrl
        End Get
        Set(ByVal value As String)
            strUrl = value
        End Set
    End Property

    Public Property Description As String
        Get
            Return strDescription
        End Get
        Set(ByVal value As String)
            strDescription = value
        End Set
    End Property
End Class
