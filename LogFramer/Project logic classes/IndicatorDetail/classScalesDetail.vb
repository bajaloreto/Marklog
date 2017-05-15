Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ScalesDetail
    Private strAgreeText, strDisagreeText As String

    Public Property AgreeText As String
        Get
            Return strAgreeText
        End Get
        Set(ByVal value As String)
            strAgreeText = value
        End Set
    End Property

    Public Property DisagreeText As String
        Get
            Return strDisagreeText
        End Get
        Set(ByVal value As String)
            strDisagreeText = value
        End Set
    End Property
End Class
