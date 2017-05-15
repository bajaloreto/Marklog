Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class AssumptionDetail
    Private strReason As String
    Private strHowToValidate As String
    Private boolValidated As Boolean
    Private intRiskResponse As Integer

    Public Sub New()

    End Sub

#Region "Properties"
    Public Property Reason() As String
        Get
            Return strReason
        End Get
        Set(ByVal value As String)
            strReason = value
        End Set
    End Property

    Public Property HowToValidate() As String
        Get
            Return strHowToValidate
        End Get
        Set(ByVal value As String)
            strHowToValidate = value
        End Set
    End Property

    Public Property Validated() As Boolean
        Get
            Return boolValidated
        End Get
        Set(ByVal value As Boolean)
            boolValidated = value
        End Set
    End Property
#End Region
End Class
