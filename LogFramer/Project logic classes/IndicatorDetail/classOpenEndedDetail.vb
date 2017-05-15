Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class OpenEndedDetail
    Private intWhiteSpace As Integer

    Public Enum WhiteSpaceUnits
        OneLine
        TwoLines
        ThreeLines
        FourLines
        FiveLines
        SixLines
        QuarterPage
        ThirdPage
        HalfPage
        ThreeQuarters
        Page
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal whitespace As Integer)
        Me.WhiteSpace = whitespace
    End Sub

#Region "Properties"
    Public Property WhiteSpace() As Integer
        Get
            Return intWhiteSpace
        End Get
        Set(ByVal value As Integer)
            intWhiteSpace = value
        End Set
    End Property
#End Region

#Region "Methods"

#End Region
End Class
