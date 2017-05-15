Imports System.Xml.Serialization
Imports System.Drawing.Printing

Public Class ReportSetup
    Private objPageSettings As New PageSettings
    Private boolPrintInColor As Boolean = True
    Private boolPrintAsLandScape As Boolean = False
    Private objPrintMargins As New Margins
    Private objPaperSize As New PaperSize
    Private strHeaderLeft As String = String.Empty, strHeaderMiddle As String = String.Empty, strHeaderRight As String = String.Empty
    Private strFooterLeft As String = String.Empty, strFooterMiddle As String = String.Empty, strFooterRight As String = String.Empty

    Public Property PrintInColor() As Boolean
        Get
            Return boolPrintInColor
        End Get
        Set(ByVal value As Boolean)
            boolPrintInColor = value
        End Set
    End Property

    Public Property PrintAsLandScape() As Boolean
        Get
            Return boolPrintAsLandScape
        End Get
        Set(ByVal value As Boolean)
            boolPrintAsLandScape = value
        End Set
    End Property

    Public Property PrintMargins() As Margins
        Get
            Return objPrintMargins
        End Get
        Set(ByVal value As Margins)
            objPrintMargins = value
        End Set
    End Property

    Public Property HeaderLeft As String
        Get
            Return strHeaderLeft
        End Get
        Set(ByVal value As String)
            strHeaderLeft = value
        End Set
    End Property

    Public Property HeaderMiddle As String
        Get
            Return strHeaderMiddle
        End Get
        Set(ByVal value As String)
            strHeaderMiddle = value
        End Set
    End Property

    Public Property HeaderRight As String
        Get
            Return strHeaderRight
        End Get
        Set(ByVal value As String)
            strHeaderRight = value
        End Set
    End Property

    Public Property FooterLeft As String
        Get
            Return strFooterLeft
        End Get
        Set(ByVal value As String)
            strFooterLeft = value
        End Set
    End Property

    Public Property FooterMiddle As String
        Get
            Return strFooterMiddle
        End Get
        Set(ByVal value As String)
            strFooterMiddle = value
        End Set
    End Property

    Public Property FooterRight As String
        Get
            Return strFooterRight
        End Get
        Set(ByVal value As String)
            strFooterRight = value
        End Set
    End Property
End Class
