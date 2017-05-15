Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public MustInherit Class LogframeObject
    Protected Event GuidChanged(ByVal sender As Object)

    Private intSection As Integer
    Private strTextValue As String
    Private strRtfValue As String
    Private objGuid As Guid
    Private bmCellImage As System.Drawing.Bitmap

    Public Enum SectionTypes As Integer
        Goal = 1
        Purpose = 2
        Output = 3
        Activity = 4
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Public Sub New(ByVal Section As Integer, ByVal RTF As String)
        Me.Section = Section
        Me.RTF = RTF
    End Sub

#Region "Properties"
    Public Overridable Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property Text() As String
        Get
            Dim strText As String = String.Empty
            If String.IsNullOrEmpty(Me.RTF) = False Then
                Using objRtfDraw As New RichTextManager
                    objRtfDraw.Rtf = Me.RTF

                    strText = objRtfDraw.Text
                End Using
            End If
            Return strText
        End Get
    End Property

    Public Property RTF() As String
        Get
            Return strRtfValue
        End Get
        Set(ByVal value As String)
            strRtfValue = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
                RaiseEvent GuidChanged(Me)
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
            RaiseEvent GuidChanged(Me)
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Property CellImage() As Bitmap
        Get
            Return bmCellImage
        End Get
        Set(ByVal value As Bitmap)
            bmCellImage = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SetText(ByVal strText As String)
        Using objRtfDraw As New RichTextManager
            objRtfDraw.Text = strText
            Me.RTF = objRtfDraw.Rtf
        End Using
    End Sub

    Public Function GetFont() As Font
        Dim selFont As Font = CurrentLogFrame.DetailsFont
        Using objRtfPainter As New RichTextManager
            With objRtfPainter
                .Rtf = Me.RTF
                .Select(0, 0)
                selFont = .SelectionFont
            End With
        End Using
        Return selFont
    End Function

    Public Function GetFontColor() As Color
        Dim selFontColor As Color = Color.Black
        Using objRtfPainter As New RichTextManager
            With objRtfPainter
                .Rtf = Me.RTF
                .Select(0, 0)
                selFontColor = .SelectionColor
            End With
        End Using
        Return selFontColor
    End Function

    Public Overrides Function ToString() As String
        Return Me.Text
    End Function

    Protected MustOverride Sub OnGuidChanged() Handles Me.GuidChanged

#End Region
End Class
