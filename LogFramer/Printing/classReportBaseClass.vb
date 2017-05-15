Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class ReportBaseClass
    Inherits PrintDocument

    Friend WithEvents RichTextManager As New RichTextManager

    Private objReportSetup As New ReportSetup
    Private objPageGraph As Graphics
    Private objPrintRectangles As New PrintRectangles

    Private intPaperWidth As Integer, intPaperHeight As Integer
    Private intPrintWidth As Integer, intPrintHeight As Integer
    Private intLeftMargin As Integer, intTopMargin As Integer
    Private intHeaderMargin As Integer = 20
    Private intFooterMargin As Integer = 20

    Private intLastRowY As Integer
    Private intRowIndex As Integer
    Private intPageNumber As Integer
    Private intTotalPages As Integer

    Public penBlack1 As Pen = New Pen(Color.Black, 1)
    Public penBlack2 As Pen = New Pen(Color.Black, 2)
    Public fntText As New Font(CurrentLogFrame.DetailsFont.FontFamily, CurrentLogFrame.DetailsFont.SizeInPoints - 2)
    Public fntTextBold As New Font(fntText, FontStyle.Bold)

#Region "Properties"
    Public Property ReportSetup As ReportSetup
        Get
            Return objReportSetup
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetup = value
            SetMargins()
        End Set
    End Property

    Public Property PrintRectangles As PrintRectangles
        Get
            Return objPrintRectangles
        End Get
        Set(ByVal value As PrintRectangles)
            objPrintRectangles = value
        End Set
    End Property

    Public Property PageGraph As Graphics
        Get
            Return objPageGraph
        End Get
        Set(ByVal value As Graphics)
            objPageGraph = value

            With objPageGraph
                .CompositingQuality = CompositingQuality.HighQuality
                .InterpolationMode = InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = PixelOffsetMode.HighQuality
                .SmoothingMode = SmoothingMode.HighQuality

                .TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
            End With
        End Set
    End Property

    Public Property PaperWidth As Integer
        Get
            Return intPaperWidth
        End Get
        Set(ByVal value As Integer)
            intPaperWidth = value
        End Set
    End Property

    Public Property PaperHeight As Integer
        Get
            Return intPaperHeight
        End Get
        Set(ByVal value As Integer)
            intPaperHeight = value
        End Set
    End Property

    Public Property LeftMargin As Integer
        Get
            Return intLeftMargin
        End Get
        Set(ByVal value As Integer)
            intLeftMargin = value
        End Set
    End Property

    Public Property TopMargin As Integer
        Get
            Return intTopMargin
        End Get
        Set(ByVal value As Integer)
            intTopMargin = value
        End Set
    End Property

    Public Property PrintWidth As Integer
        Get
            Return intPrintWidth
        End Get
        Set(ByVal value As Integer)
            intPrintWidth = value
        End Set
    End Property

    Public Property PrintHeight As Integer
        Get
            Return intPrintHeight
        End Get
        Set(ByVal value As Integer)
            intPrintHeight = value
        End Set
    End Property

    Public ReadOnly Property ContentLeft As Integer
        Get
            Return intLeftMargin
        End Get
    End Property

    Public ReadOnly Property ContentRight As Integer
        Get
            Return intLeftMargin + PrintWidth
        End Get
    End Property

    Public ReadOnly Property ContentTop As Integer
        Get
            Return intTopMargin + HeaderHeight
        End Get
    End Property

    Public ReadOnly Property ContentBottom As Integer
        Get
            Return ContentTop + ContentHeight
        End Get
    End Property

    Public ReadOnly Property ContentWidth As Integer
        Get
            Return PrintWidth
        End Get
    End Property

    Public ReadOnly Property ContentHeight As Integer
        Get
            Dim intHeight As Integer
            intHeight = Me.PrintHeight - Me.HeaderHeight - Me.FooterHeight - intHeaderMargin - intFooterMargin
            Return intHeight
        End Get
    End Property

    Public ReadOnly Property HeaderHeight As Integer
        Get
            Dim intHeight As Integer = Me.LeftHeaderHeight
            If Me.MiddleHeaderHeight > intHeight Then intHeight = Me.MiddleHeaderHeight
            If Me.RightHeaderHeight > intHeight Then intHeight = Me.RightHeaderHeight

            intHeight += HeaderMargin

            Return intHeight
        End Get
    End Property

    Public Property HeaderMargin As Integer
        Get
            Return intHeaderMargin
        End Get
        Set(ByVal value As Integer)
            intHeaderMargin = value
        End Set
    End Property

    Public ReadOnly Property LeftHeaderWidth As Integer
        Get
            Dim strLeftHeader As String = RichTextToText(ReportSetup.HeaderLeft)
            Dim strMiddleHeader As String = RichTextToText(ReportSetup.HeaderMiddle)
            Dim strRightHeader As String = RichTextToText(ReportSetup.HeaderRight)


            If String.IsNullOrEmpty(strMiddleHeader) Then
                If String.IsNullOrEmpty(strLeftHeader) And String.IsNullOrEmpty(strRightHeader) = False Then
                    Return 0
                ElseIf String.IsNullOrEmpty(strLeftHeader) = False And String.IsNullOrEmpty(strRightHeader) Then
                    Return PrintWidth
                Else
                    Return PrintWidth / 2
                End If
            Else
                If String.IsNullOrEmpty(strLeftHeader) And String.IsNullOrEmpty(strRightHeader) Then
                    Return 0
                Else
                    Return PrintWidth / 3
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property LeftHeaderHeight As Integer
        Get
            Dim strLeftHeaderRtf As String = ReportSetup.HeaderLeft
            Dim strLeftHeader As String = RichTextToText(ReportSetup.HeaderLeft)

            If String.IsNullOrEmpty(strLeftHeader) = False Then
                RichTextManager.Width = LeftHeaderWidth
                RichTextManager.Rtf = strLeftHeaderRtf
                Dim intHeight As Integer = RichTextManager.Height

                Return intHeight
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property MiddleHeaderWidth As Integer
        Get
            Dim strLeftHeader As String = RichTextToText(ReportSetup.HeaderLeft)
            Dim strMiddleHeader As String = RichTextToText(ReportSetup.HeaderMiddle)
            Dim strRightHeader As String = RichTextToText(ReportSetup.HeaderRight)


            If String.IsNullOrEmpty(strMiddleHeader) Then
                Return 0
            Else
                If String.IsNullOrEmpty(strLeftHeader) And String.IsNullOrEmpty(strRightHeader) Then
                    Return PrintWidth
                Else
                    Return PrintWidth / 3
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property MiddleHeaderHeight As Integer
        Get
            Dim strMiddleHeaderRtf As String = ReportSetup.HeaderMiddle
            Dim strMiddleHeader As String = RichTextToText(ReportSetup.HeaderMiddle)

            If String.IsNullOrEmpty(strMiddleHeader) = False Then
                RichTextManager.Width = MiddleHeaderWidth
                RichTextManager.Rtf = strMiddleHeaderRtf
                Dim intHeight As Integer = RichTextManager.Height

                Return intHeight
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property RightHeaderWidth As Integer
        Get
            Dim strLeftHeader As String = RichTextToText(ReportSetup.HeaderLeft)
            Dim strMiddleHeader As String = RichTextToText(ReportSetup.HeaderMiddle)
            Dim strRightHeader As String = RichTextToText(ReportSetup.HeaderRight)


            If String.IsNullOrEmpty(strMiddleHeader) Then
                If String.IsNullOrEmpty(strLeftHeader) And String.IsNullOrEmpty(strRightHeader) = False Then
                    Return PrintWidth
                ElseIf String.IsNullOrEmpty(strLeftHeader) = False And String.IsNullOrEmpty(strRightHeader) Then
                    Return 0
                Else
                    Return PrintWidth / 2
                End If
            Else
                If String.IsNullOrEmpty(strLeftHeader) And String.IsNullOrEmpty(strRightHeader) Then
                    Return 0
                Else
                    Return PrintWidth / 3
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property RightHeaderHeight As Integer
        Get
            Dim strRightHeaderRtf As String = ReportSetup.HeaderRight
            Dim strRightHeader As String = RichTextToText(ReportSetup.HeaderRight)

            If String.IsNullOrEmpty(strRightHeader) = False Then
                RichTextManager.Width = RightHeaderWidth
                RichTextManager.Rtf = strRightHeaderRtf
                Dim intHeight As Integer = RichTextManager.Height

                Return intHeight
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property FooterHeight As Integer
        Get
            Dim intHeight As Integer = Me.LeftFooterHeight
            If Me.MiddleFooterHeight > intHeight Then intHeight = Me.MiddleFooterHeight
            If Me.RightFooterHeight > intHeight Then intHeight = Me.RightFooterHeight

            Return intHeight
        End Get
    End Property

    Public Property FooterMargin As Integer
        Get
            Return intFooterMargin
        End Get
        Set(ByVal value As Integer)
            intFooterMargin = value
        End Set
    End Property

    Public ReadOnly Property LeftFooterWidth As Integer
        Get
            Dim strLeftFooter As String = RichTextToText(ReportSetup.FooterLeft)
            Dim strMiddleFooter As String = RichTextToText(ReportSetup.FooterMiddle)
            Dim strRightFooter As String = RichTextToText(ReportSetup.FooterRight)


            If String.IsNullOrEmpty(strMiddleFooter) Then
                If String.IsNullOrEmpty(strLeftFooter) And String.IsNullOrEmpty(strRightFooter) = False Then
                    Return 0
                ElseIf String.IsNullOrEmpty(strLeftFooter) = False And String.IsNullOrEmpty(strRightFooter) Then
                    Return PrintWidth
                Else
                    Return PrintWidth / 2
                End If
            Else
                If String.IsNullOrEmpty(strLeftFooter) And String.IsNullOrEmpty(strRightFooter) Then
                    Return 0
                Else
                    Return PrintWidth / 3
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property LeftFooterHeight As Integer
        Get
            Dim strLeftFooterRtf As String = ReportSetup.FooterLeft
            Dim strLeftFooter As String = RichTextToText(ReportSetup.FooterLeft)

            If String.IsNullOrEmpty(strLeftFooter) = False Then
                RichTextManager.Width = LeftFooterWidth
                RichTextManager.Rtf = strLeftFooterRtf
                Dim intHeight As Integer = RichTextManager.Height

                Return intHeight
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property MiddleFooterWidth As Integer
        Get
            Dim strLeftFooter As String = RichTextToText(ReportSetup.FooterLeft)
            Dim strMiddleFooter As String = RichTextToText(ReportSetup.FooterMiddle)
            Dim strRightFooter As String = RichTextToText(ReportSetup.FooterRight)


            If String.IsNullOrEmpty(strMiddleFooter) Then
                Return 0
            Else
                If String.IsNullOrEmpty(strLeftFooter) And String.IsNullOrEmpty(strRightFooter) Then
                    Return PrintWidth
                Else
                    Return PrintWidth / 3
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property MiddleFooterHeight As Integer
        Get
            Dim strMiddleFooterRtf As String = ReportSetup.FooterMiddle
            Dim strMiddleFooter As String = RichTextToText(ReportSetup.FooterMiddle)

            If String.IsNullOrEmpty(strMiddleFooter) = False Then
                RichTextManager.Width = MiddleFooterWidth
                RichTextManager.Rtf = strMiddleFooterRtf
                Dim intHeight As Integer = RichTextManager.Height

                Return intHeight
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property RightFooterWidth As Integer
        Get
            Dim strLeftFooter As String = RichTextToText(ReportSetup.FooterLeft)
            Dim strMiddleFooter As String = RichTextToText(ReportSetup.FooterMiddle)
            Dim strRightFooter As String = RichTextToText(ReportSetup.FooterRight)


            If String.IsNullOrEmpty(strMiddleFooter) Then
                If String.IsNullOrEmpty(strLeftFooter) And String.IsNullOrEmpty(strRightFooter) = False Then
                    Return PrintWidth
                ElseIf String.IsNullOrEmpty(strLeftFooter) = False And String.IsNullOrEmpty(strRightFooter) Then
                    Return 0
                Else
                    Return PrintWidth / 2
                End If
            Else
                If String.IsNullOrEmpty(strLeftFooter) And String.IsNullOrEmpty(strRightFooter) Then
                    Return 0
                Else
                    Return PrintWidth / 3
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property RightFooterHeight As Integer
        Get
            Dim strRightFooterRtf As String = ReportSetup.FooterRight
            Dim strRightFooter As String = RichTextToText(ReportSetup.FooterRight)

            If String.IsNullOrEmpty(strRightFooter) = False Then
                RichTextManager.Width = RightFooterWidth
                RichTextManager.Rtf = strRightFooterRtf
                Dim intHeight As Integer = RichTextManager.Height

                Return intHeight
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property RowIndex As Integer
        Get
            Return intRowIndex
        End Get
        Set(ByVal value As Integer)
            intRowIndex = value
        End Set
    End Property

    Public Property LastRowY As Integer
        Get
            Return intLastRowY
        End Get
        Set(ByVal value As Integer)
            intLastRowY = value
        End Set
    End Property

    Public Property PageNumber() As Integer
        Get
            Return intPageNumber
        End Get
        Set(ByVal value As Integer)
            intPageNumber = value
            If TotalPages < value Then TotalPages = value
        End Set
    End Property

    Public Property TotalPages() As Integer
        Get
            Return intTotalPages
        End Get
        Set(ByVal value As Integer)
            intTotalPages = value
        End Set
    End Property
#End Region

#Region "Print page"
    Private Sub SetMargins()
        With Me.DefaultPageSettings
            .Color = ReportSetup.PrintInColor
            .Landscape = ReportSetup.PrintAsLandScape
            .Margins = ReportSetup.PrintMargins
        End With

        With Me.DefaultPageSettings
            If .Landscape = False Then
                PaperWidth = .PaperSize.Width
                PaperHeight = .PaperSize.Height
                LeftMargin = .Margins.Left
                TopMargin = .Margins.Top
            Else
                PaperWidth = .PaperSize.Height
                PaperHeight = .PaperSize.Width
                LeftMargin = .Margins.Top
                TopMargin = .Margins.Right
            End If
            PrintWidth = PaperWidth - .Margins.Left - .Margins.Right
            PrintHeight = PaperHeight - .Margins.Top - .Margins.Bottom
        End With
    End Sub

    Protected Sub PrintHeader()
        Dim strText As String
        Dim strLeftHeader As String = RichTextToText(ReportSetup.HeaderLeft)
        Dim strMiddleHeader As String = RichTextToText(ReportSetup.HeaderMiddle)
        Dim strRightHeader As String = RichTextToText(ReportSetup.HeaderRight)
        Dim x, y As Integer
        Dim intBottom As Integer
        Dim rCell As Rectangle
        Dim bmHeader As Bitmap

        y = TopMargin
        intBottom = TopMargin + HeaderHeight

        If PageGraph IsNot Nothing Then
            With RichTextManager
                If String.IsNullOrEmpty(strLeftHeader) = False Then
                    x = LeftMargin
                    rCell = New Rectangle(x, y, Me.LeftHeaderWidth, Me.HeaderHeight)
                    strText = InsertPageNumber(ReportSetup.HeaderLeft)
                    bmHeader = .RichTextWithPaddingToBitmap(Me.LeftHeaderWidth, strText, False)
                    PageGraph.DrawImage(bmHeader, x, y)
                End If
                If String.IsNullOrEmpty(strMiddleHeader) = False Then
                    x = LeftMargin + LeftHeaderWidth
                    rCell = New Rectangle(x, y, Me.MiddleHeaderWidth, Me.HeaderHeight)
                    strText = InsertPageNumber(ReportSetup.HeaderMiddle)
                    bmHeader = .RichTextWithPaddingToBitmap(Me.MiddleHeaderWidth, strText, False, 0, HorizontalAlignment.Center)
                    PageGraph.DrawImage(bmHeader, x, y)
                End If
                If String.IsNullOrEmpty(strRightHeader) = False Then
                    x = LeftMargin + LeftHeaderWidth + MiddleHeaderWidth
                    rCell = New Rectangle(x, y, Me.RightHeaderWidth, Me.HeaderHeight)
                    strText = InsertPageNumber(ReportSetup.HeaderRight)
                    bmHeader = .RichTextWithPaddingToBitmap(Me.RightHeaderWidth, strText, False, 0, HorizontalAlignment.Right)
                    PageGraph.DrawImage(bmHeader, x, y)
                End If
            End With
        End If
    End Sub

    Protected Sub PrintFooter()
        Dim strRtf As String
        Dim strLeftFooter As String = RichTextToText(ReportSetup.FooterLeft)
        Dim strMiddleFooter As String = RichTextToText(ReportSetup.FooterMiddle)
        Dim strRightFooter As String = RichTextToText(ReportSetup.FooterRight)
        Dim x, y As Integer
        Dim intBottom As Integer
        Dim bmHeader As Bitmap

        y = TopMargin + PrintHeight - FooterHeight
        intBottom = TopMargin + PrintHeight

        If PageGraph IsNot Nothing Then
            With RichTextManager
                If String.IsNullOrEmpty(strLeftFooter) = False Then
                    x = LeftMargin
                    strRtf = InsertPageNumber(ReportSetup.FooterLeft)
                    bmHeader = .RichTextWithPaddingToBitmap(Me.LeftFooterWidth, strRtf, False)
                    PageGraph.DrawImage(bmHeader, x, y)
                End If
                If String.IsNullOrEmpty(strMiddleFooter) = False Then
                    x = LeftMargin + LeftFooterWidth
                    strRtf = InsertPageNumber(ReportSetup.FooterMiddle)
                    bmHeader = .RichTextWithPaddingToBitmap(Me.MiddleFooterWidth, strRtf, False, 0, HorizontalAlignment.Center)
                    PageGraph.DrawImage(bmHeader, x, y)
                End If
                If String.IsNullOrEmpty(strRightFooter) = False Then
                    x = LeftMargin + LeftFooterWidth + MiddleFooterWidth
                    strRtf = InsertPageNumber(ReportSetup.FooterRight)
                    bmHeader = .RichTextWithPaddingToBitmap(Me.RightFooterWidth, strRtf, False, 0, HorizontalAlignment.Right)
                    PageGraph.DrawImage(bmHeader, x, y)
                End If
            End With
        End If
    End Sub

    Protected Function InsertPageNumber(ByVal strRtf As String) As String
        With RichTextManager
            .Rtf = strRtf
            If .Text.Contains("{&Page}") Then
                .Text = .Text.Replace("{&Page}", PageNumber.ToString)
            End If
            strRtf = .Rtf
        End With

        strRtf = InsertTotalPages(strRtf)
        Return strRtf
    End Function

    Protected Function InsertTotalPages(ByVal strRtf As String) As String
        With RichTextManager
            .Rtf = strRtf
            If .Text.Contains("{&Pages}") Then
                .Text = .Text.Replace("{&Pages}", TotalPages.ToString)
            End If
            strRtf = .Rtf
        End With
        Return strRtf
    End Function

    Protected Overrides Sub OnBeginPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnBeginPrint(e)

        RowIndex = 0
        LastRowY = 0
        PageNumber = 1
        SetMargins()
    End Sub

    Protected Overrides Sub OnQueryPageSettings(ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs)
        MyBase.OnQueryPageSettings(e)
    End Sub

    Protected Overrides Sub OnEndPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnEndPrint(e)
    End Sub

    Protected Overrides Sub OnPrintPage(ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        MyBase.OnPrintPage(e)
    End Sub

#End Region 'Print page

#Region "Other methods"
    Public Function GetTextRectangle(ByVal rCell As Rectangle) As Rectangle
        Dim rText As Rectangle = Nothing

        If rCell <> Nothing Then _
            rText = New Rectangle(rCell.X + CONST_Padding, rCell.Y + CONST_Padding, rCell.Width - CONST_HorizontalPadding, rCell.Height - CONST_VerticalPadding)

        Return rText
    End Function
#End Region
End Class

Public Class LinePrintedEventArgs
    Inherits EventArgs

    Public Property LineIndex As Integer
    Public Property RowCount As Integer

    Public Sub New(ByVal intIndex As Integer, ByVal intRowCount As Integer)
        MyBase.New()

        Me.LineIndex = intIndex
        Me.RowCount = intRowCount
    End Sub
End Class

Public Class MinimumPageWidthChangedEventArgs
    Inherits EventArgs

    Public Property PageWidth As Integer

    Public Sub New(ByVal intPageWidth As Integer)
        MyBase.New()

        Me.PageWidth = intPageWidth
    End Sub
End Class
