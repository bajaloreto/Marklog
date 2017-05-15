Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintBudget
    Inherits ReportBaseClass

    Private Const CONST_MinTextColumnWidth As Integer = 100
    Private Const CONST_BudgetYearHeaderHeight As Integer = 40
    Private Const CONST_TotalBudgetRowPadding As Integer = 5

    Private objLogFrame As New LogFrame
    Private objGuid As Guid
    Private objPrintList As New PrintBudgetRows
    Private boolTotalBudget As Boolean
    Private boolColumnsWidthSet As Boolean
    Private intBudgetYearIndex As Integer
    Private boolShowDurationColumns, boolShowLocalCurrencyColumns, boolShowExchangeRates As Boolean
    Private intPagesWide As Integer

    Private CurrentSectionRectangle As New Rectangle
    Private rSortColumnRectangle, rTextColumnRectangle As New PrintRectangle
    Private rDurationColumnRectangle, rNumberColumnRectangle As New PrintRectangle
    Private rUnitCostColumnRectangle, rTotalLocalCostColumnRectangle As New PrintRectangle
    Private rTotalCostColumnRectangle As New PrintRectangle
    Private intStartIndex As Integer
    Private ColumnHeadersHeight As Integer

    Private intHorPages As Integer, intCurrentHorPage As Integer = 1
    Private intSectionBorder As Integer
    Private intLastColumnBorder As Integer
    Private intHorizontalPageIndex As Integer

    Private colForeGround As Color = Color.Black, colBackGround As Color = Color.White
    Private colBaseGrey As Color = Color.FromArgb(255, 50, 50, 50)

    Public Event PagePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
    Public Event MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)

    Public Sub New(ByVal logframe As LogFrame, ByVal budgetyearindex As Integer, ByVal showdurationcolumns As Boolean, ByVal showlocalcurrencycolumns As Boolean, _
                   ByVal showexchangerates As Boolean, ByVal pageswide As Integer)
        Me.LogFrame = logframe
        Me.BudgetYearIndex = budgetyearindex
        Me.ShowDurationColumns = showdurationcolumns
        Me.ShowLocalCurrencyColumns = showlocalcurrencycolumns
        Me.ShowExchangeRates = showexchangerates
        Me.PagesWide = pageswide

        Me.ReportSetup = logframe.ReportSetupBudget
    End Sub

#Region "Properties"
    Public Property LogFrame() As LogFrame
        Get
            Return objLogFrame
        End Get
        Set(ByVal value As LogFrame)
            objLogFrame = value
        End Set
    End Property

    Public Property PrintList() As PrintBudgetRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintBudgetRows)
            objPrintList = value
        End Set
    End Property

    Public Property BudgetYearIndex() As Integer
        Get
            Return intBudgetYearIndex
        End Get
        Set(ByVal value As Integer)
            intBudgetYearIndex = value
        End Set
    End Property

    Public Property TotalBudget As Boolean
        Get
            Return boolTotalBudget
        End Get
        Set(ByVal value As Boolean)
            boolTotalBudget = value
        End Set
    End Property

    Public Property ShowDurationColumns As Boolean
        Get
            Return boolShowDurationColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowDurationColumns = value
        End Set
    End Property

    Public Property ShowLocalCurrencyColumns As Boolean
        Get
            Return boolShowLocalCurrencyColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowLocalCurrencyColumns = value
        End Set
    End Property

    Public Property ShowExchangeRates As Boolean
        Get
            Return boolShowExchangeRates
        End Get
        Set(ByVal value As Boolean)
            boolShowExchangeRates = value
        End Set
    End Property

    Public Property PagesWide As Integer
        Get
            Return intPagesWide
        End Get
        Set(value As Integer)
            intPagesWide = value
        End Set
    End Property

    Private Property SortColumnRectangle As PrintRectangle
        Get
            Return rSortColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rSortColumnRectangle = value
        End Set
    End Property

    Private Property TextColumnRectangle As PrintRectangle
        Get
            Return rTextColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rTextColumnRectangle = value
        End Set
    End Property

    Private Property DurationColumnRectangle As PrintRectangle
        Get
            Return rDurationColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rDurationColumnRectangle = value
        End Set
    End Property

    Private Property NumberColumnRectangle As PrintRectangle
        Get
            Return rNumberColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rNumberColumnRectangle = value
        End Set
    End Property

    Private Property UnitCostColumnRectangle As PrintRectangle
        Get
            Return rUnitCostColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rUnitCostColumnRectangle = value
        End Set
    End Property

    Private Property TotalLocalCostColumnRectangle As PrintRectangle
        Get
            Return rTotalLocalCostColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rTotalLocalCostColumnRectangle = value
        End Set
    End Property

    Private Property TotalCostColumnRectangle As PrintRectangle
        Get
            Return rTotalCostColumnRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rTotalCostColumnRectangle = value
        End Set
    End Property

    Private Property LastColumnBorder As Integer
        Get
            Return intLastColumnBorder
        End Get
        Set(ByVal value As Integer)
            intLastColumnBorder = value
        End Set
    End Property

    Private Property SectionBorder As Integer
        Get
            Return intSectionBorder
        End Get
        Set(ByVal value As Integer)
            intSectionBorder = value
        End Set
    End Property

    Private Property HorPages As Integer
        Get
            Return intHorPages
        End Get
        Set(ByVal value As Integer)
            intHorPages = value
        End Set
    End Property

    Private Property CurrentHorPage As Integer
        Get
            Return intCurrentHorPage
        End Get
        Set(ByVal value As Integer)
            intCurrentHorPage = value
        End Set
    End Property
#End Region

#Region "Create Budget table"
    Public Sub CreateList()
        PrintList.Clear()

        If Me.LogFrame IsNot Nothing Then
            If BudgetYearIndex = -1 Then
                Dim intBudgetIndex As Integer
                Dim strBudgetTitle As String

                For Each selBudgetYear As BudgetYear In Me.LogFrame.Budget.BudgetYears
                    If intBudgetIndex = 0 Then
                        strBudgetTitle = LANG_TotalBudget
                    Else
                        strBudgetTitle = selBudgetYear.BudgetYear.ToString("yyyy")
                    End If

                    Dim selBudgetItem As New BudgetItem(strBudgetTitle)
                    Dim NewGridItem As New PrintBudgetRow(selBudgetItem)

                    With NewGridItem
                        .Indent = 0
                        .RowType = PrintBudgetRow.RowTypes.TitleRow
                    End With

                    Me.PrintList.Add(NewGridItem)

                    If intBudgetIndex = 0 Then
                        CreateList_BudgetItems(selBudgetYear.BudgetItems, 0, String.Empty, True)
                    Else
                        CreateList_BudgetItems(selBudgetYear.BudgetItems, 0, String.Empty, False)
                    End If
                    CreateList_TotalBudget(selBudgetYear)

                    intBudgetIndex += 1
                Next
            Else
                Dim selBudgetYear As BudgetYear = Me.LogFrame.Budget.BudgetYears(BudgetYearIndex)

                If Me.LogFrame.Budget.MultiYearBudget And BudgetYearIndex = 0 Then Me.TotalBudget = True
                If selBudgetYear IsNot Nothing Then
                    CreateList_BudgetItems(selBudgetYear.BudgetItems, 0, String.Empty, False)
                    CreateList_TotalBudget(selBudgetYear)
                End If

                End If

                If ShowExchangeRates = True Then CreateList_ExchangeRates()
        End If
    End Sub

    Private Sub CreateList_BudgetItems(ByVal selBudgetItems As BudgetItems, ByVal intLevel As Integer, ByVal strParentSort As String, ByVal boolTotal As Boolean)
        Dim boolLastHasSubBudgetItems As Boolean
        Dim strBudgetSort As String = String.Empty
        Dim intIndex As Integer

        For Each selBudgetItem As BudgetItem In selBudgetItems
            Dim NewGridItem As New PrintBudgetRow(selBudgetItem)
            boolLastHasSubBudgetItems = False
            intIndex = selBudgetItems.IndexOf(selBudgetItem)
            strBudgetSort = CurrentLogFrame.CreateSortNumber(intIndex, strParentSort)

            With NewGridItem
                .SortNumber = strBudgetSort
                .Indent = intLevel
                .BudgetItem = selBudgetItem

                If boolTotal = True Then .RowType = PrintBudgetRow.RowTypes.TotalBudget
            End With

            Me.PrintList.Add(NewGridItem)

            If selBudgetItem.BudgetItems.Count > 0 Then
                CreateList_BudgetItems(selBudgetItem.BudgetItems, intLevel + 1, strBudgetSort, boolTotal)

                boolLastHasSubBudgetItems = True
            End If
        Next
    End Sub

    Private Sub CreateList_ExchangeRates()
        Dim selBudgetItem As New BudgetItem(LANG_ExchangeRates)
        Dim TitleRow As New PrintBudgetRow(selBudgetItem)

        TitleRow.RowType = PrintBudgetRow.RowTypes.TitleRow

        Me.PrintList.Add(TitleRow)

        Dim ColumnHeaderRow As New PrintBudgetRow()

        With ColumnHeaderRow
            .RowType = PrintBudgetRow.RowTypes.ExchangeRate
            .CurrencyText = LANG_Currency
            .ExchangeRateText = LANG_ExchangeRate
            .Conversion1 = LANG_Conversion
            .Conversion2 = LANG_Conversion
        End With

        Me.PrintList.Add(ColumnHeaderRow)

        For Each selExchangeRate As ExchangeRate In CurrentLogFrame.Budget.ExchangeRates
            Dim NewGridItem As New PrintBudgetRow(selExchangeRate, CurrentLogFrame.Budget.DefaultCurrencyCode)

            Me.PrintList.Add(NewGridItem)
        Next
    End Sub

    Private Sub CreateList_TotalBudget(ByVal selBudgetYear As BudgetYear)
        Dim strLabel As String = TextToRichText(LANG_TotalBudget, True)
        Dim TotalBudgetItem As New BudgetItem(strLabel)

        TotalBudgetItem.TotalCost = selBudgetYear.TotalCost

        Dim TotalBudgetRow As New PrintBudgetRow(TotalBudgetItem)

        With TotalBudgetRow
            .RowType = PrintBudgetRow.RowTypes.TotalBudgetRow
        End With

        Me.PrintList.Add(TotalBudgetRow)
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth * PagesWide
        Dim intSortColumnWidth, intDurationColumnWidth, intNumberColumnWidth, intUnitCostColumnWidth, intTotalLocalCostColumnWidth, intTotalCostColumnWidth As Integer
        Dim intTextColumnWidth, intNumberColumnsWidth As Integer
        Dim intX As Integer

        SectionBorder = ContentRight

        'calculate column widths
        intSortColumnWidth = GetSortColumnWidth()
        intDurationColumnWidth = GetDurationColumnWidth()
        intNumberColumnWidth = GetNumberColumnWidth()
        intUnitCostColumnWidth = GetUnitCostColumnWidth()
        intTotalLocalCostColumnWidth = GetTotalLocalCostColumnWidth()
        intTotalCostColumnWidth = GetTotalCostColumnWidth()

        intNumberColumnsWidth = intDurationColumnWidth + intNumberColumnWidth + intUnitCostColumnWidth + intTotalLocalCostColumnWidth + intTotalCostColumnWidth
        intTextColumnWidth = intAvailableWidth - intSortColumnWidth - intNumberColumnsWidth

        If intTextColumnWidth < CONST_MinTextColumnWidth Then intTextColumnWidth = CONST_MinTextColumnWidth
        If intTextColumnWidth + intSortColumnWidth > ContentWidth Then intTextColumnWidth = ContentWidth - intSortColumnWidth

        'set text column rectangles
        SortColumnRectangle = New PrintRectangle(LeftMargin, ContentTop, intSortColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(SortColumnRectangle)

        TextColumnRectangle = New PrintRectangle(SortColumnRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(TextColumnRectangle)

        If TextColumnRectangle.Left < SectionBorder And TextColumnRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(TextColumnRectangle)
        End If
        intX = TextColumnRectangle.Right

        'set duration column rectangle
        If ShowDurationColumns = True And TotalBudget = False Then
            DurationColumnRectangle = New PrintRectangle(TextColumnRectangle.Right, ContentTop, intDurationColumnWidth, ContentHeight, intHorizontalPageIndex)
            PrintRectangles.Add(DurationColumnRectangle)

            If DurationColumnRectangle.Left <= SectionBorder And DurationColumnRectangle.Right > SectionBorder Then
                StretchRectanglesOfPreviousPage(DurationColumnRectangle)
            End If
            intX = DurationColumnRectangle.Right
        End If

        'set number column rectangle
        NumberColumnRectangle = New PrintRectangle(intX, ContentTop, intNumberColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(NumberColumnRectangle)

        If NumberColumnRectangle.Left <= SectionBorder And NumberColumnRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(NumberColumnRectangle)
        End If
        intX = NumberColumnRectangle.Right

        If TotalBudget = False Then
            UnitCostColumnRectangle = New PrintRectangle(intX, ContentTop, intUnitCostColumnWidth, ContentHeight, intHorizontalPageIndex)
            PrintRectangles.Add(UnitCostColumnRectangle)

            If UnitCostColumnRectangle.Left <= SectionBorder And UnitCostColumnRectangle.Right > SectionBorder Then
                StretchRectanglesOfPreviousPage(UnitCostColumnRectangle)
            End If
            intX = UnitCostColumnRectangle.Right

            If ShowLocalCurrencyColumns = True Then
                TotalLocalCostColumnRectangle = New PrintRectangle(intX, ContentTop, intTotalCostColumnWidth, ContentHeight, intHorizontalPageIndex)
                PrintRectangles.Add(TotalLocalCostColumnRectangle)

                If TotalLocalCostColumnRectangle.Left <= SectionBorder And TotalLocalCostColumnRectangle.Right > SectionBorder Then
                    StretchRectanglesOfPreviousPage(TotalLocalCostColumnRectangle)
                End If
                intX = TotalLocalCostColumnRectangle.Right
            End If
        End If

        TotalCostColumnRectangle = New PrintRectangle(intX, ContentTop, intTotalCostColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(TotalCostColumnRectangle)

        If TotalCostColumnRectangle.Left <= SectionBorder And TotalCostColumnRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(TotalCostColumnRectangle)
        End If

        HorPages = Math.Ceiling(SectionBorder / ContentRight)
        RaiseEvent MinimumPageWidthChanged(Me, New MinimumPageWidthChangedEventArgs(HorPages))
    End Sub

    Private Sub StretchRectanglesOfPreviousPage(ByVal selPrintRectangle As PrintRectangle)
        'move current print rectangle to next page
        intHorizontalPageIndex += 1
        selPrintRectangle.X = LeftMargin + SectionBorder + 1
        selPrintRectangle.HorizontalPageIndex = intHorizontalPageIndex

        SectionBorder += ContentRight

        'stretch print rectangles on previous page
        Dim PreviousPageRectangles As PrintRectangles = Me.PrintRectangles.GetRectanglesOfPage(intHorizontalPageIndex - 1)

        If PreviousPageRectangles IsNot Nothing Then
            Dim intSpace As Integer = ContentWidth - PreviousPageRectangles.GetTotalWidth - 2
            Dim intTextRectanglesCount = PreviousPageRectangles.CountTextRectangles

            If intTextRectanglesCount > 0 Then
                Dim selRectangle As PrintRectangle
                intSpace /= intTextRectanglesCount

                For Each selRectangle In PreviousPageRectangles
                    If selRectangle.IsTextRectangle = True Then selRectangle.Width += intSpace
                Next
                For i = 1 To PreviousPageRectangles.Count - 1
                    selRectangle = PreviousPageRectangles(i)
                    selRectangle.X = PreviousPageRectangles(i - 1).Right
                    If selRectangle.Right > SectionBorder Then selRectangle.Width = SectionBorder - selRectangle.X - 1
                Next
            End If
        End If
    End Sub

    Private Function GetSortColumnWidth() As Integer
        Dim intWidth, intSortWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintBudgetRow In Me.PrintList
                If selRow.SortNumber Is Nothing Then selRow.SortNumber = String.Empty

                intWidth = PageGraph.MeasureString(selRow.SortNumber, fntTextBold).Width

                If intWidth > intSortWidth Then intSortWidth = intWidth
            Next

            intSortWidth += (CONST_HorizontalPadding * 2)

        End If

        Return intSortWidth
    End Function

    Private Function GetDurationColumnWidth() As Integer
        Dim intWidth, intDurationWidth As Integer

        If ShowDurationColumns = True Then
            If PageGraph IsNot Nothing Then
                For Each selRow As PrintBudgetRow In Me.PrintList
                    intWidth = PageGraph.MeasureString(selRow.DurationText, fntText).Width

                    If intWidth > intDurationWidth Then intDurationWidth = intWidth
                Next
                If intDurationWidth > 0 Then intDurationWidth += (CONST_HorizontalPadding * 2)
            End If
        End If

        Return intDurationWidth
    End Function

    Private Function GetNumberColumnWidth() As Integer
        Dim intWidth, intNumberWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintBudgetRow In Me.PrintList
                intWidth = PageGraph.MeasureString(selRow.NumberText, fntText).Width

                If intWidth > intNumberWidth Then intNumberWidth = intWidth
            Next

            If intNumberWidth > 0 Then intNumberWidth += (CONST_HorizontalPadding * 2)
        End If

        Return intNumberWidth
    End Function

    Private Function GetUnitCostColumnWidth() As Integer
        Dim intWidth, intUnitCostWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintBudgetRow In Me.PrintList
                intWidth = PageGraph.MeasureString(selRow.UnitCostText, fntText).Width

                If intWidth > intUnitCostWidth Then intUnitCostWidth = intWidth
            Next

            If intUnitCostWidth > 0 Then intUnitCostWidth += (CONST_HorizontalPadding * 2)
        End If

        Return intUnitCostWidth
    End Function

    Private Function GetTotalLocalCostColumnWidth() As Integer
        Dim intWidth, intTotalLocalCostWidth As Integer

        If ShowLocalCurrencyColumns = True Then
            If PageGraph IsNot Nothing Then
                For Each selRow As PrintBudgetRow In Me.PrintList
                    intWidth = PageGraph.MeasureString(selRow.TotalLocalCostText, fntText).Width

                    If intWidth > intTotalLocalCostWidth Then intTotalLocalCostWidth = intWidth
                Next

                If intTotalLocalCostWidth > 0 Then intTotalLocalCostWidth += (CONST_HorizontalPadding * 2)
            End If
        End If

        Return intTotalLocalCostWidth
    End Function

    Private Function GetTotalCostColumnWidth() As Integer
        Dim intWidth, intTotalCostWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintBudgetRow In Me.PrintList
                intWidth = PageGraph.MeasureString(selRow.TotalCostText, fntText).Width

                If intWidth > intTotalCostWidth Then intTotalCostWidth = intWidth
            Next

            If intTotalCostWidth > 0 Then intTotalCostWidth += (CONST_HorizontalPadding * 2)
        End If

        Return intTotalCostWidth
    End Function
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintBudgetRow In Me.PrintList
            If selRow.RowType = PrintBudgetRow.RowTypes.Normal Or selRow.RowType = PrintBudgetRow.RowTypes.TotalBudget Or selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then
                ReloadImages_Normal(selRow)
            End If
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintBudgetRow)
        Dim intColumnWidth As Integer = TextColumnRectangle.Width

        SetBackGroundColor(selRow)

        With RichTextManager
            If String.IsNullOrEmpty(selRow.BudgetItem.Text) Then
                selRow.BudgetItemImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, BudgetItem.ItemName, selRow.SortNumber, False)
            Else
                selRow.BudgetItemImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.BudgetItem.RTF, False, selRow.Indent, HorizontalAlignment.Left, colForeGround, colBackGround)
            End If
            selRow.BudgetItemHeight = selRow.BudgetItemImage.Height
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As PrintBudgetRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer

        If selRow IsNot Nothing Then
            Select Case selRow.RowType
                Case PrintBudgetRow.RowTypes.Normal, PrintBudgetRow.RowTypes.TotalBudget
                    intRowHeight = selRow.BudgetItemHeight
                Case PrintBudgetRow.RowTypes.ExchangeRate
                    SetExchangeRateHeight(selRow)
                Case PrintBudgetRow.RowTypes.TotalBudgetRow
                    intRowHeight = selRow.BudgetItemHeight + (CONST_TotalBudgetRowPadding * 2)
                Case Else
                    intRowHeight = CONST_BudgetYearHeaderHeight
            End Select

        End If

        If intRowHeight > 0 Then selRow.RowHeight = intRowHeight Else selRow.RowHeight = NewCellHeight()
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Sub SetColumnHeadersHeight()
        Dim intHeight As Integer

        If PageGraph IsNot Nothing Then
            Dim fntHeader As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
            Dim intSortNumberHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, SortColumnRectangle.Width).Height
            If intSortNumberHeight > intHeight Then intHeight = intSortNumberHeight

            Dim intDescriptionHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, TextColumnRectangle.Width).Height
            If intDescriptionHeight > intHeight Then intHeight = intDescriptionHeight

            If ShowDurationColumns = True Then
                Dim intDurationHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, DurationColumnRectangle.Width).Height
                If intDurationHeight > intHeight Then intHeight = intDurationHeight
            End If

            Dim intNumberHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, NumberColumnRectangle.Width).Height
            If intNumberHeight > intHeight Then intHeight = intNumberHeight

            Dim intUnitCostHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, UnitCostColumnRectangle.Width).Height
            If intUnitCostHeight > intHeight Then intHeight = intUnitCostHeight

            If ShowLocalCurrencyColumns = True Then
                Dim intTotalLocalCostHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, TotalLocalCostColumnRectangle.Width).Height
                If intTotalLocalCostHeight > intHeight Then intHeight = intTotalLocalCostHeight
            End If

            Dim intTotalCostHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, TotalCostColumnRectangle.Width).Height
            If intTotalCostHeight > intHeight Then intHeight = intTotalCostHeight

            ColumnHeadersHeight = intHeight
        End If
    End Sub

    Private Function SetExchangeRateHeight(ByVal selRow As PrintBudgetRow)
        Dim intHeight As Integer
        Dim intWidth As Integer = ContentWidth / 4
        If PageGraph IsNot Nothing Then
            Dim intCurrencyHeight As Integer = PageGraph.MeasureString(selRow.CurrencyText, fntText, intWidth).Height
            If intCurrencyHeight > intHeight Then intHeight = intCurrencyHeight

            Dim intExchangeRateHeight As Integer = PageGraph.MeasureString(selRow.ExchangeRateText, fntText, intWidth).Height
            If intExchangeRateHeight > intHeight Then intHeight = intExchangeRateHeight

            Dim intConversionHeight As Integer = PageGraph.MeasureString(selRow.Conversion1, fntText, intWidth).Height
            If intConversionHeight > intHeight Then intHeight = intConversionHeight

            intConversionHeight = PageGraph.MeasureString(selRow.Conversion2, fntText, intWidth).Height
            If intConversionHeight > intHeight Then intHeight = intConversionHeight
        End If

        Return intHeight
    End Function
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight - ColumnHeadersHeight

        For Each selRow As PrintBudgetRow In PrintList
            intTotalHeight += selRow.RowHeight
        Next

        decPages = intTotalHeight / intAvailableHeight
        decPages = Math.Ceiling(decPages)
        decPages *= HorPages

        Return decPages
    End Function

    Private Sub SetBackGroundColor(ByVal selRow As PrintBudgetRow)
        Select Case selRow.RowType
            Case PrintBudgetRow.RowTypes.TotalBudgetRow
                colForeGround = Color.White
                colBackGround = Color.Black
            Case Else
                If selRow.IsSubTotal Then
                    colBackGround = LighterColor(colBaseGrey, selRow.Indent)
                    If selRow.Indent > 3 Then colForeGround = Color.Black Else colForeGround = Color.White
                Else
                    colForeGround = Color.Black
                    colBackGround = Color.White
                End If
        End Select
    End Sub

    Private Function GetCoordinateX(ByVal intColumnX As Integer) As Integer
        Dim intX As Integer = intColumnX - CurrentSectionRectangle.X

        Return intX
    End Function
#End Region

#Region "Print page"
    Protected Overrides Sub OnBeginPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnBeginPrint(e)

        boolColumnsWidthSet = False
        PrintRectangles.Clear()
        HorPages = 0
        CurrentHorPage = 1
        CurrentSectionRectangle = New Rectangle(0, 0, ContentRight, ContentHeight)

        intStartIndex = 0
        SectionBorder = 0
        LastColumnBorder = 0
        intHorizontalPageIndex = 0

        LastRowY = ContentTop

        CreateList()
    End Sub

    Protected Overrides Sub OnQueryPageSettings(ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs)
        MyBase.OnQueryPageSettings(e)
    End Sub

    Protected Overrides Sub OnEndPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnEndPrint(e)
    End Sub

    Protected Overrides Sub OnPrintPage(ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        MyBase.OnPrintPage(e)

        PageGraph = e.Graphics

        If boolColumnsWidthSet = False Then
            SetColumnsWidth()
            ReloadImages()
            SetColumnHeadersHeight()
            Me.TotalPages = GetTotalPages()

            boolColumnsWidthSet = True
        End If

        Dim intRowCount As Integer = PrintList.Count
        Dim selRow As PrintBudgetRow = PrintList(RowIndex)

        LastRowY = ContentTop

        If CurrentHorPage = 1 Then intStartIndex = RowIndex

        'Print Header
        PrintHeader()
        PrintColumnHeaders()

        If PrintList.Count > 0 Then
            Do While LastRowY + PrintList(RowIndex).RowHeight < Me.ContentBottom
                selRow = PrintList(RowIndex)

                Select Case selRow.RowType
                    Case PrintBudgetRow.RowTypes.Normal, PrintBudgetRow.RowTypes.TotalBudget, PrintBudgetRow.RowTypes.TotalBudgetRow
                        SetBackGroundColor(selRow)

                        PrintPage_PrintSortNumber(selRow)
                        PrintPage_PrintRtf(selRow)

                        PrintPage_PrintDuration(selRow)
                        PrintPage_PrintNumber(selRow)
                        PrintPage_PrintUnitCost(selRow)
                        PrintPage_PrintTotalLocalCost(selRow)

                        PrintPage_PrintTotalCost(selRow)
                    Case PrintBudgetRow.RowTypes.ExchangeRate
                        PrintPage_PrintExchangeRate(selRow)
                    Case PrintBudgetRow.RowTypes.TitleRow
                        PrintPage_PrintBudgetYearHeader(selRow)
                End Select

                LastRowY += selRow.RowHeight
                RowIndex += 1
                If RowIndex > PrintList.Count - 1 Then Exit Do
            Loop
        Else
            RaiseEvent PagePrinted(Me, New LinePrintedEventArgs(0, 0))
        End If

        PrintFooter()
        RaiseEvent PagePrinted(Me, New LinePrintedEventArgs(PageNumber - 1, TotalPages))
        LastRowY = ContentTop
        PageNumber += 1

        If CurrentHorPage < HorPages Then
            RowIndex = intStartIndex
            CurrentSectionRectangle.X += (LeftMargin + ContentWidth)
            CurrentHorPage += 1

            e.HasMorePages = True
        Else
            If RowIndex < intRowCount Then
                CurrentHorPage = 1
                CurrentSectionRectangle.X = 0

                e.HasMorePages = True
            Else
                e.HasMorePages = False
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintBudgetYearHeader(ByVal selRow As PrintBudgetRow)
        If SortColumnRectangle.Right >= CurrentSectionRectangle.Left Then
            Dim rBudgetYear As New Rectangle(SortColumnRectangle.X, LastRowY, ContentWidth, selRow.RowHeight)

            Dim formatCells As New StringFormat()
            Dim brText As SolidBrush = New SolidBrush(Color.DarkGray)
            Dim rText As Rectangle = GetTextRectangle(rBudgetYear)
            Dim fntBudgetYear As New Font(CurrentLogFrame.DetailsFont.FontFamily, CurrentLogFrame.DetailsFont.SizeInPoints + 6, FontStyle.Bold)

            formatCells.Alignment = StringAlignment.Near
            formatCells.LineAlignment = StringAlignment.Center

            PageGraph.DrawString(selRow.BudgetItem.RTF, fntBudgetYear, brText, rText, formatCells)
        End If
    End Sub

    Private Sub PrintPage_PrintSortNumber(ByVal selRow As PrintBudgetRow)
        If SortColumnRectangle.Right >= CurrentSectionRectangle.Left Then
            Dim rSortNumber As New Rectangle(SortColumnRectangle.X, LastRowY, SortColumnRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.SortNumber, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintRtf(ByVal selRow As PrintBudgetRow)
        If TextColumnRectangle.Left >= CurrentSectionRectangle.Left And TextColumnRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rImage As New Rectangle(TextColumnRectangle.X, LastRowY, TextColumnRectangle.Width, selRow.RowHeight)
            Dim boolTotalBudgetRow As Boolean

            If selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then boolTotalBudgetRow = True

            PrintPage_PrintImage(selRow.BudgetItemImage, rImage, boolTotalBudgetRow)
        End If
    End Sub

    Private Sub PrintPage_PrintDuration(ByVal selRow As PrintBudgetRow)
        If ShowDurationColumns = True Then
            If DurationColumnRectangle.Left >= CurrentSectionRectangle.Left And DurationColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim intX As Integer = GetCoordinateX(DurationColumnRectangle.X)
                Dim rDuration As New Rectangle(intX, LastRowY, DurationColumnRectangle.Width, selRow.RowHeight)
                Dim strText = selRow.DurationText

                If selRow.IsSubTotal = True Then strText = String.Empty

                If selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then
                    PrintPage_PrintTotalText(String.Empty, rDuration)
                Else
                    PrintPage_PrintText(strText, rDuration)
                End If

            End If
        End If
    End Sub

    Private Sub PrintPage_PrintNumber(ByVal selRow As PrintBudgetRow)
        If NumberColumnRectangle.Left >= CurrentSectionRectangle.Left And NumberColumnRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(NumberColumnRectangle.X)
            Dim rNumber As New Rectangle(intX, LastRowY, NumberColumnRectangle.Width, selRow.RowHeight)

            If selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then
                PrintPage_PrintTotalText(String.Empty, rNumber)
            Else
                PrintPage_PrintText(selRow.NumberText, rNumber)
            End If

        End If
    End Sub

    Private Sub PrintPage_PrintUnitCost(ByVal selRow As PrintBudgetRow)

        If UnitCostColumnRectangle.Left >= CurrentSectionRectangle.Left And UnitCostColumnRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(UnitCostColumnRectangle.X)
            Dim rUnitCost As New Rectangle(intX, LastRowY, UnitCostColumnRectangle.Width, selRow.RowHeight)
            Dim strText = selRow.UnitCostText

            If selRow.IsSubTotal = True Then strText = String.Empty

            If selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then
                PrintPage_PrintTotalText(String.Empty, rUnitCost)
            Else
                PrintPage_PrintText(strText, rUnitCost, True)
            End If

        End If
    End Sub

    Private Sub PrintPage_PrintTotalLocalCost(ByVal selRow As PrintBudgetRow)
        If ShowLocalCurrencyColumns = True Then
            If TotalLocalCostColumnRectangle.Left >= CurrentSectionRectangle.Left And TotalLocalCostColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim intX As Integer = GetCoordinateX(TotalLocalCostColumnRectangle.X)
                Dim rTotalLocalCost As New Rectangle(intX, LastRowY, TotalLocalCostColumnRectangle.Width, selRow.RowHeight)

                If selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then
                    PrintPage_PrintTotalText(String.Empty, rTotalLocalCost)
                Else
                    PrintPage_PrintText(selRow.TotalLocalCostText, rTotalLocalCost, True)
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintTotalCost(ByVal selRow As PrintBudgetRow)
        If TotalCostColumnRectangle.Left >= CurrentSectionRectangle.Left And TotalCostColumnRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(TotalCostColumnRectangle.X)
            Dim rTotalCost As New Rectangle(intX, LastRowY, TotalCostColumnRectangle.Width, selRow.RowHeight)

            If selRow.RowType = PrintBudgetRow.RowTypes.TotalBudgetRow Then
                PrintPage_PrintTotalText(selRow.TotalCostText, rTotalCost, True)
            Else
                PrintPage_PrintText(selRow.TotalCostText, rTotalCost, True)
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintExchangeRate(ByVal selRow As PrintBudgetRow)
        Dim boolHeader As Boolean

        If selRow.ExchangeRate Is Nothing Then
            boolHeader = True
        Else
            SetBackGroundColor(selRow)
        End If

        If SortColumnRectangle.Right >= CurrentSectionRectangle.Left Then
            Dim rExchangeRate As New Rectangle(SortColumnRectangle.X, LastRowY, ContentWidth / 4, selRow.RowHeight)
            PrintPage_PrintText(selRow.CurrencyText, rExchangeRate, False, boolHeader)

            rExchangeRate = New Rectangle(rExchangeRate.Right, LastRowY, ContentWidth / 4, selRow.RowHeight)
            PrintPage_PrintText(selRow.ExchangeRateText, rExchangeRate, True, boolHeader)

            rExchangeRate = New Rectangle(rExchangeRate.Right, LastRowY, ContentWidth / 4, selRow.RowHeight)
            PrintPage_PrintText(selRow.Conversion1, rExchangeRate, False, boolHeader)

            rExchangeRate = New Rectangle(rExchangeRate.Right, LastRowY, ContentWidth / 4, selRow.RowHeight)
            PrintPage_PrintText(selRow.Conversion2, rExchangeRate, False, boolHeader)
        End If
    End Sub

    Private Sub PrintPage_PrintText(ByVal strValue As String, ByVal rCell As Rectangle, Optional ByVal AlignRight As Boolean = False, Optional ByVal boolHeader As Boolean = False, _
                                    Optional ByVal boolTotalBudgetRow As Boolean = False)
        If PageGraph IsNot Nothing Then
            If boolHeader = True Then
                colForeGround = Color.Black
                colBackGround = Color.LightGray
            End If
            PageGraph.FillRectangle(New SolidBrush(colBackGround), rCell)

            Dim formatCells As New StringFormat()
            Dim brText As SolidBrush = New SolidBrush(colForeGround)
            Dim rText As Rectangle = GetTextRectangle(rCell)

            If boolTotalBudgetRow = True Then
                rText.Y += CONST_TotalBudgetRowPadding
                rText.Height -= (CONST_TotalBudgetRowPadding * 2)
            End If

            If boolHeader = True Then
                formatCells.Alignment = StringAlignment.Center
                formatCells.LineAlignment = StringAlignment.Center
                PageGraph.DrawString(strValue, fntTextBold, brText, rText, formatCells)
            Else
                If AlignRight = True Then
                    formatCells.Alignment = StringAlignment.Far
                Else
                    formatCells.Alignment = StringAlignment.Near
                End If

                formatCells.LineAlignment = StringAlignment.Near

                PageGraph.DrawString(strValue, fntText, brText, rText, formatCells)
            End If

            PageGraph.DrawRectangle(penBlack1, rCell)
        End If
    End Sub

    Private Sub PrintPage_PrintTotalText(ByVal strValue As String, ByVal rCell As Rectangle, Optional ByVal AlignRight As Boolean = False)
        If PageGraph IsNot Nothing Then
            colForeGround = Color.White
            colBackGround = Color.Black
            PageGraph.FillRectangle(New SolidBrush(colBackGround), rCell)

            Dim formatCells As New StringFormat()
            Dim brText As SolidBrush = New SolidBrush(colForeGround)
            Dim rText As Rectangle = GetTextRectangle(rCell)

            rText.Y += CONST_TotalBudgetRowPadding
            rText.Height -= (CONST_TotalBudgetRowPadding * 2)

            If AlignRight = True Then
                formatCells.Alignment = StringAlignment.Far
            Else
                formatCells.Alignment = StringAlignment.Near
            End If

            formatCells.LineAlignment = StringAlignment.Near

            PageGraph.DrawString(strValue, fntTextBold, brText, rText, formatCells)

            PageGraph.DrawRectangle(penBlack1, rCell)
        End If
    End Sub

    Private Sub PrintPage_PrintImage(ByVal bmCellImage As Bitmap, ByVal rImage As Rectangle, Optional ByVal boolTotalBudgetRow As Boolean = False)
        If PageGraph IsNot Nothing Then
            Dim rCell As Rectangle = rImage
            PageGraph.FillRectangle(New SolidBrush(colBackGround), rCell)

            If bmCellImage IsNot Nothing Then
                If boolTotalBudgetRow = True Then
                    rImage.Y += CONST_TotalBudgetRowPadding
                    rImage.Height -= (CONST_TotalBudgetRowPadding * 2)
                End If

                PageGraph.DrawImage(bmCellImage, rImage)
            End If
            PageGraph.DrawRectangle(penBlack1, rCell)
        End If
    End Sub

    Private Sub PrintPage_Background(ByVal intRowHeight As Integer)
        Dim rBackGround As New Rectangle
        'Dim rAvailableArea As Rectangle = GetAvailableArea()
        'Dim brDuration As Brush = Brushes.Gainsboro
        'Dim brWeekEnd As New SolidBrush(Me.DarkerColor(Color.Gainsboro))

        'Dim intX, intY As Integer
        'Dim penGrid As New Pen(SystemColors.ControlDark, 1)
        'Dim intRightBorder As Integer = GetCoordinateX(PeriodUntil.AddDays(1)) - rCurrentSection.X
        'Dim intBudgetColumnWidth As Integer
        'Dim intCurrentBudgetColumnX As Integer = GetCurrentBudgetColumnX()

        'If PageGraph IsNot Nothing Then
        '    If intRightBorder < GetCurrentBudgetColumnWidth() Then
        '        intBudgetColumnWidth = intRightBorder
        '    Else
        '        intBudgetColumnWidth = GetCurrentBudgetColumnWidth()
        '    End If

        '    'Project duration
        '    rBackGround = rProjectDuration
        '    rBackGround.Y = LastRowY
        '    rBackGround.Height = intRowHeight

        '    If rBackGround.Left <= CurrentSectionRectangle.Right And rBackGround.Right >= Me.CurrentSectionRectangle.Left Then
        '        rBackGround.X -= CurrentSectionRectangle.X
        '        If rBackGround.X < 0 Then
        '            rBackGround.Width += rBackGround.X
        '            rBackGround.X = 0
        '        End If
        '        rBackGround.X += GetCurrentBudgetColumnX()

        '        If rBackGround.Right > rAvailableArea.Right Then
        '            rBackGround.Width = rAvailableArea.Right - rBackGround.X
        '        End If
        '        PageGraph.FillRectangle(brDuration, rBackGround)
        '    End If

        '    rBudgetCell = New Rectangle(GetCurrentBudgetColumnX, LastRowY, intBudgetColumnWidth, intRowHeight)
        '    PageGraph.DrawRectangle(penBlack1, rBudgetCell)
        'End If
    End Sub
#End Region

#Region "Column headers"
    Private Sub PrintColumnHeaders()
        Dim intX As Integer

        If PageGraph IsNot Nothing Then
            If SortColumnRectangle.Right >= CurrentSectionRectangle.Left Then
                Dim rSortNumber As New Rectangle(SortColumnRectangle.X, LastRowY, SortColumnRectangle.Width, ColumnHeadersHeight)
                PrintPage_PrintText(LANG_Number, rSortNumber, False, True)
            End If
            If TextColumnRectangle.Left >= CurrentSectionRectangle.Left And TextColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim rDescription As New Rectangle(TextColumnRectangle.X, LastRowY, TextColumnRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Description, rDescription, False, True)
            End If
            If ShowDurationColumns = True Then
                If DurationColumnRectangle.Left >= CurrentSectionRectangle.Left And DurationColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(DurationColumnRectangle.X)
                    Dim rDuration As New Rectangle(intX, LastRowY, DurationColumnRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_Duration, rDuration, False, True)
                End If
            End If
            If NumberColumnRectangle.Left >= CurrentSectionRectangle.Left And NumberColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(NumberColumnRectangle.X)
                Dim rNumber As New Rectangle(intX, LastRowY, NumberColumnRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Quantity, rNumber, False, True)
            End If
            If UnitCostColumnRectangle.Left >= CurrentSectionRectangle.Left And UnitCostColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(UnitCostColumnRectangle.X)
                Dim rUnitCost As New Rectangle(intX, LastRowY, UnitCostColumnRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_UnitCost, rUnitCost, False, True)
            End If
            If ShowLocalCurrencyColumns = True Then
                If TotalLocalCostColumnRectangle.Left >= CurrentSectionRectangle.Left And TotalLocalCostColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(TotalLocalCostColumnRectangle.X)
                    Dim rTotalLocalCost As New Rectangle(intX, LastRowY, TotalLocalCostColumnRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_TotalLocalCost, rTotalLocalCost, False, True)
                End If
            End If
            If TotalCostColumnRectangle.Left >= CurrentSectionRectangle.Left And TotalCostColumnRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(TotalCostColumnRectangle.X)
                Dim rTotalCost As New Rectangle(intX, LastRowY, TotalCostColumnRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_TotalCost, rTotalCost, False, True)
            End If
        End If

        LastRowY += ColumnHeadersHeight
    End Sub
#End Region
End Class


