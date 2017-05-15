Public Class PrintResources
    Inherits ReportBaseClass

    Private Const CONST_MinTextColumnWidth As Integer = 100
    Private Const CONST_TabSpace As Integer = 15

    Private objLogFrame As New LogFrame
    Private objPrintList As New PrintResourceRows
    Private objClippedRow As PrintResourceRow = Nothing
    Private intPagesWide As Integer

    Private boolColumnsWidthSet As Boolean
    Private CurrentSectionRectangle As New Rectangle
    Private intColumnHeadersHeight As Integer
    Private rStructSortRectangle, rStructRectangle As New PrintRectangle
    Private rResourceSortRectangle, rResourceRectangle As New PrintRectangle
    Private rBudgetItemRectangle, rBudgetRectangle, rPercentageRectangle, rTotalCostRectangle As New PrintRectangle

    Private intHorPages As Integer, intCurrentHorPage As Integer = 1
    Private intStartIndex As Integer
    Private ColumnHeadersHeight, TargetRowTitleHeight As Integer
    Private intSectionBorder As Integer
    Private intLastColumnBorder As Integer
    Private intHorizontalPageIndex As Integer

    Private colLightGrey As Color = Color.FromArgb(212, 212, 212)
    Private colMiddleGrey As Color = Color.FromArgb(150, 150, 150)
    Private colDarkGrey As Color = Color.FromArgb(128, 128, 128)
    Private brLightGrey As New SolidBrush(colLightGrey)
    Private brMiddleGrey As New SolidBrush(colMiddleGrey)
    Private brDarkGrey As New SolidBrush(colDarkGrey)

    Public Event PagePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
    Public Event MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)

    Public Sub New(ByVal logframe As LogFrame, ByVal pageswide As Integer)
        Me.LogFrame = logframe
        Me.ReportSetup = logframe.ReportSetupResources

        Me.PagesWide = pageswide
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

    Public Property PrintList() As PrintResourceRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintResourceRows)
            objPrintList = value
        End Set
    End Property

    Public Property PagesWide As Integer
        Get
            Return intPagesWide
        End Get
        Set(ByVal value As Integer)
            intPagesWide = value
        End Set
    End Property

    Public Property ClippedRow As PrintResourceRow
        Get
            Return objClippedRow
        End Get
        Set(ByVal value As PrintResourceRow)
            objClippedRow = value
        End Set
    End Property

    Private Property StructSortRectangle As PrintRectangle
        Get
            Return rStructSortRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rStructSortRectangle = value
        End Set
    End Property

    Private Property StructRectangle As PrintRectangle
        Get
            Return rStructRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rStructRectangle = value
        End Set
    End Property

    Private Property ResourceSortRectangle As PrintRectangle
        Get
            Return rResourceSortRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rResourceSortRectangle = value
        End Set
    End Property

    Private Property ResourceRectangle As PrintRectangle
        Get
            Return rResourceRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rResourceRectangle = value
        End Set
    End Property

    Private Property BudgetItemRectangle As PrintRectangle
        Get
            Return rBudgetItemRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rBudgetItemRectangle = value
        End Set
    End Property

    Private Property BudgetAmountRectangle As PrintRectangle
        Get
            Return rBudgetRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rBudgetRectangle = value
        End Set
    End Property

    Private Property PercentageRectangle As PrintRectangle
        Get
            Return rPercentageRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rPercentageRectangle = value
        End Set
    End Property

    Private Property TotalCostRectangle As PrintRectangle
        Get
            Return rTotalCostRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rTotalCostRectangle = value
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

#Region "Create resource table"
    Public Sub CreateList()
        Dim strPurposeSort As String = String.Empty, strOutputSort As String = String.Empty

        PrintList.Clear()

        For Each selPurpose As Purpose In Me.LogFrame.Purposes
            strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))
            PrintList.Add(New PrintResourceRow(PrintResourceRow.Types.Purpose, strPurposeSort, New Purpose(selPurpose.RTF)))

            For Each selOutput As Output In selPurpose.Outputs
                strOutputSort = LogFrame.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)
                PrintList.Add(New PrintResourceRow(PrintResourceRow.Types.Output, strOutputSort, New Output(selOutput.RTF)))

                CreateList_Activities(selOutput.Activities, strOutputSort)
            Next
        Next
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String)
        Dim strActivitySort As String = String.Empty

        For Each selActivity As Activity In selActivities
            strActivitySort = LogFrame.CreateSortNumber(selActivities.IndexOf(selActivity), strParentSort)

            Dim selRow As New PrintResourceRow(PrintResourceRow.Types.Activity, strActivitySort, New Activity(selActivity.RTF))
            PrintList.Add(selRow)
            CreateList_Resources(selActivity.Resources, strActivitySort)

            If selActivity.Activities.Count > 0 Then
                CreateList_Activities(selActivity.Activities, strActivitySort)
            End If
        Next
    End Sub

    Private Sub CreateList_Resources(ByVal selResources As Resources, ByVal strActivitySort As String)
        Dim strResourceSort As String = String.Empty

        For Each selResource As Resource In selResources
            Dim objResourceCopy As New Resource(selResource.RTF)
            objResourceCopy.TotalCostAmount = selResource.TotalCostAmount

            strResourceSort = LogFrame.CreateSortNumber(selResources.IndexOf(selResource), strActivitySort)

            Dim selRow As New PrintResourceRow(strResourceSort, objResourceCopy)
            PrintList.Add(selRow)

            CreateList_BudgetItems(selResource.BudgetItemReferences)
        Next
    End Sub

    Private Sub CreateList_BudgetItems(ByVal selBudgetItemReferences As BudgetItemReferences)
        Dim intBudgetNr As Integer

        For Each selBudgetItemReference As BudgetItemReference In selBudgetItemReferences
            intBudgetNr += 1

            Dim selBudgetItem As BudgetItem = LogFrame.GetBudgetItemByGuid(selBudgetItemReference.ReferenceBudgetItemGuid)
            Dim objBudgetItemCopy As New BudgetItem(selBudgetItem.RTF)
            objBudgetItemCopy.TotalCost = selBudgetItem.TotalCost

            Dim NewRow As New PrintResourceRow(intBudgetNr.ToString, objBudgetItemCopy, selBudgetItemReference.Percentage, selBudgetItemReference.TotalCost)
            PrintList.Add(NewRow)
        Next
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth * PagesWide
        Dim intStructSortWidth, intResourceSortWidth As Integer
        Dim intSortWidth As Integer
        Dim intStructColumnWidth, intResourceColumnWidth, intBudgetItemColumnWidth As Integer
        Dim intBudgetAmountWidth, intPercentageWidth, intTotalCostColumnWidth As Integer

        SectionBorder = ContentRight

        'calculate column widths
        intStructSortWidth = GetStructSortColumnWidth()
        intResourceSortWidth = GetResourceSortColumnWidth()
        intBudgetAmountWidth = GetBudgetAmountColumnWidth()
        intPercentageWidth = intBudgetAmountWidth
        intTotalCostColumnWidth = GetTotalCostColumnWidth()
        intStructColumnWidth = ContentWidth - intStructSortWidth
        intResourceColumnWidth = intStructColumnWidth - intResourceSortWidth - intTotalCostColumnWidth

        intSortWidth = intStructSortWidth + intResourceSortWidth
        intBudgetItemColumnWidth = intAvailableWidth - intSortWidth - intBudgetAmountWidth - intPercentageWidth - intTotalCostColumnWidth

        If intBudgetItemColumnWidth < CONST_MinTextColumnWidth Then intBudgetItemColumnWidth = CONST_MinTextColumnWidth
        If intBudgetItemColumnWidth + intStructSortWidth + intResourceSortWidth > ContentWidth Then intBudgetItemColumnWidth = ContentWidth - intStructSortWidth - intResourceSortWidth

        'set struct column rectangles
        StructSortRectangle = New PrintRectangle(LeftMargin, ContentTop, intStructSortWidth, ContentHeight, intHorizontalPageIndex)
        StructRectangle = New PrintRectangle(StructSortRectangle.Right, ContentTop, intStructColumnWidth, ContentHeight, intHorizontalPageIndex)


        'set resource column rectangles
        ResourceSortRectangle = New PrintRectangle(StructSortRectangle.Right, ContentTop, intResourceSortWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(ResourceSortRectangle)

        ResourceRectangle = New PrintRectangle(ResourceSortRectangle.Right, ContentTop, intResourceColumnWidth, ContentHeight, intHorizontalPageIndex)

        'referenced budget item column
        BudgetItemRectangle = New PrintRectangle(ResourceSortRectangle.Right, ContentTop, intBudgetItemColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(BudgetItemRectangle)

        If BudgetItemRectangle.Right > intSectionBorder Then BudgetItemRectangle.Width = ContentRight - BudgetItemRectangle.Left

        'referenced budget amount column
        BudgetAmountRectangle = New PrintRectangle(BudgetItemRectangle.Right, ContentTop, intBudgetAmountWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(BudgetAmountRectangle)

        If BudgetAmountRectangle.Left <= intSectionBorder And BudgetAmountRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(BudgetAmountRectangle)
        End If

        'percentage column
        PercentageRectangle = New PrintRectangle(BudgetAmountRectangle.Right, ContentTop, intBudgetAmountWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(PercentageRectangle)

        If PercentageRectangle.Left <= intSectionBorder And PercentageRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(PercentageRectangle)
        End If

        TotalCostRectangle = New PrintRectangle(PercentageRectangle.Right, ContentTop, intTotalCostColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(TotalCostRectangle)

        If TotalCostRectangle.Left <= intSectionBorder And TotalCostRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(TotalCostRectangle)
        End If

        HorPages = Math.Ceiling(intSectionBorder / ContentRight)
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
            Dim intSpace As Integer = ContentWidth - StructSortRectangle.Width - PreviousPageRectangles.GetTotalWidth - 2
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

    Private Function GetStructSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintResourceRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.StructSort) = False AndAlso selRow.StructSort.Length > strSort.Length Then strSort = selRow.StructSort
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + (CONST_HorizontalPadding * 2)
        End If
        Return intWidth
    End Function

    Private Function GetResourceSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintResourceRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.ResourceSort) = False AndAlso selRow.ResourceSort.Length > strSort.Length Then strSort = selRow.ResourceSort
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + (CONST_HorizontalPadding * 2)
        End If
        Return intWidth
    End Function

    Private Function GetBudgetAmountColumnWidth() As Integer
        Dim intWidth, intBudgetWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintResourceRow In Me.PrintList
                If selRow.BudgetItem IsNot Nothing Then
                    intWidth = PageGraph.MeasureString(selRow.BudgetItem.TotalCost.ToString, fntText).Width

                    If intWidth > intBudgetWidth Then intBudgetWidth = intWidth
                End If
            Next

            If intBudgetWidth > 0 Then intBudgetWidth += (CONST_HorizontalPadding * 2)
        End If

        Return intBudgetWidth
    End Function

    Private Function GetTotalCostColumnWidth() As Integer
        Dim intWidth, intTotalCostWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintResourceRow In Me.PrintList
                Select Case selRow.Type
                    Case PrintResourceRow.Types.Resource
                        Dim strTotalCost As String = String.Format("{0} {1}", selRow.Resource.TotalCostAmount.ToString("N2"), CurrentLogFrame.CurrencyCode)
                        intWidth = PageGraph.MeasureString(strTotalCost.ToString, fntText).Width

                        If intWidth > intTotalCostWidth Then intTotalCostWidth = intWidth
                    Case PrintResourceRow.Types.BudgetItem
                        If selRow.TotalCost IsNot Nothing Then
                            intWidth = PageGraph.MeasureString(selRow.TotalCost.ToString, fntText).Width

                            If intWidth > intTotalCostWidth Then intTotalCostWidth = intWidth
                        End If
                End Select

            Next

            If intTotalCostWidth > 0 Then intTotalCostWidth += (CONST_HorizontalPadding * 2)
        End If

        Return intTotalCostWidth
    End Function
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintResourceRow In Me.PrintList
            ReloadImages_Normal(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintResourceRow)
        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Struct.Text) Then
                    selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.GetItemName(), selRow.StructSort, False)
                Else
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.RTF, False)
                End If
            End If
            If selRow.Resource IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Resource.Text) Then
                    selRow.Resource.CellImage = .EmptyTextWithPaddingToBitmap(ResourceRectangle.Width, Resource.ItemName, selRow.ResourceSort, False)
                Else
                    selRow.Resource.CellImage = .RichTextWithPaddingToBitmap(ResourceRectangle.Width, selRow.Resource.RTF, False)
                End If
            End If
            If selRow.BudgetItem IsNot Nothing Then
                selRow.BudgetItem.CellImage = .RichTextWithPaddingToBitmap(BudgetItemRectangle.Width, selRow.BudgetItem.RTF, False)
                selRow.BudgetItemHeight = selRow.BudgetItem.CellImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintResourceRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > 0 Then selPrintListRow.RowHeight = intRowHeight Else selPrintListRow.RowHeight = NewCellHeight()
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selRow As PrintResourceRow = Me.PrintList(RowIndex)

        If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
            intRowHeight = selRow.Struct.CellImage.Height
        ElseIf selRow.Resource IsNot Nothing AndAlso selRow.Resource.CellImage IsNot Nothing Then
            intRowHeight = selRow.Resource.CellImage.Height
        ElseIf selRow.BudgetItem IsNot Nothing AndAlso selRow.BudgetItem.CellImage IsNot Nothing Then
            intRowHeight = selRow.BudgetItem.CellImage.Height
        End If

        Return intRowHeight
    End Function

    Private Sub SetColumnHeadersHeight()
        Dim intHeight As Integer

        If PageGraph IsNot Nothing Then
            Dim fntHeader As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)

            Dim intDescriptionWidth As Integer = StructSortRectangle.Width + ResourceSortRectangle.Width + BudgetItemRectangle.Width
            Dim intDescriptionHeight As Integer = PageGraph.MeasureString(LANG_Description, fntHeader, intDescriptionWidth).Height
            If intDescriptionHeight > intHeight Then intHeight = intDescriptionHeight

            Dim intBudgetHeight As Integer = PageGraph.MeasureString(LANG_Budget, fntHeader, BudgetAmountRectangle.Width).Height
            If intBudgetHeight > intHeight Then intHeight = intBudgetHeight

            Dim intPercentageHeight As Integer = PageGraph.MeasureString(LANG_Percentage, fntHeader, PercentageRectangle.Width).Height
            If intPercentageHeight > intHeight Then intHeight = intPercentageHeight

            Dim intTotalCostHeight As Integer = PageGraph.MeasureString(LANG_TotalCost, fntHeader, TotalCostRectangle.Width).Height
            If intTotalCostHeight > intHeight Then intHeight = intTotalCostHeight

            ColumnHeadersHeight = intHeight
        End If
    End Sub
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight - ColumnHeadersHeight

        For Each selRow As PrintResourceRow In PrintList
            intTotalHeight += selRow.RowHeight
        Next

        decPages = intTotalHeight / intAvailableHeight
        decPages = Math.Ceiling(decPages)
        decPages *= HorPages

        Return decPages
    End Function

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
        ClippedRow = Nothing
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
        Dim selRow As PrintResourceRow
        Dim intMinHeight As Integer

        LastRowY = ContentTop
        If CurrentHorPage = 1 Then intStartIndex = RowIndex

        'Print Header
        PrintHeader()
        PrintColumnHeaders()

        If PrintList.Count > 0 Then
            Do
                selRow = PrintList(RowIndex)
                If selRow Is Nothing Then Exit Do

                Select Case selRow.Type
                    Case PrintResourceRow.Types.BudgetItem
                        PrintPage_PrintBudgetItemSortNumber(selRow)
                        PrintPage_PrintBudgetItem(selRow)
                        PrintPage_PrintBudget(selRow)
                        PrintPage_PrintPercentage(selRow)
                        PrintPage_PrintTotalCost(selRow)
                    Case PrintResourceRow.Types.Resource
                        PrintPage_PrintResourceSortNumber(selRow)
                        PrintPage_PrintResource(selRow)
                        PrintPage_PrintTotalCostResource(selRow)
                    Case Else
                        PrintPage_PrintStructSortNumber(selRow)
                        PrintPage_PrintStruct(selRow)
                End Select

                LastRowY += selRow.RowHeight
                RowIndex += 1

                If RowIndex > PrintList.Count - 1 Then Exit Do

                intMinHeight = LastRowY + PrintList(RowIndex).RowHeight
            Loop While intMinHeight < Me.ContentBottom
        Else
            RaiseEvent PagePrinted(Me, New LinePrintedEventArgs(0, 0))
        End If

        If RowIndex >= intRowCount Then
            PageGraph.DrawLine(penBlack1, LeftMargin, LastRowY, LastColumnBorder, LastRowY)
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

    Private Sub PrintPage_PrintText(ByVal strValue As String, ByVal rCell As Rectangle, Optional ByVal AlignRight As Boolean = False, Optional ByVal boolHeader As Boolean = False)
        If PageGraph IsNot Nothing Then
            Dim formatCells As New StringFormat()
            Dim brText As SolidBrush = New SolidBrush(Color.Black)
            Dim rText As Rectangle = GetTextRectangle(rCell)

            If boolHeader = True Then
                PageGraph.FillRectangle(New SolidBrush(Color.LightGray), rCell)

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

            If boolHeader = True Then
                PageGraph.DrawRectangle(penBlack1, rCell)
            End If

        End If
    End Sub
#End Region

#Region "Print struct"
    Private Sub PrintPage_PrintStructSortNumber(ByVal selRow As PrintResourceRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.StructSort, rSortNumber)

            PageGraph.DrawLine(penBlack1, rSortNumber.Left, rSortNumber.Top, rSortNumber.Left, rSortNumber.Bottom)
        End If
    End Sub

    Private Sub PrintPage_PrintStruct(ByVal selRow As PrintResourceRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(StructRectangle.X)

            If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Struct.CellImage.Width, selRow.Struct.CellImage.Height)

                PageGraph.DrawImage(selRow.Struct.CellImage, rImage)
            End If
        End If
        If selRow.Struct IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Struct.RTF) = False Then
            PageGraph.DrawLine(penBlack1, LeftMargin, LastRowY, LastColumnBorder, LastRowY)
            PageGraph.DrawLine(penBlack1, LastColumnBorder, LastRowY, LastColumnBorder, LastRowY + selRow.RowHeight)
        End If
    End Sub
#End Region

#Region "Print resource"
    Private Sub PrintPage_PrintResourceSortNumber(ByVal selRow As PrintResourceRow)
        If ResourceSortRectangle.Left >= CurrentSectionRectangle.Left And ResourceRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(ResourceSortRectangle.X, LastRowY, ResourceSortRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.ResourceSort, rSortNumber)

            PageGraph.DrawLine(penBlack1, rSortNumber.Left, rSortNumber.Top, rSortNumber.Left, rSortNumber.Bottom)
            PageGraph.DrawLine(penBlack1, StructSortRectangle.Left, rSortNumber.Top, StructSortRectangle.Left, rSortNumber.Bottom)
        End If
    End Sub

    Private Sub PrintPage_PrintResource(ByVal selRow As PrintResourceRow)
        Dim intLineX As Integer = ContentLeft

        If ResourceSortRectangle.Left >= CurrentSectionRectangle.Left And ResourceRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ResourceRectangle.X)

            If selRow.Resource IsNot Nothing AndAlso selRow.Resource.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Resource.CellImage.Width, selRow.Resource.CellImage.Height)

                PageGraph.DrawImage(selRow.Resource.CellImage, rImage)
            End If
            intLineX = ResourceSortRectangle.Left
        End If

        If selRow.Resource IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Resource.RTF) = False Then _
            PageGraph.DrawLine(penBlack1, intLineX, LastRowY, LastColumnBorder, LastRowY)
        PageGraph.DrawLine(penBlack1, LastColumnBorder, LastRowY, LastColumnBorder, LastRowY + selRow.RowHeight)
    End Sub

    Private Sub PrintPage_PrintTotalCostResource(ByVal selRow As PrintResourceRow)
        If TotalCostRectangle.Left >= CurrentSectionRectangle.Left And TotalCostRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(TotalCostRectangle.X)
            Dim rTotalCost As New Rectangle(intX, LastRowY, TotalCostRectangle.Width, selRow.RowHeight)
            Dim strTotalCost As String = String.Format("{0} {1}", selRow.Resource.TotalCostAmount.ToString("N2"), CurrentLogFrame.CurrencyCode)

            PrintPage_PrintText(strTotalCost, rTotalCost, True)

            PageGraph.DrawLine(penBlack1, rTotalCost.Left, rTotalCost.Top, rTotalCost.Left, rTotalCost.Bottom)
            PageGraph.DrawLine(penBlack1, rTotalCost.Right, rTotalCost.Top, rTotalCost.Right, rTotalCost.Bottom)
        End If
    End Sub
#End Region

#Region "Print budget item"
    Private Sub PrintPage_PrintBudgetItemSortNumber(ByVal selRow As PrintResourceRow)
        If ResourceSortRectangle.Left >= CurrentSectionRectangle.Left And BudgetItemRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(ResourceSortRectangle.X, LastRowY, ResourceSortRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.BudgetItemSortNumber, rSortNumber, True)

            PageGraph.DrawLine(penBlack1, rSortNumber.Left, rSortNumber.Top, rSortNumber.Left, rSortNumber.Bottom)
            PageGraph.DrawLine(penBlack1, StructSortRectangle.Left, rSortNumber.Top, StructSortRectangle.Left, rSortNumber.Bottom)
        End If
    End Sub

    Private Sub PrintPage_PrintBudgetItem(ByVal selRow As PrintResourceRow)
        Dim intLineX As Integer = ContentLeft

        If ResourceSortRectangle.Left >= CurrentSectionRectangle.Left And BudgetItemRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(BudgetItemRectangle.X)

            If selRow.BudgetItem IsNot Nothing AndAlso selRow.BudgetItem.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.BudgetItem.CellImage.Width, selRow.BudgetItem.CellImage.Height)

                PageGraph.DrawImage(selRow.BudgetItem.CellImage, rImage)
            End If
            intLineX = ResourceSortRectangle.Left
        End If

        If selRow.BudgetItem IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.BudgetItem.RTF) = False Then _
            PageGraph.DrawLine(penBlack1, intLineX, LastRowY, LastColumnBorder, LastRowY)
        PageGraph.DrawLine(penBlack1, LastColumnBorder, LastRowY, LastColumnBorder, LastRowY + selRow.RowHeight)
    End Sub

    Private Sub PrintPage_PrintBudget(ByVal selRow As PrintResourceRow)
        If BudgetAmountRectangle.Left >= CurrentSectionRectangle.Left And BudgetAmountRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(BudgetAmountRectangle.X)
            Dim rBudget As New Rectangle(intX, LastRowY, BudgetAmountRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.BudgetItem.TotalCost.ToString, rBudget, True)

            PageGraph.DrawLine(penBlack1, rBudget.Left, rBudget.Top, rBudget.Left, rBudget.Bottom)
        End If
    End Sub

    Private Sub PrintPage_PrintPercentage(ByVal selRow As PrintResourceRow)
        If PercentageRectangle.Left >= CurrentSectionRectangle.Left And PercentageRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(PercentageRectangle.X)
            Dim rPercentage As New Rectangle(intX, LastRowY, PercentageRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.Percentage.ToString("P2"), rPercentage, True)

            PageGraph.DrawLine(penBlack1, rPercentage.Left, rPercentage.Top, rPercentage.Left, rPercentage.Bottom)
        End If
    End Sub

    Private Sub PrintPage_PrintTotalCost(ByVal selRow As PrintResourceRow)
        If TotalCostRectangle.Left >= CurrentSectionRectangle.Left And TotalCostRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(TotalCostRectangle.X)
            Dim rTotalCost As New Rectangle(intX, LastRowY, TotalCostRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.TotalCost.ToString, rTotalCost, True)

            PageGraph.DrawLine(penBlack1, rTotalCost.Left, rTotalCost.Top, rTotalCost.Left, rTotalCost.Bottom)
            PageGraph.DrawLine(penBlack1, rTotalCost.Right, rTotalCost.Top, rTotalCost.Right, rTotalCost.Bottom)
        End If
    End Sub
#End Region

#Region "Column headers"
    Private Sub PrintColumnHeaders()
        Dim intX As Integer

        If PageGraph IsNot Nothing Then
            If StructSortRectangle.Left >= CurrentSectionRectangle.Left And BudgetItemRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim rDescription As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width + ResourceSortRectangle.Width + BudgetItemRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Description, rDescription, False, True)

                LastColumnBorder = rDescription.Right
            End If

            If BudgetAmountRectangle.Left >= CurrentSectionRectangle.Left And BudgetAmountRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(BudgetAmountRectangle.X)
                Dim rBudget As New Rectangle(intX, LastRowY, BudgetAmountRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Budget, rBudget, False, True)

                LastColumnBorder = rBudget.Right
            End If

            If PercentageRectangle.Left >= CurrentSectionRectangle.Left And PercentageRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(PercentageRectangle.X)
                Dim rPercentage As New Rectangle(intX, LastRowY, PercentageRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Percentage, rPercentage, False, True)

                LastColumnBorder = rPercentage.Right
            End If

            If TotalCostRectangle.Left >= CurrentSectionRectangle.Left And TotalCostRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(TotalCostRectangle.X)
                Dim rTotalCost As New Rectangle(intX, LastRowY, TotalCostRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_TotalCost, rTotalCost, False, True)

                LastColumnBorder = rTotalCost.Right
            End If
        End If

        LastRowY += ColumnHeadersHeight
    End Sub
#End Region

End Class


