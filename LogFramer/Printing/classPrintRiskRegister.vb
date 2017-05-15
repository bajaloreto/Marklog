Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintRiskRegister
    Inherits ReportBaseClass

    Private Const CONST_MinTextColumnWidth As Integer = 100

    Private objLogFrame As New LogFrame
    Private objPrintList As New PrintRiskRegisterRows
    Private objPrintListOperational As New PrintRiskRegisterRows
    Private objPrintListFinancial As New PrintRiskRegisterRows
    Private objPrintListObjectives As New PrintRiskRegisterRows
    Private objPrintListReputation As New PrintRiskRegisterRows
    Private objPrintListOther As New PrintRiskRegisterRows
    Private objPrintListNotDefined As New PrintRiskRegisterRows
    Private objClippedRow As PrintRiskRegisterRow = Nothing
    Private intRiskCategory, CurrentRiskCategory As Integer
    Private intPagesWide As Integer

    Private strClippedTextTop, strClippedTextBottom As String

    Private boolColumnsWidthSet As Boolean
    Private CurrentSectionRectangle As New Rectangle
    Private rStructSortRectangle, rAssumptionSortRectangle As New PrintRectangle
    Private rStructRectangle, rAssumptionRectangle As New PrintRectangle
    Private rRiskResponseRectangle, rRiskLevelRectangle As New PrintRectangle

    'Private intMinHeight As Integer

    Private intHorPages As Integer, intCurrentHorPage As Integer = 1
    Private intStartIndex As Integer
    Private intColumnHeadersHeight As Integer
    Private intSectionBorder As Integer
    Private intLastColumnBorder As Integer
    Private intHorizontalPageIndex As Integer

    Private colForeGround As Color = Color.Black, colBackGround As Color = Color.White
    Private colBaseGrey As Color = Color.FromArgb(255, 50, 50, 50)

    Public Event PagePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)
    Public Event MinimumPageWidthChanged(ByVal sender As Object, ByVal e As MinimumPageWidthChangedEventArgs)

    Public Enum RiskCategories As Integer
        NotDefined = 0
        Operational = 1
        Financial = 2
        Objectives = 3
        Reputation = 4
        Other = 5
        All = 6
    End Enum

    Public Sub New(ByVal logframe As LogFrame, ByVal riskcategory As Integer, ByVal pageswide As Integer)
        Me.LogFrame = logframe

        Me.ReportSetup = logframe.ReportSetupRiskRegister
        Me.RiskCategory = riskcategory
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

    Public Property PrintList() As PrintRiskRegisterRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintRiskRegisterRows)
            objPrintList = value
        End Set
    End Property

    Public Property ClippedRow As PrintRiskRegisterRow
        Get
            Return objClippedRow
        End Get
        Set(ByVal value As PrintRiskRegisterRow)
            objClippedRow = value
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

    Public Property RiskCategory() As Integer
        Get
            Return intRiskCategory
        End Get
        Set(ByVal value As Integer)
            intRiskCategory = value
        End Set
    End Property

    Private Property AssumptionSortRectangle As PrintRectangle
        Get
            Return rAssumptionSortRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rAssumptionSortRectangle = value
        End Set
    End Property

    Private Property AssumptionRectangle As PrintRectangle
        Get
            Return rAssumptionRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rAssumptionRectangle = value
        End Set
    End Property

    Private Property RiskResponseRectangle As PrintRectangle
        Get
            Return rRiskResponseRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rRiskResponseRectangle = value
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

    Private Property RiskLevelRectangle As PrintRectangle
        Get
            Return rRiskLevelRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rRiskLevelRectangle = value
        End Set
    End Property

    Private Property ColumnHeadersHeight() As Integer
        Get
            Return intColumnHeadersHeight
        End Get
        Set(ByVal value As Integer)
            intColumnHeadersHeight = value
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

#Region "Create Risks table"
    Public Sub CreateList()
        Dim strSortNumber As String

        objPrintList.Clear()
        objPrintListNotDefined.Clear()
        objPrintListOperational.Clear()
        objPrintListFinancial.Clear()
        objPrintListObjectives.Clear()
        objPrintListReputation.Clear()
        objPrintListOther.Clear()

        CreateList_SectionTitles()

        'assumptions of goals
        For Each selGoal As Goal In Me.LogFrame.Goals
            strSortNumber = Me.LogFrame.GetStructSortNumber(selGoal)
            CreateList_Assumptions(selGoal, strSortNumber)
        Next

        'assumptions of purposes
        For Each selPurpose As Purpose In Me.LogFrame.Purposes
            strSortNumber = Me.LogFrame.GetStructSortNumber(selPurpose)

            CreateList_Assumptions(selPurpose, strSortNumber)

            'assumptions of outputs
            For Each selOutput As Output In selPurpose.Outputs
                strSortNumber = Me.LogFrame.GetStructSortNumber(selOutput)

                CreateList_Assumptions(selOutput, strSortNumber)

                'assumptions of activities
                CreateList_Activities(selOutput.Activities)
            Next
        Next
        CreateList_NoRisksRow()

        PrintList.AddRange(objPrintListOperational)
        PrintList.AddRange(objPrintListFinancial)
        PrintList.AddRange(objPrintListObjectives)
        PrintList.AddRange(objPrintListReputation)
        PrintList.AddRange(objPrintListNotDefined)
        PrintList.AddRange(objPrintListOther)
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities)
        Dim strSortNumber As String

        For Each selActivity As Activity In selActivities
            strSortNumber = Me.LogFrame.GetStructSortNumber(selActivity)
            CreateList_Assumptions(selActivity, strSortNumber)

            If selActivity.Activities.Count > 0 Then
                CreateList_Activities(selActivity.Activities)
            End If
        Next
    End Sub

    Private Sub CreateList_SectionTitles()
        If RiskCategory = RiskCategories.NotDefined Or RiskCategory = RiskCategories.All Then _
            objPrintListNotDefined.Add(New PrintRiskRegisterRow(LANG_RisksNotDefined))
        If RiskCategory = RiskCategories.Operational Or RiskCategory = RiskCategories.All Then _
            objPrintListOperational.Add(New PrintRiskRegisterRow(LANG_RisksOperational))
        If RiskCategory = RiskCategories.Financial Or RiskCategory = RiskCategories.All Then _
            objPrintListFinancial.Add(New PrintRiskRegisterRow(LANG_RisksFinancial))
        If RiskCategory = RiskCategories.Objectives Or RiskCategory = RiskCategories.All Then _
            objPrintListObjectives.Add(New PrintRiskRegisterRow(LANG_RisksObjectives))
        If RiskCategory = RiskCategories.Reputation Or RiskCategory = RiskCategories.All Then _
            objPrintListReputation.Add(New PrintRiskRegisterRow(LANG_RisksReputation))
        If RiskCategory = RiskCategories.Other Or RiskCategory = RiskCategories.All Then _
            objPrintListOther.Add(New PrintRiskRegisterRow(LANG_RisksOther))
    End Sub

    Private Sub CreateList_Assumptions(ByVal selStruct As Struct, ByVal strSortNumber As String)
        Dim i As Integer
        Dim strAsmSortNumber As String
        For Each selAssumption As Assumption In selStruct.Assumptions
            If selAssumption.RaidType = Assumption.RaidTypes.Risk And selAssumption.RiskDetail IsNot Nothing Then
                strAsmSortNumber = LogFrame.CreateSortNumber(i, strSortNumber)
                Dim objRow As New PrintRiskRegisterRow(strAsmSortNumber, selAssumption, strSortNumber, selStruct)

                Select Case selAssumption.RiskDetail.RiskCategory
                    Case RiskDetail.RiskCategories.NotDefined
                        If RiskCategory = RiskCategories.NotDefined Or RiskCategory = RiskCategories.All Then objPrintListNotDefined.Add(objRow)
                    Case RiskDetail.RiskCategories.Operational
                        If RiskCategory = RiskCategories.Operational Or RiskCategory = RiskCategories.All Then objPrintListOperational.Add(objRow)
                    Case RiskDetail.RiskCategories.Financial
                        If RiskCategory = RiskCategories.Financial Or RiskCategory = RiskCategories.All Then objPrintListFinancial.Add(objRow)
                    Case RiskDetail.RiskCategories.Objectives
                        If RiskCategory = RiskCategories.Objectives Or RiskCategory = RiskCategories.All Then objPrintListObjectives.Add(objRow)
                    Case RiskDetail.RiskCategories.Reputation
                        If RiskCategory = RiskCategories.Reputation Or RiskCategory = RiskCategories.All Then objPrintListReputation.Add(objRow)
                    Case RiskDetail.RiskCategories.Other
                        If RiskCategory = RiskCategories.Other Or RiskCategory = RiskCategories.All Then objPrintListOther.Add(objRow)
                End Select
            End If

            i += 1
        Next
    End Sub

    Private Sub CreateList_NoRisksRow()
        Dim objAssumption As New Assumption(TextToRichText(LANG_RisksNoRisks))
        Dim objRow As New PrintRiskRegisterRow

        objRow.Assumption = objAssumption

        'if the list has only a title row (count=1) then add an row with an "no risks" message
        If objPrintListNotDefined.Count = 1 Then objPrintListNotDefined.Add(objRow)
        If objPrintListOperational.Count = 1 Then objPrintListOperational.Add(objRow)
        If objPrintListFinancial.Count = 1 Then objPrintListFinancial.Add(objRow)
        If objPrintListObjectives.Count = 1 Then objPrintListObjectives.Add(objRow)
        If objPrintListReputation.Count = 1 Then objPrintListReputation.Add(objRow)
        If objPrintListOther.Count = 1 Then objPrintListOther.Add(objRow)
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth * PagesWide
        Dim intAssumptionSortWidth, intStructSortWidth
        Dim intTextColumnWidth, intSortWidth As Integer

        SectionBorder = LeftMargin + Me.ContentWidth

        'calculate column widths
        intStructSortWidth = GetStructSortColumnWidth()
        intAssumptionSortWidth = GetAssumptionSortColumnWidth()

        intSortWidth = intStructSortWidth + intAssumptionSortWidth
        intTextColumnWidth = intAvailableWidth - intSortWidth
        intTextColumnWidth /= 4

        If intTextColumnWidth < CONST_MinTextColumnWidth Then intTextColumnWidth = CONST_MinTextColumnWidth
        If intTextColumnWidth + intStructSortWidth > ContentWidth Then intTextColumnWidth = ContentWidth - intStructSortWidth

        'set assumption column rectangles
        AssumptionSortRectangle = New PrintRectangle(LeftMargin, ContentTop, intAssumptionSortWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(AssumptionSortRectangle)

        AssumptionRectangle = New PrintRectangle(AssumptionSortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(AssumptionRectangle)

        If AssumptionRectangle.Left < intSectionBorder And AssumptionRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(AssumptionRectangle)
        End If

        'set risk response column rectangle
        RiskResponseRectangle = New PrintRectangle(AssumptionRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(RiskResponseRectangle)

        If RiskResponseRectangle.Left <= intSectionBorder And RiskResponseRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(RiskResponseRectangle)
        End If

        'set struct column rectangles
        StructSortRectangle = New PrintRectangle(RiskResponseRectangle.Right, ContentTop, intStructSortWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(StructSortRectangle)

        If StructSortRectangle.Left <= intSectionBorder And StructSortRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(StructSortRectangle)
        End If

        StructRectangle = New PrintRectangle(StructSortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(StructRectangle)

        If StructRectangle.Left <= intSectionBorder And StructRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(StructRectangle)
        End If

        'set risk level column rectangle
        RiskLevelRectangle = New PrintRectangle(StructRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(RiskLevelRectangle)

        If RiskLevelRectangle.Left <= intSectionBorder And RiskLevelRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(RiskLevelRectangle)
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

    Private Function GetStructSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintRiskRegisterRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.StructSortNumber) = False AndAlso selRow.StructSortNumber.Length > strSort.Length Then strSort = selRow.StructSortNumber
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.DetailsFont).Width + (CONST_HorizontalPadding * 2)
        End If
        Return intWidth
    End Function

    Private Function GetAssumptionSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintRiskRegisterRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.AssumptionSortNumber) = False AndAlso selRow.AssumptionSortNumber.Length > strSort.Length Then strSort = selRow.AssumptionSortNumber
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.DetailsFont).Width + (CONST_HorizontalPadding * 2)
        End If
        Return intWidth
    End Function
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintRiskRegisterRow In Me.PrintList
            ReloadImages_Normal(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintRiskRegisterRow)
        With RichTextManager
            If selRow.Assumption IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Assumption.Text) Then
                    selRow.Assumption.CellImage = .EmptyTextWithPaddingToBitmap(AssumptionRectangle.Width, Assumption.ItemName, selRow.AssumptionSortNumber, False)
                Else
                    selRow.Assumption.CellImage = .RichTextWithPaddingToBitmap(AssumptionRectangle.Width, selRow.Assumption.RTF, False)
                End If
            End If
            If String.IsNullOrEmpty(selRow.RiskResponse) = False Then
                selRow.RiskResponseImage = .RichTextWithPaddingToBitmap(RiskResponseRectangle.Width, selRow.RiskResponse, False)
            End If
            If selRow.Struct IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Struct.Text) Then
                    selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.GetItemName(), selRow.StructSortNumber, False)
                Else
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.RTF, False)
                End If
            End If
            If String.IsNullOrEmpty(selRow.RiskLevel) = False Then
                selRow.RiskLevelImage = .RichTextWithPaddingToBitmap(RiskLevelRectangle.Width, selRow.RiskLevel, False)
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintRiskRegisterRow = Me.PrintList(RowIndex)
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
        Dim selRow As PrintRiskRegisterRow = Me.PrintList(RowIndex)

        If selRow.Assumption IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Assumption.RTF) = False Then
            If selRow.Assumption.CellImage.Height > intRowHeight Then intRowHeight = selRow.Assumption.CellImage.Height
        End If
        If String.IsNullOrEmpty(selRow.RiskResponse) = False Then
            If selRow.RiskResponseImage.Height > intRowHeight Then intRowHeight = selRow.RiskResponseImage.Height
        End If
        If selRow.Struct IsNot Nothing Then
            If selRow.Struct.CellImage.Height > intRowHeight Then intRowHeight = selRow.Struct.CellImage.Height
        End If
        If String.IsNullOrEmpty(selRow.RiskLevel) = False Then
            If selRow.RiskLevelImage.Height > intRowHeight Then intRowHeight = selRow.RiskLevelImage.Height
        End If

        Return intRowHeight
    End Function

    Private Sub SetColumnHeadersHeight()
        Dim intHeight As Integer

        If PageGraph IsNot Nothing Then
            Dim fntHeader As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)

            Dim intAssumptionSortHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, AssumptionSortRectangle.Width).Height
            If intAssumptionSortHeight > intHeight Then intHeight = intAssumptionSortHeight

            Dim intAssumptionHeight As Integer = PageGraph.MeasureString(LANG_Assumption, fntHeader, AssumptionRectangle.Width).Height
            If intAssumptionHeight > intHeight Then intHeight = intAssumptionHeight

            Dim intRiskResponseHeight As Integer = PageGraph.MeasureString(LANG_RiskResponse, fntHeader, RiskResponseRectangle.Width).Height
            If intRiskResponseHeight > intHeight Then intHeight = intRiskResponseHeight

            Dim intStructHeight As Integer = PageGraph.MeasureString(LogFrame.StructNamePlural(2), fntHeader, StructRectangle.Width).Height
            If intStructHeight > intHeight Then intHeight = intStructHeight

            Dim intRiskLevelHeight As Integer = PageGraph.MeasureString(LANG_RiskLevel, fntHeader, RiskLevelRectangle.Width).Height
            If intRiskLevelHeight > intHeight Then intHeight = intRiskLevelHeight

            ColumnHeadersHeight = intHeight
        End If
    End Sub
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight - ColumnHeadersHeight

        For Each selRow As PrintRiskRegisterRow In PrintList
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
        Dim selRow As PrintRiskRegisterRow = PrintList(RowIndex)
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

                If String.IsNullOrEmpty(selRow.SectionTitle) = False Then
                    Dim rSectionTitle As New Rectangle(LeftMargin, LastRowY, LastColumnBorder - LeftMargin, selRow.RowHeight)

                    PrintPage_PrintText(selRow.SectionTitle, rSectionTitle, False, True)
                Else
                    PrintPage_PrintAssumptionSortNumber(selRow)
                    PrintPage_PrintAssumption(selRow)
                    PrintPage_PrintRiskResponse(selRow)
                    PrintPage_PrintStructSortNumber(selRow)
                    PrintPage_PrintStruct(selRow)
                    PrintPage_PrintRiskLevel(selRow)
                End If

                PrintPage_PrintLines(selRow)

                If LastRowY < ContentBottom And LastRowY + selRow.RowHeight > ContentBottom Then
                    If ClippedRow IsNot Nothing Then
                        PrintList.Insert(RowIndex + 1, ClippedRow)

                        ReloadImages_Normal(ClippedRow)
                        SetRowHeight(RowIndex + 1)
                        ClippedRow.RowHeight += CONST_VerticalPadding
                    End If
                End If

                LastRowY += selRow.RowHeight
                RowIndex += 1
                If RowIndex > PrintList.Count - 1 Then
                    Exit Do
                End If

                intMinHeight = LastRowY + PrintPage_GetMinHeight(PrintList(RowIndex))

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

    Private Sub PrintPage_PrintText(ByVal strValue As String, ByVal rCell As Rectangle, Optional ByVal boolHeader As Boolean = False, Optional ByVal boolSectionTitle As Boolean = False)
        If PageGraph IsNot Nothing Then
            With PageGraph
                If boolHeader = True Then
                    .FillRectangle(Brushes.LightGray, rCell)
                    .DrawRectangle(penBlack1, rCell)
                ElseIf boolSectionTitle = True Then
                    .FillRectangle(Brushes.LightSlateGray, rCell)
                    .DrawRectangle(penBlack1, rCell)
                Else
                    .DrawLine(penBlack1, rCell.Left, rCell.Top, rCell.Right, rCell.Top)
                    '.DrawLine(penBlack1, rCell.Right, rCell.Top, rCell.Right, rCell.Bottom)
                End If

                Dim formatCells As New StringFormat()
                Dim brText As SolidBrush = New SolidBrush(Color.Black)
                Dim fntText As Font
                Dim rText As New Rectangle(rCell.X + CONST_Padding, rCell.Y + CONST_Padding, rCell.Width - CONST_HorizontalPadding, rCell.Height)

                If boolHeader = True Then
                    formatCells.Alignment = StringAlignment.Center
                    formatCells.LineAlignment = StringAlignment.Center
                    fntText = New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
                ElseIf boolSectionTitle = True Then
                    formatCells.Alignment = StringAlignment.Near
                    formatCells.LineAlignment = StringAlignment.Center
                    fntText = New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
                Else
                    formatCells.Alignment = StringAlignment.Near
                    formatCells.LineAlignment = StringAlignment.Near
                    fntText = New Font(CurrentLogFrame.DetailsFont, FontStyle.Regular)
                End If

                .DrawString(strValue, fntText, brText, rText, formatCells)
            End With
        End If
    End Sub

    Private Sub PrintPage_PrintLines(ByVal selRow As PrintRiskRegisterRow)
        Dim intTop As Integer = LastRowY
        Dim intBottom As Integer = LastRowY + selRow.RowHeight
        Dim intLeft, intRight As Integer

        If intBottom > ContentBottom Then intBottom = ContentBottom

        PageGraph.DrawLine(penBlack1, Me.LeftMargin, intTop, Me.LeftMargin, intBottom + 1)

        If AssumptionRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(AssumptionRectangle.X)
            intRight = intLeft + AssumptionRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If RiskResponseRectangle.Left >= CurrentSectionRectangle.Left And RiskResponseRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(RiskResponseRectangle.X)
            intRight = intLeft + RiskResponseRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If StructRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(StructRectangle.X)
            intRight = intLeft + StructRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If RiskLevelRectangle.Left >= CurrentSectionRectangle.Left And RiskLevelRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(RiskLevelRectangle.X)
            intRight = intLeft + RiskLevelRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If
    End Sub
#End Region

#Region "Clipped text"
    Private Function PrintPage_GetMinHeight(ByVal selRow As PrintRiskRegisterRow)
        Dim intMinHeight As Integer

        With RichTextManager
            If selRow.Assumption IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(AssumptionRectangle.Width, selRow.Assumption.RTF)
            ElseIf String.IsNullOrEmpty(selRow.RiskResponse) = False Then
                intMinHeight = .GetFirstLineSpacing(RiskResponseRectangle.Width, selRow.RiskResponse)
            ElseIf selRow.Struct IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(StructRectangle.Width, selRow.Struct.RTF)
            ElseIf String.IsNullOrEmpty(selRow.RiskLevel) = False Then
                intMinHeight = .GetFirstLineSpacing(RiskLevelRectangle.Width, selRow.RiskLevel)
            Else
                intMinHeight = NewCellHeight()
            End If
        End With

        Return intMinHeight
    End Function

    Private Sub PrintClippedText(ByVal strRtf As String, ByVal rImage As Rectangle, ByVal intColumnWidth As Integer)
        With RichTextManager
            Dim bmClip As Bitmap = .PrintClippedRichText(intColumnWidth, strRtf, ContentBottom - rImage.Y)
            strClippedTextTop = .ClippedTextTop
            strClippedTextBottom = .ClippedTextBottom

            rImage.X += 2
            rImage.Y += 2
            rImage.Height = bmClip.Height
            PageGraph.DrawImage(bmClip, rImage)
        End With
    End Sub
#End Region

#Region "Print assumptions"
    Private Sub PrintPage_PrintAssumptionSortNumber(ByVal selRow As PrintRiskRegisterRow)
        If AssumptionSortRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(AssumptionSortRectangle.X)
            Dim rSortNumber As New Rectangle(intX, LastRowY, AssumptionSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.AssumptionSortNumber, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintAssumption(ByVal selRow As PrintRiskRegisterRow)
        If AssumptionSortRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(AssumptionRectangle.X)

            If selRow.Assumption IsNot Nothing AndAlso selRow.Assumption.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Assumption.CellImage.Width, selRow.Assumption.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Assumption.CellImage, rImage)
                Else
                    PrintClippedText(selRow.Assumption.RTF, rImage, AssumptionRectangle.Width)
                    selRow.Assumption.RTF = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintRiskRegisterRow()
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.Assumption = New Assumption(strClippedTextBottom)
                End If
            End If
        End If

        If selRow.Assumption IsNot Nothing AndAlso (String.IsNullOrEmpty(selRow.Assumption.RTF) = False And selRow.ClippedRow = False) Then
            Dim intSortX As Integer = GetCoordinateX(AssumptionSortRectangle.X)
            If intSortX < ContentLeft Then intSortX = ContentLeft
            PageGraph.DrawLine(penBlack1, intSortX, LastRowY, LastColumnBorder, LastRowY)
        End If
    End Sub
#End Region

#Region "Risk response"
    Private Sub PrintPage_PrintRiskResponse(ByVal selRow As PrintRiskRegisterRow)
        If RiskResponseRectangle.Left >= CurrentSectionRectangle.Left And RiskResponseRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(RiskResponseRectangle.X)

            If String.IsNullOrEmpty(selRow.RiskResponse) = False AndAlso selRow.RiskResponseImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.RiskResponseImage.Width, selRow.RiskResponseImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.RiskResponseImage, rImage)
                Else
                    PrintClippedText(selRow.RiskResponse, rImage, RiskResponseRectangle.Width)
                    selRow.RiskResponse = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintRiskRegisterRow()
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.RiskResponse = strClippedTextBottom
                End If
            End If
        End If
    End Sub
#End Region

#Region "Print struct"
    Private Sub PrintPage_PrintStructSortNumber(ByVal selRow As PrintRiskRegisterRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.StructSortNumber, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintStruct(ByVal selRow As PrintRiskRegisterRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(StructRectangle.X)

            If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Struct.CellImage.Width, selRow.Struct.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Struct.CellImage, rImage)
                Else
                    PrintClippedText(selRow.Struct.RTF, rImage, StructRectangle.Width)
                    selRow.Struct.RTF = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintRiskRegisterRow()
                        ClippedRow.ClippedRow = True
                    End If
                    Select Case selRow.Struct.GetType
                        Case GetType(Goal)
                            ClippedRow.Struct = New Goal(strClippedTextBottom)
                        Case GetType(Purpose)
                            ClippedRow.Struct = New Purpose(strClippedTextBottom)
                        Case GetType(Output)
                            ClippedRow.Struct = New Output(strClippedTextBottom)
                        Case GetType(Activity)
                            ClippedRow.Struct = New Activity(strClippedTextBottom)
                    End Select
                End If
            End If
        End If
        If selRow.Struct IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Struct.RTF) = False Then
            PageGraph.DrawLine(penBlack1, LeftMargin, LastRowY, LastColumnBorder, LastRowY)
        End If
    End Sub
#End Region

#Region "Risk level"
    Private Sub PrintPage_PrintRiskLevel(ByVal selRow As PrintRiskRegisterRow)
        If RiskLevelRectangle.Left >= CurrentSectionRectangle.Left And RiskLevelRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(RiskLevelRectangle.X)

            If String.IsNullOrEmpty(selRow.RiskLevel) = False AndAlso selRow.RiskLevelImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.RiskLevelImage.Width, selRow.RiskLevelImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.RiskLevelImage, rImage)
                Else
                    PrintClippedText(selRow.RiskLevel, rImage, RiskLevelRectangle.Width)
                    selRow.RiskLevel = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintRiskRegisterRow()
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.RiskLevel = strClippedTextBottom
                End If
            End If
        End If
    End Sub
#End Region

#Region "Column headers"
    Private Sub PrintColumnHeaders()
        Dim intX As Integer

        If PageGraph IsNot Nothing Then
            If AssumptionSortRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(AssumptionSortRectangle.X)
                Dim rAssumption As New Rectangle(intX, LastRowY, AssumptionSortRectangle.Width + AssumptionRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Assumptions, rAssumption, True, False)

                LastColumnBorder = rAssumption.Right
            End If

            If RiskResponseRectangle.Left >= CurrentSectionRectangle.Left And RiskResponseRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(RiskResponseRectangle.X)
                Dim rRiskResponse As New Rectangle(intX, LastRowY, RiskResponseRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_RiskResponse, rRiskResponse, True, False)

                LastColumnBorder = rRiskResponse.Right
            End If

            If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(StructSortRectangle.X)
                Dim rStruct As New Rectangle(intX, LastRowY, StructSortRectangle.Width + StructRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Objectives, rStruct, True, False)

                LastColumnBorder = rStruct.Right
            End If

            If RiskLevelRectangle.Left >= CurrentSectionRectangle.Left And RiskLevelRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(RiskLevelRectangle.X)
                Dim rRiskLevel As New Rectangle(intX, LastRowY, RiskLevelRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_RiskLevel, rRiskLevel, True, False)

                LastColumnBorder = rRiskLevel.Right
            End If
        End If

        LastRowY += ColumnHeadersHeight
    End Sub
#End Region
End Class
