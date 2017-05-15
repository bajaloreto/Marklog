Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintDependencies
    Inherits ReportBaseClass

    Private Const CONST_MinTextColumnWidth As Integer = 100
    Private Const CONST_DateColumnWidth As Integer = 75

    Private objLogFrame As New LogFrame
    Private objPrintList As New PrintDependencyRows
    Private objClippedRow As PrintDependencyRow = Nothing
    Private intPrintSection As Integer
    Private boolShowDeliverables As Boolean
    Private intPagesWide As Integer

    Private strClippedTextTop, strClippedTextBottom As String

    Private boolColumnsWidthSet As Boolean
    Private CurrentSectionRectangle As New Rectangle
    Private rStructSortRectangle, rAssumptionSortRectangle As New PrintRectangle
    Private rStructRectangle, rAssumptionRectangle As New PrintRectangle
    Private rDependencyTypeRectangle, rInputTypeRectangle, rImportanceLevelRectangle, rImpactRectangle As New PrintRectangle
    Private rSupplierRectangle, rDeliverableTypeRectangle, rDateExpectedRectangle, rDateDeliveredRectangle, rDeliverablesRectangle As New PrintRectangle
    Private rResponseStrategyRectangle, rOwnerRectangle As New PrintRectangle

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

    Public Enum PrintSections As Integer
        NotSelected = -1
        PrintGoals = 0
        PrintPurposes = 1
        PrintOutputs = 2
        PrintActivities = 3
        PrintAll = 4
    End Enum

    Public Sub New(ByVal logframe As LogFrame, ByVal printsection As Integer, ByVal showdeliverables As Boolean, ByVal pageswide As Integer)
        Me.LogFrame = logframe
        Me.ReportSetup = logframe.ReportSetupDependencies
        Me.PrintSection = printsection
        Me.ShowDeliverables = showdeliverables
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

    Public Property PrintList() As PrintDependencyRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintDependencyRows)
            objPrintList = value
        End Set
    End Property

    Public Property ClippedRow As PrintDependencyRow
        Get
            Return objClippedRow
        End Get
        Set(ByVal value As PrintDependencyRow)
            objClippedRow = value
        End Set
    End Property

    Public Property PrintSection() As Integer
        Get
            Return intPrintSection
        End Get
        Set(ByVal value As Integer)
            intPrintSection = value
        End Set
    End Property

    Public Property ShowDeliverables As Boolean
        Get
            Return boolShowDeliverables
        End Get
        Set(ByVal value As Boolean)
            boolShowDeliverables = value
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

    Private Property DependencyTypeRectangle As PrintRectangle
        Get
            Return rDependencyTypeRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rDependencyTypeRectangle = value
        End Set
    End Property

    Private Property InputTypeRectangle As PrintRectangle
        Get
            Return rInputTypeRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rInputTypeRectangle = value
        End Set
    End Property

    Private Property ImportanceLevelRectangle As PrintRectangle
        Get
            Return rImportanceLevelRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rImportanceLevelRectangle = value
        End Set
    End Property

    Private Property ImpactRectangle As PrintRectangle
        Get
            Return rImpactRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rImpactRectangle = value
        End Set
    End Property

    Private Property SupplierRectangle As PrintRectangle
        Get
            Return rSupplierRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rSupplierRectangle = value
        End Set
    End Property

    Private Property DeliverableTypeRectangle As PrintRectangle
        Get
            Return rDeliverableTypeRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rDeliverableTypeRectangle = value
        End Set
    End Property

    Private Property DateExpectedRectangle As PrintRectangle
        Get
            Return rDateExpectedRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rDateExpectedRectangle = value
        End Set
    End Property

    Private Property DateDeliveredRectangle As PrintRectangle
        Get
            Return rDateDeliveredRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rDateDeliveredRectangle = value
        End Set
    End Property

    Private Property DeliverablesRectangle As PrintRectangle
        Get
            Return rDeliverablesRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rDeliverablesRectangle = value
        End Set
    End Property

    Private Property ResponseStrategyRectangle As PrintRectangle
        Get
            Return rResponseStrategyRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rResponseStrategyRectangle = value
        End Set
    End Property

    Private Property OwnerRectangle As PrintRectangle
        Get
            Return rOwnerRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rOwnerRectangle = value
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

#Region "Create Dependencies table"
    Public Sub CreateList()
        Dim strGoalSort As String = String.Empty, strPurposeSort As String = String.Empty
        Dim strOutputSort As String = String.Empty

        PrintList.Clear()

        'add goals
        If Me.PrintSection = PrintSections.PrintGoals Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selGoal As Goal In Me.LogFrame.Goals
                strGoalSort = LogFrame.CreateSortNumber(LogFrame.Goals.IndexOf(selGoal))
                Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.GoalsSection, strGoalSort, New Goal(selGoal.RTF))
                PrintList.Add(NewRow)

                CreateList_Dependencies(LogFrame.SectionTypes.GoalsSection, selGoal.Assumptions, strGoalSort)
            Next
        End If

        'add outputs
        If Me.PrintSection = PrintSections.PrintPurposes Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))

                If Me.PrintSection = PrintSections.PrintPurposes Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
                    PrintList.Add(NewRow)

                    CreateList_Dependencies(LogFrame.SectionTypes.PurposesSection, selPurpose.Assumptions, strPurposeSort)
                End If
            Next
        End If

        'add outputs
        If Me.PrintSection = PrintSections.PrintOutputs Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))

                If My.Settings.setRepeatPurposes = True Then
                    Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
                    PrintList.Add(NewRow)
                End If

                'add outputs
                CreateList_Outputs(selPurpose.Outputs, strPurposeSort)
            Next
        End If

        'add activities
        If Me.PrintSection = PrintSections.PrintActivities Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))

                If My.Settings.setRepeatPurposes = True Then
                    Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
                    PrintList.Add(NewRow)
                End If

                For Each selOutput As Output In selPurpose.Outputs
                    strOutputSort = LogFrame.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)

                    If My.Settings.setRepeatOutputs = True Then
                        Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.OutputsSection, strOutputSort, New Output(selOutput.RTF))
                        PrintList.Add(NewRow)
                    End If

                    CreateList_Activities(selOutput.Activities, strOutputSort)
                Next
            Next
        End If
    End Sub

    Private Sub CreateList_Outputs(ByVal selOutputs As Outputs, ByVal strParentSort As String)
        Dim strOutputSort As String = String.Empty

        If Me.PrintSection > PrintSections.PrintPurposes Then
            For Each selOutput As Output In selOutputs
                strOutputSort = LogFrame.CreateSortNumber(selOutputs.IndexOf(selOutput), strParentSort)

                If Me.PrintSection = PrintSections.PrintOutputs Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.OutputsSection, strOutputSort, New Output(selOutput.RTF))
                    PrintList.Add(NewRow)

                    CreateList_Dependencies(LogFrame.SectionTypes.OutputsSection, selOutput.Assumptions, strOutputSort)
                End If
            Next
        End If
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String)
        Dim strActivitySort As String = String.Empty

        If Me.PrintSection > PrintSections.PrintOutputs Then
            For Each selActivity As Activity In selActivities
                strActivitySort = LogFrame.CreateSortNumber(selActivities.IndexOf(selActivity), strParentSort)

                If Me.PrintSection = PrintSections.PrintActivities Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintDependencyRow(LogFrame.SectionTypes.ActivitiesSection, strActivitySort, New Activity(selActivity.RTF))
                    PrintList.Add(NewRow)

                    CreateList_Dependencies(LogFrame.SectionTypes.ActivitiesSection, selActivity.Assumptions, strActivitySort)

                    If selActivity.Activities.Count > 0 Then
                        CreateList_Activities(selActivity.Activities, strActivitySort)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub CreateList_Dependencies(ByVal intSection As Integer, ByVal selAssumptions As Assumptions, ByVal strParentSort As String)
        Dim strAssumptionSort As String
        Dim intIndex As Integer
        Dim objRow As PrintDependencyRow

        If selAssumptions.Count > 0 Then
            For Each selAssumption As Assumption In selAssumptions
                If selAssumption.RaidType = Assumption.RaidTypes.Dependency Then
                    Dim objAssumption As Assumption
                    strAssumptionSort = LogFrame.CreateSortNumber(intIndex, strParentSort)

                    Using objCopier As New ObjectCopy
                        objAssumption = objCopier.CopyObject(selAssumption)
                    End Using

                    objRow = New PrintDependencyRow(intSection, strAssumptionSort, objAssumption)
                    PrintList.Add(objRow)
                End If
                intIndex += 1
            Next
        End If
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth * PagesWide
        Dim intAssumptionSortWidth, intStructSortWidth
        Dim intTextColumnWidth As Integer
        Dim intColumnX As Integer
        Dim intTextColumnCount As Integer = 7

        SectionBorder = Me.ContentRight

        'calculate column widths
        If ShowDeliverables = True Then intTextColumnCount += 3

        intStructSortWidth = GetStructSortColumnWidth()
        intAssumptionSortWidth = GetAssumptionSortColumnWidth()

        If intStructSortWidth < intAssumptionSortWidth Then intStructSortWidth = intAssumptionSortWidth
        If intAssumptionSortWidth < intStructSortWidth Then intAssumptionSortWidth = intStructSortWidth

        intTextColumnWidth = intAvailableWidth - intAssumptionSortWidth
        intTextColumnWidth /= intTextColumnCount

        If intTextColumnWidth < CONST_MinTextColumnWidth Then intTextColumnWidth = CONST_MinTextColumnWidth
        If intTextColumnWidth + intAssumptionSortWidth > ContentWidth Then intTextColumnWidth = ContentWidth - intAssumptionSortWidth

        'set struct column rectangles
        StructSortRectangle = New PrintRectangle(LeftMargin, ContentTop, intStructSortWidth, ContentHeight, intHorizontalPageIndex)
        StructRectangle = New PrintRectangle(StructSortRectangle.Right, ContentTop, ContentWidth - intStructSortWidth, ContentHeight, intHorizontalPageIndex, True)

        'set assumption column rectangles
        AssumptionSortRectangle = New PrintRectangle(LeftMargin, ContentTop, intAssumptionSortWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(AssumptionSortRectangle)

        AssumptionRectangle = New PrintRectangle(AssumptionSortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(AssumptionRectangle)

        If AssumptionRectangle.Left < intSectionBorder And AssumptionRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(AssumptionRectangle)
        End If

        'set dependency type column rectangle
        DependencyTypeRectangle = New PrintRectangle(AssumptionRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(DependencyTypeRectangle)

        If DependencyTypeRectangle.Left <= intSectionBorder And DependencyTypeRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(DependencyTypeRectangle)
        End If

        'set input type column rectangle
        InputTypeRectangle = New PrintRectangle(DependencyTypeRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(InputTypeRectangle)

        If InputTypeRectangle.Left <= intSectionBorder And InputTypeRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(InputTypeRectangle)
        End If

        'set importance level column rectangle
        ImportanceLevelRectangle = New PrintRectangle(InputTypeRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(ImportanceLevelRectangle)

        If ImportanceLevelRectangle.Left <= intSectionBorder And ImportanceLevelRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(ImportanceLevelRectangle)
        End If

        'set impact column rectangle
        ImpactRectangle = New PrintRectangle(ImportanceLevelRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(ImpactRectangle)

        If ImpactRectangle.Left <= intSectionBorder And ImpactRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(ImpactRectangle)
        End If

        intColumnX = ImpactRectangle.Right

        If ShowDeliverables = True Then
            'set supplier column rectangle
            SupplierRectangle = New PrintRectangle(ImpactRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
            PrintRectangles.Add(SupplierRectangle)

            If SupplierRectangle.Left <= intSectionBorder And SupplierRectangle.Right > intSectionBorder Then
                StretchRectanglesOfPreviousPage(SupplierRectangle)
            End If

            'set deliverable type column rectangle
            DeliverableTypeRectangle = New PrintRectangle(SupplierRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
            PrintRectangles.Add(DeliverableTypeRectangle)

            If DeliverableTypeRectangle.Left <= intSectionBorder And DeliverableTypeRectangle.Right > intSectionBorder Then
                StretchRectanglesOfPreviousPage(DeliverableTypeRectangle)
            End If

            'set expected delivery date column rectangle
            DateExpectedRectangle = New PrintRectangle(DeliverableTypeRectangle.Right, ContentTop, CONST_DateColumnWidth, ContentHeight, intHorizontalPageIndex)
            PrintRectangles.Add(DateExpectedRectangle)

            If DateExpectedRectangle.Left <= intSectionBorder And DateExpectedRectangle.Right > intSectionBorder Then
                StretchRectanglesOfPreviousPage(DateExpectedRectangle)
            End If

            'set delivery date column rectangle
            DateDeliveredRectangle = New PrintRectangle(DateExpectedRectangle.Right, ContentTop, CONST_DateColumnWidth, ContentHeight, intHorizontalPageIndex)
            PrintRectangles.Add(DateDeliveredRectangle)

            If DateDeliveredRectangle.Left <= intSectionBorder And DateDeliveredRectangle.Right > intSectionBorder Then
                StretchRectanglesOfPreviousPage(DateDeliveredRectangle)
            End If

            'set deliverables column rectangle
            DeliverablesRectangle = New PrintRectangle(DateDeliveredRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
            PrintRectangles.Add(DeliverablesRectangle)

            If DeliverablesRectangle.Left <= intSectionBorder And DeliverablesRectangle.Right > intSectionBorder Then
                StretchRectanglesOfPreviousPage(DeliverablesRectangle)
            End If

            intColumnX = DeliverablesRectangle.Right
        End If

        'set response strategy column rectangle
        ResponseStrategyRectangle = New PrintRectangle(intColumnX, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(ResponseStrategyRectangle)

        If ResponseStrategyRectangle.Left <= intSectionBorder And ResponseStrategyRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(ResponseStrategyRectangle)
        End If

        'set owner column rectangle
        OwnerRectangle = New PrintRectangle(ResponseStrategyRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(OwnerRectangle)

        If OwnerRectangle.Left <= intSectionBorder And OwnerRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(OwnerRectangle)
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

        For Each selRow As PrintDependencyRow In Me.PrintList
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

        For Each selRow As PrintDependencyRow In Me.PrintList
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
        For Each selRow As PrintDependencyRow In Me.PrintList
            ReloadImages_Normal(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintDependencyRow)
        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Struct.Text) Then
                    selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.GetItemName(), selRow.StructSortNumber, False)
                Else
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.RTF, False)
                End If
            ElseIf selRow.Assumption IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Assumption.Text) Then
                    selRow.Assumption.CellImage = .EmptyTextWithPaddingToBitmap(AssumptionRectangle.Width, Assumption.ItemName, selRow.AssumptionSortNumber, False)
                Else
                    selRow.Assumption.CellImage = .RichTextWithPaddingToBitmap(AssumptionRectangle.Width, selRow.Assumption.RTF, False)
                End If

                If String.IsNullOrEmpty(selRow.DependencyTypeText) = False Then
                    selRow.DependencyTypeImage = .TextToBitmap(DependencyTypeRectangle.Width, selRow.DependencyTypeText)
                End If
                If String.IsNullOrEmpty(selRow.InputTypeText) = False Then
                    selRow.InputTypeImage = .TextToBitmap(InputTypeRectangle.Width, selRow.InputTypeText)
                End If
                If String.IsNullOrEmpty(selRow.ImportanceLevelText) = False Then
                    selRow.ImportanceLevelImage = .TextToBitmap(ImportanceLevelRectangle.Width, selRow.ImportanceLevelText)
                End If
                If String.IsNullOrEmpty(selRow.Impact) = False Then
                    selRow.ImpactImage = .TextToBitmap(ImpactRectangle.Width, selRow.Impact)
                End If

                If ShowDeliverables = True Then
                    If String.IsNullOrEmpty(selRow.Supplier) = False Then
                        selRow.SupplierImage = .TextToBitmap(SupplierRectangle.Width, selRow.Supplier)
                    End If
                    If String.IsNullOrEmpty(selRow.DeliverableTypeText) = False Then
                        selRow.DeliverableTypeImage = .TextToBitmap(DeliverableTypeRectangle.Width, selRow.DeliverableTypeText)
                    End If
                    If String.IsNullOrEmpty(selRow.Deliverables) = False Then
                        selRow.DeliverablesImage = .TextToBitmap(DeliverablesRectangle.Width, selRow.Deliverables)
                    End If
                End If

                If String.IsNullOrEmpty(selRow.ResponseStrategy) = False Then
                    selRow.ResponseStrategyImage = .TextToBitmap(ResponseStrategyRectangle.Width, selRow.ResponseStrategy)
                End If
                If String.IsNullOrEmpty(selRow.Owner) = False Then
                    selRow.OwnerImage = .TextToBitmap(OwnerRectangle.Width, selRow.Owner)
                End If
            End If

        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintDependencyRow = Me.PrintList(RowIndex)
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
        Dim selRow As PrintDependencyRow = Me.PrintList(RowIndex)

        If selRow.Struct IsNot Nothing Then
            If selRow.Struct.CellImage.Height > intRowHeight Then intRowHeight = selRow.Struct.CellImage.Height
        ElseIf selRow.Assumption IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Assumption.RTF) = False Then
            If selRow.Assumption.CellImage.Height > intRowHeight Then intRowHeight = selRow.Assumption.CellImage.Height

            If String.IsNullOrEmpty(selRow.DependencyTypeText) = False Then
                If selRow.DependencyTypeImage.Height > intRowHeight Then intRowHeight = selRow.DependencyTypeImage.Height
            End If

            If String.IsNullOrEmpty(selRow.InputTypeText) = False Then
                If selRow.InputTypeImage.Height > intRowHeight Then intRowHeight = selRow.InputTypeImage.Height
            End If

            If String.IsNullOrEmpty(selRow.ImportanceLevelText) = False Then
                If selRow.ImportanceLevelImage.Height > intRowHeight Then intRowHeight = selRow.ImportanceLevelImage.Height
            End If

            If String.IsNullOrEmpty(selRow.Impact) = False Then
                If selRow.ImpactImage.Height > intRowHeight Then intRowHeight = selRow.ImpactImage.Height
            End If

            If ShowDeliverables = True Then
                If String.IsNullOrEmpty(selRow.Supplier) = False Then
                    If selRow.SupplierImage.Height > intRowHeight Then intRowHeight = selRow.SupplierImage.Height
                End If

                If String.IsNullOrEmpty(selRow.DeliverableTypeText) = False Then
                    If selRow.DeliverableTypeImage.Height > intRowHeight Then intRowHeight = selRow.DeliverableTypeImage.Height
                End If

                If String.IsNullOrEmpty(selRow.Deliverables) = False Then
                    If selRow.DeliverablesImage.Height > intRowHeight Then intRowHeight = selRow.DeliverablesImage.Height
                End If
            End If

            If String.IsNullOrEmpty(selRow.ResponseStrategy) = False Then
                If selRow.ResponseStrategyImage.Height > intRowHeight Then intRowHeight = selRow.ResponseStrategyImage.Height
            End If

            If String.IsNullOrEmpty(selRow.Owner) = False Then
                If selRow.OwnerImage.Height > intRowHeight Then intRowHeight = selRow.OwnerImage.Height
            End If
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

            Dim intDependencyTypeHeight As Integer = PageGraph.MeasureString(LANG_DependencyType, fntHeader, DependencyTypeRectangle.Width).Height
            If intDependencyTypeHeight > intHeight Then intHeight = intDependencyTypeHeight

            Dim intInputTypeHeight As Integer = PageGraph.MeasureString(LANG_InputType, fntHeader, InputTypeRectangle.Width).Height
            If intInputTypeHeight > intHeight Then intHeight = intInputTypeHeight

            Dim intImportanceLevelHeight As Integer = PageGraph.MeasureString(LANG_ImportanceLevel, fntHeader, ImportanceLevelRectangle.Width).Height
            If intImportanceLevelHeight > intHeight Then intHeight = intImportanceLevelHeight

            Dim intImpactHeight As Integer = PageGraph.MeasureString(LANG_Impact, fntHeader, ImpactRectangle.Width).Height
            If intImpactHeight > intHeight Then intHeight = intImpactHeight

            If ShowDeliverables = True Then
                Dim intSupplierHeight As Integer = PageGraph.MeasureString(LANG_Supplier, fntHeader, SupplierRectangle.Width).Height
                If intSupplierHeight > intHeight Then intHeight = intSupplierHeight

                Dim intDeliverableTypeHeight As Integer = PageGraph.MeasureString(LANG_DeliverableType, fntHeader, DeliverableTypeRectangle.Width).Height
                If intDeliverableTypeHeight > intHeight Then intHeight = intDeliverableTypeHeight

                Dim intDateExpectedHeight As Integer = PageGraph.MeasureString(LANG_DateExpected, fntHeader, DateExpectedRectangle.Width).Height
                If intDateExpectedHeight > intHeight Then intHeight = intDateExpectedHeight

                Dim intDateDeliveredHeight As Integer = PageGraph.MeasureString(LANG_DateDelivered, fntHeader, DateDeliveredRectangle.Width).Height
                If intDateDeliveredHeight > intHeight Then intHeight = intDateDeliveredHeight

                Dim intDeliverablesHeight As Integer = PageGraph.MeasureString(LANG_Deliverables, fntHeader, DeliverablesRectangle.Width).Height
                If intDeliverablesHeight > intHeight Then intHeight = intDeliverablesHeight
            End If

            Dim intResponseStrategyHeight As Integer = PageGraph.MeasureString(LANG_ResponseStrategy, fntHeader, ResponseStrategyRectangle.Width).Height
            If intResponseStrategyHeight > intHeight Then intHeight = intResponseStrategyHeight

            Dim intOwnerHeight As Integer = PageGraph.MeasureString(LANG_Owner, fntHeader, OwnerRectangle.Width).Height
            If intOwnerHeight > intHeight Then intHeight = intOwnerHeight

            ColumnHeadersHeight = intHeight + CONST_VerticalPadding
        End If
    End Sub
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight - ColumnHeadersHeight

        For Each selRow As PrintDependencyRow In PrintList
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
        Dim selRow As PrintDependencyRow = PrintList(RowIndex)
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

                If selRow.Struct IsNot Nothing Then
                    PrintPage_PrintStructSortNumber(selRow)
                    PrintPage_PrintStruct(selRow)
                ElseIf selRow.Assumption IsNot Nothing Then
                    PrintPage_PrintAssumptionSortNumber(selRow)
                    PrintPage_PrintAssumption(selRow)
                    PrintPage_PrintDependencyType(selRow)
                    PrintPage_PrintInputType(selRow)
                    PrintPage_PrintImportanceLevel(selRow)
                    PrintPage_PrintImpact(selRow)

                    If ShowDeliverables = True Then
                        PrintPage_PrintSupplier(selRow)
                        PrintPage_PrintDeliverableType(selRow)
                        PrintPage_PrintDateExpected(selRow)
                        PrintPage_PrintDateDelivered(selRow)
                        PrintPage_PrintDeliverables(selRow)
                    End If

                    PrintPage_PrintResponseStrategy(selRow)
                    PrintPage_PrintOwner(selRow)
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

                intMinHeight = LastRowY + PrintList(RowIndex).RowHeight 'PrintPage_GetMinHeight(PrintList(RowIndex))

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
                    .DrawLine(penBlack1, rCell.Right, rCell.Top, rCell.Right, rCell.Bottom)
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

    Private Sub PrintPage_PrintLines(ByVal selRow As PrintDependencyRow)
        Dim intTop As Integer = LastRowY
        Dim intBottom As Integer = LastRowY + selRow.RowHeight
        Dim intLeft, intRight As Integer

        If intBottom > ContentBottom Then intBottom = ContentBottom

        PageGraph.DrawLine(penBlack1, Me.LeftMargin, intTop, Me.LeftMargin, intBottom + 1)

        If selRow.Struct IsNot Nothing Then
            'If StructRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(StructRectangle.X)
            intRight = LastColumnBorder

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            'End If
        Else
            If AssumptionRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(AssumptionRectangle.X)
                intRight = intLeft + AssumptionRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If DependencyTypeRectangle.Left >= CurrentSectionRectangle.Left And DependencyTypeRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(DependencyTypeRectangle.X)
                intRight = intLeft + DependencyTypeRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If InputTypeRectangle.Left >= CurrentSectionRectangle.Left And InputTypeRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(InputTypeRectangle.X)
                intRight = intLeft + InputTypeRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If ImportanceLevelRectangle.Left >= CurrentSectionRectangle.Left And ImportanceLevelRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(ImportanceLevelRectangle.X)
                intRight = intLeft + ImportanceLevelRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If ImpactRectangle.Left >= CurrentSectionRectangle.Left And ImpactRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(ImpactRectangle.X)
                intRight = intLeft + ImpactRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If ShowDeliverables = True Then
                If SupplierRectangle.Left >= CurrentSectionRectangle.Left And SupplierRectangle.Right <= CurrentSectionRectangle.Right Then
                    intLeft = GetCoordinateX(SupplierRectangle.X)
                    intRight = intLeft + SupplierRectangle.Width

                    PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
                End If

                If DeliverableTypeRectangle.Left >= CurrentSectionRectangle.Left And DeliverableTypeRectangle.Right <= CurrentSectionRectangle.Right Then
                    intLeft = GetCoordinateX(DeliverableTypeRectangle.X)
                    intRight = intLeft + DeliverableTypeRectangle.Width

                    PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
                End If

                If DateExpectedRectangle.Left >= CurrentSectionRectangle.Left And DateExpectedRectangle.Right <= CurrentSectionRectangle.Right Then
                    intLeft = GetCoordinateX(DateExpectedRectangle.X)
                    intRight = intLeft + DateExpectedRectangle.Width

                    PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
                End If

                If DateDeliveredRectangle.Left >= CurrentSectionRectangle.Left And DateDeliveredRectangle.Right <= CurrentSectionRectangle.Right Then
                    intLeft = GetCoordinateX(DateDeliveredRectangle.X)
                    intRight = intLeft + DateDeliveredRectangle.Width

                    PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
                End If

                If DeliverablesRectangle.Left >= CurrentSectionRectangle.Left And DeliverablesRectangle.Right <= CurrentSectionRectangle.Right Then
                    intLeft = GetCoordinateX(DeliverablesRectangle.X)
                    intRight = intLeft + DeliverablesRectangle.Width

                    PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
                End If
            End If

            If ResponseStrategyRectangle.Left >= CurrentSectionRectangle.Left And ResponseStrategyRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(ResponseStrategyRectangle.X)
                intRight = intLeft + ResponseStrategyRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If OwnerRectangle.Left >= CurrentSectionRectangle.Left And OwnerRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(OwnerRectangle.X)
                intRight = intLeft + OwnerRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If
        End If
    End Sub
#End Region

#Region "Clipped text"
    Private Function PrintPage_GetMinHeight(ByVal selRow As PrintDependencyRow)
        Dim intMinHeight As Integer

        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(StructRectangle.Width, selRow.Struct.RTF)
            ElseIf selRow.Assumption IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(AssumptionRectangle.Width, selRow.Assumption.RTF)
            ElseIf String.IsNullOrEmpty(selRow.DependencyTypeText) = False Then
                intMinHeight = .GetFirstLineSpacing(DependencyTypeRectangle.Width, selRow.DependencyTypeText)
            ElseIf String.IsNullOrEmpty(selRow.InputTypeText) = False Then
                intMinHeight = .GetFirstLineSpacing(InputTypeRectangle.Width, selRow.InputTypeText)
            ElseIf String.IsNullOrEmpty(selRow.ImportanceLevelText) = False Then
                intMinHeight = .GetFirstLineSpacing(ImportanceLevelRectangle.Width, selRow.ImportanceLevelText)
            ElseIf String.IsNullOrEmpty(selRow.Impact) = False Then
                intMinHeight = .GetFirstLineSpacing(ImpactRectangle.Width, selRow.Impact)
            ElseIf String.IsNullOrEmpty(selRow.Supplier) = False Then
                intMinHeight = .GetFirstLineSpacing(SupplierRectangle.Width, selRow.Supplier)
            ElseIf String.IsNullOrEmpty(selRow.DeliverableTypeText) = False Then
                intMinHeight = .GetFirstLineSpacing(DeliverableTypeRectangle.Width, selRow.DeliverableTypeText)
            ElseIf String.IsNullOrEmpty(selRow.Deliverables) = False Then
                intMinHeight = .GetFirstLineSpacing(DeliverablesRectangle.Width, selRow.Deliverables)
            ElseIf String.IsNullOrEmpty(selRow.ResponseStrategy) = False Then
                intMinHeight = .GetFirstLineSpacing(ResponseStrategyRectangle.Width, selRow.ResponseStrategy)
            ElseIf String.IsNullOrEmpty(selRow.Owner) = False Then
                intMinHeight = .GetFirstLineSpacing(OwnerRectangle.Width, selRow.Owner)
            Else
                intMinHeight = NewCellHeight()
            End If
        End With

        Return intMinHeight
    End Function

    Private Sub PrintClippedRichText(ByVal strRtf As String, ByVal rImage As Rectangle, ByVal intColumnWidth As Integer)
        With RichTextManager
            Dim bmClip As Bitmap = .PrintClippedRichText(intColumnWidth, strRtf, ContentBottom - rImage.Y)
            strClippedTextTop = .ClippedTextTop
            strClippedTextBottom = .ClippedTextBottom

            rImage.X += CONST_Padding
            rImage.Y += CONST_Padding
            rImage.Height = bmClip.Height
            PageGraph.DrawImage(bmClip, rImage)
        End With
    End Sub

    Private Sub PrintClippedText(ByVal strText As String, ByVal rImage As Rectangle, ByVal intColumnWidth As Integer)
        With RichTextManager
            Dim bmClip As Bitmap = .PrintClippedText(intColumnWidth, strText, ContentBottom - rImage.Y)
            strClippedTextTop = .ClippedTextTop
            strClippedTextBottom = .ClippedTextBottom

            rImage.X += CONST_Padding
            rImage.Y += CONST_Padding
            rImage.Height = bmClip.Height
            PageGraph.DrawImage(bmClip, rImage)
        End With
    End Sub
#End Region

#Region "Print struct"
    Private Sub PrintPage_PrintStructSortNumber(ByVal selRow As PrintDependencyRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructSortRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.StructSortNumber, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintStruct(ByVal selRow As PrintDependencyRow)
        If StructRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(StructRectangle.X)

            If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Struct.CellImage.Width, selRow.Struct.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Struct.CellImage, rImage)
                    'Else
                    '    PrintClippedRichText(selRow.Struct.RTF, rImage, StructRectangle.Width)
                    '    selRow.Struct.RTF = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    Select Case selRow.Struct.GetType
                    '        Case GetType(Goal)
                    '            ClippedRow.Struct = New Goal(strClippedTextBottom)
                    '        Case GetType(Purpose)
                    '            ClippedRow.Struct = New Purpose(strClippedTextBottom)
                    '        Case GetType(Output)
                    '            ClippedRow.Struct = New Output(strClippedTextBottom)
                    '        Case GetType(Activity)
                    '            ClippedRow.Struct = New Activity(strClippedTextBottom)
                    '    End Select
                End If
            End If
        End If
        If selRow.Struct IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Struct.RTF) = False Then
            PageGraph.DrawLine(penBlack1, LeftMargin, LastRowY, LastColumnBorder, LastRowY)
        End If
    End Sub
#End Region

#Region "Print assumptions"
    Private Sub PrintPage_PrintAssumptionSortNumber(ByVal selRow As PrintDependencyRow)
        If AssumptionSortRectangle.Left >= CurrentSectionRectangle.Left And AssumptionSortRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(AssumptionSortRectangle.X)
            Dim rSortNumber As New Rectangle(intX, LastRowY, AssumptionSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.AssumptionSortNumber, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintAssumption(ByVal selRow As PrintDependencyRow)
        If AssumptionRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(AssumptionRectangle.X)

            If selRow.Assumption IsNot Nothing AndAlso selRow.Assumption.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Assumption.CellImage.Width, selRow.Assumption.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Assumption.CellImage, rImage)
                    'Else
                    '    PrintClippedRichText(selRow.Assumption.RTF, rImage, AssumptionRectangle.Width)
                    '    selRow.Assumption.RTF = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.Assumption = New Assumption(strClippedTextBottom)
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

#Region "Dependency detail properties"
    Private Sub PrintPage_PrintDependencyType(ByVal selRow As PrintDependencyRow)
        If DependencyTypeRectangle.Left >= CurrentSectionRectangle.Left And DependencyTypeRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(DependencyTypeRectangle.X)

            If String.IsNullOrEmpty(selRow.DependencyTypeText) = False AndAlso selRow.DependencyTypeImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.DependencyTypeImage.Width, selRow.DependencyTypeImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.DependencyTypeImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.DependencyTypeText, rImage, DependencyTypeRectangle.Width)
                    '    selRow.DependencyTypeText = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.DependencyTypeText = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintInputType(ByVal selRow As PrintDependencyRow)
        If InputTypeRectangle.Left >= CurrentSectionRectangle.Left And InputTypeRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(InputTypeRectangle.X)

            If String.IsNullOrEmpty(selRow.InputTypeText) = False AndAlso selRow.InputTypeImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.InputTypeImage.Width, selRow.InputTypeImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.InputTypeImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.InputTypeText, rImage, InputTypeRectangle.Width)
                    '    selRow.InputTypeText = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.InputTypeText = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintImportanceLevel(ByVal selRow As PrintDependencyRow)
        If ImportanceLevelRectangle.Left >= CurrentSectionRectangle.Left And ImportanceLevelRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ImportanceLevelRectangle.X)

            If String.IsNullOrEmpty(selRow.ImportanceLevelText) = False AndAlso selRow.ImportanceLevelImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.ImportanceLevelImage.Width, selRow.ImportanceLevelImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ImportanceLevelImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.ImportanceLevelText, rImage, ImportanceLevelRectangle.Width)
                    '    selRow.ImportanceLevelText = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.ImportanceLevelText = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintImpact(ByVal selRow As PrintDependencyRow)
        If ImpactRectangle.Left >= CurrentSectionRectangle.Left And ImpactRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ImpactRectangle.X)

            If String.IsNullOrEmpty(selRow.Impact) = False AndAlso selRow.ImpactImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.ImpactImage.Width, selRow.ImpactImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ImpactImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.Impact, rImage, ImpactRectangle.Width)
                    '    selRow.Impact = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.Impact = strClippedTextBottom
                End If
            End If
        End If
    End Sub
#End Region

#Region "Deliverables"
    Private Sub PrintPage_PrintSupplier(ByVal selRow As PrintDependencyRow)
        If SupplierRectangle.Left >= CurrentSectionRectangle.Left And SupplierRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(SupplierRectangle.X)

            If String.IsNullOrEmpty(selRow.Supplier) = False AndAlso selRow.SupplierImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.SupplierImage.Width, selRow.SupplierImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.SupplierImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.Supplier, rImage, SupplierRectangle.Width)
                    '    selRow.Supplier = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.Supplier = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintDeliverableType(ByVal selRow As PrintDependencyRow)
        If DeliverableTypeRectangle.Left >= CurrentSectionRectangle.Left And DeliverableTypeRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(DeliverableTypeRectangle.X)

            If String.IsNullOrEmpty(selRow.DeliverableTypeText) = False AndAlso selRow.DeliverableTypeImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.DeliverableTypeImage.Width, selRow.DeliverableTypeImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.DeliverableTypeImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.DeliverableTypeText, rImage, DeliverableTypeRectangle.Width)
                    '    selRow.DeliverableType = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.DeliverableTypeText = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintDateExpected(ByVal selRow As PrintDependencyRow)
        If DateExpectedRectangle.Left >= CurrentSectionRectangle.Left And DateExpectedRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(DateExpectedRectangle.X)
            Dim rDateExpected As New Rectangle(intX, LastRowY, DateExpectedRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.DateExpected, rDateExpected)
        End If
    End Sub

    Private Sub PrintPage_PrintDateDelivered(ByVal selRow As PrintDependencyRow)
        If DateDeliveredRectangle.Left >= CurrentSectionRectangle.Left And DateDeliveredRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(DateDeliveredRectangle.X)
            Dim rDateDelivered As New Rectangle(intX, LastRowY, DateDeliveredRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.DateDelivered, rDateDelivered)
        End If
    End Sub

    Private Sub PrintPage_PrintDeliverables(ByVal selRow As PrintDependencyRow)
        If DeliverablesRectangle.Left >= CurrentSectionRectangle.Left And DeliverablesRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(DeliverablesRectangle.X)

            If String.IsNullOrEmpty(selRow.Deliverables) = False AndAlso selRow.DeliverablesImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.DeliverablesImage.Width, selRow.DeliverablesImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.DeliverablesImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.Deliverables, rImage, DeliverablesRectangle.Width)
                    '    selRow.Deliverables = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.Deliverables = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintResponseStrategy(ByVal selRow As PrintDependencyRow)
        If ResponseStrategyRectangle.Left >= CurrentSectionRectangle.Left And ResponseStrategyRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ResponseStrategyRectangle.X)

            If String.IsNullOrEmpty(selRow.ResponseStrategy) = False AndAlso selRow.ResponseStrategyImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.ResponseStrategyImage.Width, selRow.ResponseStrategyImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ResponseStrategyImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.ResponseStrategy, rImage, ResponseStrategyRectangle.Width)
                    '    selRow.ResponseStrategy = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.ResponseStrategy = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintOwner(ByVal selRow As PrintDependencyRow)
        If OwnerRectangle.Left >= CurrentSectionRectangle.Left And OwnerRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(OwnerRectangle.X)

            If String.IsNullOrEmpty(selRow.Owner) = False AndAlso selRow.OwnerImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.OwnerImage.Width, selRow.OwnerImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.OwnerImage, rImage)
                    'Else
                    '    PrintClippedText(selRow.Owner, rImage, OwnerRectangle.Width)
                    '    selRow.Owner = strClippedTextTop

                    '    If ClippedRow Is Nothing Then
                    '        ClippedRow = New PrintDependencyRow()
                    '        ClippedRow.ClippedRow = True
                    '    End If
                    '    ClippedRow.Owner = strClippedTextBottom
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

                PrintPage_PrintText(LANG_Dependencies, rAssumption, True, False)

                LastColumnBorder = rAssumption.Right
            End If

            If DependencyTypeRectangle.Left >= CurrentSectionRectangle.Left And DependencyTypeRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(DependencyTypeRectangle.X)
                Dim rReason As New Rectangle(intX, LastRowY, DependencyTypeRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_DependencyType, rReason, True, False)

                LastColumnBorder = rReason.Right
            End If

            If InputTypeRectangle.Left >= CurrentSectionRectangle.Left And InputTypeRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(InputTypeRectangle.X)
                Dim rHowToValidate As New Rectangle(intX, LastRowY, InputTypeRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_InputType, rHowToValidate, True, False)

                LastColumnBorder = rHowToValidate.Right
            End If

            If ImportanceLevelRectangle.Left >= CurrentSectionRectangle.Left And ImportanceLevelRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ImportanceLevelRectangle.X)
                Dim rValidated As New Rectangle(intX, LastRowY, ImportanceLevelRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_ImportanceLevel, rValidated, True, False)

                LastColumnBorder = rValidated.Right
            End If

            If ImpactRectangle.Left >= CurrentSectionRectangle.Left And ImpactRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ImpactRectangle.X)
                Dim rImpact As New Rectangle(intX, LastRowY, ImpactRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Impact, rImpact, True, False)

                LastColumnBorder = rImpact.Right
            End If

            If ShowDeliverables = True Then
                If SupplierRectangle.Left >= CurrentSectionRectangle.Left And SupplierRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(SupplierRectangle.X)
                    Dim rSupplier As New Rectangle(intX, LastRowY, SupplierRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_Supplier, rSupplier, True, False)

                    LastColumnBorder = rSupplier.Right
                End If

                If DeliverableTypeRectangle.Left >= CurrentSectionRectangle.Left And DeliverableTypeRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(DeliverableTypeRectangle.X)
                    Dim rHowToValidate As New Rectangle(intX, LastRowY, DeliverableTypeRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_DeliverableType, rHowToValidate, True, False)

                    LastColumnBorder = rHowToValidate.Right
                End If

                If DateExpectedRectangle.Left >= CurrentSectionRectangle.Left And DateExpectedRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(DateExpectedRectangle.X)
                    Dim rDateExpected As New Rectangle(intX, LastRowY, DateExpectedRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_DateExpected, rDateExpected, True, False)

                    LastColumnBorder = rDateExpected.Right
                End If

                If DateDeliveredRectangle.Left >= CurrentSectionRectangle.Left And DateDeliveredRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(DateDeliveredRectangle.X)
                    Dim rDateDelivered As New Rectangle(intX, LastRowY, DateDeliveredRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_DateDelivered, rDateDelivered, True, False)

                    LastColumnBorder = rDateDelivered.Right
                End If

                If DeliverablesRectangle.Left >= CurrentSectionRectangle.Left And DeliverablesRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(DeliverablesRectangle.X)
                    Dim rDeliverables As New Rectangle(intX, LastRowY, DeliverablesRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_Deliverables, rDeliverables, True, False)

                    LastColumnBorder = rDeliverables.Right
                End If
            End If

            If ResponseStrategyRectangle.Left >= CurrentSectionRectangle.Left And ResponseStrategyRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ResponseStrategyRectangle.X)
                Dim rResponseStrategy As New Rectangle(intX, LastRowY, ResponseStrategyRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_ResponseStrategy, rResponseStrategy, True, False)

                LastColumnBorder = rResponseStrategy.Right
            End If

            If OwnerRectangle.Left >= CurrentSectionRectangle.Left And OwnerRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(OwnerRectangle.X)
                Dim rOwner As New Rectangle(intX, LastRowY, OwnerRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Owner, rOwner, True, False)

                LastColumnBorder = rOwner.Right
            End If
        End If

        LastRowY += ColumnHeadersHeight
    End Sub
#End Region
End Class
