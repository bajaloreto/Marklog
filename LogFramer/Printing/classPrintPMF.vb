Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintPMF
    Inherits ReportBaseClass

    Private Const CONST_MinTextColumnWidth As Integer = 100
    Private Const CONST_TargetColumnWidth As Integer = 80

    Private objLogFrame As New LogFrame
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objPrintList As New PrintPmfRows
    Private objClippedRow As PrintPmfRow = Nothing
    Private intSection As Integer
    Private intPagesWide As Integer
    Private strClippedTextTop, strClippedTextBottom As String
    Private bmStructOverFlow As Bitmap, bmIndOverFlow As Bitmap
    Private boolShowTargetRowTitles As Boolean

    Private boolColumnsWidthSet As Boolean
    Private CurrentSectionRectangle As New Rectangle
    Private rStructSortRectangle, rIndicatorSortRectangle, rVerificationSourceSortRectangle As New PrintRectangle
    Private rStructRectangle, rIndicatorRectangle, rVerificationSourceRectangle As New PrintRectangle
    Private rTargetRowTitleRectangle As New PrintRectangle
    Private rBaselineRectangle As New PrintRectangle
    Private rTargetRectangles As New PrintRectangles
    Private rCollectionMethodRectangle, rFrequencyRectangle, rResponsibilityRectangle As New PrintRectangle

    Private intHorPages As Integer, intCurrentHorPage As Integer = 1
    Private intSectionBorder As Integer
    Private intLastColumnBorder As Integer
    Private intHorizontalPageIndex As Integer
    Private intStartIndex As Integer
    Private ColumnHeadersHeight, TargetRowTitleHeight As Integer

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

    Public Sub New(ByVal logframe As LogFrame, ByVal printsection As Integer, ByVal pageswide As Integer, ByVal showtargetrowtitles As Boolean)
        Me.LogFrame = logframe
        Me.ReportSetup = logframe.ReportSetupPmf
        Me.PrintSection = printsection
        Me.PagesWide = pageswide
        Me.ShowTargetRowTitles = showtargetrowtitles

        Dim intSection As Integer = printsection
        If intSection = PrintSections.NotSelected Or intSection = PrintSections.PrintAll Then intSection = PrintSections.PrintOutputs
        intSection += 1
        Me.TargetDeadlinesSection = logframe.GetTargetDeadlinesSection(intSection)
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

    Public Property TargetDeadlinesSection As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesSection
        End Get
        Set(ByVal value As TargetDeadlinesSection)
            objTargetDeadlinesSection = value
        End Set
    End Property

    Public Property PrintList() As PrintPmfRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintPmfRows)
            objPrintList = value
        End Set
    End Property

    Public Property PrintSection() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
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

    Public Property ShowTargetRowTitles As Boolean
        Get
            Return boolShowTargetRowTitles
        End Get
        Set(ByVal value As Boolean)
            boolShowTargetRowTitles = value
        End Set
    End Property

    Public Property ClippedRow As PrintPmfRow
        Get
            Return objClippedRow
        End Get
        Set(ByVal value As PrintPmfRow)
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

    Private Property IndicatorSortRectangle As PrintRectangle
        Get
            Return rIndicatorSortRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rIndicatorSortRectangle = value
        End Set
    End Property

    Private Property IndicatorRectangle As PrintRectangle
        Get
            Return rIndicatorRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rIndicatorRectangle = value
        End Set
    End Property

    Private Property TargetRowTitleRectangle As PrintRectangle
        Get
            Return rTargetRowTitleRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rTargetRowTitleRectangle = value
        End Set
    End Property

    Private Property BaselineRectangle As PrintRectangle
        Get
            Return rBaselineRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rBaselineRectangle = value
        End Set
    End Property

    Public Property TargetRectangles As PrintRectangles
        Get
            Return rTargetRectangles
        End Get
        Set(ByVal value As PrintRectangles)
            rTargetRectangles = value
        End Set
    End Property

    Private Property VerificationSourceSortRectangle As PrintRectangle
        Get
            Return rVerificationSourceSortRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rVerificationSourceSortRectangle = value
        End Set
    End Property

    Private Property VerificationSourceRectangle As PrintRectangle
        Get
            Return rVerificationSourceRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rVerificationSourceRectangle = value
        End Set
    End Property

    Private Property CollectionMethodRectangle As PrintRectangle
        Get
            Return rCollectionMethodRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rCollectionMethodRectangle = value
        End Set
    End Property

    Private Property FrequencyRectangle As PrintRectangle
        Get
            Return rFrequencyRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rFrequencyRectangle = value
        End Set
    End Property

    Private Property ResponsibilityRectangle As PrintRectangle
        Get
            Return rResponsibilityRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rResponsibilityRectangle = value
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

#Region "Create Performance Measurement Framework table"
    Public Sub CreateList()
        Dim strGoalSort As String = String.Empty, strPurposeSort As String = String.Empty

        PrintList.Clear()

        'add goals
        If Me.PrintSection = PrintSections.PrintGoals Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selGoal As Goal In Me.LogFrame.Goals
                strGoalSort = LogFrame.CreateSortNumber(LogFrame.Goals.IndexOf(selGoal))
                Dim NewRow As New PrintPmfRow(LogFrame.SectionTypes.GoalsSection, strGoalSort, New Goal(selGoal.RTF))
                NewRow.RowCountStruct = GetRowCountOfStruct(selGoal)
                PrintList.Add(NewRow)

                CreateList_Indicators(selGoal.Indicators, strGoalSort, NewRow)
            Next
        End If

        'add purposes
        If Me.PrintSection > PrintSections.PrintGoals Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))

                If Me.PrintSection = PrintSections.PrintPurposes Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintPmfRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
                    NewRow.RowCountStruct = GetRowCountOfStruct(selPurpose)
                    PrintList.Add(NewRow)

                    CreateList_Indicators(selPurpose.Indicators, strPurposeSort, NewRow)
                End If

                'add outputs
                CreateList_Outputs(selPurpose.Outputs, strPurposeSort)
            Next
        End If
    End Sub

    Private Sub CreateList_Outputs(ByVal selOutputs As Outputs, ByVal strParentSort As String)
        Dim strOutputSort As String = String.Empty

        If Me.PrintSection > PrintSections.PrintPurposes Then
            For Each selOutput As Output In selOutputs
                strOutputSort = LogFrame.CreateSortNumber(selOutputs.IndexOf(selOutput), strParentSort)

                If Me.PrintSection = PrintSections.PrintOutputs Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintPmfRow(LogFrame.SectionTypes.OutputsSection, strOutputSort, New Output(selOutput.RTF))
                    NewRow.RowCountStruct = GetRowCountOfStruct(selOutput)
                    PrintList.Add(NewRow)

                    CreateList_Indicators(selOutput.Indicators, strOutputSort, NewRow)
                End If

                CreateList_Activities(selOutput.Activities, strOutputSort)
            Next
        End If
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String)
        Dim strActivitySort As String = String.Empty

        If Me.PrintSection > PrintSections.PrintOutputs Then
            For Each selActivity As Activity In selActivities
                strActivitySort = LogFrame.CreateSortNumber(selActivities.IndexOf(selActivity), strParentSort)

                If Me.PrintSection = PrintSections.PrintActivities Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintPmfRow(LogFrame.SectionTypes.ActivitiesSection, strActivitySort, New Activity(selActivity.RTF))
                    NewRow.RowCountStruct = GetRowCountOfStruct(selActivity)
                    PrintList.Add(NewRow)

                    CreateList_Indicators(selActivity.Indicators, strActivitySort, NewRow)

                    If selActivity.Activities.Count > 0 Then
                        CreateList_Activities(selActivity.Activities, strActivitySort)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub CreateList_Indicators(ByVal selIndicators As Indicators, ByVal strParentSort As String, ByVal objParentRow As PrintPmfRow)
        Dim strIndicatorSort As String
        Dim intIndex As Integer
        Dim objRow As PrintPmfRow

        If selIndicators.Count > 0 Then
            For Each selIndicator As Indicator In selIndicators
                If intIndex = 0 And objParentRow IsNot Nothing Then
                    objRow = objParentRow
                Else
                    objRow = New PrintPmfRow()
                    PrintList.Add(objRow)
                End If
                strIndicatorSort = LogFrame.CreateSortNumber(intIndex, strParentSort)
                CreateList_Indicators_Info(selIndicator, strIndicatorSort, objRow)

                If selIndicator.Indicators.Count > 0 Then
                    CreateList_Indicators(selIndicator.Indicators, strIndicatorSort, Nothing)
                End If
                intIndex += 1
            Next
        End If
    End Sub

    Private Function CreateList_Indicators_Info(ByVal selIndicator As Indicator, ByVal strIndicatorSort As String, _
                                                ByVal objRow As PrintPmfRow) As PrintPmfRow
        With objRow
            objRow.IndicatorSort = strIndicatorSort
            Using objCopier As New ObjectCopy
                objRow.Indicator = objCopier.CopyObject(selIndicator)
            End Using
            objRow.RowCountIndicator = GetRowCountOfIndicator(selIndicator)
        End With
        CreateList_VerificationSources(selIndicator.VerificationSources, strIndicatorSort, objRow)
        Return objRow
    End Function

    Private Sub CreateList_VerificationSources(ByVal selVerificationSources As VerificationSources, ByVal strParentSort As String, ByVal objParentRow As PrintPmfRow)
        Dim strVerificationSourceSort As String
        Dim intIndex As Integer
        Dim objRow As PrintPmfRow

        If selVerificationSources.Count > 0 Then
            For Each selVerificationSource As VerificationSource In selVerificationSources
                If intIndex = 0 And objParentRow IsNot Nothing Then
                    objRow = objParentRow
                Else
                    objRow = New PrintPmfRow()
                    PrintList.Add(objRow)
                End If
                strVerificationSourceSort = LogFrame.CreateSortNumber(intIndex, strParentSort)
                CreateList_VerificationSources_Info(selVerificationSource, strVerificationSourceSort, objRow)

                intIndex += 1
            Next
        End If
    End Sub

    Private Function CreateList_VerificationSources_Info(ByVal selVerificationSource As VerificationSource, ByVal strVerificationSourceSort As String, _
                                                         ByVal objRow As PrintPmfRow) As PrintPmfRow
        With objRow
            .VerificationSourceSort = strVerificationSourceSort
            .VerificationSource = New VerificationSource(selVerificationSource.RTF)
            .CollectionMethod = TextToRichText(selVerificationSource.CollectionMethod)
            .Frequency = selVerificationSource.Frequency
            .Responsibility = selVerificationSource.Responsibility
        End With
        Return objRow
    End Function
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth * PagesWide
        Dim intStructSortWidth, intIndicatorSortWidth, intVerificationSourceSortWidth As Integer
        Dim intTargetRowTitleWidth As Integer
        Dim intTargetColumnWidth, intTargetsWidth As Integer
        Dim intTextColumnWidth, intSortWidth As Integer
        Dim intX As Integer

        SectionBorder = LeftMargin + Me.ContentWidth

        'calculate column widths
        intStructSortWidth = GetStructSortColumnWidth()
        intIndicatorSortWidth = GetIndicatorSortColumnWidth()
        intVerificationSourceSortWidth = GetVerificationSourceSortColumnWidth()
        intTargetRowTitleWidth = GetTargetRowTitleColumnWidth()
        intTargetColumnWidth = GetTargetColumnWidth()

        If intTargetColumnWidth < CONST_TargetColumnWidth Then intTargetColumnWidth = CONST_TargetColumnWidth

        'If intTargetRowTitleWidth > 0 Then ShowTargetRowTitles = True
        If TargetDeadlinesSection IsNot Nothing Then intTargetsWidth = (TargetDeadlinesSection.TargetDeadlines.Count + 1) * intTargetColumnWidth

        intSortWidth = intStructSortWidth + intIndicatorSortWidth + intVerificationSourceSortWidth
        intTextColumnWidth = intAvailableWidth - intSortWidth - intTargetRowTitleWidth - intTargetsWidth
        intTextColumnWidth /= 6

        If intTextColumnWidth < CONST_MinTextColumnWidth Then intTextColumnWidth = CONST_MinTextColumnWidth
        If intTextColumnWidth + intStructSortWidth > ContentWidth Then intTextColumnWidth = ContentWidth - intStructSortWidth

        'set struct column rectangles
        StructSortRectangle = New PrintRectangle(LeftMargin, ContentTop, intStructSortWidth, ContentHeight, intHorizontalPageIndex)
        StructRectangle = New PrintRectangle(StructSortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        With PrintRectangles
            .Add(StructSortRectangle)
            .Add(StructRectangle)
        End With
        If StructRectangle.Left < SectionBorder And StructRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(StructRectangle)
        End If

        'set indicator column rectangles
        IndicatorSortRectangle = New PrintRectangle(StructRectangle.Right, ContentTop, intIndicatorSortWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(IndicatorSortRectangle)

        If IndicatorSortRectangle.Left <= SectionBorder And IndicatorSortRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(IndicatorSortRectangle)
        End If

        IndicatorRectangle = New PrintRectangle(IndicatorSortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(IndicatorRectangle)

        If IndicatorRectangle.Left <= SectionBorder And IndicatorRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(IndicatorRectangle)
        End If

        'set target title row rectangle
        TargetRowTitleRectangle = New PrintRectangle(IndicatorRectangle.Right, ContentTop, intTargetRowTitleWidth, ContentHeight, intHorizontalPageIndex, True)

        If ShowTargetRowTitles = True Then
            PrintRectangles.Add(TargetRowTitleRectangle)
            If TargetRowTitleRectangle.Left <= SectionBorder And TargetRowTitleRectangle.Right > SectionBorder Then
                StretchRectanglesOfPreviousPage(TargetRowTitleRectangle)
            End If
        End If

        'set baseline and target rectangles
        BaselineRectangle = New PrintRectangle(TargetRowTitleRectangle.Right, ContentTop, intTargetColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(BaselineRectangle)
        If BaselineRectangle.Left <= SectionBorder And BaselineRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(BaselineRectangle)
        End If

        intX = BaselineRectangle.Right
        If Me.TargetDeadlinesSection IsNot Nothing Then
            For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
                Dim TargetRectangle As New PrintRectangle(intX, ContentTop, intTargetColumnWidth, ContentHeight, intHorizontalPageIndex)
                If TargetRectangle.Left <= SectionBorder And TargetRectangle.Right > SectionBorder Then
                    StretchRectanglesOfPreviousPage(TargetRectangle)
                End If
                TargetRectangles.Add(TargetRectangle)
                PrintRectangles.Add(TargetRectangle)
                intX = TargetRectangle.Right
            Next
        End If

        'set verification source rectangle
        VerificationSourceSortRectangle = New PrintRectangle(intX, ContentTop, intVerificationSourceSortWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(VerificationSourceSortRectangle)

        If VerificationSourceSortRectangle.Left <= SectionBorder And VerificationSourceSortRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(VerificationSourceSortRectangle)
        End If

        VerificationSourceRectangle = New PrintRectangle(VerificationSourceSortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(VerificationSourceRectangle)

        If VerificationSourceRectangle.Left <= SectionBorder And VerificationSourceRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(VerificationSourceRectangle)
        End If

        'set collection method rectangle
        CollectionMethodRectangle = New PrintRectangle(VerificationSourceRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(CollectionMethodRectangle)
        If CollectionMethodRectangle.Left <= SectionBorder And CollectionMethodRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(CollectionMethodRectangle)
        End If

        'set frequency rectangle
        FrequencyRectangle = New PrintRectangle(CollectionMethodRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(FrequencyRectangle)
        If FrequencyRectangle.Left <= SectionBorder And FrequencyRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(FrequencyRectangle)
        End If

        'set responsibility rectangle
        ResponsibilityRectangle = New PrintRectangle(FrequencyRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(ResponsibilityRectangle)
        If ResponsibilityRectangle.Left <= SectionBorder And ResponsibilityRectangle.Right > SectionBorder Then
            StretchRectanglesOfPreviousPage(ResponsibilityRectangle)
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

        For Each selRow As PrintPmfRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.StructSort) = False AndAlso selRow.StructSort.Length > strSort.Length Then strSort = selRow.StructSort
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + CONST_HorizontalPadding
        End If
        Return intWidth
    End Function

    Private Function GetIndicatorSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintPmfRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.IndicatorSort) = False AndAlso selRow.IndicatorSort.Length > strSort.Length Then strSort = selRow.IndicatorSort
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + CONST_HorizontalPadding
        End If
        Return intWidth
    End Function

    Private Function GetVerificationSourceSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintPmfRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.VerificationSourceSort) = False AndAlso selRow.VerificationSourceSort.Length > strSort.Length Then strSort = selRow.VerificationSourceSort
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + CONST_HorizontalPadding
        End If
        Return intWidth
    End Function

    Private Function GetTargetRowTitleColumnWidth() As Integer
        Dim boolBeneficiaryLevel As Boolean
        Dim intWidth As Integer
        Dim strTitle As String

        If ShowTargetRowTitles = True Then
            For Each selRow As PrintPmfRow In Me.PrintList
                If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then boolBeneficiaryLevel = True
            Next

            If boolBeneficiaryLevel = True Then
                strTitle = String.Format("{0} {1}", LANG_ScoreValueTargetGroup, "MMMMMM")
            Else
                strTitle = LANG_Target
            End If
            intWidth = PageGraph.MeasureString(strTitle, fntTextBold).Width + CONST_HorizontalPadding
        End If

        Return intWidth
    End Function

    Private Function GetTargetColumnWidth() As Integer
        Dim intTargetWidth, intWidth As Integer

        For Each selRow As PrintPmfRow In Me.PrintList
            If selRow.Indicator IsNot Nothing Then
                Dim selIndicator As Indicator = selRow.Indicator
                Dim intTargetIndex As Integer = selIndicator.Targets.Count - 1
                Dim strValue As String

                If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                    If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                        strValue = selRow.Indicator.GetPopulationTargetFormattedValue(intTargetIndex)
                    Else
                        strValue = selRow.Indicator.GetPopulationTargetFormattedScore(intTargetIndex)
                    End If
                Else
                    If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                        strValue = selRow.Indicator.GetTargetFormattedValue(intTargetIndex)
                    Else
                        strValue = selRow.Indicator.GetTargetFormattedScore(intTargetIndex)
                    End If
                End If

                intTargetWidth = PageGraph.MeasureString(strValue, fntTextBold).Width
                If intTargetWidth > intWidth Then intWidth = intTargetWidth
            End If
        Next

        Return intWidth
    End Function
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintPmfRow In Me.PrintList
            ReloadImages_Normal(selRow)
        Next

        ResetRowHeights()
        For Each selRow As PrintPmfRow In Me.PrintList
            ResetRowHeights_MergedIndicators(selRow)
        Next

        For Each selRow As PrintPmfRow In Me.PrintList
            ResetRowHeights_MergedStructs(selRow)
        Next
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintPmfRow)
        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Struct.Text) Then
                    selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.GetItemName(selRow.Section), selRow.StructSort, False)
                Else
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(StructRectangle.Width, selRow.Struct.RTF, False)
                End If
            End If
            If selRow.Indicator IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Indicator.Text) Then
                    selRow.Indicator.CellImage = .EmptyTextWithPaddingToBitmap(IndicatorRectangle.Width, Indicator.ItemName, selRow.IndicatorSort, False)
                Else
                    selRow.Indicator.CellImage = .RichTextWithPaddingToBitmap(IndicatorRectangle.Width, selRow.Indicator.RTF, False, selRow.IndicatorIndent)
                End If
            End If
            If selRow.VerificationSource IsNot Nothing Then
                selRow.VerificationSource.CellImage = .RichTextWithPaddingToBitmap(VerificationSourceRectangle.Width, selRow.VerificationSource.RTF, False)
                selRow.VerificationSourceHeight = selRow.VerificationSource.CellImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintPmfRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > 0 Then selPrintListRow.RowHeight = intRowHeight Else selPrintListRow.RowHeight = NewCellHeight()
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Function GetTargetRowTitleHeight() As Integer
        Dim intRowHeaderHeight As Integer

        intRowHeaderHeight = fntTextBold.Height + CONST_VerticalPadding

        Return intRowHeaderHeight
    End Function

    Private Function GetTargetRowTitlesHeight(ByVal intRegistration As Integer) As Integer
        Dim intRowHeaders As Integer = 1
        Dim intRowHeadersHeight As Integer

        If intRegistration = Indicator.RegistrationOptions.BeneficiaryLevel Then intRowHeaders = 3

        intRowHeadersHeight = GetTargetRowTitleHeight() * intRowHeaders
    End Function

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selRow As PrintPmfRow = Me.PrintList(RowIndex)

        If ShowTargetRowTitles = True Then
            If selRow.Targets IsNot Nothing And selRow.Indicator IsNot Nothing Then
                Dim intRowHeadersHeight As Integer = GetTargetRowTitlesHeight(selRow.Indicator.Registration)

                If intRowHeadersHeight > intRowHeight Then intRowHeight = intRowHeadersHeight
            End If
        End If

        If selRow.VerificationSource IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.VerificationSource.RTF) = False Then
            If selRow.VerificationSourceHeight > intRowHeight Then intRowHeight = selRow.VerificationSourceHeight
        Else
            If selRow.Indicator IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Indicator.RTF) = False Then
                If selRow.IndicatorHeight > intRowHeight Then intRowHeight = selRow.IndicatorHeight
            ElseIf selRow.Struct IsNot Nothing Then
                If selRow.StructHeight > intRowHeight Then intRowHeight = selRow.StructHeight
            End If
        End If

        Return intRowHeight
    End Function

    Private Sub ResetRowHeights_MergedIndicators(ByVal selRow As PrintPmfRow)
        Dim intRowIndex As Integer = Me.PrintList.IndexOf(selRow)

        If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.CellImage IsNot Nothing Then
            Dim selIndicator As Indicator = selRow.Indicator
            Dim intRowCount As Integer = selRow.RowCountIndicator
            If ClippedRow IsNot Nothing Then intRowCount += 1

            If intRowCount <= 1 Then
                selRow.IndicatorHeight = selRow.Indicator.CellImage.Height
                If selRow.RowHeight < selRow.IndicatorHeight Then selRow.RowHeight = selRow.IndicatorHeight
            Else
                Dim intIndicatorHeight As Integer = selRow.Indicator.CellImage.Height


                Dim intTop As Integer = 0
                For i = 0 To intRowCount - 1
                    If intRowIndex + i > Me.PrintList.Count - 1 Then Exit For

                    Dim objPrintListRow As PrintPmfRow = PrintList(intRowIndex + i)
                    Dim intAvailableHeight As Integer = objPrintListRow.RowHeight

                    If i = intRowCount - 1 Then
                        Dim intNeededHeight As Integer = intIndicatorHeight - intTop

                        If ShowTargetRowTitles = True Then
                            Dim intNeededHeightTargets As Integer = GetTargetRowTitlesHeight(selRow.Indicator.Registration)
                            If intNeededHeightTargets > intNeededHeight Then intNeededHeight = intNeededHeightTargets
                        End If

                        If intNeededHeight > intAvailableHeight Then objPrintListRow.RowHeight = intNeededHeight
                    End If

                    If intAvailableHeight > 0 Then
                        objPrintListRow.IndicatorHeight = intAvailableHeight
                        intTop += objPrintListRow.RowHeight
                    Else
                        objPrintListRow.IndicatorHeight = 0
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ResetRowHeights_MergedStructs(ByVal selRow As PrintPmfRow)
        Dim intRowIndex As Integer = Me.PrintList.IndexOf(selRow)

        If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
            Dim selStruct As Struct = selRow.Struct
            Dim intRowCount As Integer = selRow.RowCountStruct
            If ClippedRow IsNot Nothing Then intRowCount += 1

            If intRowCount <= 1 Then
                selRow.StructHeight = selRow.Struct.CellImage.Height
                If selRow.RowHeight < selRow.StructHeight Then selRow.RowHeight = selRow.StructHeight
            Else
                Dim intStructHeight As Integer = selRow.Struct.CellImage.Height
                Dim intTop As Integer = 0
                For i = 0 To intRowCount - 1
                    If intRowIndex + i > Me.PrintList.Count - 1 Then Exit For

                    Dim objPrintListRow As PrintPmfRow = PrintList(intRowIndex + i)
                    Dim intAvailableHeight As Integer = objPrintListRow.RowHeight

                    If i = intRowCount - 1 Then
                        Dim intNeededHeight As Integer = intStructHeight - intTop
                        If intNeededHeight > intAvailableHeight Then objPrintListRow.RowHeight = intNeededHeight
                    End If

                    If intAvailableHeight > 0 Then
                        objPrintListRow.StructHeight = intAvailableHeight
                        intTop += objPrintListRow.RowHeight
                    Else
                        objPrintListRow.StructHeight = 0
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub SetColumnHeadersHeight()
        Dim intHeight As Integer

        If PageGraph IsNot Nothing Then
            Dim fntHeader As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
            Dim intStructHeight As Integer = PageGraph.MeasureString(LogFrame.StructNamePlural(2), fntHeader, StructRectangle.Width).Height
            If intStructHeight > intHeight Then intHeight = intStructHeight

            Dim intIndicatorSortHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, IndicatorSortRectangle.Width).Height
            If intIndicatorSortHeight > intHeight Then intHeight = intIndicatorSortHeight

            Dim intIndicatorHeight As Integer = PageGraph.MeasureString(LANG_Indicator, fntHeader, IndicatorRectangle.Width).Height
            If intIndicatorHeight > intHeight Then intHeight = intIndicatorHeight

            Dim intTargetRowTitleHeight As Integer = PageGraph.MeasureString(LANG_TargetType, fntHeader, CONST_TargetColumnWidth).Height
            If intTargetRowTitleHeight > intHeight Then intHeight = intTargetRowTitleHeight

            Dim intBaselineHeight As Integer = PageGraph.MeasureString(LANG_Baseline, fntHeader, CONST_TargetColumnWidth).Height
            If intBaselineHeight > intHeight Then intHeight = intBaselineHeight

            Dim intTargetHeight As Integer = PageGraph.MeasureString(LANG_Target, fntHeader, CONST_TargetColumnWidth).Height
            If intTargetHeight > intHeight Then intHeight = intTargetHeight

            Dim intVerificationSourceSortHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, VerificationSourceSortRectangle.Width).Height
            If intVerificationSourceSortHeight > intHeight Then intHeight = intVerificationSourceSortHeight

            Dim intVerificationSourceHeight As Integer = PageGraph.MeasureString(LANG_VerificationSource, fntHeader, VerificationSourceRectangle.Width).Height
            If intVerificationSourceHeight > intHeight Then intHeight = intVerificationSourceHeight

            Dim intCollectionMethodHeight As Integer = PageGraph.MeasureString(LANG_DataCollection, fntHeader, CollectionMethodRectangle.Width).Height
            If intCollectionMethodHeight > intHeight Then intHeight = intCollectionMethodHeight

            Dim intFrequencyHeight As Integer = PageGraph.MeasureString(LANG_Number, fntHeader, FrequencyRectangle.Width).Height
            If intFrequencyHeight > intHeight Then intHeight = intFrequencyHeight

            Dim intResponsibilityHeight As Integer = PageGraph.MeasureString(LANG_Responsibility, fntHeader, ResponsibilityRectangle.Width).Height
            If intResponsibilityHeight > intHeight Then intHeight = intResponsibilityHeight

            ColumnHeadersHeight = intHeight
        End If
    End Sub
#End Region

#Region "Get row counts"
    Private Function GetRowCountOfStruct(ByVal objStruct As Struct) As Integer
        Dim NrRows As Integer

        If objStruct IsNot Nothing Then
            Dim intRowCount As Integer = GetRowCountOfStruct_Indicators(objStruct)
            Dim intRowCountAsm As Integer = GetRowCountOfStruct_Assumptions(objStruct)

            If intRowCount >= intRowCountAsm Then NrRows = intRowCount Else NrRows = intRowCountAsm
        End If

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Indicators(ByVal objStruct As Struct) As Integer
        Dim NrRows As Integer

        If objStruct IsNot Nothing Then
            NrRows = GetRowCountOfStruct_Indicators_Count(objStruct.Indicators)
        End If

        If NrRows = 0 Then NrRows = 1

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Indicators_Count(ByVal selIndicators As Indicators) As Integer
        Dim NrRows As Integer
        Dim boolLastHasSubIndicators As Boolean

        For Each selIndicator As Indicator In selIndicators
            boolLastHasSubIndicators = False

            NrRows += GetRowCountOfIndicator(selIndicator)
            If selIndicator.Indicators.Count > 0 Then
                NrRows += GetRowCountOfStruct_Indicators_Count(selIndicator.Indicators)
                boolLastHasSubIndicators = True
            End If
        Next

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Assumptions(ByVal objStruct As Struct) As Integer
        Dim NmbrRows As Integer

        If objStruct IsNot Nothing Then NmbrRows = objStruct.Assumptions.Count

        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function

    Private Function GetRowCountOfIndicator(ByVal selIndicator As Indicator) As Integer
        Dim NmbrRows As Integer

        If selIndicator IsNot Nothing AndAlso selIndicator.VerificationSources.Count > 0 Then
            NmbrRows = selIndicator.VerificationSources.Count
        End If

        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight - ColumnHeadersHeight

        For Each selRow As PrintPmfRow In PrintList
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
        TargetRectangles.Clear()
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
        Dim selRow As PrintPmfRow = PrintList(RowIndex)
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

                PrintPage_PrintStructSortNumber(selRow)
                PrintPage_PrintStruct(selRow)
                PrintPage_PrintIndicatorSortNumber(selRow)
                PrintPage_PrintIndicator(selRow)
                PrintPage_PrintTargetRowTitle(selRow)
                PrintPage_PrintBaseline(selRow)
                PrintPage_PrintTargets(selRow)
                PrintPage_PrintVerificationSourceSortNumber(selRow)
                PrintPage_PrintVerificationSource(selRow)
                PrintPage_PrintCollectionMethod(selRow)
                PrintPage_PrintFrequency(selRow)
                PrintPage_PrintResponsibility(selRow)

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
                If RowIndex > PrintList.Count - 1 Then Exit Do

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

    Private Sub PrintPage_PrintText(ByVal strValue As String, ByVal rCell As Rectangle, Optional ByVal AlignRight As Boolean = False, Optional ByVal boolHeader As Boolean = False)
        If PageGraph IsNot Nothing Then
            If boolHeader = True Then
                PageGraph.FillRectangle(New SolidBrush(Color.LightGray), rCell)
            End If

            Dim formatCells As New StringFormat()
            Dim brText As SolidBrush = New SolidBrush(colForeGround)
            Dim rText As Rectangle = GetTextRectangle(rCell)

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

            If boolHeader = True Then
                PageGraph.DrawRectangle(penBlack1, rCell)
            End If

        End If
    End Sub

    Private Sub PrintPage_PrintTarget(ByVal strValue As String, ByVal rCell As Rectangle, Optional ByVal boolHeader As Boolean = False)
        If PageGraph IsNot Nothing Then
            Dim formatCells As New StringFormat()
            Dim brTarget As SolidBrush = New SolidBrush(colForeGround)
            Dim rTarget As Rectangle = GetTextRectangle(rCell)
            Dim fntTarget As Font = fntText

            formatCells.LineAlignment = StringAlignment.Near
            If boolHeader = True Then
                fntTarget = fntTextBold
                formatCells.Alignment = StringAlignment.Near
            Else
                formatCells.Alignment = StringAlignment.Far
            End If

            PageGraph.DrawString(strValue, fntTarget, brTarget, rTarget, formatCells)
            PageGraph.DrawLine(penBlack1, rCell.Left, rCell.Top, rCell.Right, rCell.Top)
        End If
    End Sub

    Private Sub PrintPage_PrintLines(ByVal selRow As PrintPmfRow)
        Dim intTop As Integer = LastRowY
        Dim intBottom As Integer = LastRowY + selRow.RowHeight
        Dim intLeft, intRight As Integer
        Dim intTargetDeadlineIndex As Integer
        Dim TargetRectangle As PrintRectangle

        If intBottom > ContentBottom Then intBottom = ContentBottom

        'draw vertical lines (columns)
        PageGraph.DrawLine(penBlack1, Me.LeftMargin, intTop, Me.LeftMargin, intBottom + 1)

        If StructRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(StructRectangle.X)
            intRight = intLeft + StructRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If IndicatorRectangle.Left >= CurrentSectionRectangle.Left And IndicatorRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(IndicatorRectangle.X)
            intRight = intLeft + IndicatorRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If ShowTargetRowTitles = True Then
            If TargetRowTitleRectangle.Left >= CurrentSectionRectangle.Left And TargetRowTitleRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(TargetRowTitleRectangle.X)
                intRight = intLeft + TargetRowTitleRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If
        End If

        If BaselineRectangle.Left >= CurrentSectionRectangle.Left And BaselineRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(BaselineRectangle.X)
            intRight = intLeft + BaselineRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If Me.TargetDeadlinesSection IsNot Nothing Then
            For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
                TargetRectangle = TargetRectangles(intTargetDeadlineIndex)

                intLeft = GetCoordinateX(TargetRectangle.X)
                intRight = intLeft + TargetRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)

                intTargetDeadlineIndex += 1
            Next
        End If

        If VerificationSourceSortRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceSortRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(VerificationSourceSortRectangle.X)
            intRight = intLeft + VerificationSourceSortRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If VerificationSourceRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(VerificationSourceRectangle.X)
            intRight = intLeft + VerificationSourceRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If CollectionMethodRectangle.Left >= CurrentSectionRectangle.Left And CollectionMethodRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(CollectionMethodRectangle.X)
            intRight = intLeft + CollectionMethodRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If FrequencyRectangle.Left >= CurrentSectionRectangle.Left And FrequencyRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(FrequencyRectangle.X)
            intRight = intLeft + FrequencyRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If

        If ResponsibilityRectangle.Left >= CurrentSectionRectangle.Left And ResponsibilityRectangle.Right <= CurrentSectionRectangle.Right Then
            intLeft = GetCoordinateX(ResponsibilityRectangle.X)
            intRight = intLeft + ResponsibilityRectangle.Width

            PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
        End If
    End Sub
#End Region

#Region "Clipped text"
    Private Function PrintPage_GetMinHeight(ByVal selRow As PrintPmfRow)
        Dim intMinHeight As Integer

        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(StructRectangle.Width, selRow.Struct.RTF)
            ElseIf selRow.Indicator IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(IndicatorRectangle.Width, selRow.Indicator.RTF)
            ElseIf selRow.VerificationSource IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(VerificationSourceRectangle.Width, selRow.VerificationSource.RTF)
            ElseIf String.IsNullOrEmpty(selRow.CollectionMethod) = False Then
                intMinHeight = .GetFirstLineSpacing(CollectionMethodRectangle.Width, selRow.CollectionMethod)
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

#Region "Print struct"
    Private Sub PrintPage_PrintStructSortNumber(ByVal selRow As PrintPmfRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.StructSort, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintStruct(ByVal selRow As PrintPmfRow)
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
                        ClippedRow = New PrintPmfRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    Select Case selRow.Section
                        Case LogFrame.SectionTypes.GoalsSection
                            ClippedRow.Struct = New Goal(strClippedTextBottom)
                        Case LogFrame.SectionTypes.PurposesSection
                            ClippedRow.Struct = New Purpose(strClippedTextBottom)
                        Case LogFrame.SectionTypes.OutputsSection
                            ClippedRow.Struct = New Output(strClippedTextBottom)
                        Case LogFrame.SectionTypes.ActivitiesSection
                            ClippedRow.Struct = New Activity(strClippedTextBottom)
                            ClippedRow.StructIndent = selRow.StructIndent
                    End Select
                End If
            End If
        End If
        If selRow.Struct IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Struct.RTF) = False Then
            PageGraph.DrawLine(penBlack1, LeftMargin, LastRowY, LastColumnBorder, LastRowY)
        End If
    End Sub
#End Region

#Region "Print indicators and their targets"
    Private Sub PrintPage_PrintIndicatorSortNumber(ByVal selRow As PrintPmfRow)
        If IndicatorSortRectangle.Left >= CurrentSectionRectangle.Left And IndicatorSortRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(IndicatorSortRectangle.X)
            Dim rSortNumber As New Rectangle(intX, LastRowY, IndicatorSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.IndicatorSort, rSortNumber)

            If selRow.Indicator IsNot Nothing AndAlso (String.IsNullOrEmpty(selRow.Indicator.RTF) = False And selRow.ClippedRow = False) Then
                PageGraph.DrawLine(penBlack1, intX, LastRowY, intX + IndicatorSortRectangle.Width, LastRowY)
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintIndicator(ByVal selRow As PrintPmfRow)
        Dim intX As Integer = GetCoordinateX(IndicatorRectangle.X)

        If IndicatorRectangle.Left >= CurrentSectionRectangle.Left And IndicatorRectangle.Right <= CurrentSectionRectangle.Right Then
            If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Indicator.CellImage.Width, selRow.Indicator.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Indicator.CellImage, rImage)
                Else
                    PrintClippedText(selRow.Indicator.RTF, rImage, IndicatorRectangle.Width)
                    selRow.Indicator.RTF = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintPmfRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.Indicator = New Indicator(strClippedTextBottom)
                End If
            End If
        End If

        If selRow.Indicator IsNot Nothing AndAlso (String.IsNullOrEmpty(selRow.Indicator.RTF) = False And selRow.ClippedRow = False) Then
            PageGraph.DrawLine(penBlack1, intX, LastRowY, intX + IndicatorRectangle.Width, LastRowY)
        End If
    End Sub

    Private Sub PrintPage_PrintTargetRowTitle(ByVal selRow As PrintPmfRow)
        If ShowTargetRowTitles = True Then
            If TargetRowTitleRectangle.Left >= CurrentSectionRectangle.Left And TargetRowTitleRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim intX As Integer = GetCoordinateX(TargetRowTitleRectangle.X)
                Dim intY As Integer = LastRowY
                Dim rTargetRowTitle As New Rectangle(intX, intY, TargetRowTitleRectangle.Width, selRow.RowHeight + 1)
                Dim strValue As String

                If rTargetRowTitle.Bottom > ContentBottom Then rTargetRowTitle.Height = ContentBottom - rTargetRowTitle.Top

                PageGraph.FillRectangle(New SolidBrush(Color.LightGray), rTargetRowTitle)

                If selRow.Indicator IsNot Nothing And selRow.Targets IsNot Nothing Then
                    Dim selIndicator As Indicator = selRow.Indicator
                    Dim intTargetRowTitleHeight As Integer = GetTargetRowTitleHeight()

                    If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                        If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                            strValue = LANG_TargetValueBeneficiary
                        Else
                            strValue = LANG_ScoreValueBeneficiary
                        End If
                        rTargetRowTitle = New Rectangle(intX, intY, TargetRowTitleRectangle.Width, intTargetRowTitleHeight)
                    Else
                        strValue = LANG_Target
                        rTargetRowTitle = New Rectangle(intX, intY, TargetRowTitleRectangle.Width, selRow.Indicator.CellImage.Height)
                    End If

                    PrintPage_PrintTarget(strValue, rTargetRowTitle, True)

                    If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                        intY += intTargetRowTitleHeight

                        rTargetRowTitle = New Rectangle(intX, intY, TargetRowTitleRectangle.Width, intTargetRowTitleHeight)
                        PrintPage_PrintTarget(LANG_PopulationTargetText, rTargetRowTitle, True)

                        intY += intTargetRowTitleHeight

                        rTargetRowTitle = New Rectangle(intX, intY, TargetRowTitleRectangle.Width, selRow.Indicator.CellImage.Height - (intTargetRowTitleHeight * 2))
                        PrintPage_PrintTarget(LANG_ScoreValueTargetGroup, rTargetRowTitle, True)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintBaseline(ByVal selRow As PrintPmfRow)
        Dim intTargetRowHeight As Integer = GetTargetRowTitleHeight()
        Dim intX As Integer = GetCoordinateX(BaselineRectangle.X)
        Dim intY As Integer = LastRowY
        Dim rBaseline As New Rectangle(intX, intY, BaselineRectangle.Width, intTargetRowHeight)

        If selRow.Baseline IsNot Nothing Then
            If BaselineRectangle.Left >= CurrentSectionRectangle.Left And BaselineRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim selIndicator As Indicator = selRow.Indicator

                Dim strValue As String

                If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                    strValue = selRow.Indicator.GetBaselineFormattedValue
                Else
                    strValue = selRow.Indicator.GetBaselineFormattedScore
                End If
                PrintPage_PrintTarget(strValue, rBaseline)

                If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                    intY += intTargetRowHeight

                    rBaseline = New Rectangle(intX, intY, BaselineRectangle.Width, intTargetRowHeight)
                    PageGraph.FillRectangle(New SolidBrush(Color.LightGray), rBaseline)
                    PageGraph.DrawLine(penBlack1, rBaseline.Left, rBaseline.Top, rBaseline.Right, rBaseline.Top)

                    intY += intTargetRowHeight
                    rBaseline = New Rectangle(intX, intY, BaselineRectangle.Width, intTargetRowHeight)

                    If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                        strValue = selRow.Indicator.GetPopulationBaselineFormattedValue
                    Else
                        strValue = selRow.Indicator.GetPopulationBaselineFormattedScore
                    End If

                    PrintPage_PrintTarget(strValue, rBaseline)
                End If
            End If
        ElseIf selRow.Indicator IsNot Nothing Then
            PageGraph.DrawLine(penBlack1, rBaseline.Left, rBaseline.Top, rBaseline.Right, rBaseline.Top)
        End If
    End Sub

    Private Sub PrintPage_PrintTargets(ByVal selRow As PrintPmfRow)
        Dim intTargetDeadlineIndex, intTargetIndex As Integer
        Dim TargetRectangle As PrintRectangle
        Dim boolFound As Boolean

        If selRow.Targets IsNot Nothing And Me.TargetDeadlinesSection IsNot Nothing Then
            Dim selIndicator As Indicator = selRow.Indicator
            Dim intTargetRowHeight As Integer = GetTargetRowTitleHeight()

            For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
                TargetRectangle = TargetRectangles(intTargetDeadlineIndex)
                intTargetIndex = 0

                If TargetRectangle.Left >= CurrentSectionRectangle.Left And TargetRectangle.Right <= CurrentSectionRectangle.Right Then
                    Dim intX As Integer = GetCoordinateX(TargetRectangle.X)
                    Dim intY As Integer = LastRowY
                    Dim rTarget As New Rectangle(intX, LastRowY, TargetRectangle.Width, selRow.RowHeight)

                    'find target that corresponds to targetdeadline
                    boolFound = False

                    For Each selTarget As Target In selRow.Targets
                        If selTarget.TargetDeadlineGuid = selTargetDeadline.Guid Then

                            Dim strValue As String

                            boolFound = True

                            If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                                strValue = selRow.Indicator.GetTargetFormattedValue(intTargetIndex)
                            Else
                                strValue = selRow.Indicator.GetTargetFormattedScore(intTargetIndex)
                            End If
                            PrintPage_PrintTarget(strValue, rTarget)

                            If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                                intY += intTargetRowHeight
                                strValue = selRow.Indicator.GetPopulationTargetFormattedValue(intTargetIndex)
                                strValue = DisplayAsUnit(selIndicator.PopulationTargets(intTargetIndex).Percentage, selIndicator.ValuesDetail.NrDecimals, "%")
                                rTarget = New Rectangle(intX, intY, TargetRectangle.Width, intTargetRowHeight)
                                PrintPage_PrintTarget(strValue, rTarget)

                                intY += intTargetRowHeight
                                rTarget = New Rectangle(intX, intY, TargetRectangle.Width, intTargetRowHeight)

                                If selIndicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                                    strValue = selRow.Indicator.GetPopulationTargetFormattedValue(intTargetIndex)
                                Else
                                    strValue = selRow.Indicator.GetPopulationTargetFormattedScore(intTargetIndex)
                                End If

                                PrintPage_PrintTarget(strValue, rTarget)
                            End If
                        End If

                        intTargetIndex += 1
                    Next

                    If boolFound = False Then
                        PageGraph.DrawLine(penBlack1, rTarget.Left, rTarget.Top, rTarget.Right, rTarget.Top)
                    End If
                End If
                intTargetDeadlineIndex += 1
            Next
        End If
    End Sub
#End Region

#Region "Print verification sources"
    Private Sub PrintPage_PrintVerificationSourceSortNumber(ByVal selRow As PrintPmfRow)
        If VerificationSourceSortRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceSortRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(VerificationSourceSortRectangle.X)
            Dim rSortNumber As New Rectangle(intX, LastRowY, VerificationSourceSortRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.VerificationSourceSort, rSortNumber)
            PageGraph.DrawLine(penBlack1, rSortNumber.Left, rSortNumber.Top, rSortNumber.Right, rSortNumber.Top)
        End If
    End Sub


    Private Sub PrintPage_PrintVerificationSource(ByVal selRow As PrintPmfRow)
        If VerificationSourceRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(VerificationSourceRectangle.X)
            Dim rVerificationSource As New Rectangle(intX, LastRowY, VerificationSourceRectangle.Width, selRow.RowHeight)

            If selRow.VerificationSource IsNot Nothing Then
                If selRow.VerificationSource.CellImage IsNot Nothing Then
                    Dim rImage As New Rectangle(intX, LastRowY, selRow.VerificationSource.CellImage.Width, selRow.VerificationSource.CellImage.Height)

                    If rImage.Bottom <= ContentBottom Then
                        PageGraph.DrawImage(selRow.VerificationSource.CellImage, rImage)
                    Else
                        PrintClippedText(selRow.VerificationSource.RTF, rImage, VerificationSourceRectangle.Width)
                        selRow.VerificationSource.RTF = strClippedTextTop

                        If ClippedRow Is Nothing Then
                            ClippedRow = New PrintPmfRow(selRow.Section)
                            ClippedRow.ClippedRow = True
                        End If
                        ClippedRow.VerificationSource = New VerificationSource(strClippedTextBottom)
                    End If
                End If
            End If
            PageGraph.DrawLine(penBlack1, rVerificationSource.Left, rVerificationSource.Top, rVerificationSource.Right, rVerificationSource.Top)
        End If
    End Sub

    Private Sub PrintPage_PrintCollectionMethod(ByVal selRow As PrintPmfRow)
        If CollectionMethodRectangle.Left >= CurrentSectionRectangle.Left And CollectionMethodRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(CollectionMethodRectangle.X)
            Dim rCollectionMethod As New Rectangle(intX, LastRowY, CollectionMethodRectangle.Width, selRow.RowHeight)

            If selRow.CollectionMethod IsNot Nothing AndAlso selRow.CollectionMethodImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.CollectionMethodImage.Width, selRow.CollectionMethodImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.CollectionMethodImage, rImage)
                Else
                    PrintClippedText(selRow.CollectionMethod, rImage, CollectionMethodRectangle.Width)
                    selRow.CollectionMethod = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintPmfRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.CollectionMethod = strClippedTextBottom
                End If
            End If
            PageGraph.DrawLine(penBlack1, rCollectionMethod.Left, rCollectionMethod.Top, rCollectionMethod.Right, rCollectionMethod.Top)
        End If
    End Sub

    Private Sub PrintPage_PrintFrequency(ByVal selRow As PrintPmfRow)
        If FrequencyRectangle.Left >= CurrentSectionRectangle.Left And FrequencyRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(FrequencyRectangle.X)
            Dim rFrequency As New Rectangle(intX, LastRowY, FrequencyRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.Frequency, rFrequency)
            PageGraph.DrawLine(penBlack1, rFrequency.Left, rFrequency.Top, rFrequency.Right, rFrequency.Top)
        End If
    End Sub

    Private Sub PrintPage_PrintResponsibility(ByVal selRow As PrintPmfRow)
        If ResponsibilityRectangle.Left >= CurrentSectionRectangle.Left And ResponsibilityRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ResponsibilityRectangle.X)
            Dim rResponsibility As New Rectangle(intX, LastRowY, ResponsibilityRectangle.Width, selRow.RowHeight)

            PrintPage_PrintText(selRow.Responsibility, rResponsibility)
            PageGraph.DrawLine(penBlack1, rResponsibility.Left, rResponsibility.Top, rResponsibility.Right, rResponsibility.Top)
        End If
    End Sub
#End Region

#Region "Column headers"
    Private Sub PrintColumnHeaders()
        Dim intX As Integer
        Dim strHeaderText, strDate As String

        If PageGraph IsNot Nothing Then
            If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim rStruct As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width + StructRectangle.Width, ColumnHeadersHeight)
                Dim strObjectives As String = GetColumnTitleObjectives()

                PrintPage_PrintText(strObjectives, rStruct, False, True)

                LastColumnBorder = rStruct.Right
            End If

            If IndicatorSortRectangle.Left >= CurrentSectionRectangle.Left And IndicatorRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(IndicatorSortRectangle.X)
                Dim rIndicator As New Rectangle(intX, LastRowY, IndicatorSortRectangle.Width + IndicatorRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Indicators, rIndicator, False, True)

                LastColumnBorder = rIndicator.Right
            ElseIf IndicatorSortRectangle.Left >= CurrentSectionRectangle.Left And IndicatorSortRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(IndicatorSortRectangle.X)
                Dim rIndicatorSort As New Rectangle(intX, LastRowY, IndicatorSortRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Number, rIndicatorSort, False, True)

                LastColumnBorder = rIndicatorSort.Right
            ElseIf IndicatorRectangle.Left >= CurrentSectionRectangle.Left And IndicatorRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(IndicatorRectangle.X)
                Dim rIndicator As New Rectangle(intX, LastRowY, IndicatorRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Indicators, rIndicator, False, True)

                LastColumnBorder = rIndicator.Right
            End If

            If ShowTargetRowTitles = True Then
                If TargetRowTitleRectangle.Left >= CurrentSectionRectangle.Left And TargetRowTitleRectangle.Right <= CurrentSectionRectangle.Right Then
                    intX = GetCoordinateX(TargetRowTitleRectangle.X)
                    Dim rTargetRowTitle As New Rectangle(intX, LastRowY, TargetRowTitleRectangle.Width, ColumnHeadersHeight)

                    PrintPage_PrintText(LANG_TargetType, rTargetRowTitle, False, True)

                    LastColumnBorder = rTargetRowTitle.Right
                End If
            End If

            If BaselineRectangle.Left >= CurrentSectionRectangle.Left And BaselineRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(BaselineRectangle.X)
                Dim rBaseline As New Rectangle(intX, LastRowY, BaselineRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Baseline, rBaseline, False, True)

                LastColumnBorder = rBaseline.Right
            End If

            If Me.TargetDeadlinesSection IsNot Nothing Then
                Dim intTargetIndex As Integer
                For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
                    Dim selTargetRectangle As PrintRectangle = TargetRectangles(intTargetIndex)
                    If selTargetRectangle.Left >= CurrentSectionRectangle.Left And selTargetRectangle.Right <= CurrentSectionRectangle.Right Then
                        intX = GetCoordinateX(selTargetRectangle.X)
                        Dim rTarget As New Rectangle(intX, LastRowY, selTargetRectangle.Width, ColumnHeadersHeight)

                        Select Case TargetDeadlinesSection.Repetition
                            Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                                strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
                            Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                                strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
                            Case Else
                                strDate = selTargetDeadline.ExactDeadline.ToShortDateString
                        End Select

                        strHeaderText = String.Format("{0} {1}", LANG_Target, strDate)

                        PrintPage_PrintText(strHeaderText, rTarget, False, True)

                        LastColumnBorder = rTarget.Right
                    End If

                    intTargetIndex += 1
                Next
            End If

            If VerificationSourceSortRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(VerificationSourceSortRectangle.X)
                Dim rVerificationSource As New Rectangle(intX, LastRowY, VerificationSourceSortRectangle.Width + VerificationSourceRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_VerificationSources, rVerificationSource, False, True)

                LastColumnBorder = rVerificationSource.Right
            ElseIf VerificationSourceSortRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceSortRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(VerificationSourceSortRectangle.X)
                Dim rVerificationSourceSort As New Rectangle(intX, LastRowY, VerificationSourceSortRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Number, rVerificationSourceSort, False, True)

                LastColumnBorder = rVerificationSourceSort.Right
            ElseIf VerificationSourceRectangle.Left >= CurrentSectionRectangle.Left And VerificationSourceRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(VerificationSourceRectangle.X)
                Dim rVerificationSource As New Rectangle(intX, LastRowY, VerificationSourceRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_VerificationSources, rVerificationSource, False, True)

                LastColumnBorder = rVerificationSource.Right
            End If

            If CollectionMethodRectangle.Left >= CurrentSectionRectangle.Left And CollectionMethodRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(CollectionMethodRectangle.X)
                Dim rCollectionMethod As New Rectangle(intX, LastRowY, CollectionMethodRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_DataCollection, rCollectionMethod, False, True)

                LastColumnBorder = rCollectionMethod.Right
            End If

            If FrequencyRectangle.Left >= CurrentSectionRectangle.Left And FrequencyRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(FrequencyRectangle.X)
                Dim rFrequency As New Rectangle(intX, LastRowY, FrequencyRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Frequency, rFrequency, False, True)

                LastColumnBorder = rFrequency.Right
            End If

            If ResponsibilityRectangle.Left >= CurrentSectionRectangle.Left And ResponsibilityRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ResponsibilityRectangle.X)
                Dim rResponsibility As New Rectangle(intX, LastRowY, ResponsibilityRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Responsibility, rResponsibility, False, True)

                LastColumnBorder = rResponsibility.Right
            End If
        End If

        LastRowY += ColumnHeadersHeight
    End Sub

    Public Function GetColumnTitleObjectives() As String
        Dim strObjectives As String

        Select Case Me.PrintSection
            Case PrintSections.PrintGoals
                strObjectives = LogFrame.StructNamePlural(0)
            Case PrintSections.PrintPurposes
                strObjectives = LogFrame.StructNamePlural(1)
            Case PrintSections.PrintOutputs
                strObjectives = LogFrame.StructNamePlural(2)
            Case Else
                strObjectives = LANG_Objectives
        End Select

        Return strObjectives
    End Function
#End Region
End Class


