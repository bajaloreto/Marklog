Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintPlanning
    Inherits ReportBaseClass

    Private objLogFrame As New LogFrame
    Private objGuid As Guid
    Private objPrintList As New PrintPlanningRows
    Private boolColumnsWidthSet As Boolean
    Private intPeriodView As Integer, intElementsView As Integer
    Private boolShowDatesColumns As Boolean
    Private boolShowActivityLinks As Boolean = True, boolShowKeyMomentLinks As Boolean = True
    Private datPeriodFrom, datPeriodUntil As Date
    Private intWeekCorrection As Integer
    Private WidthOfDay As Single

    Private rPlanningCell, rTotalPlanning, rCurrentSection As Rectangle
    Private rProjectDuration As Rectangle
    Private rActivity, rKeyMoment, rPreparation, rFollowUp As Rectangle
    Private rSortColumnRectangle, rTextColumnRectangle As Rectangle
    Private rColumnStartDate, rColumnEndDate As Rectangle
    Private rPlanningColumnRectangle, rFirstPlanningColumnRectangle As Rectangle
    Private rHeadingFirstRow, rHeadingSecondRow As Rectangle
    Private intStartIndex As Integer
    Private ColumnHeadersHeight As Integer
    Private DrawHorArrow, DrawVerArrow As Boolean

    Private intHorPages As Integer, intCurrentHorPage As Integer = 1
    Private ActivityColor As Color

    Public Event PagePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)

    Public Enum PeriodViews
        Day = 0
        Week = 1
        Month = 2
        Trimester = 3
        Semester = 4
        Year = 5
    End Enum

    Public Enum ShowElements
        ShowBoth = 0
        ShowActivities = 1
        ShowKeyMoments = 2
    End Enum

    Public Sub New(ByVal logframe As LogFrame, ByVal outputguid As Guid, ByVal periodview As Integer, ByVal planningelements As Integer, ByVal periodfrom As Date, ByVal perioduntil As Date, _
                   ByVal showactivitylinks As Boolean, ByVal showkeymomentlinks As Boolean, ByVal repeatrowheaders As Boolean, ByVal showdatescolumns As Boolean)
        Me.OutputGuid = outputguid

        Me.LogFrame = logframe
        Me.PeriodView = periodview
        Me.ElementsView = planningelements
        Me.ShowActivityLinks = showactivitylinks
        Me.ShowKeyMomentLinks = showkeymomentlinks
        Me.RepeatRowHeaders = repeatrowheaders
        Me.ShowDatesColumns = showdatescolumns

        Initialise_Period(periodfrom, perioduntil)

        Me.ReportSetup = logframe.ReportSetupPlanning
    End Sub

    Private Sub Initialise_Period(ByVal periodfrom As Date, ByVal perioduntil As Date)
        If periodfrom > Date.MinValue And perioduntil > Date.MinValue Then
            Select Case intPeriodView
                Case DataGridViewPlanning.PeriodViews.Day
                    Me.PeriodFrom = periodfrom.AddDays(-14)
                    Me.PeriodUntil = perioduntil.AddDays(14)
                Case DataGridViewPlanning.PeriodViews.Week
                    Me.PeriodFrom = periodfrom.AddMonths(-1)
                    Me.PeriodUntil = perioduntil.AddMonths(1)
                Case DataGridViewPlanning.PeriodViews.Month
                    Me.PeriodFrom = periodfrom.AddMonths(-3)
                    Me.PeriodUntil = perioduntil.AddMonths(3)
                Case DataGridViewPlanning.PeriodViews.Trimester
                    Me.PeriodFrom = periodfrom.AddMonths(-6)
                    Me.PeriodUntil = perioduntil.AddMonths(6)
                Case DataGridViewPlanning.PeriodViews.Semester
                    Me.PeriodFrom = periodfrom.AddMonths(-6)
                    Me.PeriodUntil = perioduntil.AddMonths(6)
                Case DataGridViewPlanning.PeriodViews.Year
                    Me.PeriodFrom = periodfrom.AddYears(-1)
                    Me.PeriodUntil = perioduntil.AddYears(1)
            End Select
        Else
            Me.PeriodFrom = New Date(Now.Year, 1, 1)
            Me.PeriodUntil = New Date(Now.Year, 12, 31)
        End If
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

    Public Property PrintList() As PrintPlanningRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintPlanningRows)
            objPrintList = value
        End Set
    End Property

    Public Property OutputGuid() As Guid
        Get
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ElementsView() As Integer
        Get
            Return intElementsView
        End Get
        Set(ByVal value As Integer)
            intElementsView = value
        End Set
    End Property

    Public Property ShowDatesColumns As Boolean
        Get
            Return boolShowDatesColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowDatesColumns = value
        End Set
    End Property

    Public Property ShowActivityLinks As Boolean
        Get
            Return boolShowActivityLinks
        End Get
        Set(ByVal value As Boolean)
            boolShowActivityLinks = value
        End Set
    End Property

    Public Property ShowKeyMomentLinks As Boolean
        Get
            Return boolShowKeyMomentLinks
        End Get
        Set(ByVal value As Boolean)
            boolShowKeyMomentLinks = value
        End Set
    End Property

    Public Property PeriodView() As Integer
        Get
            Return intPeriodView
        End Get
        Set(ByVal value As Integer)
            intPeriodView = value
        End Set
    End Property

    Public Property PeriodFrom() As Date
        Get
            Return datPeriodFrom
        End Get
        Set(ByVal value As Date)
            datPeriodFrom = value
        End Set
    End Property

    Public Property PeriodUntil() As Date
        Get
            Return datPeriodUntil
        End Get
        Set(ByVal value As Date)
            datPeriodUntil = value
        End Set
    End Property

    Private Property SortColumnRectangle As Rectangle
        Get
            Return rSortColumnRectangle
        End Get
        Set(ByVal value As Rectangle)
            rSortColumnRectangle = value
        End Set
    End Property

    Private Property TextColumnRectangle As Rectangle
        Get
            Return rTextColumnRectangle
        End Get
        Set(ByVal value As Rectangle)
            rTextColumnRectangle = value
        End Set
    End Property

    Private Property StartDateColumnRectangle As Rectangle
        Get
            Return rColumnStartDate
        End Get
        Set(ByVal value As Rectangle)
            rColumnStartDate = value
        End Set
    End Property

    Private Property EndDateColumnRectangle As Rectangle
        Get
            Return rColumnEndDate
        End Get
        Set(ByVal value As Rectangle)
            rColumnEndDate = value
        End Set
    End Property

    Private Property PlanningColumnRectangle As Rectangle
        Get
            Return rPlanningColumnRectangle
        End Get
        Set(ByVal value As Rectangle)
            rPlanningColumnRectangle = value
        End Set
    End Property

    Private Property FirstPlanningColumnRectangle As Rectangle
        Get
            Return rFirstPlanningColumnRectangle
        End Get
        Set(ByVal value As Rectangle)
            rFirstPlanningColumnRectangle = value
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

    Public Property RepeatRowHeaders As Boolean
        Get
            Return My.Settings.setPrintPlanningRepeatRowHeaders
        End Get
        Set(ByVal value As Boolean)
            My.Settings.setPrintPlanningRepeatRowHeaders = value
        End Set
    End Property
#End Region

#Region "Create Planning table"
    Public Sub CreateList()
        Dim intPurposeIndex, intOutputIndex As Integer
        Dim strPurposeSort As String = String.Empty, strOutputSort As String = String.Empty

        PrintList.Clear()

        If Me.LogFrame IsNot Nothing Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(intPurposeIndex)

                If Me.LogFrame.Purposes.Count > 1 Then
                    Dim rowPurpose As New PrintPlanningRow(LogFrame.SectionTypes.PurposesSection)
                    With rowPurpose
                        .SortNumber = strPurposeSort
                        .Struct = selPurpose
                        .RowType = PrintPlanningRow.RowTypes.RepeatPurpose
                    End With
                    PrintList.Add(rowPurpose)
                End If

                For Each selOutput As Output In selPurpose.Outputs
                    strOutputSort = LogFrame.CreateSortNumber(intOutputIndex, strPurposeSort)

                    If selPurpose.Outputs.Count > 1 Then
                        Dim rowOutput As New PrintPlanningRow(LogFrame.SectionTypes.OutputsSection)
                        With rowOutput
                            .SortNumber = strOutputSort
                            .Struct = selOutput
                            .RowType = PrintPlanningRow.RowTypes.RepeatOutput
                        End With
                        PrintList.Add(rowOutput)
                    End If
                    CreateList_Output(selOutput, strOutputSort)

                    intOutputIndex += 1
                Next

                intPurposeIndex += 1
            Next
        End If

        CreateList_LinksIndices()
    End Sub

    Private Sub CreateList_Output(ByVal selOutput As Output, ByVal strOutputSort As String)

        If selOutput IsNot Nothing Then
            If Me.ElementsView <> ShowElements.ShowActivities Then
                Dim intIndex As Integer

                'Load key moments of this output
                For Each selKeyMoment As KeyMoment In selOutput.KeyMoments.Sort
                    Dim strKeyMomentSort As String = String.Format("#{0}", CurrentLogFrame.CreateSortNumber(intIndex, strOutputSort))
                    Dim NewGridItem As New PrintPlanningRow(LogFrame.SectionTypes.OutputsSection)
                    With NewGridItem
                        .SortNumber = strKeyMomentSort
                        .KeyMoment = selKeyMoment
                        .RowType = PrintPlanningRow.RowTypes.KeyMoment
                    End With
                    PrintList.Add(NewGridItem)
                    intIndex += 1
                Next
            End If

            If Me.ElementsView <> ShowElements.ShowKeyMoments Then
                'Load activities

                CreateList_Activities(selOutput.Activities, strOutputSort, 0)
            End If
        End If
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String, ByVal intLevel As Integer)
        Dim strActivitySort As String
        Dim boolLastHasSubActivities As Boolean
        Dim intIndex As Integer

        For Each selActivity As Activity In selActivities
            Dim NewGridItem As New PrintPlanningRow(LogFrame.SectionTypes.ActivitiesSection)
            strActivitySort = LogFrame.CreateSortNumber(intIndex, strParentSort)
            boolLastHasSubActivities = False

            With NewGridItem
                .SortNumber = strActivitySort
                .Indent = intLevel
                .Struct = selActivity
                .RowType = PrintPlanningRow.RowTypes.Activity
            End With

            PrintList.Add(NewGridItem)

            If selActivity.Activities.Count > 0 Then
                CreateList_Activities(selActivity.Activities, strActivitySort, intLevel + 1)
                boolLastHasSubActivities = True
            End If
            intIndex += 1
        Next
    End Sub

    Private Sub CreateList_LinksIndices()
        For Each selPrintListRow As PrintPlanningRow In Me.PrintList
            If selPrintListRow.Struct IsNot Nothing Then
                Dim selActivity As Activity = TryCast(selPrintListRow.Struct, Activity)

                If selActivity IsNot Nothing AndAlso selActivity.IsProcess = False Then
                    Dim selActivityDetail As ActivityDetail = selActivity.ActivityDetail
                    Dim intRowIndex As Integer = PrintList.IndexOf(selPrintListRow)

                    If selActivityDetail.Relative = True Then
                        Dim intSourceRowIndex As Integer

                        'If selActivityDetail.PeriodDirection = ActivityDetail.DirectionValues.After Then
                        intSourceRowIndex = CreateList_GetRowIndexByReferenceGuid(selActivityDetail.GuidReferenceMoment)

                        If intSourceRowIndex >= 0 Then
                            Dim objSourceRow As PrintPlanningRow = Me.PrintList(intSourceRowIndex)

                            objSourceRow.OutgoingLinksIndices.Add(intRowIndex)
                            selPrintListRow.IncomingLinkIndices.Add(intSourceRowIndex)

                        End If
                        'Else
                        '    Dim intTargetRowIndex As Integer
                        '    intTargetRowIndex = CreateList_GetRowIndexByReferenceGuid(selActivityDetail.GuidReferenceMoment)

                        '    If intTargetRowIndex >= 0 Then
                        '        Dim objTargetRow As PrintPlanningRow = Me.PrintList(intTargetRowIndex)

                        '        selPrintListRow.OutgoingLinksIndices.Add(intTargetRowIndex)
                        '        objTargetRow.IncomingLinkIndices.Add(intRowIndex)

                        '    End If
                        'End If
                    End If
                End If
            ElseIf selPrintListRow.KeyMoment IsNot Nothing Then
                Dim selKeyMoment As KeyMoment = selPrintListRow.KeyMoment
                Dim intRowIndex As Integer = PrintList.IndexOf(selPrintListRow)

                If selKeyMoment.Relative = True Then
                    Dim intSourceRowIndex As Integer

                    'If selKeyMoment.PeriodDirection = KeyMoment.DirectionValues.After Then
                    intSourceRowIndex = CreateList_GetRowIndexByReferenceGuid(selKeyMoment.GuidReferenceMoment)

                    If intSourceRowIndex >= 0 Then
                        Dim objSourceRow As PrintPlanningRow = Me.PrintList(intSourceRowIndex)

                        objSourceRow.OutgoingLinksIndices.Add(intRowIndex)
                        selPrintListRow.IncomingLinkIndices.Add(intSourceRowIndex)

                    End If
                    'Else
                    '    Dim intTargetRowIndex As Integer
                    '    intTargetRowIndex = CreateList_GetRowIndexByReferenceGuid(selKeyMoment.GuidReferenceMoment)

                    '    If intTargetRowIndex >= 0 Then
                    '        Dim objTargetRow As PrintPlanningRow = Me.PrintList(intTargetRowIndex)

                    '        selPrintListRow.OutgoingLinksIndices.Add(intTargetRowIndex)
                    '        objTargetRow.IncomingLinkIndices.Add(intRowIndex)

                    '    End If
                    'End If
                End If
            End If
        Next
    End Sub

    Private Function CreateList_GetRowIndexByReferenceGuid(ByVal selGuid As Guid) As Integer
        Dim intRowIndex As Integer = -1

        For Each selGridRow As PrintPlanningRow In Me.PrintList
            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.KeyMoment
                    Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
                    If selKeyMoment IsNot Nothing AndAlso selKeyMoment.Guid = selGuid Then
                        intRowIndex = PrintList.IndexOf(selGridRow)
                        Exit For
                    End If
                Case PlanningGridRow.RowTypes.Activity
                    If selGridRow.Struct IsNot Nothing Then
                        Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
                        If selActivity IsNot Nothing AndAlso selActivity.Guid = selGuid Then
                            intRowIndex = PrintList.IndexOf(selGridRow)
                            Exit For
                        End If
                    End If
            End Select
        Next

        Return intRowIndex
    End Function
#End Region

#Region "Set column widths"
    Private Sub SetSortColumnRectangle()
        Dim intWidth, intSortWidth As Integer

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintPlanningRow In Me.PrintList
                If selRow.SortNumber Is Nothing Then selRow.SortNumber = String.Empty

                intWidth = PageGraph.MeasureString(selRow.SortNumber, fntTextBold).Width

                If intWidth > intSortWidth Then intSortWidth = intWidth
            Next

            intSortWidth += (CONST_HorizontalPadding * 2)
            SortColumnRectangle = New Rectangle(LeftMargin, ContentTop, intSortWidth, ContentHeight)
        End If
    End Sub

    Private Function SetDateColumnRectangle() As Integer
        Dim intWidthStart, intWidthEnd, intDateWidth As Integer
        Dim selActivity As Activity

        If PageGraph IsNot Nothing Then
            For Each selRow As PrintPlanningRow In Me.PrintList
                Select Case selRow.RowType
                    Case PrintPlanningRow.RowTypes.KeyMoment
                        intWidthStart = PageGraph.MeasureString(selRow.KeyMoment.ExactDateKeyMoment.ToString("d"), fntText).Width

                        If intWidthStart > intDateWidth Then intDateWidth = intWidthStart
                    Case PrintPlanningRow.RowTypes.Activity
                        selActivity = CType(selRow.Struct, Activity)
                        intWidthStart = PageGraph.MeasureString(selActivity.ExactStartDate.ToString("d"), fntText).Width
                        If intWidthStart > intDateWidth Then intDateWidth = intWidthStart

                        intWidthEnd = PageGraph.MeasureString(selActivity.ExactEndDate.ToString("d"), fntText).Width
                        If intWidthEnd > intDateWidth Then intDateWidth = intWidthEnd
                End Select
            Next

            intDateWidth += (CONST_HorizontalPadding * 2)
        End If

        Return intDateWidth
    End Function

    Private Sub SetColumnsWidth()
        Dim spanPeriod As TimeSpan = Me.PeriodUntil - Me.PeriodFrom
        Dim intDateWidth As Integer

        spanPeriod = spanPeriod.Add(New TimeSpan(1, 0, 0, 0))

        'sort column
        SetSortColumnRectangle()

        If ShowDatesColumns = True Then intDateWidth = SetDateColumnRectangle()

        Dim intAvailableWidth As Integer = Me.ContentWidth - SortColumnRectangle.Width - (intDateWidth * 2)
        Dim intFirstPlanningColumnWidth As Integer = SetFirstPlanningColumnWidth(intAvailableWidth * 0.8, spanPeriod)

        'text column
        Dim intTextColumnWidth As Integer = intAvailableWidth - intFirstPlanningColumnWidth
        TextColumnRectangle = New Rectangle(SortColumnRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight)

        'dates column
        StartDateColumnRectangle = New Rectangle(TextColumnRectangle.Right, ContentTop, intDateWidth, ContentHeight)
        EndDateColumnRectangle = New Rectangle(StartDateColumnRectangle.Right, ContentTop, intDateWidth, ContentHeight)

        'planning column
        FirstPlanningColumnRectangle = New Rectangle(EndDateColumnRectangle.Right, ContentTop, intFirstPlanningColumnWidth, ContentHeight)

        If RepeatRowHeaders = False Then
            intAvailableWidth = Me.ContentWidth - SortColumnRectangle.Width
            Dim intPlanningColumnWidth As Integer = SetPlanningColumnWidth(intAvailableWidth)

            PlanningColumnRectangle = New Rectangle(SortColumnRectangle.Right, ContentTop, intPlanningColumnWidth, ContentHeight)
        Else
            PlanningColumnRectangle = FirstPlanningColumnRectangle
        End If

        'Total planning
        rTotalPlanning.Width = spanPeriod.Days * WidthOfDay
        If rTotalPlanning.Width > FirstPlanningColumnRectangle.Width Then
            Dim intWidth As Integer
            intWidth = rTotalPlanning.Width - FirstPlanningColumnRectangle.Width
            HorPages = Math.Ceiling(intWidth / PlanningColumnRectangle.Width) + 1
        Else
            HorPages = 1
        End If

        'Project duration
        Dim datProjectStart As Date = LimitStartDateToPeriod(Me.LogFrame.StartDate)
        Dim datProjectEnd As Date = LimitEndDateToPeriod(Me.LogFrame.EndDate)

        rProjectDuration = New Rectangle(GetCoordinateX(datProjectStart), ContentTop, GetActivityWidth(datProjectStart, datProjectEnd), ContentHeight)
    End Sub

    Private Function SetFirstPlanningColumnWidth(ByVal intFirstPlanningColumnWidth As Integer, ByVal spanPeriod As TimeSpan) As Integer
        Dim intPlanningUnits As Integer

        Select Case PeriodView
            Case DataGridViewPlanning.PeriodViews.Day
                WidthOfDay = 30
                intPlanningUnits = intFirstPlanningColumnWidth / WidthOfDay
                intFirstPlanningColumnWidth = intPlanningUnits * WidthOfDay
            Case DataGridViewPlanning.PeriodViews.Week
                WidthOfDay = 10
                intPlanningUnits = intFirstPlanningColumnWidth / (WidthOfDay * 7)
                intFirstPlanningColumnWidth = intPlanningUnits * (WidthOfDay * 7)

                If PeriodFrom.DayOfWeek <> DayOfWeek.Monday Then
                    intWeekCorrection = DayOfWeek.Monday - PeriodFrom.DayOfWeek
                    PeriodFrom = PeriodFrom.AddDays(intWeekCorrection)
                End If
                If PeriodUntil.DayOfWeek <> DayOfWeek.Sunday Then
                    PeriodUntil = PeriodUntil.AddDays(PeriodUntil.DayOfWeek)
                End If
            Case DataGridViewPlanning.PeriodViews.Month
                WidthOfDay = 1
            Case DataGridViewPlanning.PeriodViews.Year
                WidthOfDay = 1 / spanPeriod.Days * intFirstPlanningColumnWidth
        End Select

        Return intFirstPlanningColumnWidth
    End Function

    Private Function SetPlanningColumnWidth(ByVal intAvailableWidth As Integer) As Integer
        Dim intPlanningUnits As Integer

        Dim intPlanningColumnWidth As Integer = intAvailableWidth
        Select Case PeriodView
            Case DataGridViewPlanning.PeriodViews.Day
                intPlanningUnits = intPlanningColumnWidth / WidthOfDay
                intPlanningColumnWidth = intPlanningUnits * WidthOfDay
            Case DataGridViewPlanning.PeriodViews.Week
                intPlanningUnits = intPlanningColumnWidth / (WidthOfDay * 7)
                intPlanningColumnWidth = intPlanningUnits * (WidthOfDay * 7)
        End Select

        Return intPlanningColumnWidth
    End Function
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintPlanningRow In Me.PrintList
            Select Case selRow.RowType
                Case PrintPlanningRow.RowTypes.RepeatOutput, PrintPlanningRow.RowTypes.RepeatPurpose
                    ReloadImages_Repeats(selRow)
                Case Else
                    ReloadImages_Normal(selRow)
            End Select
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Repeats(ByVal selRow As PrintPlanningRow)
        Dim intTotalTextWidth As Integer = Me.ContentWidth - SortColumnRectangle.Width

        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                If String.IsNullOrEmpty(selRow.Struct.Text) = False Then
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(intTotalTextWidth, selRow.Struct.RTF, False)
                End If
            End If
        End With
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintPlanningRow)
        Dim intColumnWidth As Integer = TextColumnRectangle.Width

        With RichTextManager
            If selRow.RowType = PrintPlanningRow.RowTypes.KeyMoment Then
                If selRow.KeyMoment IsNot Nothing Then
                    selRow.KeyMomentCellImage = .TextWithPaddingToBitmap(intColumnWidth, selRow.KeyMoment.Description)
                End If
            Else
                If selRow.Struct IsNot Nothing Then
                    If String.IsNullOrEmpty(selRow.Struct.Text) Then
                        selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, selRow.Struct.GetItemName(selRow.Section), selRow.SortNumber, False)
                    Else
                        selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.Struct.RTF, False)
                    End If
                End If
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintPlanningRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > 0 Then selPrintListRow.RowHeight = intRowHeight Else selPrintListRow.RowHeight = NewCellHeight()

        'column headers
        If PageGraph IsNot Nothing Then _
            ColumnHeadersHeight = PageGraph.MeasureString(LANG_DataCollection, CurrentLogFrame.DetailsFont, Me.TextColumnRectangle.Width).Height
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selPrintListRow As PrintPlanningRow = Me.PrintList(RowIndex)

        If selPrintListRow.RowType = PrintPlanningRow.RowTypes.KeyMoment Then
            If selPrintListRow.KeyMoment IsNot Nothing AndAlso String.IsNullOrEmpty(selPrintListRow.KeyMoment.Description) = False Then
                If selPrintListRow.KeyMomentCellImage.Height > intRowHeight Then intRowHeight = selPrintListRow.KeyMomentCellImage.Height

                Dim intKeyMomentHeight As Integer = CurrentLogFrame.DetailsFont.Height
                intKeyMomentHeight += 10

                If intRowHeight < intKeyMomentHeight Then intRowHeight = intKeyMomentHeight
            End If
        ElseIf selPrintListRow.Struct IsNot Nothing AndAlso String.IsNullOrEmpty(selPrintListRow.Struct.RTF) = False Then
            If selPrintListRow.Struct.CellImage.Height > intRowHeight Then intRowHeight = selPrintListRow.Struct.CellImage.Height
        End If

        Return intRowHeight
    End Function

    Public Sub SetColumnHeadersHeight()
        If PageGraph IsNot Nothing Then
            Dim fntHeader As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
            Dim intHeight As Integer = PageGraph.MeasureString(Activity.ItemName, fntHeader, SortColumnRectangle.Width + TextColumnRectangle.Width).Height

            rHeadingFirstRow = New Rectangle(LeftMargin, ContentTop, ContentWidth, intHeight)
            rHeadingSecondRow = New Rectangle(LeftMargin, rHeadingFirstRow.Bottom, ContentWidth, 30)

            ColumnHeadersHeight = rHeadingFirstRow.Height + rHeadingSecondRow.Height
        End If
    End Sub
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight - ColumnHeadersHeight

        For Each selRow As PrintPlanningRow In PrintList
            intTotalHeight += selRow.RowHeight
        Next

        decPages = intTotalHeight / intAvailableHeight
        decPages = Math.Ceiling(decPages)
        decPages *= HorPages

        Return decPages
    End Function

    Private Function LimitStartDateToPeriod(ByVal datStartDate As Date) As Date
        If datStartDate < Me.PeriodFrom Then datStartDate = Me.PeriodFrom

        Return datStartDate
    End Function

    Private Function LimitEndDateToPeriod(ByVal datEndDate As Date) As Date
        If datEndDate > Me.PeriodUntil Then datEndDate = Me.PeriodUntil

        Return datEndDate
    End Function

    Private Function GetCoordinateX(ByVal datStart As Date) As Integer
        Dim intX As Integer
        intX = (datStart.Subtract(Me.PeriodFrom).Days * WidthOfDay)
        Return intX
    End Function

    Private Function GetCoordinateY(ByVal intRowHeight As Integer, Optional ByVal boolPreparation As Boolean = False) As Integer
        Dim intY As Integer

        If boolPreparation = False Then
            intY = LastRowY + ((intRowHeight - CONST_BarHeight) / 2)
        Else
            intY = LastRowY + ((intRowHeight - CONST_PreparationHeight) / 2)
        End If

        Return intY
    End Function

    Private Function GetActivityWidth(ByVal datStart As Date, ByVal datEnd As Date) As Integer
        Dim intWidth As Integer
        intWidth = ((datEnd.Subtract(datStart).Days) + 1) * WidthOfDay
        Return intWidth
    End Function

    Private Function GetFollowUpWidth(ByVal datEndMainActivity As Date, ByVal datEndFollowUp As Date) As Integer
        Dim intWidth As Integer
        intWidth = ((datEndFollowUp.Subtract(datEndMainActivity).Days)) * WidthOfDay
        Return intWidth
    End Function

    Private Function GetCurrentPlanningColumnWidth() As Integer
        If CurrentHorPage = 1 Then
            Return FirstPlanningColumnRectangle.Width
        Else
            Return PlanningColumnRectangle.Width
        End If
    End Function

    Private Function GetCurrentPlanningColumnX() As Integer
        If CurrentHorPage = 1 Then
            Return FirstPlanningColumnRectangle.X
        Else
            Return PlanningColumnRectangle.X
        End If
    End Function

    Private Function GetCurrentPlanningColumnRight() As Integer
        If CurrentHorPage = 1 Then
            Return FirstPlanningColumnRectangle.X + FirstPlanningColumnRectangle.Width
        Else
            Return PlanningColumnRectangle.X + PlanningColumnRectangle.Width
        End If
    End Function

    Private Function GetAvailableArea() As Rectangle
        Dim rArea As New Rectangle(GetCurrentPlanningColumnX(), ContentTop + ColumnHeadersHeight, 0, ContentHeight - ColumnHeadersHeight)
        Dim intRightBorder As Integer = GetCoordinateX(PeriodUntil.AddDays(1)) - rCurrentSection.X

        If intRightBorder < GetCurrentPlanningColumnWidth() Then
            rArea.Width = intRightBorder
        Else
            rArea.Width = GetCurrentPlanningColumnWidth()
        End If

        Return rArea
    End Function

    Private Function CoordinateToDate(ByVal intX As Integer) As Date
        Dim selDate As Date

        intX = intX + rCurrentSection.X - GetCurrentPlanningColumnX()

        selDate = datPeriodFrom.AddDays(intX / WidthOfDay)
        Return selDate
    End Function

    Private Function CoordinateToCurrentSection(ByVal intX As Integer) As Integer
        intX -= rCurrentSection.X
        If intX < 0 Then intX = 0
        intX += GetCurrentPlanningColumnX()

        Return intX
    End Function

    Private Sub DetermineActivityStyle(ByVal intType As Integer)
        Select Case intType
            Case ActivityDetail.Types.Other
                ActivityColor = Color.DarkSeaGreen

            Case ActivityDetail.Types.Planning, ActivityDetail.Types.ProjectProposal, ActivityDetail.Types.DonorSigned
                ActivityColor = Color.Plum

            Case ActivityDetail.Types.Evaluation, ActivityDetail.Types.Audit, ActivityDetail.Types.Monitoring, ActivityDetail.Types.Report
                ActivityColor = Color.Tomato

            Case ActivityDetail.Types.Identification, ActivityDetail.Types.Selection, ActivityDetail.Types.Registration
                ActivityColor = Color.Orange

            Case ActivityDetail.Types.HiringStaff, ActivityDetail.Types.Procurement, ActivityDetail.Types.Logistics, ActivityDetail.Types.Debriefing
                ActivityColor = Color.CadetBlue

            Case ActivityDetail.Types.Meeting
                ActivityColor = Color.Blue

            Case ActivityDetail.Types.Travel
                ActivityColor = Color.MediumBlue

            Case ActivityDetail.Types.Activity, ActivityDetail.Types.Construction, ActivityDetail.Types.Distribution, _
                ActivityDetail.Types.MedicalTreatment, ActivityDetail.Types.HumanitarianAssistance, ActivityDetail.Types.PeaceBuilding
                ActivityColor = Color.LightGreen

            Case ActivityDetail.Types.Research, ActivityDetail.Types.Training, ActivityDetail.Types.CapacityDevelopment, ActivityDetail.Types.Awareness
                ActivityColor = Color.Yellow

            Case Else
                ActivityColor = Color.PaleGreen
        End Select
    End Sub

    Public Function MakeGradientBrush(ByVal BaseColor As Color, ByVal intX As Integer, ByVal intY As Integer, ByVal intHeight As Integer) As LinearGradientBrush
        Dim colLight As Color = BaseColor
        Dim colDark As Color = DarkerColor(colLight)
        If intX < 0 Then intX = 0
        If intHeight = 0 Then Return Nothing
        Dim NewBrush As New LinearGradientBrush(New Point(intX, intY), New Point(intX, intY + intHeight), colLight, colDark)
        Return NewBrush
    End Function
#End Region

#Region "Print page"
    Protected Overrides Sub OnBeginPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnBeginPrint(e)

        PrintRectangles.Clear()
        rCurrentSection = New Rectangle

        boolColumnsWidthSet = False
        intWeekCorrection = 0
        WidthOfDay = 0
        ColumnHeadersHeight = 0
        HorPages = 0
        CurrentHorPage = 1
        intStartIndex = 0
        rTotalPlanning.X = 0
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
            SetCurrentSectionWidth()

            boolColumnsWidthSet = True
        End If

        Dim intRowCount As Integer = PrintList.Count
        Dim selRow As PrintPlanningRow = PrintList(RowIndex)

        LastRowY = ContentTop

        If CurrentHorPage = 1 Then intStartIndex = RowIndex

        'Print Header
        PrintHeader()
        PrintColumnHeaders()

        If PrintList.Count > 0 Then
            Do While LastRowY + PrintList(RowIndex).RowHeight < Me.ContentBottom
                selRow = PrintList(RowIndex)

                'Print sort number
                PrintPage_PrintSortNumber(selRow)

                'Print activity name/key moment name
                If CurrentHorPage = 1 Or RepeatRowHeaders = True Then
                    'only print the label on the first page when HideTextColumn is true, otherwise print it on all the pages
                    PrintPage_PrintRtf(selRow)
                    PrintPage_PrintDates(selRow)
                End If

                'draw bars
                PrintPage_PrintGanttChart(selRow)

                LastRowY += selRow.RowHeight
                RowIndex += 1
                If RowIndex > PrintList.Count - 1 Then Exit Do
            Loop
            'DrawLinks()
        Else
            RaiseEvent PagePrinted(Me, New LinePrintedEventArgs(0, 0))
        End If
        PrintFooter()
        RaiseEvent PagePrinted(Me, New LinePrintedEventArgs(PageNumber - 1, TotalPages))
        LastRowY = ContentTop
        PageNumber += 1


        If CurrentHorPage < HorPages Then
            RowIndex = intStartIndex
            rCurrentSection.X += GetCurrentPlanningColumnWidth()
            CurrentHorPage += 1

            SetCurrentSectionWidth()

            e.HasMorePages = True
        Else
            If RowIndex < intRowCount Then
                CurrentHorPage = 1
                rCurrentSection.X = 0
                SetCurrentSectionWidth()

                e.HasMorePages = True
            Else
                e.HasMorePages = False
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintSortNumber(ByVal selRow As PrintPlanningRow)
        Dim rSortNumber As New Rectangle(SortColumnRectangle.X, LastRowY, SortColumnRectangle.Width, selRow.RowHeight)
        PrintPage_PrintText(selRow.SortNumber, rSortNumber)
    End Sub

    Private Sub PrintPage_PrintRtf(ByVal selRow As PrintPlanningRow)
        Dim rImage As Rectangle

        Select Case selRow.RowType
            Case PrintPlanningRow.RowTypes.RepeatPurpose, PrintPlanningRow.RowTypes.RepeatOutput
                Dim intRightBorder As Integer = GetCoordinateX(PeriodUntil.AddDays(1)) - rCurrentSection.X
                Dim intPlanningColumnWidth As Integer
                Dim intCurrentPlanningColumnWidth As Integer = GetCurrentPlanningColumnWidth()

                If intRightBorder < intCurrentPlanningColumnWidth Then
                    intPlanningColumnWidth = intRightBorder
                Else
                    intPlanningColumnWidth = intCurrentPlanningColumnWidth
                End If
                intPlanningColumnWidth += TextColumnRectangle.Width + StartDateColumnRectangle.Width + EndDateColumnRectangle.Width
                rImage = New Rectangle(TextColumnRectangle.X, LastRowY, intPlanningColumnWidth, selRow.RowHeight)

                PrintPage_PrintImage(selRow.Struct.CellImage, rImage)
            Case PrintPlanningRow.RowTypes.KeyMoment
                rImage = New Rectangle(TextColumnRectangle.X, LastRowY, TextColumnRectangle.Width, selRow.RowHeight)

                PrintPage_PrintImage(selRow.KeyMomentCellImage, rImage)
            Case PrintPlanningRow.RowTypes.Activity
                rImage = New Rectangle(TextColumnRectangle.X, LastRowY, TextColumnRectangle.Width, selRow.RowHeight)

                PrintPage_PrintImage(selRow.Struct.CellImage, rImage)
        End Select
    End Sub

    Private Sub PrintPage_PrintDates(ByVal selRow As PrintPlanningRow)
        If selRow.RowType = PrintPlanningRow.RowTypes.KeyMoment Or selRow.RowType = PrintPlanningRow.RowTypes.Activity Then
            Dim rStartDate As New Rectangle(StartDateColumnRectangle.X, LastRowY, StartDateColumnRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.StartDate.ToString("d"), rStartDate)
            Dim rEndDate As New Rectangle(EndDateColumnRectangle.X, LastRowY, EndDateColumnRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.EndDate.ToString("d"), rEndDate)
        End If
    End Sub

    Private Sub PrintPage_PrintGanttChart(ByVal selRow As PrintPlanningRow)
        Select Case selRow.RowType
            Case PrintPlanningRow.RowTypes.RepeatPurpose, PrintPlanningRow.RowTypes.RepeatOutput
                PrintPage_PrintBorderRepeats(selRow)
            Case PrintPlanningRow.RowTypes.KeyMoment
                PrintPage_Background(selRow.RowHeight)
                setPrintCoordinates_KeyMoments(selRow)

                PrintKeyMoment(selRow)
                DrawActivityLinks_VerticalActivityLinks(selRow)
                DrawKeyMomentLinks_VerticalKeyMomentLinks(selRow)
            Case PrintPlanningRow.RowTypes.Activity
                If selRow.Struct IsNot Nothing Then
                    Dim selActivity As Activity = CType(selRow.Struct, Activity)
                    PrintPage_Background(selRow.RowHeight)
                    If selActivity.IsProcess = False Then
                        SetPrintCoordinates_Activity(selRow)
                        DetermineActivityStyle(selActivity.ActivityDetail.Type)

                        PrintPreparation(selRow)
                        PrintFollowUp(selRow)
                        PrintActivity(selRow)
                        PrintRepeats(selRow)
                    Else
                        SetPrintCoordinates_Process(selRow)

                        PrintProcess(selRow)
                    End If

                End If

                DrawActivityLinks_VerticalActivityLinks(selRow)
                DrawKeyMomentLinks_VerticalKeyMomentLinks(selRow)
        End Select
    End Sub

    Private Sub PrintPage_PrintText(ByVal strValue As String, ByVal rPrint As Rectangle, _
                          Optional ByVal boolHeader As Boolean = False, _
                          Optional ByVal boolIsKeyMoment As Boolean = False)
        If PageGraph IsNot Nothing Then
            If boolHeader = True Then
                PageGraph.FillRectangle(Brushes.LightGray, rPrint)
            End If
            PageGraph.DrawRectangle(penBlack1, rPrint)

            Dim formatCells As New StringFormat()
            Dim brText As SolidBrush = New SolidBrush(Color.Black)
            Dim rText As Rectangle = GetTextRectangle(rPrint)

            If boolHeader = True Then
                formatCells.Alignment = StringAlignment.Center
                formatCells.LineAlignment = StringAlignment.Center
                PageGraph.DrawString(strValue, fntTextBold, brText, rText, formatCells)
            ElseIf boolIsKeyMoment = True Then
                formatCells.Alignment = StringAlignment.Near
                formatCells.LineAlignment = StringAlignment.Center
                PageGraph.DrawString(strValue, fntText, brText, rText, formatCells)
            Else
                formatCells.Alignment = StringAlignment.Near
                formatCells.LineAlignment = StringAlignment.Near
                PageGraph.DrawString(strValue, fntText, brText, rText, formatCells)
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintImage(ByVal bmCellImage As Bitmap, ByVal rImage As Rectangle)
        If PageGraph IsNot Nothing Then
            If bmCellImage IsNot Nothing Then PageGraph.DrawImage(bmCellImage, New Point(rImage.X, rImage.Y))
            PageGraph.DrawRectangle(penBlack1, rImage)
        End If
    End Sub

    Private Sub PrintPage_PrintBorderRepeats(ByVal selRow As PrintPlanningRow)
        Dim rImage As Rectangle

        If CurrentHorPage <> 1 And RepeatRowHeaders = False Then
            Dim intRightBorder As Integer = GetCoordinateX(PeriodUntil.AddDays(1)) - rCurrentSection.X
            Dim intPlanningColumnWidth As Integer

            If intRightBorder < GetCurrentPlanningColumnWidth() Then
                intPlanningColumnWidth = intRightBorder
            Else
                intPlanningColumnWidth = GetCurrentPlanningColumnWidth()
            End If

            rImage = New Rectangle(TextColumnRectangle.X, LastRowY, intPlanningColumnWidth, selRow.RowHeight)
            PageGraph.DrawRectangle(penBlack1, rImage)
        End If
    End Sub

    Private Sub PrintPage_Background(ByVal intRowHeight As Integer)
        'background
        Dim rBackGround As New Rectangle
        Dim rAvailableArea As Rectangle = GetAvailableArea()
        Dim brDuration As Brush = Brushes.Gainsboro
        Dim brWeekEnd As New SolidBrush(DarkerColor(Color.Gainsboro))

        Dim intX, intY As Integer
        Dim penGrid As New Pen(SystemColors.ControlDark, 1)
        Dim intRightBorder As Integer = GetCoordinateX(PeriodUntil.AddDays(1)) - rCurrentSection.X
        Dim intPlanningColumnWidth As Integer
        Dim intCurrentPlanningColumnX As Integer = GetCurrentPlanningColumnX()

        If PageGraph IsNot Nothing Then
            If intRightBorder < GetCurrentPlanningColumnWidth() Then
                intPlanningColumnWidth = intRightBorder
            Else
                intPlanningColumnWidth = GetCurrentPlanningColumnWidth()
            End If

            'Project duration
            rBackGround = rProjectDuration
            rBackGround.Y = LastRowY
            rBackGround.Height = intRowHeight

            If rBackGround.Left <= rCurrentSection.Right And rBackGround.Right >= Me.rCurrentSection.Left Then
                rBackGround.X -= rCurrentSection.X
                If rBackGround.X < 0 Then
                    rBackGround.Width += rBackGround.X
                    rBackGround.X = 0
                End If
                rBackGround.X += GetCurrentPlanningColumnX()

                If rBackGround.Right > rAvailableArea.Right Then
                    rBackGround.Width = rAvailableArea.Right - rBackGround.X
                End If
                PageGraph.FillRectangle(brDuration, rBackGround)
            End If

            For i = Me.PeriodFrom.Year To Me.PeriodUntil.Year
                Dim NewYear As Date = New Date(i, 1, 1, 0, 0, 0)
                intX = GetCoordinateX(NewYear)
                intY = LastRowY
                If intX >= Me.rCurrentSection.X And intX <= rCurrentSection.Right Then
                    intX -= rCurrentSection.X
                    intX += intCurrentPlanningColumnX
                    PageGraph.DrawLine(penGrid, intX, intY, intX, intY + intRowHeight)
                End If

                If Me.PeriodView = DataGridViewPlanning.PeriodViews.Month Then
                    For j = 2 To 12
                        Dim NewMonth As Date = New Date(i, j, 1, 0, 0, 0)
                        intX = GetCoordinateX(NewMonth)
                        If intX >= rCurrentSection.X And intX <= rCurrentSection.Right Then
                            intX -= rCurrentSection.X
                            intX += intCurrentPlanningColumnX
                            PageGraph.DrawLine(penGrid, intX, intY, intX, intY + intRowHeight)
                        End If
                    Next

                ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Then
                    Dim intDayOfWeek As Integer = DatePart(DateInterval.Weekday, NewYear, FirstDayOfWeek.Monday, FirstWeekOfYear.System) - 1
                    Dim FirstMonday As Date = New Date(i, 1, 8 - intDayOfWeek)
                    Dim NewWeek As Date = FirstMonday

                    Do While NewWeek.Year < i + 1
                        intX = GetCoordinateX(NewWeek)
                        If intX >= Me.rCurrentSection.X And intX <= rCurrentSection.Right Then
                            intX -= rCurrentSection.X
                            intX += intCurrentPlanningColumnX
                            PageGraph.DrawLine(penGrid, intX, intY, intX, intY + intRowHeight)
                        End If
                        NewWeek = NewWeek.AddDays(7)
                    Loop
                ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                    Dim NewDay As Date = NewYear
                    Do While NewDay.Year < i + 1
                        intX = GetCoordinateX(NewDay)
                        If intX >= Me.rCurrentSection.X And intX + WidthOfDay <= rCurrentSection.Right Then
                            intX -= rCurrentSection.X
                            intX += intCurrentPlanningColumnX
                            If NewDay.DayOfWeek = DayOfWeek.Saturday Or NewDay.DayOfWeek = DayOfWeek.Sunday Then
                                PageGraph.FillRectangle(brWeekEnd, intX, intY, WidthOfDay, intRowHeight)
                            End If
                            PageGraph.DrawLine(penGrid, intX, intY, intX, intY + intRowHeight)
                        End If
                        NewDay = NewDay.AddDays(1)
                    Loop
                End If
            Next

            rPlanningCell = New Rectangle(GetCurrentPlanningColumnX, LastRowY, intPlanningColumnWidth, intRowHeight)
            PageGraph.DrawRectangle(penBlack1, rPlanningCell)
        End If
    End Sub

    Private Sub SetCurrentSectionWidth()
        Dim intRightBorder As Integer = GetCoordinateX(PeriodUntil.AddDays(1)) - rCurrentSection.X
        If intRightBorder < GetCurrentPlanningColumnWidth() Then
            rCurrentSection.Width = intRightBorder
        Else
            rCurrentSection.Width = GetCurrentPlanningColumnWidth()
        End If
    End Sub
#End Region

#Region "set print coordinates"
    Private Sub SetPrintCoordinates_Activity(ByVal selRow As PrintPlanningRow)
        Dim datStartDate, datEndDate, datStartPreparation, datEndFollowUp As Date
        Dim intActivityX As Integer, intActivityY As Integer
        Dim intActivityWidth, intActivityHeight As Integer
        Dim intPreparationX, intPreparationY As Integer
        Dim intPreparationWidth, intPreparationHeight As Integer
        Dim intFollowUpX, intFollowUpY As Integer
        Dim intFollowUpWidth As Integer, intFollowUpHeight As Integer
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity.ActivityDetail
            If selRow.StartDate < PeriodUntil And selRow.EndDate > PeriodFrom Then
                datStartDate = LimitStartDateToPeriod(selRow.StartDate)
                datEndDate = LimitEndDateToPeriod(selRow.EndDate)
                intActivityX = GetCoordinateX(datStartDate)
                intActivityY = GetCoordinateY(selRow.RowHeight)
                intActivityWidth = GetActivityWidth(datStartDate, datEndDate)
                intActivityHeight = CONST_BarHeight

                rActivity = New Rectangle(intActivityX, intActivityY, intActivityWidth, intActivityHeight)
            Else
                rActivity = Nothing
            End If

            If .StartDatePreparation < PeriodUntil And selRow.StartDate > PeriodFrom Then
                datStartPreparation = LimitStartDateToPeriod(.StartDatePreparation)
                intPreparationX = GetCoordinateX(datStartPreparation)
                intPreparationY = GetCoordinateY(selRow.RowHeight, True)
                intPreparationWidth = intActivityX - intPreparationX
                intPreparationHeight = CONST_PreparationHeight


                rPreparation = New Rectangle(intPreparationX, intPreparationY, intPreparationWidth, intPreparationHeight)
            Else
                rPreparation = Nothing
            End If

            If selRow.EndDate < PeriodUntil And .EndDateFollowUp > PeriodFrom Then
                datEndFollowUp = LimitEndDateToPeriod(.EndDateFollowUp)
                intFollowUpX = rActivity.Right
                intFollowUpY = GetCoordinateY(selRow.RowHeight, True)
                intFollowUpWidth = GetActivityWidth(datEndDate.AddDays(1), datEndFollowUp)
                intFollowUpHeight = CONST_FollowUpHeight


                rFollowUp = New Rectangle(intFollowUpX, intFollowUpY, intFollowUpWidth, intFollowUpHeight)
            Else
                rFollowUp = Nothing
            End If
        End With
    End Sub

    Private Sub SetPrintCoordinates_Process(ByVal selRow As PrintPlanningRow)
        Dim datStartDate, datEndDate As Date
        Dim intActivityX As Integer, intActivityY As Integer
        Dim intActivityWidth, intActivityHeight As Integer
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity
            If .ExactStartDate < PeriodUntil And .ExactEndDate > PeriodFrom Then
                datStartDate = LimitStartDateToPeriod(.ExactStartDate)
                datEndDate = LimitEndDateToPeriod(.ExactEndDate)

                intActivityX = GetCoordinateX(.ExactStartDate)
                intActivityY = GetCoordinateY(selRow.RowHeight)
                intActivityWidth = GetActivityWidth(.ExactStartDate, .ExactEndDate)
                intActivityHeight = CONST_PreparationHeight

                rActivity = New Rectangle(intActivityX, intActivityY, intActivityWidth, intActivityHeight)
            Else
                rActivity = Nothing
            End If
        End With
    End Sub

    Private Function setPrintCoordinates_Repeats(ByVal selStartDate As Date, ByVal selEndDate As Date, ByVal intRowHeight As Integer) As Rectangle
        Dim intX, intY As Integer
        Dim intWidth, intHeight As Integer

        selStartDate = LimitStartDateToPeriod(selStartDate)
        selEndDate = LimitEndDateToPeriod(selEndDate)
        intX = GetCoordinateX(selStartDate)
        intWidth = GetActivityWidth(selStartDate, selEndDate)
        intHeight = CurrentLogFrame.DetailsFont.Height
        intY = LastRowY + ((intRowHeight - intHeight) / 2)

        Dim rect As New Rectangle(intX, intY, intWidth, intHeight)

        Return rect
    End Function

    Private Sub setPrintCoordinates_KeyMoments(ByVal selRow As PrintPlanningRow)
        Dim intKeyMomentX As Integer, intKeyMomentY As Integer
        Dim intKeyMomentWidth, intKeyMomentHeight As Integer

        If selRow.StartDate < PeriodUntil And selRow.EndDate > PeriodFrom Then
            intKeyMomentX = GetCoordinateX(selRow.StartDate)
            If Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then intKeyMomentX += (WidthOfDay / 2)

            intKeyMomentWidth = GetActivityWidth(selRow.StartDate, selRow.EndDate)
            intKeyMomentHeight = CurrentLogFrame.DetailsFont.Height
            intKeyMomentY = LastRowY + ((selRow.RowHeight - intKeyMomentHeight) / 2)

            rKeyMoment = New Rectangle(intKeyMomentX, intKeyMomentY, intKeyMomentWidth, intKeyMomentHeight)
        Else
            rKeyMoment = Nothing
        End If

    End Sub
#End Region

#Region "Print Gantt chart area"
    Private Sub PrintActivity(ByVal selRow As PrintPlanningRow)
        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity.ActivityDetail
            If rActivity.Left <= rCurrentSection.Right And rActivity.Right > rCurrentSection.Left Then
                If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Then
                    strStart = selRow.StartDate.ToString("d-MMM")
                    strEnd = selRow.EndDate.ToString("d-MMM")
                ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                    strStart = selRow.StartDate.ToString("d-MMM")
                    If String.IsNullOrEmpty(.Location) = False Then strStart = String.Format("{0} - {1}", strStart, .Location)
                    If String.IsNullOrEmpty(.Organisation) = False Then strStart = String.Format("{0} - {1}", strStart, .Organisation)
                    strEnd = selRow.EndDate.ToString("d-MMM")
                End If
                DrawActivity(rActivity, strStart, strEnd)
                DrawActivityLinks(selRow)
            End If
        End With
    End Sub

    Private Sub PrintProcess(ByVal selRow As PrintPlanningRow)
        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity.ActivityDetail
            If rActivity.Left <= rCurrentSection.Right And rActivity.Right > rCurrentSection.Left Then
                If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Then
                    strStart = selActivity.ExactStartDate.ToString("d-MMM")
                    strEnd = selActivity.ExactEndDate.ToString("d-MMM")
                ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                    strStart = selActivity.ExactStartDate.ToString("d-MMM")
                    If String.IsNullOrEmpty(.Location) = False Then strStart = String.Format("{0} - {1}", strStart, .Location)
                    If String.IsNullOrEmpty(.Organisation) = False Then strStart = String.Format("{0} - {1}", strStart, .Organisation)
                    strEnd = selActivity.ExactEndDate.ToString("d-MMM")
                End If

                DrawProcess(rActivity, selRow.RowHeight, strStart, strEnd)
                DrawActivityLinks(selRow)
            End If
        End With
    End Sub

    Private Sub PrintPreparation(ByVal selRow As PrintPlanningRow)
        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity.ActivityDetail
            If .StartDatePreparation <> selRow.StartDate Then
                If rPreparation.Left <= rCurrentSection.Right And rPreparation.Right > Me.rCurrentSection.Left Then
                    If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                        strStart = .StartDatePreparation.ToString("d-MMM")
                        strEnd = String.Empty
                    End If
                    DrawActivity(rPreparation, strStart, strEnd)
                End If
            End If
        End With
    End Sub

    Private Sub PrintFollowUp(ByVal selRow As PrintPlanningRow)
        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity.ActivityDetail
            If .EndDateFollowUp <> selRow.EndDate Then
                If rFollowUp.Left <= rCurrentSection.Right And rFollowUp.Right > Me.rCurrentSection.Left Then
                    If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                        strStart = String.Empty
                        strEnd = .EndDateFollowUp.ToString("d-MMM")
                    End If
                    DrawActivity(rFollowUp, strStart, strEnd)
                End If
            End If
        End With
    End Sub

    Private Sub PrintRepeats(ByVal selRow As PrintPlanningRow)
        Dim rBase As Rectangle
        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim selActivity As Activity = CType(selRow.Struct, Activity)

        If selActivity Is Nothing Then Exit Sub

        With selActivity.ActivityDetail
            If .RepeatStartDates.Count > 0 Then
                Dim selStartDate As Date
                Dim selEndDate As Date
                For i = 0 To .RepeatStartDates.Count - 1
                    selStartDate = .RepeatStartDates(i)
                    selEndDate = .RepeatEndDates(i)

                    rBase = setPrintCoordinates_Repeats(selStartDate, selEndDate, selRow.RowHeight)

                    'draw preparation period of repeated activity
                    If .PreparationRepeat = True And .StartDatePreparation <> selRow.StartDate Then
                        rPreparation.X = rBase.Left - rPreparation.Width
                        If rPreparation.Left <= rCurrentSection.Right And rPreparation.Right > rCurrentSection.Left Then
                            If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                                strStart = selStartDate.AddDays((selRow.StartDate.Subtract(.StartDatePreparation).Days) * -1).ToString("d-MMM")
                                strEnd = String.Empty
                            End If
                            DrawActivity(rPreparation, strStart, strEnd)
                        End If
                    End If

                    'draw follow-up period of repeated activity
                    If .FollowUpRepeat = True And .EndDateFollowUp <> selRow.EndDate Then
                        rFollowUp.X = rBase.Left + rBase.Width
                        If rFollowUp.Left <= rCurrentSection.Right And rFollowUp.Right > rCurrentSection.Left Then
                            If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                                strStart = ""
                                strEnd = selEndDate.AddDays((.EndDateFollowUp.Subtract(selRow.EndDate).Days)).ToString("d-MMM")
                            End If
                            DrawActivity(rFollowUp, strStart, strEnd)
                        End If
                    End If

                    'draw repeated activity
                    If rBase.Left <= rCurrentSection.Right And rBase.Right > rCurrentSection.Left Then
                        If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Then
                            strStart = selStartDate.ToString("d-MMM")
                            strEnd = selEndDate.ToString("d-MMM")
                        ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                            strStart = selStartDate.ToString("d-MMM")
                            If String.IsNullOrEmpty(.Location) = False Then strStart = String.Format("{0} - {1}", strStart, .Location)
                            If String.IsNullOrEmpty(.Organisation) = False Then strStart = String.Format("{0} - {1}", strStart, .Organisation)
                            strEnd = selEndDate.ToString("d-MMM")
                        End If
                        rActivity = rBase
                        DrawActivity(rActivity, strStart, strEnd)
                    End If
                Next
            End If
        End With
    End Sub

    Private Sub PrintKeyMoment(ByVal selRow As PrintPlanningRow)
        Dim strStart As String = String.Empty

        If rKeyMoment.X <= rCurrentSection.Right And rKeyMoment.X > rCurrentSection.Left Then
            If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                strStart = selRow.StartDate.ToString("d-MMM")
            End If
            ActivityColor = Color.DarkRed
            DrawKeyMoment(strStart)
            DrawKeyMomentLinks(selRow)
        End If
    End Sub
#End Region

#Region "Draw bars"
    Private Sub DrawActivity(ByVal rBar As Rectangle, ByVal strStart As String, ByVal strEnd As String)
        Dim rAvailableArea As Rectangle = GetAvailableArea()
        Dim intStartDateWidth As Integer, intEndDateWidth As Integer
        Dim intTextX As Integer, intTextY, intTextHeight As Integer
        Dim brText As Brush = Brushes.Black

        If rBar.Width = 0 Or rBar.Height = 0 Then Exit Sub

        rBar.X -= rCurrentSection.X
        If rBar.Left < 0 Then
            rBar.Width += rBar.X
            rBar.X = 0
            strStart = String.Empty
        End If
        rBar.X += GetCurrentPlanningColumnX()

        If rBar.Right > rAvailableArea.Right Then
            rBar.Width = rAvailableArea.Right - rBar.X
            strEnd = String.Empty
        End If

        If PageGraph IsNot Nothing Then
            Dim brBar As LinearGradientBrush = MakeGradientBrush(ActivityColor, rBar.X, rBar.Y, rBar.Height)
            PageGraph.FillRectangle(brBar, rBar)

            If String.IsNullOrEmpty(strStart) = False Then
                intStartDateWidth = PageGraph.MeasureString(strStart, fntText).Width
                intTextHeight = PageGraph.MeasureString(strStart, fntText).Height
                intTextX = rBar.X
                If intTextX + intStartDateWidth > GetCurrentPlanningColumnRight() Then intTextX = GetCurrentPlanningColumnRight() - intStartDateWidth

                intTextY = rBar.Y - intTextHeight
                If intTextY < rPlanningCell.Y Then intTextY = rPlanningCell.Y
                PageGraph.DrawString(strStart, fntText, brText, intTextX, intTextY)
            End If
            If String.IsNullOrEmpty(strEnd) = False Then
                intEndDateWidth = PageGraph.MeasureString(strEnd, fntText).Width
                intTextHeight = PageGraph.MeasureString(strEnd, fntText).Height
                intTextX = rBar.X + rBar.Width
                If rBar.Width > intStartDateWidth + intEndDateWidth Then intTextX -= intEndDateWidth
                If intTextX + intEndDateWidth > GetCurrentPlanningColumnRight() Then intTextX = GetCurrentPlanningColumnRight() - intEndDateWidth

                intTextY = rBar.Y + rBar.Height
                If intTextY + intTextHeight > rPlanningCell.Bottom Then intTextY = rPlanningCell.Bottom - intTextHeight
                PageGraph.DrawString(strEnd, fntText, brText, intTextX, intTextY)
            End If
        End If
    End Sub

    Private Sub DrawProcess(ByVal rActivity As Rectangle, ByVal intRowHeight As Integer, ByVal strStart As String, ByVal strEnd As String)
        Dim rProcess As New Rectangle(rActivity.X - 4, rActivity.Y, rActivity.Width + 8, rActivity.Height)
        Dim rAvailableArea As Rectangle = GetAvailableArea()
        Dim intStartDateWidth As Integer, intEndDateWidth As Integer
        Dim intTextX As Integer, intTextY, intTextHeight As Integer
        Dim brText As Brush = Brushes.Black
        Dim boolDrawStartArrow As Boolean = True, boolDrawEndArrow As Boolean = True
        If rProcess.Width = 0 Or rProcess.Height = 0 Then Exit Sub

        rProcess.X -= rCurrentSection.X
        If rProcess.Left < 0 Then
            rProcess.Width += rProcess.X
            rProcess.X = 0
            strStart = String.Empty
            boolDrawStartArrow = False
        End If
        rProcess.X += GetCurrentPlanningColumnX()

        If rProcess.Right > rAvailableArea.Right Then
            rProcess.Width = rAvailableArea.Right - rProcess.X
            strEnd = String.Empty
            boolDrawEndArrow = False
        End If

        If PageGraph IsNot Nothing Then
            PageGraph.FillRectangle(Brushes.Black, rProcess)

            Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line}
            Dim Points(2) As Point
            If boolDrawStartArrow = True Then
                Points(0) = New Point(rProcess.Left, rProcess.Bottom - 1)
                Points(1) = New Point(rProcess.Left + 4, rProcess.Bottom + 8)
                Points(2) = New Point(rProcess.Left + 8, rProcess.Bottom - 1)
                PageGraph.FillPath(Brushes.Black, New GraphicsPath(Points, PathPoints))
            End If

            If boolDrawEndArrow = True Then
                Points(0) = New Point(rProcess.Right, rProcess.Bottom - 1)
                Points(1) = New Point(rProcess.Right - 4, rProcess.Bottom + 8)
                Points(2) = New Point(rProcess.Right - 8, rProcess.Bottom - 1)
                PageGraph.FillPath(Brushes.Black, New GraphicsPath(Points, PathPoints))
            End If

            If String.IsNullOrEmpty(strStart) = False And boolDrawStartArrow = True Then
                intStartDateWidth = PageGraph.MeasureString(strStart, fntText).Width
                intTextX = rProcess.X
                intTextY = LastRowY
                PageGraph.DrawString(strStart, fntText, brText, intTextX, intTextY)
            End If
            If String.IsNullOrEmpty(strEnd) = False And boolDrawEndArrow = True Then
                intEndDateWidth = PageGraph.MeasureString(strEnd, fntText).Width
                intTextHeight = PageGraph.MeasureString(strEnd, fntText).Height
                intTextX = rProcess.X + rProcess.Width
                If rProcess.Width > intStartDateWidth + intEndDateWidth Then intTextX -= intEndDateWidth
                intTextY = LastRowY + intRowHeight - intTextHeight

                PageGraph.DrawString(strEnd, fntText, brText, intTextX, intTextY)
            End If
        End If
    End Sub

    Private Sub DrawKeyMoment(ByVal strStart As String)
        rKeyMoment.X -= rCurrentSection.X
        If rKeyMoment.Left >= 0 Then
            rKeyMoment.X += GetCurrentPlanningColumnX()

            Dim brKeyMoment As LinearGradientBrush = MakeGradientBrush(ActivityColor, rKeyMoment.Left, rKeyMoment.Top, rKeyMoment.Height)
            Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line, PathPointType.Line}
            Dim Points(3) As Point
            Dim fntText As New Font(CurrentLogFrame.DetailsFont.FontFamily, CurrentLogFrame.DetailsFont.SizeInPoints - 2)
            Dim brText As Brush = Brushes.Black

            If PageGraph IsNot Nothing Then
                Points(0) = New Point(rKeyMoment.Left, rKeyMoment.Top)
                Points(1) = New Point(rKeyMoment.Left - (rKeyMoment.Height / 2), rKeyMoment.Top + (rKeyMoment.Height / 2))
                Points(2) = New Point(rKeyMoment.Left, rKeyMoment.Bottom)
                Points(3) = New Point(rKeyMoment.Left + (rKeyMoment.Height / 2), rKeyMoment.Top + (rKeyMoment.Height / 2))
                PageGraph.FillPath(brKeyMoment, New GraphicsPath(Points, PathPoints))

                Dim intTextHeight = PageGraph.MeasureString(strStart, fntText).Height
                Dim intTextX As Integer = rKeyMoment.Left + 2
                Dim intTextY As Integer = rKeyMoment.Top + (rKeyMoment.Height / 2) - intTextHeight
                If intTextY < rPlanningCell.Y Then intTextY = rPlanningCell.Y

                If String.IsNullOrEmpty(strStart) = False Then
                    Dim intStartDateWidth As Integer = PageGraph.MeasureString(strStart, fntText).Width
                    PageGraph.DrawString(strStart, fntText, brText, intTextX, intTextY) '_
                End If
            End If
        End If
    End Sub
#End Region

#Region "Draw Key Moment links"
    Private Sub DrawKeyMomentLinks(ByVal selRow As PrintPlanningRow)
        If selRow.KeyMoment Is Nothing Then Exit Sub

        If selRow.OutgoingLinksIndices.Count > 0 Then
            DrawKeyMomentLinks_Outgoing(selRow)
        End If

        If selRow.IncomingLinkIndices.Count > 0 Then
            DrawKeyMomentLinks_Incoming(selRow)
        End If
    End Sub

    Private Sub DrawKeyMomentLinks_Outgoing(ByVal selRow As PrintPlanningRow)
        Dim selKeyMoment As KeyMoment = selRow.KeyMoment
        Dim intIndex As Integer = PrintList.IndexOf(selRow)
        Dim intLinkX As Integer = GetCoordinateX(selKeyMoment.ExactDateKeyMoment) + (CONST_BarHeight / 2)
        Dim intLinkY As Integer = LastRowY + (selRow.RowHeight / 2)
        Dim intTargetX As Integer

        intLinkX = CoordinateToCurrentSection(intLinkX)

        Dim LinkStart As Point = New Point(intLinkX, intLinkY)

        For i = 0 To selRow.OutgoingLinksIndices.Count - 1
            Dim intTargetIndex As Integer = selRow.OutgoingLinksIndices(i)
            Dim TargetRow As PrintPlanningRow = Me.PrintList(intTargetIndex)

            If TargetRow.RowType = PrintPlanningRow.RowTypes.KeyMoment Then
                Dim selTargetKeyMoment As KeyMoment = TargetRow.KeyMoment

                intTargetX = GetCoordinateX(selTargetKeyMoment.ExactDateKeyMoment) 'use the current cell to calculate since that is certainly visible


            ElseIf TargetRow.RowType = PrintPlanningRow.RowTypes.Activity Then
                Dim selTargetActivity As Activity = TryCast(TargetRow.Struct, Activity)

                intTargetX = GetCoordinateX(selTargetActivity.ExactStartDate)
            Else
                Exit Sub
            End If

            intTargetX = CoordinateToCurrentSection(intTargetX)

            If Me.ShowKeyMomentLinks = True Then DrawKeyMomentLinks_Outgoing_Draw(selRow, intIndex, intTargetIndex, intTargetX, LinkStart)
        Next
    End Sub

    Private Sub DrawKeyMomentLinks_Outgoing_Draw(ByVal selRow As PrintPlanningRow, ByVal intIndex As Integer, ByVal intTargetIndex As Integer, ByVal intTargetX As Integer, _
                                                 ByVal LinkStart As Point)
        Dim FirstElbow, SecondElbow, ThirdElbow, EndPoint As New Point

        If intTargetX >= LinkStart.X + CONST_DistanceToKeyMoment Then
            'if the target key moment/activity is not directly below the key moment, draw a simple line with one right angle
            FirstElbow = New Point(intTargetX, LinkStart.Y)

            If intTargetIndex > intIndex Then
                EndPoint = New Point(FirstElbow.X, LastRowY + selRow.RowHeight)
            Else
                EndPoint = New Point(FirstElbow.X, LastRowY)
            End If
            PageGraph.DrawLine(penLightBlue2, LinkStart, FirstElbow)
            PageGraph.DrawLine(penLightBlue2, FirstElbow, EndPoint)
            PageGraph.DrawLine(penDarkBlue1, LinkStart, FirstElbow)
            PageGraph.DrawLine(penDarkBlue1, FirstElbow, EndPoint)
        Else
            'if the target key moment/activity starts directly below the key moment, draw a half loop
            FirstElbow = New Point(LinkStart.X + CONST_DistanceToKeyMoment, LinkStart.Y)
            If intTargetIndex > intIndex Then
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y + 10)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, LastRowY + selRow.RowHeight)
            Else
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y - 10)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, LastRowY)
            End If

            PageGraph.DrawLine(penLightBlue2, LinkStart, FirstElbow)
            PageGraph.DrawLine(penLightBlue2, FirstElbow, SecondElbow)
            PageGraph.DrawLine(penLightBlue2, SecondElbow, ThirdElbow)
            PageGraph.DrawLine(penLightBlue2, ThirdElbow, EndPoint)

            PageGraph.DrawLine(penDarkBlue1, LinkStart, FirstElbow)
            PageGraph.DrawLine(penDarkBlue1, FirstElbow, SecondElbow)
            PageGraph.DrawLine(penDarkBlue1, SecondElbow, ThirdElbow)
            PageGraph.DrawLine(penDarkBlue1, ThirdElbow, EndPoint)
        End If

        DrawKeyMomentLinks_VerticalLinks(intIndex, intTargetIndex, EndPoint)
    End Sub

    Private Sub DrawKeyMomentLinks_Incoming(ByVal selRow As PrintPlanningRow)
        Dim selKeyMoment As KeyMoment = selRow.KeyMoment
        Dim intLinkX As Integer
        Dim intLinkY As Integer = GetCoordinateY(selRow.RowHeight)
        Dim intIndex As Integer = PrintList.IndexOf(selRow)

        If intLinkX > rPlanningColumnRectangle.Left Then
            For i = 0 To selRow.IncomingLinkIndices.Count - 1
                Dim intSourceIndex As Integer = selRow.IncomingLinkIndices(i)
                Dim SourceRow As PrintPlanningRow = Me.PrintList(intSourceIndex)
                Dim selSourceKeyMoment As KeyMoment = SourceRow.KeyMoment
                Dim selSourceActivity As Activity = TryCast(SourceRow.Struct, Activity)

                If SourceRow.RowType = PrintPlanningRow.RowTypes.KeyMoment Then
                    intLinkX = GetCoordinateX(selKeyMoment.ExactDateKeyMoment)
                ElseIf SourceRow.RowType = PrintPlanningRow.RowTypes.Activity Then
                    intLinkX = GetCoordinateX(selSourceActivity.ExactStartDate)
                Else
                    Exit Sub
                End If

                intLinkX = CoordinateToCurrentSection(intLinkX)

                Dim LinkEnd As New Point(intLinkX, intLinkY)

                If Me.ShowKeyMomentLinks = True Then DrawKeyMomentLinks_Incoming_Draw(selRow, intIndex, intSourceIndex, LinkEnd)
            Next
        End If
    End Sub

    Private Sub DrawKeyMomentLinks_Incoming_Draw(ByVal selRow As PrintPlanningRow, ByVal intIndex As Integer, ByVal intSourceIndex As Integer, ByVal LinkEnd As Point)
        Dim EntryPoint As New Point
        Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line}
        Dim PointsLight(2) As Point
        Dim PointsDark(2) As Point

        If intSourceIndex < intIndex Then
            EntryPoint = New Point(LinkEnd.X, LastRowY)
            PointsLight(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsLight(1) = New Point(LinkEnd.X - 4, LinkEnd.Y - 8)
            PointsLight(2) = New Point(LinkEnd.X, LinkEnd.Y - 8)
            PointsDark(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsDark(1) = New Point(LinkEnd.X + 4, LinkEnd.Y - 8)
            PointsDark(2) = New Point(LinkEnd.X, LinkEnd.Y - 8)
        Else

            'key moments are ordered by date, so only an arrow pointing downwards is needed
            LinkEnd.Y += CONST_BarHeight
            LinkEnd = New Point(LinkEnd.X, LinkEnd.Y)
            EntryPoint = New Point(LinkEnd.X, LastRowY = selRow.RowHeight)
            PointsLight(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsLight(1) = New Point(LinkEnd.X + 4, LinkEnd.Y + 8)
            PointsLight(2) = New Point(LinkEnd.X, LinkEnd.Y + 8)
            PointsDark(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsDark(1) = New Point(LinkEnd.X - 4, LinkEnd.Y + 8)
            PointsDark(2) = New Point(LinkEnd.X, LinkEnd.Y + 8)
        End If

        PageGraph.DrawLine(penLightBlue2, EntryPoint, LinkEnd)
        PageGraph.DrawLine(penDarkBlue1, EntryPoint, LinkEnd)
        PageGraph.FillPath(Brushes.DeepSkyBlue, New GraphicsPath(PointsLight, PathPoints))
        PageGraph.FillPath(Brushes.RoyalBlue, New GraphicsPath(PointsDark, PathPoints))
    End Sub

    Private Sub DrawKeyMomentLinks_VerticalLinks(ByVal intIndex As Integer, ByVal intTargetIndex As Integer, ByVal EndPoint As Point)
        'set the co-ordinates of the vertical lines that need to be drawn in the rows between the source and target rows
        Dim selDate As Date = CoordinateToDate(EndPoint.X)

        If intTargetIndex > intIndex Then
            For j = intIndex + 1 To intTargetIndex - 1
                PrintList(j).KeyMomentLinksAsDate.Add(selDate)
            Next
        Else
            For j = intTargetIndex + 1 To intIndex - 1
                PrintList(j).KeyMomentLinksAsDate.Add(selDate)
            Next
        End If
    End Sub

    Private Sub DrawKeyMomentLinks_VerticalKeyMomentLinks(ByVal selRow As PrintPlanningRow)
        Dim intLinkX As Integer

        If selRow.KeyMomentLinksAsDate.Count > 0 And Me.ShowKeyMomentLinks = True Then
            For i = 0 To selRow.KeyMomentLinksAsDate.Count - 1
                intLinkX = GetCoordinateX(selRow.KeyMomentLinksAsDate(i))
                intLinkX = CoordinateToCurrentSection(intLinkX)

                If intLinkX > rPlanningColumnRectangle.Left Then
                    PageGraph.DrawLine(penLightBlue2, intLinkX, LastRowY, intLinkX, LastRowY + selRow.RowHeight)
                    PageGraph.DrawLine(penDarkBlue1, intLinkX, LastRowY, intLinkX, LastRowY + selRow.RowHeight)
                End If

            Next
        End If
    End Sub
#End Region

#Region "Draw Activity links"
    Private Sub DrawActivityLinks(ByVal selRow As PrintPlanningRow)
        If selRow.OutgoingLinksIndices.Count > 0 Then
            DrawActivityLinks_Outgoing(selRow)
        End If

        If selRow.IncomingLinkIndices.Count > 0 Then
            DrawActivityLinks_Incoming(selRow)
        End If
    End Sub

    Private Sub DrawActivityLinks_Outgoing(ByVal selRow As PrintPlanningRow)
        Dim selActivity As Activity = DirectCast(selRow.Struct, Activity)
        Dim intIndex As Integer = PrintList.IndexOf(selRow)
        Dim intLinkX As Integer = GetCoordinateX(selActivity.ExactEndDate)
        Dim intLinkY As Integer = GetCoordinateY(selRow.RowHeight)
        Dim intTargetX As Integer
        Dim LinkStart As New Point(intLinkX, intLinkY)

        LinkStart.X = CoordinateToCurrentSection(intLinkX)

        For i = 0 To selRow.OutgoingLinksIndices.Count - 1
            Dim intTargetIndex As Integer = selRow.OutgoingLinksIndices(i)
            Dim TargetRow As PrintPlanningRow = Me.PrintList(intTargetIndex)

            If intTargetIndex > intIndex Then
                LinkStart.Y = intLinkY + CONST_BarHeight - 2
            Else
                LinkStart.Y = intLinkY + 2
            End If

            If TargetRow.RowType = PrintPlanningRow.RowTypes.Activity Then
                Dim selTargetActivity As Activity = TryCast(TargetRow.Struct, Activity)

                intTargetX = GetCoordinateX(selTargetActivity.ExactStartDate) 'use the current cell to calculate since that is certainly visible
            ElseIf TargetRow.RowType = PrintPlanningRow.RowTypes.KeyMoment Then
                Dim selTargetKeyMoment As KeyMoment = TargetRow.KeyMoment

                intTargetX = GetCoordinateX(selTargetKeyMoment.ExactDateKeyMoment) 'use the current cell to calculate since that is certainly visible
            Else
                Exit Sub
            End If
            intTargetX = CoordinateToCurrentSection(intTargetX)

            If Me.ShowActivityLinks = True Then DrawActivityLinks_Outgoing_Draw(selRow, intIndex, intTargetIndex, intTargetX, LinkStart)
        Next
    End Sub

    Private Sub DrawActivityLinks_Outgoing_Draw(ByVal selRow As PrintPlanningRow, ByVal intIndex As Integer, ByVal intTargetIndex As Integer, _
                                                ByVal intTargetX As Integer, ByVal LinkStart As Point)
        Dim selActivity As Activity = DirectCast(selRow.Struct, Activity)
        Dim FirstElbow, SecondElbow, ThirdElbow, EndPoint As New Point

        If intTargetX >= LinkStart.X + CONST_DistanceToActivity Then
            'if the target key moment/activity is not directly below or above the activity
            If selActivity.ActivityDetail.PeriodDirection = ActivityDetail.DirectionValues.Before Then
                'target activity is above source activity
                'draw small line up, then horizontal line to begin of target activity, then up to border
                If intTargetIndex > intIndex Then
                    FirstElbow = New Point(LinkStart.X, LinkStart.Y + 6)
                    SecondElbow = New Point(intTargetX, FirstElbow.Y)
                    EndPoint = New Point(SecondElbow.X, LastRowY + selRow.RowHeight)
                Else
                    FirstElbow = New Point(LinkStart.X, LinkStart.Y - 6)
                    SecondElbow = New Point(intTargetX, FirstElbow.Y)
                    EndPoint = New Point(SecondElbow.X, LastRowY)
                End If

                PageGraph.DrawLine(penBlack2, LinkStart, FirstElbow)
                PageGraph.DrawLine(penBlack2, FirstElbow, SecondElbow)
                PageGraph.DrawLine(penBlack2, SecondElbow, EndPoint)
            Else
                'if the target key moment/activity is not directly below the key moment, draw a simple line with one right angle
                FirstElbow = New Point(intTargetX, LinkStart.Y)
                If intTargetIndex > intIndex Then
                    EndPoint = New Point(FirstElbow.X, LastRowY + selRow.RowHeight)
                Else
                    EndPoint = New Point(FirstElbow.X, LastRowY)
                End If

                PageGraph.DrawLine(penBlack2, LinkStart, FirstElbow)
                PageGraph.DrawLine(penBlack2, FirstElbow, EndPoint)
            End If

        Else
            'target activity is directly above/below source activity
            'draw half loop
            FirstElbow = New Point(LinkStart.X + CONST_DistanceToActivity, LinkStart.Y)
            If intTargetIndex > intIndex Then
                'loop below
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y + 6)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, LastRowY + selRow.RowHeight)
            Else
                'loop above
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y - 6)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, LastRowY)
            End If

            PageGraph.DrawLine(penBlack2, LinkStart, FirstElbow)
            PageGraph.DrawLine(penBlack2, FirstElbow, SecondElbow)
            PageGraph.DrawLine(penBlack2, SecondElbow, ThirdElbow)
            PageGraph.DrawLine(penBlack2, ThirdElbow, EndPoint)
        End If

        DrawActivityLinks_VerticalLinks(intIndex, intTargetIndex, EndPoint)
    End Sub

    Private Sub DrawActivityLinks_Incoming(ByVal selRow As PrintPlanningRow)
        Dim selActivity As Activity = DirectCast(selRow.Struct, Activity)
        Dim intIndex As Integer = PrintList.IndexOf(selRow)
        Dim intLinkX As Integer = GetCoordinateX(selActivity.ExactStartDate)
        Dim intLinkY As Integer = GetCoordinateY(selRow.RowHeight)
        Dim EntryPoint As New Point
        Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line}

        intLinkX = CoordinateToCurrentSection(intLinkX)

        If intLinkX > rPlanningColumnRectangle.Left Then
            Dim LinkEnd As New Point(intLinkX, intLinkY)

            For i = 0 To selRow.IncomingLinkIndices.Count - 1
                Dim intSourceIndex As Integer = selRow.IncomingLinkIndices(i)
                Dim SourceRow As PrintPlanningRow = Me.PrintList(intSourceIndex)

                If SourceRow.RowType = PrintPlanningRow.RowTypes.Activity And Me.ShowActivityLinks = True Then
                    Dim Points(2) As Point
                    Dim selSourceActivity As Activity = TryCast(SourceRow.Struct, Activity)
                    If selSourceActivity Is Nothing Then Exit Sub

                    If intSourceIndex < intIndex Then
                        EntryPoint = New Point(intLinkX, LastRowY)
                        Points(0) = New Point(intLinkX, intLinkY)
                        Points(1) = New Point(intLinkX + 4, intLinkY - 8)
                        Points(2) = New Point(intLinkX - 4, intLinkY - 8)
                    Else
                        intLinkY += CONST_BarHeight
                        LinkEnd = New Point(intLinkX, intLinkY)
                        EntryPoint = New Point(intLinkX, LastRowY + selRow.RowHeight)
                        Points(0) = New Point(intLinkX, intLinkY)
                        Points(1) = New Point(intLinkX + 4, intLinkY + 8)
                        Points(2) = New Point(intLinkX - 4, intLinkY + 8)
                    End If

                    PageGraph.DrawLine(penBlack2, EntryPoint, LinkEnd)
                    PageGraph.FillPath(Brushes.Black, New GraphicsPath(Points, PathPoints))
                ElseIf SourceRow.RowType = PrintPlanningRow.RowTypes.KeyMoment And Me.ShowKeyMomentLinks = True Then
                    Dim PointsLight(2) As Point
                    Dim PointsDark(2) As Point

                    If intSourceIndex < intIndex Then
                        EntryPoint = New Point(intLinkX, LastRowY)
                        PointsLight(0) = New Point(intLinkX, intLinkY)
                        PointsLight(1) = New Point(intLinkX - 4, intLinkY - 8)
                        PointsLight(2) = New Point(intLinkX, intLinkY - 8)
                        PointsDark(0) = New Point(intLinkX, intLinkY)
                        PointsDark(1) = New Point(intLinkX + 4, intLinkY - 8)
                        PointsDark(2) = New Point(intLinkX, intLinkY - 8)
                    Else
                        intLinkY += CONST_BarHeight
                        LinkEnd = New Point(intLinkX, intLinkY)
                        EntryPoint = New Point(intLinkX, LastRowY + selRow.RowHeight)
                        PointsLight(0) = New Point(intLinkX, intLinkY)
                        PointsLight(1) = New Point(intLinkX + 4, intLinkY + 8)
                        PointsLight(2) = New Point(intLinkX, intLinkY + 8)
                        PointsDark(0) = New Point(intLinkX, intLinkY)
                        PointsDark(1) = New Point(intLinkX - 4, intLinkY + 8)
                        PointsDark(2) = New Point(intLinkX, intLinkY + 8)
                    End If

                    PageGraph.DrawLine(penLightBlue2, EntryPoint, LinkEnd)
                    PageGraph.DrawLine(penDarkBlue1, EntryPoint, LinkEnd)
                    PageGraph.FillPath(Brushes.DeepSkyBlue, New GraphicsPath(PointsLight, PathPoints))
                    PageGraph.FillPath(Brushes.RoyalBlue, New GraphicsPath(PointsDark, PathPoints))
                End If
            Next
        End If
    End Sub

    Private Sub DrawActivityLinks_VerticalLinks(ByVal intIndex As Integer, ByVal intTargetIndex As Integer, ByVal EndPoint As Point)
        'set the co-ordinates of the vertical lines that need to be drawn in the rows between the source and target rows
        Dim selDate As Date = CoordinateToDate(EndPoint.X)

        If intTargetIndex > intIndex Then
            For j = intIndex + 1 To intTargetIndex - 1
                PrintList(j).ActivityLinksAsDate.Add(selDate)
            Next
        Else
            For j = intTargetIndex + 1 To intIndex - 1
                PrintList(j).ActivityLinksAsDate.Add(selDate)
            Next
        End If
    End Sub

    Private Sub DrawActivityLinks_VerticalActivityLinks(ByVal selRow As PrintPlanningRow)
        Dim intLinkX As Integer

        If selRow.ActivityLinksAsDate.Count > 0 And Me.ShowActivityLinks = True Then
            For i = 0 To selRow.ActivityLinksAsDate.Count - 1
                intLinkX = GetCoordinateX(selRow.ActivityLinksAsDate(i))
                intLinkX = CoordinateToCurrentSection(intLinkX)

                If intLinkX > rPlanningColumnRectangle.Left Then
                    PageGraph.DrawLine(penBlack2, intLinkX, LastRowY, intLinkX, LastRowY + selRow.RowHeight)
                End If
            Next
        End If
    End Sub
#End Region

#Region "Column headers"
    Private Sub PrintColumnHeaders()

        'Dim intPlanningColumnWidth As Integer
        Dim strText As String = String.Empty
        Dim rText As Rectangle

        If PageGraph IsNot Nothing Then
            'first header row
            'text and dates column header

            If CurrentHorPage = 1 Or RepeatRowHeaders = True Then
                'only print the label on the first page when HideTextColumn is true, otherwise print it on all the pages
                rText = New Rectangle(rHeadingFirstRow.X, rHeadingFirstRow.Y, SortColumnRectangle.Width + TextColumnRectangle.Width, ColumnHeadersHeight)

                Select Case Me.ElementsView
                    Case ShowElements.ShowActivities
                        strText = Activity.ItemName
                    Case ShowElements.ShowKeyMoments
                        strText = KeyMoment.ItemName
                    Case ShowElements.ShowBoth
                        strText = String.Format("{0}/{1}", KeyMoment.ItemName, Activity.ItemName)
                End Select
                PrintPage_PrintText(strText, rText, True)

                If ShowDatesColumns = True Then
                    rText = New Rectangle(StartDateColumnRectangle.X, rHeadingFirstRow.Y, StartDateColumnRectangle.Width, ColumnHeadersHeight)
                    PrintPage_PrintText(LANG_StartDate, rText, True)
                    rText = New Rectangle(EndDateColumnRectangle.X, rHeadingFirstRow.Y, EndDateColumnRectangle.Width, ColumnHeadersHeight)
                    PrintPage_PrintText(LANG_EndDate, rText, True)
                End If
            End If

            'planning column header
            Dim rPlanning As New Rectangle(GetCurrentPlanningColumnX, rHeadingFirstRow.Y, rCurrentSection.Width, rHeadingFirstRow.Height)

            PrintPage_PrintText(LANG_Planning, rPlanning, True)

            'second header row
            PrintColumnHeaders_Dates(PageGraph)
        End If

        LastRowY += ColumnHeadersHeight
    End Sub

    Private Sub PrintColumnHeaders_Dates(ByVal graph As Graphics)
        Dim rAvailableArea As Rectangle = GetAvailableArea()
        Dim intX As Integer
        Dim intCurrentPlanningColumnX As Integer = GetCurrentPlanningColumnX()
        Dim intHalfOfSecondRowHeight As Integer = rHeadingSecondRow.Height / 2
        Dim intHalfOfSecondRowY As Integer = rHeadingSecondRow.Y + intHalfOfSecondRowHeight
        Dim strYear As String, strMonth As String, strWeek As String
        Dim intNrOfDays As Integer
        Dim intTextWidth, intMonthX As Integer
        Dim brWeekEnd As New SolidBrush(Color.DarkSlateGray)

        Dim fntHeader As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
        Dim fntYear As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
        Dim fntMonth As New Font(CurrentLogFrame.DetailsFont.FontFamily, CurrentLogFrame.DetailsFont.SizeInPoints - 2)
        Dim fntDetail As New Font(CurrentLogFrame.DetailsFont.FontFamily, CurrentLogFrame.DetailsFont.SizeInPoints - 2)

        For intYear = Me.PeriodFrom.Year To Me.PeriodUntil.Year
            Dim NewYear As Date = New Date(intYear, 1, 1, 0, 0, 0)
            intX = GetCoordinateX(NewYear)
            If intX >= rCurrentSection.Left And intX <= rCurrentSection.Right Then
                intX -= rCurrentSection.X
                intX += intCurrentPlanningColumnX
                strYear = NewYear.ToString("yyyy")
                intTextWidth = graph.MeasureString(strYear, fntYear).Width
                graph.DrawLine(penBlack1, intX, rHeadingSecondRow.Top, intX, rHeadingSecondRow.Bottom)

                If intX + intTextWidth <= intCurrentPlanningColumnX + rCurrentSection.Width Then
                    graph.DrawString(strYear, fntYear, Brushes.Black, intX, rHeadingSecondRow.Y)
                Else
                    graph.DrawString(strYear, fntYear, Brushes.Black, intX - intTextWidth, rHeadingSecondRow.Y)
                End If
            End If

            'mark months
            For intMonth = 1 To 12
                Dim NewMonth As Date = New Date(intYear, intMonth, 1, 0, 0, 0)
                intX = GetCoordinateX(NewMonth)
                intNrOfDays = Date.DaysInMonth(intYear, intMonth) * WidthOfDay

                If Me.PeriodView = DataGridViewPlanning.PeriodViews.Month Then
                    If intX >= rCurrentSection.Left And intX + intNrOfDays <= rCurrentSection.Right Then
                        intX -= rCurrentSection.X
                        intX += intCurrentPlanningColumnX
                        strMonth = NewMonth.ToString("MMM")
                        graph.DrawLine(penBlack1, intX, intHalfOfSecondRowY, intX, rHeadingSecondRow.Bottom)
                        graph.DrawString(strMonth, fntMonth, Brushes.Black, intX, rHeadingSecondRow.Bottom - fntMonth.Height)
                    End If
                ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                    If (intX >= rCurrentSection.Left And intX <= rCurrentSection.Right) Or (intX + intNrOfDays >= rCurrentSection.Left And intX + intNrOfDays <= rCurrentSection.Right) Then
                        intX -= rCurrentSection.X
                        intX += intCurrentPlanningColumnX

                        If intX < intCurrentPlanningColumnX Then
                            intNrOfDays = intNrOfDays - (intCurrentPlanningColumnX - intX)
                            intX = intCurrentPlanningColumnX
                        End If
                        If intX + intNrOfDays > rAvailableArea.Right Then
                            intNrOfDays = rAvailableArea.Right - intX
                        End If

                        If intMonth / 2 = Int(intMonth / 2) Then
                            graph.FillRectangle(SystemBrushes.ControlDark, New Rectangle(intX, rHeadingSecondRow.Y, intNrOfDays, rHeadingSecondRow.Height))
                        End If
                        strMonth = NewMonth.ToString("MMMM \'yy")

                        fntMonth = New Font(fntMonth, FontStyle.Bold)
                        intTextWidth = graph.MeasureString(strMonth, fntMonth).Width
                        If intTextWidth <= intNrOfDays Then
                            intMonthX = intX + Int((intNrOfDays / 2) - (intTextWidth / 2))
                            If intMonthX + intTextWidth < intCurrentPlanningColumnX + rCurrentSection.Width Then _
                                graph.DrawString(strMonth, fntMonth, Brushes.Black, intMonthX, rHeadingSecondRow.Y + 2)
                        End If
                    End If
                End If
            Next

            If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Then
                'print week numbers
                Dim intDayOfWeek As Integer = DatePart(DateInterval.Weekday, NewYear, FirstDayOfWeek.Monday, FirstWeekOfYear.System) - 1
                Dim FirstMonday As Date = New Date(intYear, 1, 8 - intDayOfWeek)
                Dim NewWeek As Date = FirstMonday
                Dim intNrWeek As Integer = 0

                Do While NewWeek.Year < intYear + 1
                    intX = GetCoordinateX(NewWeek)
                    intNrWeek += 1
                    If intX >= rCurrentSection.Left And intX + (7 * WidthOfDay) <= rCurrentSection.Right Then
                        intX -= rCurrentSection.X
                        intX += intCurrentPlanningColumnX
                        graph.DrawLine(penBlack1, intX, intHalfOfSecondRowY, intX, rHeadingSecondRow.Bottom)
                        strWeek = intNrWeek.ToString()
                        intTextWidth = graph.MeasureString(strWeek, fntDetail).Width
                        If intX + intTextWidth <= intCurrentPlanningColumnX + rCurrentSection.Width Then
                            graph.DrawString(strWeek, fntDetail, Brushes.Black, intX, intHalfOfSecondRowY)
                        End If
                    End If
                    NewWeek = NewWeek.AddDays(7)
                Loop
            ElseIf Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                'print days of week
                Dim NewDay As Date = NewYear
                Do While NewDay.Year < intYear + 1
                    Dim tsSinceStart As New TimeSpan
                    tsSinceStart = NewDay.Subtract(Me.PeriodFrom)
                    intX = (tsSinceStart.Days * WidthOfDay)

                    If intX >= rCurrentSection.Left And intX + WidthOfDay <= rCurrentSection.Right Then
                        intX -= rCurrentSection.X
                        intX += intCurrentPlanningColumnX
                        If NewDay.DayOfWeek = DayOfWeek.Saturday Or NewDay.DayOfWeek = DayOfWeek.Sunday Then
                            Dim rWeekend As New Rectangle(intX, intHalfOfSecondRowY, WidthOfDay, intHalfOfSecondRowHeight)
                            graph.FillRectangle(brWeekEnd, rWeekend)
                        End If
                        graph.DrawLine(penBlack1, intX, intHalfOfSecondRowY, intX, rHeadingSecondRow.Bottom)
                        graph.DrawString(NewDay.Day.ToString, fntDetail, Brushes.Black, intX, intHalfOfSecondRowY)
                    End If
                    NewDay = NewDay.AddDays(1)
                Loop
            End If
        Next
        Dim rDates As New Rectangle(intCurrentPlanningColumnX, rHeadingSecondRow.Y, rCurrentSection.Width, rHeadingSecondRow.Height)
        graph.DrawRectangle(penBlack1, rDates)
    End Sub
#End Region
End Class


