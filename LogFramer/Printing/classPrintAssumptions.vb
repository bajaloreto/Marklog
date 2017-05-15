Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintAssumptions
    Inherits ReportBaseClass

    Private Const CONST_MinTextColumnWidth As Integer = 100
    Private Const CONST_ValidatedColumnWidth As Integer = 70

    Private objLogFrame As New LogFrame
    Private objPrintList As New PrintAssumptionRows
    Private objClippedRow As PrintAssumptionRow = Nothing
    Private intPrintSection As Integer
    Private intPagesWide As Integer

    Private strClippedTextTop, strClippedTextBottom As String

    Private boolColumnsWidthSet As Boolean
    Private CurrentSectionRectangle As New Rectangle
    Private rStructSortRectangle, rAssumptionSortRectangle As New PrintRectangle
    Private rStructRectangle, rAssumptionRectangle As New PrintRectangle
    Private rReasonRectangle, rHowToValidateRectangle, rValidatedRectangle, rImpactRectangle, rResponseStrategyRectangle, rOwnerRectangle As New PrintRectangle

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

    Public Sub New(ByVal logframe As LogFrame, ByVal printsection As Integer, ByVal pageswide As Integer)
        Me.LogFrame = logframe
        Me.ReportSetup = logframe.ReportSetupAssumptions
        Me.PrintSection = printsection
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

    Public Property PrintList() As PrintAssumptionRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintAssumptionRows)
            objPrintList = value
        End Set
    End Property

    Public Property ClippedRow As PrintAssumptionRow
        Get
            Return objClippedRow
        End Get
        Set(ByVal value As PrintAssumptionRow)
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

    Private Property ReasonRectangle As PrintRectangle
        Get
            Return rReasonRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rReasonRectangle = value
        End Set
    End Property

    Private Property HowToValidateRectangle As PrintRectangle
        Get
            Return rHowToValidateRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rHowToValidateRectangle = value
        End Set
    End Property

    Private Property ValidatedRectangle As PrintRectangle
        Get
            Return rValidatedRectangle
        End Get
        Set(ByVal value As PrintRectangle)
            rValidatedRectangle = value
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

#Region "Create Assumptions table"
    Public Sub CreateList()
        Dim strGoalSort As String = String.Empty, strPurposeSort As String = String.Empty
        Dim strOutputSort As String = String.Empty

        PrintList.Clear()

        'add goals
        If Me.PrintSection = PrintSections.PrintGoals Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selGoal As Goal In Me.LogFrame.Goals
                strGoalSort = LogFrame.CreateSortNumber(LogFrame.Goals.IndexOf(selGoal))
                Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.GoalsSection, strGoalSort, New Goal(selGoal.RTF))
                PrintList.Add(NewRow)

                CreateList_Assumptions(LogFrame.SectionTypes.GoalsSection, selGoal.Assumptions, strGoalSort)
            Next
        End If

        'add purposes
        If Me.PrintSection = PrintSections.PrintPurposes Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))

                If Me.PrintSection = PrintSections.PrintPurposes Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
                    PrintList.Add(NewRow)

                    CreateList_Assumptions(LogFrame.SectionTypes.PurposesSection, selPurpose.Assumptions, strPurposeSort)
                End If
            Next
        End If

        'add outputs
        If Me.PrintSection = PrintSections.PrintOutputs Or Me.PrintSection = PrintSections.PrintAll Then
            For Each selPurpose As Purpose In Me.LogFrame.Purposes
                strPurposeSort = LogFrame.CreateSortNumber(LogFrame.Purposes.IndexOf(selPurpose))

                If My.Settings.setRepeatPurposes = True Then
                    Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
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
                    Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.PurposesSection, strPurposeSort, New Purpose(selPurpose.RTF))
                    PrintList.Add(NewRow)
                End If

                For Each selOutput As Output In selPurpose.Outputs
                    strOutputSort = LogFrame.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)

                    If My.Settings.setRepeatOutputs = True Then
                        Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.OutputsSection, strOutputSort, New Output(selOutput.RTF))
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
                    Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.OutputsSection, strOutputSort, New Output(selOutput.RTF))
                    PrintList.Add(NewRow)

                    CreateList_Assumptions(LogFrame.SectionTypes.OutputsSection, selOutput.Assumptions, strOutputSort)
                End If

                'CreateList_Activities(selOutput.Activities, strOutputSort)
            Next
        End If
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String)
        Dim strActivitySort As String = String.Empty

        If Me.PrintSection > PrintSections.PrintOutputs Then
            For Each selActivity As Activity In selActivities
                strActivitySort = LogFrame.CreateSortNumber(selActivities.IndexOf(selActivity), strParentSort)

                If Me.PrintSection = PrintSections.PrintActivities Or Me.PrintSection = PrintSections.PrintAll Then
                    Dim NewRow As New PrintAssumptionRow(LogFrame.SectionTypes.ActivitiesSection, strActivitySort, New Activity(selActivity.RTF))
                    PrintList.Add(NewRow)

                    CreateList_Assumptions(LogFrame.SectionTypes.ActivitiesSection, selActivity.Assumptions, strActivitySort)

                    If selActivity.Activities.Count > 0 Then
                        CreateList_Activities(selActivity.Activities, strActivitySort)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub CreateList_Assumptions(ByVal intSection As Integer, ByVal selAssumptions As Assumptions, ByVal strParentSort As String)
        Dim strAssumptionSort As String
        Dim intIndex As Integer
        Dim objRow As PrintAssumptionRow

        If selAssumptions.Count > 0 Then
            For Each selAssumption As Assumption In selAssumptions
                If selAssumption.RaidType = Assumption.RaidTypes.Assumption Then
                    Dim objAssumption As Assumption
                    strAssumptionSort = LogFrame.CreateSortNumber(intIndex, strParentSort)

                    Using objCopier As New ObjectCopy
                        objAssumption = objCopier.CopyObject(selAssumption)
                    End Using

                    objRow = New PrintAssumptionRow(intSection, strAssumptionSort, objAssumption)
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

        SectionBorder = Me.ContentRight

        'calculate column widths
        intStructSortWidth = GetStructSortColumnWidth()
        intAssumptionSortWidth = GetAssumptionSortColumnWidth()

        If intStructSortWidth < intAssumptionSortWidth Then intStructSortWidth = intAssumptionSortWidth
        If intAssumptionSortWidth < intStructSortWidth Then intAssumptionSortWidth = intStructSortWidth

        intTextColumnWidth = intAvailableWidth - intAssumptionSortWidth
        intTextColumnWidth /= 6

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

        'set reason column rectangle
        ReasonRectangle = New PrintRectangle(AssumptionRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(ReasonRectangle)

        If ReasonRectangle.Left <= intSectionBorder And ReasonRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(ReasonRectangle)
        End If

        'set how-to-validate column rectangle
        HowToValidateRectangle = New PrintRectangle(ReasonRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(HowToValidateRectangle)

        If HowToValidateRectangle.Left <= intSectionBorder And HowToValidateRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(HowToValidateRectangle)
        End If

        'set validated column rectangle
        ValidatedRectangle = New PrintRectangle(HowToValidateRectangle.Right, ContentTop, CONST_ValidatedColumnWidth, ContentHeight, intHorizontalPageIndex)
        PrintRectangles.Add(ValidatedRectangle)

        If ValidatedRectangle.Left <= intSectionBorder And ValidatedRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(ValidatedRectangle)
        End If

        'set impact column rectangle
        ImpactRectangle = New PrintRectangle(ValidatedRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
        PrintRectangles.Add(ImpactRectangle)

        If ImpactRectangle.Left <= intSectionBorder And ImpactRectangle.Right > intSectionBorder Then
            StretchRectanglesOfPreviousPage(ImpactRectangle)
        End If

        'set response strategy column rectangle
        ResponseStrategyRectangle = New PrintRectangle(ImpactRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight, intHorizontalPageIndex, True)
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

        For Each selRow As PrintAssumptionRow In Me.PrintList
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

        For Each selRow As PrintAssumptionRow In Me.PrintList
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
        For Each selRow As PrintAssumptionRow In Me.PrintList
            ReloadImages_Normal(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintAssumptionRow)
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

                If String.IsNullOrEmpty(selRow.Reason) = False Then
                    selRow.ReasonImage = .TextWithPaddingToBitmap(ReasonRectangle.Width, selRow.Reason)
                End If
                If String.IsNullOrEmpty(selRow.HowToValidate) = False Then
                    selRow.HowToValidateImage = .TextWithPaddingToBitmap(HowToValidateRectangle.Width, selRow.HowToValidate)
                End If
                If String.IsNullOrEmpty(selRow.Impact) = False Then
                    selRow.ImpactImage = .TextWithPaddingToBitmap(ImpactRectangle.Width, selRow.Impact)
                End If
                If String.IsNullOrEmpty(selRow.ResponseStrategy) = False Then
                    selRow.ResponseStrategyImage = .TextWithPaddingToBitmap(ResponseStrategyRectangle.Width, selRow.ResponseStrategy)
                End If
                If String.IsNullOrEmpty(selRow.Owner) = False Then
                    selRow.OwnerImage = .TextWithPaddingToBitmap(OwnerRectangle.Width, selRow.Owner)
                End If
            End If
            
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintAssumptionRow = Me.PrintList(RowIndex)
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
        Dim selRow As PrintAssumptionRow = Me.PrintList(RowIndex)

        If selRow.Struct IsNot Nothing Then
            If selRow.Struct.CellImage.Height > intRowHeight Then intRowHeight = selRow.Struct.CellImage.Height
        Else
            If selRow.Assumption IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Assumption.RTF) = False Then
                If selRow.Assumption.CellImage.Height > intRowHeight Then intRowHeight = selRow.Assumption.CellImage.Height
            End If

            If String.IsNullOrEmpty(selRow.Reason) = False Then
                If selRow.ReasonImage.Height > intRowHeight Then intRowHeight = selRow.ReasonImage.Height
            End If

            If String.IsNullOrEmpty(selRow.HowToValidate) = False Then
                If selRow.HowToValidateImage.Height > intRowHeight Then intRowHeight = selRow.HowToValidateImage.Height
            End If

            If String.IsNullOrEmpty(selRow.Impact) = False Then
                If selRow.ImpactImage.Height > intRowHeight Then intRowHeight = selRow.ImpactImage.Height
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

            Dim intReasonHeight As Integer = PageGraph.MeasureString(LANG_Reason, fntHeader, ReasonRectangle.Width).Height
            If intReasonHeight > intHeight Then intHeight = intReasonHeight

            Dim intHowToValidateHeight As Integer = PageGraph.MeasureString(LANG_HowToValidate, fntHeader, HowToValidateRectangle.Width).Height
            If intHowToValidateHeight > intHeight Then intHeight = intHowToValidateHeight

            Dim intValidatedHeight As Integer = PageGraph.MeasureString(LANG_Validated, fntHeader, ValidatedRectangle.Width).Height
            If intValidatedHeight > intHeight Then intHeight = intValidatedHeight

            Dim intImpactHeight As Integer = PageGraph.MeasureString(LANG_Impact, fntHeader, ImpactRectangle.Width).Height
            If intImpactHeight > intHeight Then intHeight = intImpactHeight

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

        For Each selRow As PrintAssumptionRow In PrintList
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
        Dim selRow As PrintAssumptionRow = PrintList(RowIndex)
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
                    PrintPage_PrintReason(selRow)
                    PrintPage_PrintHowToValidate(selRow)
                    PrintPage_PrintValidated(selRow)
                    PrintPage_PrintImpact(selRow)
                    PrintPage_PrintResponseStrategy(selRow)
                    PrintPage_PrintOwner(selRow)
                End If

                PrintPage_PrintLines(selRow)

                If selRow.ClippedRow = True Then
                    ClippedRow = Nothing
                End If

                If LastRowY < ContentBottom And LastRowY + selRow.RowHeight > ContentBottom Then
                    If ClippedRow IsNot Nothing And CurrentHorPage = HorPages Then
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

    Private Sub PrintPage_PrintLines(ByVal selRow As PrintAssumptionRow)
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

            If ReasonRectangle.Left >= CurrentSectionRectangle.Left And ReasonRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(ReasonRectangle.X)
                intRight = intLeft + ReasonRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If HowToValidateRectangle.Left >= CurrentSectionRectangle.Left And HowToValidateRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(HowToValidateRectangle.X)
                intRight = intLeft + HowToValidateRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If ValidatedRectangle.Left >= CurrentSectionRectangle.Left And ValidatedRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(ValidatedRectangle.X)
                intRight = intLeft + ValidatedRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
            End If

            If ImpactRectangle.Left >= CurrentSectionRectangle.Left And ImpactRectangle.Right <= CurrentSectionRectangle.Right Then
                intLeft = GetCoordinateX(ImpactRectangle.X)
                intRight = intLeft + ImpactRectangle.Width

                PageGraph.DrawLine(penBlack1, intRight, intTop, intRight, intBottom + 1)
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
    Private Function PrintPage_GetMinHeight(ByVal selRow As PrintAssumptionRow)
        Dim intMinHeight As Integer

        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(StructRectangle.Width, selRow.Struct.RTF)
            ElseIf selRow.Assumption IsNot Nothing Then
                intMinHeight = .GetFirstLineSpacing(AssumptionRectangle.Width, selRow.Assumption.RTF)
            ElseIf String.IsNullOrEmpty(selRow.Reason) = False Then
                intMinHeight = .GetFirstLineSpacing(ReasonRectangle.Width, selRow.Reason)
            ElseIf String.IsNullOrEmpty(selRow.HowToValidate) = False Then
                intMinHeight = .GetFirstLineSpacing(HowToValidateRectangle.Width, selRow.HowToValidate)
            ElseIf String.IsNullOrEmpty(selRow.Impact) = False Then
                intMinHeight = .GetFirstLineSpacing(ImpactRectangle.Width, selRow.Impact)
            ElseIf String.IsNullOrEmpty(selRow.ResponseStrategy) = False Then
                intMinHeight = .GetFirstLineSpacing(ResponseStrategyRectangle.Width, selRow.ResponseStrategy)
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

            rImage.Height = bmClip.Height
            PageGraph.DrawImage(bmClip, rImage)
        End With
    End Sub

    Private Sub PrintClippedText(ByVal strText As String, ByVal rImage As Rectangle, ByVal intColumnWidth As Integer)
        With RichTextManager
            Dim bmClip As Bitmap = .PrintClippedText(intColumnWidth, strText, ContentBottom - rImage.Y)
            strClippedTextTop = .ClippedTextTop
            strClippedTextBottom = .ClippedTextBottom

            rImage.Height = bmClip.Height
            PageGraph.DrawImage(bmClip, rImage.X, rImage.Y)
        End With
    End Sub
#End Region

#Region "Print struct"
    Private Sub PrintPage_PrintStructSortNumber(ByVal selRow As PrintAssumptionRow)
        If StructSortRectangle.Left >= CurrentSectionRectangle.Left And StructSortRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim rSortNumber As New Rectangle(StructSortRectangle.X, LastRowY, StructSortRectangle.Width, selRow.RowHeight)
            PrintPage_PrintText(selRow.StructSortNumber, rSortNumber)
        End If
    End Sub

    Private Sub PrintPage_PrintStruct(ByVal selRow As PrintAssumptionRow)
        If StructRectangle.Left >= CurrentSectionRectangle.Left And StructRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(StructRectangle.X)

            If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.Struct.CellImage.Width, selRow.Struct.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Struct.CellImage, rImage)
                Else
                    PrintClippedRichText(selRow.Struct.RTF, rImage, StructRectangle.Width)
                    selRow.Struct.RTF = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintAssumptionRow()
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

#Region "Print assumptions"
    Private Sub PrintPage_PrintAssumptionSortNumber(ByVal selRow As PrintAssumptionRow)
        If String.IsNullOrEmpty(selRow.AssumptionSortNumber) = False Then
            If AssumptionSortRectangle.Left >= CurrentSectionRectangle.Left And AssumptionSortRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim intX As Integer = GetCoordinateX(AssumptionSortRectangle.X)
                Dim rSortNumber As New Rectangle(intX, LastRowY, AssumptionSortRectangle.Width, selRow.RowHeight)
                PrintPage_PrintText(selRow.AssumptionSortNumber, rSortNumber)
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintAssumption(ByVal selRow As PrintAssumptionRow)
        If AssumptionRectangle.Left >= CurrentSectionRectangle.Left And AssumptionRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(AssumptionRectangle.X)

            If selRow.Assumption IsNot Nothing AndAlso String.IsNullOrEmpty(selRow.Assumption.RTF) = False Then
                If selRow.Assumption.CellImage IsNot Nothing Then
                    Dim rImage As New Rectangle(intX, LastRowY, selRow.Assumption.CellImage.Width, selRow.Assumption.CellImage.Height)

                    If rImage.Bottom <= ContentBottom Then
                        PageGraph.DrawImage(selRow.Assumption.CellImage, rImage)
                    Else
                        PrintClippedRichText(selRow.Assumption.RTF, rImage, AssumptionRectangle.Width)
                        selRow.Assumption.RTF = strClippedTextTop

                        If ClippedRow Is Nothing Then
                            ClippedRow = New PrintAssumptionRow()
                            ClippedRow.ClippedRow = True
                        End If
                        ClippedRow.Assumption = New Assumption(strClippedTextBottom)
                    End If
                End If

                If selRow.ClippedRow = False Then
                    Dim intSortX As Integer = GetCoordinateX(AssumptionSortRectangle.X)
                    If intSortX < ContentLeft Then intSortX = ContentLeft
                    PageGraph.DrawLine(penBlack1, intSortX, LastRowY, LastColumnBorder, LastRowY)
                End If
            End If
        End If
    End Sub
#End Region

#Region "Assumption detail properties"
    Private Sub PrintPage_PrintReason(ByVal selRow As PrintAssumptionRow)
        If ReasonRectangle.Left >= CurrentSectionRectangle.Left And ReasonRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ReasonRectangle.X)

            If String.IsNullOrEmpty(selRow.Reason) = False AndAlso selRow.ReasonImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.ReasonImage.Width, selRow.ReasonImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ReasonImage, rImage)
                Else
                    PrintClippedText(selRow.Reason, rImage, ReasonRectangle.Width)
                    selRow.Reason = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintAssumptionRow()
                        ClippedRow.Assumption = New Assumption
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.Reason = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintHowToValidate(ByVal selRow As PrintAssumptionRow)
        If HowToValidateRectangle.Left >= CurrentSectionRectangle.Left And HowToValidateRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(HowToValidateRectangle.X)

            If String.IsNullOrEmpty(selRow.HowToValidate) = False AndAlso selRow.HowToValidateImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.HowToValidateImage.Width, selRow.HowToValidateImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.HowToValidateImage, rImage)
                Else
                    PrintClippedText(selRow.HowToValidate, rImage, HowToValidateRectangle.Width)
                    selRow.HowToValidate = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintAssumptionRow()
                        ClippedRow.Assumption = New Assumption
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.HowToValidate = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintValidated(ByVal selRow As PrintAssumptionRow)
        If selRow.ClippedRow = False Then
            If ValidatedRectangle.Left >= CurrentSectionRectangle.Left And ValidatedRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim intX As Integer = GetCoordinateX(ValidatedRectangle.X)
                Dim rValidated As New Rectangle(intX, LastRowY, ValidatedRectangle.Width, selRow.RowHeight)
                Dim strValidated As String

                If selRow.Validated = True Then
                    strValidated = LANG_Yes
                Else
                    strValidated = LANG_No
                End If

                PrintPage_PrintText(strValidated, rValidated)
            End If
        End If
    End Sub
#End Region

#Region "Assumption properties"
    Private Sub PrintPage_PrintImpact(ByVal selRow As PrintAssumptionRow)
        If ImpactRectangle.Left >= CurrentSectionRectangle.Left And ImpactRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ImpactRectangle.X)

            If String.IsNullOrEmpty(selRow.Impact) = False AndAlso selRow.ImpactImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.ImpactImage.Width, selRow.ImpactImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ImpactImage, rImage)
                Else
                    PrintClippedText(selRow.Impact, rImage, ImpactRectangle.Width)
                    selRow.Impact = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintAssumptionRow()
                        ClippedRow.Assumption = New Assumption
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.Impact = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintResponseStrategy(ByVal selRow As PrintAssumptionRow)
        If ResponseStrategyRectangle.Left >= CurrentSectionRectangle.Left And ResponseStrategyRectangle.Right <= CurrentSectionRectangle.Right Then
            Dim intX As Integer = GetCoordinateX(ResponseStrategyRectangle.X)

            If String.IsNullOrEmpty(selRow.ResponseStrategy) = False AndAlso selRow.ResponseStrategyImage IsNot Nothing Then
                Dim rImage As New Rectangle(intX, LastRowY, selRow.ResponseStrategyImage.Width, selRow.ResponseStrategyImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ResponseStrategyImage, rImage)
                Else
                    PrintClippedText(selRow.ResponseStrategy, rImage, ResponseStrategyRectangle.Width)
                    selRow.ResponseStrategy = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintAssumptionRow()
                        ClippedRow.Assumption = New Assumption
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.ResponseStrategy = strClippedTextBottom
                End If
            End If
        End If
    End Sub

    Private Sub PrintPage_PrintOwner(ByVal selRow As PrintAssumptionRow)
        If selRow.ClippedRow = False Then
            If OwnerRectangle.Left >= CurrentSectionRectangle.Left And OwnerRectangle.Right <= CurrentSectionRectangle.Right Then
                Dim intX As Integer = GetCoordinateX(OwnerRectangle.X)
                Dim rOwner As New Rectangle(intX, LastRowY, OwnerRectangle.Width, selRow.RowHeight)

                PrintPage_PrintText(selRow.Owner, rOwner)
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

            If ReasonRectangle.Left >= CurrentSectionRectangle.Left And ReasonRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ReasonRectangle.X)
                Dim rReason As New Rectangle(intX, LastRowY, ReasonRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Reason, rReason, True, False)

                LastColumnBorder = rReason.Right
            End If

            If HowToValidateRectangle.Left >= CurrentSectionRectangle.Left And HowToValidateRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(HowToValidateRectangle.X)
                Dim rHowToValidate As New Rectangle(intX, LastRowY, HowToValidateRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_HowToValidate, rHowToValidate, True, False)

                LastColumnBorder = rHowToValidate.Right
            End If

            If ValidatedRectangle.Left >= CurrentSectionRectangle.Left And ValidatedRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ValidatedRectangle.X)
                Dim rValidated As New Rectangle(intX, LastRowY, ValidatedRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Validated, rValidated, True, False)

                LastColumnBorder = rValidated.Right
            End If

            If ImpactRectangle.Left >= CurrentSectionRectangle.Left And ImpactRectangle.Right <= CurrentSectionRectangle.Right Then
                intX = GetCoordinateX(ImpactRectangle.X)
                Dim rImpact As New Rectangle(intX, LastRowY, ImpactRectangle.Width, ColumnHeadersHeight)

                PrintPage_PrintText(LANG_Impact, rImpact, True, False)

                LastColumnBorder = rImpact.Right
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
