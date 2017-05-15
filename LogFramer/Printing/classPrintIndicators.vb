Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintIndicators
    Inherits ReportBaseClass

    Private objLogFrame As New LogFrame
    Private objTargetGroupGuid As Guid
    Private intSection As Integer
    Private objPrintList As New PrintIndicatorRows
    Private objClippedRow As PrintIndicatorRow = Nothing
    Private strClippedTextTop, strClippedTextBottom As String

    Private boolColumnsWidthSet As Boolean
    Private intMeasurement As Integer
    Private boolPrintGoals, boolPrintPurposes As Boolean
    Private boolPrintOutputs, boolPrintActivities As Boolean
    Private boolPrintOptionValues, boolPrintRanges, boolPrintTargets As Boolean

    Private rSortRectangle As Rectangle
    Private rTextRectangle As Rectangle
    Private MaxDiffLeftRectangle, MaxDiffCenterRectangle, MaxDiffRightRectangle As Rectangle
    Private ScaleStatementRectangle, ScaleAgreeRectangle, ScaleDisagreeRectangle As Rectangle
    Private RowHeaderRectangle As Rectangle
    Private TableColumnRectangle As Rectangle()

    Private intLeftLabelWidth, intRightLabelWidth, intResponseClassWidth As Integer
    Private sfColumnText As New StringFormat

    Private Const CONST_OpenEndedLineSpacing As Integer = 30
    Private Const CONST_CheckBoxWidth As Integer = 80
    Private Const CONST_RowHeaderMaxWidth As Integer = 200
    Private Const CONST_CheckBoxIconSize As Integer = 10
    Private Const CONST_LeftRightLabelSize As Integer = 120

    Public Event LinePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)

    Public Sub New(ByVal logframe As LogFrame, ByVal printsection As Integer, ByVal targetgroupguid As Guid, ByVal measurement As Integer, _
                   ByVal printpurposes As Boolean, ByVal printoutputs As Boolean, ByVal printactivities As Boolean, ByVal printoptionvalues As Boolean, _
                   ByVal printranges As Boolean, ByVal printtargets As Boolean)
        Me.LogFrame = logframe
        Me.PrintSection = printsection
        Me.TargetGroupGuid = targetgroupguid
        Me.Measurement = measurement
        Me.PrintPurposes = printpurposes
        Me.PrintOutputs = printoutputs
        Me.PrintActivities = printactivities
        Me.PrintOptionValues = printoptionvalues
        Me.PrintValueRanges = printranges
        Me.PrintTargets = printtargets

        Me.ReportSetup = logframe.ReportSetupIndicators
    End Sub

#Region "Properties"
    Public Enum PrintSections As Integer
        NotSelected = -1
        PrintGoals = 0
        PrintPurposes = 1
        PrintOutputs = 2
        PrintActivities = 3
        PrintAll = 4
    End Enum

    Public Property LogFrame() As LogFrame
        Get
            Return objLogFrame
        End Get
        Set(ByVal value As LogFrame)
            objLogFrame = value
        End Set
    End Property

    Public Property PrintSection() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
            Select Case intSection
                Case PrintSections.PrintGoals
                    boolPrintGoals = True
                Case PrintSections.PrintPurposes
                    boolPrintPurposes = True
                Case PrintSections.PrintOutputs
                    boolPrintOutputs = True
                Case PrintSections.PrintActivities
                    boolPrintActivities = True
                Case PrintSections.PrintAll
                    boolPrintGoals = True
                    boolPrintPurposes = True
                    boolPrintOutputs = True
                    boolPrintActivities = True
            End Select
        End Set
    End Property

    Private Property Measurement As Integer
        Get
            Return intMeasurement
        End Get
        Set(value As Integer)
            intMeasurement = value
        End Set
    End Property

    Public Property PrintPurposes As Boolean
        Get
            Return boolPrintPurposes
        End Get
        Set(ByVal value As Boolean)
            boolPrintPurposes = value
        End Set
    End Property

    Public Property PrintOutputs As Boolean
        Get
            Return boolPrintOutputs
        End Get
        Set(ByVal value As Boolean)
            boolPrintOutputs = value
        End Set
    End Property

    Public Property PrintActivities As Boolean
        Get
            Return boolPrintActivities
        End Get
        Set(ByVal value As Boolean)
            boolPrintActivities = value
        End Set
    End Property

    Public Property PrintOptionValues As Boolean
        Get
            Return boolPrintOptionValues
        End Get
        Set(ByVal value As Boolean)
            boolPrintOptionValues = value
        End Set
    End Property

    Public Property PrintValueRanges As Boolean
        Get
            Return boolPrintRanges
        End Get
        Set(ByVal value As Boolean)
            boolPrintRanges = value
        End Set
    End Property

    Public Property PrintTargets As Boolean
        Get
            Return boolPrintTargets
        End Get
        Set(ByVal value As Boolean)
            boolPrintTargets = value
        End Set
    End Property

    Public Property TargetGroupGuid() As Guid
        Get
            Return objTargetGroupGuid
        End Get
        Set(ByVal value As Guid)
            objTargetGroupGuid = value
        End Set
    End Property

    Public Property PrintList() As PrintIndicatorRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintIndicatorRows)
            objPrintList = value
        End Set
    End Property

    Public Property ClippedRow As PrintIndicatorRow
        Get
            Return objClippedRow
        End Get
        Set(value As PrintIndicatorRow)
            objClippedRow = value
        End Set
    End Property

    Private Property SortRectangle As Rectangle
        Get
            Return rSortRectangle
        End Get
        Set(ByVal value As Rectangle)
            rSortRectangle = value
        End Set
    End Property

    Private Property TextRectangle As Rectangle
        Get
            Return rTextRectangle
        End Get
        Set(ByVal value As Rectangle)
            rTextRectangle = value
        End Set
    End Property
#End Region

#Region "Create list of objectives, indicators and statements"
    Public Sub CreateList()
        PrintList.Clear()
        'add goals and their indicators
        If Me.PrintSection = PrintSections.PrintGoals Or Me.PrintSection = PrintSections.PrintAll Then
            CreateList_Goals()
        End If

        'add purposes, outputs and activities with their indicators
        If Me.PrintSection > PrintSections.PrintGoals Then
            CreateList_Purposes()
        End If

        If PrintList.Count > 0 AndAlso PrintList(0).ObjectType = PrintIndicatorRow.ObjectTypes.WhiteSpace Then
            PrintList.Remove(0)
        End If
    End Sub

    Private Sub CreateList_Goals()
        Dim strSortNumber As String

        For Each selGoal As Goal In Me.LogFrame.Goals
            If selGoal.Indicators.Count > 0 Then
                Dim objWhiteSpace As New PrintIndicatorRow(60)
                PrintList.Add(objWhiteSpace)

                strSortNumber = Me.LogFrame.GetStructSortNumber(selGoal)
                Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.Goal, strSortNumber, selGoal.RTF)
                PrintList.Add(objRow)

                If Me.PrintSection = PrintSections.PrintGoals Or Me.PrintSection = PrintSections.PrintAll Then
                    CreateList_Indicators(selGoal.Indicators, strSortNumber, LogFrame.SectionTypes.GoalsSection)
                End If
            End If
        Next
    End Sub

    Private Sub CreateList_Purposes()
        Dim strSortNumber As String

        For Each selPurpose As Purpose In Me.LogFrame.Purposes
            strSortNumber = Me.LogFrame.GetStructSortNumber(selPurpose)
            If PrintPurposes = True Then
                Dim objWhiteSpace As New PrintIndicatorRow(60)
                PrintList.Add(objWhiteSpace)

                Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.Purpose, strSortNumber, selPurpose.RTF)
                PrintList.Add(objRow)
            End If
            If Me.PrintSection = PrintSections.PrintPurposes Or Me.PrintSection = PrintSections.PrintAll Then
                CreateList_Indicators(selPurpose.Indicators, strSortNumber, LogFrame.SectionTypes.PurposesSection)
            End If

            'add outputs
            If Me.PrintSection > PrintSections.PrintPurposes Then
                CreateList_Outputs(selPurpose.Outputs)
            End If
        Next
    End Sub

    Private Sub CreateList_Outputs(ByVal selOutputs As Outputs)
        Dim strSortNumber As String

        For Each selOutput As Output In selOutputs
            strSortNumber = Me.LogFrame.GetStructSortNumber(selOutput)
            If PrintOutputs = True Then
                Dim objWhiteSpace As New PrintIndicatorRow(20)
                PrintList.Add(objWhiteSpace)

                Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.Output, strSortNumber, selOutput.RTF)
                PrintList.Add(objRow)
            End If
            If Me.PrintSection = PrintSections.PrintOutputs Or Me.PrintSection = PrintSections.PrintAll Then
                CreateList_Indicators(selOutput.Indicators, strSortNumber, LogFrame.SectionTypes.OutputsSection)
            End If

            'add Activities
            If Me.PrintSection > PrintSections.PrintOutputs Then
                CreateList_Activities(selOutput.Activities, strSortNumber)
            End If
        Next
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String)
        Dim strActivitySort As String
        Dim intIndex As Integer

        For Each selActivity As Activity In selActivities
            strActivitySort = LogFrame.CreateSortNumber(intIndex, strParentSort)

            If selActivity.Indicators.Count > 0 Then
                If PrintActivities = True Then
                    Dim objWhiteSpace As New PrintIndicatorRow(20)
                    PrintList.Add(objWhiteSpace)

                    Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.Activity, strActivitySort, selActivity.RTF)
                    PrintList.Add(objRow)
                End If
                If Me.PrintSection = PrintSections.PrintActivities Or Me.PrintSection = PrintSections.PrintAll Then
                    CreateList_Indicators(selActivity.Indicators, strActivitySort, LogFrame.SectionTypes.ActivitiesSection)
                End If
            End If

            If selActivity.Activities.Count > 0 Then
                CreateList_Activities(selActivity.Activities, strActivitySort)
            End If

            intIndex += 1
        Next
    End Sub

    Private Sub CreateList_Indicators(ByVal selIndicators As Indicators, ByVal strParentSort As String, ByVal intSection As Integer)
        Dim strIndicatorSort As String
        Dim intIndex As Integer

        For Each selIndicator As Indicator In selIndicators
            If Me.TargetGroupGuid = Guid.Empty OrElse selIndicator.TargetGroupGuid = Me.TargetGroupGuid Then
                'leading white space above indicator
                Dim objWhiteSpace As New PrintIndicatorRow(20)
                PrintList.Add(objWhiteSpace)

                'indicator number and text
                strIndicatorSort = LogFrame.CreateSortNumber(intIndex, strParentSort)
                Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.Indicator, strIndicatorSort, selIndicator.RTF)
                PrintList.Add(objRow)

                'answer area (text lines, table...)
                CreateList_AnswerArea(selIndicator, intSection)

                If selIndicator.Indicators.Count > 0 Then
                    CreateList_Indicators(selIndicator.Indicators, strIndicatorSort, intSection)
                End If
            End If
            intIndex += 1
        Next
    End Sub

    Private Sub CreateList_AnswerArea(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Select Case selIndicator.QuestionType
            Case Indicator.QuestionTypes.OpenEnded
                CreateList_AnswerArea_OpenEnded(selIndicator, intSection)
            Case Indicator.QuestionTypes.MaxDiff
                CreateList_AnswerArea_MaxDiff(selIndicator, intSection)
            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                CreateList_AnswerArea_Value(selIndicator, intSection)
            Case Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.Ratio
                CreateList_AnswerArea_Formula(selIndicator, intSection)
            Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.Ranking
                CreateList_AnswerArea_MultipleOptions(selIndicator, intSection)
            Case Indicator.QuestionTypes.LikertTypeScale
                CreateList_AnswerArea_LikertType(selIndicator, intSection)
            Case Indicator.QuestionTypes.SemanticDiff
                CreateList_AnswerArea_SemanticDiff(selIndicator, intSection)
            Case Indicator.QuestionTypes.Scale
                CreateList_AnswerArea_Scale(selIndicator, intSection)
            Case Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert
                CreateList_AnswerArea_LikertScale(selIndicator, intSection)
            Case Indicator.QuestionTypes.CumulativeScale
                CreateList_AnswerArea_CumulativeScale(selIndicator, intSection)
        End Select
    End Sub

    Private Function CreateList_AnswerArea_ScoringLegend(ByVal selIndicator As Indicator, ByVal strScore As String) As String
        Dim strRowHeaderScore As String = String.Empty
        Dim intTargetIndex As Integer

        strRowHeaderScore = ToLabel(strScore)
        For Each selTarget As Target In selIndicator.Targets
            strRowHeaderScore &= String.Format("{0}  {1}: {2}", vbCrLf, selIndicator.ValuesDetail.DisplayField(selTarget), selIndicator.GetTargetFormattedScore(intTargetIndex))
            intTargetIndex += 1
        Next

        Return strRowHeaderScore
    End Function

    Private Sub CreateList_AnswerArea_OpenEnded(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim intRowHeight As Integer

        If selIndicator.Indicators.Count = 0 Then
            Select Case selIndicator.OpenEndedDetail.WhiteSpace
                Case OpenEndedDetail.WhiteSpaceUnits.QuarterPage
                    intRowHeight = Me.ContentHeight / 4
                Case OpenEndedDetail.WhiteSpaceUnits.ThirdPage
                    intRowHeight = Me.ContentHeight / 3
                Case OpenEndedDetail.WhiteSpaceUnits.HalfPage
                    intRowHeight = Me.ContentHeight / 2
                Case OpenEndedDetail.WhiteSpaceUnits.ThreeQuarters
                    intRowHeight = Me.ContentHeight * 0.75
                Case OpenEndedDetail.WhiteSpaceUnits.Page
                    intRowHeight = Me.ContentHeight
                Case Else
                    Dim intLines As Integer = selIndicator.OpenEndedDetail.WhiteSpace + 1
                    intRowHeight = CONST_OpenEndedLineSpacing * intLines
            End Select

            If Measurement = 0 Then
                Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaOpenEnded, selIndicator)
                objRow.RowHeight = intRowHeight
                PrintList.Add(objRow)
            Else
                Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)

                Dim objBaselineAreaRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaOpenEnded, selIndicator)
                objBaselineAreaRow.RowHeight = intRowHeight
                objBaselineAreaRow.TargetDeadlineDate = LANG_Baseline
                PrintList.Add(objBaselineAreaRow)

                If objTargetDeadlinesSection IsNot Nothing Then
                    For Each selTargetDeadline As TargetDeadline In objTargetDeadlinesSection.TargetDeadlines
                        Dim strDate As String = objTargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)
                        Dim objTargetAreaRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaOpenEnded, selIndicator)
                        objTargetAreaRow.RowHeight = intRowHeight
                        objTargetAreaRow.TargetDeadlineDate = strDate
                        PrintList.Add(objTargetAreaRow)
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_MaxDiff(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim strColumns() As String = New String() {selIndicator.ScalesDetail.AgreeText, LANG_Options, selIndicator.ScalesDetail.DisagreeText}

        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, strColumns)
                PrintList.Add(objColumnHeaderRow)

                For Each selStatement As Statement In selIndicator.Statements
                    Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, String.Empty, selStatement.FirstLabel)
                    PrintList.Add(objRow)
                Next
            Else
                Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
                Dim objBaselineTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, LANG_Baseline)
                PrintList.Add(objBaselineTableHeaderRow)
                Dim objBaselineColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, strColumns)
                PrintList.Add(objBaselineColumnHeaderRow)

                For Each selStatement As Statement In selIndicator.Statements
                    Dim objBaselineRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, String.Empty, selStatement.FirstLabel)
                    objBaselineRow.TargetDeadlineDate = LANG_Baseline
                    PrintList.Add(objBaselineRow)
                Next

                If objTargetDeadlinesSection IsNot Nothing Then
                    For Each selTargetDeadline As TargetDeadline In objTargetDeadlinesSection.TargetDeadlines
                        Dim strDate As String = objTargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)

                        Dim objTargetTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, strDate)
                        PrintList.Add(objTargetTableHeaderRow)
                        Dim objTargetColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, strColumns)
                        PrintList.Add(objTargetColumnHeaderRow)

                        For Each selStatement As Statement In selIndicator.Statements
                            Dim objTargetRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, String.Empty, selStatement.FirstLabel)
                            objTargetRow.TargetDeadlineDate = strDate
                            PrintList.Add(objTargetRow)
                        Next
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_Value(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_Value_SingleMeasurement(selIndicator)
            Else
                CreateList_AnswerArea_Value_Table(selIndicator)
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_Value_SingleMeasurement(ByVal selIndicator As Indicator)
        Dim strRowHeaderTarget As String, strRowHeaderScore As String = String.Empty
        Dim intRowHeaderWidth
        Dim intColumnwidth As Integer
        Dim strCells(0) As String

        'row header
        If Me.PrintValueRanges And selIndicator.ValuesDetail.ValueRange.ValueRangeSet = True Then
            strRowHeaderTarget = String.Format("{0}:{1}{2}", LANG_ValueWithinRange, vbCrLf, selIndicator.ValuesDetail.DisplayField(selIndicator.ValuesDetail.ValueRange))
        Else
            strRowHeaderTarget = ToLabel(LANG_Value)
        End If

        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            strRowHeaderScore = CreateList_AnswerArea_ScoringLegend(selIndicator, LANG_Score)
        End If

        intRowHeaderWidth = CreateList_AnswerArea_Value_GetRowHeaderWidth(strRowHeaderTarget, strRowHeaderScore, selIndicator.ScoringSystem)
        intColumnwidth = ContentWidth - intRowHeaderWidth

        'value cell
        strCells(0) = selIndicator.ValuesDetail.Unit

        Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaValue, strRowHeaderTarget, strCells, intRowHeaderWidth, intColumnwidth)
        PrintList.Add(objRow)

        'scoring cell
        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            Dim strCellsScore(0) As String

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaValue, strRowHeaderScore, strCellsScore, intRowHeaderWidth, intColumnwidth)
            PrintList.Add(objRow)
        End If
    End Sub

    Private Sub CreateList_AnswerArea_Value_Table(ByVal selIndicator As Indicator)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines
        Dim intRowHeaderWidth As Integer
        Dim intColumnwidth As Integer
        Dim strRowHeaderTarget As String, strRowHeaderScore As String = String.Empty
        Dim intTargetCount As Integer = objTargetDeadlines.Count
        Dim intColumns As Integer = intTargetCount + 1
        Dim strColumns(intTargetCount) As String
        Dim strCells(intTargetCount) As String
        Dim strColumnHeader As String

        'row header
        If Me.PrintValueRanges And selIndicator.ValuesDetail.ValueRange.ValueRangeSet = True Then
            strRowHeaderTarget = String.Format("{0}:{1}{2}", LANG_ValuesWithinRange, vbCrLf, selIndicator.ValuesDetail.DisplayField(selIndicator.ValuesDetail.ValueRange))
        Else
            strRowHeaderTarget = ToLabel(LANG_Values)
        End If

        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            strRowHeaderScore = CreateList_AnswerArea_ScoringLegend(selIndicator, LANG_Scores)
        End If

        intRowHeaderWidth = CreateList_AnswerArea_Value_GetRowHeaderWidth(strRowHeaderTarget, strRowHeaderScore, selIndicator.ScoringSystem)
        intColumnwidth = (ContentWidth - intRowHeaderWidth) / (intColumns)

        'column headers
        strColumns(0) = LANG_Baseline
        For i = 0 To intTargetCount - 1
            strColumnHeader = objTargetDeadlinesSection.FormatTargetDeadlineDate(objTargetDeadlines(i))

            If Me.PrintTargets Then
                Dim strTarget As String = String.Empty
                If i <= selIndicator.Targets.Count - 1 Then strTarget = selIndicator.GetTargetFormattedValue(i)
                strColumnHeader = String.Format("{0}{1}{1}{2}: {3}", strColumnHeader, vbCrLf, LANG_Target, strTarget)
            End If
            strColumns(i + 1) = strColumnHeader
        Next

        'value cells
        For i = 0 To intColumns - 1
            strCells(i) = selIndicator.ValuesDetail.Unit
        Next

        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaValue, strColumns, intRowHeaderWidth, intColumnwidth)
        PrintList.Add(objColumnHeaderRow)
        Dim objRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaValue, strRowHeaderTarget, strCells, intRowHeaderWidth, intColumnwidth)
        PrintList.Add(objRow)

        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            Dim strCellsScore(intTargetCount) As String

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaValue, strRowHeaderScore, strCellsScore, intRowHeaderWidth, intColumnwidth)
            PrintList.Add(objRow)
        End If
    End Sub

    Private Function CreateList_AnswerArea_Value_GetRowHeaderWidth(ByVal strRowHeaderTarget As String, ByVal strRowHeaderScore As String, ByVal intScoringSystem As Integer) As Integer
        Dim intRowHeaderWidth, intRowHeaderWidthScore As Integer

        intRowHeaderWidth = GetTextWidth(strRowHeaderTarget, False) + CONST_VerticalPadding
        If intScoringSystem = Indicator.ScoringSystems.Score Then
            intRowHeaderWidthScore = GetTextWidth(strRowHeaderScore, False) + CONST_VerticalPadding
            If intRowHeaderWidthScore > intRowHeaderWidth Then intRowHeaderWidth = intRowHeaderWidthScore
        End If
        If intRowHeaderWidth > (CONST_RowHeaderMaxWidth * 2) Then intRowHeaderWidth = (CONST_RowHeaderMaxWidth * 2)

        Return intRowHeaderWidth
    End Function

    Private Sub CreateList_AnswerArea_Formula(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_Formula_SingleMeasurement(selIndicator)
            Else
                CreateList_AnswerArea_Formula_Table(selIndicator)
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_Formula_SingleMeasurement(ByVal selIndicator As Indicator)
        Dim intStatementWidth
        Dim intColumnwidth As Integer
        Dim objRow As PrintIndicatorRow
        Dim strTotal As String

        intStatementWidth = CONST_RowHeaderMaxWidth
        intColumnwidth = ContentWidth - intStatementWidth

        For Each selStatement As Statement In selIndicator.Statements
            Dim strCells(0) As String
            strCells(0) = selStatement.ValuesDetail.Unit

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, selStatement.FirstLabel, strCells, intStatementWidth, intColumnwidth)
            PrintList.Add(objRow)
        Next

        Dim strCellsTotal(0) As String
        strCellsTotal(0) = selIndicator.ValuesDetail.Unit
        strTotal = String.Format("{0} ({1})", LANG_Total, selIndicator.ValuesDetail.Formula)

        strTotal = TextToRichText(strTotal)
        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, strTotal, strCellsTotal, intStatementWidth, intColumnwidth)
        PrintList.Add(objRow)

        'Scoring
        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            Dim strRowHeaderScore As String = CreateList_AnswerArea_ScoringLegend(selIndicator, LANG_Score)
            Dim strCellsScore(0) As String

            strRowHeaderScore = TextToRichText(strRowHeaderScore)
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, strRowHeaderScore, strCellsScore, intStatementWidth, intColumnwidth)
            PrintList.Add(objRow)
        End If

    End Sub

    Private Sub CreateList_AnswerArea_Formula_Table(ByVal selIndicator As Indicator)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines
        Dim objRow As PrintIndicatorRow
        Dim intStatementWidth As Integer
        Dim intColumnwidth As Integer
        Dim intTargetCount As Integer = objTargetDeadlines.Count
        Dim intColumns As Integer = intTargetCount + 1
        Dim strColumns(intTargetCount) As String
        Dim strColumnHeader As String
        Dim strTotal As String

        intStatementWidth = CONST_RowHeaderMaxWidth
        intColumnwidth = (ContentWidth - intStatementWidth) / (intColumns)

        'column headers
        strColumns(0) = LANG_Baseline
        For i = 0 To intTargetCount - 1
            strColumnHeader = objTargetDeadlinesSection.FormatTargetDeadlineDate(objTargetDeadlines(i))

            If Me.PrintTargets Then
                Dim strTarget As String = String.Empty
                If i <= selIndicator.Targets.Count - 1 Then strTarget = selIndicator.GetTargetFormattedValue(i)
                strColumnHeader = String.Format("{0}{1}{1}{2}: {3}", strColumnHeader, vbCrLf, LANG_Target, strTarget)
            End If
            strColumns(i + 1) = strColumnHeader
        Next

        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, strColumns, intStatementWidth, intColumnwidth)
        PrintList.Add(objColumnHeaderRow)

        For Each selStatement As Statement In selIndicator.Statements
            'value cells
            Dim strCells(intTargetCount) As String

            For i = 0 To intColumns - 1
                strCells(i) = selStatement.ValuesDetail.Unit
            Next
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, selStatement.FirstLabel, strCells, intStatementWidth, intColumnwidth)
            PrintList.Add(objRow)
        Next

        Dim strCellsTotal(intTargetCount) As String

        For i = 0 To intColumns - 1
            strCellsTotal(i) = selIndicator.ValuesDetail.Unit
        Next
        strTotal = String.Format("{0} ({1})", LANG_Total, selIndicator.ValuesDetail.Formula)

        strTotal = TextToRichText(strTotal)
        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, strTotal, strCellsTotal, intStatementWidth, intColumnwidth)
        PrintList.Add(objRow)

        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            Dim strRowHeaderScore As String = CreateList_AnswerArea_ScoringLegend(selIndicator, LANG_Score)
            Dim strCellsScore(intTargetCount) As String

            strRowHeaderScore = TextToRichText(strRowHeaderScore)

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, strRowHeaderScore, strCellsScore, intStatementWidth, intColumnwidth)
            PrintList.Add(objRow)
        End If
    End Sub

    Private Sub CreateList_AnswerArea_MultipleOptions(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_MultipleOptions_SingleMeasurement(selIndicator)
            Else
                CreateList_AnswerArea_MultipleOptions_Table(selIndicator)
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_MultipleOptions_SingleMeasurement(ByVal selIndicator As Indicator)
        Dim intOptionWidth
        Dim intColumnwidth As Integer
        Dim objRow As PrintIndicatorRow
        Dim strTableHeader As String

        intOptionWidth = CONST_RowHeaderMaxWidth
        intColumnwidth = ContentWidth - intOptionWidth

        If selIndicator.QuestionType = Indicator.QuestionTypes.Ranking Then
            strTableHeader = String.Format(LANG_RankingEnterNumber, selIndicator.ResponseClasses.Count)
        ElseIf selIndicator.QuestionType = Indicator.QuestionTypes.MultipleOptions Then
            strTableHeader = LANG_CheckMultipleOptions
        Else
            strTableHeader = LANG_CheckOneOption
        End If

        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strTableHeader)
        PrintList.Add(objRow)

        'Options
        For Each selOption As ResponseClass In selIndicator.ResponseClasses
            Dim strOption As String = selOption.ClassName
            Dim strCells(0) As String

            If PrintOptionValues And selIndicator.QuestionType <> Indicator.QuestionTypes.Ranking Then
                strOption = String.Format("{0} ({1})", strOption, selOption.Value)
            End If
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strOption, strCells, intOptionWidth, intColumnwidth)
            If selIndicator.QuestionType = Indicator.QuestionTypes.Ranking Then objRow.NoCheckBox = True
            PrintList.Add(objRow)
        Next

        'Scoring
        If selIndicator.QuestionType <> Indicator.QuestionTypes.Ranking Then
            Dim strRowHeaderScore As String = LANG_Score
            Dim strCellsScore(0) As String

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strRowHeaderScore, strCellsScore, intOptionWidth, intColumnwidth, True)
            PrintList.Add(objRow)
        End If
    End Sub

    Private Sub CreateList_AnswerArea_MultipleOptions_Table(ByVal selIndicator As Indicator)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines
        Dim objRow As PrintIndicatorRow
        Dim intOptionWidth As Integer
        Dim intColumnwidth As Integer
        Dim intTargetCount As Integer = objTargetDeadlines.Count
        Dim intColumns As Integer = intTargetCount + 1
        Dim strColumns(intTargetCount) As String
        Dim strColumnHeader As String
        Dim strTableHeader As String

        intOptionWidth = CONST_RowHeaderMaxWidth
        intColumnwidth = (ContentWidth - intOptionWidth) / (intColumns)

        If selIndicator.QuestionType = Indicator.QuestionTypes.Ranking Then
            strTableHeader = String.Format(LANG_RankingEnterNumber, selIndicator.ResponseClasses.Count)
        ElseIf selIndicator.QuestionType = Indicator.QuestionTypes.MultipleOptions Then
            strTableHeader = LANG_CheckMultipleOptions
        Else
            strTableHeader = LANG_CheckOneOption
        End If
        strTableHeader = String.Format("{0} ({1})", strTableHeader, LANG_PerColumn)

        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strTableHeader)
        PrintList.Add(objRow)

        'column headers
        strColumns(0) = LANG_Baseline
        For i = 0 To intTargetCount - 1
            strColumnHeader = objTargetDeadlinesSection.FormatTargetDeadlineDate(objTargetDeadlines(i))

            If Me.PrintTargets And selIndicator.QuestionType <> Indicator.QuestionTypes.Ranking Then
                Dim strTarget As String = String.Empty
                If i <= selIndicator.Targets.Count - 1 Then strTarget = selIndicator.GetTargetFormattedScore(i)

                strColumnHeader = String.Format("{0}{1}{1}{2}: {3}", strColumnHeader, vbCrLf, LANG_Target, strTarget)
            End If
            strColumns(i + 1) = strColumnHeader
        Next

        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strColumns, intOptionWidth, intColumnwidth)
        PrintList.Add(objColumnHeaderRow)

        'Options
        For Each selOption As ResponseClass In selIndicator.ResponseClasses
            Dim strOption As String = selOption.ClassName
            Dim strCells(intTargetCount) As String

            If PrintOptionValues And selIndicator.QuestionType <> Indicator.QuestionTypes.Ranking Then
                strOption = String.Format("{0} ({1})", strOption, selOption.Value)
            End If
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strOption, strCells, intOptionWidth, intColumnwidth)
            If selIndicator.QuestionType = Indicator.QuestionTypes.Ranking Then objRow.NoCheckBox = True
            PrintList.Add(objRow)
        Next

        'scoring
        If selIndicator.QuestionType <> Indicator.QuestionTypes.Ranking Then
            Dim strRowHeaderScore As String = LANG_Score
            Dim strCellsScore(intTargetCount) As String

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strRowHeaderScore, strCellsScore, intOptionWidth, intColumnwidth, True)
            PrintList.Add(objRow)
        End If
    End Sub

    Private Sub CreateList_AnswerArea_LikertType(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines
        Dim objRow As PrintIndicatorRow
        Dim intRowHeaderWidth As Integer
        Dim intColumnwidth As Integer
        Dim intTargetCount As Integer = objTargetDeadlines.Count
        Dim intOptionCount As Integer = selIndicator.ResponseClasses.Count
        Dim strColumns(intOptionCount - 1) As String
        Dim strRowHeader As String

        If selIndicator.Indicators.Count = 0 Then
            intRowHeaderWidth = CONST_RowHeaderMaxWidth
            intColumnwidth = (ContentWidth - intRowHeaderWidth) / (intOptionCount)

            'column headers
            For i = 0 To selIndicator.ResponseClasses.Count - 1
                Dim strOption As String = selIndicator.ResponseClasses(i).ClassName

                If PrintOptionValues Then
                    strOption = String.Format("{0}{1}({2})", strOption, vbCrLf, selIndicator.ResponseClasses(i).Value)
                End If

                strColumns(i) = strOption
            Next

            Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaLikertType, strColumns, intRowHeaderWidth, intColumnwidth)
            PrintList.Add(objColumnHeaderRow)

            'Rows
            Dim strCellsBaseline(intOptionCount - 1) As String

            If Measurement = 0 Then
                strRowHeader = LANG_SelectOption
            Else
                strRowHeader = LANG_Baseline
            End If
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaLikertType, strRowHeader, strCellsBaseline, intRowHeaderWidth, intColumnwidth)
            PrintList.Add(objRow)

            If Measurement <> 0 Then
                For i = 0 To intTargetCount - 1
                    strRowHeader = objTargetDeadlinesSection.FormatTargetDeadlineDate(objTargetDeadlines(i))
                    Dim strCells(intOptionCount - 1) As String

                    If Me.PrintTargets Then
                        Dim strTarget As String = String.Empty
                        If i <= selIndicator.Targets.Count - 1 Then strTarget = selIndicator.GetTargetFormattedScore(i)

                        strRowHeader = String.Format("{0} - {1}: {2}", strRowHeader, LANG_Target, strTarget)
                    End If

                    objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaLikertType, strRowHeader, strCells, intRowHeaderWidth, intColumnwidth)
                    PrintList.Add(objRow)
                Next
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_SemanticDiff(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines

        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_SemanticDiff_Table(selIndicator)
            Else
                Dim objBaselineTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, LANG_Baseline)
                PrintList.Add(objBaselineTableHeaderRow)
                CreateList_AnswerArea_SemanticDiff_Table(selIndicator)

                For Each selTargetDeadline As TargetDeadline In objTargetDeadlines
                    Dim strDate As String = objTargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)

                    Dim objTargetTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, strDate)
                    PrintList.Add(objTargetTableHeaderRow)
                    CreateList_AnswerArea_SemanticDiff_Table(selIndicator)
                Next
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_SemanticDiff_Table(ByVal selIndicator As Indicator)
        Dim objRow As PrintIndicatorRow
        Dim intLabelWidth As Integer
        Dim intColumnwidth As Integer
        Dim intOptionCount As Integer = selIndicator.ResponseClasses.Count
        Dim strColumns(intOptionCount - 1) As String

        intLabelWidth = CONST_LeftRightLabelSize
        intColumnwidth = (ContentWidth - (intLabelWidth * 2)) / (intOptionCount)

        'column headers
        For i = 0 To selIndicator.ResponseClasses.Count - 1
            Dim strOption As String = selIndicator.ResponseClasses(i).ClassName

            If PrintOptionValues Then
                strOption = String.Format("{0}{1}({2})", strOption, vbCrLf, selIndicator.ResponseClasses(i).Value)
            End If

            strColumns(i) = strOption
        Next

        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff, strColumns, intLabelWidth, intColumnwidth)
        PrintList.Add(objColumnHeaderRow)

        'Rows
        For Each selStatement As Statement In selIndicator.Statements
            'value cells
            Dim strCells(intOptionCount) As String
            strCells(intOptionCount) = selStatement.SecondLabel
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff, selStatement.FirstLabel, strCells, intLabelWidth, intColumnwidth)
            PrintList.Add(objRow)
        Next

        'Scoring
        Dim strRowHeaderScore As String = TextToRichText(LANG_Score)
        Dim strCellsScore(intOptionCount) As String

        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff, strRowHeaderScore, strCellsScore, intLabelWidth, intColumnwidth, True)
        PrintList.Add(objRow)
    End Sub

    Private Sub CreateList_AnswerArea_Scale(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines

        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_Scale_Table(selIndicator)
            Else
                Dim objBaselineTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, LANG_Baseline)
                PrintList.Add(objBaselineTableHeaderRow)
                CreateList_AnswerArea_Scale_Table(selIndicator)

                For Each selTargetDeadline As TargetDeadline In objTargetDeadlines
                    Dim strDate As String = objTargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)

                    Dim objTargetTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, strDate)
                    PrintList.Add(objTargetTableHeaderRow)
                    CreateList_AnswerArea_Scale_Table(selIndicator)
                Next
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_Scale_Table(ByVal selIndicator As Indicator)
        Dim objRow As PrintIndicatorRow
        Dim selStatement As Statement = Nothing
        Dim selResponseClass As ResponseClass = Nothing
        Dim strColumns() As String = New String() {LANG_Statements, selIndicator.ScalesDetail.AgreeText, selIndicator.ScalesDetail.DisagreeText}
        Dim strStatement As String

        'column headers
        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, strColumns)
        PrintList.Add(objColumnHeaderRow)

        'Rows
        For i = 0 To selIndicator.Statements.Count - 1
            selStatement = selIndicator.Statements(i)

            strStatement = selStatement.FirstLabel

            If PrintOptionValues Then
                selResponseClass = selIndicator.ResponseClasses(i)

                If selResponseClass IsNot Nothing Then
                    With RichTextManager
                        .Rtf = selStatement.FirstLabel
                        .AppendText(String.Format(" ({0})", selResponseClass.Value))

                        strStatement = .Rtf
                    End With
                End If
            End If

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, String.Empty, strStatement)
            PrintList.Add(objRow)
        Next

        'Scoring
        Dim strRowHeaderScore As String = TextToRichText(LANG_Score)

        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, String.Empty, strRowHeaderScore)
        objRow.NoCheckBox = True
        PrintList.Add(objRow)
    End Sub

    Private Sub CreateList_AnswerArea_LikertScale(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines

        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_LikertScale_Table(selIndicator)
            Else
                Dim objBaselineTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, LANG_Baseline)
                PrintList.Add(objBaselineTableHeaderRow)
                CreateList_AnswerArea_LikertScale_Table(selIndicator)

                For Each selTargetDeadline As TargetDeadline In objTargetDeadlines
                    Dim strDate As String = objTargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)

                    Dim objTargetTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaScale, strDate)
                    PrintList.Add(objTargetTableHeaderRow)
                    CreateList_AnswerArea_LikertScale_Table(selIndicator)
                Next
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_LikertScale_Table(ByVal selIndicator As Indicator)
        Dim objRow As PrintIndicatorRow
        Dim intOptionCount As Integer = selIndicator.ResponseClasses.Count
        Dim strColumns(intOptionCount - 1) As String
        Dim intRowHeaderWidth As Integer
        Dim intColumnwidth As Integer

        intRowHeaderWidth = CONST_RowHeaderMaxWidth
        intColumnwidth = (ContentWidth - intRowHeaderWidth) / (intOptionCount)

        'column headers
        For i = 0 To selIndicator.ResponseClasses.Count - 1
            Dim strOption As String = selIndicator.ResponseClasses(i).ClassName

            If PrintOptionValues Then
                strOption = String.Format("{0}{1}({2})", strOption, vbCrLf, selIndicator.ResponseClasses(i).Value)
            End If

            strColumns(i) = strOption
        Next

        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale, strColumns, intRowHeaderWidth, intColumnwidth)
        PrintList.Add(objColumnHeaderRow)

        'Rows
        Dim strCells(intOptionCount - 1) As String

        For Each selStatement As Statement In selIndicator.Statements
            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale, selStatement.FirstLabel, strCells, intRowHeaderWidth, intColumnwidth)
            If selIndicator.QuestionType = Indicator.QuestionTypes.FrequencyLikert Then objRow.NoCheckBox = True
            PrintList.Add(objRow)
        Next

        'Scoring
        Dim strRowHeaderScore As String = TextToRichText(LANG_Score)
        Dim strCellsScore(intOptionCount - 1) As String

        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale, strRowHeaderScore, strCellsScore, intRowHeaderWidth, intColumnwidth, True)
        PrintList.Add(objRow)
    End Sub

    Private Sub CreateList_AnswerArea_CumulativeScale(ByVal selIndicator As Indicator, ByVal intSection As Integer)
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = Me.LogFrame.GetTargetDeadlinesSection(intSection)
        Dim objTargetDeadlines As TargetDeadlines = objTargetDeadlinesSection.TargetDeadlines

        If selIndicator.Indicators.Count = 0 Then
            If Measurement = 0 Then
                CreateList_AnswerArea_CumulativeScale_Table(selIndicator)
            Else
                Dim objBaselineTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale, LANG_Baseline)
                PrintList.Add(objBaselineTableHeaderRow)
                CreateList_AnswerArea_CumulativeScale_Table(selIndicator)

                For Each selTargetDeadline As TargetDeadline In objTargetDeadlines
                    Dim strDate As String = objTargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)

                    Dim objTargetTableHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale, strDate)
                    PrintList.Add(objTargetTableHeaderRow)
                    CreateList_AnswerArea_CumulativeScale_Table(selIndicator)
                Next
            End If
        End If
    End Sub

    Private Sub CreateList_AnswerArea_CumulativeScale_Table(ByVal selIndicator As Indicator)
        Dim objRow As PrintIndicatorRow
        Dim selStatement As Statement = Nothing
        Dim selResponseClass As ResponseClass = selIndicator.ResponseClasses(0)
        Dim dblValue As Double = selResponseClass.Value
        Dim dblTotal As Double = dblValue
        Dim strColumns() As String = New String() {LANG_Statements, selIndicator.ScalesDetail.AgreeText, selIndicator.ScalesDetail.DisagreeText}
        Dim strStatement As String

        'column headers
        Dim objColumnHeaderRow As New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale, strColumns)
        PrintList.Add(objColumnHeaderRow)

        'Rows
        For i = 0 To selIndicator.Statements.Count - 1
            selStatement = selIndicator.Statements(i)

            If PrintOptionValues Then
                With RichTextManager
                    .Rtf = selStatement.FirstLabel
                    .AppendText(String.Format(" ({0})", dblTotal))

                    strStatement = .Rtf
                    dblTotal += dblValue
                End With
            Else
                strStatement = selStatement.FirstLabel
            End If

            objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale, String.Empty, strStatement)
            PrintList.Add(objRow)
        Next

        'Scoring
        Dim strRowHeaderScore As String = TextToRichText(LANG_Score)

        objRow = New PrintIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale, String.Empty, strRowHeaderScore)
        objRow.NoCheckBox = True
        PrintList.Add(objRow)
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intSortWidth As Integer = GetSortColumnWidth()
        Dim intTextColumnWidth As Integer = ContentWidth - intSortWidth

        SortRectangle = New Rectangle(LeftMargin, ContentTop, intSortWidth, ContentHeight)
        TextRectangle = New Rectangle(SortRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight)

        MaxDiffLeftRectangle = New Rectangle(LeftMargin, ContentTop, CONST_CheckBoxWidth, ContentHeight)
        MaxDiffCenterRectangle = New Rectangle(MaxDiffLeftRectangle.Right, ContentTop, ContentWidth - (CONST_CheckBoxWidth * 2), ContentHeight)
        MaxDiffRightRectangle = New Rectangle(MaxDiffCenterRectangle.Right, ContentTop, CONST_CheckBoxWidth, ContentHeight)

        ScaleDisagreeRectangle = New Rectangle(ContentRight - CONST_CheckBoxWidth, ContentTop, CONST_CheckBoxWidth, ContentHeight)
        ScaleAgreeRectangle = New Rectangle(ScaleDisagreeRectangle.Left - CONST_CheckBoxWidth, ContentTop, CONST_CheckBoxWidth, ContentHeight)
        ScaleStatementRectangle = New Rectangle(ContentLeft, ContentTop, ContentWidth - (CONST_CheckBoxWidth * 2), ContentHeight)
    End Sub

    Private Function GetSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintIndicatorRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.SortNumber) = False AndAlso selRow.SortNumber.Length > strSort.Length Then strSort = selRow.SortNumber
        Next
        If String.IsNullOrEmpty(strSort) = False Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + (CONST_HorizontalPadding * 2)
        End If
        Return intWidth
    End Function
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintIndicatorRow In Me.PrintList
            ReloadImages_Normal(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintIndicatorRow)
        Dim intWidth As Integer

        With RichTextManager
            Select Case selRow.ObjectType
                Case PrintIndicatorRow.ObjectTypes.Goal, PrintIndicatorRow.ObjectTypes.Purpose, PrintIndicatorRow.ObjectTypes.Output, PrintIndicatorRow.ObjectTypes.Activity, PrintIndicatorRow.ObjectTypes.Indicator
                    intWidth = TextRectangle.Width
                    If String.IsNullOrEmpty(selRow.Rtf) = False Then _
                        selRow.FirstImage = .RichTextWithPaddingToBitmap(intWidth, selRow.Rtf, False)
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff
                    If String.IsNullOrEmpty(selRow.Rtf) = False Then
                        intWidth = MaxDiffCenterRectangle.Width
                        selRow.FirstImage = .RichTextWithPaddingToBitmap(intWidth, selRow.Rtf, False, 0, HorizontalAlignment.Center)
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaFormula
                    If String.IsNullOrEmpty(selRow.RowHeader) = False Then
                        intWidth = selRow.RowHeaderWidth - CONST_VerticalPadding
                        selRow.FirstImage = .RichTextWithPaddingToBitmap(intWidth, selRow.RowHeader, False, 0, HorizontalAlignment.Left, Color.Black, Color.LightGray)
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff
                    If String.IsNullOrEmpty(selRow.RowHeader) = False Then
                        intWidth = selRow.RowHeaderWidth - CONST_VerticalPadding
                        selRow.FirstImage = .RichTextWithPaddingToBitmap(intWidth, selRow.RowHeader, False, 0, HorizontalAlignment.Center, Color.Black, Color.LightGray)

                        If selRow.Cells.Count > 0 Then
                            Dim intIndex As Integer = selRow.Cells.Count - 1

                            If String.IsNullOrEmpty(selRow.Cells(intIndex)) = False Then _
                                selRow.SecondImage = .RichTextWithPaddingToBitmap(intWidth, selRow.Cells(intIndex), False, 0, HorizontalAlignment.Center, Color.Black, Color.LightGray)
                        End If
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaScale, PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale
                    If String.IsNullOrEmpty(selRow.Rtf) = False Then
                        intWidth = ScaleStatementRectangle.Width
                        selRow.FirstImage = .RichTextWithPaddingToBitmap(intWidth, selRow.Rtf, False, 0, HorizontalAlignment.Left)
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale
                    If String.IsNullOrEmpty(selRow.RowHeader) = False Then
                        intWidth = selRow.RowHeaderWidth - CONST_VerticalPadding
                        selRow.FirstImage = .RichTextWithPaddingToBitmap(intWidth, selRow.RowHeader, False, 0, HorizontalAlignment.Left, Color.Black, Color.LightGray)
                    End If
            End Select
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As PrintIndicatorRow = Me.PrintList(RowIndex)
        Dim intHeight, intRowHeight As Integer

        If selRow IsNot Nothing Then
            Select Case selRow.ObjectType
                Case PrintIndicatorRow.ObjectTypes.Goal, PrintIndicatorRow.ObjectTypes.Purpose, PrintIndicatorRow.ObjectTypes.Output, PrintIndicatorRow.ObjectTypes.Activity, _
                    PrintIndicatorRow.ObjectTypes.Indicator
                    If selRow.FirstImage IsNot Nothing Then
                        intRowHeight = selRow.FirstImage.Height
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff
                    If String.IsNullOrEmpty(selRow.TableHeader) = False Then
                        intRowHeight = GetTextHeight(selRow.TableHeader, ContentWidth, True)
                    ElseIf selRow.ColumnHeaders IsNot Nothing Then
                        intHeight = GetTextHeight(selRow.ColumnHeaders(0), MaxDiffLeftRectangle.Width, True)
                        If intHeight > intRowHeight Then intRowHeight = intHeight
                        intHeight = GetTextHeight(selRow.ColumnHeaders(1), MaxDiffCenterRectangle.Width, True)
                        If intHeight > intRowHeight Then intRowHeight = intHeight
                        intHeight = GetTextHeight(selRow.ColumnHeaders(2), MaxDiffRightRectangle.Width, True)
                        If intHeight > intRowHeight Then intRowHeight = intHeight
                    Else
                        If selRow.FirstImage IsNot Nothing Then _
                            intRowHeight = selRow.FirstImage.Height

                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaValue, PrintIndicatorRow.ObjectTypes.AnswerAreaFormula, PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, _
                    PrintIndicatorRow.ObjectTypes.AnswerAreaLikertType, PrintIndicatorRow.ObjectTypes.AnswerAreaScale, PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale, _
                    PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale

                    If String.IsNullOrEmpty(selRow.TableHeader) = False Then
                        intRowHeight = GetTextHeight(selRow.TableHeader, ContentWidth, True)
                    ElseIf selRow.ColumnHeaders IsNot Nothing Then
                        For i = 0 To selRow.ColumnHeaders.Count - 1
                            intHeight = GetTextHeight(selRow.ColumnHeaders(i), selRow.ColumnWidth, True)
                            If intHeight > intRowHeight Then intRowHeight = intHeight
                        Next
                    Else
                        If selRow.FirstImage IsNot Nothing Then
                            intRowHeight = selRow.FirstImage.Height
                        Else
                            intRowHeight = GetTextHeight(selRow.RowHeader, selRow.RowHeaderWidth, True)
                        End If
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff
                    If String.IsNullOrEmpty(selRow.TableHeader) = False Then
                        intRowHeight = GetTextHeight(selRow.TableHeader, ContentWidth, True)
                    ElseIf selRow.ColumnHeaders IsNot Nothing Then
                        For i = 0 To selRow.ColumnHeaders.Count - 1
                            intHeight = GetTextHeight(selRow.ColumnHeaders(i), selRow.ColumnWidth, True)
                            If intHeight > intRowHeight Then intRowHeight = intHeight
                        Next
                    Else
                        If selRow.FirstImage IsNot Nothing Then intRowHeight = selRow.FirstImage.Height
                        If selRow.SecondImage IsNot Nothing Then
                            If selRow.SecondImage.Height > intRowHeight Then intRowHeight = selRow.SecondImage.Height
                        End If
                    End If
                Case Else
                    intRowHeight = selRow.RowHeight
            End Select
        End If

        If intRowHeight > 0 Then selRow.RowHeight = intRowHeight Else selRow.RowHeight = NewCellHeight()
    End Sub
#End Region

#Region "General methods"
    Public Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal

        For Each selRow As PrintIndicatorRow In PrintList
            intTotalHeight += selRow.RowHeight
        Next

        decPages = intTotalHeight / Me.ContentHeight
        Return Math.Ceiling(decPages)
    End Function

    Private Function GetTextHeight(ByVal strText As String, ByVal intColumnWidth As Integer, Optional ByVal boolHeader As Boolean = False) As Integer
        Dim intHeight As Integer

        If PageGraph IsNot Nothing Then
            If boolHeader = True Then
                intHeight = PageGraph.MeasureString(strText, fntTextBold, intColumnWidth).Height
            Else
                intHeight = PageGraph.MeasureString(strText, fntText, intColumnWidth).Height
            End If
        End If

        Return intHeight
    End Function

    Private Function GetTextWidth(ByVal strText As String, Optional ByVal boolHeader As Boolean = False) As Integer
        Dim intWidth As Integer

        If PageGraph IsNot Nothing Then
            If boolHeader = True Then
                intWidth = PageGraph.MeasureString(strText, fntTextBold).Width
            Else
                intWidth = PageGraph.MeasureString(strText, fntText).Width
            End If
        End If

        Return intWidth
    End Function
#End Region

#Region "Set heights"
    Public Sub SetRowHeights(ByVal graph As Graphics)

        For Each selRow As PrintIndicatorRow In PrintList
            Select Case selRow.ObjectType
                Case LogFrame.ObjectTypes.ResponseClass
                    selRow.RowHeight = GetResponseClassesHeight(selRow, graph)
                Case LogFrame.ObjectTypes.Statement
                    selRow.RowHeight = GetResponseHeight(selRow, graph)
                Case Else
                    selRow.RowHeight = GetTextHeight(Me.TextRectangle.Width, selRow.Rtf, False, graph)
            End Select
        Next
    End Sub

    Public Function GetTextHeight(ByVal intColWidth As Integer, ByVal strRtf As String, ByVal boolIsUniformText As Boolean, ByVal graph As Graphics) As Integer
        Dim intHeight As Integer

        If strRtf <> String.Empty Then
            Dim rCell As New Rectangle
            rCell.Width = intColWidth - (CONST_HorizontalPadding * 2)
            rCell.Height = NewCellHeight()
            'intHeight = RichTextManager.GetHeight(strRtf, boolIsUniformText, rCell, graph) + (VerticalPadding * 2)
        End If
        Return intHeight
    End Function

    Public Function GetResponseClassesHeight(ByVal selRow As PrintIndicatorRow, ByVal graph As Graphics) As Integer
        Dim selResponseClasses As ResponseClasses = selRow.ResponseClasses
        Dim intHeight As Integer
        Dim size As SizeF

        Select Case selRow.QuestionType
            Case Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.MultipleChoice, _
                Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.Ranking, _
                Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.LikertScale, _
                Indicator.QuestionTypes.CumulativeScale, Indicator.QuestionTypes.SemanticDiff

                Dim intColCount As Integer = selResponseClasses.Count + 2 'left label gets 2 column widths
                If selRow.QuestionType = Indicator.QuestionTypes.SemanticDiff Then _
                    intColCount += 2
                If PrintValueRanges = True Then intColCount += 1
                If PrintTargets = True Then intColCount += 1
                selRow.ResponseClassWidth = TextRectangle.Width / intColCount
                intResponseClassWidth = selRow.ResponseClassWidth

                selRow.LeftLabelWidth = selRow.ResponseClassWidth * 2
                intLeftLabelWidth = selRow.LeftLabelWidth

                If selRow.QuestionType = Indicator.QuestionTypes.SemanticDiff Then
                    selRow.RightLabelWidth = selRow.LeftLabelWidth
                    intRightLabelWidth = selRow.RightLabelWidth
                Else
                    selRow.RightLabelWidth = 0
                    intRightLabelWidth = 0
                End If

                For Each selClass As ResponseClass In selResponseClasses
                    size = graph.MeasureString(selClass.ClassName, CurrentLogFrame.DetailsFont, selRow.ResponseClassWidth)
                    If size.Height > intHeight Then intHeight = size.Height
                Next

                If PrintValueRanges = True Then
                    size = graph.MeasureString("Range", CurrentLogFrame.DetailsFont, selRow.ResponseClassWidth)
                    If size.Height > intHeight Then intHeight = size.Height
                End If
                If PrintTargets = True Then
                    size = graph.MeasureString("Target", CurrentLogFrame.DetailsFont, selRow.ResponseClassWidth)
                    If size.Height > intHeight Then intHeight = size.Height
                End If
            Case Indicator.QuestionTypes.MaxDiff
                selRow.ResponseClassWidth = Me.ContentWidth - Me.SortRectangle.Width - 200
                For Each selClass As ResponseClass In selResponseClasses
                    size = graph.MeasureString(selClass.ClassName, CurrentLogFrame.DetailsFont, 100)
                    If size.Height > intHeight Then intHeight = size.Height
                Next
        End Select

        Return intHeight
    End Function

    Public Function GetResponseHeight(ByVal selRow As PrintIndicatorRow, ByVal graph As Graphics) As Integer
        Dim selStatement As Statement = selRow.Statement

        If selStatement IsNot Nothing Then
            selRow.LeftLabelWidth = intLeftLabelWidth
            selRow.RightLabelWidth = intRightLabelWidth
            selRow.ResponseClassWidth = intResponseClassWidth

            Select Case selRow.QuestionType
                Case Indicator.QuestionTypes.OpenEnded
                    Return GetResponseHeight_OpenEnded(selRow, graph)

                Case Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.MultipleChoice, _
                    Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.Ranking, _
                    Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.LikertScale, _
                    Indicator.QuestionTypes.CumulativeScale, Indicator.QuestionTypes.SemanticDiff

                    Return GetResponseHeight_Classes(selRow, graph)

                Case Indicator.QuestionTypes.MaxDiff
                    Return GetResponseHeight_MaxDiff(selRow, graph)

                Case Indicator.QuestionTypes.AbsoluteValue, _
                    Indicator.QuestionTypes.PercentageValue

                    Return GetResponseHeight_Values(selRow, graph)
            End Select
        End If
    End Function

    Private Function GetResponseHeight_OpenEnded(ByVal selRow As PrintIndicatorRow, ByVal graph As Graphics) As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim intLeftLabelHeight, intWhiteSpaceHeight, intHeight As Integer
        Dim intLineSpacing As Integer = 30

        If selStatement IsNot Nothing Then
            intLeftLabelHeight = GetTextHeight(selRow.LeftLabelWidth, selStatement.FirstLabel, False, graph)
            intHeight = intLeftLabelHeight
            'If selStatement.Label2 <> String.Empty Then
            '    intRightLabelHeight = GetTextHeight(selRow.RightLabelWidth, selStatement.Label2, False, graph)
            '    If intRightLabelHeight > intHeight Then intHeight = intRightLabelHeight
            'End If
            'If selStatement.WhiteSpace <> Nothing Then
            '    If selStatement.WhiteSpace.Contains("l") Then
            '        intWhiteSpaceHeight = selStatement.NrLines * intLineSpacing

            '    ElseIf selStatement.WhiteSpace.Contains("p") Then
            '        intWhiteSpaceHeight = Me.ContentHeight * selStatement.Page
            '    End If
            'End If
        End If
        If intWhiteSpaceHeight > intHeight Then intHeight = intWhiteSpaceHeight

        Return intHeight
    End Function

    Private Function GetResponseHeight_Classes(ByVal selRow As PrintIndicatorRow, ByVal graph As Graphics) As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim intHeight As Integer ', intTempHeight As Integer

        If selRow.ResponseClassesNumber > 0 Then
            'determine height of the row
            If selStatement.FirstLabel IsNot Nothing Then _
                intHeight = GetTextHeight(selRow.LeftLabelWidth, selStatement.FirstLabel, False, graph)
            'If PrintRanges = True And selStatement.RangeField <> "" Then
            '    intTempHeight = graph.MeasureString(selStatement.RangeField, CurrentLogFrame.DetailsFont, selRow.ResponseClassWidth).Height
            '    If intTempHeight > intHeight Then intHeight = intTempHeight
            'End If
            'If PrintTargets = True And selStatement.TargetField <> "" Then
            '    intTempHeight = graph.MeasureString(selStatement.TargetField, CurrentLogFrame.DetailsFont, selRow.ResponseClassWidth).Height
            '    If intTempHeight > intHeight Then intHeight = intTempHeight
            'End If
        End If
        Return intHeight
    End Function

    Private Function GetResponseHeight_MaxDiff(ByVal selRow As PrintIndicatorRow, ByVal graph As Graphics) As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim intHeight As Integer

        If selRow.ResponseClassesNumber > 0 Then
            If selStatement.FirstLabel <> String.Empty Then _
                intHeight = GetTextHeight(selRow.ResponseClassWidth, selStatement.FirstLabel, False, graph)
        End If

        Return intHeight
    End Function

    Private Function GetResponseHeight_Values(ByVal selRow As PrintIndicatorRow, ByVal graph As Graphics) As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim intWidthOfSpacing As Integer = 10
        Dim intWidthOfWhiteSpace As Integer = 150
        Dim intWidthOfUnits As Integer = 100
        Dim intHeight As Integer

        selRow.LeftLabelWidth = Me.ContentWidth - Me.SortRectangle.Width - intWidthOfSpacing - intWidthOfWhiteSpace - intWidthOfUnits
        If selRow.Rtf <> String.Empty Then _
            intHeight = GetTextHeight(selRow.LeftLabelWidth, selRow.Rtf, False, graph)

        Return intHeight
    End Function
#End Region

#Region "Print page"
    Protected Overrides Sub OnBeginPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnBeginPrint(e)
        boolColumnsWidthSet = False
        sfColumnText.Alignment = StringAlignment.Center
        sfColumnText.LineAlignment = StringAlignment.Center

        boolColumnsWidthSet = False
        PrintRectangles.Clear()
        ClippedRow = Nothing

        LastRowY = ContentTop
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
            'create print list is done here because PageGraph is needed to determine width of table header cells
            CreateList()

            SetColumnsWidth()
            ReloadImages()
            Me.TotalPages = GetTotalPages()

            boolColumnsWidthSet = True
        End If

        Dim intRowCount As Integer = PrintList.Count
        Dim intMinHeight As Integer
        Dim selRow As PrintIndicatorRow

        'Print Header
        PrintHeader()

        Do While RowIndex <= PrintList.Count - 1
            selRow = PrintList(RowIndex)

            intMinHeight = PrintPage_HeightToAvoidOrphans(selRow)

            If intMinHeight < Me.ContentBottom Then
                PrintPage_PrintRow(selRow)

                If selRow.ClippedRow = True Then ClippedRow = Nothing
            Else
                'set the minimum height to the line spacing of the first line of the row
                intMinHeight = PrintPage_GetMinHeight(selRow)

                If intMinHeight < Me.ContentBottom Then
                    PrintPage_PrintRow(selRow)

                    If ClippedRow IsNot Nothing Then
                        PrintList.Insert(RowIndex, ClippedRow)

                        ReloadImages_Normal(ClippedRow)
                        SetRowHeight(RowIndex)
                    End If

                    PrintFooter()

                    LastRowY = ContentTop
                    PageNumber += 1
                    e.HasMorePages = True
                    Exit Do
                Else
                    'else go to a new page and print the line there
                    PrintFooter()

                    LastRowY = ContentTop
                    PageNumber += 1
                    e.HasMorePages = True
                    Exit Do
                End If
            End If

            If RowIndex > PrintList.Count - 1 Then
                PrintFooter()
                e.HasMorePages = False
            End If
        Loop
        If PrintList.Count = 0 Then RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(0, 0))
    End Sub

    Private Sub PrintPage_PrintRow(ByVal selRow As PrintIndicatorRow)
        Select Case selRow.ObjectType
            Case PrintIndicatorRow.ObjectTypes.Goal, PrintIndicatorRow.ObjectTypes.Purpose, PrintIndicatorRow.ObjectTypes.Output, PrintIndicatorRow.ObjectTypes.Activity, _
                PrintIndicatorRow.ObjectTypes.Indicator
                If selRow.FirstImage IsNot Nothing Then
                    Dim rImage As New Rectangle(TextRectangle.Left, LastRowY, TextRectangle.Width, selRow.FirstImage.Height)

                    PrintPage_PrintSortNumber(selRow)
                    PrintPage_PrintRtf(selRow, rImage)
                End If
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaOpenEnded
                PrintAnswerArea_OpenEnded(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff
                PrintAnswerArea_MaxDiff(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaValue
                PrintAnswerArea_Value(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaFormula
                PrintAnswerArea_Formula(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, PrintIndicatorRow.ObjectTypes.AnswerAreaLikertType
                PrintAnswerArea_MultipleOptions(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff
                PrintAnswerArea_SemanticDiff(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaScale, PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale
                PrintAnswerArea_Scale(selRow)
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale
                PrintAnswerArea_LikertScale(selRow)
        End Select

        RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(RowIndex, PrintList.Count))
        LastRowY += selRow.RowHeight
        RowIndex += 1
    End Sub

    Private Sub PrintPage_PrintSortNumber(ByVal selRow As PrintIndicatorRow)
        Dim formatCells As New StringFormat()
        Dim brText As SolidBrush = New SolidBrush(Color.Black)
        Dim rCell As New Rectangle(SortRectangle.Left, LastRowY + CONST_VerticalPadding, SortRectangle.Width, selRow.RowHeight)

        If PageGraph IsNot Nothing Then
            Select Case selRow.ObjectType
                Case LogFrame.ObjectTypes.Goal, LogFrame.ObjectTypes.Purpose, LogFrame.ObjectTypes.Output, LogFrame.ObjectTypes.Activity, LogFrame.ObjectTypes.Indicator
                    formatCells.Alignment = StringAlignment.Near
                    formatCells.LineAlignment = StringAlignment.Near
            End Select

            PageGraph.DrawString(selRow.SortNumber, CurrentLogFrame.SortNumberFont, brText, rCell, formatCells)
        End If
    End Sub

    Private Sub PrintPage_PrintRtf(ByVal selRow As PrintIndicatorRow, ByVal rImage As Rectangle)

        If PageGraph IsNot Nothing And String.IsNullOrEmpty(selRow.Rtf) = False Then

            If selRow.FirstImage IsNot Nothing Then
                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.FirstImage, rImage)
                Else
                    PrintClippedText(selRow.Rtf, rImage, TextRectangle.Width)
                    selRow.Rtf = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintIndicatorRow(selRow.ObjectType, String.Empty, strClippedTextBottom)
                        ClippedRow.ClippedRow = True
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub PrintBackGround(ByVal intTextHeight As Integer, ByVal intObjectType As Integer, ByVal graph As Graphics)
        Dim rFill As New Rectangle(LeftMargin, LastRowY, Me.PrintWidth, intTextHeight)
        If intObjectType > 0 Then
            Select Case intObjectType
                Case LogFrame.ObjectTypes.Goal, LogFrame.ObjectTypes.Purpose
                    Dim objColor As Color = Color.FromArgb(120, 120, 120)
                    Dim brFill As New SolidBrush(objColor)

                    graph.FillRectangle(brFill, rFill)
                    graph.DrawRectangle(penBlack1, rFill)
                Case LogFrame.ObjectTypes.Output
                    Dim objColor As Color = Color.FromArgb(160, 160, 160)
                    Dim brFill As New SolidBrush(objColor)

                    graph.FillRectangle(brFill, rFill)
                    graph.DrawRectangle(penBlack1, rFill)
                Case LogFrame.ObjectTypes.Activity
                    Dim objColor As Color = Color.FromArgb(200, 200, 200)
                    Dim brFill As New SolidBrush(objColor)

                    graph.FillRectangle(brFill, rFill)
                    graph.DrawRectangle(penBlack1, rFill)
            End Select
        End If
    End Sub

    Private Function PrintPage_GetMinHeight(ByVal selRow As PrintIndicatorRow)
        Dim intMinHeight As Integer

        With RichTextManager
            Select Case selRow.ObjectType
                Case PrintIndicatorRow.ObjectTypes.Goal, PrintIndicatorRow.ObjectTypes.Purpose, PrintIndicatorRow.ObjectTypes.Output, PrintIndicatorRow.ObjectTypes.Activity, _
                    PrintIndicatorRow.ObjectTypes.Indicator

                    If String.IsNullOrEmpty(selRow.Rtf) = False Then
                        intMinHeight = .GetFirstLineSpacing(TextRectangle.Width, selRow.Rtf)
                        intMinHeight += LastRowY
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaOpenEnded
                    intMinHeight = LastRowY + CONST_OpenEndedLineSpacing
                Case Else
                    intMinHeight = PrintPage_HeightToAvoidOrphans(selRow)
            End Select
        End With

        Return intMinHeight
    End Function

    Private Function PrintPage_HeightToAvoidOrphans(ByVal selRow As PrintIndicatorRow) As Integer
        Dim intMinHeight As Integer

        Select Case selRow.ObjectType
            Case PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, PrintIndicatorRow.ObjectTypes.AnswerAreaValue
                If String.IsNullOrEmpty(selRow.TableHeader) = False And RowIndex < PrintList.Count - 2 Then
                    intMinHeight = selRow.RowHeight + PrintList(RowIndex + 1).RowHeight + PrintList(RowIndex + 2).RowHeight
                ElseIf selRow.ColumnHeaders IsNot Nothing And RowIndex < PrintList.Count - 1 Then
                    intMinHeight = selRow.RowHeight + PrintList(RowIndex + 1).RowHeight
                Else
                    intMinHeight = selRow.RowHeight
                End If

                intMinHeight += LastRowY
            Case Else
                intMinHeight = selRow.RowHeight + LastRowY
        End Select

        Return intMinHeight
    End Function
#End Region

#Region "Clipped text"
    Private Sub PrintClippedText(ByVal strRtf As String, ByVal rImage As Rectangle, intColumnWidth As Integer)
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

#Region "Answer areas"
    Private Sub PrintAnswerArea_OpenEnded(ByVal selRow As PrintIndicatorRow)
        If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.OpenEndedDetail IsNot Nothing Then
            Dim selOpenEndedDetail As OpenEndedDetail = selRow.Indicator.OpenEndedDetail
            Dim LinesRectangle As New Rectangle(LeftMargin, LastRowY, ContentWidth, selRow.RowHeight)
            Dim intY As Integer = LastRowY + CONST_OpenEndedLineSpacing - 1
            Dim penGrayDotted As New Pen(Color.DarkGray, 1)
            Dim boolFirst As Boolean = True

            penGrayDotted.DashStyle = DashStyle.Dot

            Do Until intY > LinesRectangle.Bottom
                If intY <= ContentBottom Then
                    PageGraph.DrawLine(penGrayDotted, LinesRectangle.Left, intY, LinesRectangle.Right, intY)

                    If String.IsNullOrEmpty(selRow.TargetDeadlineDate) = False And boolFirst = True Then
                        Dim strDate As String = ToLabel(selRow.TargetDeadlineDate)
                        Dim intWidth As Integer = PageGraph.MeasureString(strDate, fntTextBold).Width
                        Dim rDate As New Rectangle(LinesRectangle.Left, LinesRectangle.Top + CONST_VerticalPadding, SortRectangle.Width, CONST_OpenEndedLineSpacing - CONST_VerticalPadding)

                        PageGraph.FillRectangle(Brushes.White, rDate)
                        PageGraph.DrawString(strDate, fntTextBold, Brushes.Black, rDate)
                        boolFirst = False
                    End If
                Else
                    ClippedRow = New PrintIndicatorRow(selRow.ObjectType, selRow.Indicator)
                    ClippedRow.RowHeight = LinesRectangle.Bottom - intY
                    ClippedRow.ClippedRow = True

                    Exit Do
                End If
                intY += CONST_OpenEndedLineSpacing
            Loop
        End If
    End Sub

    Private Sub PrintAnswerArea_MaxDiff(ByVal selRow As PrintIndicatorRow)
        Dim rCheckLeft As New Rectangle(MaxDiffLeftRectangle.Left, LastRowY, MaxDiffLeftRectangle.Width, selRow.RowHeight)
        Dim rStatement As New Rectangle(MaxDiffCenterRectangle.Left, LastRowY, MaxDiffCenterRectangle.Width, selRow.RowHeight)
        Dim rCheckRight As New Rectangle(MaxDiffRightRectangle.Left, LastRowY, MaxDiffRightRectangle.Width, selRow.RowHeight)
        Dim objFormat As New StringFormat

        If String.IsNullOrEmpty(selRow.TableHeader) = False Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Far

            Dim rTitle As New Rectangle(LeftMargin, LastRowY, ContentWidth, selRow.RowHeight)
            PageGraph.DrawString(selRow.TableHeader, fntTextBold, Brushes.Black, rTitle, objFormat)
        ElseIf selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            PageGraph.FillRectangle(Brushes.LightGray, rCheckLeft)
            PageGraph.FillRectangle(Brushes.LightGray, rStatement)
            PageGraph.FillRectangle(Brushes.LightGray, rCheckRight)
            PageGraph.DrawString(selRow.ColumnHeaders(0), fntTextBold, Brushes.Black, rCheckLeft, objFormat)
            PageGraph.DrawString(selRow.ColumnHeaders(1), fntTextBold, Brushes.Black, rStatement, objFormat)
            PageGraph.DrawString(selRow.ColumnHeaders(2), fntTextBold, Brushes.Black, rCheckRight, objFormat)
        ElseIf selRow.FirstImage IsNot Nothing Then
            PrintPage_PrintRtf(selRow, rStatement)
        End If

        If String.IsNullOrEmpty(selRow.TableHeader) Then
            PageGraph.DrawRectangle(penBlack1, rCheckLeft)
            PageGraph.DrawRectangle(penBlack1, rStatement)
            PageGraph.DrawRectangle(penBlack1, rCheckRight)
        End If
    End Sub

    Private Sub PrintAnswerArea_Value(ByVal selRow As PrintIndicatorRow)
        Dim rRowHeader As New Rectangle(LeftMargin, LastRowY, selRow.RowHeaderWidth, selRow.RowHeight)
        Dim rColumn As Rectangle
        Dim intX As Integer = rRowHeader.Right
        Dim objFormat As New StringFormat

        If selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            For i = 0 To selRow.ColumnHeaders.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.FillRectangle(Brushes.LightGray, rColumn)
                PageGraph.DrawString(selRow.ColumnHeaders(i), fntTextBold, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        ElseIf selRow.Cells IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Center

            PageGraph.FillRectangle(Brushes.LightGray, rRowHeader)
            PageGraph.DrawString(selRow.RowHeader, fntText, Brushes.Black, rRowHeader, objFormat)
            PageGraph.DrawRectangle(penBlack1, rRowHeader)

            objFormat.Alignment = StringAlignment.Far

            For i = 0 To selRow.Cells.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.DrawString(selRow.Cells(i), fntText, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        End If
    End Sub

    Private Sub PrintAnswerArea_Formula(ByVal selRow As PrintIndicatorRow)
        Dim rStatement As New Rectangle(LeftMargin, LastRowY, selRow.RowHeaderWidth, selRow.RowHeight)
        Dim rColumn As Rectangle
        Dim intX As Integer = rStatement.Right
        Dim objFormat As New StringFormat

        If selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            For i = 0 To selRow.ColumnHeaders.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.FillRectangle(Brushes.LightGray, rColumn)
                PageGraph.DrawString(selRow.ColumnHeaders(i), fntTextBold, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        ElseIf selRow.Cells IsNot Nothing Then
            PageGraph.FillRectangle(Brushes.LightGray, rStatement)
            If selRow.FirstImage IsNot Nothing Then
                PageGraph.DrawImage(selRow.FirstImage, New Point(rStatement.Left, rStatement.Top))
            End If
            PageGraph.DrawRectangle(penBlack1, rStatement)

            objFormat.Alignment = StringAlignment.Far

            For i = 0 To selRow.Cells.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.DrawString(selRow.Cells(i), fntText, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        End If
    End Sub

    Private Sub PrintAnswerArea_MultipleOptions(ByVal selRow As PrintIndicatorRow)
        Dim rOption As New Rectangle(LeftMargin, LastRowY, selRow.RowHeaderWidth, selRow.RowHeight)
        Dim rColumn, rCheck As Rectangle
        Dim intX As Integer = rOption.Right
        Dim objFormat As New StringFormat

        If String.IsNullOrEmpty(selRow.TableHeader) = False Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Far

            Dim rTitle As New Rectangle(LeftMargin, LastRowY, ContentWidth, selRow.RowHeight)
            PageGraph.DrawString(selRow.TableHeader, fntText, Brushes.Black, rTitle, objFormat)
        ElseIf selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            For i = 0 To selRow.ColumnHeaders.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.FillRectangle(Brushes.LightGray, rColumn)
                PageGraph.DrawString(selRow.ColumnHeaders(i), fntTextBold, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        ElseIf selRow.Cells IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Center

            PageGraph.FillRectangle(Brushes.LightGray, rOption)
            PageGraph.DrawString(selRow.RowHeader, fntText, Brushes.Black, rOption, objFormat)
            PageGraph.DrawRectangle(penBlack1, rOption)

            For i = 0 To selRow.Cells.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                rCheck = New Rectangle(rColumn.Left + (rColumn.Width / 2) - (CONST_CheckBoxIconSize / 2), rColumn.Top + (rColumn.Height / 2) - (CONST_CheckBoxIconSize / 2), CONST_CheckBoxIconSize, CONST_CheckBoxIconSize)
                If selRow.NoCheckBox = False Then PageGraph.DrawRectangle(penBlack1, rCheck)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        End If
    End Sub

    Private Sub PrintAnswerArea_SemanticDiff(ByVal selRow As PrintIndicatorRow)
        Dim rFirstLabel As New Rectangle(LeftMargin, LastRowY, selRow.RowHeaderWidth, selRow.RowHeight)
        Dim rColumn, rCheck As Rectangle
        Dim intX As Integer = rFirstLabel.Right
        Dim objFormat As New StringFormat

        If String.IsNullOrEmpty(selRow.TableHeader) = False Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Far

            Dim rTitle As New Rectangle(LeftMargin, LastRowY, ContentWidth, selRow.RowHeight)
            PageGraph.DrawString(selRow.TableHeader, fntText, Brushes.Black, rTitle, objFormat)
        ElseIf selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            For i = 0 To selRow.ColumnHeaders.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.FillRectangle(Brushes.LightGray, rColumn)
                PageGraph.DrawString(selRow.ColumnHeaders(i), fntTextBold, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        ElseIf selRow.Cells IsNot Nothing Then
            PageGraph.FillRectangle(Brushes.LightGray, rFirstLabel)
            If selRow.FirstImage IsNot Nothing Then
                PageGraph.DrawImage(selRow.FirstImage, New Point(rFirstLabel.Left, rFirstLabel.Top))
            End If
            PageGraph.DrawRectangle(penBlack1, rFirstLabel)

            For i = 0 To selRow.Cells.Count - 2
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                rCheck = New Rectangle(rColumn.Left + (rColumn.Width / 2) - (CONST_CheckBoxIconSize / 2), rColumn.Top + (rColumn.Height / 2) - (CONST_CheckBoxIconSize / 2), CONST_CheckBoxIconSize, CONST_CheckBoxIconSize)
                If selRow.NoCheckBox = False Then PageGraph.DrawRectangle(penBlack1, rCheck)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next

            Dim rSecondLabel As New Rectangle(intX, LastRowY, selRow.RowHeaderWidth, selRow.RowHeight)
            PageGraph.FillRectangle(Brushes.LightGray, rSecondLabel)
            If selRow.SecondImage IsNot Nothing Then
                PageGraph.DrawImage(selRow.SecondImage, New Point(rSecondLabel.Left, rSecondLabel.Top))
            End If
            PageGraph.DrawRectangle(penBlack1, rSecondLabel)
        End If
    End Sub

    Private Sub PrintAnswerArea_Scale(ByVal selRow As PrintIndicatorRow)
        Dim rStatement As New Rectangle(ScaleStatementRectangle.Left, LastRowY, ScaleStatementRectangle.Width, selRow.RowHeight)
        Dim rCheckAgree As New Rectangle(ScaleAgreeRectangle.Left, LastRowY, ScaleAgreeRectangle.Width, selRow.RowHeight)
        Dim rCheckDisagree As New Rectangle(ScaleDisagreeRectangle.Left, LastRowY, ScaleDisagreeRectangle.Width, selRow.RowHeight)
        Dim rCheck As Rectangle
        Dim objFormat As New StringFormat

        If String.IsNullOrEmpty(selRow.TableHeader) = False Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Far

            Dim rTitle As New Rectangle(LeftMargin, LastRowY, ContentWidth, selRow.RowHeight)
            PageGraph.DrawString(selRow.TableHeader, fntTextBold, Brushes.Black, rTitle, objFormat)
        ElseIf selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            PageGraph.FillRectangle(Brushes.LightGray, rStatement)
            PageGraph.FillRectangle(Brushes.LightGray, rCheckAgree)
            PageGraph.FillRectangle(Brushes.LightGray, rCheckDisagree)

            PageGraph.DrawString(selRow.ColumnHeaders(0), fntTextBold, Brushes.Black, rStatement, objFormat)
            PageGraph.DrawString(selRow.ColumnHeaders(1), fntTextBold, Brushes.Black, rCheckAgree, objFormat)
            PageGraph.DrawString(selRow.ColumnHeaders(2), fntTextBold, Brushes.Black, rCheckDisagree, objFormat)
        ElseIf selRow.FirstImage IsNot Nothing Then
            PrintPage_PrintRtf(selRow, rStatement)
        End If

        If String.IsNullOrEmpty(selRow.TableHeader) Then
            PageGraph.DrawRectangle(penBlack1, rStatement)
            PageGraph.DrawRectangle(penBlack1, rCheckAgree)
            PageGraph.DrawRectangle(penBlack1, rCheckDisagree)

            If selRow.ColumnHeaders Is Nothing And selRow.NoCheckBox = False Then
                rCheck = New Rectangle(rCheckAgree.Left + (rCheckAgree.Width / 2) - (CONST_CheckBoxIconSize / 2), rCheckAgree.Top + (rCheckAgree.Height / 2) - (CONST_CheckBoxIconSize / 2), CONST_CheckBoxIconSize, CONST_CheckBoxIconSize)
                PageGraph.DrawRectangle(penBlack1, rCheck)
                rCheck = New Rectangle(rCheckDisagree.Left + (rCheckDisagree.Width / 2) - (CONST_CheckBoxIconSize / 2), rCheckDisagree.Top + (rCheckDisagree.Height / 2) - (CONST_CheckBoxIconSize / 2), CONST_CheckBoxIconSize, CONST_CheckBoxIconSize)
                PageGraph.DrawRectangle(penBlack1, rCheck)
            End If
        End If
    End Sub

    Private Sub PrintAnswerArea_LikertScale(ByVal selRow As PrintIndicatorRow)
        Dim rStatement As New Rectangle(LeftMargin, LastRowY, selRow.RowHeaderWidth, selRow.RowHeight)
        Dim rColumn, rCheck As Rectangle
        Dim intX As Integer = rStatement.Right
        Dim objFormat As New StringFormat

        If String.IsNullOrEmpty(selRow.TableHeader) = False Then
            objFormat.Alignment = StringAlignment.Near
            objFormat.LineAlignment = StringAlignment.Far

            Dim rTitle As New Rectangle(LeftMargin, LastRowY, ContentWidth, selRow.RowHeight)
            PageGraph.DrawString(selRow.TableHeader, fntText, Brushes.Black, rTitle, objFormat)
        ElseIf selRow.ColumnHeaders IsNot Nothing Then
            objFormat.Alignment = StringAlignment.Center
            objFormat.LineAlignment = StringAlignment.Center

            For i = 0 To selRow.ColumnHeaders.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                PageGraph.FillRectangle(Brushes.LightGray, rColumn)
                PageGraph.DrawString(selRow.ColumnHeaders(i), fntTextBold, Brushes.Black, rColumn, objFormat)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        ElseIf selRow.Cells IsNot Nothing Then
            PageGraph.FillRectangle(Brushes.LightGray, rStatement)
            If selRow.FirstImage IsNot Nothing Then
                PageGraph.DrawImage(selRow.FirstImage, New Point(rStatement.Left, rStatement.Top))
            End If
            PageGraph.DrawRectangle(penBlack1, rStatement)

            For i = 0 To selRow.Cells.Count - 1
                rColumn = New Rectangle(intX, LastRowY, selRow.ColumnWidth, selRow.RowHeight)
                rCheck = New Rectangle(rColumn.Left + (rColumn.Width / 2) - (CONST_CheckBoxIconSize / 2), rColumn.Top + (rColumn.Height / 2) - (CONST_CheckBoxIconSize / 2), CONST_CheckBoxIconSize, CONST_CheckBoxIconSize)
                If selRow.NoCheckBox = False Then PageGraph.DrawRectangle(penBlack1, rCheck)
                PageGraph.DrawRectangle(penBlack1, rColumn)
                intX = rColumn.Right
            Next
        End If
    End Sub
#End Region

End Class

