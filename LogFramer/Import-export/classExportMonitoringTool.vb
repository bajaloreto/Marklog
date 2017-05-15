Public Class ExportMonitoringTool
    Implements System.IDisposable

    'Private objPrintIndicators As PrintIndicators
    Private objTableGrid As New ExportIndicatorRows
    Private objLogFrame As New LogFrame
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroupGuid As Guid
    Private intSection As Integer
    Private intMeasurement As Integer
    Private boolExportGoals, boolExportPurposes As Boolean
    Private boolExportOutputs, boolExportActivities As Boolean
    Private boolExportOptionValues, boolExportRanges, boolExportTargets As Boolean

    Public Enum ExportSections As Integer
        NotSelected = 0
        ExportGoals = 1
        ExportPurposes = 2
        ExportOutputs = 3
        ExportActivities = 4
        ExportAll = 5
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, ByVal exportsection As Integer, ByVal targetgroupguid As Guid, ByVal measurement As Integer, _
                   ByVal exportpurposes As Boolean, ByVal exportoutputs As Boolean, ByVal exportactivities As Boolean, ByVal exportoptionvalues As Boolean, _
                   ByVal exportranges As Boolean, ByVal exporttargets As Boolean)
        Me.LogFrame = logframe
        Me.ExportSection = exportsection
        Me.TargetGroupGuid = targetgroupguid
        Me.Measurement = measurement
        Me.ExportPurposes = exportpurposes
        Me.ExportOutputs = exportoutputs
        Me.ExportActivities = exportactivities
        Me.ExportOptionValues = exportoptionvalues
        Me.ExportValueRanges = exportranges
        Me.ExportTargets = exporttargets

        Me.TargetDeadlinesSection = logframe.GetTargetDeadlinesSection(exportsection)
    End Sub

#Region "Properties"
    Public Property TableGrid As ExportIndicatorRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportIndicatorRows)
            objTableGrid = value
        End Set
    End Property

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

    Public ReadOnly Property TargetDeadlines As TargetDeadlines
        Get
            If objTargetDeadlinesSection IsNot Nothing Then
                Return objTargetDeadlinesSection.TargetDeadlines
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property ExportSection() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
            Select Case intSection
                Case ExportSections.ExportGoals
                    boolExportGoals = True
                Case ExportSections.ExportPurposes
                    boolExportPurposes = True
                Case ExportSections.ExportOutputs
                    boolExportOutputs = True
                Case ExportSections.ExportActivities
                    boolExportActivities = True
                Case ExportSections.ExportAll
                    boolExportGoals = True
                    boolExportPurposes = True
                    boolExportOutputs = True
                    boolExportActivities = True
            End Select
        End Set
    End Property

    Private Property Measurement As Integer
        Get
            Return intMeasurement
        End Get
        Set(ByVal value As Integer)
            intMeasurement = value
        End Set
    End Property

    Public Property ExportPurposes As Boolean
        Get
            Return boolExportPurposes
        End Get
        Set(ByVal value As Boolean)
            boolExportPurposes = value
        End Set
    End Property

    Public Property ExportOutputs As Boolean
        Get
            Return boolExportOutputs
        End Get
        Set(ByVal value As Boolean)
            boolExportOutputs = value
        End Set
    End Property

    Public Property ExportActivities As Boolean
        Get
            Return boolExportActivities
        End Get
        Set(ByVal value As Boolean)
            boolExportActivities = value
        End Set
    End Property

    Public Property ExportOptionValues As Boolean
        Get
            Return boolExportOptionValues
        End Get
        Set(ByVal value As Boolean)
            boolExportOptionValues = value
        End Set
    End Property

    Public Property ExportValueRanges As Boolean
        Get
            Return boolExportRanges
        End Get
        Set(ByVal value As Boolean)
            boolExportRanges = value
        End Set
    End Property

    Public Property ExportTargets As Boolean
        Get
            Return boolExportTargets
        End Get
        Set(ByVal value As Boolean)
            boolExportTargets = value
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
#End Region

#Region "Create list of objectives, indicators and statements"
    Public Sub LoadTable()
        objTableGrid.Clear()

        'add goals and their indicators
        If Me.ExportSection = ExportSections.ExportGoals Then
            CreateList_Goals()
        ElseIf Me.ExportSection > ExportSections.ExportGoals Then
            CreateList_Purposes()
        End If
    End Sub

    Private Sub CreateList_Goals()
        Dim strSortNumber As String

        For Each selGoal As Goal In Me.LogFrame.Goals
            If selGoal.Indicators.Count > 0 Then
                strSortNumber = Me.LogFrame.GetStructSortNumber(selGoal)
                Dim objRow As New ExportIndicatorRow(PrintIndicatorRow.ObjectTypes.Goal, strSortNumber, selGoal)
                TableGrid.Add(objRow)

                If Me.ExportSection = ExportSections.ExportGoals Or Me.ExportSection = ExportSections.ExportAll Then
                    CreateList_Indicators(selGoal.Indicators, strSortNumber)
                End If
            End If
        Next
    End Sub

    Private Sub CreateList_Purposes()
        Dim strSortNumber As String

        For Each selPurpose As Purpose In Me.LogFrame.Purposes
            strSortNumber = Me.LogFrame.GetStructSortNumber(selPurpose)
            If ExportPurposes = True Then
                Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.Purpose, strSortNumber, selPurpose)
                TableGrid.Add(objRow)
            End If
            If Me.ExportSection = ExportSections.ExportPurposes Then
                CreateList_Indicators(selPurpose.Indicators, strSortNumber)
            End If

            'add outputs
            If Me.ExportSection > ExportSections.ExportPurposes Then
                CreateList_Outputs(selPurpose.Outputs)
            End If
        Next
    End Sub

    Private Sub CreateList_Outputs(ByVal selOutputs As Outputs)
        Dim strSortNumber As String

        For Each selOutput As Output In selOutputs
            strSortNumber = Me.LogFrame.GetStructSortNumber(selOutput)
            If ExportOutputs = True Then
                Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.Output, strSortNumber, selOutput)
                objTableGrid.Add(objRow)
            End If
            If Me.ExportSection = ExportSections.ExportOutputs Then
                CreateList_Indicators(selOutput.Indicators, strSortNumber)
            End If

            'add Activities
            If Me.ExportSection > ExportSections.ExportOutputs Then
                CreateList_Activities(selOutput.Activities, strSortNumber)
            End If
        Next
    End Sub

    Private Sub CreateList_Activities(ByVal selActivities As Activities, ByVal strParentSort As String)
        Dim strActivitySort As String
        Dim intIndex As Integer

        For Each selActivity As Activity In selActivities
            strActivitySort = LogFrame.CreateSortNumber(intIndex, strParentSort)

            'If selActivity.Indicators.Count > 0 Then
            If ExportActivities = True Then
                Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.Activity, strActivitySort, selActivity)
                TableGrid.Add(objRow)
            End If
            If Me.ExportSection = ExportSections.ExportActivities Then
                CreateList_Indicators(selActivity.Indicators, strActivitySort)
            End If
            'End If

            If selActivity.Activities.Count > 0 Then
                CreateList_Activities(selActivity.Activities, strActivitySort)
            End If

            intIndex += 1
        Next
    End Sub

    Private Sub CreateList_Indicators(ByVal selIndicators As Indicators, ByVal strParentSort As String)
        Dim strIndicatorSort As String
        Dim intIndex As Integer
        Dim intRowIndex As Integer

        For Each selIndicator As Indicator In selIndicators
            If Me.TargetGroupGuid = Guid.Empty OrElse selIndicator.TargetGroupGuid = Me.TargetGroupGuid Then
                'indicator number and text
                strIndicatorSort = LogFrame.CreateSortNumber(intIndex, strParentSort)
                Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.Indicator, strIndicatorSort, selIndicator)
                objTableGrid.Add(objRow)

                'answer area (text lines, table...)
                If selIndicator.Indicators.Count = 0 Then
                    CreateList_AnswerArea(selIndicator)
                Else
                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.Image, Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Ranking
                            'no totals
                        Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.PercentageValue, _
                            Indicator.QuestionTypes.Ratio
                            Dim intTargetCount As Integer = selIndicator.Targets.Count
                            Dim strCells(intTargetCount * 3) As String
                            Dim strTotal As String = AggregationString(selIndicator)

                            With objRow
                                .ScoringSystem = selIndicator.ScoringSystem
                                .Unit = selIndicator.ValuesDetail.Unit

                                If selIndicator.ValuesDetail.ValueRange.ValueRangeSet Then
                                    .ValueRangeSet = True
                                    .ValueRangeMin = selIndicator.ValuesDetail.ValueRange.MinValue
                                    .ValueRangeMax = selIndicator.ValuesDetail.ValueRange.MaxValue
                                End If
                                .NrDecimals = selIndicator.ValuesDetail.NrDecimals

                                .BaselineValue = strTotal
                                If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                                    .BaselineScore = strTotal
                                    .MaximumScore = selIndicator.GetTargetMaximumScore
                                End If

                                For i = 0 To intTargetCount - 1
                                    intIndex = (i * 3)

                                    strCells(intIndex) = strTotal
                                    If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                                        strCells(intIndex + 1) = strTotal
                                    ElseIf selIndicator.ScoringSystem = Indicator.ScoringSystems.Percentage Then
                                        intRowIndex = TableGrid.Count
                                        strCells(intIndex + 1) = CreateExcelFormulaPercentageScore(intIndex + 10, intRowIndex - 1)
                                    End If
                                    strCells(intIndex + 2) = strTotal
                                Next
                                .Cells = strCells
                            End With
                        Case Else
                            Dim intTargetCount As Integer = selIndicator.Targets.Count
                            Dim strCells(intTargetCount * 3) As String
                            Dim strTotal As String = AggregationString(selIndicator)

                            With objRow
                                .ScoringSystem = selIndicator.ScoringSystem
                                .NrDecimals = selIndicator.ValuesDetail.NrDecimals
                                .ScoringLegend = strTotal
                                .BaselineScore = strTotal
                                .MaximumScore = selIndicator.GetTargetMaximumScore

                                For i = 0 To intTargetCount - 1
                                    intIndex = (i * 3)

                                    strCells(intIndex + 1) = strTotal
                                    strCells(intIndex + 2) = strTotal
                                Next
                                .Cells = strCells
                            End With
                    End Select
                    CreateList_Indicators(selIndicator.Indicators, strIndicatorSort)
                End If
            End If
            intIndex += 1
        Next
    End Sub
#End Region

#Region "Create answer area"
    Private Sub CreateList_AnswerArea(ByVal selIndicator As Indicator)
        Select Case selIndicator.QuestionType
            Case Indicator.QuestionTypes.OpenEnded
                'do nothing
            Case Indicator.QuestionTypes.MaxDiff
                CreateList_AnswerArea_MaxDiff(selIndicator)
            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                CreateList_AnswerArea_Value(selIndicator)
            Case Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.Ratio
                CreateList_AnswerArea_Formula(selIndicator)
            Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.LikertTypeScale
                CreateList_AnswerArea_MultipleOptions(selIndicator)
            Case Indicator.QuestionTypes.Ranking
                CreateList_AnswerArea_Ranking(selIndicator)
            Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert, Indicator.QuestionTypes.CumulativeScale, Indicator.QuestionTypes.SemanticDiff
                CreateList_AnswerArea_Scale(selIndicator)
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

    Private Sub CreateList_AnswerArea_MaxDiff(ByVal selIndicator As Indicator)
        Dim strColumns() As String = New String() {selIndicator.ScalesDetail.AgreeText, LANG_Options, selIndicator.ScalesDetail.DisagreeText}

        If selIndicator.Indicators.Count = 0 Then
            For Each selStatement As Statement In selIndicator.Statements
                Dim objRow As New ExportIndicatorRow(PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff, selStatement)
                TableGrid.Add(objRow)
            Next
        End If
    End Sub

    Private Sub CreateList_AnswerArea_Value(ByVal selIndicator As Indicator)
        Dim strRowHeader As String = String.Empty
        Dim intTargetCount As Integer = selIndicator.Targets.Count
        Dim strCells(intTargetCount * 3) As String
        Dim strUnit As String
        Dim sngValueRangeMin, sngValueRangeMax As Single
        Dim boolValueRangeSet As Boolean
        Dim intNrDecimals As Integer
        Dim strScoringLegend As String = String.Empty
        Dim intScoringSystem As Integer
        Dim dblMaximumScore As Double
        Dim objBaselineValue, objBaselineScore As Object
        Dim intIndex As Integer
        Dim intRowIndex As Integer = TableGrid.Count

        If Me.ExportValueRanges And selIndicator.ValuesDetail.ValueRange.ValueRangeSet = True Then
            strRowHeader = String.Format("{0}:{1}{2}", LANG_ValuesWithinRange, vbCrLf, selIndicator.ValuesDetail.DisplayField(selIndicator.ValuesDetail.ValueRange))
        Else
            strRowHeader = ToLabel(LANG_Values)
        End If

        strUnit = selIndicator.ValuesDetail.Unit
        sngValueRangeMin = selIndicator.ValuesDetail.ValueRange.MinValue
        sngValueRangeMax = selIndicator.ValuesDetail.ValueRange.MaxValue
        boolValueRangeSet = selIndicator.ValuesDetail.ValueRange.ValueRangeSet
        intNrDecimals = selIndicator.ValuesDetail.NrDecimals

        intScoringSystem = selIndicator.ScoringSystem
        If intScoringSystem = Indicator.ScoringSystems.Score Then
            strScoringLegend = CreateList_AnswerArea_ScoringLegend(selIndicator, LANG_Scores)
        End If
        dblMaximumScore = selIndicator.GetTargetMaximumScore

        objBaselineValue = selIndicator.GetBaselineTotalValue
        objBaselineScore = selIndicator.GetBaselineTotalScore

        For i = 0 To intTargetCount - 1
            intIndex = (i * 3)
            strCells(intIndex) = "0"
            Select Case selIndicator.TargetSystem
                Case Indicator.TargetSystems.ValueRange
                    strCells(intIndex + 2) = selIndicator.GetTargetFormattedScore(i)
                Case Else
                    If selIndicator.ScoringSystem = Indicator.ScoringSystems.Percentage Then
                        strCells(intIndex + 1) = CreateExcelFormulaPercentageScore(intIndex + 10, intRowIndex)
                    End If
                    If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                        strCells(intIndex + 1) = "0"
                        strCells(intIndex + 2) = selIndicator.GetTargetScoresTotal(i).Score
                    Else
                        strCells(intIndex + 2) = selIndicator.GetTargetValuesTotal(i).MinValue
                    End If
            End Select
        Next

        Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaValue, strRowHeader, strUnit, sngValueRangeMin, sngValueRangeMax, boolValueRangeSet, intNrDecimals, _
                                             intScoringSystem, strScoringLegend, dblMaximumScore, objBaselineValue, objBaselineScore, strCells)
        TableGrid.Add(objRow)
    End Sub

    Private Sub CreateList_AnswerArea_Formula(ByVal selIndicator As Indicator)
        Dim lstIndices As New Dictionary(Of String, Integer)
        Dim intCurrentIndex As Integer
        Dim strRowCode As String
        Dim strRowHeader As String = String.Empty
        Dim intTargetCount As Integer = selIndicator.Targets.Count
        Dim strUnit As String
        Dim sngValueRangeMin, sngValueRangeMax As Single
        Dim boolValueRangeSet As Boolean
        Dim intNrDecimals As Integer
        Dim intScoringSystem As Integer
        Dim strScoringLegend As String = String.Empty
        Dim dblMaximumScore As Double
        Dim objBaselineValue, objBaselineScore As Object
        Dim intIndex, intStatementIndex As Integer

        intScoringSystem = selIndicator.ScoringSystem
        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            strScoringLegend = CreateList_AnswerArea_ScoringLegend(selIndicator, LANG_Scores)
        End If
        dblMaximumScore = selIndicator.GetTargetMaximumScore

        For Each selStatement As Statement In selIndicator.Statements
            intCurrentIndex = TableGrid.Count
            strRowCode = GetRowCode(intStatementIndex)
            lstIndices.Add(strRowCode, intCurrentIndex)

            strUnit = selStatement.ValuesDetail.Unit
            sngValueRangeMin = selStatement.ValuesDetail.ValueRange.MinValue
            sngValueRangeMax = selStatement.ValuesDetail.ValueRange.MaxValue
            boolValueRangeSet = selStatement.ValuesDetail.ValueRange.ValueRangeSet
            intNrDecimals = selStatement.ValuesDetail.NrDecimals
            objBaselineValue = selIndicator.Baseline.DoubleValues(intStatementIndex).Value

            Dim strCells(intTargetCount * 3) As String

            For i = 0 To intTargetCount - 1
                intIndex = (i * 3)
                strCells(intIndex) = "0"
                If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                    'nothing: statements are not scored
                End If
                strCells(intIndex + 2) = selIndicator.Targets(i).DoubleValues(intStatementIndex).Value
            Next
            Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaFormula, strRowHeader, strUnit, sngValueRangeMin, sngValueRangeMax, boolValueRangeSet, intNrDecimals, _
                                                 intScoringSystem, String.Empty, dblMaximumScore, objBaselineValue, Nothing, strCells)
            TableGrid.Add(objRow)

            intStatementIndex += 1
        Next

        'total
        strUnit = selIndicator.ValuesDetail.Unit
        sngValueRangeMin = selIndicator.ValuesDetail.ValueRange.MinValue
        sngValueRangeMax = selIndicator.ValuesDetail.ValueRange.MaxValue
        boolValueRangeSet = selIndicator.ValuesDetail.ValueRange.ValueRangeSet
        intNrDecimals = selIndicator.ValuesDetail.NrDecimals

        objBaselineValue = CreateExcelFormula(selIndicator.ValuesDetail.Formula, 7, lstIndices)
        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
            objBaselineScore = selIndicator.GetBaselineTotalScore
        Else
            objBaselineScore = CreateExcelFormula(selIndicator.ValuesDetail.Formula, 8, lstIndices)
        End If

        Dim strCellsTotal(intTargetCount * 3) As String

        For i = 0 To intTargetCount - 1
            intIndex = (i * 3)
            strCellsTotal(intIndex) = CreateExcelFormula(selIndicator.ValuesDetail.Formula, intIndex + 9, lstIndices)
            intIndex += 1
            If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                strCellsTotal(intIndex) = "0"
            End If
            intIndex += 1

            If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                Select Case selIndicator.TargetSystem
                    Case Indicator.TargetSystems.ValueRange
                        strCellsTotal(intIndex) = selIndicator.GetTargetFormattedScore(i)
                    Case Else
                        strCellsTotal(intIndex) = selIndicator.GetTargetScoresTotal(i).Score
                End Select
            Else
                strCellsTotal(intIndex) = CreateExcelFormula(selIndicator.ValuesDetail.Formula, intIndex + 9, lstIndices)
            End If
        Next

        Dim objRowTotal As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaFormula, strRowHeader, strUnit, sngValueRangeMin, sngValueRangeMax, boolValueRangeSet, intNrDecimals, _
                                             intScoringSystem, strScoringLegend, dblMaximumScore, objBaselineValue, objBaselineScore, strCellsTotal)
        objRowTotal.Formula = selIndicator.ValuesDetail.Formula
        TableGrid.Add(objRowTotal)
    End Sub

    Private Sub CreateList_AnswerArea_MultipleOptions(ByVal selIndicator As Indicator)
        Dim objRow As ExportIndicatorRow
        Dim intIndex, intOptionIndex As Integer
        Dim objBaselineScore As Object
        Dim intTargetCount As Integer = selIndicator.Targets.Count
        Dim boolValue As Boolean
        Dim strScoringLegend As String = String.Empty
        Dim dblMaximumScore As Double = selIndicator.GetTargetMaximumScore
        Dim strCellsTotal(intTargetCount * 3) As String
        Dim intStartIndex As Integer = TableGrid.Count
        Dim intLastIndex As Integer

        'Options
        For Each selOption As ResponseClass In selIndicator.ResponseClasses
            Dim strOption As String = selOption.ClassName
            strScoringLegend = selOption.Value

            boolValue = selIndicator.Baseline.BooleanValues(intOptionIndex).Value
            If boolValue = True Then objBaselineScore = "X" Else objBaselineScore = Nothing

            Dim strCells(intTargetCount * 3) As String

            For i = 0 To intTargetCount - 1
                intIndex = (i * 3)
                boolValue = selIndicator.Targets(i).BooleanValues(intOptionIndex).Value
                If boolValue = True Then strCells(2 + intIndex) = "X"
            Next

            objRow = New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strOption, 0, strScoringLegend, dblMaximumScore, objBaselineScore, strCells)
            TableGrid.Add(objRow)

            intOptionIndex += 1
        Next

        intLastIndex = TableGrid.Count - 1
        strScoringLegend = selIndicator.GetBaselineMaximumScore
        'objBaselineScore = selIndicator.GetBaselineTotalScore
        objBaselineScore = CreateExcelFormulaMultipleChoice(intStartIndex, intLastIndex, 8)

        For i = 0 To intTargetCount - 1
            intIndex = (i * 3)
            'strCellsTotal(intIndex) = CreateExcelFormulaMultipleChoice(intStartIndex, intLastIndex, intIndex + 9)
            intIndex += 1
            strCellsTotal(intIndex) = CreateExcelFormulaMultipleChoice(intStartIndex, intLastIndex, intIndex + 9)
            intIndex += 1
            strCellsTotal(intIndex) = CreateExcelFormulaMultipleChoice(intStartIndex, intLastIndex, intIndex + 9)
        Next

        objRow = New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, LANG_Score, selIndicator.ValuesDetail.NrDecimals, strScoringLegend, dblMaximumScore, _
                                        objBaselineScore, strCellsTotal)
        TableGrid.Add(objRow)
    End Sub

    Private Sub CreateList_AnswerArea_Ranking(ByVal selIndicator As Indicator)
        Dim objRow As ExportIndicatorRow
        Dim intIndex, intOptionIndex As Integer
        Dim objBaselineScore As Object
        Dim intTargetCount As Integer = selIndicator.Targets.Count
        Dim strScoringLegend As String = String.Empty
        Dim dblMaximumScore As Double = selIndicator.GetTargetMaximumScore
        Dim dblValue As Double

        'Options
        For Each selOption As ResponseClass In selIndicator.ResponseClasses
            Dim strOption As String = selOption.ClassName
            strScoringLegend = selOption.Value

            objBaselineScore = selIndicator.Baseline.DoubleValues(intOptionIndex).Value

            Dim strCells(intTargetCount * 3) As String

            For i = 0 To intTargetCount - 1
                intIndex = (i * 3)
                dblValue = selIndicator.Targets(i).DoubleValues(intOptionIndex).Value
                If dblValue <> 0 Then strCells(intIndex + 2) = dblValue
            Next

            objRow = New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, strOption, 0, strScoringLegend, dblMaximumScore, objBaselineScore, strCells)
            TableGrid.Add(objRow)

            intOptionIndex += 1
        Next
    End Sub

    Private Sub CreateList_AnswerArea_Scale(ByVal selIndicator As Indicator)
        Dim strRowHeader As String = String.Empty
        Dim intTargetCount As Integer = selIndicator.Targets.Count
        Dim strCells(intTargetCount * 3) As String
        Dim intNrDecimals As Integer
        Dim strScoringLegend As String
        Dim sngBaselineScore As Single
        Dim dblMaximumScore As Double = selIndicator.GetTargetMaximumScore
        Dim intIndex As Integer

        strRowHeader = ToLabel(LANG_Scores)
        intNrDecimals = selIndicator.ValuesDetail.NrDecimals
        strScoringLegend = selIndicator.GetBaselineMaximumScore
        sngBaselineScore = selIndicator.GetBaselineTotalScore

        For i = 0 To intTargetCount - 1
            intIndex = (i * 3)
            strCells(intIndex + 1) = "0"
            strCells(intIndex + 2) = selIndicator.GetTargetScoresTotal(i).Score
        Next

        Dim objRow As New ExportIndicatorRow(ExportIndicatorRow.ObjectTypes.AnswerAreaScale, strRowHeader, intNrDecimals, _
                                             strScoringLegend, dblMaximumScore, sngBaselineScore, strCells)
        TableGrid.Add(objRow)

    End Sub
#End Region
    Private Function AggregationString(ByVal selIndicator As Indicator) As String
        Dim strAggregate As String = String.Empty

        Select Case selIndicator.AggregateVertical
            Case Indicator.AggregationOptions.Average
                strAggregate = "AVG"
            Case Indicator.AggregationOptions.Maximum
                strAggregate = "MAX"
            Case Indicator.AggregationOptions.Minimum
                strAggregate = "MIN"
            Case Indicator.AggregationOptions.Sum
                strAggregate = "SUM"
        End Select

        Return strAggregate
    End Function

    Private Function CreateExcelFormulaPercentageScore(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer) As String
        Dim strExcelFormula As String
        Dim strColumnNameValue As String = GetRowCode(intColumnIndex - 1)
        Dim strColumnNameTarget As String = GetRowCode(intColumnIndex + 1)

        intRowIndex += 2

        strExcelFormula = String.Format("=IFERROR({1}{0}/{2}{0},""-"")", intRowIndex, strColumnNameValue, strColumnNameTarget)

        Return strExcelFormula
    End Function

    Private Function CreateExcelFormula(ByVal strFormula As String, ByVal intColumnIndex As Integer, ByVal lstIndices As Dictionary(Of String, Integer)) As String
        Dim strExcelFormula As String = strFormula
        Dim strColumnName As String = GetRowCode(intColumnIndex)
        Dim strCellReference As String
        

        If String.IsNullOrEmpty(strFormula) = False Then
            Dim strDelimiter As Char() = New Char() {"+", "-", "/", "*", "^", "(", ")"}
            Dim strParts() As String = strFormula.Split(strDelimiter)
            Dim strRowCode As String
            Dim intRowIndex As Integer

            For i = 0 To strParts.Length - 1
                strRowCode = Trim(strParts(i))
                If lstIndices.ContainsKey(strRowCode) Then
                    intRowIndex = lstIndices(strRowCode)
                    intRowIndex += 2 'Excel rowindex is base 1 instead of base 0 + additional title row
                    strCellReference = String.Format("{0}{1}", strColumnName, intRowIndex)

                    strExcelFormula = strExcelFormula.Replace(strRowCode, strCellReference)
                End If
            Next
        End If
        If String.IsNullOrEmpty(strExcelFormula) = False Then
            If strExcelFormula.Contains("/") Then
                strExcelFormula = String.Format("IFERROR({0},""-"")", strExcelFormula)
            End If
            strExcelFormula = String.Format("={0}", strExcelFormula)
        End If

        Return strExcelFormula
    End Function

    Private Function CreateExcelFormulaMultipleChoice(ByVal intStartIndex As Integer, ByVal intLastIndex As Integer, ByVal intColumnIndex As Integer) As String
        Dim strExcelFormula As String
        Dim strColumnName As String = GetRowCode(intColumnIndex)

        intStartIndex += 2
        intLastIndex += 2

        strExcelFormula = String.Format("=SUMIF({0}{1}:{0}{2},""<>"",$G${1}:$G${2})", strColumnName, intStartIndex, intLastIndex)

        Return strExcelFormula
    End Function

    Private Function GetRowCode(ByVal intIndex As Integer) As String
        Dim strRowCode As String = "A"
        Dim strFirst As String = String.Empty
        Dim strSecond As String = String.Empty
        Dim intIndexFirst As Integer
        Dim intIndexSecond As Integer = intIndex

        If intIndex >= 26 Then
            intIndexFirst = Math.Floor(intIndex / 26) - 1
            strFirst = Chr(intIndexFirst + 65)

            Math.DivRem(intIndex, 26, intIndexSecond)
        End If
        strSecond = Chr(intIndexSecond + 65)

        strRowCode = strFirst & strSecond
        
        Return strRowCode
    End Function

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 10

        If TargetDeadlinesSection IsNot Nothing Then
            intColumns += (TargetDeadlines.Count * 3)
        End If

        Return intColumns
    End Function

    Public Function GetRowCount() As Integer
        Return Me.TableGrid.Count
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls


    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class


