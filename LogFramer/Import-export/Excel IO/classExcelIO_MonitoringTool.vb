Partial Public Class ExcelIO
    Public Sub ExportMonitoringToolToExcel(ByVal logframe As LogFrame)
        'Try
        InitialiseWorkBook()

        With My.Settings
            If .setPrintIndicatorsSection = PrintSettingsIndicators.PrintSections.printall Then
                For intSection = 1 To 4
                    Dim objWorkSheet As Object

                    objWorkSheet = WorkBook.Worksheets.Add(After:=WorkBook.Sheets(WorkBook.Sheets.Count))

                    objWorkSheet.Name = logframe.StructNamePlural(intSection - 1)

                    ExportMonitoringToolToExcel_Section(logframe, intSection, objWorkSheet)
                Next
            Else
                Dim objWorkSheet As Object = WorkBook.Worksheets(1)

                ExportMonitoringToolToExcel_Section(logframe, .setPrintIndicatorsSection, objWorkSheet)
            End If

        End With
        FinaliseWorkBook()

        MsgBox(LANG_ExportToExcelComplete)
        'Catch ex As Exception
        '    MsgBox(LANG_ExportToExcelError)
        'End Try

    End Sub

    Private Sub ExportMonitoringToolToExcel_Section(ByVal logframe As LogFrame, ByVal intSection As Integer, ByVal objWorkSheet As Object)
        With My.Settings
            Using objExportMonitoringTool As New ExportMonitoringTool(CurrentLogFrame, intSection, .setPrintIndicatorsTargetgroupGuid, .setPrintIndicatorsMeasurement, _
                                              .setPrintIndicatorsPrintPurposes, .setPrintIndicatorsPrintOutputs, .setPrintIndicatorsPrintActivities, _
                                              .setPrintIndicatorsPrintOptionValues, .setPrintIndicatorsPrintValueRanges, .setPrintIndicatorsPrintTargets)
                objExportMonitoringTool.LoadTable()

                Dim objTableGrid As Object = ExportMonitoringTool_CreateTableGrid(objExportMonitoringTool)
                Dim intColumnCount As Integer = objExportMonitoringTool.GetColumnCount
                Dim intRowCount As Integer = objExportMonitoringTool.GetRowCount + 1

                If objTableGrid IsNot Nothing Then
                    objTableGrid = ExportMonitoringTool_CreateFormulas(objTableGrid)

                    ExportMonitoringTool_SetColumns(objWorkSheet)
                    With objWorkSheet
                        .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount)).Value = objTableGrid
                    End With

                    ExportMonitoringTool_LayOut(objWorkSheet, intRowCount, intColumnCount - 1)
                    ExportMonitoringTool_LayOut_Rows(objExportMonitoringTool, objWorkSheet, intRowCount, intColumnCount)
                End If
            End Using
        End With
    End Sub

#Region "Table grid and formulas"
    Private Function ExportMonitoringTool_CreateTableGrid(ByVal objExportMonitoringTool As ExportMonitoringTool) As Object
        Dim intColumnCount As Integer = objExportMonitoringTool.GetColumnCount
        Dim intRowCount As Integer = objExportMonitoringTool.GetRowCount
        Dim intColumnIndex As Integer
        Dim intRow As Integer = 1
        Dim strTargetDeadline As String

        If intRowCount = 0 Then Return Nothing

        Dim objTableGrid(intRowCount, intColumnCount - 1) As Object

        'Title row
        objTableGrid(0, 0) = LANG_Number
        objTableGrid(0, 1) = LANG_Description
        objTableGrid(0, 2) = LANG_Unit
        objTableGrid(0, 3) = LANG_ValueRange
        objTableGrid(0, 5) = LANG_Decimals
        objTableGrid(0, 6) = LANG_ScoringLegend
        objTableGrid(0, 7) = LANG_BaselineValue
        objTableGrid(0, 8) = LANG_BaselineScore

        For i = 0 To objExportMonitoringTool.TargetDeadlines.Count - 1
            intColumnIndex = (i * 3)
            strTargetDeadline = objExportMonitoringTool.TargetDeadlinesSection.FormatTargetDeadlineDate(objExportMonitoringTool.TargetDeadlines(i))

            objTableGrid(0, 9 + intColumnIndex) = String.Format("{0} {1}", LANG_Value, strTargetDeadline)
            objTableGrid(0, 10 + intColumnIndex) = String.Format("{0} {1}", LANG_Score, strTargetDeadline)
            objTableGrid(0, 11 + intColumnIndex) = String.Format("{0} {1}", LANG_Target, strTargetDeadline)
        Next

        'Values

        For Each selGridRow As ExportIndicatorRow In objExportMonitoringTool.TableGrid
            Select Case selGridRow.ObjectType
                Case PrintIndicatorRow.ObjectTypes.Goal, PrintIndicatorRow.ObjectTypes.Purpose, PrintIndicatorRow.ObjectTypes.Output, PrintIndicatorRow.ObjectTypes.Activity
                    objTableGrid(intRow, 0) = selGridRow.SortNumber
                    objTableGrid(intRow, 1) = selGridRow.Struct.Text
                Case PrintIndicatorRow.ObjectTypes.Indicator
                    objTableGrid(intRow, 0) = selGridRow.SortNumber
                    objTableGrid(intRow, 1) = selGridRow.Indicator.Text

                    If selGridRow.Indicator.Indicators.Count > 0 Then
                        Dim selIndicator As Indicator = selGridRow.Indicator
                        Select Case selIndicator.QuestionType
                            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.PercentageValue, _
                                Indicator.QuestionTypes.Ratio
                                objTableGrid(intRow, 2) = selGridRow.Unit

                                If selGridRow.ValueRangeSet Then
                                    objTableGrid(intRow, 3) = selGridRow.ValueRangeMin
                                    objTableGrid(intRow, 4) = selGridRow.ValueRangeMax
                                End If

                                objTableGrid(intRow, 5) = selGridRow.NrDecimals
                                'objTableGrid(intRow, 6) = selGridRow.ScoringLegend
                            Case Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, _
                                Indicator.QuestionTypes.LikertTypeScale, Indicator.QuestionTypes.SemanticDiff, Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.LikertScale, _
                                Indicator.QuestionTypes.FrequencyLikert, Indicator.QuestionTypes.CumulativeScale
                                objTableGrid(intRow, 5) = selGridRow.NrDecimals
                                objTableGrid(intRow, 6) = selGridRow.ScoringLegend
                        End Select

                        objTableGrid(intRow, 7) = selGridRow.BaselineValue
                        objTableGrid(intRow, 8) = selGridRow.BaselineScore
                        If selGridRow.Cells IsNot Nothing Then
                            For i = 0 To selGridRow.Cells.Count - 1
                                objTableGrid(intRow, 9 + i) = selGridRow.Cells(i)
                            Next
                        End If
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaOpenEnded
                    'do nothing
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaMaxDiff
                    If selGridRow.Statement IsNot Nothing Then
                        objTableGrid(intRow, 1) = RichTextToText(selGridRow.Statement.FirstLabel)
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaValue, PrintIndicatorRow.ObjectTypes.AnswerAreaFormula
                    If String.IsNullOrEmpty(selGridRow.Formula) = False Then
                        objTableGrid(intRow, 1) = selGridRow.Formula
                    Else
                        objTableGrid(intRow, 1) = selGridRow.RowHeader
                    End If

                    objTableGrid(intRow, 2) = selGridRow.Unit

                    If selGridRow.ValueRangeSet Then
                        objTableGrid(intRow, 3) = selGridRow.ValueRangeMin
                        objTableGrid(intRow, 4) = selGridRow.ValueRangeMax
                    End If

                    objTableGrid(intRow, 5) = selGridRow.NrDecimals
                    objTableGrid(intRow, 6) = selGridRow.ScoringLegend
                    objTableGrid(intRow, 7) = selGridRow.BaselineValue

                    If selGridRow.ScoringSystem = Indicator.ScoringSystems.Score Then
                        objTableGrid(intRow, 8) = selGridRow.BaselineScore
                    End If
                    If selGridRow.Cells IsNot Nothing Then
                        For i = 0 To selGridRow.Cells.Count - 1
                            objTableGrid(intRow, 9 + i) = selGridRow.Cells(i)
                        Next
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, PrintIndicatorRow.ObjectTypes.AnswerAreaLikertType
                    objTableGrid(intRow, 1) = selGridRow.RowHeader
                    objTableGrid(intRow, 2) = selGridRow.Unit
                    If selGridRow.RowHeader = LANG_Score Then
                        objTableGrid(intRow, 5) = selGridRow.NrDecimals
                    End If
                    objTableGrid(intRow, 6) = selGridRow.ScoringLegend
                    objTableGrid(intRow, 8) = selGridRow.BaselineScore

                    If selGridRow.Cells IsNot Nothing Then
                        For i = 0 To selGridRow.Cells.Count - 1
                            objTableGrid(intRow, 9 + i) = selGridRow.Cells(i)
                        Next
                    End If
                Case PrintIndicatorRow.ObjectTypes.AnswerAreaScale, PrintIndicatorRow.ObjectTypes.AnswerAreaLikertScale, PrintIndicatorRow.ObjectTypes.AnswerAreaCumulativeScale, _
                    PrintIndicatorRow.ObjectTypes.AnswerAreaSemanticDiff

                    objTableGrid(intRow, 1) = selGridRow.RowHeader
                    objTableGrid(intRow, 5) = selGridRow.NrDecimals
                    objTableGrid(intRow, 6) = selGridRow.ScoringLegend
                    objTableGrid(intRow, 8) = selGridRow.BaselineScore

                    If selGridRow.Cells IsNot Nothing Then
                        For i = 0 To selGridRow.Cells.Count - 1
                            objTableGrid(intRow, 9 + i) = selGridRow.Cells(i)
                        Next
                    End If
            End Select
            'objTableGrid(intRow, intColumnCount - 1) = selGridRow.ObjectType
            intRow += 1
        Next

        Return objTableGrid
    End Function

    Private Function ExportMonitoringTool_CreateFormulas(ByVal objTableGrid As Object(,)) As Object
        Dim strFormula, strValue As String
        Dim objValue As String
        Dim intRowCount As Integer = objTableGrid.GetLength(0)
        Dim intColumnCount As Integer = objTableGrid.GetLength(1)
        Dim intRowIndex As Integer
        Dim strSortNumber As String

        Do Until intRowIndex = intRowCount - 1
            objValue = objTableGrid(intRowIndex, 0)
            If objValue IsNot Nothing Then
                strSortNumber = objValue.ToString

                For intColumnIndex = 6 To intColumnCount - 1
                    objValue = objTableGrid(intRowIndex, intColumnIndex)
                    If objValue IsNot Nothing Then
                        strValue = objValue.ToString

                        Select Case strValue
                            Case "AVG", "MAX", "MIN", "SUM"
                                Dim alRows As ArrayList = ExportMonitoringTool_ListChildren(objTableGrid, strSortNumber, intRowIndex, intRowCount, intColumnIndex)

                                If alRows.Count > 0 Then
                                    Select Case strValue
                                        Case "AVG"
                                            strFormula = ExportMonitoringTool_ParseAverage(alRows, intColumnIndex)
                                        Case "SUM"
                                            strFormula = ExportMonitoringTool_ParseSum(alRows, intColumnIndex)
                                        Case "MIN"
                                            strFormula = ExportMonitoringTool_ParseMinimum(alRows, intColumnIndex)
                                        Case "MAX"
                                            strFormula = ExportMonitoringTool_ParseMaximum(alRows, intColumnIndex)
                                        Case Else
                                            strFormula = String.Empty
                                    End Select
                                    If String.IsNullOrEmpty(strFormula) = False Then objTableGrid(intRowIndex, intColumnIndex) = strFormula
                                End If
                        End Select
                    End If
                Next
            End If
            intRowIndex += 1
        Loop

        Return objTableGrid
    End Function

    Private Function ExportMonitoringTool_ListChildren(ByVal objTableGrid As Object(,), ByVal strParentSortNumber As String, ByVal intStartRowIndex As Integer, ByVal intRowCount As Integer, ByVal intColumnIndex As Integer) As ArrayList
        Dim strFormula As String = String.Empty
        Dim strSortNumber As String
        Dim strDivider = CurrentLogFrame.SortNumberDivider
        Dim alRows As New ArrayList

        strParentSortNumber &= strDivider

        Dim intParentLength As Integer = strParentSortNumber.Length

        intStartRowIndex += 1

        For intRowIndex = intStartRowIndex To intRowCount - 1
            strSortNumber = objTableGrid(intRowIndex, 0)
            If String.IsNullOrEmpty(strSortNumber) = False AndAlso strSortNumber.StartsWith(strParentSortNumber) Then
                strSortNumber = strSortNumber.Remove(0, intParentLength)

                If strSortNumber.Contains(strDivider) = False Then
                    For intStatementIndex = intRowIndex + 1 To intRowCount - 1
                        strSortNumber = objTableGrid(intStatementIndex, 0)
                        If String.IsNullOrEmpty(strSortNumber) = False Then
                            alRows.Add(intStatementIndex - 1)
                            Exit For
                        End If
                    Next
                End If
            End If
        Next

        Return alRows
    End Function

    Private Function ExportMonitoringTool_ParseSum(ByVal alRows As ArrayList, ByVal intColumnIndex As Integer) As String
        Dim strColumn As String = Chr(intColumnIndex + 65)

        Dim strFormula As String = ExportMonitoringTool_ParseSum_Parse(alRows, strColumn)

        strFormula = String.Format("=IFERROR({0},""-"")", strFormula)

        Return strFormula
    End Function

    Private Function ExportMonitoringTool_ParseSum_Parse(ByVal alRows As ArrayList, ByVal strColumn As String, Optional ByVal strFormula As String = "") As String
        Dim intCurrent, intNext As Integer
        Dim intIndex, intCounter As Integer


        If String.IsNullOrEmpty(strFormula) = False Then
            strFormula &= "+"
            intIndex = 0
        End If

        If alRows.Count > 1 Then
            'find consecutive rows
            Do Until intIndex >= alRows.Count - 1
                intCurrent = alRows(intIndex)
                If intIndex = alRows.Count - 1 Then
                    intNext = 0
                Else
                    intNext = alRows(intIndex + 1)
                End If

                If intNext - intCurrent = 1 Then
                    intCounter += 1
                Else
                    Exit Do
                End If
                intIndex += 1
            Loop

            If intCounter = 0 Then
                'if there are no consecutive rows: A1 + A3
                strFormula &= String.Format("{0}{1}", strColumn, alRows(0) + 1)
                alRows.RemoveAt(0)

                If alRows.Count > 0 Then
                    strFormula = ExportMonitoringTool_ParseSum_Parse(alRows, strColumn, strFormula)
                End If
            ElseIf intCounter > 0 Then
                'if there are consecutive rows: SUM(A1:A3)
                strFormula &= String.Format("SUM({0}{1}:{0}{2})", strColumn, alRows(0) + 1, alRows(intCounter) + 1)
                alRows.RemoveRange(0, intCounter + 1)

                If alRows.Count > 0 Then
                    strFormula = ExportMonitoringTool_ParseSum_Parse(alRows, strColumn, strFormula)
                End If
            End If
        ElseIf alRows.Count = 1 Then
            strFormula &= String.Format("{0}{1}", strColumn, alRows(0) + 1)
        End If

        Return strFormula
    End Function

    Private Function ExportMonitoringTool_ParseAverage(ByVal alRows As ArrayList, ByVal intColumnIndex As Integer) As String
        Dim intCount As Integer = alRows.Count
        Dim strFormula As String = ExportMonitoringTool_ListCells(alRows, intColumnIndex)

        strFormula = String.Format("=IFERROR(AVERAGE({0}),""-"")", strFormula)

        Return strFormula
    End Function

    Private Function ExportMonitoringTool_ParseMinimum(ByVal alRows As ArrayList, ByVal intColumnIndex As Integer) As String
        Dim intCount As Integer = alRows.Count
        Dim strFormula As String = ExportMonitoringTool_ListCells(alRows, intColumnIndex)

        strFormula = String.Format("=IFERROR(MIN({0}),""-"")", strFormula)

        Return strFormula
    End Function

    Private Function ExportMonitoringTool_ParseMaximum(ByVal alRows As ArrayList, ByVal intColumnIndex As Integer) As String
        Dim intCount As Integer = alRows.Count
        Dim strFormula As String = ExportMonitoringTool_ListCells(alRows, intColumnIndex)

        strFormula = String.Format("=IFERROR(MAX({0}),""-"")", strFormula)

        Return strFormula
    End Function

    Private Function ExportMonitoringTool_ListCells(ByVal alRows As ArrayList, ByVal intColumnIndex As Integer) As String
        Dim strFormula As String = String.Empty
        Dim strColumn As String = Chr(intColumnIndex + 65)

        For i = 0 To alRows.Count - 1
            If String.IsNullOrEmpty(strFormula) = False Then strFormula &= ","
            strFormula &= String.Format("{0}{1}", strColumn, alRows(i) + 1)
        Next

        Return strFormula

    End Function
#End Region

#Region "Lay-out"
    Private Sub ExportMonitoringTool_SetColumns(ByVal objWorkSheet As Object)
        With objWorkSheet
            .Cells.VerticalAlignment = xlVerticalAlignment_Top
            .Columns("A:A").NumberFormat = "@"

            With .Columns("B:B")
                .ColumnWidth = 60
                .WrapText = True
            End With

            With .Columns("G:G")
                .ColumnWidth = 25
                .WrapText = True
            End With
        End With
    End Sub

    Private Sub ExportMonitoringTool_LayOut(ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer)
        With objWorkSheet
            Dim intRow As Integer

            intRow = 1

            'header row
            With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount))
                .Font.Bold = True
                .HorizontalAlignment = xlHorizontalAlignment_Center
                .VerticalAlignment = xlVerticalAlignment_Center
                .WrapText = True
            End With
            .Range(.Cells(intRow, 4), .Cells(intRow, 5)).Merge()
            With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount)).Interior
                .ColorIndex = 15
                .Pattern = xlPattern_Solid
            End With
            .Rows("1:1").EntireRow.AutoFit()
            .Range("C2").Select()
            .Application.ActiveWindow.FreezePanes = True

            'borders
            With .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount))
                With .Cells.Borders(xlBorders_EdgeTop)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeLeft)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeBottom)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeRight)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_InsideVertical)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_InsideHorizontal)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
            End With

            With .Range(.Cells(1, 2), .Cells(intRowCount, 2))
                With .Cells.Borders(xlBorders_EdgeRight)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
            End With
            With .Range(.Cells(1, 7), .Cells(intRowCount, 7))
                With .Cells.Borders(xlBorders_EdgeRight)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
            End With
            With .Range(.Cells(1, 9), .Cells(intRowCount, 9))
                With .Cells.Borders(xlBorders_EdgeRight)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Medium
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
            End With

            For i = 10 To intColumnCount - 1 Step 3
                With .Range(.Cells(1, i + 2), .Cells(intRowCount, i + 2))
                    With .Cells.Borders(xlBorders_EdgeRight)
                        .LineStyle = xlLineStyle_Continuous
                        .Weight = xlBorderWeight_Medium
                        .ColorIndex = xlColorIndex_ColorIndexAutomatic
                    End With
                End With
            Next

            .Columns("A:A").EntireColumn.AutoFit()
        End With
    End Sub

    Private Sub ExportMonitoringTool_LayOut_Rows(ByVal objExportMonitoringTool As ExportMonitoringTool, ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer)
        Dim selRow As ExportIndicatorRow
        Dim intRow, intIndex As Integer

        With objWorkSheet
            For intRow = 2 To intRowCount
                intIndex = intRow - 2
                selRow = objExportMonitoringTool.TableGrid(intIndex)
                'objValue = .Cells(intRow, intColumnCount).value

                'If objValue IsNot Nothing Then
                'intObjectType = selRow.ObjectType 'ParseInteger(objValue)

                Select Case selRow.ObjectType
                    Case ExportIndicatorRow.ObjectTypes.Goal
                        .Range(.Cells(intRow, 2), .Cells(intRow, intColumnCount - 1)).Merge()
                        With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount - 1)).Interior
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = -0.249977111117893
                        End With
                    Case ExportIndicatorRow.ObjectTypes.Purpose
                        .Range(.Cells(intRow, 2), .Cells(intRow, intColumnCount - 1)).Merge()
                        With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount - 1)).Interior
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = -0.249977111117893
                        End With
                    Case ExportIndicatorRow.ObjectTypes.Output
                        .Range(.Cells(intRow, 2), .Cells(intRow, intColumnCount - 1)).Merge()
                        With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount - 1)).Interior
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = 0.399975585192419
                        End With
                    Case ExportIndicatorRow.ObjectTypes.Activity
                        .Range(.Cells(intRow, 2), .Cells(intRow, intColumnCount - 1)).Merge()
                        With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount - 1)).Interior
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = 0.599993896298105
                        End With
                    Case ExportIndicatorRow.ObjectTypes.Indicator
                        Select Case selRow.Indicator.QuestionType
                            Case Indicator.QuestionTypes.Image, Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff
                                With .Range(.Cells(intRow, 3), .Cells(intRow, 7)).Interior
                                    .Pattern = xlPattern_Solid
                                    .PatternColorIndex = xlColorIndex_ColorIndexAutomatic
                                    .ThemeColor = xlThemeColorLight1
                                    .TintAndShade = 0.349986266670736
                                    .PatternTintAndShade = 0
                                End With
                        End Select
                        Select Case selRow.Indicator.QuestionType
                            Case Indicator.QuestionTypes.OpenEnded
                                .Range(.Cells(intRow, 8), .Cells(intRow, 9)).Merge()

                                For i = 10 To intColumnCount - 1 Step 3
                                    .Range(.Cells(intRow, i), .Cells(intRow, i + 2)).Merge()
                                Next
                            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio
                                ExportMonitoringTool_LayOut_Values(objWorkSheet, selRow, intRow, intColumnCount)
                            Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.LikertTypeScale,
                                Indicator.QuestionTypes.SemanticDiff, Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert, _
                                Indicator.QuestionTypes.CumulativeScale
                                ExportMonitoringTool_LayOut_MultipleOptions(objWorkSheet, selRow, intRow, intColumnCount)
                        End Select
                    Case ExportIndicatorRow.ObjectTypes.AnswerAreaMaxDiff
                        With .Range(.Cells(intRow, 3), .Cells(intRow, 7)).Interior
                            .Pattern = xlPattern_Solid
                            .PatternColorIndex = xlColorIndex_ColorIndexAutomatic
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0.349986266670736
                            .PatternTintAndShade = 0
                        End With
                        With .Cells(intRow, 9).Interior
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0.349986266670736
                        End With

                        For i = 10 To intColumnCount - 1 Step 3
                            With .Range(.Cells(intRow, i + 1), .Cells(intRow, i + 2)).Interior
                                .ThemeColor = xlThemeColorLight1
                                .TintAndShade = 0.349986266670736
                            End With
                        Next
                    Case ExportIndicatorRow.ObjectTypes.AnswerAreaValue, ExportIndicatorRow.ObjectTypes.AnswerAreaFormula
                        ExportMonitoringTool_LayOut_Values(objWorkSheet, selRow, intRow, intColumnCount)
                    Case ExportIndicatorRow.ObjectTypes.AnswerAreaMultipleOptions, ExportIndicatorRow.ObjectTypes.AnswerAreaScale
                        ExportMonitoringTool_LayOut_MultipleOptions(objWorkSheet, selRow, intRow, intColumnCount)
                End Select
            Next
        End With
    End Sub

    Private Sub ExportMonitoringTool_LayOut_Values(ByVal objWorkSheet As Object, ByVal selRow As ExportIndicatorRow, ByVal intRow As Integer, ByVal intColumnCount As Integer)
        Dim strUnit As String = selRow.Unit
        Dim intNrDecimals As Integer = selRow.NrDecimals
        Dim strNumberFormat As String = GetNumberFormat(intNrDecimals, strUnit)

        If String.IsNullOrEmpty(strNumberFormat) = False Then
            With objWorkSheet
                .Range(.Cells(intRow, 4), .Cells(intRow, 5)).NumberFormat = strNumberFormat

                Select Case selRow.ScoringSystem
                    Case Indicator.ScoringSystems.Value
                        .Range(.Cells(intRow, 8), .Cells(intRow, intColumnCount - 1)).NumberFormat = strNumberFormat

                        With .Cells(intRow, 9).Interior
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0.349986266670736
                        End With

                        For i = 10 To intColumnCount - 1 Step 3
                            With .Cells(intRow, i + 1).Interior
                                .ThemeColor = xlThemeColorLight1
                                .TintAndShade = 0.349986266670736
                            End With
                        Next
                    Case Indicator.ScoringSystems.Percentage
                        Dim strNumberFormatPercentage As String = GetNumberFormat(intNrDecimals, "%")

                        .Range(.Cells(intRow, 8), .Cells(intRow, intColumnCount - 1)).NumberFormat = strNumberFormat

                        With .Cells(intRow, 9).Interior
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0.349986266670736
                        End With

                        For i = 10 To intColumnCount - 1 Step 3
                            .Cells(intRow, i).NumberFormat = strNumberFormat
                            .Cells(intRow, i + 1).NumberFormat = strNumberFormatPercentage
                            .Cells(intRow, i + 2).NumberFormat = strNumberFormat
                        Next
                    Case Indicator.ScoringSystems.Score
                        Dim strNumberFormatScore As String = GetNumberFormatScore(intNrDecimals, selRow.MaximumScore)

                        .Cells(intRow, 8).NumberFormat = strNumberFormat

                        If selRow.ObjectType = ExportIndicatorRow.ObjectTypes.AnswerAreaFormula And String.IsNullOrEmpty(selRow.Formula) Then
                            With .Cells(intRow, 9).Interior
                                .ThemeColor = xlThemeColorLight1
                                .TintAndShade = 0.349986266670736
                            End With
                        Else
                            .Cells(intRow, 9).NumberFormat = strNumberFormatScore
                        End If

                        For i = 10 To intColumnCount - 1 Step 3
                            .Cells(intRow, i).NumberFormat = strNumberFormat
                            If selRow.ObjectType = ExportIndicatorRow.ObjectTypes.AnswerAreaFormula And String.IsNullOrEmpty(selRow.Formula) Then
                                With .Cells(intRow, i + 1).Interior
                                    .ThemeColor = xlThemeColorLight1
                                    .TintAndShade = 0.349986266670736
                                End With
                            Else
                                .Cells(intRow, i + 1).NumberFormat = strNumberFormatScore
                            End If
                            .Cells(intRow, i + 2).NumberFormat = strNumberFormatScore
                        Next
                End Select
            End With
        End If
    End Sub

    Private Sub ExportMonitoringTool_LayOut_MultipleOptions(ByVal objWorkSheet As Object, ByVal selRow As ExportIndicatorRow, ByVal intRow As Integer, ByVal intColumnCount As Integer)
        Dim intNrDecimals As Integer
        Dim strNumberFormatScore As String = String.Empty

        If selRow.ObjectType = ExportIndicatorRow.ObjectTypes.Indicator Or selRow.ObjectType = ExportIndicatorRow.ObjectTypes.AnswerAreaScale Or selRow.RowHeader = LANG_Score Then
            intNrDecimals = selRow.NrDecimals
            strNumberFormatScore = GetNumberFormatScore(intNrDecimals, selRow.MaximumScore)
        End If

        With objWorkSheet
            With .Range(.Cells(intRow, 3), .Cells(intRow, 5)).Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349986266670736
            End With
            With .Cells(intRow, 8).Interior
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349986266670736
            End With
            If String.IsNullOrEmpty(strNumberFormatScore) = False Then
                .Cells(intRow, 9).NumberFormat = strNumberFormatScore
            Else
                .Cells(intRow, 9).HorizontalAlignment = xlHorizontalAlignment_Center
            End If

            For i = 10 To intColumnCount - 1 Step 3
                With .Cells(intRow, i).Interior
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0.349986266670736
                End With
                If String.IsNullOrEmpty(strNumberFormatScore) = False Then
                    .Range(.Cells(intRow, i + 1), .Cells(intRow, i + 2)).NumberFormat = strNumberFormatScore
                Else
                    .Range(.Cells(intRow, i + 1), .Cells(intRow, i + 2)).HorizontalAlignment = xlHorizontalAlignment_Center
                End If
            Next
        End With
    End Sub

    Private Function GetNumberFormat(ByVal intNrDecimals As Integer, ByVal strUnit As String) As String
        Dim strNumberFormat As String = "#,##0"
        If intNrDecimals > 0 Then
            strNumberFormat &= "."
            For i = 1 To intNrDecimals
                strNumberFormat &= "0"
            Next
        End If
        If String.IsNullOrEmpty(strUnit) = False Then strNumberFormat = String.Format("{0} ""{1}""", strNumberFormat, strUnit)

        Return strNumberFormat
    End Function

    Private Function GetNumberFormatScore(ByVal intNrDecimals As Integer, ByVal dblMaximumScore As Double) As String
        Dim strNumberFormat As String = "#,##0"
        Dim strFormatedMaxScore As String = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

        If intNrDecimals > 0 Then
            strNumberFormat &= "."
            For i = 1 To intNrDecimals
                strNumberFormat &= "0"
            Next
        End If
        strNumberFormat = String.Format("{0}""/{1}""", strNumberFormat, strFormatedMaxScore)

        Return strNumberFormat
    End Function
#End Region
End Class
