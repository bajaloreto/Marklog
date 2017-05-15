Partial Public Class WordIO
    Public Sub ExportPlanningToNewDocument(ByVal intPlanningElements As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        ExportPlanning(intPlanningElements, objDoc, objRange)
    End Sub

    Public Sub ExportPlanningToDocument(ByVal intPlanningElements As Integer, ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportPlanning(intPlanningElements, objDoc, objRange)
    End Sub

    Public Sub ExportPlanningToBookmark(ByVal intPlanningElements As Integer, ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportPlanning(intPlanningElements, objDoc, objRange)
    End Sub

    Private Sub ExportPlanning(ByVal intPlanningElements As Integer, ByVal objDoc As Object, ByVal objRange As Object)
        Dim datExecutionStart As Date = CurrentLogFrame.GetExecutionStart
        Dim datExecutionEnd As Date = CurrentLogFrame.GetExecutionEnd
        Dim intYearStart As Integer = datExecutionStart.Year
        Dim intYearEnd As Integer = datExecutionEnd.Year
        Dim intDuration = intYearEnd - intYearStart + 1

        Dim lstRanges As New ArrayList

        'Set styles
        ExportPlanning_Styles(objDoc.Application, objDoc)

        'create planning table for each year of the project
        For intYear = intYearStart To intYearEnd
            Dim datPeriodStart As New Date(intYear, 1, 1)
            Dim datPeriodEnd As New Date(intYear, 12, 31)

            objRange.Select()
            With objDoc.Application.Selection
                .Style = "Planning year"
                .TypeText(intYear.ToString)
                .TypeParagraph()
                .Style = wdStyleNormal
                objRange = .Range
            End With

            Using objExportPlanning As New ExportPlanning(CurrentLogFrame, My.Settings.setPlanningPeriodView, My.Settings.setPlanningElementsView, datPeriodStart, datPeriodEnd, True, False)
                objExportPlanning.LoadTable()

                Dim objTableGrid As Object = ExportPlanning_CreateTableGrid(objExportPlanning, datPeriodStart, datPeriodEnd)
                Dim objTable As Object = Nothing

                If objTableGrid IsNot Nothing Then objTable = CreateTable(objDoc, objRange, objTableGrid)

                If objTable IsNot Nothing Then
                    With objTable
                        .Rows(1).HeadingFormat = True
                        .Style = "PlanningTable"

                        '*** Takes (much) more time ***
                        'Dim intLastRow As Integer = .Rows.Count
                        'Dim intLastColumn As Integer = .Columns.Count
                        'Dim PlanningCells As Object

                        'For i = 1 To intLastRow
                        '    PlanningCells = objDoc.Range(.Cell(i, 3).Range.Start, .Cell(i, intLastColumn).Range.End)
                        '    PlanningCells.Style = "Planning column"
                        'Next

                        With .Columns(1)
                            .PreferredWidthType = wdPreferredWidthPercent
                            .PreferredWidth = 20
                        End With

                        .AutoFitBehavior(wdAutofitBehaviour_FitContent)
                    End With

                    objRange.Collapse(wdCollapseEnd)
                    objRange.InsertParagraphAfter()
                End If
            End Using
        Next
    End Sub

    Private Function ExportPlanning_CreateTableGrid(ByVal objExportPlanning As ExportPlanning, ByVal datPeriodStart As Date, ByVal datPeriodEnd As Date) As Object
        Dim intColumnCount As Integer = GetMonthsBetween(datPeriodStart, datPeriodEnd) + 2
        Dim intRowCount As Integer = objExportPlanning.GetRowCount + 1
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object

        If intRowCount = 0 Then Return Nothing

        Dim intRow As Integer = 1
        Dim intStartCell As Integer, intEndCell As Integer

        objTableGrid(0, 0) = LANG_Number
        objTableGrid(0, 1) = LANG_Description

        For i = 1 To 12
            objTableGrid(0, 1 + i) = New Date(datPeriodStart.Year, i, 1).ToString("MMM")
        Next
        For Each selGridRow As ExportPlanningRow In objExportPlanning.TableGrid
            objTableGrid(intRow, 0) = selGridRow.SortNumber

            If selGridRow.RowType = ExportPlanningRow.RowTypes.KeyMoment Then
                objTableGrid(intRow, 1) = selGridRow.KeyMomentText

                If selGridRow.StartDate <> Date.MinValue Then
                    If selGridRow.StartDate < datPeriodEnd And selGridRow.StartDate > datPeriodStart Then
                        intStartCell = GetMonthsBetween(datPeriodStart, selGridRow.StartDate) + 1

                        If intStartCell < 2 Then intStartCell = 2
                        If intStartCell > intColumnCount - 1 Then intStartCell = intColumnCount - 1

                        objTableGrid(intRow, intStartCell) = "*"
                    End If
                End If
            Else
                objTableGrid(intRow, 1) = selGridRow.StructText

                If selGridRow.RowType = ExportPlanningRow.RowTypes.Activity Then
                    Dim selActivity As Activity = CType(selGridRow.Struct, Activity)
                    Dim datStart As Date = selGridRow.StartDate
                    Dim datEnd As Date = selGridRow.EndDate

                    If selActivity.IsProcess = True Then
                        datStart = selActivity.ExactStartDate
                        datEnd = selActivity.ExactEndDate
                    End If

                    If datStart <> Date.MinValue And datEnd <> Date.MinValue Then
                        If datStart < datPeriodEnd And datEnd > datPeriodStart Then
                            Dim strProcessStart As String = "|"
                            Dim strProcessEnd As String = "|"

                            If datStart < datPeriodStart Then
                                datStart = datPeriodStart
                                strProcessStart = Chr(151)
                            End If
                            If datEnd > datPeriodEnd Then
                                datEnd = datPeriodEnd
                                strProcessEnd = Chr(151)
                            End If

                            intStartCell = GetMonthsBetween(datPeriodStart, datStart) + 1
                            intEndCell = GetMonthsBetween(datPeriodStart, datEnd) + 1

                            If intStartCell < 2 Then intStartCell = 2
                            If intEndCell > intColumnCount - 1 Then intEndCell = intColumnCount - 1

                            If selActivity.IsProcess = True Then
                                objTableGrid(intRow, intStartCell) = strProcessStart
                                For i = intStartCell + 1 To intEndCell - 1
                                    objTableGrid(intRow, i) = Chr(151)
                                Next
                                objTableGrid(intRow, intEndCell) = strProcessEnd
                            Else
                                For i = intStartCell To intEndCell
                                    objTableGrid(intRow, i) = "X"
                                Next
                            End If

                        End If
                    End If
                End If
            End If

            intRow += 1
        Next

        Return objTableGrid
    End Function

    Private Function GetMonthsBetween(ByVal datStart As Date, ByVal datEnd As Date) As Integer
        Dim intYears As Integer = Math.Abs((datEnd.Year - datStart.Year))
        Dim intMonths As Integer = ((intYears * 12) + Math.Abs((datEnd.Month - datStart.Month))) + 1

        Return intMonths
    End Function

    Private Sub ExportPlanning_Styles(ByVal MsWord As Object, ByVal objDoc As Object)
        'Title row
        Dim StyleTitleRow As Object = objDoc.Styles.Add("Title row", wdStyleParagraph)
        With StyleTitleRow.Font
            .Size = 16
            .Bold = True
            .Color = -738131969
        End With
        With StyleTitleRow.ParagraphFormat
            .SpaceBefore = 18
            .SpaceBeforeAuto = False
            .SpaceAfter = 12
            .SpaceAfterAuto = False
        End With

        'Year
        Dim StyleYear As Object = objDoc.Styles.Add("Planning year", wdStyleParagraph)
        With StyleYear.Font
            .Size = 16
            .Bold = True
            .Color = -738131969
        End With
        With StyleYear.ParagraphFormat
            .SpaceBefore = 18
            .SpaceBeforeAuto = False
            .SpaceAfter = 12
            .SpaceAfterAuto = False
        End With

        'Planning column content
        Dim StylePlanning As Object = objDoc.Styles.Add("Planning column", wdStyleParagraph)
        With StylePlanning.Font
            .Size = 11
            .Bold = False
        End With
        With StylePlanning.ParagraphFormat
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RightIndent = MsWord.CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphCenter
        End With

        'Planning table
        Dim TableStylePlanning As Object = objDoc.Styles.Add("PlanningTable", wdStyleTypeTable)
        With TableStylePlanning.Font
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = 0
            .StrikeThrough = False
            .Color = wdColorAutomatic
        End With
        With TableStylePlanning.ParagraphFormat
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RightIndent = MsWord.CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
        End With
        TableStylePlanning.NoSpaceBetweenParagraphsOfSameStyle = False
        TableStylePlanning.ParagraphFormat.TabStops.ClearAll()
        TableStylePlanning.Frame.Delete()

        With TableStylePlanning.Table
            .TableDirection = 1
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)
            .Alignment = 0
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowBreakAcrossPage = True
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RowStripe = 0
            .ColumnStripe = 0

            With .Shading
                .Texture = 0
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            With .Borders
                .InsideLineStyle = wdLineStyleSingle
                .InsideLineWidth = wdLineWidth025pt
                .InsideColor = wdColorAutomatic
                .OutsideLineStyle = wdLineStyleSingle
                .OutsideLineWidth = wdLineWidth025pt
                .OutsideColor = wdColorAutomatic
                .Shadow = False
            End With
        End With
        With TableStylePlanning.Table.Condition(wdFirstRow)
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)

            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = -587137152
            End With
            With .Borders
                .InsideLineStyle = wdLineStyleSingle
                .InsideLineWidth = wdLineWidth025pt
                .InsideColor = wdColorAutomatic
                .OutsideLineStyle = wdLineStyleSingle
                .OutsideLineWidth = wdLineWidth025pt
                .OutsideColor = wdColorAutomatic
                .Shadow = False
            End With
            With .ParagraphFormat
                .LeftIndent = MsWord.CentimetersToPoints(0)
                .RightIndent = MsWord.CentimetersToPoints(0)
                .SpaceBefore = 4
                .SpaceBeforeAuto = False
                .SpaceAfter = 4
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .Alignment = wdAlignParagraphCenter
                .WidowControl = True
                .LineUnitBefore = 0.8
                .LineUnitAfter = 0.8
            End With
            With .Font
                .Bold = True
            End With
            .ParagraphFormat.TabStops.ClearAll()
        End With
    End Sub
End Class
