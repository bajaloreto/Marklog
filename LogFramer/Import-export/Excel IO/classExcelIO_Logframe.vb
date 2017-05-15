Partial Public Class ExcelIO
    Private IndexSections As New ArrayList
    Private IndexRepeatedPurposes As New ArrayList
    Private IndexRepeatedOutputs As New ArrayList
    Private IndexBudget As Integer

    Public Sub ExportLogFrameToExcel(ByVal logframe As LogFrame, ByVal boolShowIndColumn As Boolean, ByVal boolShowVerColumn As Boolean, ByVal boolShowAsmColumn As Boolean, _
                                     ByVal boolShowGoals As Boolean, boolShowPurposes As Boolean, boolShowOutputs As Boolean, boolShowActivities As Boolean, boolShowResourcesBudget As Boolean)
        InitialiseWorkBook()

        Using objExportLogframe As New ExportLogframe(CurrentLogFrame, boolShowIndColumn, boolShowVerColumn, boolShowAsmColumn, _
                                                          boolShowGoals, boolShowPurposes, boolShowOutputs, boolShowActivities, boolShowResourcesBudget)
            objExportLogframe.LoadTable()

            Dim objTableGrid As Object = ExportLogframe_CreateTableGrid(objExportLogframe)
            Dim intColumnCount As Integer = objExportLogframe.GetColumnCount
            Dim intRowCount As Integer = objExportLogframe.GetRowCount

            If objTableGrid IsNot Nothing Then
                Dim objWorkSheet As Object = WorkBook.Worksheets(1)

                ExportLogframe_SetColumns(objWorkSheet)
                With objWorkSheet
                    .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount)).Value = objTableGrid
                End With

                ExportLogframe_LayOut(objWorkSheet, intRowCount, intColumnCount, boolShowIndColumn, boolShowResourcesBudget)
            End If

        End Using

        FinaliseWorkBook()
    End Sub

    Private Function ExportLogframe_CreateTableGrid(ByVal objExportLogframe As ExportLogframe) As Object
        Dim intStructColIndex As Integer = 0, intIndColIndex As Integer = 2, intVerColIndex As Integer = 4, intAsmColIndex As Integer = 6
        Dim intColumnCount As Integer = 8 'objExportLogframe.GetColumnCount
        Dim intRowCount As Integer = objExportLogframe.GetRowCount
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object
        Dim intRowIndex As Integer

        If intRowCount = 0 Then Return Nothing

        intRowCount += 1

        For Each selGridRow As ExportLogframeRow In objExportLogframe.TableGrid
            If selGridRow.RowType = ExportLogframeRow.RowTypes.Section Then
                IndexSections.Add(intRowIndex + 1)

                If selGridRow.StructRtf IsNot Nothing Then _
                    objTableGrid(intRowIndex, intStructColIndex) = selGridRow.Struct.Text
                If objExportLogframe.IsResourceBudgetRow(selGridRow) = False Then
                    If objExportLogframe.ShowIndicatorColumn = True And String.IsNullOrEmpty(selGridRow.IndicatorText) = False Then _
                        objTableGrid(intRowIndex, intIndColIndex) = selGridRow.IndicatorText
                    If objExportLogframe.ShowVerificationSourceColumn = True And String.IsNullOrEmpty(selGridRow.VerificationSourceText) = False Then _
                        objTableGrid(intRowIndex, intVerColIndex) = selGridRow.VerificationSourceText
                Else
                    If objExportLogframe.ShowIndicatorColumn = True And String.IsNullOrEmpty(selGridRow.ResourceText) = False Then _
                        objTableGrid(intRowIndex, intIndColIndex) = selGridRow.ResourceText
                    If objExportLogframe.ShowVerificationSourceColumn = True And String.IsNullOrEmpty(selGridRow.BudgetText) = False Then
                        objTableGrid(intRowIndex, intVerColIndex) = LANG_Budget
                        IndexBudget = intRowIndex + 1
                    End If
                End If
                If objExportLogframe.ShowAssumptionColumn = True And String.IsNullOrEmpty(selGridRow.AssumptionText) = False Then _
                    objTableGrid(intRowIndex, intAsmColIndex) = selGridRow.AssumptionText

            ElseIf selGridRow.RowType = ExportLogframeRow.RowTypes.RepeatPurpose Or _
                selGridRow.RowType = ExportLogframeRow.RowTypes.RepeatOutput Then

                If selGridRow.RowType = ExportLogframeRow.RowTypes.RepeatPurpose Then
                    IndexRepeatedPurposes.Add(intRowIndex + 1)
                ElseIf selGridRow.RowType = ExportLogframeRow.RowTypes.RepeatOutput Then
                    IndexRepeatedOutputs.Add(intRowIndex + 1)
                End If

                If String.IsNullOrEmpty(selGridRow.StructSort) = False Then _
                    objTableGrid(intRowIndex, intStructColIndex) = selGridRow.StructSort.ToString
                If String.IsNullOrEmpty(selGridRow.StructText) = False Then _
                    objTableGrid(intRowIndex, intStructColIndex + 1) = selGridRow.StructText
            Else
                If String.IsNullOrEmpty(selGridRow.StructSort) = False Then
                    objTableGrid(intRowIndex, intStructColIndex) = selGridRow.StructSort
                End If

                If String.IsNullOrEmpty(selGridRow.StructRtf) = False Then _
                    objTableGrid(intRowIndex, intStructColIndex + 1) = selGridRow.StructText
                If objExportLogframe.IsResourceBudgetRow(selGridRow) = False Then
                    If String.IsNullOrEmpty(selGridRow.IndicatorSort) = False Then _
                        objTableGrid(intRowIndex, intIndColIndex) = selGridRow.IndicatorSort
                    If String.IsNullOrEmpty(selGridRow.IndicatorText) = False Then _
                        objTableGrid(intRowIndex, intIndColIndex + 1) = selGridRow.IndicatorText
                    If String.IsNullOrEmpty(selGridRow.VerificationSourceSort) = False Then _
                        objTableGrid(intRowIndex, intVerColIndex) = selGridRow.VerificationSourceSort
                    If String.IsNullOrEmpty(selGridRow.VerificationSourceText) = False Then _
                        objTableGrid(intRowIndex, intVerColIndex + 1) = selGridRow.VerificationSourceText
                Else
                    If String.IsNullOrEmpty(selGridRow.ResourceSort) = False Then _
                        objTableGrid(intRowIndex, intIndColIndex) = selGridRow.ResourceSort
                    If String.IsNullOrEmpty(selGridRow.ResourceText) = False Then _
                        objTableGrid(intRowIndex, intIndColIndex + 1) = selGridRow.ResourceText
                    If String.IsNullOrEmpty(selGridRow.BudgetText) = False And objExportLogframe.ShowVerificationSourceColumn = True Then
                        objTableGrid(intRowIndex, intVerColIndex + 1) = selGridRow.TotalCostAmount
                    End If
                End If
                If String.IsNullOrEmpty(selGridRow.AssumptionSort) = False Then _
                    objTableGrid(intRowIndex, intAsmColIndex) = selGridRow.AssumptionSort
                If String.IsNullOrEmpty(selGridRow.AssumptionText) = False Then _
                    objTableGrid(intRowIndex, intAsmColIndex + 1) = selGridRow.AssumptionText
            End If

            intRowIndex += 1
        Next

        Return objTableGrid
    End Function

    Private Sub ExportLogframe_SetColumns(ByVal objWorkSheet As Object)
        With objWorkSheet
            .Cells.VerticalAlignment = xlVerticalAlignment_Top
            .Columns("A:A").NumberFormat = "@"
            .Columns("C:C").NumberFormat = "@"
            .Columns("E:E").NumberFormat = "@"
            .Columns("G:G").NumberFormat = "@"

            With .Columns("B:B")
                .ColumnWidth = 25
                .WrapText = True
            End With
            With .Columns("D:D")
                .ColumnWidth = 25
                .WrapText = True
            End With
            With .Columns("F:F")
                .ColumnWidth = 25
                .WrapText = True
            End With
            With .Columns("H:H")
                .ColumnWidth = 25
                .WrapText = True
            End With
        End With
    End Sub

    Private Sub ExportLogframe_LayOut(ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer, ByVal boolShowIndicators As Boolean, ByVal boolShowResourcesBudget As Boolean)
        With objWorkSheet
            Dim intRow As Integer

            For i = 0 To IndexSections.Count - 1
                intRow = IndexSections(i)

                With .Range(.Cells(intRow, 1), .Cells(intRow, 8))
                    .Font.Bold = True
                    .HorizontalAlignment = xlHorizontalAlignment_Center
                End With
                For j = 1 To intColumnCount Step 2
                    .Range(.Cells(intRow, j), .Cells(intRow, j + 1)).Merge()
                Next

                With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount)).Interior
                    .ColorIndex = 15
                    .Pattern = xlPattern_Solid
                End With
            Next

            For i = 0 To IndexRepeatedPurposes.Count - 1
                intRow = IndexRepeatedPurposes(i)

                With .Range(.Cells(intRow, 2), .Cells(intRow, intColumnCount))
                    .Merge()
                    .HorizontalAlignment = xlHorizontalAlignment_Left
                End With
                With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount))
                    .Font.Bold = True
                    With .Interior
                        .ColorIndex = 33
                        .Pattern = xlPattern_Solid
                    End With
                End With
            Next

            For i = 0 To IndexRepeatedOutputs.Count - 1
                intRow = IndexRepeatedOutputs(i)

                With .Range(.Cells(intRow, 2), .Cells(intRow, intColumnCount))
                    .Merge()
                    .HorizontalAlignment = xlHorizontalAlignment_Left
                End With
                With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount))
                    .Font.Bold = True
                    With .Interior
                        .ColorIndex = 37
                        .Pattern = xlPattern_Solid
                    End With
                End With
            Next

            'budget column
            If boolShowResourcesBudget = True Then
                Dim intColumnIndexBudget As Integer = 6
                If boolShowIndicators = False Then intColumnIndexBudget -= 2

                With .Range(.Cells(IndexBudget, intColumnIndexBudget - 1), .Cells(IndexBudget, intColumnIndexBudget))
                    .Font.Bold = True
                    .HorizontalAlignment = xlHorizontalAlignment_Center 'Excel.XlHAlign.xlHAlignRight
                End With

                With .Range(.Cells(IndexBudget + 1, intColumnIndexBudget), .Cells(intRowCount, intColumnIndexBudget))
                    .Font.Bold = False
                    .NumberFormat = String.Format("#,##0.00 ""{0}""", CurrentLogFrame.Budget.DefaultCurrencyCode)
                    .HorizontalAlignment = xlHorizontalAlignment_Right 'Excel.XlHAlign.xlHAlignRight
                End With
            End If

            'borders
            With .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount))
                With .Cells.Borders(xlBorders_EdgeTop)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeLeft)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeTop)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeBottom)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
                    .ColorIndex = xlColorIndex_ColorIndexAutomatic
                End With
                With .Cells.Borders(xlBorders_EdgeRight)
                    .LineStyle = xlLineStyle_Continuous
                    .Weight = xlBorderWeight_Thin
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

            If My.Settings.setShowAssumptionColumn = False Then
                .Columns("G:H").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
            End If
            If My.Settings.setShowVerificationSourceColumn = False Then
                .Columns("E:F").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
            End If
            If My.Settings.setShowIndicatorColumn = False Then
                .Columns("C:D").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
            End If
        End With
    End Sub

End Class
