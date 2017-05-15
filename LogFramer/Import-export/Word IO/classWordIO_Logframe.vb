Partial Public Class WordIO
    Public Sub ExportLogframeToNewDocument(ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        ExportLogFrame(objDoc, objRange)
    End Sub

    Public Sub ExportLogframeToDocument(ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc = Nothing Or objRange = Nothing Then Exit Sub

        ExportLogFrame(objDoc, objRange)
    End Sub

    Public Sub ExportLogframeToBookmark(ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportLogFrame(objDoc, objRange)
    End Sub

    Private Sub ExportLogFrame(ByVal objDoc As Object, ByVal objRange As Object)
        Dim intIndColIndex As Integer = 1, intVerColIndex As Integer = 1, intAsmColIndex As Integer

        With CurrentProjectForm.dgvLogframe
            Using objExportLogframe As New ExportLogframe(CurrentLogFrame, .ShowIndicatorColumn, .ShowVerificationSourceColumn, .ShowAssumptionColumn, _
                                                          .ShowGoals, .ShowPurposes, .ShowOutputs, .ShowActivities, .ShowResourcesBudget)
                objExportLogframe.LoadTable()

                Dim intColumnCount As Integer = objExportLogframe.GetColumnCount
                Dim intRowCount As Integer = objExportLogframe.GetRowCount

                With objExportLogframe
                    If .ShowIndicatorColumn = True Then
                        intIndColIndex = 3
                    End If
                    If .ShowVerificationSourceColumn = True Then
                        intVerColIndex = intIndColIndex + 2
                    End If
                    If .ShowAssumptionColumn = True Then
                        If .ShowVerificationSourceColumn = True Then
                            intAsmColIndex = intVerColIndex + 2
                        Else
                            intAsmColIndex = intIndColIndex + 2
                        End If
                    End If
                End With

                Dim objTable As Object = _
                    objDoc.Tables.Add(objRange, intRowCount, intColumnCount, wdTableBehaviour9, wdAutofitBehaviour_FitContent)
                Dim intRow As Integer = 1

                With objTable
                    For Each selGridRow As ExportLogframeRow In objExportLogframe.TableGrid
                        If selGridRow.RowType = ExportLogframeRow.RowTypes.Section Then
                            If selGridRow.StructRtf IsNot Nothing Then _
                                .Cell(intRow, 1).Range.Text = selGridRow.Struct.Text
                            If objExportLogframe.IsResourceBudgetRow(selGridRow) = False Then
                                If objExportLogframe.ShowIndicatorColumn = True And String.IsNullOrEmpty(selGridRow.IndicatorText) = False Then _
                                    .Cell(intRow, intIndColIndex).Range.Text = selGridRow.IndicatorText
                                If objExportLogframe.ShowVerificationSourceColumn = True And String.IsNullOrEmpty(selGridRow.VerificationSourceText) = False Then _
                                    .Cell(intRow, intVerColIndex).Range.Text = selGridRow.VerificationSourceText
                            Else
                                If objExportLogframe.ShowIndicatorColumn = True And String.IsNullOrEmpty(selGridRow.ResourceText) = False Then _
                                    .Cell(intRow, intIndColIndex).Range.Text = selGridRow.ResourceText
                                If objExportLogframe.ShowVerificationSourceColumn = True And String.IsNullOrEmpty(selGridRow.BudgetText) = False Then _
                                    .Cell(intRow, intVerColIndex).Range.Text = selGridRow.BudgetText
                            End If
                            If objExportLogframe.ShowAssumptionColumn = True And String.IsNullOrEmpty(selGridRow.AssumptionText) = False Then _
                                .Cell(intRow, intAsmColIndex).Range.Text = selGridRow.AssumptionText

                            For i = intColumnCount To 2 Step -2
                                .Cell(intRow, i - 1).Merge(.Cell(intRow, i))
                            Next i

                            With .Rows(intRow).Range
                                .Font.Bold = True
                                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                                .Shading.ForegroundPatternColor = wdColorAutomatic
                                .Shading.BackgroundPatternColor = wdColor_Gray25
                            End With

                        ElseIf selGridRow.RowType = ExportLogframeRow.RowTypes.RepeatPurpose Or _
                            selGridRow.RowType = ExportLogframeRow.RowTypes.RepeatOutput Then

                            .Cell(intRow, 2).Merge(.Cell(intRow, intColumnCount))
                            .Rows(intRow).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft

                            Dim intColor As Integer
                            If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose Then _
                                intColor = 16776960 Else intColor = 16777164
                            With .Rows(intRow).Range
                                .Font.Bold = True
                                .Shading.ForegroundPatternColor = wdColorAutomatic
                                .Shading.BackgroundPatternColor = intColor
                            End With

                            If String.IsNullOrEmpty(selGridRow.StructSort) = False Then _
                                .Cell(intRow, 1).Range.Text = selGridRow.StructSort.ToString
                            If String.IsNullOrEmpty(selGridRow.StructText) = False Then _
                                .Cell(intRow, 2).Range.Text = selGridRow.StructText
                        Else
                            If String.IsNullOrEmpty(selGridRow.StructSort) = False Then _
                                .Cell(intRow, 1).Range.Text = selGridRow.StructSort
                            If String.IsNullOrEmpty(selGridRow.StructRtf) = False Then _
                                .Cell(intRow, 2).Range.Text = selGridRow.StructText
                            If objExportLogframe.IsResourceBudgetRow(selGridRow) = False Then
                                If String.IsNullOrEmpty(selGridRow.IndicatorSort) = False Then _
                                    .Cell(intRow, intIndColIndex).Range.Text = selGridRow.IndicatorSort
                                If String.IsNullOrEmpty(selGridRow.IndicatorText) = False Then _
                                    .Cell(intRow, intIndColIndex + 1).Range.Text = selGridRow.IndicatorText
                                If String.IsNullOrEmpty(selGridRow.VerificationSourceSort) = False Then _
                                    .Cell(intRow, intVerColIndex).Range.Text = selGridRow.VerificationSourceSort
                                If String.IsNullOrEmpty(selGridRow.VerificationSourceText) = False Then _
                                    .Cell(intRow, intVerColIndex + 1).Range.Text = selGridRow.VerificationSourceText
                            Else
                                If String.IsNullOrEmpty(selGridRow.ResourceSort) = False Then _
                                    .Cell(intRow, intIndColIndex).Range.Text = selGridRow.ResourceSort
                                If String.IsNullOrEmpty(selGridRow.ResourceText) = False Then _
                                    .Cell(intRow, intIndColIndex + 1).Range.Text = selGridRow.ResourceText
                                If String.IsNullOrEmpty(selGridRow.BudgetText) = False And objExportLogframe.ShowVerificationSourceColumn = True Then
                                    .Cell(intRow, intVerColIndex + 1).Range.Text = selGridRow.BudgetText
                                    .Cell(intRow, intVerColIndex + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                                End If
                            End If
                            If String.IsNullOrEmpty(selGridRow.AssumptionSort) = False Then _
                                    .Cell(intRow, intAsmColIndex).Range.Text = selGridRow.AssumptionSort
                            If String.IsNullOrEmpty(selGridRow.AssumptionText) = False Then _
                                    .Cell(intRow, intAsmColIndex + 1).Range.Text = selGridRow.AssumptionText
                        End If

                        intRow += 1
                    Next
                End With
            End Using
        End With
    End Sub
End Class
