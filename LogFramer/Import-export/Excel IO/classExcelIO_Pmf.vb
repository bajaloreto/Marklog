Partial Public Class ExcelIO
    Public Sub ExportPmfToExcel(ByVal logframe As LogFrame)
        'Try
        InitialiseWorkBook()

        With My.Settings
            If .setPrintPmfSection = PrintSettingsPmf.PrintSections.printall Then
                For intSection = 0 To 3
                    Dim objWorkSheet As Object

                    objWorkSheet = WorkBook.Worksheets.Add(After:=WorkBook.Sheets(WorkBook.Sheets.Count))

                    objWorkSheet.Name = logframe.StructNamePlural(intSection)

                    ExportPmfToExcel_Section(logframe, intSection, .setPrintPmfShowTargetRowTitles, objWorkSheet)
                Next
            Else
                Dim objWorkSheet As Object = WorkBook.Worksheets(1)

                ExportPmfToExcel_Section(logframe, .setPrintPmfSection, .setPrintPmfShowTargetRowTitles, objWorkSheet)

            End If

        End With
        FinaliseWorkBook()

        MsgBox(LANG_ExportToExcelComplete)
        'Catch ex As Exception
        '    MsgBox(LANG_ExportToExcelError)
        'End Try

    End Sub

    Public Sub ExportPmfToExcel_Section(ByVal logframe As LogFrame, ByVal intSection As Integer, ByVal boolShowTargetRowTitles As Boolean, ByVal objWorkSheet As Object)

        Using objExportPmf As New ExportPmf(logframe, intSection, boolShowTargetRowTitles)
            objExportPmf.LoadTable()

            Dim objTableGrid As Object = ExportPmf_CreateTableGrid(objExportPmf)
            Dim intColumnCount As Integer = objExportPmf.GetColumnCount
            Dim intRowCount As Integer = objExportPmf.GetRowCount + 1
            Dim intTargetCount As Integer = objExportPmf.GetTargetColumnCount

            If objTableGrid IsNot Nothing Then

                ExportPmf_SetColumns(objWorkSheet, intTargetCount)
                With objWorkSheet
                    .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount)).Value = objTableGrid
                End With

                ExportPmf_LayOut(objWorkSheet, intRowCount, intColumnCount, intTargetCount)
                ExportPmf_LayOut_Rows(objExportPmf, objWorkSheet, intRowCount, intColumnCount, intTargetCount)
            End If

        End Using
    End Sub

#Region "Table grid"
    Private Function ExportPmf_CreateTableGrid(ByVal objExportPmf As ExportPmf) As Object
        Dim intColumnCount As Integer = objExportPmf.GetColumnCount + 1
        Dim intRowCount As Integer = objExportPmf.GetRowCount + 1
        Dim intTargetCount As Integer = objExportPmf.GetTargetColumnCount
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object
        Dim intColumnIndex As Integer
        Dim intRowIndex As Integer = 1

        If intRowCount = 0 Then Return Nothing

        'Title row
        objTableGrid = ExportPmf_CreateTableGrid_TitleRow(objExportPmf, objTableGrid)

        For Each selGridRow As ExportPmfRow In objExportPmf.TableGrid
            Dim selIndicator As Indicator = selGridRow.Indicator

            objTableGrid(intRowIndex, 0) = selGridRow.StructSort
            objTableGrid(intRowIndex, 1) = selGridRow.StructText
            objTableGrid(intRowIndex, 2) = selGridRow.IndicatorSort
            objTableGrid(intRowIndex, 3) = selGridRow.IndicatorText

            If selIndicator IsNot Nothing Then
                Select Case selIndicator.QuestionType
                    Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Image, Indicator.QuestionTypes.ImageWithTargets
                        'no targets
                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                        If selIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                            objTableGrid = ExportPmf_CreateTableGrid_TargetValues(objTableGrid, selIndicator, intTargetCount, intRowIndex)
                        Else
                            objTableGrid = ExportPmf_CreateTableGrid_TargetScores(objTableGrid, selIndicator, intTargetCount, intRowIndex)
                        End If
                    Case Else
                        objTableGrid = ExportPmf_CreateTableGrid_TargetScores(objTableGrid, selIndicator, intTargetCount, intRowIndex)
                End Select
            End If

            intColumnIndex = intTargetCount + 6
            objTableGrid(intRowIndex, intColumnIndex) = selGridRow.VerificationSourceSort
            intColumnIndex += 1
            objTableGrid(intRowIndex, intColumnIndex) = selGridRow.VerificationSourceText
            intColumnIndex += 1
            If selGridRow.VerificationSource IsNot Nothing Then
                objTableGrid(intRowIndex, intColumnIndex) = selGridRow.VerificationSource.CollectionMethod
            End If
            intColumnIndex += 1
            objTableGrid(intRowIndex, intColumnIndex) = selGridRow.Frequency
            intColumnIndex += 1
            objTableGrid(intRowIndex, intColumnIndex) = selGridRow.Responsibility

            If selIndicator IsNot Nothing Then
                Select Case selIndicator.QuestionType
                    Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Image, Indicator.QuestionTypes.ImageWithTargets
                        intRowIndex += 1
                    Case Else
                        If selIndicator IsNot Nothing AndAlso selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                            intRowIndex += 2
                        Else
                            intRowIndex += 1
                        End If
                End Select
            Else
                intRowIndex += 1
            End If
        Next

        Return objTableGrid
    End Function

    Private Function ExportPmf_CreateTableGrid_TargetValues(ByVal objTableGrid(,) As Object, ByVal selIndicator As Indicator, ByVal intTargetCount As Integer, ByVal intRowIndex As Integer) As Object
        Dim strTargetType As String = String.Empty, strTargetTypeBeneficiary As String = String.Empty
        Dim objBaseline As Object
        Dim strTarget(intTargetCount) As String, strPopulationTarget(intTargetCount) As String
        Dim intColumnIndex As Integer = 6

        If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            strTargetType = LANG_TargetValueBeneficiary
        Else
            strTargetType = LANG_Target
        End If
        objBaseline = selIndicator.GetBaselineTotalValue

        objTableGrid(intRowIndex, 4) = strTargetType
        objTableGrid(intRowIndex, 5) = objBaseline

        For intTargetIndex = 0 To selIndicator.Targets.Count - 1
            Select Case selIndicator.TargetSystem
                Case Indicator.TargetSystems.ValueRange
                    strTarget(intTargetIndex) = selIndicator.GetTargetFormattedScore(intTargetIndex)
                Case Else
                    strTarget(intTargetIndex) = selIndicator.GetTargetValuesTotal(intTargetIndex).MinValue
            End Select

            objTableGrid(intRowIndex, intColumnIndex) = strTarget(intTargetIndex)
            intColumnIndex += 1
        Next

        If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            objTableGrid = ExportPmf_CreateTableGrid_PopulationTargetValues(objTableGrid, selIndicator, intTargetCount, intRowIndex)
        End If

        Return objTableGrid
    End Function

    Private Function ExportPmf_CreateTableGrid_TargetScores(ByVal objTableGrid(,) As Object, ByVal selIndicator As Indicator, ByVal intTargetCount As Integer, ByVal intRowIndex As Integer) As Object
        Dim strTargetType As String = String.Empty, strTargetTypeBeneficiary As String = String.Empty
        Dim objBaseline As Object
        Dim strTarget(intTargetCount) As String
        Dim intColumnIndex As Integer = 6

        If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            strTargetType = LANG_ScoreValueBeneficiary
        Else
            strTargetType = LANG_Target
        End If
        objBaseline = selIndicator.GetBaselineTotalScore

        objTableGrid(intRowIndex, 4) = strTargetType
        objTableGrid(intRowIndex, 5) = objBaseline

        For intTargetIndex = 0 To selIndicator.Targets.Count - 1
            strTarget(intTargetIndex) = selIndicator.GetTargetScoresTotal(intTargetIndex).Score

            objTableGrid(intRowIndex, intColumnIndex) = strTarget(intTargetIndex)
            intColumnIndex += 1
        Next

        If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            objTableGrid = ExportPmf_CreateTableGrid_PopulationTargetScores(objTableGrid, selIndicator, intTargetCount, intRowIndex)
        End If

        Return objTableGrid
    End Function

    Private Function ExportPmf_CreateTableGrid_PopulationTargetValues(ByVal objTableGrid(,) As Object, ByVal selIndicator As Indicator, ByVal intTargetCount As Integer, ByVal intRowIndex As Integer) As Object
        Dim intColumnIndex As Integer = 6

        objTableGrid(intRowIndex + 1, 4) = LANG_ScoreValueTargetGroup
        objTableGrid(intRowIndex + 1, 5) = selIndicator.GetPopulationBaselineValue

        If intTargetCount > 0 Then
            For intTargetIndex = 0 To intTargetCount - 1
                objTableGrid(intRowIndex + 1, intColumnIndex) = selIndicator.GetPopulationTargetValues(intTargetIndex).MinValue
                intColumnIndex += 1
            Next
        End If

        Return objTableGrid
    End Function

    Private Function ExportPmf_CreateTableGrid_PopulationTargetScores(ByVal objTableGrid(,) As Object, ByVal selIndicator As Indicator, ByVal intTargetCount As Integer, ByVal intRowIndex As Integer) As Object
        Dim intColumnIndex As Integer = 6

        objTableGrid(intRowIndex + 1, 4) = LANG_ScoreValueTargetGroup
        objTableGrid(intRowIndex + 1, 5) = selIndicator.GetPopulationBaselineScore

        If intTargetCount > 0 Then
            For intTargetIndex = 0 To intTargetCount - 1
                objTableGrid(intRowIndex + 1, intColumnIndex) = selIndicator.GetPopulationTargetScores(intTargetIndex).Score
                intColumnIndex += 1
            Next
        End If

        Return objTableGrid
    End Function

    Private Function ExportPmf_CreateTableGrid_TitleRow(ByVal objExportPmf As ExportPmf, ByVal objTableGrid(,) As Object) As Object
        Dim strObjectives As String = objExportPmf.GetColumnTitleObjectives()
        Dim intColumnIndex As Integer
        Dim strDate As String

        objTableGrid(0, 0) = strObjectives
        objTableGrid(0, 2) = LANG_Indicators
        objTableGrid(0, 4) = LANG_TargetType
        objTableGrid(0, 5) = LANG_Baseline

        intColumnIndex = 6
        If objExportPmf.TargetDeadlinesSection IsNot Nothing Then
            For Each selTargetDeadline As TargetDeadline In objExportPmf.TargetDeadlinesSection.TargetDeadlines
                Select Case objExportPmf.TargetDeadlinesSection.Repetition
                    Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                        strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
                    Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                        strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
                    Case Else
                        strDate = selTargetDeadline.ExactDeadline.ToShortDateString
                End Select

                objTableGrid(0, intColumnIndex) = String.Format("{0} {1}", LANG_Target, strDate)

                intColumnIndex += 1
            Next
        End If
        objTableGrid(0, intColumnIndex) = LANG_VerificationSources
        intColumnIndex += 2
        objTableGrid(0, intColumnIndex) = LANG_DataCollection
        intColumnIndex += 1
        objTableGrid(0, intColumnIndex) = LANG_Frequency
        intColumnIndex += 1
        objTableGrid(0, intColumnIndex) = LANG_Responsibility

        Return objTableGrid
    End Function
#End Region

#Region "Lay-out"
    Private Sub ExportPmf_SetColumns(ByVal objWorkSheet As Object, ByVal intTargetCount As Integer)
        Dim strVerSortColumnRange As String = GetColumnRange(intTargetCount + 7)
        Dim strVerColumnRange As String = GetColumnRange(intTargetCount + 8)
        Dim strCollectionColumnRange As String = GetColumnRange(intTargetCount + 9)
        Dim strFrequencyColumnRange As String = GetColumnRange(intTargetCount + 10)
        Dim strResponsibleColumnRange As String = GetColumnRange(intTargetCount + 11)

        With objWorkSheet
            .Cells.VerticalAlignment = xlVerticalAlignment_Top
            .Columns("A:A").NumberFormat = "@"
            .Columns("C:C").NumberFormat = "@"
            .Columns(strVerSortColumnRange).NumberFormat = "@"

            With .Columns("B:B")
                .ColumnWidth = 40
                .WrapText = True
            End With
            With .Columns("D:D")
                .ColumnWidth = 60
                .WrapText = True
            End With
            With .Columns("E:E")
                .ColumnWidth = 20
                .WrapText = True
            End With
            With .Columns(strVerColumnRange)
                .ColumnWidth = 40
                .WrapText = True
            End With
            With .Columns(strCollectionColumnRange)
                .ColumnWidth = 40
                .WrapText = True
            End With
            With .Columns(strFrequencyColumnRange)
                .ColumnWidth = 20
                .WrapText = True
            End With
            With .Columns(strResponsibleColumnRange)
                .ColumnWidth = 20
                .WrapText = True
            End With
        End With
    End Sub

    Private Sub ExportPmf_LayOut(ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer, ByVal intTargetCount As Integer)
        Dim strTargetsRange As String = GetColumnRange(5, 7 + intTargetCount)

        With objWorkSheet
            Dim intRow As Integer = 1

            .Columns(strTargetsRange).HorizontalAlignment = xlHorizontalAlignment_Right

            'header row
            With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount))
                .Font.Bold = True
                .HorizontalAlignment = xlHorizontalAlignment_Center
                .VerticalAlignment = xlVerticalAlignment_Center
                .WrapText = True
            End With
            .Range(.Cells(intRow, 1), .Cells(intRow, 2)).Merge()
            .Range(.Cells(intRow, 3), .Cells(intRow, 4)).Merge()
            .Range(.Cells(intRow, intTargetCount + 7), .Cells(intRow, intTargetCount + 8)).Merge()
            With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount)).Interior
                .ColorIndex = 15
                .Pattern = xlPattern_Solid
            End With
            .Rows("1:1").EntireRow.AutoFit()

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

            .Columns(strTargetsRange).EntireColumn.AutoFit()
        End With
    End Sub

    Private Sub ExportPmf_LayOut_Rows(ByVal objExportPmfTool As ExportPmf, ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer, ByVal intTargetCount As Integer)
        Dim selRow As ExportPmfRow
        Dim intRow As Integer = 2
        Dim intIndex As Integer
        Dim selIndicator As Indicator
        Dim intFirstIndex As Integer = 6
        Dim intLastIndex As Integer = intTargetCount + 6
        Dim strNumberFormat As String = String.Empty
        Dim strNumberFormatPopulation As String = String.Empty
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        With objWorkSheet
            For intIndex = 0 To objExportPmfTool.TableGrid.Count - 1
                selRow = objExportPmfTool.TableGrid(intIndex)

                If selRow.Indicator IsNot Nothing Then
                    selIndicator = selRow.Indicator

                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Image, Indicator.QuestionTypes.ImageWithTargets
                            'no targets
                            'Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                            '    If selIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                            '        objTableGrid = ExportPmf_CreateTableGrid_TargetValues(objTableGrid, selIndicator, intTargetCount, intRowIndex)
                            '    Else
                            '        objTableGrid = ExportPmf_CreateTableGrid_TargetScores(objTableGrid, selIndicator, intTargetCount, intRowIndex)
                            '    End If
                        Case Else
                            If selIndicator.ValuesDetail IsNot Nothing Then
                                intNrDecimals = selIndicator.ValuesDetail.NrDecimals
                                strUnit = selIndicator.ValuesDetail.Unit
                            End If
                            Select Case selIndicator.ScoringSystem
                                Case Indicator.ScoringSystems.Value
                                    strNumberFormat = GetNumberFormat(intNrDecimals, strUnit)
                                    strNumberFormatPopulation = strNumberFormat
                                Case Indicator.ScoringSystems.Percentage
                                    strNumberFormat = GetNumberFormat(intNrDecimals, "%")
                                    strNumberFormatPopulation = strNumberFormat
                                Case Indicator.ScoringSystems.Score
                                    strNumberFormat = GetNumberFormatScore(intNrDecimals, selIndicator.GetTargetMaximumScore)
                                    strNumberFormatPopulation = GetNumberFormatScore(intNrDecimals, selIndicator.GetPopulationBaselineMaximumScore)
                            End Select

                            If String.IsNullOrEmpty(strNumberFormat) = False Then
                                .Range(.Cells(intRow, intFirstIndex), .Cells(intRow, intLastIndex)).NumberFormat = strNumberFormat
                                If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                                    .Range(.Cells(intRow + 1, intFirstIndex), .Cells(intRow + 1, intLastIndex)).NumberFormat = strNumberFormatPopulation
                                End If
                            End If
                    End Select

                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Image, Indicator.QuestionTypes.ImageWithTargets
                            intRow += 1
                        Case Else
                            If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                                intRow += 2
                            Else
                                intRow += 1
                            End If
                    End Select
                Else
                    intRow += 1
                End If
            Next
        End With
    End Sub
#End Region
End Class
