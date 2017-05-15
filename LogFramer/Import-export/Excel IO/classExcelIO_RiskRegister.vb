Imports Excel = Microsoft.Office.Interop.Excel

Partial Public Class ExcelIO
    Public Sub ExportRiskRegisterToExcel(ByVal logFrame As LogFrame)
        Try
            InitialiseWorkBook()

            With My.Settings
                Dim objWorkSheetLookUp As Object
                objWorkSheetLookUp = WorkBook.Worksheets.Add(After:=WorkBook.Sheets(1))
                objWorkSheetLookUp.Name = LANG_Validation
                ExportRiskRegisterToExcel_LookUpTable(objWorkSheetLookUp)

                Dim objWorkSheet As Object = WorkBook.Worksheets(1)
                objWorkSheet.Name = LANG_RiskRegister
                ExportRiskRegisterToExcel_RiskCategory(logFrame, .setPrintRiskRegisterRiskCategory, objWorkSheet)

            End With
            FinaliseWorkBook()

            MsgBox(LANG_ExportToExcelComplete)
        Catch ex As Exception
            MsgBox(LANG_ExportToExcelError)
        End Try
    End Sub

#Region "Look-up table"
    Private Sub ExportRiskRegisterToExcel_LookUpTable(ByVal objWorkSheet As Object)

        Dim objTableGrid As Object = ExportRiskRegister_CreateLookUpGrid()
        Dim intColumnCount As Integer = 3
        Dim intRowCount As Integer = 12

        If objTableGrid IsNot Nothing Then

            ExportRiskRegister_SetColumnsLookUp(objWorkSheet)
            With objWorkSheet
                .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount)).Value = objTableGrid
            End With

            ExportRiskRegister_LookUpLayOut(objWorkSheet)
        End If
    End Sub

    Private Function ExportRiskRegister_CreateLookUpGrid() As Object
        Dim intColumnCount As Integer = 3
        Dim intRowCount As Integer = 12
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object

        objTableGrid(0, 0) = LANG_Likelihood
        For i = 0 To 4
            objTableGrid(i + 1, 0) = LIST_Likelihoods(i)
            objTableGrid(i + 1, 2) = i
        Next

        objTableGrid(6, 0) = LANG_Impact
        For i = 0 To 4
            Dim strImpact() = LIST_RiskImpacts(i).Split("-")
            objTableGrid(i + 7, 0) = Trim(strImpact(0))
            If strImpact.Length > 1 Then objTableGrid(i + 7, 1) = Trim(strImpact(1))
            objTableGrid(i + 7, 2) = i
        Next

        Return objTableGrid
    End Function

    Private Sub ExportRiskRegister_SetColumnsLookUp(ByVal objWorkSheet As Object)
        With objWorkSheet
            .Cells.VerticalAlignment = xlVerticalAlignment_Top
            '.Columns("A:A").NumberFormat = "@"

            With .Columns("A:B")
                .ColumnWidth = 40
                .WrapText = True
            End With
        End With
    End Sub

    Private Sub ExportRiskRegister_LookUpLayOut(ByVal objWorkSheet As Object)
        With objWorkSheet
            Dim intRow As Integer = 1

            'header rows
            With .Range(.Cells(1, 1), .Cells(1, 3))
                .Merge()
                .Font.Bold = True
                .HorizontalAlignment = xlHorizontalAlignment_Center
                .VerticalAlignment = xlVerticalAlignment_Center
                .WrapText = True
                With .Interior
                    .ColorIndex = 15
                    .Pattern = xlPattern_Solid
                End With
            End With

            With .Range(.Cells(7, 1), .Cells(7, 3))
                .Merge()
                .Font.Bold = True
                .HorizontalAlignment = xlHorizontalAlignment_Center
                .VerticalAlignment = xlVerticalAlignment_Center
                .WrapText = True
                With .Interior
                    .ColorIndex = 15
                    .Pattern = xlPattern_Solid
                End With
            End With

            'borders
            With .Range(.Cells(1, 1), .Cells(12, 3))
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

            .Columns("A").EntireColumn.AutoFit()
        End With
    End Sub
#End Region

#Region "Risk register"
    Private Sub ExportRiskRegisterToExcel_RiskCategory(ByVal logframe As LogFrame, ByVal intRiskCategory As Integer, ByVal objWorkSheet As Object)

        Using objExportRiskRegister As New ExportRiskRegister(logframe, intRiskCategory)
            objExportRiskRegister.LoadTable()

            Dim objTableGrid As Object = ExportRiskRegister_CreateTableGrid(objExportRiskRegister)
            Dim intColumnCount As Integer = objExportRiskRegister.GetColumnCount
            Dim intRowCount As Integer = objExportRiskRegister.GetRowCount + 2
            Dim intDeadlinesCount As Integer = objExportRiskRegister.GetDeadlinesCount

            If objTableGrid IsNot Nothing Then

                ExportRiskRegister_SetColumns(objWorkSheet, intDeadlinesCount)
                With objWorkSheet
                    .Range(.Cells(1, 1), .Cells(intRowCount, intColumnCount)).Value = objTableGrid
                End With
                ExportRiskRegister_Validation(objExportRiskRegister, objWorkSheet, intDeadlinesCount)
                ExportRiskRegister_LayOut(objWorkSheet, intRowCount, intColumnCount, intDeadlinesCount)
                ExportRiskRegister_LayOut_SectionTitles(objExportRiskRegister, objWorkSheet, intColumnCount)
            End If

        End Using
    End Sub

    Private Function ExportRiskRegister_CreateTableGrid(ByVal objExportRiskRegister As ExportRiskRegister) As Object
        Dim intColumnCount As Integer = objExportRiskRegister.GetColumnCount
        Dim intRowCount As Integer = objExportRiskRegister.GetRowCount + 2
        Dim intDeadlinesCount As Integer = objExportRiskRegister.GetDeadlinesCount
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object
        Dim strDate, strFormula As String
        Dim intLikelihoodIndex, intImpactIndex, intRiskLevelIndex As Integer
        Dim strLikelihoodColumnName, strImpactColumnName, strRiskLevelColumnName As String
        Dim intColumnIndex As Integer
        Dim intRowIndex As Integer = 2
        Dim intRowIndexExcel As Integer

        If intRowCount = 0 Then Return Nothing

        'Title row
        objTableGrid(0, 0) = LANG_Risks
        objTableGrid(0, 2) = LANG_RiskResponse
        objTableGrid(0, 3) = LANG_ResponseStrategy
        objTableGrid(0, 4) = LANG_Objectives
        objTableGrid(0, 6) = LANG_Baseline
        objTableGrid(1, 6) = LANG_Likelihood
        objTableGrid(1, 7) = LANG_Impact
        objTableGrid(1, 8) = LANG_RiskLevel

        intColumnIndex = 9
        If objExportRiskRegister.RiskMonitoring IsNot Nothing Then
            For Each selDeadline As RiskMonitoringDeadline In objExportRiskRegister.RiskMonitoring.RiskMonitoringDeadlines
                Select Case objExportRiskRegister.RiskMonitoring.Repetition
                    Case RiskMonitoring.RepetitionOptions.Monthly, RiskMonitoring.RepetitionOptions.Quarterly, RiskMonitoring.RepetitionOptions.TwiceYear
                        strDate = selDeadline.ExactDeadline.ToString("MMM-yyyy")
                    Case RiskMonitoring.RepetitionOptions.SingleMoment, RiskMonitoring.RepetitionOptions.Yearly
                        strDate = selDeadline.ExactDeadline.ToString("yyyy")
                    Case Else
                        strDate = selDeadline.ExactDeadline.ToShortDateString
                End Select

                objTableGrid(0, intColumnIndex) = strDate
                objTableGrid(1, intColumnIndex) = LANG_Likelihood
                intColumnIndex += 1
                objTableGrid(1, intColumnIndex) = LANG_Impact
                intColumnIndex += 1
                objTableGrid(1, intColumnIndex) = LANG_RiskLevel
                intColumnIndex += 1
            Next
        End If

        For Each selGridRow As ExportRiskRegisterRow In objExportRiskRegister.TableGrid
            If String.IsNullOrEmpty(selGridRow.SectionTitle) = False Then
                objTableGrid(intRowIndex, 0) = selGridRow.SectionTitle
            Else
                Dim selAssumption As Assumption = selGridRow.Assumption
                objTableGrid(intRowIndex, 0) = selGridRow.AssumptionSortNumber
                objTableGrid(intRowIndex, 1) = selGridRow.AssumptionText

                If selAssumption IsNot Nothing Then
                    objTableGrid(intRowIndex, 2) = selAssumption.RiskDetail.RiskResponseText
                    objTableGrid(intRowIndex, 3) = selAssumption.ResponseStrategy
                End If
                objTableGrid(intRowIndex, 4) = selGridRow.StructSortNumber
                objTableGrid(intRowIndex, 5) = selGridRow.StructText
                If selAssumption IsNot Nothing Then
                    objTableGrid(intRowIndex, 6) = selGridRow.LikelihoodText
                    objTableGrid(intRowIndex, 7) = selGridRow.ImpactText
                    objTableGrid(intRowIndex, 8) = selAssumption.RiskDetail.RiskLevel
                End If

                For i = 0 To intDeadlinesCount
                    intRowIndexExcel = intRowIndex + 1
                    intColumnIndex = 6 + (i * 3)
                    intLikelihoodIndex = intColumnIndex
                    intImpactIndex = intColumnIndex + 1
                    intRiskLevelIndex = intColumnIndex + 2

                    strLikelihoodColumnName = GetColumnName(intLikelihoodIndex + 1)
                    strImpactColumnName = GetColumnName(intImpactIndex + 1)
                    strRiskLevelColumnName = GetColumnName(intRiskLevelIndex + 1)

                    'VLOOKUP($G$4;Validation!$A$2:$C$6;3;FALSE) * VLOOKUP($H$4;Validation!$A$8:$C$12;3;FALSE) / 16
                    strFormula = String.Format("=IFERROR(VLOOKUP(${2}${1},{0}!$A$2:$C$6,3,FALSE) * VLOOKUP(${3}${1},{0}!$A$8:$C$12,3,FALSE) / 16,""-"")", LANG_Validation, intRowIndexExcel, strLikelihoodColumnName, strImpactColumnName)
                    objTableGrid(intRowIndex, intColumnIndex + 2) = strFormula
                Next

            End If

            intRowIndex += 1
        Next

        Return objTableGrid
    End Function

    Private Sub ExportRiskRegister_SetColumns(ByVal objWorkSheet As Object, ByVal intDeadlinesCount As Integer)
        Dim intFirstIndex, intColumnIndex As Integer
        Dim strColumnRange As String

        With objWorkSheet
            .Cells.VerticalAlignment = xlVerticalAlignment_Top
            .Columns("A:A").NumberFormat = "@"
            .Columns("E:E").NumberFormat = "@"

            With .Columns("B:B")
                .ColumnWidth = 40
                .WrapText = True
            End With
            With .Columns("C:C")
                .ColumnWidth = 20
                .WrapText = True
            End With
            With .Columns("D:D")
                .ColumnWidth = 40
                .WrapText = True
            End With
            With .Columns("F:F")
                .ColumnWidth = 40
                .WrapText = True
            End With

            For i = 0 To intDeadlinesCount
                intFirstIndex = 7 + (i * 3)
                intColumnIndex = intFirstIndex + 2

                strColumnRange = GetColumnRange(intFirstIndex, intColumnIndex)

                With .Columns(strColumnRange)
                    .ColumnWidth = 10
                End With

                strColumnRange = GetColumnRange(intColumnIndex)

                With .Columns(strColumnRange)
                    .NumberFormat = "0%"
                End With
            Next
        End With
    End Sub

    Private Sub ExportRiskRegister_Validation(ByVal objExportRiskRegister As ExportRiskRegister, ByVal objWorkSheet As Object, ByVal intDeadlinesCount As Integer)
        Dim intRowIndex, intColumnIndex As Integer
        Dim strFormula As String

        intRowIndex = 3
        With objWorkSheet
            For Each selGridRow As ExportRiskRegisterRow In objExportRiskRegister.TableGrid
                If String.IsNullOrEmpty(selGridRow.SectionTitle) Then
                    Dim selAssumption As Assumption = selGridRow.Assumption

                    For i = 0 To intDeadlinesCount - 1
                        intColumnIndex = 7 + (i * 3)
                        strFormula = String.Format("={0}!$A$2:$A$6", LANG_Validation)
                        ExportRiskRegister_Validation_Cell(.Cells(intRowIndex, intColumnIndex), strFormula)

                        intColumnIndex += 1
                        strFormula = String.Format("={0}!$A$8:$A$12", LANG_Validation)
                        ExportRiskRegister_Validation_Cell(.Cells(intRowIndex, intColumnIndex), strFormula)
                    Next
                End If

                intRowIndex += 1
            Next
        End With
    End Sub

    Private Sub ExportRiskRegister_Validation_Cell(ByVal objCell As Object, ByVal strFormula As String)
        With objCell.Validation
            .Delete()
            .Add(Type:=xlDVType_ValidateList, AlertStyle:=xlDVAlertStyle_Stop, Operator:=xlFormatConditionOperator_Between, Formula1:=strFormula)
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End Sub

    Private Sub ExportRiskRegister_LayOut(ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer, ByVal intDeadlinesCount As Integer)
        Dim intColumnIndex As Integer

        With objWorkSheet
            Dim intRow As Integer = 1

            'header row
            With .Range(.Cells(intRow, 1), .Cells(intRow + 1, intColumnCount))
                .Font.Bold = True
                .HorizontalAlignment = xlHorizontalAlignment_Center
                .VerticalAlignment = xlVerticalAlignment_Center
                .WrapText = True
            End With

            .Range(.Cells(intRow, 1), .Cells(intRow + 1, 2)).Merge()
            .Range(.Cells(intRow, 3), .Cells(intRow + 1, 3)).Merge()
            .Range(.Cells(intRow, 4), .Cells(intRow + 1, 4)).Merge()
            .Range(.Cells(intRow, 5), .Cells(intRow + 1, 6)).Merge()

            For i = 0 To intDeadlinesCount
                intColumnIndex = 7 + (i * 3)

                .Range(.Cells(intRow, intColumnIndex), .Cells(intRow, intColumnIndex + 2)).Merge()
            Next

            With .Range(.Cells(intRow, 1), .Cells(intRow + 1, intColumnCount)).Interior
                .ColorIndex = 15
                .Pattern = xlPattern_Solid
            End With
            .Rows("1:2").EntireRow.AutoFit()

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

            For i = 0 To intDeadlinesCount
                intColumnIndex = 7 + (i * 3)

                With .Range(.Cells(1, intColumnIndex), .Cells(intRowCount, intColumnIndex))
                    With .Cells.Borders(xlBorders_EdgeLeft)
                        .LineStyle = xlLineStyle_Continuous
                        .Weight = xlBorderWeight_Medium
                        .ColorIndex = xlColorIndex_ColorIndexAutomatic
                    End With
                End With
            Next
            '.Columns(strTargetsRange).EntireColumn.AutoFit()
        End With
    End Sub

    Private Sub ExportRiskRegister_LayOut_SectionTitles(ByVal objExportRiskRegister As ExportRiskRegister, ByVal objWorkSheet As Object, ByVal intColumnCount As Integer)
        Dim intRowIndex As Integer = 3

        With objWorkSheet
            For Each selGridRow As ExportRiskRegisterRow In objExportRiskRegister.TableGrid
                If String.IsNullOrEmpty(selGridRow.SectionTitle) = False Then
                    With .Range(.Cells(intRowIndex, 1), .Cells(intRowIndex, intColumnCount))
                        .Merge()
                        .Font.Bold = True
                        With .Interior
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = 0.599993896298105
                        End With
                    End With
                End If

                intRowIndex += 1
            Next
        End With
    End Sub
#End Region
End Class
