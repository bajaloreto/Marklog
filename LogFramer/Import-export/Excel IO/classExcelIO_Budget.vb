Partial Public Class ExcelIO
    Public Sub ExportBudgetToExcel(ByVal logframe As LogFrame)
        Dim intBudgetYearIndex As Integer
        Dim intBudgetYearsCount As Integer = logframe.Budget.BudgetYears.Count
        Dim objTableGrids As New List(Of Object)
        Dim intColumnCount(intBudgetYearsCount - 1) As Integer
        Dim intRowCount(intBudgetYearsCount - 1) As Integer
        Dim strBudgetTitle(intBudgetYearsCount - 1) As String
        Dim selBudgetYear As BudgetYear

        InitialiseWorkBook()

        'Create tables for each year (if the budget is a multi-year budget)
        If logframe.Budget.MultiYearBudget = True Then
            For intBudgetYearIndex = 1 To intBudgetYearsCount - 1
                selBudgetYear = logframe.Budget.BudgetYears(intBudgetYearIndex)

                Using objExportBudget As New ExportBudget(logframe, intBudgetYearIndex)
                    objExportBudget.LoadTable()
                    intColumnCount(intBudgetYearIndex) = objExportBudget.GetColumnCount
                    intRowCount(intBudgetYearIndex) = objExportBudget.GetRowCount + 1
                    strBudgetTitle(intBudgetYearIndex) = selBudgetYear.BudgetYear.ToString("yyyy")

                    Dim objTableGrid(intRowCount(intBudgetYearIndex) - 1, intColumnCount(intBudgetYearIndex) - 1) As Object

                    objTableGrid = ExportBudget_CreateTableGrid(objExportBudget)
                    ExportBudget_CreateFormulas(objTableGrid)

                    objTableGrids.Add(objTableGrid)
                End Using
            Next
        End If

        'create a table for the totals
        Using objExportBudget As New ExportBudget(logframe, 0)
            objExportBudget.LoadTable()
            intColumnCount(0) = objExportBudget.GetColumnCount
            intRowCount(0) = objExportBudget.GetRowCount + 1
            strBudgetTitle(0) = LANG_TotalBudget

            Dim objTableGrid(intRowCount(0) - 1, intColumnCount(0) - 1) As Object

            objTableGrid = ExportBudget_CreateTableGrid(objExportBudget)
            objTableGrids.Insert(0, objTableGrid)

            If logframe.Budget.MultiYearBudget = True Then
                ExportBudget_CreateReferences(objTableGrids, strBudgetTitle)
            End If
            ExportBudget_CreateFormulas(objTableGrid)
        End Using

        ExportBudget_TableGrids_Finalise(objTableGrids)

        'export tables to different worksheets
        For intBudgetYearIndex = intBudgetYearsCount - 1 To 0 Step -1
            selBudgetYear = logframe.Budget.BudgetYears(intBudgetYearIndex)
            Dim objTableGrid As Object = objTableGrids(intBudgetYearIndex)

            If objTableGrid IsNot Nothing Then
                Dim objWorkSheet As Object

                objWorkSheet = WorkBook.Worksheets.Add(Before:=WorkBook.Sheets(1))
                'objWorkSheet = WorkBook.Worksheets.Add(After:=WorkBook.Sheets(WorkBook.Sheets.Count))

                objWorkSheet.Name = strBudgetTitle(intBudgetYearIndex)

                ExportBudget_SetColumns(objWorkSheet)

                With objWorkSheet
                    .Range(.Cells(1, 1), .Cells(intRowCount(intBudgetYearIndex), intColumnCount(intBudgetYearIndex))).Value = objTableGrid
                End With

                ExportBudget_LayOut(objWorkSheet, intRowCount(intBudgetYearIndex), intColumnCount(intBudgetYearIndex))
            End If
        Next

        ExportBudget_RemoveColumns(intBudgetYearsCount, logframe.Budget.MultiYearBudget)

        FinaliseWorkBook()
    End Sub

    Private Function ExportBudget_CreateTableGrid(ByVal objExportBudget As ExportBudget) As Object
        Dim intColumnCount As Integer = objExportBudget.GetColumnCount
        Dim intRowCount As Integer = objExportBudget.GetRowCount
        Dim intRow As Integer = 1

        If intRowCount = 0 Then Return Nothing

        Dim objTableGrid(intRowCount, intColumnCount - 1) As Object

        'Title row
        objTableGrid(0, 0) = LANG_Number
        objTableGrid(0, 1) = LANG_Description
        objTableGrid(0, 2) = LANG_Duration
        objTableGrid(0, 3) = LANG_Unit
        objTableGrid(0, 4) = LANG_Quantity
        objTableGrid(0, 5) = LANG_Unit
        objTableGrid(0, 6) = LANG_UnitCost
        objTableGrid(0, 8) = LANG_TotalLocalCost
        objTableGrid(0, 10) = LANG_TotalCost

        'Values
        For Each selGridRow As ExportBudgetRow In objExportBudget.TableGrid
            objTableGrid(intRow, 0) = selGridRow.SortNumber
            objTableGrid(intRow, 1) = selGridRow.BudgetItem.Text

            If selGridRow.DurationSet Then
                objTableGrid(intRow, 2) = selGridRow.Duration
                objTableGrid(intRow, 3) = selGridRow.DurationUnit
            End If
            If selGridRow.NumberSet Then
                objTableGrid(intRow, 4) = selGridRow.Number
                objTableGrid(intRow, 5) = selGridRow.NumberUnit
            End If
            If selGridRow.BudgetItem.BudgetItems.Count = 0 Then
                objTableGrid(intRow, 6) = selGridRow.UnitCost.Amount
                objTableGrid(intRow, 7) = selGridRow.UnitCost.CurrencyCode
            Else
                objTableGrid(intRow, 6) = "SUM"
            End If
            objTableGrid(intRow, 8) = selGridRow.TotalLocalCost.Amount
            objTableGrid(intRow, 9) = selGridRow.TotalLocalCost.CurrencyCode
            objTableGrid(intRow, 10) = selGridRow.TotalCost.Amount
            objTableGrid(intRow, 11) = selGridRow.TotalCost.CurrencyCode

            intRow += 1
        Next

        Return objTableGrid
    End Function

    Private Sub ExportBudget_SetColumns(ByVal objWorkSheet As Object)
        With objWorkSheet
            .Cells.VerticalAlignment = xlVerticalAlignment_Top
            .Columns("A:A").NumberFormat = "@"
            .Columns("G:G").NumberFormat = "#,##0.00"
            .Columns("I:I").NumberFormat = "#,##0.00"
            .Columns("K:K").NumberFormat = "#,##0.00"

            With .Columns("B:B")
                .ColumnWidth = 60
                .WrapText = True
            End With
        End With
    End Sub

    Private Function ExportBudget_CreateFormulas(ByVal objTableGrid As Object(,)) As Object

        Dim strFormula, strValue As String
        Dim objValue As Object
        Dim intRowCount As Integer = objTableGrid.GetLength(0)
        Dim strSortNumber As String

        For i = 1 To intRowCount - 1
            objValue = objTableGrid(i, 6)
            strValue = objValue.ToString

            If strValue = "SUM" Then
                objValue = objTableGrid(i, 0)
                strSortNumber = objValue.ToString

                objValue = objTableGrid(i, 4)
                If objValue <> Nothing Then

                    strFormula = ExportBudget_ParseSums(objTableGrid, strSortNumber, i, intRowCount, 4)
                    If String.IsNullOrEmpty(strFormula) = False Then objTableGrid(i, 4) = strFormula
                End If

                strFormula = ExportBudget_ParseSums(objTableGrid, strSortNumber, i, intRowCount, 8)
                If String.IsNullOrEmpty(strFormula) = False Then objTableGrid(i, 8) = strFormula

                strFormula = ExportBudget_ParseSums(objTableGrid, strSortNumber, i, intRowCount, 10)
                If String.IsNullOrEmpty(strFormula) = False Then objTableGrid(i, 10) = strFormula
                'objTableGrid(i, 6) = Nothing
            End If
        Next
        Return objTableGrid
    End Function

    Private Function ExportBudget_ParseSums(ByVal objTableGrid As Object(,), ByVal strParentSortNumber As String, ByVal intStartRowIndex As Integer, ByVal intRowCount As Integer, ByVal intColumnIndex As Integer) As String
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
                    alRows.Add(intRowIndex)
                End If
            End If
        Next

        If alRows.Count > 0 Then
            strFormula = ExportBudget_ParseSums_Parse(alRows, intColumnIndex)
        End If

        Return strFormula
    End Function

    Private Function ExportBudget_ParseSums_Parse(ByVal alRows As ArrayList, ByVal intColumnIndex As Integer, Optional ByVal strFormula As String = "") As String
        Dim intCurrent, intNext As Integer
        Dim intIndex, intCounter As Integer
        Dim strColumn As String = Chr(intColumnIndex + 65)

        If strFormula = String.Empty Then
            strFormula = "="
        Else
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
                    strFormula = ExportBudget_ParseSums_Parse(alRows, intColumnIndex, strFormula)
                End If
            ElseIf intCounter > 0 Then
                'if there are consecutive rows: SUM(A1:A3)
                strFormula &= String.Format("SUM({0}{1}:{0}{2})", strColumn, alRows(0) + 1, alRows(intCounter) + 1)
                alRows.RemoveRange(0, intCounter + 1)

                If alRows.Count > 0 Then
                    strFormula = ExportBudget_ParseSums_Parse(alRows, intColumnIndex, strFormula)
                End If
            End If
        ElseIf alRows.Count = 1 Then
            strFormula &= String.Format("{0}{1}", strColumn, alRows(0) + 1)
        End If

        Return strFormula
    End Function

    Private Function ExportBudget_CreateReferences(ByVal objTableGrids As List(Of Object), ByVal strBudgetTitle() As String) As Object
        Dim strReference, strValue As String
        Dim objValue As Object
        Dim strSortNumber As String

        Dim objTableGridTotal As Object = objTableGrids(0)
        Dim intRowCount As Integer = objTableGridTotal.GetLength(0)

        For intRowIndex = 1 To intRowCount - 1
            objValue = objTableGridTotal(intRowIndex, 6)
            strValue = objValue.ToString

            If strValue <> "SUM" Then
                objValue = objTableGridTotal(intRowIndex, 0)
                If objValue IsNot Nothing Then
                    strSortNumber = objValue.ToString

                    strReference = ExportBudget_CreateReferences_Reference(objTableGrids, strSortNumber, strBudgetTitle)
                    If String.IsNullOrEmpty(strReference) = False Then objTableGridTotal(intRowIndex, 10) = strReference
                End If
            End If
        Next

        Return objTableGridTotal
    End Function

    Private Function ExportBudget_CreateReferences_Reference(ByVal objTableGrids As List(Of Object), ByVal strParentSortNumber As String, ByVal strBudgetTitle() As String) As String
        Dim strReference As String = String.Empty

        For intTableIndex As Integer = 1 To objTableGrids.Count - 1
            Dim objTableGrid As Object = objTableGrids(intTableIndex)
            Dim intRowCount As Integer = objTableGrid.GetLength(0)
            Dim objValue As Object
            Dim strSortNumber As String
            Dim strPageName As String = strBudgetTitle(intTableIndex)

            For intReferenceRowIndex = 1 To intRowCount - 1
                objValue = objTableGrid(intReferenceRowIndex, 0)
                If objValue IsNot Nothing Then
                    strSortNumber = objValue.ToString

                    If strSortNumber = strParentSortNumber Then
                        If strReference = String.Empty Then
                            strReference = "="
                        Else
                            strReference &= "+"
                        End If

                        '='2009'!K2+'2010'!K2+'2011'!K2
                        strReference &= String.Format("'{0}'!K{1}", strPageName, intReferenceRowIndex + 1)
                    End If
                End If
            Next
        Next

        Return strReference
    End Function

    Private Sub ExportBudget_TableGrids_Finalise(ByVal objTableGrids As List(Of Object))
        For intTableIndex As Integer = 1 To objTableGrids.Count - 1
            Dim objTableGrid As Object = objTableGrids(intTableIndex)
            Dim intRowCount As Integer = objTableGrid.GetLength(0)
            Dim objValue As Object
            Dim strValue As String

            For intReferenceRowIndex = 1 To intRowCount - 1
                objValue = objTableGrid(intReferenceRowIndex, 6)
                strValue = objValue.ToString

                If strValue = "SUM" Then
                    objTableGrid(intReferenceRowIndex, 6) = Nothing
                End If
            Next
        Next
    End Sub

    Private Sub ExportBudget_LayOut(ByVal objWorkSheet As Object, ByVal intRowCount As Integer, ByVal intColumnCount As Integer)
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
            .Range(.Cells(intRow, 7), .Cells(intRow, 8)).Merge()
            .Range(.Cells(intRow, 9), .Cells(intRow, 10)).Merge()
            .Range(.Cells(intRow, 11), .Cells(intRow, 12)).Merge()
            With .Range(.Cells(intRow, 1), .Cells(intRow, intColumnCount)).Interior
                .ColorIndex = 15
                .Pattern = xlPattern_Solid
            End With
            .Rows("1:1").EntireRow.AutoFit()

            'bottom (totals) row
            With .Range(.Cells(intRowCount, 1), .Cells(intRowCount, intColumnCount))
                .Font.Bold = True
                .RowHeight = .RowHeight * 2
                .VerticalAlignment = xlVerticalAlignment_Center
            End With
            With .Range(.Cells(intRowCount, 3), .Cells(intRowCount, 10))
                .ClearContents()
            End With

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

            .Columns("C:L").EntireColumn.AutoFit()
        End With
    End Sub

    Private Sub ExportBudget_RemoveColumns(ByVal intBudgetYearsCount As Integer, ByVal boolMultiYearBudget As Boolean)
        Dim objWorkSheet As Object
        Dim boolTotalPage As Boolean

        For intWorkSheetIndex = 1 To intBudgetYearsCount
            objWorkSheet = WorkBook.Worksheets(intWorkSheetIndex)

            If boolMultiYearBudget = True And intWorkSheetIndex = 1 Then boolTotalPage = True Else boolTotalPage = False

            With objWorkSheet
                If boolTotalPage = True Then
                    .Columns("G:J").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
                    .Columns("C:D").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
                Else
                    If My.Settings.setPrintBudgetShowLocalCurrencyColumns = False Then
                        .Columns("I:J").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
                    End If
                    If My.Settings.setPrintBudgetShowDurationColumns = False Then
                        .Columns("C:D").Delete(Shift:=xlDeleteShiftDirection_ShiftToLeft)
                    End If
                End If
            End With
        Next
    End Sub
End Class
