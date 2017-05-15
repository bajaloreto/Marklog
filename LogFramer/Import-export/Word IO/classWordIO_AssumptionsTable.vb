Partial Public Class WordIO
    Public Sub ExportAssumptionsTableToNewDocument(ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        ExportAssumptionsTable(objDoc, objRange)
    End Sub

    Public Sub ExportAssumptionsTableToDocument(ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportAssumptionsTable(objDoc, objRange)
    End Sub

    Public Sub ExportAssumptionsTableToBookmark(ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportAssumptionsTable(objDoc, objRange)
    End Sub

    Private Sub ExportAssumptionsTable(ByVal objDoc As Object, ByVal objRange As Object)
        'Set styles
        SetTableStyle_Basic(objDoc.Application, objDoc)

        Using objExportAssumptionsTable As New ExportAssumptionsTable(CurrentLogFrame, My.Settings.setPrintAssumptionsSection)
            objExportAssumptionsTable.LoadTable()

            Dim objTableGrid As Object = ExportAssumptionsTable_CreateTableGrid(objExportAssumptionsTable)
            Dim objTable As Object = Nothing

            If objTableGrid IsNot Nothing Then objTable = CreateTable(objDoc, objRange, objTableGrid)

            If objTable IsNot Nothing Then
                With objTable
                    .Rows(1).HeadingFormat = True
                    .Cell(1, 1).Merge(.Cell(1, 2))
                    .Style = "Logframer table basic"

                    .AutoFitBehavior(wdAutofitBehaviour_FitContent)
                End With

                objRange.Collapse(wdCollapseEnd)
                objRange.InsertParagraphAfter()
            End If
        End Using
    End Sub

    Private Function ExportAssumptionsTable_CreateTableGrid(ByVal objExportAssumptionsTable As ExportAssumptionsTable) As Object
        Dim intColumnCount As Integer = objExportAssumptionsTable.GetColumnCount
        Dim intRowCount As Integer = objExportAssumptionsTable.GetRowCount + 1
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object
        Dim intRow As Integer = 1

        If intRowCount = 0 Then Return Nothing

        'Title row
        objTableGrid(0, 1) = LANG_Assumptions
        objTableGrid(0, 2) = LANG_Reason
        objTableGrid(0, 3) = LANG_HowToValidate
        objTableGrid(0, 4) = LANG_Validated
        objTableGrid(0, 5) = LANG_Impact
        objTableGrid(0, 6) = LANG_ResponseStrategy
        objTableGrid(0, 7) = LANG_Owner

        'Values
        For Each selGridRow As ExportAssumptionTableRow In objExportAssumptionsTable.TableGrid
            If String.IsNullOrEmpty(selGridRow.StructText) = False Then
                objTableGrid(intRow, 0) = selGridRow.StructSortNumber
                objTableGrid(intRow, 1) = selGridRow.StructText
            Else
                objTableGrid(intRow, 0) = selGridRow.AssumptionSortNumber
                objTableGrid(intRow, 1) = selGridRow.AssumptionText
                objTableGrid(intRow, 2) = selGridRow.Reason
                objTableGrid(intRow, 3) = selGridRow.HowToValidate
                objTableGrid(intRow, 4) = selGridRow.Validated
                objTableGrid(intRow, 5) = selGridRow.Impact
                objTableGrid(intRow, 6) = selGridRow.ResponseStrategy
                objTableGrid(intRow, 7) = selGridRow.Owner
            End If

            intRow += 1
        Next

        Return objTableGrid
    End Function
End Class
