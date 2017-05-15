Partial Public Class WordIO
    Public Sub ExportRisksTableToNewDocument(ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        ExportRisksTable(objDoc, objRange)
    End Sub

    Public Sub ExportRisksTableToDocument(ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportRisksTable(objDoc, objRange)
    End Sub

    Public Sub ExportRisksTableToBookmark(ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportRisksTable(objDoc, objRange)
    End Sub

    Private Sub ExportRisksTable(ByVal objDoc As Object, ByVal objRange As Object)
        'Set styles
        SetTableStyle_Basic(objDoc.Application, objDoc)

        Using objExportRiskTable As New ExportRisksTable(CurrentLogFrame, My.Settings.setPrintRiskRegisterRiskCategory)
            objExportRiskTable.LoadTable()

            Dim objTableGrid As Object = ExportRiskTable_CreateTableGrid(objExportRiskTable)
            Dim objTable As Object = Nothing

            If objTableGrid IsNot Nothing Then objTable = CreateTable(objDoc, objRange, objTableGrid)

            If objTable IsNot Nothing Then
                With objTable
                    .Rows(1).HeadingFormat = True
                    .Cell(1, 4).Merge(.Cell(1, 5))
                    .Cell(1, 1).Merge(.Cell(1, 2))
                    .Style = "Logframer table basic"

                    .AutoFitBehavior(wdAutofitBehaviour_FitContent)
                End With

                objRange.Collapse(wdCollapseEnd)
                objRange.InsertParagraphAfter()
            End If
        End Using
    End Sub

    Private Function ExportRiskTable_CreateTableGrid(ByVal objExportRiskTable As ExportRisksTable) As Object
        Dim intColumnCount As Integer = objExportRiskTable.GetColumnCount
        Dim intRowCount As Integer = objExportRiskTable.GetRowCount + 1
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object
        Dim intRow As Integer = 1

        If intRowCount = 0 Then Return Nothing

        'Title row
        objTableGrid(0, 1) = LANG_RiskDefinition
        objTableGrid(0, 2) = LANG_RiskResponse
        objTableGrid(0, 4) = LANG_Objective
        objTableGrid(0, 5) = LANG_Likelihood
        objTableGrid(0, 6) = LANG_Impact
        objTableGrid(0, 7) = LANG_RiskLevel

        'Values
        For Each selGridRow As ExportRiskTableRow In objExportRiskTable.TableGrid
            If String.IsNullOrEmpty(selGridRow.SectionTitle) = False Then
                objTableGrid(intRow, 1) = selGridRow.SectionTitle
            Else
                objTableGrid(intRow, 0) = selGridRow.AssumptionSortNumber
                objTableGrid(intRow, 1) = selGridRow.AssumptionText
                objTableGrid(intRow, 2) = selGridRow.RiskResponseText
                objTableGrid(intRow, 3) = selGridRow.StructSortNumber
                objTableGrid(intRow, 4) = selGridRow.StructText
                objTableGrid(intRow, 5) = selGridRow.LikelihoodText
                objTableGrid(intRow, 6) = selGridRow.ImpactText
                objTableGrid(intRow, 7) = selGridRow.RiskLevelText
            End If

            intRow += 1
        Next

        Return objTableGrid
    End Function
End Class
