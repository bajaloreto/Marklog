Imports Word = Microsoft.Office.Interop.Word

Partial Public Class WordIO
    Private objTargetGroupIdFormRows As New PrintTargetGroupIdFormRows

    Public Sub ExportTargetGroupIdFormToNewDocument(ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        ExportTargetGroupIdForm(objDoc, objRange)
    End Sub

    Public Sub ExportTargetGroupIdFormToDocument(ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportTargetGroupIdForm(objDoc, objRange)
    End Sub

    Public Sub ExportTargetGroupIdFormToBookmark(ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportTargetGroupIdForm(objDoc, objRange)
    End Sub

    Private Sub ExportTargetGroupIdForm(ByVal objDoc As Object, ByVal objRange As Object)
        Dim intRowIndex As Integer
        Dim objTargetGroups As TargetGroups = CurrentLogFrame.GetTargetGroups()
        Dim selGridRow As ExportTargetGroupIdRow

        If objTargetGroups Is Nothing Then Exit Sub

        'Set styles
        SetTableStyle_StripedGreen(objDoc.Application, objDoc)

        Using objExportTargetGroupIdForm As New ExportTargetGroupIdForm(CurrentLogFrame.GetTargetGroups())
            objExportTargetGroupIdForm.LoadTable()

            For intRowIndex = 0 To objExportTargetGroupIdForm.TableGrid.Count - 1
                selGridRow = objExportTargetGroupIdForm.TableGrid(intRowIndex)

                If selGridRow.IsTitle Then
                    Dim objTableGrid As Object = ExportTargetGroupIdForm_CreateTableGrid(objExportTargetGroupIdForm, intRowIndex)
                    Dim objTable As Object = Nothing

                    If objTableGrid IsNot Nothing Then objTable = CreateTable(objDoc, objRange, objTableGrid)

                    If objTable IsNot Nothing Then
                        With objTable
                            .Cell(1, 1).Merge(.Cell(1, 2))
                            .Style = "Logframer table striped"
                        End With

                        objRange.Collapse(wdCollapseEnd)
                        objRange.InsertParagraphAfter()

                        objRange.Select()
                        With objDoc.Application.Selection
                            .TypeParagraph()
                            .Style = wdStyleNormal
                            objRange = .Range
                        End With
                    End If
                End If
            Next
        End Using
    End Sub

    Private Function ExportTargetGroupIdForm_CreateTableGrid(ByVal objExportTargetGroupIdForm As ExportTargetGroupIdForm, ByVal intStartIndex As Integer) As Object
        Dim intColumnCount As Integer = objExportTargetGroupIdForm.GetColumnCount
        Dim intRowCount As Integer = objExportTargetGroupIdForm.GetRowCountOfTargetGroup(intStartIndex)
        Dim objTableGrid(intRowCount, intColumnCount - 1) As Object
        Dim selGridRow As ExportTargetGroupIdRow
        Dim intRowIndex As Integer

        If intRowCount = 0 Then Return Nothing

        For i = intStartIndex To objExportTargetGroupIdForm.TableGrid.Count - 1
            selGridRow = objExportTargetGroupIdForm.TableGrid(i)

            objTableGrid(intRowIndex, 0) = selGridRow.PropertyName
            objTableGrid(intRowIndex, 1) = selGridRow.PropertyValue

            intRowIndex += 1
            If i >= objExportTargetGroupIdForm.TableGrid.Count - 1 OrElse objExportTargetGroupIdForm.TableGrid(i + 1).IsTitle Then Exit For

        Next

        Return objTableGrid
    End Function
End Class
