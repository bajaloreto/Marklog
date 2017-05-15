Partial Public Class WordIO
    Public Sub ExportPartnerListToNewDocument(ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        ExportPartnerListToWord(objDoc, objRange)
    End Sub

    Public Sub ExportPartnerListToDocument(ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportPartnerListToWord(objDoc, objRange)
    End Sub

    Public Sub ExportPartnerListToBookmark(ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        ExportPartnerListToWord(objDoc, objRange)
    End Sub

    Private Sub ExportPartnerListToWord(ByVal objDoc As Object, ByVal objRange As Object)
        Dim intRowIndex As Integer
        Dim selGridRow As ExportPartnerListRow

        'Set styles
        SetTableStyle_StripedGreen(objDoc.Application, objDoc)

        Using objExportPartnerList As New ExportPartnerList(CurrentLogFrame)
            objExportPartnerList.LoadTable()

            For intRowIndex = 0 To objExportPartnerList.TableGrid.Count - 1
                selGridRow = objExportPartnerList.TableGrid(intRowIndex)

                If selGridRow.PropertyName.StartsWith(LANG_Organisation) Then
                    Dim objTableGrid As Object = ExportPartnerList_CreateTableGrid(objExportPartnerList, intRowIndex)
                    Dim objTable As Object = Nothing

                    If objTableGrid IsNot Nothing Then objTable = CreateTable(objDoc, objRange, objTableGrid)

                    If objTable IsNot Nothing Then
                        With objTable
                            .Style = "Logframer table striped"

                            .AutoFitBehavior(wdAutofitBehaviour_FitContent)
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

    Private Function ExportPartnerList_CreateTableGrid(ByVal objExportPartnerList As ExportPartnerList, ByVal intStartIndex As Integer) As Object
        Dim intColumnCount As Integer = objExportPartnerList.GetColumnCount
        Dim intRowCount As Integer = objExportPartnerList.GetRowCountOfOrganisation(intStartIndex)
        Dim objTableGrid(intRowCount, intColumnCount - 1) As Object
        Dim selGridRow As ExportPartnerListRow
        Dim intRowIndex As Integer

        If intRowCount = 0 Then Return Nothing

        For i = intStartIndex To objExportPartnerList.TableGrid.Count - 1
            selGridRow = objExportPartnerList.TableGrid(i)

            objTableGrid(intRowIndex, 0) = selGridRow.PropertyName
            objTableGrid(intRowIndex, 1) = selGridRow.PropertyValue

            intRowIndex += 1
            If i >= objExportPartnerList.TableGrid.Count - 1 OrElse objExportPartnerList.TableGrid(i + 1).PropertyName.StartsWith(LANG_Organisation) Then Exit For

        Next

        Return objTableGrid
    End Function
End Class
