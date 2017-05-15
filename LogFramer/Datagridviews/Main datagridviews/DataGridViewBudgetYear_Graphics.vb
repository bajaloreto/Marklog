Imports System.Drawing
Imports System.Drawing.Drawing2D

Partial Public Class DataGridViewBudgetYear

#Region "Custom painting"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim selGridRow As BudgetGridRow = Grid(intRowIndex)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds

        If intColumnIndex >= 0 Then
            If Me.Grid.Count > 0 Then
                If intRowIndex >= 0 And intRowIndex < IndexTotal Then
                    Select Case selGridRow.Type
                        Case BudgetItem.BudgetItemTypes.Normal, BudgetItem.BudgetItemTypes.Category, BudgetItem.BudgetItemTypes.ReferenceSource, BudgetItem.BudgetItemTypes.Reference
                            CellPainting_Normal(e)
                        Case BudgetItem.BudgetItemTypes.Ratio
                            CellPainting_Ratios(e)
                            e.Handled = True
                    End Select
                ElseIf intRowIndex = IndexTotal Then
                    CellPainting_Totals(intColumnIndex, selGridRow, rCell, CellGraphics)
                    e.Handled = True
                End If
            Else
                If intRowIndex >= 0 Then
                    CellPainting_Totals(intColumnIndex, selGridRow, rCell, CellGraphics)
                    e.Handled = True
                End If
            End If
        End If
    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As BudgetGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow.BudgetItemImage IsNot Nothing Then
            rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.BudgetItemImage.Width, selGridRow.BudgetItemImage.Height)
            CellGraphics.DrawImage(selGridRow.BudgetItemImage, rImage)
        End If
    End Sub

    Private Sub CellPainting_Normal(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim selGridRow As BudgetGridRow = Grid(intRowIndex)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds

        If selGridRow.IsSubTotal Then
            CellPainting_SubTotals(e.Value, intColumnIndex, selGridRow, rCell, CellGraphics)

            e.Paint(rCell, DataGridViewPaintParts.Border)
            e.Handled = True
        Else
            If intColumnIndex = 1 Then
                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectBudgetItems Then
                    CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
                Else
                    e.PaintBackground(rCell, False)
                End If

                If selGridRow IsNot Nothing Then _
                    CellPainting_RichText(selGridRow, rCell, CellGraphics)

                e.Paint(rCell, DataGridViewPaintParts.Border)
                e.Handled = True
            Else
                If Me.TotalBudget = True Then
                    CellPainting_TotalBudgetHeaders(e)

                    e.Handled = True
                Else
                    If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectBudgetItems Then
                        CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
                        e.PaintContent(rCell)
                        e.Paint(rCell, DataGridViewPaintParts.Border)
                        e.Handled = True
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub CellPainting_TotalBudgetHeaders(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim selGridRow As BudgetGridRow = Grid(e.RowIndex)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim strColumnName As String = Columns(e.ColumnIndex).Name
        Dim rText As Rectangle = GetTextRectangle(rCell)
        Dim objFormat As New StringFormat

        CellGraphics.FillRectangle(Brushes.White, rCell)

        Select Case strColumnName
            Case "Number", "TotalLocalCost", "TotalCost"
                objFormat.Alignment = StringAlignment.Far
            Case "SortNumber", "NumberUnit"
                objFormat.Alignment = StringAlignment.Near
        End Select

        CellGraphics.DrawString(e.Value, Me.Font, New SolidBrush(Color.Black), rText, objFormat)
        e.Paint(rCell, DataGridViewPaintParts.Border)
    End Sub

    Private Sub CellPainting_Ratios(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim selGridRow As BudgetGridRow = Grid(e.RowIndex)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim strColumnName As String = Columns(e.ColumnIndex).Name

        Select Case strColumnName
            Case "SortNumber", "TotalCost"
                e.PaintBackground(rCell, False)
                e.PaintContent(rCell)
            Case "RTF"
                e.PaintBackground(rCell, False)
                CellPainting_RichText(selGridRow, rCell, CellGraphics)
                e.Paint(rCell, DataGridViewPaintParts.Border)
            Case Else
                e.PaintBackground(rCell, False)
                CellGraphics.FillRectangle(Brushes.White, rCell)
                e.Paint(rCell, DataGridViewPaintParts.Border)
        End Select

    End Sub

    Private Sub CellPainting_SubTotals(ByVal objValue As Object, ByVal intColumnIndex As Integer, ByVal selGridRow As BudgetGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim BackGroundColor As Color = Color.LightGreen
        Dim ForeGroundColor As Color = Color.Black
        Dim intLevels As Integer = selGridRow.BudgetItem.GetLevelsOfChildren - 1
        Dim strColumnName As String = Columns(intColumnIndex).Name
        Dim rText As Rectangle = GetTextRectangle(rCell)

        If intLevels > 0 Then BackGroundColor = DarkerColor(BackGroundColor, intLevels)
        If intLevels > 5 Then ForeGroundColor = Color.White

        CellGraphics.FillRectangle(New SolidBrush(BackGroundColor), rCell)

        Select Case strColumnName
            Case "RTF"
                ReloadImages_BudgetItem(selGridRow, ForeGroundColor, BackGroundColor)

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectBudgetItems Then
                    CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
                End If

                CellPainting_RichText(selGridRow, rCell, CellGraphics)
            Case "Number", "TotalLocalCost", "TotalCost"
                Dim objFormat As New StringFormat
                objFormat.Alignment = StringAlignment.Far

                CellGraphics.DrawString(objValue, Me.Font, New SolidBrush(ForeGroundColor), rText, objFormat)
            Case "SortNumber", "NumberUnit"
                Dim objFormat As New StringFormat
                objFormat.Alignment = StringAlignment.Near

                CellGraphics.DrawString(objValue, Me.Font, New SolidBrush(ForeGroundColor), rText, objFormat)
        End Select
    End Sub

    Private Sub CellPainting_Totals(ByVal intColumnIndex As Integer, ByVal selGridRow As BudgetGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim strColumnName As String = Columns(intColumnIndex).Name
        Dim fntTotal As New Font(Me.Font, FontStyle.Bold)
        Dim rText As Rectangle = GetTextRectangle(rCell)

        CellGraphics.FillRectangle(Brushes.DarkGreen, rCell)

        Select Case strColumnName
            Case "RTF"
                CellGraphics.DrawString(LANG_TotalBudget, fntTotal, Brushes.White, rText)
            Case "TotalCost"
                Dim objFormat As New StringFormat
                Dim strValue As String = Me.BudgetYear.TotalCost.ToString
                objFormat.Alignment = StringAlignment.Far

                CellGraphics.DrawString(strValue, fntTotal, Brushes.White, rText, objFormat)
        End Select

    End Sub

    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim RowGraphics As Graphics = e.Graphics

        'draw focus rectangle
        DrawSelectionRectangle(RowGraphics)
    End Sub
#End Region 'Custom painting

#Region "Selection rectangle"
    Private Sub InvalidateSelectionRectangle()
        With SelectionRectangle
            Dim rSelection As New Rectangle(.Rectangle.X - 2, .Rectangle.Y - 2, .Rectangle.Width + 4, .Rectangle.Height + 4)
            Me.Invalidate(rSelection)
        End With
    End Sub

    Private Sub DrawSelectionRectangle(ByVal graph As Graphics)
        SetSelectionRectangle()

        With SelectionRectangle
            Dim IndexLastDisplayedRow As Integer = Me.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            If .FirstRowIndex > IndexLastDisplayedRow And .LastRowIndex > IndexLastDisplayedRow Then Exit Sub

            Dim IndexFirstDisplayedRow As Integer = Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
            If .FirstRowIndex < IndexFirstDisplayedRow And .LastRowIndex < IndexFirstDisplayedRow Then Exit Sub


            graph.DrawRectangle(penSelection, .Rectangle)
        End With
    End Sub

    Public Sub SetSelectionRectangle()
        Dim rTopLeftCell As New Rectangle
        Dim rBottomRightCell As New Rectangle
        Dim FirstRowIndexSelection As Integer

        If EditingControl IsNot Nothing Then
            With SelectionRectangle
                .FirstRowIndex = CurrentCell.RowIndex
                .LastRowIndex = CurrentCell.RowIndex
                .FirstColumnIndex = CurrentCell.ColumnIndex
                .LastColumnIndex = CurrentCell.ColumnIndex
            End With
        ElseIf SelectedCells.Count > 0 Then
            SetSelectionRectangleGridArea()
        Else
            Exit Sub
        End If

        With SelectionRectangle
            If .FirstRowIndex < 0 Or .LastRowIndex < 0 Then Exit Sub
            FirstRowIndexSelection = .FirstRowIndex

            rTopLeftCell = GetCellDisplayRectangle(.FirstColumnIndex, .FirstRowIndex, False)


            If Me.Rows(.LastRowIndex).Displayed = False And .LastRowIndex >= Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed) Then
                .LastRowIndex = Me.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            End If
            rBottomRightCell = GetCellDisplayRectangle(.LastColumnIndex, .LastRowIndex, True)

            Dim intLeft As Integer = GetColumnDisplayRectangle(.FirstColumnIndex, True).X
            Dim intTop As Integer = rTopLeftCell.Y
            Dim intFirstColumnWidth As Integer = GetColumnDisplayRectangle(.FirstColumnIndex, True).Width
            Dim intFirstRowHeight As Integer = rTopLeftCell.Height
            Dim intRight As Integer = GetColumnDisplayRectangle(.LastColumnIndex, True).X
            Dim intBottom As Integer = rBottomRightCell.Y
            Dim intLastColumnWidth As Integer = GetColumnDisplayRectangle(.LastColumnIndex, True).Width
            Dim intLastRowHeight As Integer = rBottomRightCell.Height

            If .FirstRowIndex = .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intFirstColumnWidth
                .Height = intFirstRowHeight
            ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intFirstColumnWidth
                .Height = intBottom + intLastRowHeight - intTop
            ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                .X = intRight
                .Y = intBottom
                .Width = intLastColumnWidth
                .Height = intTop + intFirstRowHeight - intBottom
            ElseIf .FirstRowIndex = .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intRight + intLastColumnWidth - intLeft
                .Height = intFirstRowHeight
            ElseIf .FirstRowIndex = .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                .X = intRight
                .Y = intBottom
                .Width = intLeft + intFirstColumnWidth - intRight
                .Height = intLastRowHeight

            ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intRight + intLastColumnWidth - intLeft
                .Height = intBottom + intLastRowHeight - intTop
            ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                .X = intLeft
                .Y = intBottom
                .Width = intRight + intLastColumnWidth - intLeft
                .Height = intTop + intFirstRowHeight - intBottom
            ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                .X = intRight
                .Y = intTop
                .Width = intLeft + intFirstColumnWidth - intRight
                .Height = intBottom + intLastRowHeight - intTop
            ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                .X = intRight
                .Y = intBottom
                .Width = intLeft + intFirstColumnWidth - intRight
                .Height = intTop + intFirstRowHeight - intBottom
            End If

            .X += 1
            .Y += 1
            .Width -= 3
            .Height -= 3

            If .Rectangle.Bottom = -1 And .LastRowIndex >= Me.FirstDisplayedScrollingRowIndex Then .Height = Me.Bottom - .Y - 1
        End With
    End Sub

    Public Sub SetSelectionRectangleGridArea()
        Dim Vdir As Integer
        Dim intSelSize As Integer = SelectedCells.Count - 1
        If intSelSize < 0 Then intSelSize = 0

        If SelectedCells(0).RowIndex >= SelectedCells(intSelSize).RowIndex Then Vdir = 1 Else Vdir = -1

        If CurrentCell IsNot Nothing Then
            SelectionRectangle.FirstRowIndex = CurrentCell.RowIndex
            SelectionRectangle.FirstColumnIndex = CurrentCell.ColumnIndex
            SelectionRectangle.LastRowIndex = CurrentCell.RowIndex
            SelectionRectangle.LastColumnIndex = CurrentCell.ColumnIndex
        End If

        If SelectedCells.Count > 1 Then
            For i = 0 To SelectedCells.Count - 1
                If SelectedCells(i).RowIndex < SelectionRectangle.FirstRowIndex Then SelectionRectangle.FirstRowIndex = SelectedCells(i).RowIndex
                If SelectedCells(i).RowIndex > SelectionRectangle.LastRowIndex Then SelectionRectangle.LastRowIndex = SelectedCells(i).RowIndex
                SelectionRectangle.FirstColumnIndex = 0
                SelectionRectangle.LastColumnIndex = Columns.Count - 1
            Next
        End If
    End Sub
#End Region
End Class
