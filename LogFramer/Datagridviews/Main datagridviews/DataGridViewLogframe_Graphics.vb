Imports System.Drawing
Imports System.Drawing.Drawing2D

Partial Public Class DataGridViewLogframe

#Region "Custom painting"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds

        If boolColumnWidthChanged = True Then
            Reload()
            boolColumnWidthChanged = False
        End If

        If e.RowIndex = -1 Then
            OnCellPainting_Headers(CellGraphics, rCell, e.ColumnIndex)
            e.Handled = True
        ElseIf e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
            Dim strColumnName As String = Columns(e.ColumnIndex).Name
            Dim selGridRow As LogframeRow = Me.Grid(e.RowIndex)
            Dim strValue As String = CType(e.Value, String)

            If selGridRow.RowType = LogframeRow.RowTypes.Section Then
                Dim brSection As New LinearGradientBrush(rCell, CurrentLogFrame.SectionColorTop, CurrentLogFrame.SectionColorBottom, LinearGradientMode.Vertical)
                CellGraphics.FillRectangle(brSection, rCell)
                If strColumnName = "StructSort" Then CellGraphics.DrawLine(penBlack1, rCell.Left, rCell.Top, rCell.Left, rCell.Bottom)
                If strColumnName.Contains("RTF") Then CellGraphics.DrawLine(penBlack2, rCell.Right - 1, rCell.Top, rCell.Right - 1, rCell.Bottom)
                CellGraphics.DrawLine(penBlack2, rCell.Left, rCell.Top + 1, rCell.Right, rCell.Top + 1)
            End If
            If selGridRow IsNot Nothing AndAlso selGridRow.RowType = LogframeRow.RowTypes.Normal Then
                If strColumnName.Contains("Ver") And IsResourceBudgetRow(selGridRow) And selGridRow.Resource Is Nothing Then strValue = String.Empty

                'Paint background
                OnCellPainting_Background(selGridRow, strColumnName, rCell, CellGraphics)

                'Draw text
                OnCellPainting_Text(strValue, selGridRow, strColumnName, rCell, CellGraphics)

                'Draw lines
                OnCellPainting_Lines(selGridRow, strColumnName, rCell, CellGraphics)
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub OnCellPainting_Headers(ByVal CellGraphics As Graphics, ByVal rCell As Rectangle, ByVal intColumnIndex As Integer)
        Dim penDark As New Pen(SystemBrushes.ControlDark, 1)
        Dim penGrid As Pen = New Pen(Me.GridColor, 1)

        CellGraphics.FillRectangle(SystemBrushes.ControlLight, rCell)

        CellGraphics.DrawLine(penGrid, rCell.Left, rCell.Bottom - 1, rCell.Right, rCell.Bottom - 1)
        If intColumnIndex Mod 2 <> 0 Then
            CellGraphics.DrawLine(penGrid, rCell.Right - 1, rCell.Top, rCell.Right - 1, rCell.Bottom)
        End If
    End Sub

    Private Sub OnCellPainting_Background(ByVal selGridRow As LogframeRow, ByVal strColumnName As String, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim colBackGround As Color = Color.White
        Dim colGrey As Color = System.Drawing.Color.FromArgb(230, 230, 230)
        Dim boolHighLight As Boolean = HighLight(strColumnName, selGridRow)

        Select Case strColumnName
            Case "StructSort"

            Case "StructRTF"

            Case "IndSort", "IndRTF"
                If selGridRow.StructImage Is Nothing Then colBackGround = colGrey
            Case "VerSort", "VerRTF"
                If IsResourceBudgetRow(selGridRow) = False Then
                    If selGridRow.IndicatorImage Is Nothing Then
                        colBackGround = colGrey
                    End If
                Else
                    If selGridRow.ResourceImage Is Nothing Then colBackGround = colGrey
                End If

            Case "AsmSort", "AsmRTF"
                If selGridRow.StructImage Is Nothing Then colBackGround = colGrey
        End Select

        Dim brNormalCell As New SolidBrush(colBackGround)

        If boolHighLight = True Then
            CellGraphics.FillRectangle(SystemBrushes.Highlight, rCell)
        Else
            CellGraphics.FillRectangle(brNormalCell, rCell)
        End If
    End Sub

    Private Sub OnCellPainting_Text(ByVal strValue As String, ByVal selGridRow As LogframeRow, ByVal strColumnName As String, ByVal rCell As Rectangle, _
                                    ByVal CellGraphics As Graphics)
        Dim boolHighLight As Boolean = HighLight(strColumnName, selGridRow)
        Dim rImage As Rectangle
        Select Case strColumnName
            Case "StructSort", "IndSort", "VerSort", "AsmSort"
                If String.IsNullOrEmpty(strValue) = False Then
                    Dim formatCells As New StringFormat()
                    Dim brText As SolidBrush

                    If boolHighLight = False Then
                        brText = New SolidBrush(Color.Black)
                    Else
                        brText = New SolidBrush(SystemColors.HighlightText)
                    End If

                    rCell.X += 2
                    rCell.Y += 2
                    formatCells.Alignment = StringAlignment.Near
                    formatCells.LineAlignment = StringAlignment.Near
                    CellGraphics.DrawString(strValue, CurrentLogFrame.SortNumberFont, brText, rCell, formatCells)
                End If
            Case "StructRTF"
                If selGridRow.StructImage IsNot Nothing Then
                    'CellGraphics.DrawRectangle(Pens.Bisque, rCell)
                    rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StructImage.Width, selGridRow.StructImage.Height)
                    CellGraphics.DrawImage(selGridRow.StructImage, rImage)
                End If
            Case "IndRTF"
                If IsResourceBudgetRow(selGridRow) = False Then
                    If selGridRow.IndicatorImage IsNot Nothing Then
                        rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.IndicatorImage.Width, selGridRow.IndicatorImage.Height)
                        CellGraphics.DrawImage(selGridRow.IndicatorImage, rImage)
                    End If
                Else
                    If selGridRow.ResourceImage IsNot Nothing Then
                        rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.ResourceImage.Width, selGridRow.ResourceImage.Height)
                        CellGraphics.DrawImage(selGridRow.ResourceImage, rImage)
                    End If
                End If
            Case "VerRTF"
                If IsResourceBudgetRow(selGridRow) = False Then
                    If selGridRow.VerificationSourceImage IsNot Nothing Then
                        rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.VerificationSourceImage.Width, selGridRow.VerificationSourceImage.Height)
                        CellGraphics.DrawImage(selGridRow.VerificationSourceImage, rImage)
                    End If
                Else
                    OnCellPainting_Text_ResourceBudget(selGridRow, rCell, CellGraphics, boolHighLight)
                End If
            Case "AsmRTF"
                If selGridRow.AssumptionImage IsNot Nothing Then
                    rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.AssumptionImage.Width, selGridRow.AssumptionImage.Height)
                    CellGraphics.DrawImage(selGridRow.AssumptionImage, rImage)
                End If
        End Select
    End Sub

    Private Sub OnCellPainting_Text_ResourceBudget(ByVal selGridRow As LogframeRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics, ByVal boolHighLight As Boolean)

        Dim selResource As Resource = selGridRow.Resource

        If selGridRow.Resource IsNot Nothing Then
            Dim formatCells As New StringFormat()
            Dim brText As New SolidBrush(selResource.GetFontColor)
            Dim selFont As Font = selResource.GetFont
            Dim strBudget As String = String.Format("{0} {1}", selResource.TotalCostAmount.ToString("N2"), CurrentLogFrame.CurrencyCode)

            formatCells = StringFormat.GenericTypographic
            formatCells.Alignment = StringAlignment.Far

            Select Case Me.TextSelectionIndex
                Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe, TextSelectionValues.SelectResourceBudgets
                    brText = New SolidBrush(SystemColors.HighlightText)
            End Select

            rCell.Y += 2
            rCell.Width -= 4
            CellGraphics.DrawString(strBudget, selFont, brText, rCell, formatCells)

        End If
    End Sub

    Private Function HighLight(ByVal strColumnName As String, ByVal selGridRow As LogframeRow) As Boolean
        Dim boolStructColumn, boolIndicatorColumn, boolVerificationSourceColumn, boolAssumptionColumn As Boolean
        Dim boolHighLight As Boolean

        If strColumnName.Contains("Struct") Then boolStructColumn = True
        If strColumnName.Contains("Ind") Then boolIndicatorColumn = True
        If strColumnName.Contains("Ver") Then boolVerificationSourceColumn = True
        If strColumnName.Contains("Asm") Then boolAssumptionColumn = True

        Select Case Me.TextSelectionIndex
            Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe
                boolHighLight = True
            Case Else
                If boolStructColumn = True And Me.TextSelectionIndex = TextSelectionValues.SelectStructs Then boolHighLight = True

                If boolIndicatorColumn = True Then
                    If IsResourceBudgetRow(selGridRow) = False And Me.TextSelectionIndex = TextSelectionValues.SelectIndicators Then
                        boolHighLight = True
                    ElseIf IsResourceBudgetRow(selGridRow) = True And Me.TextSelectionIndex = TextSelectionValues.SelectResources Then
                        boolHighLight = True
                    End If
                End If

                If boolVerificationSourceColumn = True Then
                    If IsResourceBudgetRow(selGridRow) = False And Me.TextSelectionIndex = TextSelectionValues.SelectVerificationSources Then
                        boolHighLight = True
                    ElseIf IsResourceBudgetRow(selGridRow) = True And Me.TextSelectionIndex = TextSelectionValues.SelectResourceBudgets Then
                        boolHighLight = True
                    End If
                End If

                If boolAssumptionColumn = True And Me.TextSelectionIndex = TextSelectionValues.SelectAssumptions Then boolHighLight = True
        End Select

        Return boolHighLight
    End Function

    Private Sub OnCellPainting_Lines(ByVal selGridRow As LogframeRow, ByVal strColumnName As String, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim boolMeansBudget As Boolean = IsResourceBudgetRow(selGridRow)

        Dim intCorrect As Integer
        Dim boolDraw As Boolean

        If Me.HideEmptyCells = True Then intCorrect = 1

        'thick vertical line to separate columns
        If strColumnName = "StructSort" Then CellGraphics.DrawLine(penBlack1, rCell.Left, rCell.Top, rCell.Left, rCell.Bottom)
        If strColumnName.Contains("RTF") Then CellGraphics.DrawLine(penBlack2, rCell.Right - 1, rCell.Top, rCell.Right - 1, rCell.Bottom)

        'horizontal lines on top of cells
        If strColumnName.Contains("Struct") Then
            boolDraw = DrawLineAboveStruct(selGridRow)
        ElseIf strColumnName.Contains("Ind") Then
            boolDraw = DrawLineAboveIndicator(selGridRow)
        ElseIf strColumnName.Contains("Ver") Then
            boolDraw = DrawLineAboveVerificationSource(selGridRow)
        ElseIf strColumnName.Contains("Asm") Then
            boolDraw = DrawLineAboveAssumption(selGridRow)
        End If

        If boolDraw = True Then CellGraphics.DrawLine(penBlack1, rCell.Left, rCell.Top, rCell.Right, rCell.Top)
    End Sub

    Private Function DrawLineAboveStruct(ByVal selGridRow As LogframeRow) As Boolean
        Dim boolDraw As Boolean

        If String.IsNullOrEmpty(selGridRow.StructRtf) = False Then
            boolDraw = True
        Else
            If Me.HideEmptyCells = False And selGridRow.StructImage Is Nothing Then boolDraw = True
        End If

        Return boolDraw
    End Function

    Private Function DrawLineAboveIndicator(ByVal selGridRow As LogframeRow) As Boolean
        Dim boolDraw As Boolean
        Dim PreviousRow As LogframeRow = Me.Grid(Grid.IndexOf(selGridRow) - 1)

        If IsResourceBudgetRow(selGridRow) = False Then
            If String.IsNullOrEmpty(selGridRow.StructRtf) = False Or _
                String.IsNullOrEmpty(selGridRow.IndicatorRtf) = False Then
                boolDraw = True
            ElseIf Me.HideEmptyCells = False Then
                If selGridRow.IndicatorImage Is Nothing And PreviousRow.IndicatorImage IsNot Nothing Then boolDraw = True
                If selGridRow.StructImage Is Nothing And selGridRow.IndicatorImage Is Nothing Then boolDraw = True
            End If
        Else
            If String.IsNullOrEmpty(selGridRow.StructRtf) = False Or _
                String.IsNullOrEmpty(selGridRow.ResourceRtf) = False Then
                boolDraw = True
            ElseIf Me.HideEmptyCells = False Then
                'If selGridRow.Struct Is Nothing Then boolDraw = True
                If selGridRow.ResourceImage Is Nothing And PreviousRow.ResourceImage IsNot Nothing Then boolDraw = True
                If selGridRow.StructImage Is Nothing And selGridRow.ResourceImage Is Nothing Then boolDraw = True
            End If
        End If

        Return boolDraw
    End Function

    Private Function DrawLineAboveVerificationSource(ByVal selGridRow As LogframeRow) As Boolean
        Dim boolDraw As Boolean
        Dim PreviousRow As LogframeRow = Me.Grid(Grid.IndexOf(selGridRow) - 1)

        If IsResourceBudgetRow(selGridRow) = False Then
            If String.IsNullOrEmpty(selGridRow.StructRtf) = False Or _
                String.IsNullOrEmpty(selGridRow.IndicatorRtf) = False Or _
                String.IsNullOrEmpty(selGridRow.VerificationSourceRtf) = False Then
                boolDraw = True
            Else
                If selGridRow.VerificationSourceImage Is Nothing And Me.HideEmptyCells = False Then
                    If selGridRow.Struct Is Nothing And selGridRow.Indicator Is Nothing And selGridRow.VerificationSource Is Nothing And selGridRow.Assumption Is Nothing Then boolDraw = True
                    If selGridRow.StructImage Is Nothing And selGridRow.IndicatorImage Is Nothing Then boolDraw = True
                    If selGridRow.IndicatorImage IsNot Nothing Then boolDraw = True
                    If PreviousRow.VerificationSourceImage IsNot Nothing Then boolDraw = True
                End If
            End If
        Else
            If String.IsNullOrEmpty(selGridRow.StructRtf) = False Or _
                String.IsNullOrEmpty(selGridRow.ResourceRtf) = False Then
                boolDraw = True
            ElseIf Me.HideEmptyCells = False Then
                'If selGridRow.StructImage Is Nothing And selGridRow.ResourceImage Is Nothing And PreviousRow.ResourceImage IsNot Nothing Then boolDraw = True
                If selGridRow.ResourceImage Is Nothing And PreviousRow.ResourceImage IsNot Nothing Then boolDraw = True
                If selGridRow.StructImage Is Nothing And selGridRow.ResourceImage Is Nothing Then boolDraw = True
            End If
        End If

        Return boolDraw
    End Function

    Private Function DrawLineAboveAssumption(ByVal selGridRow As LogframeRow) As Boolean
        Dim boolDraw As Boolean
        Dim PreviousRow As LogframeRow = Me.Grid(Grid.IndexOf(selGridRow) - 1)

        If String.IsNullOrEmpty(selGridRow.StructRtf) = False Or String.IsNullOrEmpty(selGridRow.AssumptionRtf) = False Then
            boolDraw = True
        ElseIf Me.HideEmptyCells = False Then
            If selGridRow.AssumptionImage Is Nothing And PreviousRow.AssumptionImage IsNot Nothing Then boolDraw = True
            If selGridRow.StructImage Is Nothing And selGridRow.AssumptionImage Is Nothing Then boolDraw = True
        End If

        Return boolDraw
    End Function
#End Region

#Region "Custom painting - Sections and repeated rows"
    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim selGridRow As LogframeRow = Grid(e.RowIndex)
        Dim RowGraphics As Graphics = e.Graphics

        If selGridRow IsNot Nothing Then

            Select Case selGridRow.RowType
                Case GridRow.RowTypes.Section
                    'section titles
                    RowPostPaint_PaintSections(RowGraphics, e.RowBounds, selGridRow, e.RowIndex)
                Case GridRow.RowTypes.RepeatPurpose, GridRow.RowTypes.RepeatOutput
                    'repeated purposes and outputs
                    RowPostPaint_PaintRepeatedRows(RowGraphics, e.RowBounds, selGridRow, e.RowIndex)
            End Select

            'draw focus rectangle
            DrawSelectionRectangle(RowGraphics)
        End If
    End Sub

    Private Sub RowPostPaint_PaintSections(ByVal RowGraphics As Graphics, ByVal rRowBounds As Rectangle, ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer)
        'draw section titles spanning two columns
        Dim objStringFormat As New StringFormat()
        Dim brFont As Brush = New SolidBrush(Logframe.SectionTitleFontColor)
        Dim strTitle As String
        Dim rStruct, rIndicator, rVerificationSource, rAssumption As Rectangle
        Dim xStruct, yStruct As Integer
        Dim xInd, yInd As Integer
        Dim xVer, yVer As Integer
        Dim xAsm, yAsm As Integer
        Dim intRowHeadersWidth As Integer

        objStringFormat = StringFormat.GenericDefault
        objStringFormat.Alignment = StringAlignment.Center
        objStringFormat.LineAlignment = StringAlignment.Center

        'draw header background with gradient
        If Me.RowHeadersVisible = True Then intRowHeadersWidth = Me.RowHeadersWidth
        Dim rRow As New Rectangle(intRowHeadersWidth + 1, rRowBounds.Top, _
                                  Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - _
                                  HorizontalScrollingOffset + 1, rRowBounds.Height)

        Dim rSection As New Rectangle(rRow.X, rRow.Y, rRow.Width, rRow.Height)
        Dim brSection As LinearGradientBrush = New LinearGradientBrush(rSection, CurrentLogFrame.SectionColorTop, CurrentLogFrame.SectionColorBottom, LinearGradientMode.Vertical)

        RowGraphics.FillRectangle(brSection, rSection)
        RowGraphics.DrawLine(penBlack2, rSection.Left, rSection.Top + 1, rSection.Right, rSection.Top + 1)

        'draw struct section title (goal, purpose, output, activity)
        xStruct = rRow.X
        yStruct = rRow.Y
        rStruct = New Rectangle(xStruct, yStruct, StructSortColumn.Width + StructRtfColumn.Width, rRow.Height)

        RowGraphics.DrawString(selGridRow.Struct.Text, Logframe.SectionTitleFont, brFont, rStruct, objStringFormat)
        RowGraphics.DrawLine(penBlack2, rStruct.Left, rStruct.Top, rStruct.Left, rStruct.Bottom)
        RowGraphics.DrawLine(penBlack2, rStruct.Right, rStruct.Top, rStruct.Right, rStruct.Bottom)

        xInd = xStruct + rStruct.Width
        yInd = yStruct

        'draw Indicator/Resource section title on canvas
        If ShowIndicatorColumn = True Then
            If IsResourceBudgetRow(selGridRow) = False Then strTitle = selGridRow.Indicator.Text Else strTitle = selGridRow.Resource.Text
            rIndicator = New Rectangle(xInd, yInd, IndicatorSortColumn.Width + IndicatorRtfColumn.Width, rRow.Height)

            RowGraphics.DrawString(strTitle, Logframe.SectionTitleFont, brFont, rIndicator, objStringFormat)
            RowGraphics.DrawLine(penBlack2, rIndicator.Right, rIndicator.Top, rIndicator.Right, rIndicator.Bottom)
            xVer = xInd + rIndicator.Width
            yVer = yInd
        Else
            xVer = xInd
            yVer = yInd
        End If


        If ShowVerificationSourceColumn = True Then
            If IsResourceBudgetRow(selGridRow) = False Then strTitle = selGridRow.VerificationSource.Text Else strTitle = Resource.BudgetName
            rVerificationSource = New Rectangle(xVer, yVer, VerificationSourceSortColumn.Width + VerificationSourceRtfColumn.Width, rRow.Height)

            RowGraphics.DrawString(strTitle, Logframe.SectionTitleFont, brFont, rVerificationSource, objStringFormat)
            RowGraphics.DrawLine(penBlack2, rVerificationSource.Right, rVerificationSource.Top, rVerificationSource.Right, rVerificationSource.Bottom)
            xAsm = xVer + rVerificationSource.Width
            yAsm = yVer
        Else
            xAsm = xVer
            yAsm = yVer
        End If
        If ShowAssumptionColumn = True Then
            rAssumption = New Rectangle(xAsm, yAsm, AssumptionSortColumn.Width + AssumptionRtfColumn.Width, rRow.Height)

            RowGraphics.DrawString(selGridRow.Assumption.Text, Logframe.SectionTitleFont, brFont, rAssumption, objStringFormat)
            RowGraphics.DrawLine(penBlack2, rAssumption.Right, rAssumption.Top, rAssumption.Right, rAssumption.Bottom)
        End If

        RowGraphics.DrawLine(penBlack1, rSection.Left, rSection.Top, rSection.Right, rSection.Top)
    End Sub

    Private Sub RowPostPaint_PaintRepeatedRows(ByVal RowGraphics As Graphics, ByVal rRowBounds As Rectangle, ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer)
        'draw repeated Purposes/Outputs
        Dim rStructSort As Rectangle, rStructRTF As Rectangle
        Dim brBackGround As Brush
        Dim intX = 1, intY As Integer
        Dim fntRepeat As Font = New Font(Me.Logframe.SortNumberFont, FontStyle.Bold)
        Dim brRepeat As SolidBrush = New SolidBrush(Color.Black)
        Dim objStringFormat As New StringFormat()

        objStringFormat.Alignment = StringAlignment.Near
        objStringFormat.LineAlignment = StringAlignment.Near

        If Me.RowHeadersVisible = True Then intX = Me.RowHeadersWidth
        intY = rRowBounds.Top

        Dim rRow As New Rectangle(intX, intY, Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - _
                                  HorizontalScrollingOffset, rRowBounds.Height)

        If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose Then brBackGround = Brushes.CornflowerBlue Else brBackGround = Brushes.LightSteelBlue
        RowGraphics.FillRectangle(brBackGround, rRow)

        'draw number
        rStructSort = GetCellDisplayRectangle(StructSortColumn.Index, intRowIndex, False)
        rStructSort.X += 2

        RowGraphics.DrawString(selGridRow.StructSort, fntRepeat, brRepeat, rStructSort, objStringFormat)

        'draw text
        rStructRTF = GetCellDisplayRectangle(StructRtfColumn.Index, intRowIndex, False)
        Dim strText As String = selGridRow.Struct.Text
        Dim rRepeat As New Rectangle(rStructRTF.X, rStructRTF.Y, _
                                     Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - rStructSort.Width - HorizontalScrollingOffset + 1, rRowBounds.Height)
        Dim intHeight As Integer = RowGraphics.MeasureString(strText, Logframe.SectionTitleFont, rRepeat.Width, objStringFormat).Height

        rRepeat.Height = intHeight
        If intHeight = 0 Then intHeight = NewCellHeight()
        Me.Rows(intRowIndex).Height = intHeight

        RowGraphics.DrawString(strText, Logframe.SectionTitleFont, brRepeat, rRepeat, objStringFormat)
        RowGraphics.DrawLine(penBlack2, rRow.Left, rRow.Top, rRow.Left, rRow.Bottom)
        RowGraphics.DrawLine(penBlack2, rRow.Right, rRow.Top, rRow.Right, rRow.Bottom)
        RowGraphics.DrawLine(penBlack1, rRow.Left, rRow.Top, rRow.Right, rRow.Top)
    End Sub
#End Region

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
        End With

        Dim penSelection As New Pen(Brushes.Red, 2)
        penSelection.DashStyle = Drawing2D.DashStyle.Dot
        graph.DrawRectangle(penSelection, SelectionRectangle.Rectangle)
    End Sub

    Public Sub SetSelectionRectangle()
        Dim rTopLeftCell As New Rectangle
        Dim rBottomRightCell As New Rectangle
        Dim FirstRowIndexSelection As Integer

        If EditingControl IsNot Nothing Then
            With SelectionRectangle
                .FirstRowIndex = CurrentCell.RowIndex
                .LastRowIndex = CurrentCell.RowIndex
                .FirstColumnIndex = CurrentCell.ColumnIndex - 1
                .LastColumnIndex = CurrentCell.ColumnIndex

                SetSelectionRectangleGridArea_Modify_Merged()
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


            If Me.Rows(.FirstRowIndex).Displayed = True And Me.Rows(.LastRowIndex).Displayed = False And .LastRowIndex >= Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed) Then
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
            .Y += 2
            .Width -= 4
            .Height -= 3

            If .Rectangle.Bottom = -1 And .LastRowIndex >= Me.FirstDisplayedScrollingRowIndex Then .Height = Me.Bottom - .Y - 1

            Dim intRowCount As Integer = Me.DisplayedRowCount(True)
            If .Y <= Me.ColumnHeadersHeight + 2 Then
                If .FirstRowIndex <= Me.FirstDisplayedScrollingRowIndex + intRowCount - 1 Then
                    .Y = Me.ColumnHeadersHeight + 2
                    .Height -= Me.ColumnHeadersHeight
                Else
                    SelectionRectangle.Rectangle = Nothing
                End If
            End If
        End With
    End Sub

    Public Sub SetSelectionRectangleGridArea()
        Dim Vdir As Integer, Hdir As Integer
        Dim intSelSize As Integer = SelectedCells.Count - 1
        If intSelSize < 0 Then intSelSize = 0

        If SelectedCells(0).RowIndex >= SelectedCells(intSelSize).RowIndex Then Vdir = 1 Else Vdir = -1
        If SelectedCells(0).ColumnIndex >= SelectedCells(intSelSize).ColumnIndex Then Hdir = 1 Else Hdir = -1

        If SelectedCells.Count = 0 Then
            SelectionRectangle.FirstRowIndex = CurrentCell.RowIndex
            SelectionRectangle.FirstColumnIndex = CurrentCell.ColumnIndex
        Else
            SelectionRectangle.FirstRowIndex = SelectedCells(0).RowIndex
            SelectionRectangle.FirstColumnIndex = SelectedCells(0).ColumnIndex
        End If
        SelectionRectangle.LastRowIndex = SelectionRectangle.FirstRowIndex
        SelectionRectangle.LastColumnIndex = SelectionRectangle.FirstColumnIndex

        For i = 0 To SelectedCells.Count - 1
            If SelectedCells(i).RowIndex < SelectionRectangle.FirstRowIndex Then SelectionRectangle.FirstRowIndex = SelectedCells(i).RowIndex
            If SelectedCells(i).RowIndex > SelectionRectangle.LastRowIndex Then SelectionRectangle.LastRowIndex = SelectedCells(i).RowIndex
            If SelectedCells(i).ColumnIndex < SelectionRectangle.FirstColumnIndex Then SelectionRectangle.FirstColumnIndex = SelectedCells(i).ColumnIndex
            If SelectedCells(i).ColumnIndex > SelectionRectangle.LastColumnIndex Then SelectionRectangle.LastColumnIndex = SelectedCells(i).ColumnIndex
        Next

        SetSelectionRectangleGridArea_Modify(Vdir, Hdir)
    End Sub

    Public Sub SetSelectionRectangleGridArea_Modify(ByVal Vdir As Integer, ByVal Hdir As Integer)
        Dim boolRest As Boolean
        Dim i As Integer

        If Me.Columns(SelectionRectangle.FirstColumnIndex).Name.Contains("RTF") Then SelectionRectangle.FirstColumnIndex = SelectionRectangle.FirstColumnIndex - 1
        If Me.Columns(SelectionRectangle.LastColumnIndex).Name.Contains("Sort") Then SelectionRectangle.LastColumnIndex = SelectionRectangle.LastColumnIndex + 1

        'if selection is larger than section, limit selection to that section
        If Vdir = 1 Then
            For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
                If Me.Grid(i).RowType = GridRow.RowTypes.Section And boolRest = False Then
                    boolRest = True
                    SelectionRectangle.LastRowIndex = i - 1
                End If
            Next
        Else
            For i = SelectionRectangle.LastRowIndex To SelectionRectangle.FirstRowIndex Step -1
                If Me.Grid(i).RowType = GridRow.RowTypes.Section And boolRest = False Then
                    boolRest = True
                    SelectionRectangle.FirstRowIndex = i + 1
                End If
            Next
        End If
        If SelectionRectangle.LastRowIndex < 0 Then SelectionRectangle.LastRowIndex = 0
        SetSelectionRectangleGridArea_Modify_Merged()
    End Sub

    Public Sub SetSelectionRectangleGridArea_Modify_Merged()

        With SelectionRectangle
            'if a merged Struct or Indicator is selected, select the whole Struct/indicator (not only the first row)
            Dim strFirstColName As String = Columns(.FirstColumnIndex).Name
            Dim selGridRowStart As LogframeRow = Me.Grid(.FirstRowIndex)
            Dim selGridRowEnd As LogframeRow = Me.Grid(.LastRowIndex)

            'modify selection rectangle
            If strFirstColName.Contains("Struct") Then
                SelectionRectangle_modify_merged_Struct(selGridRowStart, selGridRowEnd)

            ElseIf strFirstColName.Contains("Ind") Then
                If IsResourceBudgetRow(selGridRowStart) = False Then
                    SelectionRectangle_modify_merged_Indicator(selGridRowStart, selGridRowEnd)
                Else
                    SelectionRectangle_modify_merged_Resource(selGridRowStart, selGridRowEnd)
                End If
            ElseIf strFirstColName.Contains("Ver") Then
                If IsResourceBudgetRow(selGridRowStart) = False Then
                    SelectionRectangle_modify_merged_VerificationSource(selGridRowStart, selGridRowEnd)
                Else
                    SelectionRectangle_modify_merged_Resource(selGridRowStart, selGridRowEnd)
                End If

            ElseIf strFirstColName.Contains("Asm") Then
                SelectionRectangle_modify_merged_Assumption(selGridRowStart, selGridRowEnd)
            End If
        End With
    End Sub

    Public Sub SelectionRectangle_modify_merged_Struct(ByVal selGridRowStart As LogframeRow, ByVal selGridRowEnd As LogframeRow)
        Dim StartStruct As Struct = Nothing, EndStruct As Struct = Nothing

        With SelectionRectangle
            If selGridRowStart.StructImage IsNot Nothing And selGridRowStart.Struct Is Nothing Then
                StartStruct = Grid.GetPreviousStruct(.FirstRowIndex)
                .FirstRowIndex = Me.Grid.GetRowIndexOfStruct(StartStruct)
            End If
            If selGridRowEnd.StructImage IsNot Nothing Then
                If .LastRowIndex >= 0 Then
                    Dim intRowCount As Integer
                    If selGridRowEnd.Struct Is Nothing Then
                        EndStruct = Grid.GetPreviousStruct(.LastRowIndex)
                        .LastRowIndex = Me.Grid.GetRowIndexOfStruct(EndStruct)
                        intRowCount = Me.GetRowCountOfStruct(EndStruct)
                    Else
                        intRowCount = Me.GetRowCountOfStruct(selGridRowEnd.Struct)
                    End If

                    .LastRowIndex += intRowCount
                    If intRowCount > 0 Then .LastRowIndex -= 1
                End If
            End If
        End With
    End Sub

    Public Sub SelectionRectangle_modify_merged_Indicator(ByVal selGridRowStart As LogframeRow, ByVal selGridRowEnd As LogframeRow)
        Dim StartIndicator As Indicator = Nothing, EndIndicator As Indicator = Nothing

        With SelectionRectangle
            If selGridRowStart.IndicatorImage IsNot Nothing And selGridRowStart.Indicator Is Nothing Then
                'user clicked within a merged indicator but not on first cell
                StartIndicator = Grid.GetPreviousIndicator(.FirstRowIndex)
                .FirstRowIndex = Grid.GetRowIndexOfIndicator(StartIndicator)
            End If
            If selGridRowEnd.IndicatorImage IsNot Nothing Then
                'get bottom cell of merged indicator
                If .LastRowIndex >= 0 Then
                    Dim intRowCount As Integer
                    If selGridRowEnd.Indicator Is Nothing Then
                        EndIndicator = Grid.GetPreviousIndicator(.LastRowIndex)
                        .LastRowIndex = Grid.GetRowIndexOfIndicator(EndIndicator)
                        intRowCount = Me.GetRowCountOfIndicator(EndIndicator)
                    Else
                        intRowCount = Me.GetRowCountOfIndicator(selGridRowEnd.Indicator)
                    End If
                    .LastRowIndex += intRowCount
                    If intRowCount > 0 Then .LastRowIndex -= 1
                End If
            ElseIf selGridRowEnd.IndicatorImage Is Nothing And selGridRowEnd.StructImage IsNot Nothing Then
                'if number of assumptions > number of indicators, the bottom indicator is merged to the last line of the merged struct
                Dim selRow As LogframeRow

                For i = .LastRowIndex To RowCount - 1
                    selRow = Grid(i)
                    If selRow.Indicator IsNot Nothing Then
                        .LastRowIndex = i - 1
                        Exit For
                    ElseIf ((selRow.Struct Is Nothing And selRow.StructImage Is Nothing) Or (selRow.Struct IsNot Nothing)) And i > .FirstRowIndex Then
                        .LastRowIndex = i - 1
                        Exit For
                    ElseIf selRow.RowType <> LogframeRow.RowTypes.Normal Then
                        .LastRowIndex = i - 1
                        Exit For
                    End If
                Next
            End If
        End With
    End Sub

    Public Sub SelectionRectangle_modify_merged_Resource(ByVal selGridRowStart As LogframeRow, ByVal selGridRowEnd As LogframeRow)
        Dim selRow As LogframeRow

        'if number of assumptions > number of resources, the bottom resource is merged to the last line of the merged struct
        With SelectionRectangle
            If selGridRowStart.StructImage IsNot Nothing And selGridRowStart.ResourceImage Is Nothing Then
                For i = .FirstRowIndex To 0 Step -1
                    selRow = Grid(i)
                    If selRow.Resource IsNot Nothing Then
                        .FirstRowIndex = i + 1
                        Exit For
                    ElseIf selRow.Struct IsNot Nothing Then
                        .FirstRowIndex = i
                        Exit For
                    ElseIf selRow.RowType <> LogframeRow.RowTypes.Normal Then
                        .FirstRowIndex = i + 1
                        Exit For
                    End If
                Next
            End If
            If selGridRowEnd.StructImage IsNot Nothing And selGridRowEnd.ResourceImage Is Nothing Then
                For i = .LastRowIndex To RowCount - 1
                    selRow = Grid(i)
                    If ((selRow.Struct Is Nothing And selRow.StructImage Is Nothing) Or (selRow.Struct IsNot Nothing)) And i > .FirstRowIndex Then
                        'If selRow.Struct IsNot Nothing And i > .FirstRowIndex Then
                        .LastRowIndex = i - 1
                        Exit For
                    ElseIf selRow.RowType <> LogframeRow.RowTypes.Normal Then
                        .LastRowIndex = i - 1
                        Exit For
                    End If
                Next
            End If
        End With
    End Sub

    Public Sub SelectionRectangle_modify_merged_VerificationSource(ByVal selGridRowStart As LogframeRow, ByVal selGridRowEnd As LogframeRow)
        'Dim selRow As LogframeRow

        'if number of assumptions > number of indicators/verification sources, the bottom verification source is merged to the last line of the merged struct
        'With SelectionRectangle
        '    If selGridRowEnd.IndicatorImage Is Nothing And selGridRowEnd.StructImage IsNot Nothing Then
        '        'if number of assumptions > number of indicators, the bottom indicator is merged to the last line of the merged struct

        '        For i = .LastRowIndex To RowCount - 1
        '            selRow = Grid(i)
        '            If selRow.Indicator IsNot Nothing Then
        '                '.LastRowIndex = i - 1
        '                'Exit For
        '            ElseIf ((selRow.Struct Is Nothing And selRow.StructImage Is Nothing) Or (selRow.Struct IsNot Nothing)) And i > .FirstRowIndex Then
        '                .LastRowIndex = i - 1
        '                Exit For
        '            ElseIf selRow.RowType <> LogframeRow.RowTypes.Normal Then
        '                .LastRowIndex = i - 1
        '                Exit For
        '            End If
        '        Next
        '    End If
        'End With
    End Sub

    Public Sub SelectionRectangle_modify_merged_Assumption(ByVal selGridRowStart As LogframeRow, ByVal selGridRowEnd As LogframeRow)
        Dim selRow As LogframeRow

        With SelectionRectangle
            If selGridRowStart.StructImage IsNot Nothing And selGridRowStart.AssumptionImage Is Nothing Then
                For i = .FirstRowIndex To 0 Step -1
                    selRow = Grid(i)
                    If selRow.Assumption IsNot Nothing Then
                        .FirstRowIndex = i + 1
                        Exit For
                    ElseIf selRow.Struct IsNot Nothing Then
                        .LastRowIndex = i
                        Exit For
                    ElseIf selRow.RowType <> LogframeRow.RowTypes.Normal Then
                        .FirstRowIndex = i + 1
                        Exit For
                    End If
                Next

            End If
            If selGridRowEnd.AssumptionImage Is Nothing Then
                For i = .LastRowIndex To RowCount - 1
                    selRow = Grid(i)
                    If ((selRow.Struct Is Nothing And selRow.StructImage Is Nothing) Or (selRow.Struct IsNot Nothing)) And i > .FirstRowIndex Then
                        'If selRow.Struct IsNot Nothing Or selRow.Assumption IsNot Nothing Or selRow.RowType <> LogframeRow.RowTypes.Normal Then
                        .LastRowIndex = i - 1
                        Exit For
                    ElseIf selRow.RowType <> LogframeRow.RowTypes.Normal Then
                        .LastRowIndex = i - 1
                        Exit For
                    End If
                Next
            End If
        End With
    End Sub
#End Region 'Selection rectangle
End Class
