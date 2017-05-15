Partial Public Class DataGridViewPlanning

#Region "Mouse actions"
    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)

        If CurrentCell.IsInEditMode = False Then
            MoveCurrentCell()
            InvalidateSelectionRectangle()
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Tab Then

            Dim x As Integer = CurrentCell.ColumnIndex
            Dim y As Integer = CurrentRow.Index
            Dim boolShift As Boolean = e.Shift

            If y = Me.RowCount - 1 Then
                CurrentCell = Me(x, y - 1)
            Else
                CurrentCell = Me(x, y + 1)
            End If

            CurrentCell = Me(x, y)

            MoveCurrentCell()
            CurrentCell.Selected = True

            InvalidateSelectionRectangle()
        End If
        MyBase.OnKeyUp(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If hit.ColumnIndex > 0 And hit.RowIndex >= 0 Then
            Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
            Dim selGridRow As PlanningGridRow = Me.Grid(hit.RowIndex)

            If e.Button = MouseButtons.Right Then
                If hit.Type = DataGridViewHitTestType.Cell Then
                    ' Create a rectangle using the DragSize, with the mouse position being at the center of the rectangle.
                    OnMouseDown_SetDragRectangle(e.Location, selCell)
                End If
            Else
                If selCell.ColumnIndex < CONST_PlanningColumnIndex Then
                    'determine where exactly in the text the user clicked
                    OnMouseDown_SetClickPoint(e.Location)
                Else
                    If DragLink = True Then
                        'start linking activity or key moment
                        OnMouseDown_StartLink(e.Location, selGridRow)
                    ElseIf Unlink = True Then
                        ' do nothing
                    Else
                        'start drag of activity or key moment
                        OnMouseDown_StartDrag(e.Location, selGridRow)
                    End If
                End If

                DragBoxFromMouseDown = Rectangle.Empty
            End If
        End If
        MyBase.OnMouseDown(e)
        Invalidate()
    End Sub

    Private Sub OnMouseDown_SetDragRectangle(ByVal MouseLocation As Point, ByVal selCell As DataGridViewCell)
        Dim dragSize As Size = SystemInformation.DragSize
        DragBoxFromMouseDown = New Rectangle(New Point(MouseLocation.X - (dragSize.Width / 2), MouseLocation.Y - (dragSize.Height / 2)), dragSize)

        If SelectionRectangle.Rectangle.Contains(MouseLocation) = False Then Me.CurrentCell = selCell
    End Sub

    Private Sub OnMouseDown_SetClickPoint(ByVal MouseLocation As Point)
        If SelectionRectangle.Rectangle.Contains(MouseLocation) = True Then
            ClickPoint = MouseLocation
        Else
            ClickPoint = Nothing
        End If
    End Sub

    Private Sub OnMouseDown_StartLink(ByVal MouseLocation As Point, ByVal selGridRow As PlanningGridRow)
        LinkRowIndex = Grid.IndexOf(selGridRow)
        If selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
            LinkActivity = TryCast(selGridRow.Struct, Activity)

            If rDragActivityStart.Contains(MouseLocation) Then
                boolLinkActivityStart = True
            ElseIf rDragActivityEnd.Contains(MouseLocation) Then
                boolLinkActivityEnd = True
            End If
        ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            LinkKeyMoment = selGridRow.KeyMoment

            If rDragKeyMoment.Contains(MouseLocation) Then
                boolLinkKeyMoment = True
            End If
        End If
    End Sub

    Private Sub OnMouseDown_StartDrag(ByVal MouseLocation As Point, ByVal selGridRow As PlanningGridRow)
        If selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
            DragActivity = TryCast(selGridRow.Struct, Activity)

            If DragActivity IsNot Nothing Then
                With DragActivity.ActivityDetail
                    If .Relative = True Then
                        Dim RefMoment As Object = CurrentLogFrame.GetReferenceMomentByGuid(.GuidReferenceMoment)

                        If RefMoment Is Nothing Then Exit Sub
                        Select Case RefMoment.GetType
                            Case GetType(Activity)
                                RefActivity = CType(RefMoment, Activity)
                                RefKeyMoment = Nothing
                            Case GetType(KeyMoment)
                                RefKeyMoment = CType(RefMoment, KeyMoment)
                                RefActivity = Nothing
                        End Select
                    End If
                End With

                If rDragActivity.Contains(MouseLocation) Then
                    Dim intRowIndex As Integer = Me.Grid.IndexOf(selGridRow)
                    Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, intRowIndex, False)
                    rCell.X -= Me.HorizontalScrollingOffset

                    datInitialDragMoment = CoordinateToDate(rCell.X, MouseLocation.X)
                    boolDragActivity = True
                ElseIf rDragActivityStart.Contains(MouseLocation) Then
                    boolDragActivityStart = True
                ElseIf rDragActivityEnd.Contains(MouseLocation) Then
                    boolDragActivityEnd = True
                ElseIf rDragPreparationStart.Contains(MouseLocation) Then
                    boolDragPreparationStart = True
                ElseIf rDragFollowUpEnd.Contains(MouseLocation) Then
                    boolDragFollowUpEnd = True
                End If
            End If
        ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            DragKeyMoment = selGridRow.KeyMoment

            If DragKeyMoment IsNot Nothing Then
                With DragKeyMoment
                    If .Relative = True Then
                        Dim RefMoment As Object = CurrentLogFrame.GetReferenceMomentByGuid(.GuidReferenceMoment)

                        If RefMoment Is Nothing Then Exit Sub
                        Select Case RefMoment.GetType
                            Case GetType(Activity)
                                RefActivity = CType(RefMoment, Activity)
                                RefKeyMoment = Nothing
                            Case GetType(KeyMoment)
                                RefKeyMoment = CType(RefMoment, KeyMoment)
                                RefActivity = Nothing
                        End Select
                    End If
                End With

                If rDragKeyMoment.Contains(MouseLocation) Then
                    boolDragKeyMoment = True
                End If
            End If
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim ptMouseLocation As New Point(e.X, e.Y)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If ((e.Button And MouseButtons.Right) = MouseButtons.Right) Then
            'drag items
            OnMouseMove_DragItems(e.Location)
        ElseIf ((e.Button And MouseButtons.Left) = MouseButtons.Left) Then

            If hit.ColumnIndex = CONST_PlanningColumnIndex Then
                Dim rColumn As Rectangle = GetColumnDisplayRectangle(CONST_PlanningColumnIndex, True)
                Dim intOffSet As Integer

                rColumn.Width -= 200
                rColumn.X += 100

                'when dragging in the vicinity of the columns edges, scroll left or right
                If e.X < rColumn.X And HorizontalScrollingOffset > 0 Then
                    intOffSet = HorizontalScrollingOffset
                    intOffSet -= 10
                    If intOffSet < 0 Then intOffSet = 0
                    HorizontalScrollingOffset = intOffSet
                ElseIf e.X > rColumn.Right Then
                    intOffSet = HorizontalScrollingOffset
                    intOffSet += 10
                    'If intOffSet < 0 Then intOffSet = 0
                    HorizontalScrollingOffset = intOffSet
                End If

                If DragLink = True Then
                    'drag to establish link
                    Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
                    Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)

                    OnMouseMove_DragLink(e.Location, selGridRow, rCell)
                ElseIf Unlink = True Then
                    'do nothing
                Else
                    'drag activity/key moment
                    OnMouseMove_DragActivityOrKeyMoment(e.Location)
                End If
            End If
        Else
            'determine areas of start of activity, end of activity, body of activity, start of preparation and end of follow-up
            'if the cursor hovers over any of these areas, change it to double arrow cursor

            If hit.ColumnIndex = CONST_PlanningColumnIndex Then
                Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset

                If selGridRow Is Nothing Then Exit Sub

                If selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
                    OnMouseMove_ActivityAreas(ptMouseLocation, selGridRow, rCell)
                ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    OnMouseMove_KeyMomentAreas(ptMouseLocation, selGridRow, rCell)
                End If
            End If

        End If
        MyBase.OnMouseMove(e)

        InvalidateSelectionRectangle()
    End Sub

    Private Sub OnMouseMove_DragItems(ByVal MouseLocation As Point)
        Dim p1 As New Point(Me.Location.X, Me.Location.Y)
        Dim p2 As New Point(Me.Location.X, Me.Location.Y + Me.Height)

        'if cursor nears the rim of the datagridview while dragging, scroll up or down
        If MouseLocation.Y < p1.Y + 50 And FirstDisplayedScrollingRowIndex > 0 Then _
            FirstDisplayedScrollingRowIndex -= 1
        If MouseLocation.Y > p2.Y - 50 And Rows(Me.RowCount - 1).Displayed = False Then _
            FirstDisplayedScrollingRowIndex += 1

        'select cursor
        If Control.ModifierKeys = Keys.Control Then
            Cursor.Current = Cursors.HSplit
            DragOperator = DragOperatorValues.Copy
        Else
            Cursor.Current = Cursors.Hand
            DragOperator = DragOperatorValues.Move
        End If

        Drag(MouseLocation)
    End Sub

    Private Sub OnMouseMove_DragLink(ByVal MouseLocation As Point, ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle)
        Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
        Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
        boolLinkHoverOverActivity = False
        boolLinkHoverOverKeyMoment = False

        If selActivity IsNot Nothing Then
            rLinkActivity = rCell

            With selActivity.ActivityDetail
                rLinkActivity.X = GetCoordinateX(rCell.X, .StartDateMainActivity)
                rLinkActivity.Width = GetActivityWidth(.StartDateMainActivity, .EndDateMainActivity)

                If rLinkActivity.Contains(MouseLocation) Then boolLinkHoverOverActivity = True
            End With
        ElseIf selKeyMoment IsNot Nothing Then
            rLinkKeyMoment = rCell

            With selKeyMoment
                rLinkKeyMoment.X = GetCoordinateX(rCell.X, .ExactDateKeyMoment)
                rLinkKeyMoment.X -= CONST_BarHeight
                rLinkKeyMoment.Width = (CONST_BarHeight * 2)

                If rLinkKeyMoment.Contains(MouseLocation) Then boolLinkHoverOverKeyMoment = True
            End With
        End If

        If boolLinkKeyMoment = True Then
            LinkFromKeyMoment(MouseLocation)
        ElseIf boolLinkActivityStart = True Then
            LinkFromActivityStart(MouseLocation)
        ElseIf boolLinkActivityEnd = True Then
            LinkFromActivityEnd(MouseLocation)
        End If

        Dim rLinkArea As New Rectangle
        If ptLinkStart.X < ptLinkEnd.X Then rLinkArea.X = ptLinkStart.X Else rLinkArea.X = ptLinkEnd.X
        If ptLinkStart.Y < ptLinkEnd.Y Then rLinkArea.Y = ptLinkStart.Y Else rLinkArea.Y = ptLinkEnd.Y
        rLinkArea.X -= 50
        rLinkArea.Y -= 50
        rLinkArea.Width = Math.Abs(ptLinkEnd.X - ptLinkStart.X) + 100
        rLinkArea.Height = Math.Abs(ptLinkEnd.Y - ptLinkStart.Y) + 100
        Invalidate(rLinkArea)
    End Sub

    Private Sub OnMouseMove_DragActivityOrKeyMoment(ByVal MouseLocation As Point)
        If boolDragKeyMoment = True Then
            DragMoveKeyMoment(MouseLocation)
        ElseIf boolDragActivity = True Then
            DragMoveActivity(MouseLocation)
        ElseIf boolDragActivityStart = True Then
            DragActivityStart(MouseLocation)
        ElseIf boolDragActivityEnd = True Then
            DragActivityEnd(MouseLocation)
        ElseIf boolDragPreparationStart = True Then
            DragPreparationStart(MouseLocation)
        ElseIf boolDragFollowUpEnd = True Then
            DragFollowUpEnd(MouseLocation)
        End If
    End Sub

    Private Sub OnMouseMove_ActivityAreas(ByVal ptMouseLocation As Point, ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle)
        Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
        Dim boolOver, boolIn As Boolean

        If selActivity IsNot Nothing Then

            rDragActivity = rCell
            rDragActivityStart = New Rectangle(rCell.X, rCell.Y, CONST_SelectRadius, rCell.Height)
            rDragActivityEnd = rDragActivityStart

            With selActivity.ActivityDetail
                rDragActivity.X = GetCoordinateX(rCell.X, .StartDateMainActivity)
                rDragActivity.Width = GetActivityWidth(.StartDateMainActivity, .EndDateMainActivity)

                rDragActivityStart.X = rDragActivity.X - CONST_SelectRadius
                rDragActivityEnd.X = rDragActivity.Right
                If rDragActivityStart.Contains(ptMouseLocation) Then
                    If Unlink = True Then
                        With selActivity.ActivityDetail
                            If .Relative = True And .PeriodDirection = ActivityDetail.DirectionValues.After Then _
                                boolOver = True
                        End With
                    Else
                        boolOver = True
                    End If

                    rDragActivityEnd = Nothing
                    rDragPreparationStart = Nothing
                    rDragFollowUpEnd = Nothing
                ElseIf rDragActivityEnd.Contains(ptMouseLocation) Then
                    If Unlink = True Then
                        With selActivity.ActivityDetail
                            If .Relative = True And .PeriodDirection = ActivityDetail.DirectionValues.Before Then _
                                boolOver = True
                        End With
                    Else
                        boolOver = True
                    End If

                    rDragActivityStart = Nothing
                    rDragPreparationStart = Nothing
                    rDragFollowUpEnd = Nothing
                ElseIf DragLink = False And Unlink = False Then
                    If .StartDatePreparation > Date.MinValue And .StartDatePreparation < .StartDateMainActivity Then
                        rDragPreparationStart = rCell
                        rDragPreparationStart.X = GetCoordinateX(rCell.X, .StartDatePreparation) - CONST_SelectRadius
                        rDragPreparationStart.Width = CONST_SelectRadius
                        If rDragPreparationStart.Contains(ptMouseLocation) Then
                            boolOver = True
                        Else
                            rDragActivityStart = Nothing
                            rDragActivityEnd = Nothing
                            rDragFollowUpEnd = Nothing
                        End If
                    End If
                    If .EndDateFollowUp > Date.MinValue And .EndDateFollowUp > .EndDateMainActivity Then
                        rDragFollowUpEnd = rCell
                        rDragFollowUpEnd.X = GetCoordinateX(rCell.X, .EndDateFollowUp)
                        rDragFollowUpEnd.Width = CONST_SelectRadius
                        If rDragFollowUpEnd.Contains(ptMouseLocation) Then
                            boolOver = True
                        Else
                            rDragActivityStart = Nothing
                            rDragActivityEnd = Nothing
                            rDragPreparationStart = Nothing
                        End If
                    End If
                End If

                If boolOver = False And DragLink = False And Unlink = False Then
                    rDragActivity.X += CONST_SelectRadius
                    rDragActivity.Width -= (CONST_SelectRadius * 2)
                    If rDragActivity.Contains(ptMouseLocation) Then
                        boolIn = True
                    Else
                        rDragActivityStart = Nothing
                        rDragActivityEnd = Nothing
                        rDragPreparationStart = Nothing
                        rDragFollowUpEnd = Nothing
                    End If
                End If
            End With
        End If

        OnMouseOver_SetCursor(boolOver, boolIn)
    End Sub

    Private Sub OnMouseMove_KeyMomentAreas(ByVal ptMouseLocation As Point, ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle)
        Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
        Dim boolIn As Boolean

        If selKeyMoment IsNot Nothing Then

            rDragKeyMoment = rCell

            With selKeyMoment
                rDragKeyMoment.X = GetCoordinateX(rCell.X, .ExactDateKeyMoment)

                rDragKeyMoment.X -= CONST_BarHeight
                rDragKeyMoment.Width = (CONST_BarHeight * 2)
                If rDragKeyMoment.Contains(ptMouseLocation) Then
                    If Unlink = True Then
                        With selKeyMoment
                            If .Relative = True Then boolIn = True
                        End With
                    Else
                        boolIn = True
                    End If
                Else
                    rDragKeyMoment = Nothing
                End If
            End With
        End If

        OnMouseOver_SetCursor(False, boolIn)
    End Sub

    Private Sub OnMouseOver_SetCursor(ByVal boolOver As Boolean, ByVal boolIn As Boolean)
        If boolOver = True Then
            If DragLink = True Then
                Cursor.Current = Cursors.Cross
            ElseIf Unlink = True Then
                Cursor.Current = Cursors.VSplit
            Else
                Cursor.Current = Cursors.SizeWE
            End If
        ElseIf boolIn = True Then
            If DragLink = True Then 'only for key moments
                Cursor.Current = Cursors.Cross
            ElseIf Unlink = True Then
                Cursor.Current = Cursors.VSplit 'only for key moments
            Else
                Cursor.Current = Cursors.Hand
            End If
        Else
            Cursor.Current = Cursors.Default
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim ptMouseLocation As New Point(e.X, e.Y)

        DragReleased = True

        InvalidateSelectionRectangle()
        If e.Button = MouseButtons.Right Then
            If Control.ModifierKeys = Keys.Control Then
                DragOperator = DragOperatorValues.Copy
            Else
                DragOperator = DragOperatorValues.Move
            End If
            Drop(e.X, e.Y)
        ElseIf e.Button = MouseButtons.Left Then
            Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)
            Dim intRowIndex As Integer = hit.RowIndex

            If DragLink = True Then
                If boolLinkKeyMoment = True Then
                    LinkFromKeyMoment_Connect(ptMouseLocation)
                ElseIf boolLinkActivityStart = True Then
                    LinkFromActivityStart_Connect(ptMouseLocation)
                ElseIf boolLinkActivityEnd = True Then
                    LinkFromActivityEnd_Connect(ptMouseLocation)
                End If

                RaiseEvent DragLinkReleased()
                Me.ClearSelection()
                Me(CONST_PlanningColumnIndex, intRowIndex).Selected = True
            ElseIf Unlink = True Then
                UnlinkClicked(ptMouseLocation)

                RaiseEvent UnlinkReleased()
                Me.ClearSelection()
                Me(CONST_PlanningColumnIndex, intRowIndex).Selected = True
            Else
                If boolDragKeyMoment = True Then
                    DropKeyMoment(ptMouseLocation)
                ElseIf boolDragActivity = True Then
                    DropActivity(ptMouseLocation)
                ElseIf boolDragActivityStart = True Then
                    DropActivityStart(ptMouseLocation)
                ElseIf boolDragActivityEnd = True Then
                    DropActivityEnd(ptMouseLocation)
                ElseIf boolDragPreparationStart = True Then
                    DropPreparationStart(ptMouseLocation)
                ElseIf boolDragFollowUpEnd = True Then
                    DropFollowUpEnd(ptMouseLocation)
                End If
            End If

        End If
        DragBoxFromMouseDown = Rectangle.Empty

        DragLink = False
        Unlink = False
        LinkActivity = Nothing
        LinkKeyMoment = Nothing
        boolLinkKeyMoment = False
        boolLinkActivityStart = False
        boolLinkActivityEnd = False
        boolLinkHoverOverActivity = False
        boolLinkHoverOverKeyMoment = False
        rLinkKeyMoment = Nothing
        rLinkActivity = Nothing
        LinkRowIndex = -1
        ptLinkStart = Nothing
        ptLinkEnd = Nothing

        DragActivity = Nothing
        DragKeyMoment = Nothing
        RefActivity = Nothing
        RefKeyMoment = Nothing
        boolDragActivity = False
        boolDragActivityStart = False
        boolDragActivityEnd = False
        boolDragPreparationStart = False
        boolDragFollowUpEnd = False
        boolDragKeyMoment = False
        rDragRectangle = Nothing
        rDragActivity = Nothing
        rDragActivityStart = Nothing
        rDragActivityEnd = Nothing
        rDragPreparationStart = Nothing
        rDragFollowUpEnd = Nothing
        rDragKeyMoment = Nothing

        MyBase.OnMouseUp(e)
        Invalidate()
    End Sub
#End Region

#Region "Drag and drop"
    Public Sub Drag(ByVal MouseLocation As Point)
        ' If the mouse moves outside the rectangle, start the drag.
        If (Rectangle.op_Inequality(DragBoxFromMouseDown, Rectangle.Empty) And _
            Not DragBoxFromMouseDown.Contains(MouseLocation)) Then

            If DragReleased = True Then
                If DragOperator = DragOperatorValues.Copy Then
                    CopyItems()
                Else
                    CutItems(False)
                End If
                DragReleased = False
            End If

            Dim hit As DataGridView.HitTestInfo = HitTest(MouseLocation.X, MouseLocation.Y)
            If hit.Type = DataGridViewHitTestType.Cell Then
                Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
                Dim intLastColIndex As Integer = 1

                If Me.ShowDatesColumns = True Then intLastColIndex = 3

                SelectionRectangle.FirstRowIndex = selCell.RowIndex
                SelectionRectangle.LastRowIndex = selCell.RowIndex
                If SelectionRectangle.LastRowIndex > Me.RowCount - 1 Then SelectionRectangle.LastRowIndex = Me.RowCount - 1

                'draw insert/copy indicator line
                Dim rSelStart As Rectangle = GetCellDisplayRectangle(0, SelectionRectangle.FirstRowIndex, False)
                Dim rSelEnd As Rectangle = GetCellDisplayRectangle(intLastColIndex, SelectionRectangle.LastRowIndex, False)

                Dim intVertDivider As Integer = rSelStart.Top + ((rSelEnd.Bottom - rSelStart.Top) / 2)
                pStartInsertLine.X = rSelStart.Left
                pEndInsertLine.X = rSelEnd.Right - 1
                If MouseLocation.Y <= intVertDivider Then pStartInsertLine.Y = rSelStart.Top Else pStartInsertLine.Y = rSelEnd.Bottom
                pEndInsertLine.Y = pStartInsertLine.Y

                Dim graph As Graphics = CreateGraphics()
                If Not (pStartInsertLine = Nothing Or pEndInsertLine = Nothing) Then
                    If pStartInsertLine <> pStartInsertLineOld Then
                        Invalidate()
                        Update()

                        Dim penGreen2 As New Pen(Color.Green, 2)
                        graph.DrawLine(penGreen2, pStartInsertLine, pEndInsertLine)
                        Dim triangleStart(2) As Point
                        triangleStart(0) = New Point(pStartInsertLine.X, pStartInsertLine.Y - 6)
                        triangleStart(1) = New Point(pStartInsertLine.X, pStartInsertLine.Y + 6)
                        triangleStart(2) = New Point(pStartInsertLine.X + 6, pStartInsertLine.Y)
                        graph.FillPolygon(Brushes.Green, triangleStart)
                        Dim triangleEnd(2) As Point
                        triangleEnd(0) = New Point(pEndInsertLine.X, pEndInsertLine.Y - 6)
                        triangleEnd(1) = New Point(pEndInsertLine.X, pEndInsertLine.Y + 6)
                        triangleEnd(2) = New Point(pEndInsertLine.X - 8, pEndInsertLine.Y)
                        graph.FillPolygon(Brushes.Green, triangleEnd)
                        pStartInsertLineOld = pStartInsertLine
                        pEndInsertLineOld = pEndInsertLine
                    End If
                End If

            Else
                pStartInsertLine = Nothing
                pEndInsertLine = Nothing
            End If
        End If
    End Sub

    Public Sub Drop(ByVal mouseX As Integer, ByVal mouseY As Integer)
        Dim hit As DataGridView.HitTestInfo = HitTest(pStartInsertLine.X, pStartInsertLine.Y)

        If hit.Type = DataGridViewHitTestType.Cell Then
            Dim CopyGroup As ClipboardItems = ItemClipboard.GetCopyGroup()
            Dim selCell As DataGridViewCell

            If hit.RowIndex < Me.RowCount - 1 Then
                selCell = Me(hit.ColumnIndex, hit.RowIndex)
            Else
                selCell = Me(hit.ColumnIndex, Me.RowCount - 2)
            End If

            PasteItems(CopyGroup, ClipboardItems.PasteOptions.PasteAll, selCell)
        End If
    End Sub

    Private Function Drop_SetPeriod(ByVal selActivityDetail As ActivityDetail, ByVal datStartPeriod As Date, ByVal datEndPeriod As Date) As ActivityDetail
        Dim intNrDays, intMonth, intMonthOnly, intYear As Integer

        intYear = datEndPeriod.Year - datStartPeriod.Year
        intMonth = datEndPeriod.Month - datStartPeriod.Month + (intYear * 12)
        intMonthOnly = intMonth - (intYear * 12)
        If intMonthOnly < 0 Then intMonthOnly = intMonth

        If datEndPeriod.Day > datStartPeriod.Day Then
            intNrDays = datEndPeriod.Day - datStartPeriod.Day + 1
        Else
            If datEndPeriod.Day <> datStartPeriod.Day Then
                intNrDays = datEndPeriod.Day + (Date.DaysInMonth(datStartPeriod.Year, datStartPeriod.Month) - datStartPeriod.Day) + 1
                intMonth -= 1
            End If
        End If

        With selActivityDetail
            If intYear > 0 And intMonthOnly = 0 And intNrDays = 0 Then
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Period", selActivityDetail.Period)
                .Period = intYear
                UndoRedo.ValueChanged(.Period)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "PeriodUnit", selActivityDetail.PeriodUnit)
                .PeriodUnit = ActivityDetail.DurationUnits.Year
                UndoRedo.OptionChanged(.PeriodUnit)
            ElseIf intMonth > 0 And intNrDays = 0 Then
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Period", selActivityDetail.Period)
                .Period = intMonth
                UndoRedo.ValueChanged(.Period)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "PeriodUnit", selActivityDetail.PeriodUnit)
                .PeriodUnit = ActivityDetail.DurationUnits.Month
                UndoRedo.OptionChanged(.PeriodUnit)
            Else
                Dim tsPeriod As TimeSpan = datEndPeriod.Subtract(datStartPeriod)
                intNrDays = tsPeriod.Days

                SetStartDateByDate_SetPeriod(selActivityDetail, intNrDays)
            End If

        End With

        Return selActivityDetail
    End Function

    Private Function Drop_SetDuration(ByVal selActivityDetail As ActivityDetail, ByVal datStartPeriod As Date, ByVal datEndPeriod As Date) As ActivityDetail
        Dim intNrDays, intMonth, intMonthOnly, intYear As Integer

        intYear = datEndPeriod.Year - datStartPeriod.Year
        intMonth = datEndPeriod.Month - datStartPeriod.Month + (intYear * 12)
        intMonthOnly = intMonth - (intYear * 12)
        If intMonthOnly < 0 Then intMonthOnly = intMonth

        If datEndPeriod.Day > datStartPeriod.Day Then
            intNrDays = datEndPeriod.Day - datStartPeriod.Day
        Else
            If datEndPeriod.Day <> datStartPeriod.Day - 1 Then
                intNrDays = datEndPeriod.Day + (Date.DaysInMonth(datStartPeriod.Year, datStartPeriod.Month) - datStartPeriod.Day)
                intMonth -= 1
            End If
        End If

        With selActivityDetail
            If intYear > 0 And intMonthOnly = 0 And intNrDays = 0 Then
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Duration", selActivityDetail.Duration)
                .Duration = intYear
                UndoRedo.ValueChanged(.Duration)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "DurationUnit", selActivityDetail.DurationUnit)
                .DurationUnit = ActivityDetail.DurationUnits.Year
                UndoRedo.OptionChanged(.DurationUnit)
            ElseIf intMonth > 0 And intNrDays = 0 Then
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Duration", selActivityDetail.Duration)
                .Duration = intMonth
                UndoRedo.ValueChanged(.Duration)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "DurationUnit", selActivityDetail.DurationUnit)
                .DurationUnit = ActivityDetail.DurationUnits.Month
                UndoRedo.OptionChanged(.DurationUnit)
            Else
                Dim tsDuration As TimeSpan = datEndPeriod.Subtract(datStartPeriod)
                intNrDays = tsDuration.Days + 1

                SetDurationByEndDate_SetDuration(selActivityDetail, intNrDays)
            End If

        End With

        Return selActivityDetail
    End Function

    Private Sub DragMoveKeyMoment(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment

            If selKeyMoment IsNot Nothing And selKeyMoment Is DragKeyMoment Then
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset
                Dim rKeyMoment As Rectangle = rCell
                datDragMoment = CoordinateToDate(rCell.X, ptMouseLocation.X)

                With selKeyMoment
                    If .Relative = True Then
                        If RefKeyMoment IsNot Nothing Then
                            If .PeriodDirection = KeyMoment.DirectionValues.Before Then
                                If datDragMoment > RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = KeyMoment.DirectionValues.After Then
                                If datDragMoment < RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        ElseIf RefActivity IsNot Nothing Then
                            If .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                                If datDragMoment > RefActivity.ActivityDetail.StartDateMainActivity Then
                                    datDragMoment = RefActivity.ActivityDetail.StartDateMainActivity
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                If datDragMoment < RefActivity.ActivityDetail.EndDateMainActivity Then
                                    datDragMoment = RefActivity.ActivityDetail.EndDateMainActivity
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        End If
                    End If
                    rKeyMoment.X = GetCoordinateX(rCell.X, datDragMoment) - (CONST_BarHeight / 2)
                    rKeyMoment.Y = GetCoordinateY(rCell.Y, rCell.Height)
                    rKeyMoment.Width = CONST_BarHeight
                    rKeyMoment.Height = CONST_BarHeight
                End With

                rDragRectangle = rKeyMoment
            Else
                boolDragKeyMoment = False
                DragKeyMoment = Nothing
                rDragRectangle = Nothing
            End If
        End If
    End Sub

    Private Sub DropKeyMoment(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment

            If selKeyMoment IsNot Nothing Then
                With selKeyMoment

                    If .Relative = False Then
                        UndoRedo.UndoBuffer_Initialise(selKeyMoment, "KeyMoment", selKeyMoment.KeyMoment)
                        .KeyMoment = datDragMoment
                        UndoRedo.DateChanged(.KeyMoment)
                    Else
                        Dim tsPeriod As TimeSpan
                        If .PeriodDirection = KeyMoment.DirectionValues.After Then
                            If RefKeyMoment IsNot Nothing Then
                                tsPeriod = datDragMoment.Subtract(RefKeyMoment.ExactDateKeyMoment)
                            ElseIf RefActivity IsNot Nothing Then
                                tsPeriod = datDragMoment.Subtract(RefActivity.ActivityDetail.EndDateMainActivity)
                            End If
                        ElseIf .PeriodDirection = KeyMoment.DirectionValues.Before Then
                            If RefKeyMoment IsNot Nothing Then
                                tsPeriod = RefKeyMoment.ExactDateKeyMoment.Subtract(datDragMoment)
                            ElseIf RefActivity IsNot Nothing Then
                                tsPeriod = RefActivity.ActivityDetail.StartDateMainActivity.Subtract(datDragMoment)
                            End If
                        End If

                        SetStartDateByDate_SetPeriod(selKeyMoment, tsPeriod.Days)
                    End If

                End With
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub DragMoveActivity(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing And selActivity Is DragActivity Then
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset
                Dim rActivity As Rectangle = rCell
                Dim datDragActivityStart, datDragActivityEnd As Date
                Dim intDaysMoved As Integer

                datDragMoment = CoordinateToDate(rCell.X, ptMouseLocation.X)
                intDaysMoved = datDragMoment.Subtract(datInitialDragMoment).Days

                With selActivity.ActivityDetail
                    datDragActivityStart = .StartDateMainActivity.AddDays(intDaysMoved)
                    datDragActivityEnd = .EndDateMainActivity.AddDays(intDaysMoved)

                    If .Relative = True Then
                        If RefActivity IsNot Nothing Then
                            If .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                                If datDragActivityEnd > RefActivity.ExactStartDate Then
                                    datDragMoment = RefActivity.ExactStartDate.AddDays((intDaysMoved * -1))
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                If datDragActivityStart < RefActivity.ExactEndDate Then
                                    datDragMoment = RefActivity.ExactEndDate.AddDays(intDaysMoved)
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        ElseIf RefKeyMoment IsNot Nothing Then
                            If .PeriodDirection = KeyMoment.DirectionValues.Before Then
                                If datDragActivityEnd > RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = KeyMoment.DirectionValues.After Then
                                If datDragActivityStart < RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        End If
                    End If
                    rActivity.X = GetCoordinateX(rCell.X, datDragActivityStart)
                    rActivity.Y = GetCoordinateY(rCell.Y, rCell.Height)
                    rActivity.Width = GetActivityWidth(.StartDateMainActivity, .EndDateMainActivity)
                    rActivity.Height = CONST_BarHeight
                End With

                rDragRectangle = rActivity
            Else
                boolDragActivityStart = False
                DragActivity = Nothing
                rDragRectangle = Nothing
                datInitialDragMoment = Nothing
            End If
        End If
    End Sub

    Private Sub DropActivity(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
            Dim datStartPeriod, datEndPeriod As Date

            If selActivity IsNot Nothing AndAlso datDragMoment <> Date.MinValue Then
                With selActivity.ActivityDetail
                    Dim intDaysMoved As Integer = datDragMoment.Subtract(datInitialDragMoment).Days
                    Dim datDragActivityStart As Date = .StartDateMainActivity.AddDays(intDaysMoved)

                    If .Relative = False Then
                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "StartDate", selActivity.ActivityDetail.StartDate)
                        .StartDate = datDragActivityStart
                        UndoRedo.DateChanged(.StartDate)
                    Else
                        If .PeriodDirection = ActivityDetail.DirectionValues.After Then
                            datEndPeriod = datDragActivityStart
                            If RefActivity IsNot Nothing Then
                                datStartPeriod = RefActivity.ActivityDetail.EndDateMainActivity
                            ElseIf RefKeyMoment IsNot Nothing Then
                                datStartPeriod = RefKeyMoment.ExactDateKeyMoment
                            End If
                        ElseIf .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                            datStartPeriod = datDragActivityStart
                            If RefActivity IsNot Nothing Then
                                datEndPeriod = RefActivity.ActivityDetail.StartDateMainActivity
                            ElseIf RefKeyMoment IsNot Nothing Then
                                datEndPeriod = RefKeyMoment.ExactDateKeyMoment
                            End If
                        End If
                        selActivity.ActivityDetail = Drop_SetPeriod(selActivity.ActivityDetail, datStartPeriod, datEndPeriod)
                    End If
                End With
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub DragActivityStart(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing And selActivity Is DragActivity Then
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset
                Dim rActivity As Rectangle = rCell
                Dim intWidth As Integer
                datDragMoment = CoordinateToDate(rCell.X, ptMouseLocation.X)

                With selActivity.ActivityDetail
                    If datDragMoment > .EndDateMainActivity Then
                        datDragMoment = .EndDateMainActivity
                        ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                    End If
                    If .Relative = True Then
                        If RefActivity IsNot Nothing Then
                            If .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                                If datDragMoment > RefActivity.ExactStartDate Then
                                    datDragMoment = RefActivity.ExactStartDate
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                If datDragMoment < RefActivity.ExactEndDate Then
                                    datDragMoment = RefActivity.ExactEndDate
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        ElseIf RefKeyMoment IsNot Nothing Then
                            If .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                                If datDragMoment > RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                If datDragMoment < RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        End If
                    End If
                    rActivity.X = GetCoordinateX(rCell.X, .StartDateMainActivity)
                    rActivity.Y = GetCoordinateY(rCell.Y, rCell.Height)
                    rActivity.Width = GetActivityWidth(.StartDateMainActivity, .EndDateMainActivity)
                    rActivity.Height = CONST_BarHeight
                End With

                intWidth = rActivity.Right - ptMouseLocation.X + 1
                rDragRectangle = New Rectangle(ptMouseLocation.X, rActivity.Y, intWidth, CONST_BarHeight)
            Else
                boolDragActivityStart = False
                DragActivity = Nothing
                rDragRectangle = Nothing
            End If
        End If
    End Sub

    Private Sub DropActivityStart(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
            Dim datStartPeriod, datEndPeriod As Date

            If selActivity IsNot Nothing Then
                With selActivity.ActivityDetail
                    If .Relative = False Then
                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "StartDate", selActivity.ActivityDetail.StartDate)
                        .StartDate = datDragMoment
                        UndoRedo.DateChanged(.StartDate)
                    Else
                        If .PeriodDirection = ActivityDetail.DirectionValues.After Then
                            datEndPeriod = datDragMoment
                            If RefActivity IsNot Nothing Then
                                datStartPeriod = RefActivity.ActivityDetail.EndDateMainActivity
                            ElseIf RefKeyMoment IsNot Nothing Then
                                datStartPeriod = RefKeyMoment.ExactDateKeyMoment
                            End If
                        ElseIf .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                            datStartPeriod = datDragMoment
                            If RefActivity IsNot Nothing Then
                                datEndPeriod = RefActivity.ActivityDetail.StartDateMainActivity
                            ElseIf RefKeyMoment IsNot Nothing Then
                                datEndPeriod = RefKeyMoment.ExactDateKeyMoment
                            End If

                        End If
                        selActivity.ActivityDetail = Drop_SetPeriod(selActivity.ActivityDetail, datStartPeriod, datEndPeriod)
                    End If

                    selActivity.ActivityDetail = Drop_SetDuration(selActivity.ActivityDetail, datDragMoment, .EndDateMainActivity)
                End With
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub DragActivityEnd(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing And selActivity Is DragActivity Then
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset
                Dim rActivity As Rectangle = rCell
                Dim intWidth As Integer
                datDragMoment = CoordinateToDate(rCell.X, ptMouseLocation.X)

                With selActivity.ActivityDetail
                    If datDragMoment < .StartDateMainActivity Then
                        datDragMoment = .StartDateMainActivity
                        ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                    End If
                    If .Relative = True Then
                        If RefActivity IsNot Nothing Then
                            If .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                                If datDragMoment > RefActivity.ExactStartDate Then
                                    datDragMoment = RefActivity.ExactStartDate
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                If datDragMoment < RefActivity.ExactEndDate Then
                                    datDragMoment = RefActivity.ExactEndDate
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        ElseIf RefKeyMoment IsNot Nothing Then
                            If .PeriodDirection = ActivityDetail.DirectionValues.Before Then
                                If datDragMoment > RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            ElseIf .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                If datDragMoment < RefKeyMoment.ExactDateKeyMoment Then
                                    datDragMoment = RefKeyMoment.ExactDateKeyMoment
                                    ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                                End If
                            End If
                        End If
                    End If
                    rActivity.X = GetCoordinateX(rCell.X, .StartDateMainActivity)
                    rActivity.Y = GetCoordinateY(rCell.Y, rCell.Height)
                    rActivity.Width = GetActivityWidth(.StartDateMainActivity, .EndDateMainActivity)
                    rActivity.Height = CONST_BarHeight
                End With

                intWidth = ptMouseLocation.X - rActivity.X + 1
                rDragRectangle = New Rectangle(rActivity.X, rActivity.Y, intWidth, CONST_BarHeight)
            Else
                boolDragActivityEnd = False
                DragActivity = Nothing
                rDragRectangle = Nothing
            End If
        End If
    End Sub

    Private Sub DropActivityEnd(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing Then
                With selActivity
                    .ActivityDetail = Drop_SetDuration(selActivity.ActivityDetail, .ActivityDetail.StartDateMainActivity, datDragMoment)
                End With
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub DragPreparationStart(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing And selActivity Is DragActivity Then
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset
                Dim rActivity As Rectangle = rCell
                Dim intWidth As Integer
                datDragMoment = CoordinateToDate(rCell.X, ptMouseLocation.X)

                With selActivity.ActivityDetail
                    If datDragMoment > .StartDateMainActivity Then
                        datDragMoment = .StartDateMainActivity
                        ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                    End If

                    rActivity.X = GetCoordinateX(rCell.X, .StartDatePreparation)
                    rActivity.Y = GetCoordinateY(rCell.Y, rCell.Height, True)
                    rActivity.Width = GetActivityWidth(.StartDatePreparation, .StartDateMainActivity)
                    rActivity.Height = CONST_BarHeight
                End With

                intWidth = rActivity.Right - ptMouseLocation.X + 1
                rDragRectangle = New Rectangle(ptMouseLocation.X, rActivity.Y, intWidth, CONST_PreparationHeight)
            Else
                boolDragPreparationStart = False
                DragActivity = Nothing
                rDragRectangle = Nothing
            End If
        End If
    End Sub

    Private Sub DropPreparationStart(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing Then
                With selActivity.ActivityDetail
                    Dim tsDuration As TimeSpan = .StartDateMainActivity.Subtract(datDragMoment)
                    Dim intPreparation As Integer = tsDuration.Days

                    If datDragMoment <> CurrentLogFrame.StartDate Then .PreparationFromStart = False

                    If intPreparation Mod 7 = 0 Then
                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "Preparation", selActivity.ActivityDetail.Preparation)
                        .Preparation = intPreparation / 7
                        UndoRedo.ValueChanged(.Preparation)

                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "PreparationUnit", selActivity.ActivityDetail.PreparationUnit)
                        .PreparationUnit = ActivityDetail.DurationUnits.Week
                        UndoRedo.OptionChanged(.PreparationUnit)
                    Else
                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "Preparation", selActivity.ActivityDetail.Preparation)
                        .Preparation = intPreparation
                        UndoRedo.ValueChanged(.Preparation)

                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "PreparationUnit", selActivity.ActivityDetail.PreparationUnit)
                        .PreparationUnit = ActivityDetail.DurationUnits.Day
                        UndoRedo.OptionChanged(.PreparationUnit)
                    End If
                End With
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub DragFollowUpEnd(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing And selActivity Is DragActivity Then
                Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, hit.RowIndex, False)
                rCell.X -= Me.HorizontalScrollingOffset
                Dim rActivity As Rectangle = rCell
                Dim intWidth As Integer
                datDragMoment = CoordinateToDate(rCell.X, ptMouseLocation.X)

                With selActivity.ActivityDetail
                    If datDragMoment < .EndDateMainActivity Then
                        datDragMoment = .EndDateMainActivity
                        ptMouseLocation.X = GetCoordinateX(rCell.X, datDragMoment)
                    End If

                    rActivity.X = GetCoordinateX(rCell.X, .EndDateMainActivity)
                    rActivity.Y = GetCoordinateY(rCell.Y, rCell.Height, True)
                    rActivity.Width = GetActivityWidth(.EndDateMainActivity, .EndDateFollowUp)
                    rActivity.Height = CONST_BarHeight
                End With

                intWidth = ptMouseLocation.X - rActivity.X + 1
                rDragRectangle = New Rectangle(rActivity.X, rActivity.Y, intWidth, CONST_PreparationHeight)
            Else
                boolDragFollowUpEnd = False
                DragActivity = Nothing
                rDragRectangle = Nothing
            End If
        End If
    End Sub

    Private Sub DropFollowUpEnd(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

            If selActivity IsNot Nothing Then
                With selActivity.ActivityDetail
                    Dim tsDuration As TimeSpan = datDragMoment.Subtract(.EndDateMainActivity)
                    Dim intFollowUp As Integer = tsDuration.Days

                    If datDragMoment <> CurrentLogFrame.EndDate Then .FollowUpUntilEnd = False

                    If intFollowUp Mod 7 = 0 Then
                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "FollowUp", selActivity.ActivityDetail.FollowUp)
                        .FollowUp = intFollowUp / 7
                        UndoRedo.ValueChanged(.FollowUp)

                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "FollowUpUnit", selActivity.ActivityDetail.FollowUpUnit)
                        .FollowUpUnit = ActivityDetail.DurationUnits.Week
                        UndoRedo.OptionChanged(.FollowUpUnit)
                    Else
                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "FollowUp", selActivity.ActivityDetail.FollowUp)
                        .FollowUp = intFollowUp
                        UndoRedo.ValueChanged(.FollowUp)

                        UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "FollowUpUnit", selActivity.ActivityDetail.FollowUpUnit)
                        .FollowUpUnit = ActivityDetail.DurationUnits.Day
                        UndoRedo.OptionChanged(.FollowUpUnit)
                    End If
                End With
            End If
        End If

        Me.Reload()
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As Object)
        Dim selGridRow As PlanningGridRow
        Dim intColIndex As Integer = 1
        Dim intRowIndex As Integer

        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.KeyMoment
                    If selGridRow.KeyMoment Is selItem Then
                        Exit For
                    End If
                Case PlanningGridRow.RowTypes.Activity
                    If selGridRow.Struct Is selItem Then
                        Exit For
                    End If
            End Select
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Public Function GetFirstItem() As Object
        For Each selGridRow As PlanningGridRow In Me.Grid
            If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                If String.IsNullOrEmpty(selGridRow.KeyMoment.Description) = False Then
                    Return selGridRow.KeyMoment
                End If
            ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
                If String.IsNullOrEmpty(selGridRow.Struct.RTF) = False Then
                    Return selGridRow.Struct
                End If
            End If
        Next

        Return Nothing
    End Function
#End Region

#Region "Add and insert items"
    Public Sub AddNewItem()
        If CurrentCell Is Nothing Then Exit Sub

        Dim intRowIndex As Integer
        Dim intColIndex = 1
        Dim strColName As String = Me.Columns(intColIndex).Name
        Dim selGridRow As PlanningGridRow

        For intRowIndex = CurrentCell.RowIndex To RowCount - 1
            selGridRow = Grid(intRowIndex)
            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.KeyMoment
                    If selGridRow.KeyMoment Is Nothing Then Exit For
                Case PlanningGridRow.RowTypes.Activity
                    If selGridRow.Struct Is Nothing Then Exit For
            End Select
        Next

        CurrentCell = Me(intColIndex, intRowIndex)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selObject As Object = Me.CurrentItem(False)
        Dim intIndex As Integer

        If selObject Is Nothing Then Exit Sub

        Select Case selObject.GetType
            Case GetType(KeyMoment)
                Dim selKeyMoment As KeyMoment = DirectCast(selObject, KeyMoment)
                If selKeyMoment.ParentOutputGuid <> Guid.Empty Then
                    Dim ParentOutput As Output = CurrentLogFrame.GetOutputByGuid(selKeyMoment.ParentOutputGuid)
                    If ParentOutput IsNot Nothing Then
                        intIndex = ParentOutput.KeyMoments.IndexOf(selKeyMoment)

                        Dim NewKeyMoment As New KeyMoment
                        ParentOutput.KeyMoments.Insert(intIndex, NewKeyMoment)
                        UndoRedo.ItemInserted(NewKeyMoment, ParentOutput.KeyMoments)
                        selObject = NewKeyMoment
                    End If
                End If
            Case GetType(Activity)
                Dim selActivity As Activity = DirectCast(selObject, Activity)
                Dim selActivities As Activities = CurrentLogFrame.GetParentCollection(selActivity)

                If selActivities IsNot Nothing Then
                    intIndex = selActivities.IndexOf(selActivity)

                    Dim NewActivity As New Activity
                    selActivities.Insert(intIndex, NewActivity)
                    UndoRedo.ItemInserted(NewActivity, selActivities)
                    selObject = NewActivity
                End If
        End Select

        Reload()

        SetFocusOnItem(selObject)

        Me.BeginEdit(False)
    End Sub

    Public Sub InsertParentItem()
        Dim selActivity As Activity
        Dim ParentActivities As Activities = Nothing

        selActivity = TryCast(Me.CurrentItem(False), Activity)

        If selActivity IsNot Nothing Then
            ParentActivities = CurrentLogFrame.GetParentCollection(selActivity)

            If ParentActivities IsNot Nothing Then
                Dim intIndex As Integer = ParentActivities.IndexOf(selActivity)
                Dim NewParent As New Activity

                ParentActivities.Remove(selActivity)

                ParentActivities.Insert(intIndex, NewParent)
                UndoRedo.ItemInserted(NewParent, ParentActivities)

                NewParent.Activities.AddToProcess(selActivity)
                UndoRedo.ItemParentInserted(selActivity, ParentActivities, intIndex, NewParent.Activities)

                Reload()
                SetFocusOnItem(NewParent)
                Me.BeginEdit(False)
            End If
        End If
    End Sub

    Public Sub InsertChildItem()
        Dim selActivity As Activity = TryCast(Me.CurrentItem(False), Activity)

        If selActivity IsNot Nothing Then
            Dim ChildActivity As New Activity

            selActivity.Activities.Insert(0, ChildActivity)
            UndoRedo.ItemChildInserted(ChildActivity, selActivity.Activities)

            Reload()
            SetFocusOnItem(ChildActivity)
            Me.BeginEdit(False)
        End If
    End Sub
#End Region

#Region "Move items"
    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selGridRow As PlanningGridRow = Grid(CurrentCell.RowIndex)
        Dim selItem As Object = Me.CurrentItem(False)

        If selItem Is Nothing Then Exit Sub

        If selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
            If intDirection < 0 Then
                selItem = MoveActivity_ToPreviousParent(selItem)
            Else
                selItem = MoveActivity_ToNextParent(selItem)
            End If
        End If

        Reload()
        SetFocusOnItem(selItem)
    End Sub

    Private Function MoveActivity_ToPreviousParent(ByVal selActivity As Activity) As Activity
        Dim objActivities As Activities = CurrentLogFrame.GetParentCollection(selActivity)
        Dim intOldIndex As Integer = objActivities.IndexOf(selActivity)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(intRowIndex)
        Dim intIndex As Integer

        If PreviousStruct Is Nothing Then Return selActivity

        Select Case PreviousStruct.GetType
            Case GetType(Activity)
                Dim objPreviousActivities As Activities = CurrentLogFrame.GetParentCollection(PreviousStruct)

                intIndex = objPreviousActivities.IndexOf(PreviousStruct)
                If intIndex = objPreviousActivities.Count - 1 Then intIndex += 1

                objActivities.Remove(selActivity)
                objPreviousActivities.Insert(intIndex, selActivity)

                UndoRedo.ItemMovedUp(selActivity, selActivity, objActivities, intOldIndex, objPreviousActivities)
            Case GetType(Output)
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToPreviousParent(selActivity)

            Case GetType(Purpose)
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToPreviousParent(selActivity)
        End Select

        Return selActivity
    End Function

    Private Function MoveActivity_ToNextParent(ByVal selActivity As Activity) As Activity
        Dim objActivities As Activities = CurrentLogFrame.GetParentCollection(selActivity)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As PlanningGridRow = Me.Grid(intRowIndex)
        Dim NextStruct As Struct = Me.Grid.GetNextStruct(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objActivities.IndexOf(selActivity)

        If NextStruct Is Nothing Then Return selActivity

        Select Case NextStruct.GetType
            Case GetType(Activity)
                If CurrentLogFrame.IsParentLineage(NextStruct, selActivity) Then
                    intRowIndex += 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    MoveActivity_ToNextParent(selActivity)
                Else
                    Dim objNextActivities As Activities
                    If CType(NextStruct, Activity).Activities.Count > 0 Then
                        'if the next activity has sub-activities, insert as first sub-activity
                        objNextActivities = CType(NextStruct, Activity).Activities
                        intIndex = 0
                    Else
                        'insert before the next activity (at the same level)
                        objNextActivities = CurrentLogFrame.GetParentCollection(NextStruct)
                        intIndex = objNextActivities.IndexOf(NextStruct)
                    End If
                    objActivities.Remove(selActivity)
                    objNextActivities.Insert(intIndex, selActivity)

                    UndoRedo.ItemMovedDown(selActivity, selActivity, objActivities, intOldIndex, objNextActivities)
                End If

            Case GetType(Output)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToNextParent(selActivity)

            Case GetType(Purpose)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToNextParent(selActivity)
        End Select

        Return selActivity
    End Function

    Public Sub ChangeSection(ByVal intDirection As Integer)
        If CurrentCell Is Nothing Then Exit Sub

        Dim selItem As Object = Nothing
        Dim selGridRow As PlanningGridRow = Grid(CurrentCell.RowIndex)


        Select Case selGridRow.RowType
            Case PlanningGridRow.RowTypes.KeyMoment
                selItem = ChangeSection_KeyMoment(intDirection, selGridRow.KeyMoment)
            Case PlanningGridRow.RowTypes.Activity
                Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
                If selActivity IsNot Nothing Then _
                    selItem = ChangeSection_Activity(intDirection, selGridRow.Struct)
        End Select

        Reload()
        If selItem IsNot Nothing Then SetFocusOnItem(selItem)
        CurrentRow.Selected = True
    End Sub

    Public Function ChangeSection_KeyMoment(ByVal intDirection As Integer, ByVal selKeyMoment As KeyMoment) As Object
        Dim intIndex As Integer, intParentIndex As Integer
        Dim selParentOutput As Output = CurrentLogFrame.GetOutputByGuid(selKeyMoment.ParentOutputGuid)
        Dim selParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selParentOutput.ParentPurposeGuid)

        intIndex = selParentOutput.KeyMoments.IndexOf(selKeyMoment)
        intParentIndex = selParentPurpose.Outputs.IndexOf(selParentOutput)

        If intDirection < 0 Then
            intParentIndex -= 1
            If intParentIndex >= 0 Then
                selParentOutput.KeyMoments.Remove(selKeyMoment)
                selParentOutput = selParentPurpose.Outputs(intParentIndex)
                If intIndex > selParentOutput.KeyMoments.Count Then intIndex = selParentOutput.KeyMoments.Count
                If intIndex < 0 Then intIndex = 0
                selParentOutput.KeyMoments.Insert(intIndex, selKeyMoment)
            Else
                Dim intPurposeIndex As Integer = CurrentLogFrame.Purposes.IndexOf(selParentPurpose)
                intPurposeIndex -= 1
                If intPurposeIndex >= 0 Then
                    selParentPurpose = CurrentLogFrame.Purposes(intPurposeIndex)
                    If selParentPurpose.Outputs.Count > 0 Then
                        selParentOutput.KeyMoments.Remove(selKeyMoment)
                        selParentOutput = selParentPurpose.Outputs(selParentPurpose.Outputs.Count - 1)
                        If intIndex > selParentOutput.KeyMoments.Count Then intIndex = selParentOutput.KeyMoments.Count
                        If intIndex < 0 Then intIndex = 0
                        selParentOutput.KeyMoments.Insert(intIndex, selKeyMoment)
                    End If
                End If
            End If
        Else
            intParentIndex += 1
            If intParentIndex <= selParentPurpose.Outputs.Count - 1 Then
                selParentOutput.KeyMoments.Remove(selKeyMoment)
                selParentOutput = selParentPurpose.Outputs(intParentIndex)
                If intIndex > selParentOutput.KeyMoments.Count Then intIndex = selParentOutput.KeyMoments.Count
                If intIndex < 0 Then intIndex = 0
                selParentOutput.KeyMoments.Insert(intIndex, selKeyMoment)
            Else
                Dim intPurposeIndex As Integer = CurrentLogFrame.Purposes.IndexOf(selParentPurpose)
                intPurposeIndex += 1
                If intPurposeIndex <= CurrentLogFrame.Purposes.Count - 1 Then
                    selParentPurpose = CurrentLogFrame.Purposes(intPurposeIndex)
                    If selParentPurpose.Outputs.Count > 0 Then
                        selParentOutput.KeyMoments.Remove(selKeyMoment)
                        selParentOutput = selParentPurpose.Outputs(selParentPurpose.Outputs.Count - 1)
                        If intIndex > selParentOutput.KeyMoments.Count Then intIndex = selParentOutput.KeyMoments.Count
                        If intIndex < 0 Then intIndex = 0
                        selParentOutput.KeyMoments.Insert(intIndex, selKeyMoment)
                    End If
                End If
            End If
        End If

        Return selKeyMoment
    End Function

    Public Function ChangeSection_Activity(ByVal intDirection As Integer, ByVal selActivity As Activity) As Object
        Dim intIndex As Integer, intParentIndex, intOldIndex As Integer
        Dim RootActivity As Activity = CurrentLogFrame.GetRootActivity(selActivity)
        Dim ParentActivities As Activities = CurrentLogFrame.GetParentCollection(selActivity)
        Dim selParentOutput As Output = CurrentLogFrame.GetOutputByGuid(RootActivity.ParentOutputGuid)
        Dim selParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selParentOutput.ParentPurposeGuid)

        intIndex = selParentOutput.Activities.IndexOf(RootActivity)
        intParentIndex = selParentPurpose.Outputs.IndexOf(selParentOutput)
        intOldIndex = ParentActivities.IndexOf(selActivity)

        If intDirection < 0 Then
            intParentIndex -= 1
            If intParentIndex >= 0 Then
                ParentActivities.Remove(selActivity)
                selParentOutput = selParentPurpose.Outputs(intParentIndex)

                If intIndex > selParentOutput.Activities.Count Then intIndex = selParentOutput.Activities.Count
                If intIndex < 0 Then intIndex = 0

                selParentOutput.Activities.Insert(intIndex, selActivity)
                UndoRedo.ItemSectionUp(selActivity, selActivity, ParentActivities, intOldIndex, selParentOutput.Activities)
            Else
                Dim intPurposeIndex As Integer = CurrentLogFrame.Purposes.IndexOf(selParentPurpose)
                intPurposeIndex -= 1
                If intPurposeIndex >= 0 Then
                    selParentPurpose = CurrentLogFrame.Purposes(intPurposeIndex)
                    If selParentPurpose.Outputs.Count > 0 Then
                        ParentActivities.Remove(selActivity)
                        selParentOutput = selParentPurpose.Outputs(selParentPurpose.Outputs.Count - 1)

                        If intIndex > selParentOutput.Activities.Count Then intIndex = selParentOutput.Activities.Count
                        If intIndex < 0 Then intIndex = 0

                        selParentOutput.Activities.Insert(intIndex, selActivity)
                        UndoRedo.ItemSectionUp(selActivity, selActivity, ParentActivities, intOldIndex, selParentOutput.Activities)
                    End If
                End If
            End If
        Else
            intParentIndex += 1
            If intParentIndex <= selParentPurpose.Outputs.Count - 1 Then
                ParentActivities.Remove(selActivity)
                selParentOutput = selParentPurpose.Outputs(intParentIndex)

                If intIndex > selParentOutput.Activities.Count Then intIndex = selParentOutput.Activities.Count
                If intIndex < 0 Then intIndex = 0

                selParentOutput.Activities.Insert(intIndex, selActivity)
                UndoRedo.ItemSectionDown(selActivity, selActivity, ParentActivities, intOldIndex, selParentOutput.Activities)
            Else
                Dim intPurposeIndex As Integer = CurrentLogFrame.Purposes.IndexOf(selParentPurpose)
                intPurposeIndex += 1
                If intPurposeIndex <= CurrentLogFrame.Purposes.Count - 1 Then
                    selParentPurpose = CurrentLogFrame.Purposes(intPurposeIndex)
                    If selParentPurpose.Outputs.Count > 0 Then
                        ParentActivities.Remove(selActivity)
                        selParentOutput = selParentPurpose.Outputs(selParentPurpose.Outputs.Count - 1)

                        If intIndex > selParentOutput.Activities.Count Then intIndex = selParentOutput.Activities.Count
                        If intIndex < 0 Then intIndex = 0

                        selParentOutput.Activities.Insert(intIndex, selActivity)
                        UndoRedo.ItemSectionDown(selActivity, selActivity, ParentActivities, intOldIndex, selParentOutput.Activities)
                    End If
                End If
            End If
        End If

        Return selActivity
    End Function

    Public Sub LevelUp()
        Dim selActivity As Activity = TryCast(Me.CurrentItem(False), Activity)
        If selActivity Is Nothing Then Exit Sub

        Dim ParentActivity As Activity = TryCast(CurrentLogFrame.GetParent(selActivity), Activity)
        Dim intOldIndex As Integer = ParentActivity.Activities.IndexOf(selActivity)

        If ParentActivity Is Nothing Then Exit Sub

        Dim objParentActivities As Activities = CurrentLogFrame.GetParentCollection(ParentActivity)
        Dim intIndex As Integer = objParentActivities.IndexOf(ParentActivity)

        intIndex += 1
        ParentActivity.Activities.Remove(selActivity)
        objParentActivities.Insert(intIndex, selActivity)

        UndoRedo.ItemLevelUp(selActivity, ParentActivity.Activities, intOldIndex, objParentActivities)

        Reload()
        SetFocusOnItem(selActivity)
    End Sub

    Public Sub LevelDown()
        Dim selActivity As Activity = TryCast(Me.CurrentItem(False), Activity)
        If selActivity Is Nothing Then Exit Sub

        Dim objActivities As Activities = CurrentLogFrame.GetParentCollection(selActivity)
        Dim intIndex As Integer = objActivities.IndexOf(selActivity)
        Dim intOldIndex As Integer = intIndex

        If intIndex = 0 Then Exit Sub

        intIndex -= 1
        objActivities.Remove(selActivity)
        Dim PreviousActivity As Activity = objActivities(intIndex)
        PreviousActivity.Activities.AddToProcess(selActivity)

        UndoRedo.ItemLevelDown(selActivity, objActivities, intOldIndex, PreviousActivity.Activities)

        Reload()
        SetFocusOnItem(selActivity)
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim strSourceColName As String
        Dim boolShift As Boolean
        Dim boolRemoveAll As Boolean
        Dim intRowIndex, intColumnIndex As Integer
        Dim strSortNumber As String = ""
        Dim objParentOutput As Output
        Dim objActivity As Activity
        Dim objActivities As Activities

        If Me.IsCurrentCellInEditMode = False Then
            intRowIndex = CurrentCell.RowIndex
            intColumnIndex = CurrentCell.ColumnIndex
            If Control.ModifierKeys = Keys.Shift Then boolShift = True

            'copy cells to delete
            strSourceColName = Columns(SelectionRectangle.FirstColumnIndex).Name
            For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
                SelectedGridRows.Add(Me.Grid(i))
            Next

            If ShowWarning = True Then
                Dim boolCancel As Boolean = RemoveItems_Warning(strSourceColName)
                If boolCancel = True Then Exit Sub
            End If

            For Each selGridRow As PlanningGridRow In SelectedGridRows
                strSortNumber = selGridRow.SortNumber

                If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment And selGridRow.KeyMoment IsNot Nothing Then
                    objParentOutput = CurrentLogFrame.GetOutputByGuid(selGridRow.KeyMoment.ParentOutputGuid)
                    If objParentOutput IsNot Nothing AndAlso objParentOutput.KeyMoments.Contains(selGridRow.KeyMoment) Then
                        If boolCut = False Then
                            UndoRedo.ItemRemoved(selGridRow.KeyMoment, objParentOutput.KeyMoments)
                        Else
                            UndoRedo.ItemCut(selGridRow.KeyMoment, objParentOutput.KeyMoments)
                        End If

                        RemoveItems_Referers(selGridRow.KeyMoment.Guid)
                        objParentOutput.KeyMoments.Remove(selGridRow.KeyMoment)
                    End If

                ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity And selGridRow.Struct IsNot Nothing Then
                    objActivity = DirectCast(selGridRow.Struct, Activity)
                    objActivities = CurrentLogFrame.GetParentCollection(objActivity)

                    If objActivities IsNot Nothing AndAlso objActivities.Contains(objActivity) Then
                        If objActivity.Activities.Count > 0 Then
                            If boolShift = True Then boolRemoveAll = True Else boolRemoveAll = False
                        Else
                            boolRemoveAll = True
                        End If

                        If boolRemoveAll = True Then
                            If boolCut = False Then
                                UndoRedo.ItemRemoved(objActivity, objActivities)
                            Else
                                UndoRedo.ItemCut(objActivity, objActivities)
                            End If

                            RemoveItems_Referers(objActivity.Guid)
                            objActivities.Remove(objActivity)
                        Else
                            If boolCut = False Then
                                UndoRedo.ItemRemovedNotVertical(objActivity, objActivities)
                            Else
                                UndoRedo.ItemCutNotVertical(objActivity, objActivities)
                            End If

                            With objActivity
                                .RTF = String.Empty
                                .Indicators.Clear()
                                .Resources.Clear()
                                .Assumptions.Clear()
                            End With
                        End If
                    End If
                End If
            Next

            SelectedGridRows.Clear()

            ClearSelection()
            CurrentCell = Me(intColumnIndex, intRowIndex)
            Me.Reload()
        Else
            If CurrentEditingControl IsNot Nothing Then
                With CurrentEditingControl
                    Dim intSelectionStart As Integer = .SelectionStart
                    Dim intLength As Integer = .SelectionLength

                    If intLength = 0 Then intLength = 1
                    .Text = .Text.Remove(intSelectionStart, intLength)

                    .Select(intSelectionStart, 0)
                End With
            End If
        End If
    End Sub

    Private Sub RemoveItems_Referers(ByVal objReferenceGuid As Guid)
        Dim ReferersList As ArrayList = CurrentLogFrame.GetReferingMomentsByReferenceGuid(objReferenceGuid)

        If ReferersList.Count > 0 Then
            For i = 0 To ReferersList.Count - 1
                Select Case ReferersList(i).GetType
                    Case GetType(KeyMoment)
                        Dim selKeyMoment As KeyMoment = ReferersList(i)

                        With selKeyMoment
                            .KeyMoment = .ExactDateKeyMoment
                            .Relative = False
                            .Period = 0
                            .PeriodDirection = KeyMoment.DirectionValues.NotDefined
                            .PeriodUnit = DurationUnits.NotDefined
                            .GuidReferenceMoment = Guid.Empty
                        End With
                    Case GetType(Activity)
                        Dim selActivity As Activity = ReferersList(i)

                        With selActivity.ActivityDetail
                            .StartDate = .StartDateMainActivity
                            .Relative = False
                            .Period = 0
                            .PeriodDirection = ActivityDetail.DirectionValues.NotDefined
                            .PeriodUnit = ActivityDetail.DurationUnits.NotDefined
                            .GuidReferenceMoment = Guid.Empty
                        End With
                End Select
            Next
        End If
    End Sub

    Private Function RemoveItems_Warning(ByVal strSourceColName As String) As Boolean
        Dim intNrActivities As Integer, intNrKeyMoments As Integer
        Dim intNrInd As Integer, intNrRsc As Integer, intNrAsm As Integer
        Dim strMsg As String, strTitle As String = String.Empty
        Dim strMsgKeyMoment As String = String.Empty, strMsgActivity As String = String.Empty

        For Each selGridRow As PlanningGridRow In SelectedGridRows
            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.Activity
                    If selGridRow.Struct IsNot Nothing Then
                        Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
                        intNrInd += selActivity.Indicators.Count
                        intNrRsc += selActivity.Resources.Count
                        intNrAsm += selActivity.Assumptions.Count
                        intNrActivities += 1
                    End If
                Case PlanningGridRow.RowTypes.KeyMoment
                    If selGridRow.KeyMoment IsNot Nothing Then
                        intNrKeyMoments += 1
                    End If
            End Select
        Next


        If intNrKeyMoments > 0 Then
            strMsgKeyMoment = String.Format("{0} {1}", intNrKeyMoments, LANG_KeyMoments.ToLower) & vbLf & vbLf
        End If
        If intNrActivities > 0 Then
            strMsgActivity = String.Format("{0} {1}", intNrActivities, CurrentLogFrame.StructNamePlural(3)).ToLower

            If intNrInd > 0 Or intNrRsc > 0 Or intNrAsm > 0 Then
                strMsgActivity &= String.Format(" {0}:{1}", LANG_With, vbLf)

                If intNrInd > 0 Then
                    strMsgActivity &= String.Format("   - {0} {1}", intNrInd, LANG_Indicators.ToLower) & vbLf
                End If
                If intNrRsc > 0 Then
                    strMsgActivity &= String.Format("   - {0} {1}", intNrRsc, LANG_Resources.ToLower) & vbLf
                End If
                If intNrAsm > 0 Then
                    strMsgActivity &= String.Format("   - {0} {1}", intNrAsm, LANG_Assumptions.ToLower) & vbLf
                End If
            End If

            strMsgActivity &= vbLf
        End If

        strTitle = LANG_Remove
        strMsg = String.Format(LANG_RemovePlanningItems, strMsgKeyMoment, strMsgActivity)

        Dim wdDeleteChildren As New DialogWarning(strMsg, strTitle)
        wdDeleteChildren.Type = DialogWarning.WarningDialogTypes.wdDeleteChildren
        wdDeleteChildren.ShowDialog()

        If wdDeleteChildren.DialogResult = Windows.Forms.DialogResult.No Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region 'remove items

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)
        CopyItems()

        RemoveItems(ShowWarning, True)
    End Sub

    Public Overrides Sub CopyItems()
        With SelectionRectangle
            Dim selRow As PlanningGridRow
            Dim strSort As String
            Dim CopyGroup As Date = Now()

            With SelectionRectangle
                For i = .FirstRowIndex To .LastRowIndex
                    selRow = Me.Grid(i)
                    If selRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                        If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.KeyMoment Then
                            strSort = KeyMoment.ItemName & " " & selRow.SortNumber
                            Dim NewItem As New ClipboardItem(CopyGroup, selRow.KeyMoment, strSort)
                            ItemClipboard.Insert(0, NewItem)
                        End If
                    ElseIf selRow.RowType = PlanningGridRow.RowTypes.Activity Then
                        If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Struct Then
                            strSort = Activity.ItemName & " " & selRow.SortNumber
                            Dim NewItem As New ClipboardItem(CopyGroup, selRow.Struct, strSort)
                            ItemClipboard.Insert(0, NewItem)
                        End If
                    End If
                Next
            End With
        End With

    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
        If PasteCell Is Nothing Then PasteCell = CurrentCell
        If PasteCell Is Nothing Then Exit Sub

        Dim intColumnIndex As Integer = PasteCell.ColumnIndex
        Dim intRowIndex As Integer = PasteCell.RowIndex
        Dim selItem As ClipboardItem

        PasteRow = Grid(intRowIndex)

        For i = 0 To PasteItems.Count - 1 'To 0 Step -1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(KeyMoment)
                    PasteItems_KeyMoment(selItem, intPasteOption)
                Case GetType(Activity)
                    PasteItems_Activity(selItem, intPasteOption)
                Case Else
                    PasteItems_Other(selItem, intPasteOption)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(intColumnIndex, intRowIndex)
    End Sub

    Private Sub PasteActivity(ByVal NewActivity As Activity)
        Dim ParentActivities As Activities
        Dim PasteRowActivity As Activity = TryCast(PasteRow.Struct, Activity)
        Dim intIndex As Integer
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRowActivity IsNot Nothing Then
            ParentActivities = CurrentLogFrame.GetParentCollection(PasteRowActivity)
            intIndex = ParentActivities.IndexOf(PasteRowActivity)
        Else
            Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))

            Select Case PreviousStruct.GetType
                Case GetType(Output)
                    ParentActivities = CType(PreviousStruct, Output).Activities
                Case GetType(Activity)
                    ParentActivities = CurrentLogFrame.GetParentCollection(PreviousStruct)
                Case Else
                    Exit Sub
            End Select
            intIndex = ParentActivities.Count
        End If
        ParentActivities.Insert(intIndex, NewActivity)

        PasteRow.Struct = NewActivity

        UndoRedo.ItemPasted(NewActivity, ParentActivities)
    End Sub

    Private Sub PasteKeyMoment(ByVal NewKeyMoment As KeyMoment)
        Dim ParentOutput As Output
        Dim PasteRowKeyMoment As KeyMoment = PasteRow.KeyMoment
        Dim strSortNumber As String = PasteRow.SortNumber
        Dim intIndex As Integer

        If PasteRowKeyMoment IsNot Nothing Then
            ParentOutput = CurrentLogFrame.GetOutputByGuid(PasteRowKeyMoment.ParentOutputGuid)
            intIndex = ParentOutput.KeyMoments.IndexOf(PasteRowKeyMoment)
        Else
            Dim PreviousKeyMoment As KeyMoment = Me.Grid.GetPreviousKeyMomentOfOutput(Grid.IndexOf(PasteRow))
            If PreviousKeyMoment IsNot Nothing Then
                ParentOutput = CurrentLogFrame.GetOutputByGuid(PreviousKeyMoment.ParentOutputGuid)
            Else
                ParentOutput = Grid.GetPreviousOutput(Grid.IndexOf(PasteRow))
            End If
            intIndex = ParentOutput.KeyMoments.Count
        End If

        ParentOutput.KeyMoments.Insert(intIndex, NewKeyMoment)
        PasteRow.KeyMoment = NewKeyMoment

        UndoRedo.ItemPasted(NewKeyMoment, ParentOutput.KeyMoments)
    End Sub

    Private Sub PasteItems_Activity(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selActivity As Activity = DirectCast(selItem.Item, Activity)

        If PasteRow.RowType = PlanningGridRow.RowTypes.Activity Then
            Dim NewActivity As New Activity

            Using copier As New ObjectCopy
                NewActivity = copier.CopyObject(selActivity)
            End Using

            PasteActivity(NewActivity)
        ElseIf PasteRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            Dim NewKeyMoment As New KeyMoment()
            NewKeyMoment.Description = selActivity.Text
            If selActivity.ActivityDetail.Relative = True Then
                With NewKeyMoment
                    .Relative = True
                    .GuidReferenceMoment = selActivity.ActivityDetail.GuidReferenceMoment
                    .Period = selActivity.ActivityDetail.Period
                    .PeriodDirection = selActivity.ActivityDetail.PeriodDirection
                    .PeriodUnit = selActivity.ActivityDetail.PeriodUnit
                End With
            Else
                NewKeyMoment.KeyMoment = selActivity.ActivityDetail.StartDateMainActivity
            End If

            PasteKeyMoment(NewKeyMoment)
        End If
    End Sub

    Private Sub PasteItems_KeyMoment(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selKeyMoment As KeyMoment = DirectCast(selItem.Item, KeyMoment)
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            Dim NewKeyMoment As New KeyMoment

            Using copier As New ObjectCopy
                NewKeyMoment = copier.CopyObject(selKeyMoment)
            End Using

            PasteKeyMoment(NewKeyMoment)
        ElseIf PasteRow.RowType = PlanningGridRow.RowTypes.Activity Then
            Dim NewActivity As New Activity()
            NewActivity.SetText(selKeyMoment.Description)
            If selKeyMoment.Relative = True Then
                With NewActivity.ActivityDetail
                    .Relative = True
                    .Duration = 1
                    .DurationUnit = ActivityDetail.DurationUnits.Day
                    .GuidReferenceMoment = selKeyMoment.GuidReferenceMoment
                    .Period = selKeyMoment.Period
                    .PeriodDirection = selKeyMoment.PeriodDirection
                    .PeriodUnit = selKeyMoment.PeriodUnit
                End With
            Else
                NewActivity.ActivityDetail.StartDate = selKeyMoment.KeyMoment
            End If

            PasteActivity(NewActivity)
        End If
    End Sub

    Private Sub PasteItems_Other(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selObject As LogframeObject = TryCast(selItem.Item, LogframeObject)
        Dim strText As String = String.Empty, strRtf As String = String.Empty

        If selObject IsNot Nothing Then
            strText = selObject.Text
            strRtf = selObject.RTF
        Else
            strText = selItem.ToString
        End If
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRow.RowType = PlanningGridRow.RowTypes.Activity Then
            Dim NewActivity As New Activity

            If String.IsNullOrEmpty(strRtf) Then
                NewActivity.SetText(strText)
            Else
                NewActivity.RTF = strRtf
            End If

            With NewActivity.ActivityDetail
                .Relative = False
                .StartDate = Now.Date
                .Duration = 1
                .DurationUnit = ActivityDetail.DurationUnits.Day
            End With

            PasteActivity(NewActivity)
        ElseIf PasteRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            Dim NewKeyMoment As New KeyMoment(Now.Date, strText)

            PasteKeyMoment(NewKeyMoment)
        End If
    End Sub
#End Region

#Region "Linking activities / key moments"
    Private Sub LinkFromKeyMoment(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, LinkRowIndex, False)

            With LinkKeyMoment
                ptLinkEnd = ptMouseLocation

                If Me.Rows(LinkRowIndex).Displayed = True Then
                    ptLinkStart = New Point()
                    ptLinkStart.X = GetCoordinateX(rCell.X, .ExactDateKeyMoment)
                    ptLinkStart.Y = GetCoordinateY(rCell.Y, rCell.Height)
                End If
            End With
        End If
    End Sub

    Private Sub LinkFromKeyMoment_Connect(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            If boolLinkHoverOverKeyMoment = True Then
                Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment

                If selKeyMoment IsNot Nothing Then
                    With selKeyMoment
                        .Relative = True
                        .GuidReferenceMoment = LinkKeyMoment.Guid
                        .Period = 1
                        .PeriodUnit = DurationUnits.Day
                        .PeriodDirection = KeyMoment.DirectionValues.After
                    End With
                End If
            ElseIf boolLinkHoverOverActivity = True Then
                Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

                If selActivity IsNot Nothing Then
                    With selActivity.ActivityDetail
                        .Relative = True
                        .GuidReferenceMoment = LinkKeyMoment.Guid
                        .Period = 1
                        .PeriodUnit = ActivityDetail.DurationUnits.Day
                        .PeriodDirection = ActivityDetail.DirectionValues.After
                    End With
                End If
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub LinkFromActivityStart(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, LinkRowIndex, False)

            With LinkActivity.ActivityDetail
                ptLinkEnd = ptMouseLocation

                If Me.Rows(LinkRowIndex).Displayed = True Then
                    ptLinkStart = New Point()
                    ptLinkStart.X = GetCoordinateX(rCell.X, .StartDateMainActivity)
                    ptLinkStart.Y = GetCoordinateY(rCell.Y, rCell.Height)
                End If
            End With
        End If
    End Sub

    Private Sub LinkFromActivityStart_Connect(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)

            If boolLinkHoverOverKeyMoment = True Then
                'Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment

                'If selKeyMoment IsNot Nothing Then
                '    With selKeyMoment
                '        .Relative = True
                '        .GuidReferenceMoment = LinkActivity.Guid
                '        .Period = 1
                '        .PeriodUnit = DurationUnits.Day
                '        .PeriodDirection = KeyMoment.DirectionValues.Before
                '    End With
                'End If
            ElseIf boolLinkHoverOverActivity = True Then
                Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

                If selActivity IsNot Nothing Then
                    With selActivity.ActivityDetail
                        .DurationUntilEnd = False
                        .Relative = True
                        .GuidReferenceMoment = LinkActivity.Guid
                        .Period = .Duration
                        .PeriodUnit = .DurationUnit
                        .PeriodDirection = ActivityDetail.DirectionValues.Before
                    End With
                End If
            End If
        End If
        Me.Reload()
    End Sub

    Private Sub LinkFromActivityEnd(ByVal ptMouseLocation)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim rCell As Rectangle = Me.GetCellDisplayRectangle(CONST_PlanningColumnIndex, LinkRowIndex, False)

            With LinkActivity.ActivityDetail
                ptLinkEnd = ptMouseLocation

                If Me.Rows(LinkRowIndex).Displayed = True Then
                    ptLinkStart = New Point()
                    ptLinkStart.X = GetCoordinateX(rCell.X, .EndDateMainActivity)
                    ptLinkStart.Y = GetCoordinateY(rCell.Y, rCell.Height)
                End If
            End With
        End If
    End Sub

    Private Sub LinkFromActivityEnd_Connect(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            If boolLinkHoverOverKeyMoment = True Then
                Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment

                If selKeyMoment IsNot Nothing Then
                    With selKeyMoment
                        .Relative = True
                        .GuidReferenceMoment = LinkActivity.Guid
                        .Period = 1
                        .PeriodUnit = DurationUnits.Day
                        .PeriodDirection = KeyMoment.DirectionValues.After
                    End With
                End If
            ElseIf boolLinkHoverOverActivity = True Then
                Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

                If selActivity IsNot Nothing Then
                    With selActivity.ActivityDetail
                        .Relative = True
                        .GuidReferenceMoment = LinkActivity.Guid
                        .Period = 1
                        .PeriodUnit = ActivityDetail.DurationUnits.Day
                        .PeriodDirection = ActivityDetail.DirectionValues.After
                    End With
                End If
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub UnlinkClicked(ByVal ptMouseLocation As Point)
        Dim hit As DataGridView.HitTestInfo = HitTest(ptMouseLocation.X, ptMouseLocation.Y)

        If hit.ColumnIndex = CONST_PlanningColumnIndex Then
            Dim selGridRow As PlanningGridRow = Grid(hit.RowIndex)
            If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment

                If selKeyMoment IsNot Nothing AndAlso selKeyMoment.Relative = True Then
                    With selKeyMoment
                        .KeyMoment = .ExactDateKeyMoment
                        .GuidReferenceMoment = Guid.Empty
                        .Period = 0
                        .PeriodUnit = DurationUnits.NotDefined
                        .PeriodDirection = KeyMoment.DirectionValues.NotDefined
                        .Relative = False
                    End With
                End If
            ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
                Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

                If selActivity IsNot Nothing AndAlso selActivity.ActivityDetail.Relative = True Then
                    With selActivity.ActivityDetail
                        .StartDate = .StartDateMainActivity
                        .GuidReferenceMoment = Guid.Empty
                        .Period = 0
                        .PeriodUnit = ActivityDetail.DurationUnits.NotDefined
                        .PeriodDirection = ActivityDetail.DirectionValues.NotDefined
                        .Relative = False
                    End With
                End If
            End If
        End If

        Me.Reload()
    End Sub

    Private Sub MoveCurrentCell()
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim CurrentGridRow As PlanningGridRow = Me.Grid(CurrentCell.RowIndex)
        Dim strColumnName As String = Me.Columns(intColumnIndex).Name
        Dim objPlanningObject As Object = Nothing

        If CurrentGridRow.RowType <> PlanningGridRow.RowTypes.Activity And CurrentGridRow.RowType <> PlanningGridRow.RowTypes.KeyMoment Then
            intRowIndex += 1
            Me.CurrentCell = Me(intColumnIndex, intRowIndex)
            MoveCurrentCell()
        End If

        If CurrentGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            If CurrentGridRow.KeyMoment IsNot Nothing Then objPlanningObject = CurrentGridRow.KeyMoment
        ElseIf CurrentGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
            Dim selActivity As Activity = TryCast(CurrentGridRow.Struct, Activity)
            If selActivity IsNot Nothing Then objPlanningObject = selActivity
        End If

        Select Case strColumnName
            Case "SortNumber"
                intColumnIndex += 1
                Me.CurrentCell = Me(intColumnIndex, intRowIndex)
                MoveCurrentCell()
            Case "RTF"
                Me.CurrentCell = Me(intColumnIndex, intRowIndex)

                If ClickPoint.IsEmpty = False Then
                    Me.BeginEdit(False)

                    Dim rCell As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex, intRowIndex, False)
                    ClickPoint.X -= rCell.X
                    ClickPoint.Y -= rCell.Y

                    If RichTextEditingControl IsNot Nothing Then
                        With RichTextEditingControl
                            Dim intCharIndex As Integer = .GetCharIndexFromPosition(ClickPoint)
                            .Select(intCharIndex, 0)
                            .SetCurrentText()

                            ClickPoint = Nothing
                        End With
                    End If
                End If
        End Select

        RaiseEvent PlanningObjectSelected(Me, New PlanningObjectSelectedEventArgs(objPlanningObject))
    End Sub

    Protected Overrides Sub OnColumnWidthChanged(ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
        If Me.IsCurrentCellInEditMode = False Then
            MyBase.OnColumnWidthChanged(e)

            boolColumnWidthChanged = True
        End If
    End Sub

    Protected Overrides Sub OnScroll(ByVal e As System.Windows.Forms.ScrollEventArgs)
        MyBase.OnScroll(e)
        If Me.ColumnCount >= CONST_PlanningColumnIndex + 1 Then Me.InvalidateColumn(CONST_PlanningColumnIndex)
        InvalidateSelectionRectangle()
    End Sub
#End Region

#Region "Select text"
    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex
        Me.ReloadImages()
    End Sub

    Public Sub HighlightTextInCell(ByVal intMatchIndex As Integer, ByVal intMatchLength As Integer)
        If CurrentItem(False) Is Nothing Then Exit Sub

        If intMatchLength > 0 Then
            If CurrentCell.IsInEditMode = True Then EndEdit()

            BeginEdit(False)
            Dim ctl As RichTextEditingControlLogframe = EditingControl
            If ctl IsNot Nothing Then ctl.Select(intMatchIndex, intMatchLength)
        End If
    End Sub
#End Region
End Class
