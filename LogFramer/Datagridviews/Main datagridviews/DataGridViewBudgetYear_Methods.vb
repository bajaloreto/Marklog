Partial Public Class DataGridViewBudgetYear

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
            Dim selGridRow As BudgetGridRow = Me.Grid(hit.RowIndex)

            If e.Button = MouseButtons.Right Then
                If hit.Type = DataGridViewHitTestType.Cell Then
                    ' Create a rectangle using the DragSize, with the mouse position being at the center of the rectangle.
                    OnMouseDown_SetDragRectangle(e.Location, selCell)
                End If
            Else
                If selCell.ColumnIndex = 1 Then
                    'determine where exactly in the text the user clicked
                    OnMouseDown_SetClickPoint(e.Location)
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

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim ptMouseLocation As New Point(e.X, e.Y)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If ((e.Button And MouseButtons.Right) = MouseButtons.Right) Then
            OnMouseMove_DragItems(e.Location)
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
        End If

        DragBoxFromMouseDown = Rectangle.Empty
        DragBudgetItem = Nothing

        MyBase.OnMouseUp(e)
        Invalidate()
    End Sub

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

    Private Sub MoveCurrentCell()
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim strColumnName As String = Me.Columns(intColumnIndex).Name

        If intRowIndex = IndexTotal Then
            intRowIndex -= 1
            If intRowIndex < 0 Then intRowIndex = 0
            Me.CurrentCell = Me(intColumnIndex, intRowIndex)
            If intRowIndex < Me.RowCount - 1 Then MoveCurrentCell()
        End If

        Dim CurrentGridRow As BudgetGridRow = Me.Grid(CurrentCell.RowIndex)
        If CurrentGridRow Is Nothing Then Exit Sub

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
            Case Else
                If CurrentGridRow.Type = BudgetItem.BudgetItemTypes.Ratio Then
                    Me.CurrentCell = Me(1, intRowIndex)
                End If
        End Select

        Dim selBudgetItem As BudgetItem = CurrentGridRow.BudgetItem
        RaiseEvent BudgetItemSelected(Me, New BudgetItemSelectedEventArgs(selBudgetItem))
    End Sub

    Protected Overrides Sub OnScroll(ByVal e As System.Windows.Forms.ScrollEventArgs)
        MyBase.OnScroll(e)

        InvalidateSelectionRectangle()
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As BudgetItem, Optional ByVal strPropertyName As String = "")
        Dim selGridRow As BudgetGridRow
        Dim intColIndex As Integer = 1
        Dim intRowIndex As Integer

        Select Case strPropertyName
            Case String.Empty
                intColIndex = 1
            Case "Duration"
                intColIndex = 2
            Case "DurationUnit"
                intColIndex = 3
            Case "Number"
                intColIndex = 4
            Case "NumberUnit"
                intColIndex = 5
            Case "UnitCost"
                intColIndex = 6
            Case "CurrencyCode"
                intColIndex = 7
            Case "TotalLocalCost"
                intColIndex = 8
            Case "TotalCost"
                intColIndex = 9
        End Select

        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.BudgetItem Is selItem Then
                Exit For
            End If
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            If Columns(intColIndex).Visible = False Then
                For i = intColIndex To 0 Step -1
                    If Columns(i).Visible = True Then
                        intColIndex = i
                        Exit For
                    End If
                Next
            End If
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Public Function GetFirstItem() As Object
        For Each selGridRow As BudgetGridRow In Me.Grid
            If selGridRow.BudgetItem IsNot Nothing AndAlso String.IsNullOrEmpty(selGridRow.BudgetItem.Text) = False Then
                Return selGridRow.BudgetItem
            End If
        Next

        Return Nothing
    End Function

    Public Function GetBudgetItemReferences() As Dictionary(Of Guid, String)
        Dim objBudgetItemReferences As New Dictionary(Of Guid, String)
        Dim selGridRow As BudgetGridRow
        Dim intRowNumber As Integer

        objBudgetItemReferences.Add(Guid.Empty, String.Empty)
        If Grid.Count > 0 Then
            For i = 0 To Grid.Count - 1
                selGridRow = Me.Grid(i)
                intRowNumber = i + 1
                If selGridRow.BudgetItem IsNot Nothing Then _
                    objBudgetItemReferences.Add(selGridRow.BudgetItem.Guid, intRowNumber.ToString)
            Next
        End If
        Return objBudgetItemReferences
    End Function

    Private Sub GetSelectedGridRows()
        SelectedGridRows.Clear()
        For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
            SelectedGridRows.Add(Me.Grid(i))
        Next
    End Sub

    Public Sub UpdateExchangeRates()
        Dim lstCurrencyCodes As List(Of IdValuePair) = CurrentLogFrame.Budget.LoadUsedCurrencyCodesList()
        With colUnitCostCurrency
            .DataSource = lstCurrencyCodes
            .ValueMember = "Id"
            .DisplayMember = "Value"
        End With

    End Sub
#End Region

#Region "Add and insert items"
    Public Sub AddNewItem()
        If CurrentCell Is Nothing Then Exit Sub

        Dim intRowIndex As Integer
        Dim intColIndex = 1
        Dim strColName As String = Me.Columns(intColIndex).Name
        Dim selGridRow As BudgetGridRow

        For intRowIndex = CurrentCell.RowIndex To RowCount - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.BudgetItem Is Nothing Then Exit For
        Next

        CurrentCell = Me(intColIndex, intRowIndex)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selBudgetItem As BudgetItem = Me.CurrentItem(False)
        Dim intIndex As Integer

        If selBudgetItem Is Nothing Then Exit Sub


        Dim selBudgetItems As BudgetItems = CurrentLogFrame.GetParentCollection(selBudgetItem)

        If selBudgetItems IsNot Nothing Then
            intIndex = selBudgetItems.IndexOf(selBudgetItem)

            Dim NewBudgetItem As New BudgetItem
            selBudgetItems.Insert(intIndex, NewBudgetItem)
            UndoRedo.ItemInserted(NewBudgetItem, selBudgetItems)

            selBudgetItem = NewBudgetItem
        End If

        Reload()

        SetFocusOnItem(selBudgetItem)

        Me.BeginEdit(False)
    End Sub

    Public Sub InsertParentItem()
        Dim selBudgetItem As BudgetItem = Me.CurrentItem(False)
        Dim ParentBudgetItems As BudgetItems = Nothing

        If selBudgetItem IsNot Nothing Then
            ParentBudgetItems = CurrentLogFrame.GetParentCollection(selBudgetItem)

            If ParentBudgetItems IsNot Nothing Then
                Dim intIndex As Integer = ParentBudgetItems.IndexOf(selBudgetItem)
                Dim NewParent As New BudgetItem

                ParentBudgetItems.Remove(selBudgetItem)

                ParentBudgetItems.Insert(intIndex, NewParent)
                UndoRedo.ItemInserted(NewParent, ParentBudgetItems)

                NewParent.BudgetItems.Add(selBudgetItem)
                UndoRedo.ItemParentInserted(selBudgetItem, ParentBudgetItems, intIndex, NewParent.BudgetItems)

                Reload()
                SetFocusOnItem(NewParent)

                If Me.TotalBudget = True Then RaiseEvent BudgetItemInsertParent(Me, New BudgetItemMovedEventArgs(selBudgetItem))

                Me.BeginEdit(False)
            End If
        End If
    End Sub

    Public Sub InsertChildItem()
        Dim selBudgetItem As BudgetItem = Me.CurrentItem(False)

        If selBudgetItem IsNot Nothing Then
            Dim ChildBudgetItem As New BudgetItem

            selBudgetItem.BudgetItems.Insert(0, ChildBudgetItem)
            UndoRedo.ItemChildInserted(ChildBudgetItem, selBudgetItem.BudgetItems)

            Reload()
            SetFocusOnItem(ChildBudgetItem)

            If Me.TotalBudget = True Then RaiseEvent BudgetItemInsertChild(Me, New BudgetItemMovedEventArgs(ChildBudgetItem))

            Me.BeginEdit(False)
        End If
    End Sub
#End Region

#Region "Move items"
    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selGridRow As BudgetGridRow = Grid(CurrentCell.RowIndex)
        Dim selBudgetItem As Object = Me.CurrentItem(False)

        If selBudgetItem Is Nothing Then Exit Sub

        If intDirection < 0 Then
            selBudgetItem = MoveBudgetItem_ToPreviousParent(selBudgetItem)
        Else
            selBudgetItem = MoveBudgetItem_ToNextParent(selBudgetItem)
        End If

        Reload()
        SetFocusOnItem(selBudgetItem)
        If Me.TotalBudget = True Then RaiseEvent BudgetItemMoved(Me, New BudgetItemMovedEventArgs(selBudgetItem))
    End Sub

    Private Function MoveBudgetItem_ToPreviousParent(ByVal selBudgetItem As BudgetItem) As BudgetItem
        Dim objParent As Object = CurrentLogFrame.GetParent(selBudgetItem)
        Dim objParentBudgetItems As BudgetItems = objParent.BudgetItems
        Dim intOldIndex As Integer = objParentBudgetItems.IndexOf(selBudgetItem)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim PreviousBudgetItem As BudgetItem = Me.Grid.GetPreviousBudgetItem(intRowIndex)
        Dim intIndex As Integer

        If PreviousBudgetItem Is Nothing Then Return selBudgetItem

        Dim objPreviousBudgetItems As BudgetItems = CurrentLogFrame.GetParentCollection(PreviousBudgetItem)
        intIndex = objPreviousBudgetItems.IndexOf(PreviousBudgetItem)
        If intIndex = objPreviousBudgetItems.Count - 1 Then intIndex += 1

        objParentBudgetItems.Remove(selBudgetItem)
        RaiseEvent BudgetUpdateChildTotalsNeeded(Me, New UpdateChildTotalsEventArgs(objParent))
        objPreviousBudgetItems.Insert(intIndex, selBudgetItem)
        RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(selBudgetItem))

        UndoRedo.ItemMovedUp(selBudgetItem, selBudgetItem, objParentBudgetItems, intOldIndex, objPreviousBudgetItems)

        Return selBudgetItem
    End Function

    Private Function MoveBudgetItem_ToNextParent(ByVal selBudgetItem As BudgetItem) As BudgetItem
        Dim objParent As Object = CurrentLogFrame.GetParent(selBudgetItem)
        Dim objParentBudgetItems As BudgetItems = objParent.BudgetItems
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As BudgetGridRow = Me.Grid(intRowIndex)
        Dim NextBudgetItem As BudgetItem = Me.Grid.GetNextBudgetItem(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objParentBudgetItems.IndexOf(selBudgetItem)

        If NextBudgetItem Is Nothing Then Return selBudgetItem

        If CurrentLogFrame.IsParentLineage(NextBudgetItem, selBudgetItem) Then
            intRowIndex += 1
            CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
            MoveBudgetItem_ToNextParent(selBudgetItem)
        Else
            Dim objNextBudgetItems As BudgetItems
            If CType(NextBudgetItem, BudgetItem).BudgetItems.Count > 0 Then
                'if the next budget item has sub-items, insert as first sub-item
                objNextBudgetItems = CType(NextBudgetItem, BudgetItem).BudgetItems
                intIndex = 0
            Else
                'insert before the next budget item (at the same level)
                objNextBudgetItems = CurrentLogFrame.GetParentCollection(NextBudgetItem)
                intIndex = objNextBudgetItems.IndexOf(NextBudgetItem)
            End If

            objParentBudgetItems.Remove(selBudgetItem)
            RaiseEvent BudgetUpdateChildTotalsNeeded(Me, New UpdateChildTotalsEventArgs(objParent))
            objNextBudgetItems.Insert(intIndex, selBudgetItem)
            RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(selBudgetItem))

            UndoRedo.ItemMovedDown(selBudgetItem, selBudgetItem, objParentBudgetItems, intOldIndex, objNextBudgetItems)
        End If

        Return selBudgetItem
    End Function

    Public Sub LevelUp()
        Dim selBudgetItem As BudgetItem = TryCast(Me.CurrentItem(False), BudgetItem)
        If selBudgetItem Is Nothing Then Exit Sub

        Dim Parent As BudgetItem = TryCast(CurrentLogFrame.GetParent(selBudgetItem), BudgetItem)

        If Parent Is Nothing Then Exit Sub

        Dim intOldIndex As Integer = Parent.BudgetItems.IndexOf(selBudgetItem)
        Dim objParentBudgetItems As BudgetItems = CurrentLogFrame.GetParentCollection(Parent)
        If objParentBudgetItems Is Nothing Then Exit Sub

        Dim intIndex As Integer = objParentBudgetItems.IndexOf(Parent)

        intIndex += 1
        Parent.BudgetItems.Remove(selBudgetItem)
        RaiseEvent BudgetUpdateChildTotalsNeeded(Me, New UpdateChildTotalsEventArgs(Parent))

        objParentBudgetItems.Insert(intIndex, selBudgetItem)
        RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(selBudgetItem))

        UndoRedo.ItemLevelUp(selBudgetItem, Parent.BudgetItems, intOldIndex, objParentBudgetItems)

        Reload()
        SetFocusOnItem(selBudgetItem)
        If Me.TotalBudget = True Then RaiseEvent BudgetItemLevelUp(Me, New BudgetItemMovedEventArgs(selBudgetItem))
    End Sub

    Public Sub LevelDown()
        GetSelectedGridRows()

        Dim selBudgetItem As BudgetItem = TryCast(SelectedGridRows(0).BudgetItem, BudgetItem)
        If selBudgetItem Is Nothing Then Exit Sub

        Dim objBudgetItems As BudgetItems = CurrentLogFrame.GetParentCollection(selBudgetItem)
        Dim intIndex As Integer = objBudgetItems.IndexOf(selBudgetItem)
        Dim intOldIndex As Integer = intIndex

        If intIndex = 0 Then Exit Sub

        Dim PreviousBudgetItem As BudgetItem = objBudgetItems(intIndex - 1)

        For Each selGridRow As BudgetGridRow In SelectedGridRows
            selBudgetItem = selGridRow.BudgetItem
            objBudgetItems.Remove(selBudgetItem)
            PreviousBudgetItem.BudgetItems.Add(selBudgetItem)
        Next

        RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(selBudgetItem))

        UndoRedo.ItemLevelDown(selBudgetItem, objBudgetItems, intOldIndex, PreviousBudgetItem.BudgetItems)

        Reload()
        SetFocusOnItem(selBudgetItem)
        If Me.TotalBudget = True Then RaiseEvent BudgetItemLevelDown(Me, New BudgetItemMovedEventArgs(selBudgetItem))
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim strSourceColName As String
        Dim boolShift As Boolean
        Dim boolRemoveAll As Boolean
        Dim intRowIndex, intColumnIndex As Integer
        Dim strSortNumber As String = String.Empty
        Dim objBudgetItem As BudgetItem
        Dim objParent As Object = Nothing
        Dim objBudgetItems As BudgetItems = Nothing
        Dim objGuid As Guid

        If Me.IsCurrentCellInEditMode = False Then
            intRowIndex = CurrentCell.RowIndex
            intColumnIndex = CurrentCell.ColumnIndex
            If Control.ModifierKeys = Keys.Shift Then boolShift = True

            'copy cells to delete
            strSourceColName = Columns(SelectionRectangle.FirstColumnIndex).Name
            GetSelectedGridRows()

            If ShowWarning = True And SelectedGridRows.Count > 1 Then
                Dim boolCancel As Boolean = RemoveItems_Warning(strSourceColName)
                If boolCancel = True Then Exit Sub
            End If

            For Each selGridRow As BudgetGridRow In SelectedGridRows
                strSortNumber = selGridRow.SortNumber

                If selGridRow.BudgetItem IsNot Nothing Then
                    objBudgetItem = DirectCast(selGridRow.BudgetItem, BudgetItem)
                    objGuid = objBudgetItem.Guid
                    objParent = CurrentLogFrame.GetParent(objBudgetItem)
                    If objParent IsNot Nothing Then objBudgetItems = objParent.BudgetItems 'CurrentLogFrame.GetParentCollection(objBudgetItem)

                    If objBudgetItems IsNot Nothing AndAlso objBudgetItems.Contains(objBudgetItem) Then
                        If objBudgetItem.BudgetItems.Count > 0 Then
                            If boolShift = True Then boolRemoveAll = True Else boolRemoveAll = False
                        Else
                            boolRemoveAll = True
                        End If

                        If boolRemoveAll = True Then
                            If boolCut = False Then
                                UndoRedo.ItemRemoved(objBudgetItem, objBudgetItems)
                            Else
                                UndoRedo.ItemCut(objBudgetItem, objBudgetItems)
                            End If

                            objBudgetItems.Remove(objBudgetItem)
                        Else
                            If boolCut = False Then
                                UndoRedo.ItemRemovedNotVertical(objBudgetItem, objBudgetItems)
                            Else
                                UndoRedo.ItemCutNotVertical(objBudgetItem, objBudgetItems)
                            End If

                            objBudgetItem.RTF = String.Empty
                        End If
                    End If

                    RaiseEvent BudgetUpdateChildTotalsNeeded(Me, New UpdateChildTotalsEventArgs(objParent))

                End If
            Next

            SelectedGridRows.Clear()

            ClearSelection()
            CurrentCell = Me(intColumnIndex, intRowIndex)
            Me.Reload()
            If Me.TotalBudget = True And objParent IsNot Nothing Then RaiseEvent BudgetItemRemoved(Me, New BudgetItemRemovedEventArgs(objGuid, objParent.Guid))
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

    Private Function RemoveItems_Warning(ByVal strSourceColName As String) As Boolean
        Dim intNrBudgetItems As Integer
        Dim intNrSub As Integer
        Dim strMsg As String, strTitle As String = String.Empty
        Dim strMsgBudgetItem As String = String.Empty

        For Each selGridRow As BudgetGridRow In SelectedGridRows
            If selGridRow.BudgetItem IsNot Nothing Then
                Dim selBudgetItem As BudgetItem = selGridRow.BudgetItem
                intNrSub += selBudgetItem.BudgetItems.Count
                intNrBudgetItems += 1
            End If
        Next

        If intNrBudgetItems > 0 Then
            strMsgBudgetItem = String.Format("{0} {1}", intNrBudgetItems, BudgetItem.ItemNamePlural.ToLower)

            If intNrSub > 0 Then
                strMsgBudgetItem &= String.Format(" {0}:{1}", LANG_With, vbLf)

                If intNrSub > 0 Then
                    strMsgBudgetItem &= String.Format("   - {0} {1}", intNrSub, BudgetItem.ItemNamePlural.ToLower) & vbLf
                End If
            End If

            strMsgBudgetItem &= vbLf
        End If

        strTitle = LANG_Remove
        strMsg = String.Format(LANG_RemoveBudgetItems, strMsgBudgetItem)

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

        RemoveItems(ShowWarning)
    End Sub

    Public Overrides Sub CopyItems()
        Dim selRow As BudgetGridRow
        Dim strSort As String
        Dim CopyGroup As Date = Now()

        With SelectionRectangle
            For i = .FirstRowIndex To .LastRowIndex
                selRow = Me.Grid(i)
                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.BudgetItem Then
                    strSort = BudgetItem.ItemName & " " & selRow.SortNumber
                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.BudgetItem, strSort)
                    ItemClipboard.Insert(0, NewItem)
                End If
            Next
        End With
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
        If PasteCell Is Nothing Then PasteCell = CurrentCell
        If PasteCell Is Nothing Then Exit Sub

        Dim intColumnIndex As Integer = PasteCell.ColumnIndex
        Dim intRowIndex As Integer = PasteCell.RowIndex
        Dim selItem As ClipboardItem

        PasteRow = Grid(intRowIndex)

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(BudgetItem)
                    PasteItems_BudgetItem(selItem, intPasteOption)
                Case Else
                    PasteItems_Other(selItem, intPasteOption)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(intColumnIndex, intRowIndex)
    End Sub

    Private Sub PasteBudgetItem(ByVal NewBudgetItem As BudgetItem)
        Dim Parent As Object
        Dim ParentBudgetItems As BudgetItems
        Dim PasteRowBudgetItem As BudgetItem = PasteRow.BudgetItem
        Dim intIndex As Integer
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRowBudgetItem IsNot Nothing Then
            Parent = CurrentLogFrame.GetParent(PasteRowBudgetItem)
            ParentBudgetItems = Parent.BudgetItems
            intIndex = ParentBudgetItems.IndexOf(PasteRowBudgetItem)

            ParentBudgetItems.Insert(intIndex, NewBudgetItem)

            UndoRedo.ItemPasted(NewBudgetItem, ParentBudgetItems)

            PasteRow.BudgetItem = NewBudgetItem

            RaiseEvent BudgetUpdateChildTotalsNeeded(Me, New UpdateChildTotalsEventArgs(Parent))
        End If

    End Sub

    Private Sub PasteItems_BudgetItem(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selBudgetItem As BudgetItem = DirectCast(selItem.Item, BudgetItem)

        Dim NewBudgetItem As New BudgetItem

        Using copier As New ObjectCopy
            NewBudgetItem = copier.CopyObject(selBudgetItem)
        End Using

        PasteBudgetItem(NewBudgetItem)
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

        Dim NewBudgetItem As New BudgetItem

        If String.IsNullOrEmpty(strRtf) Then
            NewBudgetItem.SetText(strText)
        Else
            NewBudgetItem.RTF = strRtf
        End If

        PasteBudgetItem(NewBudgetItem)
    End Sub
#End Region

#Region "Select text"
    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex
        ReloadImages()
        Invalidate()
    End Sub

    Public Sub HighlightTextInCell(ByVal intMatchIndex As Integer, ByVal intMatchLength As Integer)
        If CurrentItem(False) Is Nothing Then Exit Sub

        If intMatchLength > 0 Then
            If CurrentCell.IsInEditMode = True Then EndEdit()

            If CurrentCell.ColumnIndex = 1 Then
                BeginEdit(False)
                Dim ctl As RichTextEditingControlLogframe = EditingControl
                If ctl IsNot Nothing Then ctl.Select(intMatchIndex, intMatchLength)
            Else
                BeginEdit(True)
            End If
        End If
    End Sub
#End Region
End Class
