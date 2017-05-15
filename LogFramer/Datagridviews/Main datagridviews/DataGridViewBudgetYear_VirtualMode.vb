Partial Public Class DataGridViewBudgetYear
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)
        Dim RowTmp As BudgetGridRow = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name
        Dim selBudgetItem As BudgetItem

        If e.RowIndex >= IndexTotal Then Return

        ' Store a reference to the planning grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), BudgetGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub
        selBudgetItem = RowTmp.BudgetItem
        If selBudgetItem Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "SortNumber"
                e.Value = RowTmp.SortNumber
            Case "RTF"
                e.Value = RowTmp.RTF
            Case "Duration"
                If RowTmp.DurationSet Then
                    e.Value = RowTmp.Duration
                End If
            Case "DurationUnit"
                If RowTmp.DurationSet Then
                    e.Value = RowTmp.DurationUnit
                End If
            Case "Number"
                If RowTmp.IsSubTotal Then
                    If RowTmp.Number > 0 Then
                        If selBudgetItem.BudgetItems.UniformNumberUnit = True Then _
                            e.Value = RowTmp.Number
                    End If
                Else
                    If RowTmp.NumberSet Then
                        e.Value = RowTmp.Number
                    End If
                End If
            Case "NumberUnit"
                If RowTmp.IsSubTotal Then
                    If RowTmp.Number > 0 Then
                        If selBudgetItem.BudgetItems.UniformNumberUnit = True Then _
                            e.Value = RowTmp.NumberUnit
                    End If
                Else
                    If RowTmp.NumberSet Then
                        e.Value = RowTmp.NumberUnit
                    End If
                End If
            Case "UnitCost"
                If ShowLocalCurrencyColumn Then
                    e.Value = RowTmp.UnitCost.Amount.ToString("N2")
                Else
                    e.Value = RowTmp.UnitCost.ToString
                End If
            Case "CurrencyCode"
                If RowTmp.UnitCost IsNot Nothing Then _
                    e.Value = RowTmp.UnitCost.CurrencyCode
            Case "TotalLocalCost"
                If RowTmp.IsSubTotal Then
                    If RowTmp.TotalLocalCost.Amount > 0 Then
                        If selBudgetItem.BudgetItems.UniformCurrencyCode = True Then _
                            e.Value = RowTmp.TotalLocalCost.ToString
                    End If
                Else
                    e.Value = RowTmp.TotalLocalCost.ToString
                End If
            Case "TotalCost"
                e.Value = RowTmp.TotalCost.ToString
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As BudgetGridRow
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name
        Dim strCellValue As String = String.Empty
        Dim selRowIndex As Integer = e.RowIndex
        Dim UpdateParentTotals As Boolean

        If e.RowIndex >= IndexTotal Then Return

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new planning grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As BudgetGridRow = CType(Me.Grid(selRowIndex), BudgetGridRow)

                Me.EditRow = New BudgetGridRow
                With EditRow
                    .BudgetItem = CurrentGridRow.BudgetItem
                    .BudgetItemHeight = CurrentGridRow.BudgetItemHeight
                    .BudgetItemImage = CurrentGridRow.BudgetItemImage
                    .BudgetItemImageDirty = CurrentGridRow.BudgetItemImageDirty
                    .Indent = CurrentGridRow.Indent
                    .SortNumber = CurrentGridRow.SortNumber
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        ' Set the appropriate objRowEdit property to the cell value entered.
        Select Case strColName
            Case "RTF"
                Dim ctlRTF As RichTextEditingControlLogframe = CType(Me.EditingControl, RichTextEditingControlLogframe)

                'necessary for pasting into non-active cell
                If ctlRTF Is Nothing Then
                    BeginEdit(False)
                    ctlRTF = CType(Me.EditingControl, RichTextEditingControlLogframe)
                End If

                strCellValue = TryCast(ctlRTF.Rtf, String)
        End Select

        If RowTmp.BudgetItem Is Nothing Then RowTmp.BudgetItem = InitialiseBudgetItem(RowTmp, selRowIndex)

        Select Case strColName
            Case "RTF"
                RowTmp.RTF = strCellValue
                RowTmp.BudgetItemImageDirty = True
            Case "Duration"
                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "Duration", RowTmp.BudgetItem.Duration)

                RowTmp.Duration = ParseSingle(e.Value)
                UndoRedo.ValueChanged(RowTmp.Duration)

                If String.IsNullOrEmpty(RowTmp.DurationUnit) Then
                    UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "DurationUnit", String.Empty)
                    RowTmp.DurationUnit = LANG_Months
                    UndoRedo.TextChanged(RowTmp.DurationUnit)
                End If
                RowTmp.BudgetItem.SetTotalCost()
                UpdateParentTotals = True
            Case "DurationUnit"
                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "DurationUnit", RowTmp.DurationUnit)
                RowTmp.DurationUnit = e.Value
                UndoRedo.TextChanged(RowTmp.DurationUnit)
            Case "Number"
                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "Number", RowTmp.Number)
                RowTmp.Number = ParseSingle(e.Value)
                UndoRedo.ValueChanged(RowTmp.Number)

                RowTmp.BudgetItem.SetTotalCost()
                UpdateParentTotals = True
            Case "NumberUnit"
                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "NumberUnit", RowTmp.NumberUnit)
                RowTmp.NumberUnit = e.Value
                UndoRedo.TextChanged(RowTmp.NumberUnit)

                UpdateParentTotals = True
            Case "UnitCost"
                Dim sngValue As Single = ParseSingle(e.Value)
                Dim curUnitCost As New Currency(sngValue, CurrentLogFrame.CurrencyCode)

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "UnitCost", RowTmp.UnitCost)
                RowTmp.UnitCost = curUnitCost
                UndoRedo.AmountChanged(curUnitCost)

                RowTmp.BudgetItem.SetTotalCost()
                UpdateParentTotals = True
            Case "CurrencyCode"
                Dim strCurrencyCode As String = e.Value

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "CurrencyCode", RowTmp.UnitCost.CurrencyCode)
                RowTmp.UnitCost.CurrencyCode = strCurrencyCode
                UndoRedo.TextChanged(strCurrencyCode)

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "CurrencyCode", RowTmp.TotalLocalCost.CurrencyCode)
                RowTmp.TotalLocalCost.CurrencyCode = strCurrencyCode
                UndoRedo.TextChanged(strCurrencyCode)

                Dim sngExchangeRate As Single = CurrentLogFrame.Budget.ExchangeRates.GetExchangeRate(RowTmp.UnitCost.CurrencyCode)

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "ExchangeRate", RowTmp.UnitCost.ExchangeRate)
                RowTmp.UnitCost.ExchangeRate = sngExchangeRate
                UndoRedo.ValueChanged(sngExchangeRate)

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItem, "ExchangeRate", RowTmp.TotalLocalCost.ExchangeRate)
                RowTmp.TotalLocalCost.ExchangeRate = sngExchangeRate
                UndoRedo.ValueChanged(sngExchangeRate)

                RowTmp.BudgetItem.SetTotalCost()
                UpdateParentTotals = True
        End Select

        If UpdateParentTotals = True Then RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(RowTmp.BudgetItem))
    End Sub

    Private Function InitialiseBudgetItem(ByVal selGridRow As BudgetGridRow, ByVal intRowIndex As Integer) As BudgetItem
        Dim NewBudgetItem As BudgetItem = Nothing

        Dim PreviousBudgetItem As BudgetItem = Grid.GetPreviousBudgetItem(intRowIndex)
        Dim ParentPurpose As Purpose = Nothing
        Dim ParentOutput As Output = Nothing
        NewBudgetItem = New BudgetItem

        If intRowIndex = 0 Then
            BudgetYear.BudgetItems.Add(NewBudgetItem)

            Return NewBudgetItem
        ElseIf intRowIndex > 0 Then
            PreviousBudgetItem = Grid(intRowIndex - 1).BudgetItem
        ElseIf Me.BudgetYear.BudgetItems.Count > 0 Then
            PreviousBudgetItem = Me.BudgetYear.BudgetItems(0)
        End If

        If PreviousBudgetItem Is Nothing Then
            'No budget items identified yet
            If BudgetYear.BudgetItems.Count = 0 Then
                BudgetYear.BudgetItems.Add(New BudgetItem())

                UndoRedo.ItemInserted(BudgetYear.BudgetItems(0), BudgetYear.BudgetItems)
            End If
            PreviousBudgetItem = Me.BudgetYear.BudgetItems(0)
        End If

        If PreviousBudgetItem.ParentBudgetYearGuid <> Guid.Empty Then
            BudgetYear.BudgetItems.Add(NewBudgetItem)

            UndoRedo.ItemInserted(NewBudgetItem, BudgetYear.BudgetItems)
        ElseIf PreviousBudgetItem.ParentBudgetItemGuid <> Guid.Empty Then
            Dim ParentBudgetItem As BudgetItem = CurrentLogFrame.GetBudgetItemByGuid(PreviousBudgetItem.ParentBudgetItemGuid)
            If ParentBudgetItem IsNot Nothing Then
                ParentBudgetItem.BudgetItems.Add(NewBudgetItem)

                UndoRedo.ItemInserted(NewBudgetItem, ParentBudgetItem.BudgetItems)
            End If
        End If

        Return NewBudgetItem
    End Function

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New BudgetGridRow
        Else
            ' If the user has canceled the edit of an existing row, 
            ' release the corresponding logframe row.
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        End If
        Me.Reload()
        MyBase.OnCancelRowEdit(e)

    End Sub

    Protected Overrides Sub OnNewRowNeeded(ByVal e As System.Windows.Forms.DataGridViewRowEventArgs)
        Me.EditRow = New BudgetGridRow()
        Me.EditRowFlag = Me.Rows.Count - 1
    End Sub

    Protected Overrides Sub OnRowDirtyStateNeeded(ByVal e As System.Windows.Forms.QuestionEventArgs)
        If Not rowScopeCommit Then

            ' In cell-level commit scope, indicate whether the value
            ' of the current cell has been modified.
            e.Response = Me.IsCurrentCellDirty

        End If
    End Sub

    Protected Overrides Sub OnRowValidated(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ' Save row changes if any were made and release the edited 
        ' planning grid row if there is one.
        If e.RowIndex >= Me.Grid.Count AndAlso e.RowIndex <> Me.Rows.Count - 1 Then

            ' Add the new planning grid row to grid.
            Me.Grid.Add(Me.EditRow)

            If Me.TotalBudget = True Then RaiseEvent BudgetItemAdded(Me, New BudgetItemAddedEventArgs(EditRow.BudgetItem))

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then

            ' Save the modified planning grid row in grid.
            Me.Grid(e.RowIndex) = Me.EditRow

            If Me.TotalBudget = True Then RaiseEvent BudgetItemAdded(Me, New BudgetItemAddedEventArgs(EditRow.BudgetItem))

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        MyBase.OnRowValidated(e)
    End Sub
End Class
