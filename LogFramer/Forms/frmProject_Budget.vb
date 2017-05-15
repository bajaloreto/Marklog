Partial Public Class frmProject
#Region "Initialise"
    Public Sub ProjectInitialise_Budget()
        Dim strBudgetTitle As String
        Dim intBudgetIndex As Integer
        Dim boolTotalBudget As Boolean

        If ProjectLogframe.Budget.BudgetYears.Count = 0 Then
            ProjectLogframe.Budget.BudgetYears.Add(New BudgetYear(ProjectLogframe.StartDate))
        End If
        With TabControlBudget
            .TabPages.Clear()
            For Each selBudgetYear As BudgetYear In CurrentLogFrame.Budget.BudgetYears
                If intBudgetIndex = 0 Then
                    strBudgetTitle = LANG_TotalBudget

                    If CurrentLogFrame.Budget.BudgetYears.Count > 1 Then boolTotalBudget = True
                Else
                    strBudgetTitle = selBudgetYear.BudgetYear.ToString("yyyy")
                    boolTotalBudget = False
                End If

                .TabPages.Add(selBudgetYear.Guid.ToString, strBudgetTitle)

                Dim NewTabPage As TabPage = .TabPages(.TabPages.Count - 1)
                Dim SplitContainerBudgetYear As New SplitContainer

                With SplitContainerBudgetYear
                    .Name = "SplitContainerBudgetYear"
                    .Orientation = Orientation.Horizontal
                    .SplitterDistance = 460
                    .Dock = DockStyle.Fill
                End With

                NewTabPage.Controls.Add(SplitContainerBudgetYear)

                Dim dgvBudgetYear As New DataGridViewBudgetYear(selBudgetYear, intBudgetIndex, boolTotalBudget)

                With dgvBudgetYear
                    .Name = "dgvBudgetYear"
                    .Dock = DockStyle.Fill
                    .ShowDurationColumns = My.Settings.setShowDurationColumns
                    .ShowLocalCurrencyColumn = My.Settings.setShowLocalCurrencyColumn
                    .LoadColumns()

                    AddHandler .Enter, AddressOf dgvBudgetYear_Enter
                    AddHandler .CellClick, AddressOf dgvBudgetYear_CellClick
                    AddHandler .MouseDown, AddressOf dgvBudgetYear_MouseDown
                    AddHandler .MouseUp, AddressOf dgvBudgetYear_MouseUp
                    AddHandler .EditingControlShowing, AddressOf dgvBudgetYear_EditingControlShowing
                    AddHandler .Reloaded, AddressOf dgvBudgetYear_Reloaded
                    AddHandler .BudgetItemSelected, AddressOf dgvBudgetYear_BudgetItemSelected
                    AddHandler .BudgetItemAdded, AddressOf dgvBudgetYear_UpdateBudgetYears_Added
                    AddHandler .BudgetItemRemoved, AddressOf dgvBudgetYear_UpdateBudgetYears_Removed
                    AddHandler .BudgetItemMoved, AddressOf dgvBudgetYear_UpdateBudgetYears_Moved
                    AddHandler .BudgetItemInsertParent, AddressOf dgvBudgetYear_UpdateBudgetYears_InsertParentHeader
                    AddHandler .BudgetItemInsertChild, AddressOf dgvBudgetYear_UpdateBudgetYears_InsertChildHeader
                    AddHandler .BudgetItemLevelDown, AddressOf dgvBudgetYear_UpdateBudgetYears_LevelDownHeader
                    AddHandler .BudgetItemLevelUp, AddressOf dgvBudgetYear_UpdateBudgetYears_LevelUpHeader
                    AddHandler .BudgetUpdateParentTotalsNeeded, AddressOf dgvBudgetYear_UpdateParentTotals
                    AddHandler .BudgetUpdateChildTotalsNeeded, AddressOf dgvBudgetYear_UpdateChildTotals
                    AddHandler .NoTextSelected, AddressOf dgvBudgetYear_NoTextSelected
                End With

                SplitContainerBudgetYear.Panel1.Controls.Add(dgvBudgetYear)
                dgvBudgetYear.Reload()

                intBudgetIndex += 1
            Next
        End With
    End Sub
#End Region

#Region "General methods"
    Public Function GetDataGridViewBudgetYear(ByVal intIndex As Integer) As DataGridViewBudgetYear
        Dim dgvBudgetYear As DataGridViewBudgetYear = Nothing

        If TabControlBudget.TabPages.Count > intIndex Then
            Dim SplitContainerBudgetYear As SplitContainer = TabControlBudget.TabPages(intIndex).Controls("SplitContainerBudgetYear")

            If SplitContainerBudgetYear IsNot Nothing Then
                dgvBudgetYear = SplitContainerBudgetYear.Panel1.Controls("dgvBudgetYear")
            End If
        End If

        Return dgvBudgetYear
    End Function

    Public Function GetCurrentDataGridViewBudgetYear() As DataGridViewBudgetYear
        Dim dgvBudgetYear As DataGridViewBudgetYear = Nothing

        If TabControlBudget.SelectedTab IsNot Nothing Then
            Dim SplitContainerBudgetYear As SplitContainer = TabControlBudget.SelectedTab.Controls("SplitContainerBudgetYear")

            If SplitContainerBudgetYear IsNot Nothing Then
                dgvBudgetYear = SplitContainerBudgetYear.Panel1.Controls("dgvBudgetYear")
            End If
        End If

        Return dgvBudgetYear
    End Function

    Public Sub ReloadAllDataGridViewBudgetYears()
        Dim dgvBudgetYear As DataGridViewBudgetYear = Nothing
        Dim SplitContainerBudgetYear As SplitContainer = Nothing

        If TabControlBudget.TabPages.Count > 0 Then
            For i = 0 To TabControlBudget.TabPages.Count - 1
                SplitContainerBudgetYear = TabControlBudget.TabPages(i).Controls("SplitContainerBudgetYear")

                If SplitContainerBudgetYear IsNot Nothing Then
                    dgvBudgetYear = SplitContainerBudgetYear.Panel1.Controls("dgvBudgetYear")

                    If dgvBudgetYear IsNot Nothing Then
                        dgvBudgetYear.Reload()
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub DetailExchangeRates_CloseButtonClicked()
        If frmParent IsNot Nothing Then
            frmParent.ExchangeRatesShowHide()
        End If
    End Sub

    Private Sub DetailExchangeRates_ExchangeRateUpdated()
        Dim dgvBudgetYear As DataGridViewBudgetYear

        If CurrentBudget.BudgetYears.Count > 0 Then
            For i = 0 To CurrentBudget.BudgetYears.Count - 1
                dgvBudgetYear = GetDataGridViewBudgetYear(i)

                If dgvBudgetYear IsNot Nothing Then dgvBudgetYear.UpdateExchangeRates()
            Next
        End If
    End Sub

    Public Sub ShowExchangeRates()
        With SplitContainerExchangeRates
            If .Panel2Collapsed Then
                DetailExchangeRates = New DetailExchangeRates(ProjectLogframe.Budget)
                DetailExchangeRates.Dock = DockStyle.Fill
                AddHandler DetailExchangeRates.CloseButtonClicked, AddressOf DetailExchangeRates_CloseButtonClicked
                AddHandler DetailExchangeRates.ExchangeRateUpdated, AddressOf DetailExchangeRates_ExchangeRateUpdated

                .Panel2.Controls.Add(DetailExchangeRates)
                .Panel2Collapsed = False
            Else
                .Panel2Collapsed = True
                DetailExchangeRates = Nothing
            End If
        End With
    End Sub
#End Region

#Region "Events"
    Private Sub dgvBudgetYear_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dgvBudgetYear As DataGridViewBudgetYear = TryCast(sender, DataGridViewBudgetYear)

        If dgvBudgetYear IsNot Nothing Then
            If Me.CurrentDataGridView IsNot dgvBudgetYear Then Me.CurrentDataGridView = dgvBudgetYear
        End If
    End Sub

    Private Sub dgvBudgetYear_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim dgvBudgetYear As DataGridViewBudgetYear = TryCast(sender, DataGridViewBudgetYear)

        If dgvBudgetYear IsNot Nothing Then
            If Me.CurrentDataGridView IsNot dgvBudgetYear Then Me.CurrentDataGridView = dgvBudgetYear
            If Me.TextSelectionIndex <> TextSelectionValues.SelectNothing Then Me.TextSelectionIndex = TextSelectionValues.SelectNothing
        End If
    End Sub

    Private Sub dgvBudgetYear_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim dgvBudgetYear As DataGridViewBudgetYear = TryCast(sender, DataGridViewBudgetYear)

        If dgvBudgetYear IsNot Nothing Then
            Dim hit As DataGridView.HitTestInfo = dgvBudgetYear.HitTest(e.X, e.Y)

            If hit.ColumnIndex > 0 And hit.RowIndex > 0 Then
                Dim selGridRow As BudgetGridRow = dgvBudgetYear.Grid(hit.RowIndex)
                Dim strObjects As String = LANG_BudgetItems
                Dim strMsg As String

                If e.Button = MouseButtons.Left Then
                    strMsg = String.Format(LANG_DragToSelectMultiple, strObjects)

                    frmParent.StatusLabelGeneral.Text = strMsg
                ElseIf e.Button = MouseButtons.Right Then
                    With frmParent
                        If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
                    End With
                    strMsg = String.Format(LANG_DragToMoveSelected, strObjects)

                    frmParent.StatusLabelGeneral.Text = strMsg
                End If

            End If
        End If
    End Sub

    Private Sub dgvBudgetYear_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        frmParent.StatusLabelGeneral.Text = String.Empty

        Dim dgvBudgetYear As DataGridViewBudgetYear = TryCast(sender, DataGridViewBudgetYear)

        If dgvBudgetYear IsNot Nothing Then
            If dgvBudgetYear.IsCurrentCellInEditMode = False Then
                With frmParent
                    If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
                End With
            End If
            'Show details window
            Dim intIndex As Integer = TabControlBudget.SelectedIndex

            If intIndex >= 0 Then
                SetTypeOfDetailWindowBudget(intIndex)
            End If
        End If
    End Sub

    Private Sub dgvBudgetYear_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
        End With
    End Sub

    Private Sub dgvBudgetYear_Reloaded(ByVal sender As Object)
        Dim dgvBudgetYear As DataGridViewBudgetYear = TryCast(sender, DataGridViewBudgetYear)

        If dgvBudgetYear IsNot Nothing Then
            If dgvBudgetYear.CurrentCell Is Nothing Then
                dgvBudgetYear.CurrentCell = dgvBudgetYear(1, 0)
                dgvBudgetYear(1, 0).Selected = True
            End If
        End If
    End Sub

    Private Sub dgvBudgetYear_NoTextSelected()
        Dim objParentForm As frmParent = MdiParent
        objParentForm.SetTextSelectionToNothing()
    End Sub
#End Region

#Region "Visibility of columns, lay-out panels and buttons"
    Private Sub dgvBudgetYear_BudgetItemSelected(ByVal sender As Object, ByVal e As BudgetItemSelectedEventArgs)
        Dim dgvBudgetYear As DataGridViewBudgetYear = CType(sender, DataGridViewBudgetYear)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        If dgvBudgetYear Is Nothing Then Exit Sub

        With frmParent
            If selBudgetItem Is Nothing Then
                .RibbonPanelReferences.Enabled = False

                .RibbonButtonInsertChild.Enabled = False
                .RibbonButtonInsertParent.Enabled = False
                .RibbonButtonLevelUp.Enabled = False
                .RibbonButtonLevelDown.Enabled = False
            Else
                Dim ParentBudgetYear As BudgetYear
                Dim ParentBudgetItem As BudgetItem
                Dim ParentBudgetItems As BudgetItems = Nothing

                If dgvBudgetYear.BudgetYearIndex = 0 Then
                    .RibbonPanelReferences.Enabled = False

                    .RibbonButtonPasteItems.Enabled = True
                    .RibbonButtonCutItems.Enabled = True
                    .RibbonButtonNewItem.Enabled = True
                    .RibbonButtonInsertItem.Enabled = True

                    .RibbonButtonInsertChild.Enabled = True
                    .RibbonButtonMoveUp.Enabled = True
                    .RibbonButtonMoveDown.Enabled = True

                    If selBudgetItem.ParentBudgetYearGuid = Guid.Empty Then
                        .RibbonButtonLevelUp.Enabled = True
                        .RibbonButtonInsertParent.Enabled = True
                    Else
                        .RibbonButtonLevelUp.Enabled = False
                        .RibbonButtonInsertParent.Enabled = False
                    End If

                    If selBudgetItem.ParentBudgetItemGuid <> Guid.Empty Then
                        ParentBudgetItem = CurrentLogFrame.GetBudgetItemByGuid(selBudgetItem.ParentBudgetItemGuid)
                        If ParentBudgetItem IsNot Nothing Then ParentBudgetItems = ParentBudgetItem.BudgetItems
                    ElseIf selBudgetItem.ParentBudgetYearGuid <> Guid.Empty Then
                        ParentBudgetYear = CurrentLogFrame.GetBudgetYearByGuid(selBudgetItem.ParentBudgetYearGuid)
                        If ParentBudgetYear IsNot Nothing Then ParentBudgetItems = ParentBudgetYear.BudgetItems
                    End If

                    If ParentBudgetItems IsNot Nothing Then
                        If ParentBudgetItems.IndexOf(selBudgetItem) > 0 Then
                            .RibbonButtonLevelDown.Enabled = True
                        Else
                            .RibbonButtonLevelDown.Enabled = False
                        End If
                    End If
                Else
                    .RibbonPanelReferences.Enabled = True

                    If dgvBudgetYear.BudgetYearIndex = 1 Then .RibbonButtonReferenceCopyPreviousYear.Enabled = False
                    If dgvBudgetYear.BudgetYearIndex = CurrentLogFrame.Budget.BudgetYears.Count - 1 Then .RibbonButtonReferenceCopyNextYear.Enabled = False

                    Select Case selBudgetItem.Type

                        Case BudgetItem.BudgetItemTypes.Reference, BudgetItem.BudgetItemTypes.Category
                            .RibbonButtonPasteItems.Enabled = False
                            .RibbonButtonCutItems.Enabled = False
                            .RibbonButtonNewItem.Enabled = False
                            .RibbonButtonInsertItem.Enabled = False
                            .RibbonButtonRemoveItem.Enabled = False
                            .RibbonButtonEditItem.Enabled = False

                            .RibbonButtonInsertParent.Enabled = False
                            .RibbonButtonMoveUp.Enabled = False
                            .RibbonButtonMoveDown.Enabled = False
                            .RibbonButtonLevelUp.Enabled = False
                            .RibbonButtonLevelDown.Enabled = False

                            If selBudgetItem.BudgetItems.Count = 0 Then
                                .RibbonButtonInsertChild.Enabled = True
                            Else
                                If selBudgetItem.BudgetItems(0).Type = BudgetItem.BudgetItemTypes.Normal Then
                                    .RibbonButtonInsertChild.Enabled = True
                                Else
                                    .RibbonButtonInsertChild.Enabled = False
                                End If
                            End If
                        Case BudgetItem.BudgetItemTypes.Normal, BudgetItem.BudgetItemTypes.Ratio
                            .RibbonButtonPasteItems.Enabled = True
                            .RibbonButtonCutItems.Enabled = True
                            .RibbonButtonNewItem.Enabled = True
                            .RibbonButtonInsertItem.Enabled = True
                            .RibbonButtonRemoveItem.Enabled = True
                            .RibbonButtonEditItem.Enabled = True

                            .RibbonButtonInsertParent.Enabled = True
                            .RibbonButtonInsertChild.Enabled = True
                            .RibbonButtonMoveUp.Enabled = True
                            .RibbonButtonMoveDown.Enabled = True

                            If selBudgetItem.ParentBudgetItemGuid <> Guid.Empty Then
                                ParentBudgetItem = CurrentLogFrame.GetBudgetItemByGuid(selBudgetItem.ParentBudgetItemGuid)

                                If ParentBudgetItem IsNot Nothing Then
                                    .RibbonButtonLevelUp.Enabled = True

                                    If ParentBudgetItem.BudgetItems.IndexOf(selBudgetItem) > 0 Then
                                        .RibbonButtonLevelDown.Enabled = True
                                    Else
                                        .RibbonButtonLevelDown.Enabled = False
                                    End If
                                End If
                            End If
                    End Select

                End If
            End If
        End With
    End Sub

    Public Sub dgvBudgetYear_SetVisibilityDurationColumns(ByVal boolVisible As Boolean)
        For i = 0 To TabControlBudget.TabPages.Count - 1
            Dim dgvBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(i)

            If CurrentDataGridViewBudgetYear IsNot Nothing Then
                With dgvBudgetYear
                    .ShowDurationColumns = boolVisible
                    .ShowColumns()
                End With
            End If
        Next
    End Sub

    Public Sub dgvBudgetYear_SetVisibilityLocalCurrencyColumn(ByVal boolVisible As Boolean)
        For i = 0 To TabControlBudget.TabPages.Count - 1
            Dim dgvBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(i)

            If dgvBudgetYear IsNot Nothing Then
                With dgvBudgetYear
                    .ShowLocalCurrencyColumn = boolVisible
                    .ShowColumns()
                End With
            End If
        Next
    End Sub

    Public Sub dgvBudgetYear_SetVisibilityEmptyCells(ByVal boolVisible As Boolean)
        For i = 0 To TabControlBudget.TabPages.Count - 1
            Dim dgvBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(i)

            If dgvBudgetYear IsNot Nothing Then
                With dgvBudgetYear
                    .HideEmptyCells = boolVisible
                    .Reload()
                End With
            End If
        Next
    End Sub
#End Region

#Region "Update budget years"
    Private Sub dgvBudgetYear_UpdateBudgetYears_Added(ByVal sender As Object, ByVal e As BudgetItemAddedEventArgs)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .InsertBudgetHeader(selBudgetItem)
                For i = 1 To .BudgetYears.Count - 1
                    Dim selDataGridViewBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(i)

                    If selDataGridViewBudgetYear IsNot Nothing Then selDataGridViewBudgetYear.Reload()
                Next
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateBudgetYears_Removed(ByVal sender As Object, ByVal e As BudgetItemRemovedEventArgs)
        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .RemoveBudgetHeader(e.Guid, e.ParentGuid)
                For i = 1 To .BudgetYears.Count - 1
                    Dim selDataGridViewBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(i)

                    If selDataGridViewBudgetYear IsNot Nothing Then selDataGridViewBudgetYear.Reload()
                Next
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateBudgetYears_Moved(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .MoveBudgetHeader(selBudgetItem)
                For i = 1 To .BudgetYears.Count - 1
                    Dim selDataGridViewBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(i)

                    If selDataGridViewBudgetYear IsNot Nothing Then selDataGridViewBudgetYear.Reload()
                Next
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateBudgetYears_InsertParentHeader(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .InsertParentHeader(selBudgetItem)
                Me.ReloadAllDataGridViewBudgetYears()
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateBudgetYears_InsertChildHeader(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .InsertChildHeader(selBudgetItem)
                Me.ReloadAllDataGridViewBudgetYears()
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateBudgetYears_LevelDownHeader(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .LevelDownHeader(selBudgetItem)
                Me.ReloadAllDataGridViewBudgetYears()
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateBudgetYears_LevelUpHeader(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
        Dim selBudgetItem As BudgetItem = e.BudgetItem

        With ProjectLogframe.Budget
            If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                .LevelUpHeader(selBudgetItem)
                Me.ReloadAllDataGridViewBudgetYears()
            End If
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateParentTotals(ByVal sender As Object, ByVal e As UpdateParentTotalsEventArgs)
        With Me.ProjectLogframe.Budget
            .UpdateParentTotals(e.ChildItem)
            .UpdateMultiYearBudget()
        End With
    End Sub

    Private Sub dgvBudgetYear_UpdateChildTotals(ByVal sender As Object, ByVal e As UpdateChildTotalsEventArgs)
        With Me.ProjectLogframe.Budget
            .UpdateChildTotals(e.Parent)
            .UpdateMultiYearBudget()
        End With
    End Sub
#End Region

#Region "Copy previous year/next year"
    Public Sub CopyReferenceValues(ByVal intDirection As Integer)
        Dim intIndexBudgetYear As Integer
        Dim selRow As BudgetGridRow
        Dim selBudgetItem, selReferenceItem As BudgetItem
        Dim selBudgetYear, selReferenceYear As BudgetYear

        If TabControlBudget.TabPages.Count > 0 AndAlso TabControlBudget.SelectedTab IsNot Nothing Then
            intIndexBudgetYear = TabControlBudget.SelectedIndex
            selBudgetYear = CurrentLogFrame.Budget.BudgetYears(intIndexBudgetYear)

            If intDirection < 0 And intIndexBudgetYear <= 1 Then
                Exit Sub
            ElseIf intDirection > 0 And intIndexBudgetYear >= CurrentLogFrame.Budget.BudgetYears.Count - 1 Then
                Exit Sub
            End If

            selReferenceYear = CurrentLogFrame.Budget.BudgetYears(intIndexBudgetYear + intDirection)
            With CurrentDataGridViewBudgetYear
                For i = .SelectionRectangle.FirstRowIndex To .SelectionRectangle.LastRowIndex
                    selRow = .Grid(i)
                    selBudgetItem = selRow.BudgetItem

                    If selBudgetItem IsNot Nothing AndAlso selBudgetItem.ReferenceBudgetItemGuid <> Guid.Empty Then
                        selReferenceItem = selReferenceYear.BudgetItems.GetBudgetItemByReferenceGuid(selBudgetItem.ReferenceBudgetItemGuid)

                        CopyReferenceValues_CopyValues(selBudgetItem, selReferenceItem, selBudgetYear, selReferenceYear)

                        With CurrentBudget
                            .UpdateParentTotals(selBudgetItem)
                            .UpdateMultiYearBudget()
                        End With
                    End If
                Next

                .Reload()
            End With
        End If

    End Sub

    Public Sub CopyReferenceValues_CopyValues(ByVal selBudgetItem As BudgetItem, ByVal selReferenceItem As BudgetItem, ByVal selBudgetYear As BudgetYear, ByVal selReferenceYear As BudgetYear)

        Select Case selReferenceItem.Type
            Case BudgetItem.BudgetItemTypes.Normal, BudgetItem.BudgetItemTypes.Reference
                selBudgetItem.Duration = selReferenceItem.Duration
                selBudgetItem.DurationUnit = selReferenceItem.DurationUnit
                selBudgetItem.Number = selReferenceItem.Number
                selBudgetItem.NumberUnit = selReferenceItem.NumberUnit
                selBudgetItem.UnitCost = New Currency(selReferenceItem.UnitCostAmount, selReferenceItem.UnitCostCurrencyCode, selReferenceItem.UnitCostExchangeRate)
                selBudgetItem.TotalLocalCost = New Currency(selReferenceItem.TotalLocalCostAmount, selReferenceItem.TotalLocalCostCurrencyCode, selReferenceItem.TotalLocalCostExchangeRate)
                selBudgetItem.TotalCost = New Currency(selReferenceItem.TotalCostAmount, selReferenceItem.TotalCostCurrencyCode)
            Case BudgetItem.BudgetItemTypes.Ratio
                Dim objReferenceSourceItem As BudgetItem = selReferenceYear.BudgetItems.GetBudgetItemByGuid(selReferenceItem.ReferenceBudgetItemGuid)
                If objReferenceSourceItem IsNot Nothing Then
                    Dim objSourceItem As BudgetItem = selBudgetYear.BudgetItems.GetBudgetItemByReferenceGuid(objReferenceSourceItem.ReferenceBudgetItemGuid)

                    If objSourceItem IsNot Nothing Then selBudgetItem.ReferenceBudgetItemGuid = objSourceItem.Guid
                    selBudgetItem.Ratio = selReferenceItem.Ratio
                    'selBudgetItem.CalculationGuidRatio = ?
                    selBudgetItem.TotalLocalCost = New Currency(selReferenceItem.TotalLocalCostAmount, selReferenceItem.TotalLocalCostCurrencyCode, selReferenceItem.TotalLocalCostExchangeRate)
                    selBudgetItem.TotalCost = New Currency(selReferenceItem.TotalCostAmount, selReferenceItem.TotalCostCurrencyCode)
                End If
                
            Case BudgetItem.BudgetItemTypes.Category
                If selReferenceItem.BudgetItems.Count > 0 Then
                    Dim FirstChild As BudgetItem = selReferenceItem.BudgetItems(0)

                    If FirstChild.Type = BudgetItem.BudgetItemTypes.Normal Or FirstChild.Type = BudgetItem.BudgetItemTypes.Ratio Then
                        Using copier As New ObjectCopy
                            selBudgetItem.BudgetItems = copier.CopyCollection(selReferenceItem.BudgetItems)
                        End Using
                        For Each selChildItem As BudgetItem In selBudgetItem.BudgetItems
                            selChildItem.ParentBudgetItemGuid = selBudgetItem.Guid
                        Next

                        selBudgetItem.TotalLocalCost = New Currency(selReferenceItem.TotalLocalCostAmount, selReferenceItem.TotalLocalCostCurrencyCode, selReferenceItem.TotalLocalCostExchangeRate)
                        selBudgetItem.TotalCost = New Currency(selReferenceItem.TotalCostAmount, selReferenceItem.TotalCostCurrencyCode)
                    End If
                End If
        End Select
    End Sub
#End Region
End Class
