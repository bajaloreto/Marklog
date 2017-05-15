Public Class DataGridViewBudgetItemReferences
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New BudgetItemReferenceGridRows
    Friend WithEvents SelectionRectangle As New SelectionRectangle

    Public Event TotalCostChanged()

    Private objCurrentResource As Resource
    Private objBudgetItemsList As Dictionary(Of Guid, String)

    Private objCellLocation As New Point
    Private colSortNumber As New DataGridViewTextBoxColumn
    Private colReferenceBudgetItem As New DataGridViewComboBoxColumn
    Private colPercentage As New DataGridViewTextBoxColumn
    Private colTotalCost As New DataGridViewTextBoxColumn

    Private RowModifyIndex As Integer = -1
    Private EditRow As BudgetItemReferenceGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private PasteRow As BudgetItemReferenceGridRow
    Private SelectedGridRows As New List(Of BudgetItemReferenceGridRow)

#Region "Properties"
    Public Property CurrentResource As Resource
        Get
            Return objCurrentResource
        End Get
        Set(ByVal value As Resource)
            objCurrentResource = value
        End Set
    End Property

    Public ReadOnly Property IndexTotal As Integer
        Get
            Return Me.Grid.Count
        End Get
    End Property

    Public Property BudgetItemsList As Dictionary(Of Guid, String)
        Get
            Return objBudgetItemsList
        End Get
        Set(value As Dictionary(Of Guid, String))
            objBudgetItemsList = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim selPlanningRow As BudgetItemReferenceGridRow = Me.Grid(Me.CurrentCell.RowIndex)
                If selPlanningRow.BudgetItemReference IsNot Nothing Then
                    selObject = selPlanningRow.BudgetItemReference
                End If
            End If
            Return selObject
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New(ByVal resource As Resource)
        'datagridview settings
        DoubleBuffered = True
        VirtualMode = True
        AutoGenerateColumns = False
        AllowUserToAddRows = True
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = False

        ShowCellToolTips = False
        BackgroundColor = Color.White
        DefaultCellStyle.Padding = New Padding(2)
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        RowHeadersVisible = True

        With ColumnHeadersDefaultCellStyle
            .Font = New Font(DefaultFont, FontStyle.Bold)
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .WrapMode = DataGridViewTriState.True
        End With
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

        'load values
        Me.CurrentResource = resource

        LoadColumns()
    End Sub

    Private Sub LoadColumns()
        Columns.Clear()

        'With colSortNumber
        '    .Name = "SortNumber"
        '    .HeaderText = LANG_Number
        '    .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        'End With
        'Me.Columns.Add(colSortNumber)

        BudgetItemsList = CurrentLogFrame.Budget.GetBudgetItemsList()
        With colReferenceBudgetItem
            .Name = "ReferenceBudgetItem"
            .HeaderText = LANG_BudgetItem
            .MinimumWidth = 120
            .FillWeight = 100
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .AutoComplete = True
            .DisplayMember = "Value"
            .ValueMember = "Key"
            .DataSource = New BindingSource(BudgetItemsList, Nothing)
        End With
        Columns.Add(colReferenceBudgetItem)

        With colPercentage
            .Name = "Percentage"
            .HeaderText = LANG_Percentage
            .MinimumWidth = 120
            .FillWeight = 50
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
        Me.Columns.Add(colPercentage)

        With colTotalCost
            .Name = "TotalCost"
            .HeaderText = LANG_Amount
            .MinimumWidth = 120
            .FillWeight = 50
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
        Me.Columns.Add(colTotalCost)

        Reload()
    End Sub

    Public Sub Reload()
        'remember current cell location
        objCellLocation = CurrentCellAddress
        If objCellLocation.X < 0 Then objCellLocation.X = 0
        If objCellLocation.Y < 0 Then objCellLocation.Y = 0

        '(re)load target columns and rows
        Me.SuspendLayout()

        LoadBudgetItemReferences()

        'set current cell location to what it was before
        Me.Invalidate()
        Me.ResumeLayout()
        CurrentCell = Me(objCellLocation.X, objCellLocation.Y)
    End Sub

    Public Sub LoadBudgetItemReferences()
        Dim intIndex As Integer
        Dim strSort As String
        Me.Grid.Clear()
        EditRow = Nothing
        Me.RowModifyIndex = -1

        For Each selBudgetItemReference As BudgetItemReference In CurrentResource.BudgetItemReferences
            strSort = CurrentLogFrame.CreateSortNumber(intIndex)
            Grid.Add(New BudgetItemReferenceGridRow(strSort, selBudgetItemReference))
        Next
        Grid.Add(New BudgetItemReferenceGridRow)

        Me.RowCount = Me.Grid.Count + 1

        Me.Invalidate()
    End Sub
#End Region

#Region "Events"
    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)

        MoveCurrentCell()
    End Sub
#End Region

#Region "General methods"
    Public Sub AddNewItem()
        If CurrentCell Is Nothing Then Exit Sub

        Dim intRowIndex As Integer
        Dim intColIndex = 1
        Dim strColName As String = Me.Columns(intColIndex).Name
        Dim selGridRow As BudgetItemReferenceGridRow

        For intRowIndex = CurrentCell.RowIndex To RowCount - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.BudgetItemReference Is Nothing Then Exit For
        Next

        CurrentCell = Me(intColIndex, intRowIndex)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selBudgetItemReference As BudgetItemReference = Me.CurrentItem(False)
        Dim intIndex As Integer

        If selBudgetItemReference Is Nothing Then Exit Sub

        intIndex = CurrentResource.BudgetItemReferences.IndexOf(selBudgetItemReference)

        Dim NewBudgetItem As New BudgetItemReference
        CurrentResource.BudgetItemReferences.Insert(intIndex, NewBudgetItem)
        UndoRedo.ItemInserted(NewBudgetItem, CurrentResource.BudgetItemReferences)

        selBudgetItemReference = NewBudgetItem

        Reload()

        SetFocusOnItem(selBudgetItemReference)

        Me.BeginEdit(False)
    End Sub

    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selGridRow As BudgetItemReferenceGridRow = Grid(CurrentCell.RowIndex)
        Dim selBudgetItemReference As BudgetItemReference = Me.CurrentItem(False)
        Dim intOldIndex As Integer = CurrentCell.RowIndex - 1

        If intOldIndex <= 0 Then intOldIndex = 0
        If selBudgetItemReference Is Nothing Then Exit Sub

        If intDirection < 0 Then
            selBudgetItemReference = MoveBudgetItemReference_ToPrevious(selBudgetItemReference)
            UndoRedo.ItemMovedUp(selBudgetItemReference, selBudgetItemReference, CurrentResource.BudgetItemReferences, intOldIndex, CurrentResource.BudgetItemReferences)
        Else
            selBudgetItemReference = MoveBudgetItemReference_ToNext(selBudgetItemReference)
            UndoRedo.ItemMovedDown(selBudgetItemReference, selBudgetItemReference, CurrentResource.BudgetItemReferences, intOldIndex, CurrentResource.BudgetItemReferences)
        End If

        Reload()
        SetFocusOnItem(selBudgetItemReference)
    End Sub

    Private Function MoveBudgetItemReference_ToPrevious(ByVal selBudgetItemReference As BudgetItemReference) As BudgetItemReference
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intIndex As Integer = intRowIndex - 1

        If intIndex <= 0 Then intIndex = 0

        CurrentResource.BudgetItemReferences.Remove(selBudgetItemReference)
        CurrentResource.BudgetItemReferences.Insert(intIndex, selBudgetItemReference)

        Return selBudgetItemReference
    End Function

    Private Function MoveBudgetItemReference_ToNext(ByVal selBudgetItemReference As BudgetItemReference) As BudgetItemReference
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intIndex As Integer = intRowIndex + 1

        If intIndex >= CurrentResource.BudgetItemReferences.Count - 1 Then intIndex = CurrentResource.BudgetItemReferences.Count - 1

        CurrentResource.BudgetItemReferences.Remove(selBudgetItemReference)
        CurrentResource.BudgetItemReferences.Insert(intIndex, selBudgetItemReference)

        Return selBudgetItemReference
    End Function

    Public Sub SetFocusOnItem(ByVal selItem As BudgetItemReference, Optional ByVal strPropertyName As String = "")
        Dim selGridRow As BudgetItemReferenceGridRow
        Dim intColIndex As Integer = 1
        Dim intRowIndex As Integer

        If String.IsNullOrEmpty(strPropertyName) = False Then
            Dim selColumn As DataGridViewColumn = Columns(strPropertyName)

            If selColumn IsNot Nothing Then intColIndex = Columns.IndexOf(selColumn)
        End If
        
        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.BudgetItemReference Is selItem Then
                Exit For
            End If
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Private Sub MoveCurrentCell()
        If CurrentCell IsNot Nothing Then
            Dim strColName As String = Columns(CurrentCell.ColumnIndex).Name

            If CurrentCell.RowIndex = IndexTotal Then
                CurrentCell = Me(CurrentCell.ColumnIndex, CurrentCell.RowIndex - 1)
                CurrentCell.Selected = True

                MoveCurrentCell()
            ElseIf strColName.Contains("Number") Then
                CurrentCell = Me(CurrentCell.ColumnIndex + 1, CurrentCell.RowIndex)
                CurrentCell.Selected = True
            End If
        End If
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim boolShift As Boolean
        Dim intRowIndex, intColumnIndex As Integer
        Dim strSortNumber As String = String.Empty
        Dim objBudgetItemReference As BudgetItemReference

        If Me.IsCurrentCellInEditMode = False Then
            intRowIndex = CurrentCell.RowIndex
            intColumnIndex = CurrentCell.ColumnIndex
            If Control.ModifierKeys = Keys.Shift Then boolShift = True

            'copy cells to delete
            For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
                SelectedGridRows.Add(Me.Grid(i))
            Next

            If ShowWarning = True Then
                Dim boolCancel As Boolean = RemoveItems_Warning()
                If boolCancel = True Then Exit Sub
            End If

            For Each selGridRow As BudgetItemReferenceGridRow In SelectedGridRows
                strSortNumber = selGridRow.SortNumber

                If selGridRow.BudgetItemReference IsNot Nothing Then
                    objBudgetItemReference = DirectCast(selGridRow.BudgetItemReference, BudgetItemReference)
                    CurrentResource.BudgetItemReferences = CurrentLogFrame.GetParentCollection(objBudgetItemReference)

                    If CurrentResource.BudgetItemReferences IsNot Nothing AndAlso CurrentResource.BudgetItemReferences.Contains(objBudgetItemReference) Then
                        If boolCut = False Then
                            UndoRedo.ItemRemoved(objBudgetItemReference, CurrentResource.BudgetItemReferences)
                        Else
                            UndoRedo.ItemCut(objBudgetItemReference, CurrentResource.BudgetItemReferences)
                        End If

                        CurrentResource.BudgetItemReferences.Remove(objBudgetItemReference)
                    End If
                End If
            Next

            SelectedGridRows.Clear()

            ClearSelection()
            CurrentCell = Me(intColumnIndex, intRowIndex)
            Me.Reload()
        Else
            With CurrentEditingControl
                Dim intSelectionStart As Integer = .SelectionStart
                Dim intLength As Integer = .SelectionLength

                If intLength = 0 Then intLength = 1
                .Text = .Text.Remove(intSelectionStart, intLength)

                .Select(intSelectionStart, 0)
            End With
        End If
    End Sub

    Private Function RemoveItems_Warning() As Boolean
        Dim intNrBudgetItemReferences As Integer
        Dim strMsg As String, strTitle As String = String.Empty
        Dim strMsgBudgetItemReference As String = String.Empty

        For Each selGridRow As BudgetItemReferenceGridRow In SelectedGridRows
            If selGridRow.BudgetItemReference IsNot Nothing Then
                Dim selBudgetItemReference As BudgetItemReference = selGridRow.BudgetItemReference
                intNrBudgetItemReferences += 1
            End If
        Next

        If intNrBudgetItemReferences > 0 Then
            strMsgBudgetItemReference = String.Format("{0} {1}{2}", intNrBudgetItemReferences, BudgetItemReference.ItemNamePlural.ToLower, vbLf)
        End If

        strTitle = LANG_Remove
        strMsg = String.Format(LANG_RemoveBudgetItemReferences, strMsgBudgetItemReference)

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
        Dim selRow As BudgetItemReferenceGridRow
        Dim strSort As String
        Dim CopyGroup As Date = Now()

        With SelectionRectangle
            For i = .FirstRowIndex To .LastRowIndex
                selRow = Me.Grid(i)
                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.BudgetItemReference Then
                    strSort = BudgetItemReference.ItemName & " " & selRow.SortNumber
                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.BudgetItemReference, strSort)
                    ItemClipboard.Insert(0, NewItem)
                End If
            Next
        End With
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
        If CurrentCell Is Nothing Then Exit Sub

        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim intRowIndex As Integer = CurrentCell.RowIndex

        PasteRow = Grid(intRowIndex)
        Dim selItem As ClipboardItem

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(BudgetItemReference)
                    PasteItems_BudgetItemReference(selItem, intPasteOption)
                Case GetType(BudgetItem)
                    PasteItems_BudgetItem(selItem, intPasteOption)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(intColumnIndex, intRowIndex)
    End Sub

    Private Sub PasteBudgetItemReference(ByVal NewBudgetItemReference As BudgetItemReference)
        Dim PasteRowBudgetItem As BudgetItemReference = PasteRow.BudgetItemReference
        Dim intIndex As Integer
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRowBudgetItem IsNot Nothing Then
            intIndex = CurrentResource.BudgetItemReferences.IndexOf(PasteRowBudgetItem)
        Else
            intIndex = 0
        End If
        CurrentResource.BudgetItemReferences.Insert(intIndex, NewBudgetItemReference)

        UndoRedo.ItemPasted(NewBudgetItemReference, CurrentResource.BudgetItemReferences)

    End Sub

    Private Sub PasteItems_BudgetItemReference(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selBudgetItemReference As BudgetItemReference = DirectCast(selItem.Item, BudgetItemReference)

        Dim NewBudgetItemReference As New BudgetItemReference

        Using copier As New ObjectCopy
            NewBudgetItemReference = copier.CopyObject(selBudgetItemReference)
        End Using

        PasteBudgetItemReference(NewBudgetItemReference)
    End Sub

    Private Sub PasteItems_BudgetItem(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selBudgetItem As BudgetItem = TryCast(selItem.Item, BudgetItem)
        Dim strText As String = String.Empty, strRtf As String = String.Empty

        If selBudgetItem IsNot Nothing AndAlso BudgetItemsList.ContainsKey(selItem.ReferenceGuid) Then
            Dim NewBudgetItemReference As New BudgetItemReference(selItem.ReferenceGuid)

            CalculateTotalCost(NewBudgetItemReference)

            PasteBudgetItemReference(NewBudgetItemReference)
        End If
    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)
        Dim RowTmp As BudgetItemReferenceGridRow = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name

        If e.RowIndex >= IndexTotal Then Return

        ' Store a reference to the planning grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), BudgetItemReferenceGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "SortNumber"
                e.Value = RowTmp.SortNumber
            Case "ReferenceBudgetItem"
                e.Value = RowTmp.ReferenceBudgetItemGuid
            Case "Percentage"
                e.Value = RowTmp.Percentage.ToString("P2")
            Case "TotalCost"
                e.Value = RowTmp.TotalCost.ToString
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As BudgetItemReferenceGridRow
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name
        Dim selRowIndex As Integer = e.RowIndex

        If e.RowIndex >= IndexTotal Then Return

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new planning grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As BudgetItemReferenceGridRow = CType(Me.Grid(selRowIndex), BudgetItemReferenceGridRow)

                Me.EditRow = New BudgetItemReferenceGridRow(CurrentGridRow.SortNumber, CurrentGridRow.BudgetItemReference)
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        If RowTmp.BudgetItemReference Is Nothing Then RowTmp.BudgetItemReference = InitialiseBudgetItemReference(RowTmp, selRowIndex)

        Select Case strColName
            Case "ReferenceBudgetItem"
                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItemReference, "ReferenceBudgetItemGuid", RowTmp.ReferenceBudgetItemGuid)
                RowTmp.ReferenceBudgetItemGuid = e.Value
                UndoRedo.OptionChanged(RowTmp.ReferenceBudgetItemGuid)

                CalculateTotalCost(RowTmp.BudgetItemReference)
            Case "Percentage"
                Dim sngPercentage As Single = ParseSingle(e.Value) / 100

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItemReference, "Percentage", RowTmp.Percentage)
                RowTmp.Percentage = sngPercentage
                UndoRedo.ValueChanged(sngPercentage)

                CalculateTotalCost(RowTmp.BudgetItemReference)
            Case "TotalCost"
                Dim sngValue As Single = ParseSingle(e.Value)

                UndoRedo.UndoBuffer_Initialise(RowTmp.BudgetItemReference, "TotalCost", RowTmp.TotalCost)
                Dim curTotalCost As New Currency(sngValue, CurrentLogFrame.CurrencyCode)
                RowTmp.TotalCost = curTotalCost
                UndoRedo.AmountChanged(curTotalCost)

                CalculatePercentage(RowTmp.BudgetItemReference)
        End Select

        'CurrentUndoList.AddToUndoList()
    End Sub

    Private Function InitialiseBudgetItemReference(ByVal selGridRow As BudgetItemReferenceGridRow, ByVal intRowIndex As Integer) As BudgetItemReference
        Dim NewBudgetItemReference As New BudgetItemReference
        Dim objReferenceGuid As Guid

        If BudgetItemsList.Count > 1 Then
            objReferenceGuid = BudgetItemsList.ElementAt(1).Key

            NewBudgetItemReference.ReferenceBudgetItemGuid = objReferenceGuid

            CalculateTotalCost(NewBudgetItemReference)
        End If

        CurrentResource.BudgetItemReferences.Add(NewBudgetItemReference)

        Return NewBudgetItemReference
    End Function

    Public Sub CalculateTotalCost(ByVal selBudgetItemReference As BudgetItemReference)
        If selBudgetItemReference.ReferenceBudgetItemGuid <> Guid.Empty Then
            Dim selBudgetItem As BudgetItem = CurrentLogFrame.Budget.GetBudgetItemByGuid(selBudgetItemReference.ReferenceBudgetItemGuid)

            If selBudgetItem IsNot Nothing Then
                Dim sngAmount = selBudgetItem.TotalCostAmount
                sngAmount *= selBudgetItemReference.Percentage

                selBudgetItemReference.TotalCost = New Currency(sngAmount, selBudgetItem.TotalCostCurrencyCode)

                RaiseEvent TotalCostChanged()
            End If
        End If
    End Sub

    Public Sub CalculatePercentage(ByVal selBudgetItemReference As BudgetItemReference)
        If selBudgetItemReference.ReferenceBudgetItemGuid <> Guid.Empty Then
            Dim selBudgetItem As BudgetItem = CurrentLogFrame.Budget.GetBudgetItemByGuid(selBudgetItemReference.ReferenceBudgetItemGuid)

            If selBudgetItem IsNot Nothing Then
                Dim sngPercentage As Single

                If selBudgetItem.TotalCost.Amount > 0 Then
                    sngPercentage = selBudgetItemReference.TotalCost.Amount / selBudgetItem.TotalCost.Amount
                End If

                selBudgetItemReference.Percentage = sngPercentage

                RaiseEvent TotalCostChanged()
            End If
        End If
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New BudgetItemReferenceGridRow
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
        Me.EditRow = New BudgetItemReferenceGridRow()
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
            If EditRow.BudgetItemReference.ReferenceBudgetItemGuid <> Guid.Empty Then _
                Me.Grid.Add(Me.EditRow)

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then

            ' Save the modified planning grid row in grid.
            With Me.Grid(e.RowIndex)
                .ReferenceBudgetItemGuid = EditRow.ReferenceBudgetItemGuid
                .Percentage = EditRow.Percentage
                .TotalCost = EditRow.TotalCost
            End With

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        MyBase.OnRowValidated(e)
    End Sub
#End Region

#Region "Custom painting"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim selGridRow As BudgetItemReferenceGridRow = Grid(intRowIndex)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds

        If intRowIndex = IndexTotal And intColumnIndex >= 0 Then
            CellPainting_Totals(intColumnIndex, selGridRow, rCell, CellGraphics)
            e.Handled = True

        End If
    End Sub

    Private Sub CellPainting_Totals(ByVal intColumnIndex As Integer, ByVal selGridRow As BudgetItemReferenceGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim strColumnName As String = Columns(intColumnIndex).Name
        Dim fntTotal As New Font(Me.Font, FontStyle.Bold)
        Dim rText As Rectangle = GetTextRectangle(rCell)

        CellGraphics.FillRectangle(Brushes.DarkGreen, rCell)

        Select Case strColumnName
            Case "ReferenceBudgetItem"
                CellGraphics.DrawString(LANG_TotalBudget, fntTotal, Brushes.White, rText)
            Case "TotalCost"
                Dim objFormat As New StringFormat
                Dim strTotalCost As String = String.Format("{0} {1}", CurrentResource.TotalCostAmount.ToString("N2"), CurrentLogFrame.CurrencyCode)

                objFormat.Alignment = StringAlignment.Far

                CellGraphics.DrawString(strTotalCost, fntTotal, Brushes.White, rText, objFormat)
        End Select

    End Sub

    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim RowGraphics As Graphics = e.Graphics

        'draw focus rectangle
        DrawSelectionRectangle(RowGraphics)
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


            graph.DrawRectangle(penSelection, .Rectangle)
        End With
    End Sub

    Public Sub SetSelectionRectangle()
        Dim rTopLeftCell As New Rectangle
        Dim rBottomRightCell As New Rectangle
        Dim FirstRowIndexSelection As Integer

        If SelectedCells.Count > 0 Then
            SetSelectionRectangleGridArea()

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
        End If
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
