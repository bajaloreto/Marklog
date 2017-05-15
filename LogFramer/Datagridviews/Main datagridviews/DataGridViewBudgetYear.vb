Public Class DataGridViewBudgetYear
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New BudgetGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents SelectionRectangle As New SelectionRectangle
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe

    Public Event Reloaded(ByVal sender As Object)
    Public Event TotalCostChanged()
    Public Event BudgetItemSelected(ByVal sender As Object, ByVal e As BudgetItemSelectedEventArgs)
    Public Event BudgetItemAdded(ByVal sender As Object, ByVal e As BudgetItemAddedEventArgs)
    Public Event BudgetItemRemoved(ByVal sender As Object, ByVal e As BudgetItemRemovedEventArgs)
    Public Event BudgetItemMoved(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
    Public Event BudgetItemInsertParent(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
    Public Event BudgetItemInsertChild(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
    Public Event BudgetItemLevelDown(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
    Public Event BudgetItemLevelUp(ByVal sender As Object, ByVal e As BudgetItemMovedEventArgs)
    Public Event BudgetUpdateParentTotalsNeeded(ByVal sender As Object, ByVal e As UpdateParentTotalsEventArgs)
    Public Event BudgetUpdateChildTotalsNeeded(ByVal sender As Object, ByVal e As UpdateChildTotalsEventArgs)
    Public Event NoTextSelected()

    Private objBudgetYear As BudgetYear
    Private boolHideEmptyCells As Boolean
    Private boolShowDurationColumns As Boolean = True
    Private boolShowLocalCurrencyColumn As Boolean
    Private boolTotalBudget As Boolean
    Private intBudgetYearIndex As Integer
    Private objCellLocation As New Point

    Private colSortNumber As New DataGridViewTextBoxColumn
    Private colRTF As New RichTextColumnLogframe
    Private colDuration As New DataGridViewTextBoxColumn
    Private colDurationUnit As New DataGridViewComboBoxColumn
    Private colNumber As New DataGridViewTextBoxColumn
    Private colNumberUnit As New StructuredComboBoxColumnNumberUnit
    Private colUnitCost As New DataGridViewTextBoxColumn
    Private colUnitCostCurrency As New DataGridViewComboBoxColumn
    Private colTotalLocalCost As New DataGridViewTextBoxColumn
    Private colTotalCost As New DataGridViewTextBoxColumn

    Private RowModifyIndex As Integer = -1
    Private EditRow As BudgetGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private SelectedGridRows As New List(Of BudgetGridRow)
    Private PasteRow As BudgetGridRow
    Private DragBudgetItem As BudgetItem
    Private sngScale As Single
    Private ClickPoint As Point
    Private pStartInsertLine, pEndInsertLine As New Point
    Private pStartInsertLineOld, pEndInsertLineOld As New Point

#Region "Enumerations"
    Public Enum TextSelectionValues
        SelectNothing = 0
        SelectAll = 1
        SelectBudgetItems = 8
    End Enum
#End Region

#Region "Properties"
    Public Property BudgetYear As BudgetYear
        Get
            Return objBudgetYear
        End Get
        Set(ByVal value As BudgetYear)
            objBudgetYear = value
        End Set
    End Property

    Public Property BudgetYearIndex As Integer
        Get
            Return intBudgetYearIndex
        End Get
        Set(ByVal value As Integer)
            intBudgetYearIndex = value
        End Set
    End Property

    Public Property TotalBudget As Boolean
        Get
            Return boolTotalBudget
        End Get
        Set(ByVal value As Boolean)
            boolTotalBudget = value
        End Set
    End Property

    Public Property ShowDurationColumns As Boolean
        Get
            Return boolShowDurationColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowDurationColumns = value
        End Set
    End Property

    Public Property ShowLocalCurrencyColumn As Boolean
        Get
            Return boolShowLocalCurrencyColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowLocalCurrencyColumn = value
        End Set
    End Property

    Public Property HideEmptyCells As Boolean
        Get
            Return boolHideEmptyCells
        End Get
        Set(ByVal value As Boolean)
            boolHideEmptyCells = value
        End Set
    End Property

    Public ReadOnly Property IndexTotal As Integer
        Get
            Return Me.Grid.Count
        End Get
    End Property

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim selPlanningRow As BudgetGridRow = Me.Grid(Me.CurrentCell.RowIndex)
                If selPlanningRow.BudgetItem IsNot Nothing Then
                    selObject = selPlanningRow.BudgetItem
                End If
            End If
            Return selObject
        End Get
    End Property
#End Region 'Properties

#Region "Initialise"
    Public Sub New(ByVal budgetyear As BudgetYear, ByVal budgetyearindex As Integer, ByVal totalbudget As Boolean)
        Me.BudgetYear = budgetyear
        Me.BudgetYearIndex = budgetyearindex
        Me.TotalBudget = totalbudget

        DoubleBuffered = True
        VirtualMode = True
        AutoGenerateColumns = False
        AllowUserToAddRows = False

        ShowCellToolTips = False
        MultiSelect = True
        RowHeadersVisible = False
        BackgroundColor = Color.White
        DefaultCellStyle.Padding = New Padding(2)

        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = False

        RowHeadersVisible = True
        RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True

        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        With ColumnHeadersDefaultCellStyle
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Font = New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)
            .WrapMode = DataGridViewTriState.True
        End With

    End Sub

    Public Sub LoadColumns()
        Me.Columns.Clear()

        With colSortNumber
            .Name = "SortNumber"
            .HeaderText = LANG_Number
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .Frozen = True
        End With
        Me.Columns.Add(colSortNumber)

        With colRTF
            .Name = "RTF"
            .HeaderText = LANG_Description
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .MinimumWidth = 250
            .FillWeight = 200
        End With
        Me.Columns.Add(colRTF)

        With colDuration
            .Name = "Duration"
            .HeaderText = LANG_Duration
            .MinimumWidth = 80
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .CellTemplate.Style.Alignment = DataGridViewContentAlignment.TopRight
        End With
        Columns.Add(colDuration)

        With colDurationUnit
            .Name = "DurationUnit"
            .HeaderText = LANG_Unit
            .MinimumWidth = 120
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .AutoComplete = True
            .DataSource = CurrentLogFrame.GetDurationUnitsList
        End With
        Columns.Add(colDurationUnit)

        With colNumber
            .Name = "Number"
            .HeaderText = LANG_Quantity
            .MinimumWidth = 80
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .CellTemplate.Style.Alignment = DataGridViewContentAlignment.TopRight
        End With
        Columns.Add(colNumber)

        With colNumberUnit
            .Name = "NumberUnit"
            .HeaderText = LANG_Unit
            .MinimumWidth = 120
            .AutoComplete = True
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .Items.AddRange(CurrentLogFrame.GetNumberUnitsList.ToArray)
            .ValueType = GetType(String)
            .DisplayMember = "Description"
            .ValueMember = "Unit"
        End With
        Columns.Add(colNumberUnit)

        With colUnitCost
            .Name = "UnitCost"
            .HeaderText = LANG_UnitCost
            .MinimumWidth = 120
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .CellTemplate.Style.Alignment = DataGridViewContentAlignment.TopRight
        End With
        Columns.Add(colUnitCost)

        Dim lstCurrencyCodes As List(Of IdValuePair) = CurrentLogFrame.Budget.LoadUsedCurrencyCodesList()
        With colUnitCostCurrency
            .Name = "CurrencyCode"
            .HeaderText = LANG_Currency
            .MinimumWidth = 120
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .AutoComplete = True
            .DataSource = lstCurrencyCodes
            .ValueMember = "Id"
            .DisplayMember = "Value"
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .DropDownWidth = 200
        End With
        Me.Columns.Add(colUnitCostCurrency)

        With colTotalLocalCost
            .Name = "TotalLocalCost"
            .HeaderText = LANG_TotalLocalCost
            .MinimumWidth = 120
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .CellTemplate.Style.Alignment = DataGridViewContentAlignment.TopRight
            .ReadOnly = True
        End With
        Columns.Add(colTotalLocalCost)

        With colTotalCost
            .Name = "TotalCost"
            .HeaderText = LANG_TotalCost
            .MinimumWidth = 120
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .CellTemplate.Style.Alignment = DataGridViewContentAlignment.TopRight
            .ReadOnly = True
        End With
        Columns.Add(colTotalCost)

        ShowColumns()
    End Sub

    Public Sub ShowColumns()
        If TotalBudget Then
            colDuration.Visible = False
            colDurationUnit.Visible = False
            colUnitCost.Visible = False
            colUnitCostCurrency.Visible = False
            colTotalLocalCost.Visible = False
        Else
            colDuration.Visible = Me.ShowDurationColumns
            colDurationUnit.Visible = Me.ShowDurationColumns
            colUnitCostCurrency.Visible = Me.ShowLocalCurrencyColumn
            colTotalLocalCost.Visible = Me.ShowLocalCurrencyColumn
        End If

        Invalidate()
    End Sub

    Private Sub ReadOnlyRows()
        Dim intRowIndex As Integer

        For Each selGridRow As BudgetGridRow In Me.Grid

            If selGridRow.BudgetItem IsNot Nothing AndAlso selGridRow.BudgetItem.BudgetItems.Count > 0 Then
                For i = 2 To Columns.Count - 1
                    Me(i, intRowIndex).ReadOnly = True
                Next
            Else
                Rows(intRowIndex).ReadOnly = False
            End If
            intRowIndex += 1
        Next
    End Sub

    Private Sub SetRowHeadersText()
        Dim intCounter As Integer
        Dim strCounter As String

        For i = 0 To Me.Grid.Count - 1
            intCounter += 1

            strCounter = intCounter.ToString

            Rows(i).HeaderCell.Value = strCounter
        Next
    End Sub

    Public Sub Reload()
        'remember current cell location
        objCellLocation = CurrentCellAddress
        If objCellLocation.X < 1 Then objCellLocation.X = 1
        If objCellLocation.Y < 0 Then objCellLocation.Y = 0

        '(re)load target columns and rows
        Me.SuspendLayout()
        LoadItems()

        Me.RowCount = Me.Grid.Count + 1
        Me.RowModifyIndex = -1
        AutoSizeSortColumn()
        ReadOnlyRows()

        SetRowHeadersText()

        ResetAllImages()
        ReloadImages()

        'set current cell location to what it was before
        Me.Invalidate()
        Me.ResumeLayout()
        CurrentCell = Me(objCellLocation.X, objCellLocation.Y)

        RaiseEvent Reloaded(Me)
    End Sub

    Private Sub AutoSizeSortColumn()
        Dim intWidth As Integer = 0

        colSortNumber.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        intWidth = colSortNumber.Width
        colSortNumber.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colSortNumber.Width = intWidth
    End Sub
#End Region

#Region "Budget items"
    Public Sub LoadItems()
        Me.Grid.Clear()
        EditRow = Nothing

        If CurrentBudget.MultiYearBudget = True Then
            CurrentBudget.UpdateMultiYearBudget()
        Else
            CurrentBudget.UpdateMultiYearBudget_CalculatedAmounts()
        End If

        LoadItems_BudgetItems(BudgetYear.BudgetItems, 0)
        Me.RowCount = Me.Grid.Count + 1

        'UpdateCalculatedItems()

        Me.Invalidate()
        RaiseEvent TotalCostChanged()
    End Sub

    Private Sub LoadItems_BudgetItems(ByVal selBudgetItems As BudgetItems, ByVal intLevel As Integer, Optional ByVal strParentSort As String = "")
        Dim boolLastHasSubBudgetItems As Boolean
        Dim strBudgetSort As String = String.Empty
        Dim intIndex As Integer

        For Each selBudgetItem As BudgetItem In selBudgetItems
            Dim NewGridItem As New BudgetGridRow(selBudgetItem)
            boolLastHasSubBudgetItems = False
            intIndex = selBudgetItems.IndexOf(selBudgetItem)
            strBudgetSort = CurrentLogFrame.CreateSortNumber(intIndex, strParentSort)

            With NewGridItem
                .SortNumber = strBudgetSort
                .Indent = intLevel
                .BudgetItem = selBudgetItem
            End With

            Me.Grid.Add(NewGridItem)

            If selBudgetItem.BudgetItems.Count > 0 Then
                LoadItems_BudgetItems(selBudgetItem.BudgetItems, intLevel + 1, strBudgetSort)

                boolLastHasSubBudgetItems = True

                If selBudgetItem.Type <> BudgetItem.BudgetItemTypes.Category Then
                    selBudgetItem.Type = BudgetItem.BudgetItemTypes.Category
                End If
            End If
        Next
        If Me.HideEmptyCells = False And boolLastHasSubBudgetItems = False Then
            If BudgetYearIndex = 0 Then
                Me.Grid.Add(New BudgetGridRow())
            ElseIf selBudgetItems.Count > 0 Then
                If selBudgetItems(0).Type = BudgetItem.BudgetItemTypes.Normal Then
                    Me.Grid.Add(New BudgetGridRow())
                End If
            End If
        End If
    End Sub

    Private Sub UpdateCalculatedItems()
        For Each selGridRow As BudgetGridRow In Me.Grid

            If selGridRow.BudgetItem IsNot Nothing AndAlso selGridRow.BudgetItem.Type = BudgetItem.BudgetItemTypes.Ratio Then
                RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(selGridRow.BudgetItem))
            End If
        Next
    End Sub
#End Region

#Region "Cell images"
    Private Sub ResetAllImages()
        For Each selRow As BudgetGridRow In Me.Grid
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As BudgetGridRow)
        selRow.BudgetItemImageDirty = True
        selRow.BudgetItemHeight = 0
    End Sub

    Private Sub ReloadImages()
        For Each selRow As BudgetGridRow In Me.Grid
            ReloadImages_BudgetItem(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_BudgetItem(ByVal selRow As BudgetGridRow, Optional ByVal colForeGround As Color = Nothing, Optional ByVal colBackGround As Color = Nothing)
        Dim intColumnWidth As Integer
        Dim boolSelected As Boolean

        If selRow.BudgetItem Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.BudgetItemImageDirty = True Then
                intColumnWidth = colRTF.Width

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectBudgetItems Then
                    boolSelected = True
                Else
                    boolSelected = False
                End If

                If String.IsNullOrEmpty(selRow.BudgetItem.Text) Then
                    selRow.BudgetItemImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, BudgetItem.ItemName, selRow.SortNumber, boolSelected)
                Else
                    selRow.BudgetItemImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.RTF, boolSelected, selRow.Indent, HorizontalAlignment.Left, colForeGround, colBackGround)
                End If
                selRow.BudgetItemHeight = selRow.BudgetItemImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > CONST_MinRowHeight Then selRow.Height = intRowHeight Else selRow.Height = CONST_MinRowHeight
    End Sub

    Public Sub ResetRowHeights()
        For Each selRow As DataGridViewRow In Me.Rows
            SetRowHeight(selRow.Index)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selGridRow As BudgetGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing AndAlso selGridRow.BudgetItemHeight > intRowHeight Then _
            intRowHeight = selGridRow.BudgetItemHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
    Private Sub RichTextEditingControl_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs) Handles RichTextEditingControl.ContentsResized
        Dim intRequiredHeight As Integer = e.NewRectangle.Height + RichTextEditingControl.Margin.Vertical + SystemInformation.VerticalResizeBorderThickness
        If CurrentRow.Height < intRequiredHeight Then
            CurrentRow.Height = intRequiredHeight
            SetSelectionRectangle()
            InvalidateSelectionRectangle()
        End If
    End Sub

    Private Sub RichTextEditingControl_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextEditingControl.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim pClick As New Point(CurrentCell.ContentBounds.X, CurrentCell.ContentBounds.Y)
            pClick.X += e.X + CONST_Padding
            pClick.Y += e.Y + CONST_Padding

            Dim dragSize As Size = SystemInformation.DragSize
            DragBoxFromMouseDown = New Rectangle(New Point(Me.Location.X + e.X - (dragSize.Width / 2), _
                                                           Me.Location.Y + e.Y - (dragSize.Height / 2)), dragSize)
            If SelectionRectangle.Rectangle.Contains(pClick.X, pClick.Y) Then DragMultipleCells = True Else DragMultipleCells = False
        Else
            DragBoxFromMouseDown = Rectangle.Empty

            Dim selItem As Object = GetCurrentItem()

            If selItem IsNot Nothing Then
                UndoRedo.UndoBuffer_Initialise(selItem, "RTF")
            End If
        End If
    End Sub

    Private Sub RichTextEditingControl_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextEditingControl.TextChanged
        If RichTextEditingControl.Text.Length > 0 Then
            Me.RowModifyIndex = Me.CurrentRow.Index
        Else
            Me.RowModifyIndex = -1
        End If

        UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.TextChanged)
    End Sub

    Private Sub RichTextEditingControl_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextEditingControl.Validated
        UndoRedo.TextChanged(RichTextEditingControl.Rtf)
    End Sub

    Protected Overrides Sub OnEditingControlShowing(ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        MyBase.OnEditingControlShowing(e)

        Dim selItem As Object = GetCurrentItem()

        If Me.TextSelectionIndex <> TextSelectionValues.SelectNothing Then
            Me.TextSelectionIndex = TextSelectionValues.SelectNothing
            ResetAllImages()
            ReloadImages()
            Invalidate()
            RaiseEvent NoTextSelected()
        End If

        Me.RichTextEditingControl = TryCast(e.Control, RichTextEditingControlLogframe)

        If selItem IsNot Nothing Then
            UndoRedo.UndoBuffer_Initialise(selItem, "RTF")
        End If
    End Sub

    Private Sub DataGridViewBudgetYear_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles Me.DataError
        Debug.Print(e.ColumnIndex.ToString & vbTab & e.Exception.Message.ToString)
        e.Cancel = True
    End Sub

    Private Sub DataGridViewBudgetYear_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Me.Scroll
        Invalidate()
    End Sub
#End Region

End Class

Public Class BudgetItemSelectedEventArgs
    Inherits EventArgs

    Public Property BudgetItem As BudgetItem

    Public Sub New(ByVal objBudgetItem As BudgetItem)
        MyBase.New()

        Me.BudgetItem = objBudgetItem
    End Sub
End Class

Public Class BudgetItemMovedEventArgs
    Inherits EventArgs

    Public Property BudgetItem As BudgetItem

    Public Sub New(ByVal objBudgetItem As BudgetItem)
        MyBase.New()

        Me.BudgetItem = objBudgetItem
    End Sub
End Class
