Public Class DataGridViewStatementsMaxDiff
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New StatementGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents SelectionRectangle As New SelectionRectangle

    Private objCurrentIndicator As Indicator
    Private colStatement As New RichTextColumnLogframe
    Private objCellLocation As New Point

    Private RowModifyIndex As Integer = -1
    Private EditRow As StatementGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private SelectedGridRows As New List(Of StatementGridRow)
    Private PasteRow As StatementGridRow
    Private boolColumnWidthChanged As Boolean

    Private ClickPoint As Point
    Private pStartInsertLine, pEndInsertLine As New Point
    Private pStartInsertLineOld, pEndInsertLineOld As New Point

    Private Const CONST_MinRowHeight As Integer = 24

    Public Sub New()

    End Sub

    Public Sub New(ByVal currentindicator As Indicator)
        Me.CurrentIndicator = currentindicator

        VirtualMode = True
        AutoGenerateColumns = False
        AllowUserToAddRows = False
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = False

        ShowCellToolTips = False
        BackgroundColor = Color.White
        DefaultCellStyle.Padding = New Padding(2)
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        RowHeadersVisible = True
        RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        With RowHeadersDefaultCellStyle
            .WrapMode = DataGridViewTriState.True
        End With

        With ColumnHeadersDefaultCellStyle
            .Font = New Font(DefaultFont, FontStyle.Bold)
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .WrapMode = DataGridViewTriState.True
        End With
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

        LoadColumns()
    End Sub

#Region "Enumerations"
    Public Enum TextSelectionValues
        SelectNothing = 0
        SelectAll = 1
    End Enum
#End Region

#Region "Properties"
    Public Property CurrentIndicator As Indicator
        Get
            Return objCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            objCurrentIndicator = value
            objCurrentIndicator.CalculateTargetsWithFormula()
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim intIndex As Integer = CurrentCell.RowIndex

                selObject = CurrentIndicator.Statements(intIndex)
            End If
            Return selObject
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub LoadColumns()
        Columns.Clear()

        With colStatement
            .Name = "Statement"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Statement
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End With

        Columns.Add(colStatement)

        Reload()
    End Sub

    Public Sub Reload()
        'remember current cell location
        objCellLocation = CurrentCellAddress
        If objCellLocation.X < 0 Then
            If CurrentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                objCellLocation.X = 1
            Else
                objCellLocation.X = 0
            End If
        End If
        If objCellLocation.Y < 0 Then objCellLocation.Y = 0

        '(re)load responses
        Me.SuspendLayout()

        LoadStatements()
        Me.RowCount = Me.Grid.Count + 1
        Me.RowModifyIndex = -1

        ResetAllImages()
        ReloadImages()

        Me.Invalidate()
        Me.ResumeLayout()

        'set current cell location to what it was before
        If objCellLocation.X <= Me.ColumnCount - 1 And objCellLocation.Y <= Me.RowCount - 1 Then
            If Me(objCellLocation.X, objCellLocation.Y).Displayed = True Then _
                CurrentCell = Me(objCellLocation.X, objCellLocation.Y)
        End If
    End Sub

    Private Sub LoadStatements()
        Me.Grid.Clear()

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                For intIndex = 0 To .Statements.Count - 1
                    Me.Grid.Add(New StatementGridRow(.Statements(intIndex), Nothing))
                Next
            End With
        End If
    End Sub
#End Region

#Region "Cell images"
    Public Sub ResetAllImages()
        For Each selRow As StatementGridRow In Me.Grid
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As StatementGridRow)
        selRow.StatementImageDirty = True
        selRow.StatementHeight = 0
    End Sub

    Public Sub ReloadImages()
        For Each selRow As StatementGridRow In Me.Grid
            ReloadImage_Statement(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImage_Statement(ByVal selRow As StatementGridRow)
        Dim intColumnWidth As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim boolSelected As Boolean

        If selStatement Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.StatementImageDirty = True Then
                intColumnWidth = colStatement.Width

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                    boolSelected = True
                Else
                    boolSelected = False
                End If

                selRow.StatementImage = .RichTextWithPaddingToBitmap(intColumnWidth, selStatement.FirstLabel, boolSelected)
                selRow.StatementHeight = selRow.StatementImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim selGridRow As StatementGridRow = Me.Grid(RowIndex)

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
        Dim selGridRow As StatementGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing Then intRowHeight = selGridRow.StatementHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
    Protected Overrides Sub OnColumnWidthChanged(ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
        If Me.IsCurrentCellInEditMode = False Then
            MyBase.OnColumnWidthChanged(e)

            boolColumnWidthChanged = True
        End If
    End Sub

    Private Sub RichTextEditingControl_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs) Handles RichTextEditingControl.ContentsResized
        Dim intRequiredHeight As Integer = e.NewRectangle.Height + RichTextEditingControl.Margin.Vertical + SystemInformation.VerticalResizeBorderThickness
        If CurrentRow.Height < intRequiredHeight Then CurrentRow.Height = intRequiredHeight
    End Sub

    Private Sub RichTextEditingControl_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextEditingControl.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim pClick As New Point(CurrentCell.ContentBounds.X, CurrentCell.ContentBounds.Y)
            pClick.X += e.X + CONST_Padding
            pClick.Y += e.Y + CONST_Padding

            Dim dragSize As Size = SystemInformation.DragSize
            DragBoxFromMouseDown = New Rectangle(New Point(Me.Location.X + e.X - (dragSize.Width / 2), _
                                                           Me.Location.Y + e.Y - (dragSize.Height / 2)), dragSize)
        Else
            DragBoxFromMouseDown = Rectangle.Empty

            Dim selItem As Object = GetCurrentItem()

            If selItem IsNot Nothing Then
                With UndoRedo
                    UndoRedo.UndoBuffer_Initialise(selItem, "FirstLabel")
                End With
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

    Protected Overrides Sub OnCurrentCellChanged(ByVal e As System.EventArgs)
        MyBase.OnCurrentCellChanged(e)

        Dim selItem As Object = GetCurrentItem()

        UndoRedo.UndoBuffer_Initialise(selItem)
    End Sub

    Protected Overrides Sub OnEditingControlShowing(ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        MyBase.OnEditingControlShowing(e)

        'set RichTextEditingControl
        Me.RichTextEditingControl = TryCast(e.Control, RichTextEditingControlLogframe)

        Dim selItem As Object = GetCurrentItem()

        If selItem IsNot Nothing Then
            UndoRedo.UndoBuffer_Initialise(selItem, "FirstLabel")
        End If
    End Sub
#End Region

#Region "General methods"
    Private Sub SetFocusOnItem(ByVal selItem As Object)
        Dim selGridRow As StatementGridRow
        Dim intColIndex As Integer = 1
        Dim intRowIndex As Integer

        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.Statement Is selItem Then
                Exit For
            End If
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
        End If
    End Sub

    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex
        Me.ReloadImages()
    End Sub

    Public Sub AddNewItem()
        Dim intColIndex As Integer
        If CurrentCell IsNot Nothing Then
            intColIndex = CurrentCell.ColumnIndex
        End If

        CurrentCell = Me(intColIndex, RowCount - 1)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selObject As Object = Me.CurrentItem(False)
        Dim intIndex As Integer

        If CurrentCell IsNot Nothing Then
            intIndex = CurrentCell.RowIndex

            Dim NewStatement As New Statement
            CurrentIndicator.Statements.Insert(intIndex, NewStatement)

            UndoRedo.ItemInserted(NewStatement, CurrentIndicator.Statements)

            selObject = NewStatement
        End If

        Reload()

        SetFocusOnItem(selObject)

        Me.BeginEdit(False)
    End Sub

    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selItem As Object = Me.CurrentItem(False)
        If selItem Is Nothing Then Exit Sub

        Dim intIndex As Integer = CurrentCell.RowIndex
        Dim intOldIndex As Integer = intIndex
        Dim selStatement As Statement = CurrentIndicator.Statements(intIndex)

        If intDirection < 0 Then
            intIndex -= 1
            If intIndex < 0 Then intIndex = 0
        Else
            intIndex += 1
            If intIndex > CurrentIndicator.Statements.Count - 1 Then intIndex = CurrentIndicator.Statements.Count - 1
        End If

        CurrentIndicator.Statements.Remove(selStatement)
        CurrentIndicator.Statements.Insert(intIndex, selStatement)

        If intDirection < 0 Then
            UndoRedo.ItemMovedUp(selStatement, selStatement, CurrentIndicator.Statements, intOldIndex, CurrentIndicator.Statements)
        Else
            UndoRedo.ItemMovedDown(selStatement, selStatement, CurrentIndicator.Statements, intOldIndex, CurrentIndicator.Statements)
        End If

        Reload()
        SetFocusOnItem(selItem)
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim intRowIndex, intColumnIndex As Integer
        Dim selStatement As Statement

        If Me.IsCurrentCellInEditMode = False Then
            intRowIndex = CurrentCell.RowIndex
            intColumnIndex = CurrentCell.ColumnIndex

            'copy cells to delete
            For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
                SelectedGridRows.Add(Me.Grid(i))
            Next

            For Each selGridRow As StatementGridRow In SelectedGridRows
                selStatement = selGridRow.Statement

                If boolCut = False Then
                    UndoRedo.ItemRemoved(selStatement, CurrentIndicator.Statements)
                Else
                    UndoRedo.ItemCut(selStatement, CurrentIndicator.Statements)
                End If

                CurrentIndicator.Statements.Remove(selStatement)
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
#End Region 'remove items

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)
        CopyItems()

        RemoveItems(False)
    End Sub

    Public Overrides Sub CopyItems()
        With SelectionRectangle
            Dim selRow As StatementGridRow
            Dim strSort As String
            Dim CopyGroup As Date = Now()

            With SelectionRectangle
                For i = .FirstRowIndex To .LastRowIndex
                    selRow = Me.Grid(i)
                    If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Statement Then
                        strSort = Statement.ItemName
                        Dim NewItem As New ClipboardItem(CopyGroup, selRow.Statement, strSort)
                        ItemClipboard.Insert(0, NewItem)
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

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Statement)
                    PasteItems_Statement(selItem, intPasteOption)
                Case Else
                    PasteItems_Other(selItem, intPasteOption)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(PasteCell.ColumnIndex, PasteCell.RowIndex)
    End Sub

    Private Sub PasteItems_Statement(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selStatement As Statement = DirectCast(selItem.Item, Statement)

        Dim NewStatement As New Statement

        Using copier As New ObjectCopy
            NewStatement = copier.CopyObject(selStatement)
        End Using

        PasteStatement(NewStatement)
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

        If String.IsNullOrEmpty(strRtf) = False Then
            Dim NewStatement As New Statement(strRtf)

            PasteStatement(NewStatement)
        End If
    End Sub

    Private Sub PasteStatement(ByVal NewStatement As Statement)
        Dim PasteRowStatement As Statement = PasteRow.Statement
        Dim intIndex As Integer = CurrentIndicator.Statements.IndexOf(PasteRowStatement)

        If intIndex = -1 Then intIndex = CurrentIndicator.Statements.Count

        CurrentIndicator.Statements.Insert(intIndex, NewStatement)

        PasteRow.Statement = NewStatement

        UndoRedo.ItemPasted(NewStatement, CurrentIndicator.Statements)
    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)
        Dim RowTmp As StatementGridRow = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name

        If e.RowIndex = RowCount - 1 Then
            Return
        End If
        ' Store a reference to the grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), StatementGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "Statement"
                e.Value = RowTmp.FirstLabel
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As StatementGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim selRowIndex As Integer = e.RowIndex

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As StatementGridRow = CType(Me.Grid(selRowIndex), StatementGridRow)

                Me.EditRow = New StatementGridRow
                With EditRow
                    If CurrentGridRow.Statement IsNot Nothing Then .Statement = CurrentGridRow.Statement
                    .StatementHeight = CurrentGridRow.StatementHeight
                    .StatementImage = CurrentGridRow.StatementImage
                    .StatementImageDirty = CurrentGridRow.StatementImageDirty
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        ' Set the appropriate objRowEdit property to the cell value entered.
        strColName = Me.Columns(e.ColumnIndex).Name

        Select Case strColName
            Case "Statement"
                Dim ctlRTF As RichTextEditingControlLogframe = CType(Me.EditingControl, RichTextEditingControlLogframe)

                'necessary for pasting into non-active cell
                If ctlRTF Is Nothing Then
                    BeginEdit(False)
                    ctlRTF = CType(Me.EditingControl, RichTextEditingControlLogframe)
                End If

                strCellValue = TryCast(ctlRTF.Rtf, String)

                If RowTmp.Statement Is Nothing Then
                    Dim NewStatement As New Statement
                    CurrentIndicator.Statements.Add(NewStatement)

                    UndoRedo.ItemInserted(NewStatement, CurrentIndicator.Statements)

                    RowTmp.Statement = NewStatement
                End If

                RowTmp.FirstLabel = strCellValue
                RowTmp.StatementImageDirty = True
        End Select
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New StatementGridRow
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
        Me.EditRow = New StatementGridRow()
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
        ' grid row if there is one.
        If e.RowIndex >= Me.Grid.Count AndAlso e.RowIndex <> Me.Rows.Count - 1 Then
            ' Add the new planning grid row to grid.
            Me.Grid.Add(Me.EditRow)
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then

            ' Save the modified planning grid row in grid.
            Me.Grid(e.RowIndex) = Me.EditRow
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        MyBase.OnRowValidated(e)
    End Sub
#End Region 'Virtual mode

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If boolColumnWidthChanged = True Then
            Reload()
            boolColumnWidthChanged = False
        End If

        If Me.Grid.Count > 0 And e.ColumnIndex = 0 Then
            Dim CellGraphics As Graphics = e.Graphics
            Dim rCell As Rectangle = e.CellBounds
            Dim intIndex As Integer = e.RowIndex
            Dim selGridRow As StatementGridRow = Grid(intIndex)

            If Me.TextSelectionIndex = TextSelectionValues.SelectAll And e.RowIndex >= 0 Then
                CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
            Else
                e.PaintBackground(rCell, False)
            End If

            CellPainting_RichText(selGridRow, rCell, CellGraphics)
            e.Paint(rCell, DataGridViewPaintParts.Border)
            e.Handled = True
        End If
    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As StatementGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow IsNot Nothing Then
            If selGridRow.StatementImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StatementImage.Width, selGridRow.StatementImage.Height)
                CellGraphics.DrawImage(selGridRow.StatementImage, rImage)
            End If
        End If
    End Sub

    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim RowGraphics As Graphics = e.Graphics

        DrawSelectionRectangle(RowGraphics)
    End Sub
#End Region

#Region "Mouse actions"
    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)

        If CurrentCell.IsInEditMode = False Then
            InvalidateSelectionRectangle()
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Tab Then

            Dim x As Integer = CurrentCell.ColumnIndex
            Dim y As Integer = CurrentRow.Index

            If y = Me.RowCount - 1 Then
                CurrentCell = Me(x, y - 1)
            Else
                CurrentCell = Me(x, y + 1)
            End If

            CurrentCell = Me(x, y)

            CurrentCell.Selected = True

            InvalidateSelectionRectangle()
        End If
        MyBase.OnKeyUp(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If hit.ColumnIndex > 0 And hit.RowIndex >= 0 Then
            Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
            Dim selGridRow As StatementGridRow = Me.Grid(hit.RowIndex)

            If e.Button = MouseButtons.Right Then
                If hit.Type = DataGridViewHitTestType.Cell Then
                    ' Create a rectangle using the DragSize, with the mouse position being at the center of the rectangle.
                    OnMouseDown_SetDragRectangle(e.Location, selCell)
                End If
            Else
                'determine where exactly in the text the user clicked
                OnMouseDown_SetClickPoint(e.Location)

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
            'drag items
            OnMouseMove_DragItems(e.Location)

            InvalidateSelectionRectangle()
        End If
        MyBase.OnMouseMove(e)
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
        End If
        DragBoxFromMouseDown = Rectangle.Empty

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
                Dim intLastColIndex As Integer

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

    Protected Overrides Sub OnScroll(ByVal e As System.Windows.Forms.ScrollEventArgs)
        MyBase.OnScroll(e)

        InvalidateSelectionRectangle()
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

        SelectionRectangle.FirstRowIndex = CurrentCell.RowIndex
        SelectionRectangle.FirstColumnIndex = CurrentCell.ColumnIndex
        SelectionRectangle.LastRowIndex = CurrentCell.RowIndex
        SelectionRectangle.LastColumnIndex = CurrentCell.ColumnIndex

        If SelectedCells.Count > 1 Then
            For i = 0 To SelectedCells.Count - 1
                If SelectedCells(i).RowIndex < SelectionRectangle.FirstRowIndex Then SelectionRectangle.FirstRowIndex = SelectedCells(i).RowIndex
                If SelectedCells(i).RowIndex > SelectionRectangle.LastRowIndex Then SelectionRectangle.LastRowIndex = SelectedCells(i).RowIndex
                SelectionRectangle.FirstColumnIndex = 0
                SelectionRectangle.LastColumnIndex = 1
            Next
        End If
    End Sub
#End Region

End Class
