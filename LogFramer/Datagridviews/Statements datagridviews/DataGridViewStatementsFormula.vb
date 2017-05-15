Imports System.Text.RegularExpressions

Public Class DataGridViewStatementsFormula
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New StatementFormulaGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents SelectionRectangle As New SelectionRectangle

    Public Event FormulaUpdated()
    Public Event UnitUpdated()

    Private objCurrentIndicator As Indicator
    Private colStatement As New RichTextColumnLogframe
    Private colNrDecimals As New DataGridViewTextBoxColumn
    Private colUnit As New StructuredComboBoxColumnNumberUnit
    Private colOpMin As New DataGridViewComboBoxColumn
    Private colOpMax As New DataGridViewComboBoxColumn
    Private colMinValue As New DataGridViewTextBoxColumn
    Private colMaxValue As New DataGridViewTextBoxColumn
    Private objCellLocation As New Point

    Private MeasureUnits As List(Of StructuredComboBoxItem)
    Private RemovedRows As New List(Of String)
    Private strFind As String

    Private RowModifyIndex As Integer = -1
    Private EditRow As StatementFormulaGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private SelectedGridRows As New List(Of StatementFormulaGridRow)
    Private PasteRow As StatementFormulaGridRow

    Private ClickPoint As Point
    Private pStartInsertLine, pEndInsertLine As New Point
    Private pStartInsertLineOld, pEndInsertLineOld As New Point
    Private boolColumnWidthChanged As Boolean

    Private Const CONST_MinRowHeight As Integer = 24

    Public Sub New()

    End Sub

    Public Sub New(ByVal currentindicator As Indicator)
        Me.CurrentIndicator = currentindicator

        VirtualMode = True
        AutoGenerateColumns = False
        BackgroundColor = Color.White
        ShowCellToolTips = False

        DefaultCellStyle.Padding = New Padding(2)
        With ColumnHeadersDefaultCellStyle
            .Font = New Font(DefaultFont, FontStyle.Bold)
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .WrapMode = DataGridViewTriState.True
        End With
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders

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
                If Me.CurrentCell.ColumnIndex = 0 Then
                    selObject = CurrentIndicator.Statements(intIndex)
                Else
                    selObject = CurrentIndicator.ResponseClasses(intIndex)
                End If
            End If
            Return selObject
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub LoadColumns()
        Columns.Clear()
        MeasureUnits = LoadMeasureUnits()

        With colStatement
            .Name = "Statement"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Statement
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .MinimumWidth = 100
        End With

        With colNrDecimals
            .Name = "NrDecimals"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Decimals
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        End With

        With colUnit
            .Name = "Unit"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Unit
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .DataSource = MeasureUnits
            .ValueType = GetType(String)
            .DisplayMember = "Description"
            .ValueMember = "Unit"
        End With

        With colOpMin
            .Name = "OpMin"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = String.Empty
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .Items.AddRange({String.Empty, CONST_LargerThan, CONST_LargerThanOrEqual})
        End With

        With colMinValue
            .Name = "MinValue"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_MinValue
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        End With

        With colOpMax
            .Name = "OpMax"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = String.Empty
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .Items.AddRange({String.Empty, CONST_SmallerThan, CONST_SmallerThanOrEqual})
        End With

        With colMaxValue
            .Name = "MaxValue"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_MaxValue
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        End With

        Columns.Add(colStatement)
        Columns.Add(colNrDecimals)
        Columns.Add(colUnit)
        Columns.Add(colOpMin)
        Columns.Add(colMinValue)
        Columns.Add(colOpMax)
        Columns.Add(colMaxValue)

        If My.Settings.setShowAdvancedIndicatorOptions = False Then
            colOpMin.Visible = False
            colMinValue.Visible = False
            colOpMax.Visible = False
            colMaxValue.Visible = False
        End If
        Reload()
    End Sub

    Private Sub SetRowHeadersText()
        Dim strFirst As String = String.Empty
        Dim strSecond As String = String.Empty
        Dim intCounter, intLeadCounter As Integer
        Dim strCounter As String
        Dim strFormula As String = String.Empty

        Me.RowCount = Me.Grid.Count + 1

        For i = 0 To Me.Grid.Count - 1
            If intCounter > 0 And Decimal.Remainder(intCounter, 26) = 0 Then
                intCounter = 0

                strFirst = Chr(intLeadCounter + 65)
                intLeadCounter += 1
            End If

            strSecond = Chr(intCounter + 65)
            strCounter = strFirst & strSecond

            Rows(i).HeaderCell.Value = strCounter
            intCounter += 1
        Next

        If CurrentIndicator.ValuesDetail.Formula = String.Empty Then
            For i = 0 To RowCount - 1
                strFormula &= Rows(i).HeaderCell.Value & "+"
            Next
            strFormula = strFormula.Remove(strFormula.Length - 2)
            CurrentIndicator.ValuesDetail.Formula = strFormula

        End If
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

        'Rows.Clear()
        LoadStatements()
        SetRowHeadersText()

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
                    Me.Grid.Add(New StatementFormulaGridRow(.Statements(intIndex)))
                Next
            End With
        End If
    End Sub
#End Region

#Region "Cell images"
    Public Sub ResetAllImages()
        For Each selRow As StatementFormulaGridRow In Me.Grid
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As StatementFormulaGridRow)
        selRow.StatementImageDirty = True
        selRow.StatementHeight = 0
    End Sub

    Public Sub ReloadImages()
        For Each selRow As StatementFormulaGridRow In Me.Grid
            ReloadImage_Statement(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImage_Statement(ByVal selRow As StatementFormulaGridRow)
        Dim intColumnWidth As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim boolSelected As Boolean

        If selStatement Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.StatementImageDirty = True And String.IsNullOrEmpty(selStatement.FirstLabel) = False Then
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
        Dim selGridRow As StatementFormulaGridRow = Me.Grid(RowIndex)

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
        Dim selGridRow As StatementFormulaGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing AndAlso selGridRow.StatementHeight > intRowHeight Then intRowHeight = selGridRow.StatementHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
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

    Protected Overrides Sub OnEditingControlShowing(ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        MyBase.OnEditingControlShowing(e)

        'set RichTextEditingControl
        Me.RichTextEditingControl = TryCast(e.Control, RichTextEditingControlLogframe)

        Dim selItem As Object = GetCurrentItem()

        If selItem IsNot Nothing Then
            UndoRedo.UndoBuffer_Initialise(selItem, "FirstLabel")
        End If
    End Sub

    Protected Overrides Sub OnCellValidating(e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)
        MyBase.OnCellValidating(e)

        Dim strColName As String = Columns(e.ColumnIndex).Name

        If strColName = "Unit" Then
            Dim strText As String = e.FormattedValue
            VerifyIfItemExists(Me(e.ColumnIndex, e.RowIndex))
        End If
    End Sub

    Protected Overrides Sub OnColumnWidthChanged(ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
        If Me.IsCurrentCellInEditMode = False Then
            MyBase.OnColumnWidthChanged(e)

            boolColumnWidthChanged = True
        End If
    End Sub

    Private Sub VerifyIfItemExists(ByVal CurrentCell As DataGridViewCell)
        Dim objComboBoxItem As Object = Nothing
        Dim objEdit As StructuredComboBoxEditingControlNumberUnit = CType(Me.EditingControl, StructuredComboBoxEditingControlNumberUnit)

        If objEdit Is Nothing Then Exit Sub

        strFind = objEdit.Text
        objComboBoxItem = MeasureUnits.Find(AddressOf FindQuestionTypeName)

        If objComboBoxItem Is Nothing Then
            Dim NewItem As New StructuredComboBoxItem(strFind, False, False, , strFind)
            MeasureUnitsUser.Add(NewItem)
            MeasureUnits = LoadMeasureUnits()

            colUnit.DataSource = MeasureUnits
            objEdit.Text = strFind
        End If
    End Sub

    Private Function FindQuestionTypeName(ByVal selItem As StructuredComboBoxItem) As Boolean
        If selItem.Description = strFind Then Return True Else Return False
    End Function
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selStatement As Statement, ByVal strPropertyName As String, ByVal intIndex As Integer, ByVal intLength As Integer)
        Dim intColIndex, intRowIndex As Integer

        'If boolSelectValue = True Then intColIndex = 1
        intRowIndex = CurrentIndicator.Statements.IndexOf(selStatement)

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)

            If strPropertyName = "FirstLabel" Then
                BeginEdit(False)

                Dim ctlEdit As RichTextEditingControlLogframe = CType(EditingControl, RichTextEditingControlLogframe)
                ctlEdit.Select(intIndex, intLength)
            Else
                BeginEdit(True)
            End If
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

        SetFocusOnStatement(selObject)

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
        SetFocusOnStatement(selItem)
    End Sub

    Private Sub SetFocusOnStatement(ByVal selStatement As Statement)
        Dim selGridRow As StatementFormulaGridRow
        Dim intColIndex As Integer
        Dim intRowIndex As Integer

        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.Statement Is selStatement Then
                Exit For
            End If
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            CurrentCell.Selected = True
        End If
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim intRowIndex, intColumnIndex As Integer
        Dim selStatement As Statement

        If Me.IsCurrentCellInEditMode = False Then
            RemovedRows.Clear()

            intRowIndex = CurrentCell.RowIndex
            intColumnIndex = CurrentCell.ColumnIndex

            'copy cells to delete
            For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
                SelectedGridRows.Add(Me.Grid(i))
                If i < Grid.Count - 1 Then
                    RemovedRows.Add(Rows(i + 1).HeaderCell.Value)
                Else
                    RemovedRows.Add(Rows(i).HeaderCell.Value)
                End If
            Next

            For Each selGridRow As StatementFormulaGridRow In SelectedGridRows
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

            RemoveFromFormula()
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
            Dim selRow As StatementFormulaGridRow
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

        For i = 0 To PasteItems.Count - 1 'To 0 Step -1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Statement)
                    PasteItems_Statement(selItem, intPasteOption)
                Case Else
                    PasteItems_Other(selItem, intPasteOption)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(intColumnIndex, intRowIndex)
    End Sub

    Private Sub PasteItems_Statement(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selStatement As Statement = DirectCast(selItem.Item, Statement)

        Dim NewStatement As New Statement

        Using copier As New ObjectCopy
            NewStatement = copier.CopyObject(selStatement)
        End Using

        PasteStatement(NewStatement)
    End Sub

    Private Sub PasteStatement(ByVal NewStatement As Statement)
        Dim PasteRowStatement As Statement = PasteRow.Statement
        Dim intIndex As Integer = CurrentIndicator.Statements.IndexOf(PasteRowStatement)

        If intIndex = -1 Then intIndex = CurrentIndicator.Statements.Count

        CurrentIndicator.Statements.Insert(intIndex, NewStatement)

        PasteRow.Statement = NewStatement

        UndoRedo.ItemPasted(NewStatement, CurrentIndicator.Statements)
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
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)
        Dim RowTmp As StatementFormulaGridRow = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name

        If e.RowIndex = RowCount - 1 Then
            Return
        End If
        ' Store a reference to the grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), StatementFormulaGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "Statement"
                e.Value = RowTmp.FirstLabel
            Case "NrDecimals"
                e.Value = RowTmp.NrDecimals
            Case "Unit"
                e.Value = RowTmp.Unit
            Case "OpMin"
                e.Value = RowTmp.OpMin
            Case "MinValue"
                e.Value = RowTmp.MinValue
            Case "OpMax"
                e.Value = RowTmp.OpMax
            Case "MaxValue"
                e.Value = RowTmp.MaxValue
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As StatementFormulaGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim dblCellValue As Double
        Dim selRowIndex As Integer = e.RowIndex

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As StatementFormulaGridRow = CType(Me.Grid(selRowIndex), StatementFormulaGridRow)

                Me.EditRow = New StatementFormulaGridRow
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

        If RowTmp.Statement Is Nothing Then
            Dim NewStatement As New Statement
            CurrentIndicator.Statements.Add(NewStatement)

            UndoRedo.ItemInserted(NewStatement, CurrentIndicator.Statements)

            RowTmp.Statement = NewStatement

            Dim ptCurrentCell As New Point(CurrentCell.ColumnIndex, CurrentCell.RowIndex)
            CurrentCell = Nothing
            CurrentCell = Me(ptCurrentCell.X, ptCurrentCell.Y)
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
            Case Else
                If e.Value = Nothing Then Exit Sub
        End Select

        Select Case strColName
            Case "Statement"
                RowTmp.FirstLabel = strCellValue
                RowTmp.StatementImageDirty = True
            Case "NrDecimals"
                dblCellValue = ParseDouble(e.Value)

                UndoRedo.UndoBuffer_Initialise(RowTmp.Statement.ValuesDetail, "NrDecimals", RowTmp.Statement.ValuesDetail.NrDecimals)

                RowTmp.NrDecimals = dblCellValue
                UndoRedo.ValueChanged(dblCellValue)
            Case "Unit"
                strCellValue = e.Value.ToString

                UndoRedo.UndoBuffer_Initialise(RowTmp.Statement.ValuesDetail, "Unit", RowTmp.Statement.ValuesDetail.Unit)
                RowTmp.Unit = strCellValue
                UndoRedo.OptionChanged(strCellValue)

                RaiseEvent UnitUpdated()
            Case "OpMin"
                strCellValue = e.Value

                UndoRedo.UndoBuffer_Initialise(RowTmp.Statement.ValuesDetail.ValueRange, "OpMin", RowTmp.Statement.ValuesDetail.ValueRange.OpMin)
                RowTmp.OpMin = strCellValue
                UndoRedo.OptionChanged(strCellValue)
            Case "MinValue"
                dblCellValue = ParseDouble(e.Value)

                UndoRedo.UndoBuffer_Initialise(RowTmp.Statement.ValuesDetail.ValueRange, "MinValue", RowTmp.Statement.ValuesDetail.ValueRange.MinValue)

                RowTmp.MinValue = dblCellValue
                UndoRedo.ValueChanged(dblCellValue)
            Case "OpMax"
                strCellValue = e.Value

                UndoRedo.UndoBuffer_Initialise(RowTmp.Statement.ValuesDetail.ValueRange, "OpMax", RowTmp.Statement.ValuesDetail.ValueRange.OpMax)
                RowTmp.OpMax = strCellValue
                UndoRedo.OptionChanged(strCellValue)
            Case "MaxValue"
                dblCellValue = ParseDouble(e.Value)

                UndoRedo.UndoBuffer_Initialise(RowTmp.Statement.ValuesDetail.ValueRange, "MaxValue", RowTmp.Statement.ValuesDetail.ValueRange.MaxValue)

                RowTmp.MaxValue = dblCellValue
                UndoRedo.ValueChanged(dblCellValue)
        End Select
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New StatementFormulaGridRow
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
        Me.EditRow = New StatementFormulaGridRow()
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

            AddToFormula()
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

    Private Sub AddToFormula()
        'when a row is added, include it in the indicator's formula
        Dim strRowHeader As String = Rows(Grid.Count - 1).HeaderCell.Value

        If String.IsNullOrEmpty(strRowHeader) = False Then
            With CurrentIndicator.ValuesDetail
                If .Formula.EndsWith(strRowHeader) = False Then
                    .Formula &= "+" & strRowHeader

                    RaiseEvent FormulaUpdated()
                End If
            End With
        End If
    End Sub

    Private Sub RemoveFromFormula()
        Dim strRowCode As String
        Dim pattern As String
        Dim strFormula As String = CurrentIndicator.ValuesDetail.Formula
        'Dim match As Match

        If RemovedRows.Count > 0 Then
            For i = 0 To RemovedRows.Count - 1
                strRowCode = RemovedRows(i)
                pattern = String.Format("[+\-\*\/]\s*{0}", strRowCode)
                'Match = Regex.Match(CurrentIndicator.ValuesDetail.Formula, pattern)

                'If Match.Success Then
                strFormula = Regex.Replace(strFormula, pattern, String.Empty)
                'End If
            Next
        End If

        If strFormula <> CurrentIndicator.ValuesDetail.Formula Then
            CurrentIndicator.ValuesDetail.Formula = strFormula
            RaiseEvent FormulaUpdated()
        End If

    End Sub
#End Region 'Virtual mode

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If Me.Grid.Count > 0 Then
            If boolColumnWidthChanged = True Then
                Reload()
                boolColumnWidthChanged = False
            End If

            Dim CellGraphics As Graphics = e.Graphics
            Dim rCell As Rectangle = e.CellBounds
            Dim intRowIndex As Integer = e.RowIndex
            Dim selGridRow As StatementFormulaGridRow = Grid(intRowIndex)

            If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
            Else
                e.PaintBackground(rCell, False)
            End If

            If intRowIndex >= 0 And e.ColumnIndex = 0 Then
                If selGridRow IsNot Nothing AndAlso selGridRow.StatementImage IsNot Nothing Then
                    Dim intRowHeight As Integer = selGridRow.StatementImage.Height
                    If Rows(intRowIndex).Height < intRowHeight Then Rows(intRowIndex).Height = intRowHeight
                End If

                CellPainting_RichText(selGridRow, rCell, CellGraphics)
                e.Paint(rCell, DataGridViewPaintParts.Border)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As StatementFormulaGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow IsNot Nothing AndAlso selGridRow.StatementImage IsNot Nothing Then
            rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StatementImage.Width, selGridRow.StatementImage.Height)
            CellGraphics.DrawImage(selGridRow.StatementImage, rImage)
        End If
    End Sub

    Private Sub CellPainting_NormalText(ByVal strText As String, ByVal boolKeyMoment As Boolean, ByVal boolAlignRight As Boolean, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim fntText As Font = DefaultCellStyle.Font
        Dim brText As New SolidBrush(Color.Black)
        Dim objFormat As New StringFormat

        If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
            brText = New SolidBrush(SystemColors.HighlightText)
        End If

        rCell.X += CONST_Padding
        rCell.Y += CONST_Padding
        rCell.Width -= CONST_HorizontalPadding
        rCell.Height -= CONST_VerticalPadding

        If boolAlignRight = True Then
            objFormat.Alignment = StringAlignment.Far
        Else
            objFormat.Alignment = StringAlignment.Near
        End If

        CellGraphics.DrawString(strText, fntText, brText, rCell, objFormat)
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

        Select Case e.ColumnIndex
            Case 0
                If ClickPoint.IsEmpty = False Then
                    Me.BeginEdit(False)

                    Dim rCell As Rectangle = Me.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False)
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
            Dim selGridRow As StatementFormulaGridRow = Me.Grid(hit.RowIndex)

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
