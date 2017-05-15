Public Class DataGridViewTargetsScaleLikert
    Inherits DataGridViewBaseClassRichText

    Friend Shadows WithEvents Grid As New TargetScalesLikertGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents EditBox As New NumericTextBox
    Friend WithEvents SelectionRectangle As New SelectionRectangle

    Public Event ScoreUpdated()
    Public Event StatementUpdated()

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlines As TargetDeadlines
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colStatement As New RichTextColumnLogframe

    Protected RowModifyIndex As Integer = -1
    Protected EditRow As TargetScalesLikertGridRow = Nothing
    Protected EditRowFlag As Integer = -1
    Protected rowScopeCommit As Boolean = True
    Protected UpdateScore As Boolean
    Private SelectedGridRows As New List(Of TargetScalesLikertGridRow)
    Private PasteRow As TargetScalesLikertGridRow

    Protected intTargetIndex As Integer = -1
    Protected strBaseline As String
    Protected strHeaderText() As String
    Protected strPopulationTarget, strPopulationScore As String
    Protected EditPopulationIndex As Integer

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
            If objCurrentIndicator IsNot Nothing Then objCurrentIndicator.CalculateTargetsWithFormula()
        End Set
    End Property

    Public Property TargetDeadlines As TargetDeadlines
        Get
            Return objTargetDeadlines
        End Get
        Set(ByVal value As TargetDeadlines)
            objTargetDeadlines = value
        End Set
    End Property

    Public Property TargetGroup As TargetGroup
        Get
            Return objTargetGroup
        End Get
        Set(ByVal value As TargetGroup)
            objTargetGroup = value
        End Set
    End Property

    Public Property TargetIndex As Integer
        Get
            Return intTargetIndex
        End Get
        Set(ByVal value As Integer)
            intTargetIndex = value
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

    Protected Overloads ReadOnly Property IndexPopulationTarget As Integer
        Get
            Return CurrentIndicator.Statements.Count + 1
        End Get
    End Property

    Protected Overloads ReadOnly Property IndexPopulationScore As Integer
        Get
            Return CurrentIndicator.Statements.Count + 2
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New()

    End Sub

    Public Sub New(ByVal currentindicator As Indicator, ByVal targetdeadlines As TargetDeadlines, ByVal targetgroup As TargetGroup, ByVal targetindex As Integer)
        'datagridview settings
        ChooseSettings()

        'edit box settings
        Me.EditBox.AcceptsReturn = True
        Me.EditBox.AcceptsTab = True

        'load values
        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlines = targetdeadlines
        Me.TargetGroup = targetgroup
        Me.TargetIndex = targetindex

        'SetBooleanValues()
        LoadColumns()
    End Sub

    Protected Sub ChooseSettings()
        DoubleBuffered = True
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

        With EditBox
            .AcceptsReturn = False
            .AcceptsTab = True
            .Multiline = False
            .WordWrap = False
            .TextAlign = HorizontalAlignment.Center
            .BorderStyle = Windows.Forms.BorderStyle.Fixed3D

        End With
    End Sub

    Public Sub SetBooleanValues()
        With Me.CurrentIndicator
            VerifyBooleanMatrix(.Baseline.BooleanValuesMatrix)

            For Each selTarget As Target In .Targets
                VerifyBooleanMatrix(selTarget.BooleanValuesMatrix)
            Next
        End With
    End Sub

    Protected Sub VerifyBooleanMatrix(ByVal objBooleanValuesMatrix As BooleanValuesMatrix)
        Dim intStatementsCount As Integer = CurrentIndicator.Statements.Count

        If objBooleanValuesMatrix.Count <> intStatementsCount Then
            objBooleanValuesMatrix.Clear()

            For i = 0 To intStatementsCount - 1
                Dim NewMatrixRow As New BooleanValuesMatrixRow

                objBooleanValuesMatrix.Add(NewMatrixRow)
            Next
        End If

        For Each selMatrixRow As BooleanValuesMatrixRow In objBooleanValuesMatrix
            VerifyBooleanMatrixRow(selMatrixRow)
        Next

    End Sub

    Protected Sub VerifyBooleanMatrixRow(ByVal selMatrixRow As BooleanValuesMatrixRow)
        Dim intResponseClassesCount As Integer = CurrentIndicator.ResponseClasses.Count

        If selMatrixRow.BooleanValues.Count <> intResponseClassesCount Then
            selMatrixRow.BooleanValues.Clear()
            For j = 0 To intResponseClassesCount - 1
                selMatrixRow.BooleanValues.Add(New BooleanValue)
            Next
        End If
    End Sub

    Public Overridable Sub LoadColumns()
        Columns.Clear()

        With colStatement
            .Name = "Statement"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Statement
            .MinimumWidth = 150
            .Width = 400
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        End With
        Me.Columns.Add(colStatement)

        LoadTargetColumns()
        SetColumnHeadersText()
        Reload()
    End Sub

    Protected Sub LoadTargetColumns()
        If Columns.Count > 1 Then
            For i = 1 To Columns.Count - 1
                Columns.RemoveAt(1)
            Next
        End If

        For Each selResponseClass As ResponseClass In CurrentIndicator.ResponseClasses
            Dim colResponseClass As New DataGridViewCheckBoxColumn
            With colResponseClass
                .Name = String.Format("Class{0}", CurrentIndicator.ResponseClasses.IndexOf(selResponseClass))
                .HeaderText = vbCrLf & vbCrLf & vbCrLf
                .MinimumWidth = 100
                .FillWeight = 1
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            Me.Columns.Add(colResponseClass)
        Next
    End Sub

    Protected Sub SetColumnHeadersText()
        Dim selResponseClass As ResponseClass
        Dim intTargetIndex As Integer

        'header text of target columns
        ReDim strHeaderText(CurrentIndicator.ResponseClasses.Count - 1)
        For intTargetIndex = 0 To CurrentIndicator.ResponseClasses.Count - 1
            selResponseClass = CurrentIndicator.ResponseClasses(intTargetIndex)
            strHeaderText(intTargetIndex) = String.Format("{0}{1}{1}{2}", selResponseClass.ClassName, vbCrLf, selResponseClass.Value)
        Next
    End Sub

    Protected Sub LoadItems()
        Dim objMatrix As BooleanValuesMatrix
        Me.Grid.Clear()

        If CurrentIndicator IsNot Nothing Then
            SetBooleanValues()
            With CurrentIndicator
                If Me.TargetIndex = -1 Then
                    objMatrix = .Baseline.BooleanValuesMatrix
                Else
                    Dim selTarget As Target = .Targets(TargetIndex)
                    objMatrix = selTarget.BooleanValuesMatrix
                End If
                For intIndex = 0 To .Statements.Count - 1
                    Me.Grid.Add(New TargetScalesLikertGridRow(.Statements(intIndex), objMatrix(intIndex)))
                Next

                'empty row
                Dim NewRow As New TargetScalesLikertGridRow
                VerifyBooleanMatrixRow(NewRow.BooleanValuesMatrixRow)
                Me.Grid.Add(NewRow)
            End With
        End If
    End Sub

    Public Sub Reload()
        'remember current cell location
        objCellLocation = CurrentCellAddress
        If objCellLocation.X < 0 Then objCellLocation.X = 0
        If objCellLocation.Y < 0 Then objCellLocation.Y = 0

        '(re)load target columns and rows
        Me.SuspendLayout()

        'Rows.Clear()
        CurrentIndicator.CalculateScores()
        LoadItems()

        Me.RowCount = Me.Grid.Count + 2
        Me.RowModifyIndex = -1

        If CurrentIndicator.Registration <> Indicator.RegistrationOptions.BeneficiaryLevel Then
            For i = 0 To IndexPopulationTarget - 1
                Rows(i).Visible = True
            Next
            Rows(IndexPopulationTarget).Visible = False
            Rows(IndexPopulationScore).Visible = False
        Else
            For i = 0 To Columns.Count - 1
                Me(i, IndexPopulationScore).ReadOnly = True
            Next
        End If

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
#End Region

#Region "Cell images"
    Public Sub ResetAllImages()
        Dim selRow As TargetScalesLikertGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ResetRowImages(selRow)
        Next
    End Sub

    Protected Overridable Sub ResetRowImages(ByVal selRow As TargetScalesLikertGridRow)
        selRow.FirstLabelImageDirty = True
        selRow.FirstLabelHeight = 0
    End Sub

    Public Sub ReloadImages()
        Dim selRow As TargetScalesLikertGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ReloadImage_Statement(selRow)
        Next

        ResetRowHeights()
    End Sub

    Protected Overridable Sub ReloadImage_Statement(ByVal selRow As TargetScalesLikertGridRow)
        Dim intColumnWidth As Integer
        Dim selStatement As Statement = selRow.Statement
        Dim boolSelected As Boolean

        If selStatement Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.FirstLabelImageDirty = True Then
                intColumnWidth = colStatement.Width

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                    boolSelected = True
                Else
                    boolSelected = False
                End If

                selRow.FirstLabelImage = .RichTextWithPaddingToBitmap(intColumnWidth, selStatement.FirstLabel, boolSelected)
                selRow.FirstLabelHeight = selRow.FirstLabelImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim selGridRow As TargetScalesLikertGridRow = Me.Grid(RowIndex)

        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > CONST_MinRowHeight Then selRow.Height = intRowHeight Else selRow.Height = CONST_MinRowHeight
    End Sub

    Public Sub ResetRowHeights()
        For i = 0 To RowCount - 2
            SetRowHeight(i)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selGridRow As TargetScalesLikertGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing AndAlso selGridRow.FirstLabelHeight > intRowHeight Then intRowHeight = selGridRow.FirstLabelHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex = IndexPopulationTarget And intColumnIndex > 0 Then
            EditPopulationPercentage(intRowIndex)
        End If
    End Sub

    Protected Overridable Sub EditBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditBox.Leave
        EditBoxClose()
    End Sub

    Private Sub EditBoxClose()
        Dim strValue As String = EditBox.Text
        Dim dblValue As Double = ParseDouble(strValue)

        Select Case EditPopulationIndex
            Case -1
                CurrentIndicator.Baseline.PopulationPercentage = dblValue
            Case Else
                CurrentIndicator.PopulationTargets(EditPopulationIndex).Percentage = dblValue
        End Select

        UndoRedo.ValueChanged(dblValue)

        Me.Controls.Remove(EditBox)
        Reload()
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As Object)
        Dim intColIndex, intRowIndex As Integer

        Select Case selItem.GetType
            Case GetType(Baseline)
                intColIndex = 1
                intRowIndex = IndexPopulationTarget
            Case GetType(PopulationTarget)
                intColIndex = 1
                intRowIndex = IndexPopulationTarget
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            EditPopulationPercentage(intRowIndex)
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

        CurrentCell = Me(intColIndex, IndexPopulationTarget - 1)
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
        Dim selGridRow As TargetScalesLikertGridRow
        Dim intColIndex As Integer = 1
        Dim intRowIndex As Integer

        For intRowIndex = 0 To CurrentIndicator.Statements.Count - 1
            selGridRow = Grid(intRowIndex)

            If selGridRow.Statement Is selStatement Then
                Exit For
            End If
        Next

        If intRowIndex >= 0 And intRowIndex < IndexPopulationTarget Then
            CurrentCell = Me(intColIndex, intRowIndex)
        End If
    End Sub

    Protected Overridable Sub EditPopulationPercentage(ByVal intRowIndex As Integer)
        Dim rSpan As Rectangle
        Dim ptLocation As Point
        Dim dblValue As Double

        Select Case TargetIndex
            Case -1
                dblValue = CurrentIndicator.Baseline.PopulationPercentage
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.Baseline, "PopulationPercentage", dblValue)
                EditPopulationIndex = -1
            Case Else
                dblValue = CurrentIndicator.PopulationTargets(TargetIndex).Percentage
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.PopulationTargets(intTargetIndex), "Percentage", dblValue)
                EditPopulationIndex = TargetIndex
        End Select

        rSpan = GetCellDisplayRectangle(1, intRowIndex, False)
        If ColumnCount > 1 Then
            For i = 2 To ColumnCount - 1
                rSpan.Width += GetColumnDisplayRectangle(i, False).Width
            Next
        End If

        rSpan.Width -= 2
        ptLocation = New Point(rSpan.X + 1, rSpan.Y + 1)

        With EditBox
            .Size = New Size(rSpan.Width, rSpan.Height)
            .Location = ptLocation
            .Text = dblValue
            .SelectAll()
        End With
        Me.Controls.Add(EditBox)
        EditBox.Focus()
    End Sub

    Private Sub MoveCurrentCell()
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex

        If CurrentCell.RowIndex = IndexPopulationScore Then
            CurrentCell = Me(CurrentCell.ColumnIndex, IndexPopulationScore - 1)
            CurrentCell.Selected = True
        Else
            Dim objCurrentCell As DataGridViewCell = Me(intColumnIndex, intRowIndex)
            CurrentCell = Nothing
            CurrentCell = objCurrentCell
        End If
        Invalidate()
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

            For Each selGridRow As TargetScalesLikertGridRow In SelectedGridRows
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
            Dim selRow As TargetScalesLikertGridRow
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

        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim RowTmp As TargetScalesLikertGridRow = Nothing

        If intRowIndex >= RowCount - 2 Then
            Return
        End If
        ' Store a reference to the grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(intRowIndex), TargetScalesLikertGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case intColumnIndex
            Case 0
                e.Value = RowTmp.FirstLabel
            Case Else
                e.Value = OnCellValueNeeded_Targets(RowTmp, intColumnIndex)
        End Select
    End Sub

    Protected Function OnCellValueNeeded_Targets(ByVal RowTmp As TargetScalesLikertGridRow, ByVal intColumnIndex As Integer) As Boolean
        Dim boolValue As Boolean

        Dim intIndex As Integer = intColumnIndex - 1
        Dim selBooleanValue As BooleanValue = RowTmp.BooleanValuesMatrixRow.BooleanValues(intIndex)

        boolValue = selBooleanValue.Value

        Return boolValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As TargetScalesLikertGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim boolValue As Boolean
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex < IndexPopulationTarget Then
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As TargetScalesLikertGridRow = CType(Me.Grid(intRowIndex), TargetScalesLikertGridRow)

                Me.EditRow = New TargetScalesLikertGridRow
                With EditRow
                    .Statement = CurrentGridRow.Statement
                    .FirstLabelHeight = CurrentGridRow.FirstLabelHeight
                    .FirstLabelImage = CurrentGridRow.FirstLabelImage
                    .FirstLabelImageDirty = CurrentGridRow.FirstLabelImageDirty
                    .BooleanValuesMatrixRow = CurrentGridRow.BooleanValuesMatrixRow
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex

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
                    Boolean.TryParse(e.Value, boolValue)
            End Select

            Select Case strColName
                Case "Statement"
                    If RowTmp.Statement IsNot Nothing Then
                        UndoRedo.UndoBuffer_Initialise(RowTmp.Statement, "FirstLabel", RowTmp.FirstLabel)

                        RowTmp.FirstLabel = strCellValue
                        RowTmp.FirstLabelImageDirty = True

                        UndoRedo.TextChanged(strCellValue)
                    End If
                Case Else
                    Dim intIndex As Integer = intColumnIndex - 1
                    Dim objOldValues As BooleanValues

                    With RowTmp.BooleanValuesMatrixRow
                        Using copier As New ObjectCopy
                            objOldValues = copier.CopyCollection(.BooleanValues)
                        End Using

                        For i = 0 To CurrentIndicator.ResponseClasses.Count - 1
                            .BooleanValues(i).Value = False
                        Next
                        .BooleanValues(intIndex).Value = boolValue

                        UndoRedo.BooleanValueChecked(RowTmp.BooleanValuesMatrixRow, intIndex, objOldValues, .BooleanValues)
                    End With
            End Select
        End If
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New TargetScalesLikertGridRow
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
        Me.EditRow = New TargetScalesLikertGridRow()
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
        If Me.EditRow IsNot Nothing AndAlso e.RowIndex >= CurrentIndicator.Statements.Count AndAlso e.RowIndex < IndexPopulationTarget Then
            ' Add the new planning grid row to grid.
            Dim NewRow As BooleanValuesMatrixRow

            With CurrentIndicator
                .Statements.Add(Me.EditRow.Statement)

                UndoRedo.ItemInserted(Me.EditRow.Statement, CurrentIndicator.Statements)

                NewRow = New BooleanValuesMatrixRow
                VerifyBooleanMatrixRow(NewRow)

                .Baseline.BooleanValuesMatrix.Add(NewRow)

                For Each selTarget As Target In .Targets
                    NewRow = New BooleanValuesMatrixRow
                    VerifyBooleanMatrixRow(NewRow)

                    selTarget.BooleanValuesMatrix.Add(NewRow)
                Next

                If TargetIndex < 0 Then
                    Dim LastBaselineRow As BooleanValuesMatrixRow = .Baseline.BooleanValuesMatrix(e.RowIndex)
                    LastBaselineRow = EditRow.BooleanValuesMatrixRow
                Else
                    Dim LastTargetRow As BooleanValuesMatrixRow = .Targets(TargetIndex).BooleanValuesMatrix(e.RowIndex)
                    LastTargetRow = EditRow.BooleanValuesMatrixRow
                End If
            End With

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

        RaiseEvent ScoreUpdated()
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex = -1 And intColumnIndex >= 0 Then
            OnCellPainting_ColumnHeaders(e, intColumnIndex)
        End If

        If intRowIndex = IndexPopulationTarget Then
            If intColumnIndex >= 0 Then
                OnCellPainting_PopulationTargets(e)

                e.Handled = True
            End If
        ElseIf intRowIndex = IndexPopulationScore Then
            If intColumnIndex >= 0 Then
                OnCellPainting_PopulationScores(e)

                e.Handled = True
            End If
        Else
            OnCellPainting_NormalCells(e)
        End If
    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As TargetScalesLikertGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow IsNot Nothing Then
            If selGridRow.FirstLabelImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.FirstLabelImage.Width, selGridRow.FirstLabelImage.Height)
                CellGraphics.DrawImage(selGridRow.FirstLabelImage, rImage)
            End If
        End If
    End Sub

    Private Sub OnCellPainting_NormalCells(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If e.RowIndex >= 0 Then
            If intColumnIndex = 0 Then
                Dim selGridRow As TargetScalesLikertGridRow = Grid(intRowIndex)

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                    CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
                Else
                    CellGraphics.FillRectangle(Brushes.White, rCell)
                End If

                CellPainting_RichText(selGridRow, rCell, CellGraphics)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                e.Handled = True
            End If
        End If
    End Sub

    Private Sub OnCellPainting_PopulationTargets(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim selGridRow As TargetScalesLikertGridRow = Grid(intRowIndex)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim strValue As String = String.Empty

        Dim sfValue As New StringFormat
        sfValue.LineAlignment = StringAlignment.Center
        sfValue.Alignment = StringAlignment.Center

        If intColumnIndex = 0 Then
            sfValue.Alignment = StringAlignment.Near
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            CellGraphics.DrawString(LANG_PopulationTargetText, Me.Font, Brushes.Black, rCell, sfValue)
            e.Paint(rCell, DataGridViewPaintParts.Border)
        Else
            Dim rSpan As Rectangle = GetCellDisplayRectangle(1, intRowIndex, False)
            If ColumnCount > 1 Then
                For i = 2 To ColumnCount - 1
                    rSpan.Width += GetColumnDisplayRectangle(i, False).Width
                Next
            End If

            rSpan.Height -= 1
            CellGraphics.FillRectangle(SystemBrushes.Control, rSpan)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            If TargetIndex = -1 Then
                strValue = DisplayAsUnit(CurrentIndicator.Baseline.PopulationPercentage, CurrentIndicator.ValuesDetail.NrDecimals, "%")
                CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
            Else
                Dim selTargetPopulation As PopulationTarget = CurrentIndicator.PopulationTargets(TargetIndex)

                If selTargetPopulation IsNot Nothing Then strValue = DisplayAsUnit(selTargetPopulation.Percentage, CurrentIndicator.ValuesDetail.NrDecimals, "%")
                CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
            End If
        End If
    End Sub

    Private Sub OnCellPainting_PopulationScores(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim strValue As String = String.Empty
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim sfValue As New StringFormat
        sfValue.LineAlignment = StringAlignment.Center
        sfValue.Alignment = StringAlignment.Center

        If intColumnIndex = 0 Then
            If Me.TargetGroup IsNot Nothing Then
                Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

                strValue = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroup.Number, strTypeName)
            Else
                strValue = LANG_ScoreValueTargetGroup
            End If

            sfValue.Alignment = StringAlignment.Near

            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rCell, sfValue)
            e.Paint(rCell, DataGridViewPaintParts.Border)
        Else
            Dim rSpan As Rectangle = GetCellDisplayRectangle(1, intRowIndex, False)
            If ColumnCount > 1 Then
                For i = 2 To ColumnCount - 1
                    rSpan.Width += GetColumnDisplayRectangle(i, False).Width
                Next
            End If

            rSpan.Height -= 1
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rSpan)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            If TargetIndex = -1 Then
                strValue = CurrentIndicator.GetPopulationBaselineFormattedScore()
                CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
            Else
                strValue = CurrentIndicator.GetPopulationTargetFormattedScore(TargetIndex)
                CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
            End If
        End If
    End Sub
#End Region

#Region "Custom painting - column headers"
    Private Sub OnCellPainting_ColumnHeaders(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs, ByVal intColumnIndex As Integer)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim sfHeader As New StringFormat
        sfHeader.LineAlignment = StringAlignment.Center
        sfHeader.Alignment = StringAlignment.Center

        If intColumnIndex = 0 Then
            e.PaintBackground(rCell, False)
            CellGraphics.DrawString(LANG_Statement, Me.Font, Brushes.Black, rCell, sfHeader)
            e.Paint(rCell, DataGridViewPaintParts.Border)
        Else
            Dim intResponseClassIndex As Integer = intColumnIndex - 1

            e.PaintBackground(rCell, False)
            CellGraphics.DrawString(strHeaderText(intResponseClassIndex), Me.Font, Brushes.Black, rCell, sfHeader)
            e.Paint(rCell, DataGridViewPaintParts.Border)
        End If
        e.Handled = True
    End Sub

    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim RowGraphics As Graphics = e.Graphics

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
