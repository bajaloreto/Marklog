Public Class DataGridViewTargetsMaxDiff
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New TargetMaxDiffGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents EditBox As New NumericTextBox

    Public Event ScoreUpdated()

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlines As TargetDeadlines
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colStatement As New RichTextColumnLogframe
    Private colMostImportant, colLeastImportant As New DataGridViewCheckBoxColumn

    Private RowModifyIndex As Integer = -1
    Private EditRow As TargetMaxDiffGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private UpdateScore As Boolean

    Private intTargetIndex As Integer = -1
    Private strBaseline As String
    Private strHeaderText() As String
    Private strPopulationTarget, strPopulationScore As String
    Private EditPopulationIndex As Integer

    Protected Const CONST_MinRowHeight As Integer = 24

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

    Protected ReadOnly Property TargetGroupNumber As Integer
        Get
            If Me.TargetGroup IsNot Nothing Then
                Return Me.TargetGroup.Number
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property TargetIndex As Integer
        Get
            Return intTargetIndex
        End Get
        Set(ByVal value As Integer)
            intTargetIndex = value
        End Set
    End Property

    Protected ReadOnly Property IndexPopulationTarget As Integer
        Get
            Return CurrentIndicator.Statements.Count
        End Get
    End Property

    Protected ReadOnly Property IndexPopulationScore As Integer
        Get
            Return CurrentIndicator.Statements.Count + 1
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New()
        ChooseSettings()
    End Sub

    Public Sub New(ByVal currentindicator As Indicator, ByVal targetdeadlines As TargetDeadlines, ByVal targetgroup As TargetGroup, ByVal targetindex As Integer)
        'datagridview settings
        ChooseSettings()

        'load values
        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlines = targetdeadlines
        Me.TargetGroup = targetgroup
        Me.TargetIndex = targetindex

        SetBooleanValues()
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
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

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
            Dim intResponseClassesCount As Integer = 2
            Dim intResponsesCount As Integer = .Statements.Count

            VerifyBooleanMatrix(.Baseline.BooleanValuesMatrix, intResponsesCount, intResponseClassesCount)
            For Each selTarget As Target In .Targets
                VerifyBooleanMatrix(selTarget.BooleanValuesMatrix, intResponsesCount, intResponseClassesCount)
            Next
        End With
    End Sub

    Protected Sub VerifyBooleanMatrix(ByVal objBooleanValuesMatrix As BooleanValuesMatrix, ByVal intResponsesCount As Integer, ByVal intResponseClassesCount As Integer)
        If objBooleanValuesMatrix.Count <> intResponsesCount Then
            objBooleanValuesMatrix.Clear()

            For i = 0 To intResponsesCount - 1
                Dim NewMatrixRow As New BooleanValuesMatrixRow

                objBooleanValuesMatrix.Add(NewMatrixRow)
            Next
        End If

        For Each selMatrixRow As BooleanValuesMatrixRow In objBooleanValuesMatrix
            If selMatrixRow.BooleanValues.Count <> intResponseClassesCount Then
                selMatrixRow.BooleanValues.Clear()
                For j = 0 To intResponseClassesCount - 1
                    selMatrixRow.BooleanValues.Add(New BooleanValue)
                Next
            End If
        Next

    End Sub

    Public Overridable Sub LoadColumns()
        Columns.Clear()

        Dim strLeastImportant As String = String.Empty
        Dim strMostImportant As String = String.Empty

        If CurrentIndicator IsNot Nothing Then
            strLeastImportant = CurrentIndicator.ScalesDetail.AgreeText
            strMostImportant = CurrentIndicator.ScalesDetail.DisagreeText
        End If

        With colMostImportant
            .Name = "Least important"
            .HeaderText = strLeastImportant
            .MinimumWidth = 100
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colMostImportant)

        With colStatement
            .Name = "Statement"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Statement
            .MinimumWidth = 150
            .Width = 400
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colStatement)

        With colLeastImportant
            .Name = "Most important"
            .HeaderText = strMostImportant
            .MinimumWidth = 100
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colLeastImportant)

        Reload()
    End Sub

    Protected Overridable Sub SetRowHeadersText()

        strPopulationTarget = LANG_PopulationTargetText

        If Me.TargetGroup IsNot Nothing Then
            Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

            strPopulationScore = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
        Else
            strPopulationScore = LANG_ScoreValueTargetGroup
        End If

        If CurrentIndicator.Registration <> Indicator.RegistrationOptions.BeneficiaryLevel Then
            Rows(IndexPopulationTarget).Visible = False
            Rows(IndexPopulationScore).Visible = False
        Else
            For i = 0 To Columns.Count - 1
                Me(i, IndexPopulationScore).ReadOnly = True
            Next
        End If
    End Sub

    Protected Sub LoadItems()
        Dim objMatrix As BooleanValuesMatrix
        Me.Grid.Clear()

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                If Me.TargetIndex = -1 Then
                    objMatrix = .Baseline.BooleanValuesMatrix
                Else
                    Dim selTarget As Target = .Targets(TargetIndex)
                    objMatrix = selTarget.BooleanValuesMatrix
                End If
                For intIndex = 0 To .Statements.Count - 1
                    Me.Grid.Add(New TargetMaxDiffGridRow(.Statements(intIndex), objMatrix(intIndex).BooleanValues(0), objMatrix(intIndex).BooleanValues(1)))
                Next
            End With
        End If
    End Sub

    Public Overridable Sub Reload()
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

        '(re)load target columns and rows
        Me.SuspendLayout()

        Rows.Clear()
        CurrentIndicator.CalculateScores()
        LoadItems()

        Me.RowCount = Me.Grid.Count + 2
        Me.RowModifyIndex = -1

        SetRowHeadersText()

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
    Protected Overridable Sub ResetAllImages()
        Dim selRow As TargetMaxDiffGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As TargetMaxDiffGridRow)
        selRow.ResponseImageDirty = True
        selRow.ResponseHeight = 0
    End Sub

    Protected Overridable Sub ReloadImages()
        Dim selRow As TargetMaxDiffGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ReloadImage_Response(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImage_Response(ByVal selRow As TargetMaxDiffGridRow)
        Dim intColumnWidth As Integer
        Dim selStatement As Statement = selRow.Response
        Dim boolSelected As Boolean

        If selStatement Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.ResponseImageDirty = True Then
                intColumnWidth = colStatement.Width

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                    boolSelected = True
                Else
                    boolSelected = False
                End If

                selRow.ResponseImage = .RichTextWithPaddingToBitmap(intColumnWidth, selStatement.FirstLabel, boolSelected)
                selRow.ResponseHeight = selRow.ResponseImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim selGridRow As TargetMaxDiffGridRow = Me.Grid(RowIndex)

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
        Dim selGridRow As TargetMaxDiffGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing AndAlso selGridRow.ResponseHeight > intRowHeight Then intRowHeight = selGridRow.ResponseHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)

        MoveCurrentCell()
    End Sub

    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex = IndexPopulationTarget And intColumnIndex > 0 Then
            EditPopulationPercentage(intColumnIndex, intRowIndex)
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

            CurrentCell.Selected = True

        End If
        MyBase.OnKeyUp(e)
    End Sub

    Protected Sub EditBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles EditBox.KeyUp
        If e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Then
            Me.Select()
        End If
    End Sub

    Protected Overridable Sub EditBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditBox.Leave
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
        'Invalidate()
    End Sub
#End Region

#Region "General methods"
    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex
        Me.ReloadImages()
    End Sub

    Protected Overridable Sub EditPopulationPercentage(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Dim rFirstCell, rSecondCell As Rectangle
        Dim rSpan As Rectangle
        Dim ptLocation As Point
        Dim dblValue As Double

        Select Case intColumnIndex
            Case 1, 2
                rFirstCell = Me.GetCellDisplayRectangle(1, intRowIndex, False)
                rSecondCell = Me.GetCellDisplayRectangle(2, intRowIndex, False)

                dblValue = CurrentIndicator.Baseline.PopulationPercentage
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.Baseline, "PopulationPercentage", dblValue)
                EditPopulationIndex = -1
            Case Else
                Dim intTargetIndex As Integer = Math.Floor((intColumnIndex - 3) / 2)
                Dim intIndexFirstCell As Integer = (intTargetIndex * 2) + 3

                rFirstCell = Me.GetCellDisplayRectangle(intIndexFirstCell, intRowIndex, False)
                rSecondCell = Me.GetCellDisplayRectangle(intIndexFirstCell + 1, intRowIndex, False)

                dblValue = CurrentIndicator.PopulationTargets(intTargetIndex).Percentage
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.PopulationTargets(intTargetIndex), "Percentage", dblValue)
                EditPopulationIndex = intTargetIndex
        End Select

        rSpan = rFirstCell
        rSpan.Width += rSecondCell.Width
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
        If CurrentCell.RowIndex = IndexPopulationScore Then
            CurrentCell = Me(CurrentCell.ColumnIndex, IndexPopulationScore - 1)
            CurrentCell.Selected = True
        Else
            Dim objCurrentCell As DataGridViewCell = CurrentCell
            CurrentCell = Nothing
            CurrentCell = objCurrentCell
        End If
        Invalidate()
    End Sub

    Public Sub Edit()
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intTargetIndex As Integer = intColumnIndex - 1

        If intRowIndex < IndexPopulationTarget Then
            If CurrentCell.ReadOnly = False Then Me.BeginEdit(True)
        ElseIf intRowIndex = IndexPopulationTarget Then
            EditPopulationPercentage(intRowIndex, intColumnIndex)
        End If
    End Sub

    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)

    End Sub
#End Region

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)
        CopyItems()

        RemoveItems(ShowWarning)
    End Sub

    Public Overrides Sub CopyItems()
        'With SelectionRectangle
        '    Dim selRow As PlanningGridRow
        '    Dim strSort As String
        '    Dim CopyGroup As Date = Now()

        '    With SelectionRectangle
        '        For i = .FirstRowIndex To .LastRowIndex
        '            selRow = Me.Grid(i)
        '            If selRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
        '                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.KeyMoment Then
        '                    strSort = KeyMoment.ItemName & " " & selRow.SortNumber
        '                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.KeyMoment, strSort)
        '                    ItemClipboard.Insert(0, NewItem)
        '                End If
        '            ElseIf selRow.RowType = PlanningGridRow.RowTypes.Activity Then
        '                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Struct Then
        '                    strSort = Activity.ItemName & " " & selRow.SortNumber
        '                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.Struct, strSort)
        '                    ItemClipboard.Insert(0, NewItem)
        '                End If
        '            End If
        '        Next
        '    End With
        'End With

    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
        'If PasteCell Is Nothing Then PasteCell = CurrentCell
        'If PasteCell Is Nothing Then Exit Sub

        'Dim intColumnIndex As Integer = PasteCell.ColumnIndex
        'Dim intRowIndex As Integer = PasteCell.RowIndex
        'Dim selItem As ClipboardItem

        'PasteRow = Grid(intRowIndex)

        'For i = 0 To PasteItems.Count - 1 'To 0 Step -1
        '    selItem = PasteItems(i)
        '    Select Case selItem.Item.GetType
        '        Case GetType(KeyMoment)
        '            PasteItems_KeyMoment(selItem, intPasteOption)
        '        Case GetType(Activity)
        '            PasteItems_Activity(selItem, intPasteOption)
        '        Case Else
        '            PasteItems_Other(selItem, intPasteOption)
        '    End Select
        'Next

        'Me.Reload()
        'Me.CurrentCell = Me(intColumnIndex, intRowIndex)
    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)

        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim RowTmp As TargetMaxDiffGridRow = Nothing

        If intRowIndex >= RowCount - 2 Then
            Return
        End If
        ' Store a reference to the grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(intRowIndex), TargetMaxDiffGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case intColumnIndex
            Case 0
                e.Value = RowTmp.LeastImportantValue.Value
            Case 1
                e.Value = RowTmp.FirstLabel
            Case 2
                e.Value = RowTmp.MostImportantValue.Value
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As TargetMaxDiffGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim boolValue As Boolean
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex < CurrentIndicator.Statements.Count Then
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As TargetMaxDiffGridRow = CType(Me.Grid(intRowIndex), TargetMaxDiffGridRow)

                Me.EditRow = New TargetMaxDiffGridRow
                With EditRow
                    If CurrentGridRow.Response IsNot Nothing Then .Response = CurrentGridRow.Response Else .Response = New Statement
                    .LeastImportantValue = CurrentGridRow.LeastImportantValue
                    .MostImportantValue = CurrentGridRow.MostImportantValue
                    .ResponseHeight = CurrentGridRow.ResponseHeight
                    .ResponseImage = CurrentGridRow.ResponseImage
                    .ResponseImageDirty = CurrentGridRow.ResponseImageDirty
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

            Select Case e.ColumnIndex
                Case 0
                    For i = 0 To CurrentIndicator.Statements.Count - 1
                        Grid(i).LeastImportantValue.Value = False
                    Next
                    RowTmp.LeastImportantValue.Value = boolValue
                Case 1
                    If RowTmp.Response IsNot Nothing Then
                        RowTmp.FirstLabel = strCellValue
                        RowTmp.ResponseImageDirty = True
                    End If
                Case 2
                    For i = 0 To CurrentIndicator.Statements.Count - 1
                        Grid(i).MostImportantValue.Value = False
                    Next
                    RowTmp.MostImportantValue.Value = boolValue
            End Select

            'CurrentUndoList.AddToUndoList()
        End If
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New TargetMaxDiffGridRow
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
        Me.EditRow = New TargetMaxDiffGridRow()
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
            'Me.Grid.Add(Me.EditRow)
            'Me.EditRow = Nothing
            'Me.EditRowFlag = -1
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then
            ' Save the modified planning grid row in grid.
            Me.Grid(e.RowIndex) = Me.EditRow
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        MyBase.OnRowValidated(e)
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex = -1 Then
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

    Private Sub CellPainting_RichText(ByVal selGridRow As TargetMaxDiffGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow IsNot Nothing Then
            If selGridRow.ResponseImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.ResponseImage.Width, selGridRow.ResponseImage.Height)
                CellGraphics.DrawImage(selGridRow.ResponseImage, rImage)
            End If
        End If
    End Sub

    Private Sub OnCellPainting_NormalCells(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If e.RowIndex >= 0 Then
            If intColumnIndex = 1 Then
                Dim selGridRow As TargetMaxDiffGridRow = Grid(intRowIndex)

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
        Dim selGridRow As TargetMaxDiffGridRow = Grid(intRowIndex)
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

                strValue = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
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
        'Dim CellGraphics As Graphics = e.Graphics
        'Dim rCell As Rectangle = e.CellBounds
        'Dim sfHeader As New StringFormat
        'sfHeader.LineAlignment = StringAlignment.Center
        'sfHeader.Alignment = StringAlignment.Center

        'If intColumnIndex = 0 Then
        '    e.PaintBackground(rCell, False)
        '    CellGraphics.DrawString(LANG_Statement, Me.Font, Brushes.Black, rCell, sfHeader)
        '    e.Paint(rCell, DataGridViewPaintParts.Border)
        'Else
        '    Dim intResponseClassIndex As Integer = intColumnIndex - 1

        '    e.PaintBackground(rCell, False)
        '    CellGraphics.DrawString(strHeaderText(intResponseClassIndex), Me.Font, Brushes.Black, rCell, sfHeader)
        '    e.Paint(rCell, DataGridViewPaintParts.Border)
        'End If
        'e.Handled = True
    End Sub
#End Region
End Class
