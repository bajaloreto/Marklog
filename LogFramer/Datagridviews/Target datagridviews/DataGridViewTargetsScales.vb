Public Class DataGridViewTargetsScales
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New TargetScalesGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents EditBox As New NumericTextBox

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colStatement As New RichTextColumnLogframe
    Private colAgree, colDisagree As New DataGridViewCheckBoxColumn

    Private RowModifyIndex As Integer = -1
    Private EditRow As TargetScalesGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private UpdateScore As Boolean

    Private strBaseline As String
    Private strHeaderText() As String
    Private strPopulationTarget, strPopulationScore As String
    Private EditPopulationIndex As Integer

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

    Public Property TargetDeadlinesSection As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesSection
        End Get
        Set(ByVal value As TargetDeadlinesSection)
            objTargetDeadlinesSection = value
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

    Public Sub New(ByVal currentindicator As Indicator, ByVal targetdeadlinessection As TargetDeadlinesSection, ByVal targetgroup As TargetGroup)
        'datagridview settings
        ChooseSettings()

        'edit box settings
        EditBox.ValueType = NumericTextBox.ValueTypes.DoubleValue

        'load values
        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlinesSection = targetdeadlinessection
        Me.TargetGroup = targetgroup

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

    Public Overridable Sub SetBooleanValues()
        With Me.CurrentIndicator
            Dim intCount As Integer
            If Me.CurrentIndicator.QuestionType = Indicator.QuestionTypes.CumulativeScale Then
                intCount = .Statements.Count
            Else
                intCount = .ResponseClasses.Count
            End If

            VerifyNumberOfBooleanValues(.Baseline.BooleanValues, intCount)
            For Each selTarget As Target In .Targets
                VerifyNumberOfBooleanValues(selTarget.BooleanValues, intCount)
            Next
        End With
    End Sub

    Protected Overridable Sub VerifyNumberOfBooleanValues(ByVal objBooleanValues As BooleanValues, ByVal intCount As Integer)
        If objBooleanValues.Count <> intCount Then
            objBooleanValues.Clear()
            For i = 0 To intCount - 1
                objBooleanValues.Add(New BooleanValue)
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
            .Width = 200
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        End With
        Me.Columns.Add(colStatement)

        With colAgree
            .Name = "Baseline Agree"
            .HeaderText = vbCrLf & vbCrLf & vbCrLf
            .MinimumWidth = 100
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colAgree)

        With colDisagree
            .Name = "Baseline Disagree"
            .MinimumWidth = 100
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colDisagree)

        Reload()
    End Sub

    Protected Overridable Sub LoadTargetColumns()
        If Columns.Count > 3 Then
            For i = 3 To Columns.Count - 1
                Columns.RemoveAt(3)
            Next
        End If

        For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
            Dim colAgree As New DataGridViewCheckBoxColumn
            With colAgree
                .Name = selTargetDeadline.Deadline
                .MinimumWidth = 100
                .FillWeight = 1
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            Columns.Add(colAgree)

            Dim colDisAgree As New DataGridViewCheckBoxColumn
            With colDisAgree
                .Name = selTargetDeadline.Deadline
                .MinimumWidth = 100
                .FillWeight = 20
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            Columns.Add(colDisAgree)
        Next
    End Sub

    Protected Overridable Sub SetColumnHeadersText()
        Dim strDate As String
        Dim strScore As String = String.Empty
        Dim intTargetIndex As Integer
        Dim strBaselineScore As String = 0.ToString

        CurrentIndicator.CalculateScores()

        'header text of baseline column
        strBaselineScore = CurrentIndicator.GetBaselineFormattedScore

        'set header text
        strBaseline = String.Format("{0}{1}{2}: {3}", LANG_Baseline, vbCrLf, LANG_ScoringValue, strBaselineScore)

        'header text of target columns
        ReDim Preserve strHeaderText(TargetDeadlinesSection.TargetDeadlines.Count - 1)
        For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
            'deadline
            intTargetIndex = TargetDeadlinesSection.TargetDeadlines.IndexOf(selTargetDeadline)

            Select Case TargetDeadlinesSection.Repetition
                Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                    strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
                Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                    strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
                Case Else
                    strDate = selTargetDeadline.ExactDeadline.ToShortDateString
            End Select

            'score
            strScore = CurrentIndicator.GetTargetFormattedScore(intTargetIndex)
            strScore = String.Format("{0}: {1}", LANG_ScoringValue, strScore)


            'set header text
            strHeaderText(intTargetIndex) = String.Format("{0} {1}{2}{3}", LANG_Target, strDate, vbCrLf, strScore)
        Next
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

    Protected Overridable Sub LoadItems()
        Me.Grid.Clear()

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                For intIndex = 0 To .Statements.Count - 1
                    Dim BaselineBooleanValue As BooleanValue = .Baseline.BooleanValues(intIndex)
                    Dim TargetBooleanValues As New BooleanValues

                    For Each selTarget As Target In CurrentIndicator.Targets
                        TargetBooleanValues.Add(selTarget.BooleanValues(intIndex))
                    Next

                    Me.Grid.Add(New TargetScalesGridRow(.Statements(intIndex), BaselineBooleanValue, TargetBooleanValues))
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

        'Rows.Clear()
        CurrentIndicator.CalculateScores()
        LoadItems()

        Me.RowCount = Me.Grid.Count + 2
        Me.RowModifyIndex = -1

        LoadTargetColumns()
        SetColumnHeadersText()
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
        Dim selRow As TargetScalesGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As TargetScalesGridRow)
        selRow.StatementImageDirty = True
        selRow.StatementHeight = 0
    End Sub

    Protected Overridable Sub ReloadImages()
        Dim selRow As TargetScalesGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ReloadImage_Response(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImage_Response(ByVal selRow As TargetScalesGridRow)
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
        Dim selGridRow As TargetScalesGridRow = Me.Grid(RowIndex)

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
        Dim selGridRow As TargetScalesGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing AndAlso selGridRow.StatementHeight > intRowHeight Then intRowHeight = selGridRow.StatementHeight

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

            CurrentCell = Nothing
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
                Dim selPopulationTarget As PopulationTarget = TryCast(selItem, PopulationTarget)

                If selPopulationTarget IsNot Nothing Then
                    intColIndex = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget) + 1
                    intColIndex *= 2
                    intColIndex += 1
                    intRowIndex = IndexPopulationTarget
                End If
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            EditPopulationPercentage(intColIndex, intRowIndex)
        End If
    End Sub

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
        Dim RowTmp As TargetScalesGridRow = Nothing

        If intRowIndex >= RowCount - 2 Then
            Return
        End If
        ' Store a reference to the grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(intRowIndex), TargetScalesGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case intColumnIndex
            Case 0
                e.Value = RowTmp.FirstLabel
            Case 1, 2
                e.Value = OnCellValueNeeded_Baseline(RowTmp, intColumnIndex)
            Case Else
                e.Value = OnCellValueNeeded_Targets(RowTmp, intColumnIndex)
        End Select
    End Sub

    Private Function OnCellValueNeeded_Baseline(ByVal RowTmp As TargetScalesGridRow, ByVal intColumnIndex As Integer) As Boolean
        Dim boolValue As Boolean

        If RowTmp.BaselineBooleanValue IsNot Nothing Then
            boolValue = RowTmp.BaselineBooleanValue.Value

            If intColumnIndex = 2 Then
                If boolValue = False Then boolValue = True Else boolValue = False
            End If
        End If

        Return boolValue
    End Function

    Private Function OnCellValueNeeded_Targets(ByVal RowTmp As TargetScalesGridRow, ByVal intColumnIndex As Integer) As Boolean
        Dim boolValue As Boolean

        intColumnIndex -= 3
        Dim intTargetIndex As Integer = Math.Floor(intColumnIndex / 2)
        Dim selTargetBooleanValue As BooleanValue = RowTmp.TargetBooleanValues(intTargetIndex)

        If selTargetBooleanValue IsNot Nothing Then
            boolValue = selTargetBooleanValue.Value
            If Decimal.Remainder(intColumnIndex, 2) <> 0 Then
                If boolValue = False Then boolValue = True Else boolValue = False
            End If
        End If

        Return boolValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As TargetScalesGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim boolValue As Boolean
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex < CurrentIndicator.Statements.Count Then
            'If the user is editing a new row, create a new grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As TargetScalesGridRow = CType(Me.Grid(intRowIndex), TargetScalesGridRow)

                Me.EditRow = New TargetScalesGridRow
                With EditRow
                    If CurrentGridRow.Statement IsNot Nothing Then .Statement = CurrentGridRow.Statement
                    .StatementHeight = CurrentGridRow.StatementHeight
                    .StatementImage = CurrentGridRow.StatementImage
                    .StatementImageDirty = CurrentGridRow.StatementImageDirty
                    .BaselineBooleanValue = CurrentGridRow.BaselineBooleanValue
                    .TargetBooleanValues = CurrentGridRow.TargetBooleanValues
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex

            ' Set the appropriate objRowEdit property to the cell value entered.
            strColName = Me.Columns(e.ColumnIndex).Name
            Select Case intColumnIndex
                Case 0 'statement
                    Dim ctlRTF As RichTextEditingControlLogframe = CType(Me.EditingControl, RichTextEditingControlLogframe)

                    'necessary for pasting into non-active cell
                    If ctlRTF Is Nothing Then
                        BeginEdit(False)
                        ctlRTF = CType(Me.EditingControl, RichTextEditingControlLogframe)
                    End If

                    strCellValue = TryCast(ctlRTF.Rtf, String)
                Case Else
                    If Boolean.TryParse(e.Value, boolValue) Then
                        If Decimal.Remainder(intColumnIndex, 2) = 0 Then
                            If boolValue = False Then boolValue = True Else boolValue = False
                        End If
                    End If
            End Select

            Select Case intColumnIndex
                Case 0
                    If RowTmp.Statement Is Nothing Then
                        Dim NewStatement As New Statement
                        CurrentIndicator.Statements.Add(NewStatement)
                        UndoRedo.ItemInserted(NewStatement, CurrentIndicator.Statements)

                        RowTmp.Statement = NewStatement
                    End If
                    UndoRedo.UndoBuffer_Initialise(RowTmp.Statement, "FirstLabel", RowTmp.FirstLabel)

                    RowTmp.FirstLabel = strCellValue
                    RowTmp.StatementImageDirty = True

                    UndoRedo.TextChanged(strCellValue)

                    Me.NotifyCurrentCellDirty(True)
                    MoveCurrentCell()
                Case 1, 2
                    UndoRedo.UndoBuffer_Initialise(RowTmp.BaselineBooleanValue, "Value", RowTmp.BaselineBooleanValue.Value)
                    RowTmp.BaselineBooleanValue.Value = boolValue
                    UndoRedo.OptionChecked(boolValue)
                    UpdateScore = True
                Case Else
                    intColumnIndex -= 3
                    Dim intTargetIndex As Integer = Math.Floor(intColumnIndex / 2)

                    UndoRedo.UndoBuffer_Initialise(RowTmp.TargetBooleanValues(intTargetIndex), "Value", RowTmp.TargetBooleanValues(intTargetIndex).Value)
                    RowTmp.TargetBooleanValues(intTargetIndex).Value = boolValue
                    UndoRedo.OptionChecked(boolValue)

                    UpdateScore = True
            End Select
        End If
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New TargetScalesGridRow
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
        Me.EditRow = New TargetScalesGridRow()
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

    Private Sub CellPainting_RichText(ByVal selGridRow As TargetScalesGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow IsNot Nothing Then
            If selGridRow.StatementImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StatementImage.Width, selGridRow.StatementImage.Height)
                CellGraphics.DrawImage(selGridRow.StatementImage, rImage)
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
                Dim selGridRow As TargetScalesGridRow = Grid(intRowIndex)

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                    CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
                Else
                    CellGraphics.FillRectangle(Brushes.White, rCell)
                End If

                CellPainting_RichText(selGridRow, rCell, CellGraphics)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                e.Handled = True
            Else
                e.PaintBackground(rCell, False)
                If ColorBackGround(e.ColumnIndex) = True Then
                    CellGraphics.FillRectangle(brBlue, rCell)
                End If
                e.PaintContent(rCell)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub OnCellPainting_PopulationTargets(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim selGridRow As TargetScalesGridRow = Grid(intRowIndex)
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
            If Decimal.Remainder(intColumnIndex, 2) <> 0 Then
                Dim rCellRight As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex + 1, -1, False)

                rCell.Width += rCellRight.Width
                CellGraphics.FillRectangle(SystemBrushes.Control, rCell)
                e.Paint(rCell, DataGridViewPaintParts.Border)
            Else
                intColumnIndex -= 1

                rCell = Me.GetCellDisplayRectangle(intColumnIndex, IndexPopulationTarget, False)
                Dim rCellRight As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex + 1, IndexPopulationTarget, False)
                Dim rSpan As Rectangle = rCell

                rSpan.Width += rCellRight.Width
                rSpan.Height -= 1
                CellGraphics.FillRectangle(SystemBrushes.Control, rSpan)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                If intColumnIndex = 1 Then
                    strValue = DisplayAsUnit(CurrentIndicator.Baseline.PopulationPercentage, CurrentIndicator.ValuesDetail.NrDecimals, "%")
                    CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
                Else
                    Dim intTargetPopulationIndex As Integer = (intColumnIndex - 3) / 2
                    Dim selTargetPopulation As PopulationTarget = CurrentIndicator.PopulationTargets(intTargetPopulationIndex)

                    If selTargetPopulation IsNot Nothing Then strValue = DisplayAsUnit(selTargetPopulation.Percentage, CurrentIndicator.ValuesDetail.NrDecimals, "%")
                    CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
                End If

                CellGraphics.DrawLine(New Pen(Me.GridColor, 1), New Point(rCellRight.Right - 1, rCellRight.Top), New Point(rCellRight.Right - 1, rCellRight.Bottom))
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
            If Decimal.Remainder(intColumnIndex, 2) <> 0 Then
                Dim rCellRight As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex + 1, -1, False)

                rCell.Width += rCellRight.Width
                CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
                e.Paint(rCell, DataGridViewPaintParts.Border)
            Else
                intColumnIndex -= 1

                rCell = Me.GetCellDisplayRectangle(intColumnIndex, IndexPopulationScore, False)
                Dim rCellRight As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex + 1, IndexPopulationScore, False)
                Dim rSpan As Rectangle = rCell

                rSpan.Width += rCellRight.Width
                rSpan.Height -= 1
                CellGraphics.FillRectangle(SystemBrushes.ControlDark, rSpan)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                If intColumnIndex = 1 Then
                    strValue = CurrentIndicator.GetPopulationBaselineFormattedScore()
                    CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
                Else
                    Dim intTargetIndex As Integer = (intColumnIndex - 3) / 2

                    strValue = CurrentIndicator.GetPopulationTargetFormattedScore(intTargetIndex)
                    CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rSpan, sfValue)
                End If

                CellGraphics.DrawLine(New Pen(Me.GridColor, 1), New Point(rCellRight.Right - 1, rCellRight.Top), New Point(rCellRight.Right - 1, rCellRight.Bottom))
            End If
        End If
    End Sub

    Private Function ColorBackGround(ByVal intColumnIndex As Integer) As Boolean
        intColumnIndex -= 1
        Dim intDivider As Integer = 2
        Dim intNormalised As Integer

        intNormalised = Math.Floor(intColumnIndex / intDivider)

        If Decimal.Remainder(intNormalised, 2) = 0 Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

#Region "Custom painting - column headers"
    Private Sub OnCellPainting_ColumnHeaders(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs, ByVal intColumnIndex As Integer)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim strAgree As String = CurrentIndicator.ScalesDetail.AgreeText
        Dim strDisagree As String = CurrentIndicator.ScalesDetail.DisagreeText
        Dim sfHeader As New StringFormat
        sfHeader.LineAlignment = StringAlignment.Center
        sfHeader.Alignment = StringAlignment.Center

        If UpdateScore = True Then
            SetColumnHeadersText()
            UpdateScore = False
        End If

        If intColumnIndex = 0 Then
            CellGraphics.FillRectangle(SystemBrushes.Control, rCell)
            CellGraphics.DrawString(LANG_Statement, Me.Font, Brushes.Black, rCell, sfHeader)

            e.Paint(rCell, DataGridViewPaintParts.Border)
            CellGraphics.DrawLine(New Pen(Me.GridColor, 1), New Point(rCell.Right - 1, rCell.Top), New Point(rCell.Right - 1, rCell.Bottom))

            e.Handled = True
        ElseIf intColumnIndex >= 1 Then
            If Decimal.Remainder(intColumnIndex, 2) <> 0 Then
                Dim rCellRight As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex + 1, -1, False)
                Dim rAgree As Rectangle = rCell

                rCell.Width += rCellRight.Width
                CellGraphics.FillRectangle(SystemBrushes.Control, rCell)

                rAgree.Height = CellGraphics.MeasureString(strAgree, Me.Font, rAgree.Width).Height
                rAgree.Height += e.CellStyle.Padding.Vertical
                rAgree.Y += (rCell.Height - rAgree.Height - 2)
                CellGraphics.DrawString(strAgree, Me.Font, Brushes.Black, rAgree, sfHeader)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                e.Handled = True
            Else
                intColumnIndex -= 1
                rCell = Me.GetCellDisplayRectangle(intColumnIndex, -1, False)
                Dim rCellRight As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex + 1, -1, False)
                Dim rHeader As Rectangle = rCell
                Dim rDisagree As Rectangle = rCellRight

                rDisagree.Height = CellGraphics.MeasureString(strDisagree, Me.Font, rDisagree.Width).Height
                rDisagree.Height += e.CellStyle.Padding.Vertical
                rDisagree.Y += (rCell.Height - rDisagree.Height - 2)
                CellGraphics.DrawString(strDisagree, Me.Font, Brushes.Black, rDisagree, sfHeader)

                rHeader.Y += 2
                rHeader.Height -= (rDisagree.Height + 2)
                rHeader.Width += rCellRight.Width
                CellGraphics.FillRectangle(SystemBrushes.Control, rHeader)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                If intColumnIndex = 1 Then
                    CellGraphics.DrawString(strBaseline, Me.Font, Brushes.Black, rHeader, sfHeader)
                Else
                    Dim intTargetIndex As Integer = (intColumnIndex - 3) / 2
                    CellGraphics.DrawString(strHeaderText(intTargetIndex), Me.Font, Brushes.Black, rHeader, sfHeader)
                End If

                CellGraphics.DrawLine(New Pen(Me.GridColor, 1), New Point(rCellRight.Right - 1, rCellRight.Top), New Point(rCellRight.Right - 1, rCellRight.Bottom))
                e.Handled = True
            End If
        End If
    End Sub
#End Region
End Class
