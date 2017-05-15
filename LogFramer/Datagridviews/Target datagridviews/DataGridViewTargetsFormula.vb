Public Class DataGridViewTargetsFormula
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New TargetFormulaGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colStatement As New RichTextColumnLogframe
    Private colBaseline As New DataGridViewTextBoxColumn

    Private RowModifyIndex As Integer = -1
    Private EditRow As TargetFormulaGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private UpdateScore As Boolean

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

    Protected ReadOnly Property IndexTotal As Integer
        Get
            Return CurrentIndicator.Statements.Count
        End Get
    End Property

    Protected ReadOnly Property IndexScore As Integer
        Get
            Return CurrentIndicator.Statements.Count + 1
        End Get
    End Property

    Protected ReadOnly Property IndexPopulationTarget As Integer
        Get
            Return CurrentIndicator.Statements.Count + 2
        End Get
    End Property

    Protected ReadOnly Property IndexPopulationScore As Integer
        Get
            Return CurrentIndicator.Statements.Count + 3
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

        'load values
        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlinesSection = targetdeadlinessection
        Me.TargetGroup = targetgroup

        SetDoubleValues()
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
        'AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

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
    End Sub

    Public Overridable Sub SetDoubleValues()
        With Me.CurrentIndicator
            Dim intCount As Integer = .Statements.Count

            VerifyNumberOfDoubleValues(.Baseline.DoubleValues, intCount)
            For Each selTarget As Target In .Targets
                VerifyNumberOfDoubleValues(selTarget.DoubleValues, intCount)
            Next
        End With
    End Sub

    Protected Overridable Sub VerifyNumberOfDoubleValues(ByVal objDoubleValues As DoubleValues, ByVal intCount As Integer)
        If objDoubleValues.Count <> intCount Then
            objDoubleValues.Clear()
            For i = 0 To intCount - 1
                objDoubleValues.Add(New DoubleValue)
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

        With colBaseline
            .Name = "Baseline"
            .HeaderText = vbCrLf & vbCrLf & vbCrLf
            .MinimumWidth = 100
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
        Me.Columns.Add(colBaseline)

        Reload()
    End Sub

    Protected Overridable Sub LoadTargetColumns()
        If Columns.Count > 2 Then
            For i = 2 To Columns.Count - 1
                Columns.RemoveAt(2)
            Next
        End If

        For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
            Dim colTarget As New DataGridViewTextBoxColumn
            With colTarget
                .Name = selTargetDeadline.Deadline
                .MinimumWidth = 100
                .FillWeight = 1
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End With
            Columns.Add(colTarget)
        Next
    End Sub

    Private Sub SetColumnHeadersText()
        Dim strHeaderText, strDate As String
        Dim strFormula As String = String.Empty
        Dim strScore As String = String.Empty
        Dim intIndex, intColumnIndex As Integer

        'header text of baseline column
        Dim strBaseline As String = LANG_Baseline
        If CurrentIndicator.TargetSystem = Indicator.TargetSystems.Formula Then
            strBaseline &= " (BL)"
        End If

        If Me.Columns.Count > 1 Then Columns(1).HeaderText = strBaseline

        'header text of target columns
        For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
            intIndex = TargetDeadlinesSection.TargetDeadlines.IndexOf(selTargetDeadline)

            Select Case TargetDeadlinesSection.Repetition
                Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                    strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
                Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                    strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
                Case Else
                    strDate = selTargetDeadline.ExactDeadline.ToShortDateString
            End Select

            If CurrentIndicator.TargetSystem = Indicator.TargetSystems.Formula Then

                Dim selTarget As Target = CurrentIndicator.Targets(intIndex)
                If String.IsNullOrEmpty(selTarget.Formula) = False Then
                    strFormula = String.Format("{0} {1}", selTarget.OpMin, selTarget.Formula)
                Else
                    strFormula = String.Empty
                End If
            End If

            Select Case CurrentIndicator.ScoringSystem
                Case Indicator.ScoringSystems.Percentage
                    strScore = String.Format("{0}: 100 %", LANG_ScoringValue)
            End Select

            If CurrentIndicator.TargetSystem = Indicator.TargetSystems.Formula Then
                strHeaderText = String.Format("{0} {1} (T{2})", LANG_Target, strDate, intIndex + 1)
            Else
                strHeaderText = String.Format("{0} {1}", LANG_Target, strDate)
            End If

            If String.IsNullOrEmpty(strFormula) = False Or String.IsNullOrEmpty(strScore) = False Then
                strHeaderText &= vbCrLf
                If String.IsNullOrEmpty(strFormula) = False Then
                    strHeaderText = String.Format("{0}{1}{2}", strHeaderText, vbCrLf, strFormula)
                End If
                If String.IsNullOrEmpty(strScore) = False Then
                    strHeaderText = String.Format("{0}{1}{2}", strHeaderText, vbCrLf, strScore)
                End If
            End If

            intColumnIndex = intIndex + 2
            If intColumnIndex <= Me.Columns.Count - 1 Then
                Columns(intColumnIndex).HeaderText = strHeaderText
            End If
        Next
    End Sub

    Private Sub SetRowHeadersText()
        Dim strFirst As String = String.Empty
        Dim strSecond As String = String.Empty
        Dim intCounter, intLeadCounter As Integer
        Dim strCounter As String
        Dim strFormula As String = String.Empty

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

        For i = 0 To Columns.Count - 1
            Me(i, IndexTotal).ReadOnly = True
        Next

        For i = 0 To IndexPopulationTarget - 1
            Rows(i).Visible = True
        Next
        If CurrentIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
            Rows(IndexScore).Visible = False
        Else
            Me(0, IndexScore).ReadOnly = True
            Me(1, IndexScore).ReadOnly = True
        End If

        If CurrentIndicator.Registration <> Indicator.RegistrationOptions.BeneficiaryLevel Then
            Rows(IndexPopulationTarget).Visible = False
            Rows(IndexPopulationScore).Visible = False
        Else
            Me(0, IndexPopulationTarget).ReadOnly = True
            If CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then Me(1, IndexPopulationTarget).ReadOnly = True

            If CurrentIndicator.QuestionType <> Indicator.QuestionTypes.Ratio Then
                For i = 0 To Columns.Count - 1
                    Me(i, IndexPopulationScore).ReadOnly = True
                Next
            Else
                Rows(IndexPopulationScore).Visible = False
            End If
        End If
    End Sub

    Protected Overridable Sub LoadItems()
        Me.Grid.Clear()

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                For intIndex = 0 To .Statements.Count - 1
                    Dim BaselineDoubleValue As DoubleValue = .Baseline.DoubleValues(intIndex)
                    Dim TargetDoubleValues As New DoubleValues

                    For Each selTarget As Target In CurrentIndicator.Targets
                        TargetDoubleValues.Add(selTarget.DoubleValues(intIndex))
                    Next

                    Me.Grid.Add(New TargetFormulaGridRow(.Statements(intIndex), BaselineDoubleValue, TargetDoubleValues))
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

        Me.RowCount = Me.Grid.Count + 4
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
        Dim selRow As TargetFormulaGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As TargetFormulaGridRow)
        selRow.StatementImageDirty = True
        selRow.StatementHeight = 0
    End Sub

    Protected Overridable Sub ReloadImages()
        Dim selRow As TargetFormulaGridRow
        For i = 0 To CurrentIndicator.Statements.Count - 1
            selRow = Me.Grid(i)
            ReloadImage_Response(selRow)
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImage_Response(ByVal selRow As TargetFormulaGridRow)
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
        Dim selGridRow As TargetFormulaGridRow = Me.Grid(RowIndex)

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
        Dim selGridRow As TargetFormulaGridRow = Me.Grid(RowIndex)

        If selGridRow IsNot Nothing AndAlso selGridRow.StatementHeight > intRowHeight Then intRowHeight = selGridRow.StatementHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)

        MoveCurrentCell()
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
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As Object, ByVal boolShowScore As Boolean)
        Dim intColIndex, intRowIndex As Integer

        Select Case selItem.GetType
            Case GetType(Baseline)
                intColIndex = 1
                If boolShowScore = False Then intRowIndex = IndexTotal Else intRowIndex = IndexScore
            Case GetType(Target)
                Dim selTarget As Target = TryCast(selItem, Target)

                If selTarget IsNot Nothing Then
                    intColIndex = CurrentIndicator.Targets.IndexOf(selTarget)
                    intColIndex += 2
                    If boolShowScore = False Then intRowIndex = IndexTotal Else intRowIndex = IndexScore
                End If
            Case GetType(PopulationTarget)
                Dim selPopulationTarget As PopulationTarget = TryCast(selItem, PopulationTarget)

                If selPopulationTarget IsNot Nothing Then
                    intColIndex = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget)
                    intColIndex += 2
                    intRowIndex = IndexPopulationTarget
                End If
            Case GetType(DoubleValue)
                Dim selValue As DoubleValue = TryCast(selItem, DoubleValue)

                If CurrentIndicator.Baseline.Guid = selValue.ParentGuid Then
                    intColIndex = 1
                    intRowIndex = CurrentIndicator.Baseline.DoubleValues.IndexOf(selValue)
                Else
                    Dim selTarget As Target
                    For i = 0 To CurrentIndicator.Targets.Count - 1
                        selTarget = CurrentIndicator.Targets(i)

                        If selTarget.Guid = selValue.ParentGuid Then
                            intColIndex = i + 2
                            intRowIndex = selTarget.DoubleValues.IndexOf(selValue)
                        End If
                    Next
                End If
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
        End If
    End Sub

    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex
        Me.ReloadImages()
    End Sub

    Private Sub MoveCurrentCell()
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex

        If intRowIndex < IndexTotal Then
            If intColumnIndex = 0 Then intColumnIndex = 1
        ElseIf intRowIndex >= IndexTotal Then
            If intColumnIndex <= 1 Then intColumnIndex = 2
        End If
        If intRowIndex = IndexTotal Or intRowIndex = IndexPopulationScore Then
            intRowIndex -= 1
        End If
        CurrentCell = Nothing
        CurrentCell = Me(intColumnIndex, intRowIndex)

        Invalidate()
    End Sub

    Public Sub Edit()
        If CurrentCell.ReadOnly = True Then MoveCurrentCell()

        Me.BeginEdit(True)
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

    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)

    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)

        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim RowTmp As TargetFormulaGridRow = Nothing


        If intRowIndex < IndexTotal Then
            e.Value = OnCellValueNeeded_TargetValues(intRowIndex, intColumnIndex)
        ElseIf intRowIndex = IndexTotal Then
            e.Value = OnCellValueNeeded_TotalValues(intColumnIndex)
        ElseIf intRowIndex = IndexScore Then
            e.Value = OnCellValueNeeded_TargetScores(intColumnIndex)
        ElseIf intRowIndex = IndexPopulationTarget Then
            e.Value = OnCellValueNeeded_PopulationTargets(intColumnIndex)
        ElseIf intRowIndex = IndexPopulationScore Then
            e.Value = OnCellValueNeeded_TotalPopulationScores(e.ColumnIndex)
        End If

    End Sub

    Private Function OnCellValueNeeded_TargetValues(ByVal intRowIndex As Integer, ByVal intColumnIndex As Integer) As String
        Dim RowTmp As TargetFormulaGridRow = Nothing
        Dim strValue As String = String.Empty

        ' Store a reference to the grid row for the row being painted.
        If intRowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(intRowIndex), TargetFormulaGridRow)
        End If
        If RowTmp Is Nothing Then Return String.Empty

        Dim intNrDecimals As Integer = RowTmp.Statement.ValuesDetail.NrDecimals
        Dim strUnit As String = RowTmp.Statement.ValuesDetail.Unit

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case intColumnIndex
            Case 0
                If intRowIndex < IndexTotal Then
                    strValue = RowTmp.FirstLabel
                End If
            Case 1
                If intRowIndex < IndexTotal Then
                    strValue = DisplayAsUnit(RowTmp.BaselineDoubleValue.Value, intNrDecimals, strUnit)
                End If
            Case Else
                If intRowIndex < IndexTotal Then
                    Dim intTargetIndex As Integer = intColumnIndex - 2
                    strValue = DisplayAsUnit(RowTmp.TargetDoubleValues(intTargetIndex).Value, intNrDecimals, strUnit)
                End If
        End Select

        Return strValue
    End Function

    Private Function OnCellValueNeeded_TotalValues(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        With CurrentIndicator
            Dim intNrDecimals As Integer = .ValuesDetail.NrDecimals
            Dim strUnit As String = .ValuesDetail.Unit

            If intColumnIndex = 0 Then
                strValue = String.Format("{0} ({1})", LANG_Total, .ValuesDetail.Formula)
            ElseIf intColumnIndex = 1 Then
                strValue = CurrentIndicator.GetBaselineFormattedValue
            ElseIf intColumnIndex > 1 Then
                Dim intTargetIndex As Integer = intColumnIndex - 2
                Dim selTarget As Target = .Targets(intTargetIndex)

                If selTarget IsNot Nothing Then
                    strValue = CurrentIndicator.GetTargetFormattedValue(intTargetIndex)
                End If
            End If
        End With

        Return strValue
    End Function

    Private Function OnCellValueNeeded_TargetScores(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex = 0 Then
            strValue = LANG_ScoreValueBeneficiary
        ElseIf intColumnIndex > 1 Then
            Dim intTargetIndex As Integer = intColumnIndex - 2
            strValue = CurrentIndicator.GetTargetFormattedScore(intTargetIndex)
        End If

        Return strValue
    End Function

    Private Function OnCellValueNeeded_PopulationTargets(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex = 0 Then
            strValue = LANG_PopulationTargetText
        ElseIf intColumnIndex > 1 Then
            Dim intTargetIndex As Integer = intColumnIndex - 2
            Dim selTargetPopulation As PopulationTarget = CurrentIndicator.PopulationTargets(intTargetIndex)

            If selTargetPopulation IsNot Nothing Then
                strValue = DisplayAsUnit(selTargetPopulation.Percentage, CurrentIndicator.ValuesDetail.NrDecimals, "%")
            End If
        End If

        Return strValue
    End Function

    Private Function OnCellValueNeeded_TotalPopulationScores(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex = 0 Then
            If Me.TargetGroup IsNot Nothing Then
                Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

                strValue = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
            Else
                strValue = LANG_ScoreValueTargetGroup
            End If
        ElseIf intColumnIndex = 1 Then
            If CurrentIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then _
                strValue = CurrentIndicator.GetPopulationBaselineFormattedScore
        Else
            Dim intTargetIndex As Integer = intColumnIndex - 2

            strValue = CurrentIndicator.GetPopulationTargetFormattedScore(intTargetIndex)
        End If

        Return strValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        If e.RowIndex < IndexScore Then
            OnCellValuePushed_TargetValues(e)
        ElseIf e.RowIndex = IndexScore Then
            If e.ColumnIndex > 1 Then OnCellValuePushed_Score(e)
        ElseIf e.RowIndex = IndexPopulationTarget Then
            If e.ColumnIndex > 0 Then OnCellValuePushed_PopulationTarget(e)
        End If
    End Sub

    Private Sub OnCellValuePushed_TargetValues(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As TargetFormulaGridRow
        Dim strColName As String
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim strCellValue As String = String.Empty
        Dim dblValue As Double

        If Me.EditRow Is Nothing Then
            Dim CurrentGridRow As TargetFormulaGridRow = CType(Me.Grid(intRowIndex), TargetFormulaGridRow)

            Me.EditRow = New TargetFormulaGridRow
            With EditRow
                .Statement = CurrentGridRow.Statement
                .StatementHeight = CurrentGridRow.StatementHeight
                .StatementImage = CurrentGridRow.StatementImage
                .StatementImageDirty = CurrentGridRow.StatementImageDirty
                .BaselineDoubleValue = CurrentGridRow.BaselineDoubleValue
                .TargetDoubleValues = CurrentGridRow.TargetDoubleValues
            End With
        End If
        RowTmp = Me.EditRow
        Me.EditRowFlag = e.RowIndex

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
                Double.TryParse(e.Value, dblValue)
                With RowTmp.Statement.ValuesDetail.ValueRange
                    dblValue = .MakeWithinRange(dblValue)
                End With
        End Select

        Select Case intColumnIndex
            Case 0
                RowTmp.FirstLabel = strCellValue
                RowTmp.StatementImageDirty = True
            Case 1
                UndoRedo.UndoBuffer_Initialise(RowTmp.BaselineDoubleValue, "Value", RowTmp.BaselineDoubleValue.Value)
                RowTmp.BaselineDoubleValue.Value = dblValue
                UndoRedo.ValueChanged(dblValue)

                UpdateScore = True
            Case Else
                Dim intTargetIndex As Integer = intColumnIndex - 2

                UndoRedo.UndoBuffer_Initialise(RowTmp.TargetDoubleValues(intTargetIndex), "Value", RowTmp.TargetDoubleValues(intTargetIndex).Value)
                RowTmp.TargetDoubleValues(intTargetIndex).Value = dblValue
                UndoRedo.ValueChanged(dblValue)

                UpdateScore = True
        End Select

        'CurrentUndoList.AddToUndoList()
    End Sub

    Private Sub OnCellValuePushed_Score(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim intTargetIndex As Integer = e.ColumnIndex - 2
        Dim dblValue As Double

        If Double.TryParse(e.Value, dblValue) Then
            If intTargetIndex >= 0 Then

                Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)

                If selTarget IsNot Nothing Then
                    UndoRedo.UndoBuffer_Initialise(selTarget, "Score", selTarget.Score)
                    selTarget.Score = dblValue
                    UndoRedo.ValueChanged(dblValue)
                End If

            End If
        End If
    End Sub

    Private Sub OnCellValuePushed_PopulationTarget(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim intPopulationTargetIndex As Integer = e.ColumnIndex - 2
        Dim dblValue As Double

        If Double.TryParse(e.Value, dblValue) Then
            If intPopulationTargetIndex < 0 Then
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.Baseline, "PopulationPercentage", CurrentIndicator.Baseline.PopulationPercentage)
                CurrentIndicator.Baseline.PopulationPercentage = dblValue
                UndoRedo.ValueChanged(dblValue)
            Else
                Dim selPopulationTarget As PopulationTarget = CurrentIndicator.PopulationTargets(intPopulationTargetIndex)

                If selPopulationTarget IsNot Nothing Then
                    UndoRedo.UndoBuffer_Initialise(selPopulationTarget, "Percentage", selPopulationTarget.Percentage)
                    selPopulationTarget.Percentage = dblValue
                End If

            End If
        End If
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New TargetFormulaGridRow
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
        Me.EditRow = New TargetFormulaGridRow()
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
            '
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then
            ' Save the modified planning grid row in grid.
            Me.Grid(e.RowIndex) = Me.EditRow
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        CurrentIndicator.CalculateValues()

        InvalidateRow(IndexTotal)
        MyBase.OnRowValidated(e)
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex >= 0 And intRowIndex < IndexTotal And intColumnIndex = 0 Then
            Dim selGridRow As TargetFormulaGridRow = Grid(intRowIndex)
            If selGridRow IsNot Nothing AndAlso selGridRow.StatementImage IsNot Nothing Then
                Dim intRowHeight As Integer = selGridRow.StatementImage.Height
                If Rows(intRowIndex).Height < intRowHeight Then Rows(intRowIndex).Height = intRowHeight
            End If

            CellPainting_RichText(selGridRow, rCell, CellGraphics)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
        ElseIf intRowIndex = IndexTotal And intColumnIndex >= 0 Then
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            e.PaintContent(rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
        ElseIf intRowIndex = IndexScore And (intColumnIndex = 0 Or intColumnIndex = 1) Then
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            e.PaintContent(rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
        ElseIf intRowIndex = IndexPopulationTarget Then
            If intColumnIndex = 0 Or intColumnIndex = 1 Then '(intColumnIndex = 1 And CurrentIndicator.ScoringSystem = Indicator.ScoringSystems.Score) Then
                CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
                e.PaintContent(rCell)
                e.Paint(rCell, DataGridViewPaintParts.Border)

                e.Handled = True
            End If
        ElseIf intRowIndex = IndexPopulationScore And intColumnIndex >= 0 Then
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            e.PaintContent(rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
        End If

    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As TargetFormulaGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow IsNot Nothing Then
            If Me.TextSelectionIndex = TextSelectionValues.SelectAll Then
                CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
            Else
                CellGraphics.FillRectangle(Brushes.White, rCell)
            End If

            If selGridRow.StatementImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StatementImage.Width, selGridRow.StatementImage.Height)
                CellGraphics.DrawImage(selGridRow.StatementImage, rImage)
            End If
        End If
    End Sub
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
