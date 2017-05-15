Public Class DataGridViewTargetsRanking
    Inherits DataGridView

    Friend WithEvents EditBox As NumericTextBox

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colBaseline As New DataGridViewComboBoxColumn

    Private EditPopulationIndex As Integer

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

    Private ReadOnly Property TargetGroupNumber As Integer
        Get
            If Me.TargetGroup IsNot Nothing Then
                Return Me.TargetGroup.Number
            Else
                Return 0
            End If
        End Get
    End Property

    Private ReadOnly Property IndexPopulationTarget As Integer
        Get
            Return CurrentIndicator.ResponseClasses.Count
        End Get
    End Property

    Private ReadOnly Property IndexPopulationScore As Integer
        Get
            Return CurrentIndicator.ResponseClasses.Count + 1
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New(ByVal currentindicator As Indicator, ByVal targetdeadlinessection As TargetDeadlinesSection, ByVal targetgroup As TargetGroup)
        'datagridview settings
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

        'load values

        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlinesSection = targetdeadlinessection
        Me.TargetGroup = targetgroup

        Dim intResponseClassesCount As Integer = Me.CurrentIndicator.ResponseClasses.Count

        With Me.CurrentIndicator
            VerifyNumberOfDoubleValues(.Baseline.DoubleValues, intResponseClassesCount)
            For Each selTarget As Target In .Targets
                VerifyNumberOfDoubleValues(selTarget.DoubleValues, intResponseClassesCount)
            Next
        End With

        LoadColumns()
    End Sub

    Private Sub VerifyNumberOfDoubleValues(ByVal objDoubleValues As DoubleValues, ByVal intResponseClassesCount As Integer)
        If objDoubleValues.Count <> intResponseClassesCount Then
            objDoubleValues.Clear()
            For i = 0 To intResponseClassesCount - 1
                objDoubleValues.Add(New DoubleValue)
            Next
        End If
    End Sub

    Private Sub LoadColumns()
        Columns.Clear()

        With colBaseline
            .Name = "Baseline"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ValueType = GetType(Integer)

            For i = 0 To CurrentIndicator.ResponseClasses.Count
                .Items.Add(i)
            Next
        End With
        Me.Columns.Add(colBaseline)

        Reload()
    End Sub

    Private Sub LoadTargetColumns()
        If Columns.Count > 1 Then
            For i = 1 To Columns.Count - 1
                Columns.RemoveAt(1)
            Next
        End If

        For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
            Dim colTarget As New DataGridViewComboBoxColumn
            With colTarget
                .Name = selTargetDeadline.Deadline
                .MinimumWidth = 100
                .FillWeight = 20
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .ValueType = GetType(Integer)
                For i = 0 To CurrentIndicator.ResponseClasses.Count
                    .Items.Add(i)
                Next
            End With
            Columns.Add(colTarget)
        Next
    End Sub

    Private Sub SetColumnHeadersText()
        Dim strHeaderText As String = String.Empty
        Dim strDate As String
        Dim strScore As String = String.Empty
        Dim intTargetIndex, intColumnIndex As Integer
        Dim strBaselineScore As String = 0.ToString

        CurrentIndicator.CalculateScores()

        'header text of baseline column
        strBaselineScore = CurrentIndicator.GetBaselineFormattedScore

        'set header text
        Dim strBaseline As String = LANG_Baseline 'String.Format("{0}{1}{2}: {3}", LANG_Baseline, vbCrLf, LANG_ScoringValue, strBaselineScore)

        If Me.Columns.Count > 0 Then
            Columns(0).HeaderText = strBaseline
        End If

        'header text of target columns
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

            ''score
            'strScore = CurrentIndicator.GetTargetFormattedScore(intTargetIndex)
            'strScore = String.Format("{0}: {1}", LANG_ScoringValue, strScore)


            'set header text
            'strHeaderText = String.Format("{0} {1}{2}{3}", LANG_Target, strDate, vbCrLf, strScore)
            strHeaderText = String.Format("{0} {1}", LANG_Target, strDate)

            intColumnIndex = intTargetIndex + 1
            If intColumnIndex <= Me.Columns.Count - 1 Then
                Columns(intColumnIndex).HeaderText = strHeaderText
            End If
        Next
    End Sub

    Private Sub SetRowHeadersText()
        Dim intRowIndex As Integer

        Me.RowCount = CurrentIndicator.ResponseClasses.Count '+ 2

        For intRowIndex = 0 To CurrentIndicator.ResponseClasses.Count - 1
            Rows(intRowIndex).HeaderCell.Value = CurrentIndicator.ResponseClasses(intRowIndex).ClassName
        Next
        'intRowIndex = CurrentIndicator.ResponseClasses.Count
        'Rows(intRowIndex).HeaderCell.Value = LANG_PopulationTargetText

        'intRowIndex += 1
        'If Me.TargetGroup IsNot Nothing Then
        '    Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

        '    Rows(intRowIndex).HeaderCell.Value = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
        'Else
        '    Rows(intRowIndex).HeaderCell.Value = LANG_ScoreValueTargetGroup
        'End If

        'If CurrentIndicator.Registration <> Indicator.RegistrationOptions.BeneficiaryLevel Then
        '    Rows(IndexPopulationTarget).Visible = False
        '    Rows(IndexPopulationScore).Visible = False
        'Else
        '    For i = 0 To Columns.Count - 1
        '        Me(i, IndexPopulationScore).ReadOnly = True
        '    Next
        'End If
    End Sub

    Public Sub Reload()
        'remember current cell location
        objCellLocation = CurrentCellAddress
        If objCellLocation.X < 0 Then
            'If CurrentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            '    objCellLocation.X = 1
            'Else
            objCellLocation.X = 0
            'End If
        End If
        If objCellLocation.Y < 0 Then objCellLocation.Y = 0

        '(re)load target columns and rows
        Me.SuspendLayout()

        Rows.Clear()
        LoadTargetColumns()
        SetColumnHeadersText()
        SetRowHeadersText()

        Me.Invalidate()
        Me.ResumeLayout()

        'set current cell location to what it was before
        If objCellLocation.X <= Me.ColumnCount - 1 And objCellLocation.Y <= Me.RowCount - 1 Then
            If Me(objCellLocation.X, objCellLocation.Y).Displayed = True Then _
                CurrentCell = Me(objCellLocation.X, objCellLocation.Y)
        End If
    End Sub
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

        'Dim intRowIndex As Integer = e.RowIndex
        'Dim intColumnIndex As Integer = e.ColumnIndex

        'If EditBox IsNot Nothing Then
        '    SetPopulationTarget()
        'End If
        'If intRowIndex = IndexPopulationTarget Then
        '    If EditBox Is Nothing Then _
        '        EditPopulationPercentage(intColumnIndex, intRowIndex)
        'End If
    End Sub

    Protected Sub EditBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles EditBox.KeyUp
        'If e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Then
        '    Me.Select()
        'End If
    End Sub

    Private Sub SetPopulationTarget()
        'Dim strValue As String = EditBox.Text
        'Dim dblValue As Double = ParseDouble(strValue)

        'Select Case EditPopulationIndex
        '    Case -1
        '        CurrentIndicator.Baseline.PopulationPercentage = dblValue
        '    Case Else
        '        CurrentIndicator.PopulationTargets(EditPopulationIndex).Percentage = dblValue
        'End Select

        'Me.Controls.Remove(EditBox)
        'EditBox = Nothing
        'Reload()

    End Sub

    Protected Overridable Sub EditBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditBox.Leave
        
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As Object)
        Dim intColIndex, intRowIndex As Integer

        Select Case selItem.GetType
            Case GetType(DoubleValue)
                Dim selValue As DoubleValue = TryCast(selItem, DoubleValue)

                If CurrentIndicator.Baseline.Guid = selValue.ParentGuid Then
                    intColIndex = 0
                    intRowIndex = CurrentIndicator.Baseline.DoubleValues.IndexOf(selValue)
                Else
                    Dim selTarget As Target
                    For i = 0 To CurrentIndicator.Targets.Count - 1
                        selTarget = CurrentIndicator.Targets(i)

                        If selTarget.Guid = selValue.ParentGuid Then
                            intColIndex = i + 1
                            intRowIndex = selTarget.DoubleValues.IndexOf(selValue)
                        End If
                    Next
                End If
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
        End If
    End Sub

    Protected Overridable Sub EditPopulationPercentage(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Dim rCell As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex, intRowIndex, False)
        Dim ptLocation As Point
        Dim dblValue As Double

        Select Case intColumnIndex
            Case 0
                dblValue = CurrentIndicator.Baseline.PopulationPercentage
                EditPopulationIndex = -1
            Case Else
                Dim intTargetIndex As Integer = intColumnIndex - 1

                dblValue = CurrentIndicator.PopulationTargets(intTargetIndex).Percentage
                EditPopulationIndex = intTargetIndex
        End Select

        rCell.Width -= 2
        ptLocation = New Point(rCell.X + 1, rCell.Y + 1)

        EditBox = New NumericTextBox
        With EditBox
            .Size = New Size(rCell.Width, rCell.Height)
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
            '    Dim objCurrentCell As DataGridViewCell = CurrentCell
            '    CurrentCell = Nothing
            '    CurrentCell = objCurrentCell
        End If
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
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)

        Dim intRowIndex As Integer = e.RowIndex

        If intRowIndex < IndexPopulationTarget Then
            e.Value = OnCellValueNeeded_Targets(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    Private Function OnCellValueNeeded_Targets(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer) As Integer
        Dim intValue As Integer

        If intColumnIndex = 0 Then
            intValue = CurrentIndicator.Baseline.DoubleValues(intRowIndex).Value
        Else
            Dim intTargetIndex As Integer = intColumnIndex - 1
            Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)

            intValue = selTarget.DoubleValues(intRowIndex).Value
        End If

        Return intValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim selDoubleValue As DoubleValue

        If intRowIndex < IndexPopulationTarget Then
            Dim dblValue As Double = ParseDouble(e.Value)
            Dim dblOldValue As Double

            If intColumnIndex = 0 Then
                selDoubleValue = CurrentIndicator.Baseline.DoubleValues(intRowIndex)

                dblOldValue = selDoubleValue.Value
                UndoRedo.UndoBuffer_Initialise(selDoubleValue, "Value", dblOldValue)
                CurrentIndicator.Baseline.DoubleValues(intRowIndex).Value = dblValue

                UndoRedo.ValueChanged(dblValue)
            Else
                Dim intTargetIndex As Integer = intColumnIndex - 1
                Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)
                selDoubleValue = selTarget.DoubleValues(intRowIndex)

                dblOldValue = selDoubleValue.Value
                UndoRedo.UndoBuffer_Initialise(selDoubleValue, "Value", dblOldValue)
                selTarget.DoubleValues(intRowIndex).Value = dblValue

                UndoRedo.ValueChanged(dblValue)
            End If

            SetColumnHeadersText()
            Invalidate()
        End If
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex >= 0 Then
            If intColumnIndex > 0 AndAlso e.Value = 0 Then
                e.PaintBackground(e.CellBounds, False)
                e.Paint(e.CellBounds, DataGridViewPaintParts.ContentBackground)
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub OnCellPainting_PopulationTargets(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs, ByVal intColumnIndex As Integer)
        'Dim strValue As String = String.Empty
        'Dim CellGraphics As Graphics = e.Graphics
        'Dim rCell As Rectangle = e.CellBounds

        'With CurrentIndicator
        '    If intColumnIndex = 0 Then
        '        strValue = DisplayAsUnit(.Baseline.PopulationPercentage, 2, "%")
        '    Else

        '        Dim intTargetPopulationIndex As Integer = intColumnIndex - 1
        '        Dim selTargetPopulation As PopulationTarget = .PopulationTargets(intTargetPopulationIndex)

        '        If selTargetPopulation IsNot Nothing Then
        '            strValue = DisplayAsUnit(selTargetPopulation.Percentage, 2, "%")
        '        End If
        '    End If
        'End With

        'Dim sfValue As New StringFormat
        'sfValue.LineAlignment = StringAlignment.Center
        'sfValue.Alignment = StringAlignment.Center
        'CellGraphics.FillRectangle(Brushes.White, rCell)
        ''e.Paint(rCell, DataGridViewPaintParts.Background)
        'CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rCell, sfValue)
        'e.Paint(rCell, DataGridViewPaintParts.Border)
    End Sub

    Private Sub OnCellPainting_PopulationScores(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs, ByVal intColumnIndex As Integer)
        'Dim strValue As String = String.Empty
        'Dim CellGraphics As Graphics = e.Graphics
        'Dim rCell As Rectangle = e.CellBounds

        'With CurrentIndicator
        '    If intColumnIndex = 0 Then
        '        strValue = .GetPopulationBaselineFormattedScore()
        '    Else
        '        Dim intTargetIndex As Integer = intColumnIndex - 1
        '        strValue = .GetPopulationTargetFormattedScore(intTargetIndex)
        '    End If
        'End With

        'Dim sfValue As New StringFormat
        'sfValue.LineAlignment = StringAlignment.Center
        'sfValue.Alignment = StringAlignment.Center

        'CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
        'CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rCell, sfValue)
        'e.Paint(rCell, DataGridViewPaintParts.Border)
    End Sub

    Private Sub DataGridViewTargetsRanking_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles Me.EditingControlShowing
        'If CurrentCell.RowIndex = IndexPopulationTarget Then Me.CancelEdit()
    End Sub
#End Region

    
End Class
