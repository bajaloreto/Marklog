Public Class DataGridViewTargetsClasses
    Inherits DataGridViewBaseClass

    Friend WithEvents EditBox As New NumericTextBox

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colBaseline As New DataGridViewCheckBoxColumn

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

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim intIndex As Integer = CurrentCell.ColumnIndex

                If intIndex = 0 Then
                    selObject = CurrentIndicator.Baseline
                Else
                    selObject = CurrentIndicator.Targets(intIndex - 1)
                End If

            End If
            Return selObject
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New(ByVal currentindicator As Indicator, ByVal targetdeadlinessection As TargetDeadlinesSection, ByVal targetgroup As TargetGroup)
        'datagridview settings
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

        'load values

        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlinesSection = targetdeadlinessection
        Me.TargetGroup = targetgroup

        Dim intResponseClassesCount As Integer = Me.CurrentIndicator.ResponseClasses.Count

        With Me.CurrentIndicator
            VerifyNumberOfBooleanValues(.Baseline.BooleanValues, intResponseClassesCount)
            For Each selTarget As Target In .Targets
                VerifyNumberOfBooleanValues(selTarget.BooleanValues, intResponseClassesCount)
            Next
        End With

        LoadColumns()
    End Sub

    Private Sub VerifyNumberOfBooleanValues(ByVal objBooleanValues As BooleanValues, ByVal intResponseClassesCount As Integer)
        If objBooleanValues.Count <> intResponseClassesCount Then
            objBooleanValues.Clear()
            For i = 0 To intResponseClassesCount - 1
                objBooleanValues.Add(New BooleanValue)
            Next
        End If
    End Sub

    Private Sub LoadColumns()
        Columns.Clear()

        With colBaseline
            .Name = "Baseline"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            Dim colTarget As New DataGridViewCheckBoxColumn
            With colTarget
                .Name = selTargetDeadline.Deadline
                .MinimumWidth = 100
                .FillWeight = 20
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
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
        Dim strBaseline As String = String.Format("{0}{1}{2}: {3}", LANG_Baseline, vbCrLf, LANG_ScoringValue, strBaselineScore)

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

            'score
            strScore = CurrentIndicator.GetTargetFormattedScore(intTargetIndex)
            strScore = String.Format("{0}: {1}", LANG_ScoringValue, strScore)


            'set header text
            strHeaderText = String.Format("{0} {1}{2}{3}", LANG_Target, strDate, vbCrLf, strScore)

            intColumnIndex = intTargetIndex + 1
            If intColumnIndex <= Me.Columns.Count - 1 Then
                Columns(intColumnIndex).HeaderText = strHeaderText
            End If
        Next
    End Sub

    Private Sub SetRowHeadersText()
        Dim intRowIndex As Integer

        Me.RowCount = CurrentIndicator.ResponseClasses.Count + 2

        For intRowIndex = 0 To CurrentIndicator.ResponseClasses.Count - 1
            Rows(intRowIndex).HeaderCell.Value = CurrentIndicator.ResponseClasses(intRowIndex).ClassName
        Next
        intRowIndex = CurrentIndicator.ResponseClasses.Count
        Rows(intRowIndex).HeaderCell.Value = LANG_PopulationTargetText

        intRowIndex += 1
        If Me.TargetGroup IsNot Nothing Then
            Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

            Rows(intRowIndex).HeaderCell.Value = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
        Else
            Rows(intRowIndex).HeaderCell.Value = LANG_ScoreValueTargetGroup
        End If

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
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(e)

        MoveCurrentCell()
    End Sub

    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)

        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex = IndexPopulationTarget Then
            EditPopulationPercentage(intColumnIndex, intRowIndex)
        End If
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
            Case GetType(PopulationTarget)
                Dim selPopulationTarget As PopulationTarget = TryCast(selItem, PopulationTarget)

                If selPopulationTarget IsNot Nothing Then
                    intColIndex = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget)
                    intColIndex += 1
                    intRowIndex = IndexPopulationTarget
                End If
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Protected Overridable Sub EditPopulationPercentage(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer)
        Dim rCell As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex, intRowIndex, False)
        Dim ptLocation As Point
        Dim dblValue As Double

        Select Case intColumnIndex
            Case 0
                dblValue = CurrentIndicator.Baseline.PopulationPercentage
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.Baseline, "PopulationPercentage", dblValue)
                EditPopulationIndex = -1
            Case Else
                Dim intTargetIndex As Integer = intColumnIndex - 1

                dblValue = CurrentIndicator.PopulationTargets(intTargetIndex).Percentage
                UndoRedo.UndoBuffer_Initialise(CurrentIndicator.PopulationTargets(intTargetIndex), "Percentage", dblValue)
                EditPopulationIndex = intTargetIndex
        End Select

        rCell.Width -= 2
        ptLocation = New Point(rCell.X + 1, rCell.Y + 1)

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
            Dim objCurrentCell As DataGridViewCell = CurrentCell
            CurrentCell = Nothing
            CurrentCell = objCurrentCell
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

    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)

    End Sub
#End Region

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)

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

        If intRowIndex < IndexPopulationTarget Then
            e.Value = OnCellValueNeeded_Targets(e.ColumnIndex, e.RowIndex)
        End If
    End Sub

    Private Function OnCellValueNeeded_Targets(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer) As Boolean
        Dim boolValue As Boolean

        If intColumnIndex = 0 Then
            boolValue = CurrentIndicator.Baseline.BooleanValues(intRowIndex).Value
        Else
            Dim intTargetIndex As Integer = intColumnIndex - 1
            Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)

            boolValue = selTarget.BooleanValues(intRowIndex).Value
        End If

        Return boolValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim boolSelected As Boolean
        If Boolean.TryParse(e.Value, boolSelected) Then


            If intRowIndex < IndexPopulationTarget Then
                If intColumnIndex = 0 Then
                    OnCellValuePushed_Baseline(intRowIndex, boolSelected)
                Else
                    OnCellValuePushed_Targets(intColumnIndex, intRowIndex, boolSelected)
                End If
            End If
        End If
    End Sub

    Private Sub OnCellValuePushed_Baseline(ByVal intRowIndex As Integer, ByVal boolSelected As Boolean)
        Dim boolSum As Boolean = CurrentIndicator.AddClassValuesForTotal
        Dim objOldValues As BooleanValues

        With CurrentIndicator.Baseline
            Using copier As New ObjectCopy
                objOldValues = copier.CopyCollection(.BooleanValues)
            End Using

            .BooleanValues.SetValue(intRowIndex, boolSelected, boolSum)
            UndoRedo.BooleanValueChecked(CurrentIndicator.Baseline, intRowIndex, objOldValues, .BooleanValues)
        End With

        Invalidate()
        SetColumnHeadersText()
    End Sub

    Private Sub OnCellValuePushed_Targets(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer, ByVal boolSelected As Boolean)
        Dim boolSum As Boolean = CurrentIndicator.AddClassValuesForTotal
        Dim intTargetIndex As Integer = intColumnIndex - 1
        Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)
        Dim objOldValues As BooleanValues

        Using copier As New ObjectCopy
            objOldValues = copier.CopyCollection(selTarget.BooleanValues)
        End Using

        selTarget.BooleanValues.SetValue(intRowIndex, boolSelected, boolSum)
        UndoRedo.BooleanValueChecked(selTarget, intRowIndex, objOldValues, selTarget.BooleanValues)

        SetFutureTargets(selTarget, intRowIndex, boolSelected)

        Invalidate()
        SetColumnHeadersText()
    End Sub

    Public Sub SetFutureTargets(ByVal selTarget As Target, ByVal intBooleanValueIndex As Integer, ByVal boolSelected As Boolean)
        Dim boolSum As Boolean = CurrentIndicator.AddClassValuesForTotal

        If boolSum = False Then
            If boolSelected = True Then

                'set targets to the right to the same value
                Dim intIndex As Integer = CurrentIndicator.Targets.IndexOf(selTarget) + 1
                If intIndex <= CurrentIndicator.Targets.Count - 1 Then
                    For i = intIndex To CurrentIndicator.Targets.Count - 1
                        With CurrentIndicator.Targets(i)
                            .BooleanValues.SetValue(intBooleanValueIndex, boolSelected, boolSum)
                        End With
                    Next
                End If
            End If
        End If
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intRowIndex = IndexPopulationTarget Then
            If intColumnIndex >= 0 Then
                OnCellPainting_PopulationTargets(e, intColumnIndex)

                e.Handled = True
            End If
        ElseIf intRowIndex = IndexPopulationScore Then
            If intColumnIndex >= 0 Then
                OnCellPainting_PopulationScores(e, intColumnIndex)

                e.Handled = True
            End If
        End If
    End Sub

    Private Sub OnCellPainting_PopulationTargets(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs, ByVal intColumnIndex As Integer)
        Dim strValue As String = String.Empty
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds

        With CurrentIndicator
            If intColumnIndex = 0 Then
                strValue = DisplayAsUnit(.Baseline.PopulationPercentage, 2, "%")
            Else

                Dim intTargetPopulationIndex As Integer = intColumnIndex - 1
                Dim selTargetPopulation As PopulationTarget = .PopulationTargets(intTargetPopulationIndex)

                If selTargetPopulation IsNot Nothing Then
                    strValue = DisplayAsUnit(selTargetPopulation.Percentage, 2, "%")
                End If
            End If
        End With

        Dim sfValue As New StringFormat
        sfValue.LineAlignment = StringAlignment.Center
        sfValue.Alignment = StringAlignment.Center

        e.PaintBackground(rCell, True)
        CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rCell, sfValue)
        e.Paint(rCell, DataGridViewPaintParts.Border)
    End Sub

    Private Sub OnCellPainting_PopulationScores(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs, ByVal intColumnIndex As Integer)
        Dim strValue As String = String.Empty
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds

        With CurrentIndicator
            If intColumnIndex = 0 Then
                strValue = .GetPopulationBaselineFormattedScore()
            Else
                Dim intTargetIndex As Integer = intColumnIndex - 1
                strValue = .GetPopulationTargetFormattedScore(intTargetIndex)
            End If
        End With

        Dim sfValue As New StringFormat
        sfValue.LineAlignment = StringAlignment.Center
        sfValue.Alignment = StringAlignment.Center

        CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
        CellGraphics.DrawString(strValue, Me.Font, Brushes.Black, rCell, sfValue)
        e.Paint(rCell, DataGridViewPaintParts.Border)
    End Sub
#End Region
End Class
