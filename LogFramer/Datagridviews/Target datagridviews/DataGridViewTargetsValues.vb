Public Class DataGridViewTargetsValues
    Inherits DataGridView

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargets As Targets
    Private objPopulationTargets As PopulationTargets
    Private objTargetGroup As TargetGroup
    Private objBaseline As ResponseValue

    Private objCellLocation As New Point
    Private colBaseline As New DataGridViewTextBoxColumn

#Region "Properties"
    Public Property CurrentIndicator As Indicator
        Get
            Return objCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            objCurrentIndicator = value
            If objCurrentIndicator.ValuesDetail Is Nothing Then objCurrentIndicator.ValuesDetail = New ValuesDetail
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

    Public Property Targets As Targets
        Get
            Return objTargets
        End Get
        Set(ByVal value As Targets)
            objTargets = value
        End Set
    End Property

    Public Property PopulationTargets As PopulationTargets
        Get
            Return objPopulationTargets
        End Get
        Set(ByVal value As PopulationTargets)
            objPopulationTargets = value
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
        RowHeadersWidth = 200
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

        Me.Targets = currentindicator.Targets
        Me.PopulationTargets = currentindicator.PopulationTargets

        LoadColumns()
    End Sub

    Private Sub LoadColumns()
        Columns.Clear()

        With colBaseline
            .Name = "Baseline"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
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
            Dim colTarget As New DataGridViewTextBoxColumn
            With colTarget
                .Name = selTargetDeadline.Deadline
                .MinimumWidth = 100
                .FillWeight = 20
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

        If Me.Columns.Count > 0 Then Columns(0).HeaderText = strBaseline

        'header text of target columns
        For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
            intIndex = TargetDeadlinesSection.TargetDeadlines.IndexOf(selTargetDeadline)
            strDate = TargetDeadlinesSection.FormatTargetDeadlineDate(selTargetDeadline)

            If CurrentIndicator.TargetSystem = Indicator.TargetSystems.Formula Then

                Dim selTarget As Target = Me.Targets(intIndex)
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

            intColumnIndex = intIndex + 1
            If intColumnIndex <= Me.Columns.Count - 1 Then
                Columns(intColumnIndex).HeaderText = strHeaderText
            End If
        Next
    End Sub

    Private Sub SetRowHeadersText()
        Me.RowCount = 4

        If CurrentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            Rows(0).HeaderCell.Value = LANG_TargetValueBeneficiary
            Rows(1).HeaderCell.Value = LANG_ScoreValueBeneficiary
        Else
            Rows(0).HeaderCell.Value = LANG_Target
            Rows(1).HeaderCell.Value = LANG_Score
        End If
        Rows(2).HeaderCell.Value = LANG_PopulationTargetText

        If Me.TargetGroup IsNot Nothing Then
            Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

            Rows(3).HeaderCell.Value = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
        Else
            Rows(3).HeaderCell.Value = LANG_ScoreValueTargetGroup
        End If

        If CurrentIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
            Rows(1).Visible = False
        Else
            Me(0, 1).ReadOnly = True
        End If

        If CurrentIndicator.Registration <> Indicator.RegistrationOptions.BeneficiaryLevel Then
            Rows(2).Visible = False
            Rows(3).Visible = False
        Else
            For i = 0 To Columns.Count - 1
                Me(i, 3).ReadOnly = True
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

        'set current cell location to what it was before
        Me.Invalidate()
        Me.ResumeLayout()
        CurrentCell = Me(objCellLocation.X, objCellLocation.Y)
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

        Dim intRowIndex As Integer = e.RowIndex
        Dim intTargetIndex As Integer = e.ColumnIndex - 1


        If intRowIndex = 0 And e.ColumnIndex > 0 Then
            If CurrentIndicator.TargetSystem <> Indicator.TargetSystems.Simple Then
                ShowDialog(intRowIndex, intTargetIndex)
            End If
        End If
    End Sub

    Protected Overrides Sub OnCellDoubleClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellDoubleClick(e)

        Dim intTargetIndex As Integer = e.ColumnIndex
        If intTargetIndex > 0 Then
            intTargetIndex -= 1
            ShowDialog(e.RowIndex, intTargetIndex)
        End If
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As Object, ByVal boolShowScore As Boolean, ByVal boolShowPopulationTarget As Boolean)
        Dim intColIndex, intRowIndex As Integer

        If boolShowScore = True Then intRowIndex = 1
        If boolShowPopulationTarget = True Then intRowIndex = 2

        Select Case selItem.GetType
            Case GetType(Baseline)
                intColIndex = 0
            Case GetType(Target)
                Dim selTarget As Target = TryCast(selItem, Target)

                If selTarget IsNot Nothing Then
                    intColIndex = Me.Targets.IndexOf(selTarget)
                    intColIndex += 1
                End If
            Case GetType(PopulationTarget)
                Dim selPopulationTarget As PopulationTarget = TryCast(selItem, PopulationTarget)

                If selPopulationTarget IsNot Nothing Then
                    intColIndex = Me.PopulationTargets.IndexOf(selPopulationTarget)
                    intColIndex += 1
                End If
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Private Sub MoveCurrentCell()
        If CurrentCell.RowIndex = 1 And CurrentCell.ColumnIndex = 0 Then
            CurrentCell = Me(1, 1)
            CurrentCell.Selected = True
        ElseIf CurrentCell.RowIndex = 2 And CurrentCell.ColumnIndex = 0 Then
            CurrentCell = Me(1, 2)
            CurrentCell.Selected = True
        ElseIf CurrentCell.RowIndex = 3 Then
            CurrentCell = Me(CurrentCell.ColumnIndex, 2)
            CurrentCell.Selected = True
        End If
    End Sub

    Private Sub ShowDialog(ByVal intRowIndex As Integer, ByVal intIndex As Integer)
        If intRowIndex = 0 Or intRowIndex = 1 Then
            ShowDialog_Targets(intIndex)
        End If

        Reload()
    End Sub

    Public Sub ShowDialog_Targets(ByVal intTargetIndex As Integer)
        Dim selTargetDeadline As TargetDeadline = TargetDeadlinesSection.TargetDeadlines(intTargetIndex)
        Dim selTarget As Target = Me.Targets(intTargetIndex)

        Select Case CurrentIndicator.TargetSystem
            Case Indicator.TargetSystems.Simple, Indicator.TargetSystems.ValueRange
                Dim DialogTarget As New DialogTarget(selTargetDeadline.ExactDeadline, selTarget, CurrentIndicator.ScoringSystem, CurrentIndicator.TargetSystem, _
                                                     CurrentIndicator.ValuesDetail.Unit, CurrentIndicator.ValuesDetail.NrDecimals)

                If DialogTarget.ShowDialog() = DialogResult.OK Then
                    Select Case CurrentIndicator.ScoringSystem
                        Case Indicator.ScoringSystems.Percentage
                            selTarget.Score = 1
                    End Select
                End If

            Case Indicator.TargetSystems.Formula
                Dim DialogTargetFormula As New DialogTargetFormula(selTargetDeadline.ExactDeadline, selTarget, CurrentIndicator.ScoringSystem)
                If DialogTargetFormula.ShowDialog() = DialogResult.OK Then
                    CurrentIndicator.CalculateTargetsWithFormula()
                    Me.Invalidate()
                End If
        End Select
    End Sub

    Public Sub Edit()
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intTargetIndex As Integer = intColumnIndex - 1

        If intRowIndex = 0 Then
            If intColumnIndex = 0 Or (intColumnIndex > 0 And CurrentIndicator.TargetSystem = Indicator.TargetSystems.Simple) Then
                Me.BeginEdit(True)
            Else
                ShowDialog(intRowIndex, intTargetIndex)
            End If
        ElseIf intRowIndex = 1 And intColumnIndex > 0 Then
            Me.BeginEdit(True)
        ElseIf intRowIndex = 2 Then
            ShowDialog(intRowIndex, intTargetIndex)
        End If
    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)

        Dim intRowIndex As Integer = e.RowIndex

        If intRowIndex = 0 Then
            e.Value = OnCellValueNeeded_TargetValues(e.ColumnIndex)
        ElseIf intRowIndex = 1 Then
            e.Value = OnCellValueNeeded_TargetScores(e.ColumnIndex)
        ElseIf intRowIndex = 2 Then
            e.Value = OnCellValueNeeded_PopulationTargets(e.ColumnIndex)
        ElseIf intRowIndex = 3 Then
            e.Value = OnCellValueNeeded_TotalPopulationScores(e.ColumnIndex)
        End If
    End Sub

    Private Function OnCellValueNeeded_TargetValues(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex = 0 Then
            strValue = CurrentIndicator.GetBaselineFormattedValue
        Else
            Dim intTargetIndex As Integer = intColumnIndex - 1
            strValue = CurrentIndicator.GetTargetFormattedValue(intTargetIndex)
        End If

        Return strValue
    End Function

    Private Function OnCellValueNeeded_TargetScores(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex > 0 Then
            Dim intTargetIndex As Integer = intColumnIndex - 1
            strValue = CurrentIndicator.GetTargetFormattedScore(intTargetIndex)
        End If

        Return strValue
    End Function

    Private Function OnCellValueNeeded_PopulationTargets(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex > 0 Then
            Dim intTargetIndex As Integer = intColumnIndex - 1
            Dim selTargetPopulation As PopulationTarget = Me.PopulationTargets(intTargetIndex)

            If selTargetPopulation IsNot Nothing Then
                strValue = DisplayAsUnit(selTargetPopulation.Percentage, CurrentIndicator.ValuesDetail.NrDecimals, "%")
            End If
        End If

        Return strValue
    End Function

    Private Function OnCellValueNeeded_TotalPopulationScores(ByVal intColumnIndex As Integer) As String
        Dim strValue As String = String.Empty

        If intColumnIndex = 0 Then
            strValue = CurrentIndicator.GetPopulationBaselineFormattedScore
        Else
            Dim intTargetIndex As Integer = intColumnIndex - 1

            strValue = CurrentIndicator.GetPopulationTargetFormattedScore(intTargetIndex)
        End If

        Return strValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim intRowIndex As Integer = e.RowIndex

        If intRowIndex = 0 Then
            OnCellValuePushed_TargetValues(e.ColumnIndex, e.Value)
        ElseIf intRowIndex = 1 Then
            OnCellValuePushed_TargetScores(e.ColumnIndex, e.Value)
        ElseIf intRowIndex = 2 Then
            OnCellValuePushed_PopulationTargets(e.ColumnIndex, e.Value)
        End If
        Invalidate()

    End Sub

    Private Sub OnCellValuePushed_TargetValues(ByVal intColumnIndex As Integer, ByVal objValue As Object)
        Dim dblValue As Double = ParseDouble(objValue)

        dblValue = CurrentIndicator.ValuesDetail.ValueRange.MakeWithinRange(dblValue)

        If intColumnIndex = 0 Then
            UndoRedo.UndoBuffer_Initialise(CurrentIndicator.Baseline, "Value", CurrentIndicator.Baseline.Value)
            CurrentIndicator.Baseline.Value = dblValue

            UndoRedo.ValueChanged(dblValue)
            If CurrentIndicator.TargetSystem = Indicator.TargetSystems.Formula Then
                CurrentIndicator.CalculateTargetsWithFormula()
                Me.Invalidate()
            End If
        Else
            Dim intTargetIndex As Integer = intColumnIndex - 1
            Dim selTarget As Target = Me.Targets(intTargetIndex)

            Select Case CurrentIndicator.TargetSystem
                Case Indicator.TargetSystems.Simple, Indicator.TargetSystems.ValueRange
                    UndoRedo.UndoBuffer_Initialise(selTarget, "MinValue", selTarget.MinValue)
                    selTarget.OpMin = CONST_LargerThanOrEqual
                    selTarget.MinValue = dblValue

                    UndoRedo.ValueChanged(dblValue)
                Case Indicator.TargetSystems.Formula
                    'use dialog to set values
            End Select
        End If
    End Sub

    Private Sub OnCellValuePushed_TargetScores(ByVal intColumnIndex As Integer, ByVal objValue As Object)
        If intColumnIndex > 0 Then
            Dim intTargetIndex As Integer = intColumnIndex - 1
            Dim selTarget As Target = Me.Targets(intTargetIndex)

            UndoRedo.UndoBuffer_Initialise(selTarget, "Score", selTarget.Score)
            selTarget.Score = ParseDouble(objValue)
            UndoRedo.ValueChanged(selTarget.Score)
        End If
    End Sub

    Private Sub OnCellValuePushed_PopulationTargets(ByVal intColumnIndex As Integer, ByVal objValue As Object)
        If intColumnIndex > 0 Then
            Dim intTargetIndex As Integer = intColumnIndex - 1
            Dim selPopulationTarget As PopulationTarget = Me.PopulationTargets(intTargetIndex)

            UndoRedo.UndoBuffer_Initialise(selPopulationTarget, "Percentage", selPopulationTarget.Percentage)
            selPopulationTarget.Percentage = ParseDouble(objValue)
            UndoRedo.ValueChanged(selPopulationTarget.Percentage)
        End If
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim CellGraphics As Graphics = e.Graphics
        Dim rCell As Rectangle = e.CellBounds
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If (intRowIndex = 1 Or intRowIndex = 2) And intColumnIndex = 0 Then
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
        ElseIf intRowIndex = 3 Then
            CellGraphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            e.PaintContent(rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
        End If
    End Sub
#End Region
End Class
