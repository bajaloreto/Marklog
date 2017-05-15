Public Class DataGridViewTargetsScaleLikertType
    Inherits DataGridView

    Friend WithEvents EditBox As New NumericTextBox

    Public Event ScoreUpdated()

    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

    Private objCellLocation As New Point
    Private colPopulationTarget, colPopulationScore As New DataGridViewTextBoxColumn

    Private RowModifyIndex As Integer = -1
    Private EditRow As TargetScalesLikertGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private UpdateScore As Boolean

    Private intTargetIndex As Integer = -1
    Private strBaseline As String
    Private strHeaderText() As String
    Private strPopulationTarget, strPopulationScore As String
    Private EditPopulationIndex As Integer

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
    Public Sub New()

    End Sub

    Public Sub New(ByVal currentindicator As Indicator, ByVal targetdeadlinessection As TargetDeadlinesSection, ByVal targetgroup As TargetGroup)
        'datagridview settings
        ChooseSettings()

        'load values
        Me.CurrentIndicator = currentindicator
        Me.TargetDeadlinesSection = targetdeadlinessection
        Me.TargetGroup = targetgroup

        SetBooleanValues()
        LoadColumns()
    End Sub

    Private Sub ChooseSettings()
        'DoubleBuffered = True
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
    End Sub

    Public Sub SetBooleanValues()
        With Me.CurrentIndicator
            Dim intCount As Integer = .ResponseClasses.Count

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

    Public Sub LoadColumns()
        Columns.Clear()

        SetColumnHeadersText()
        LoadTargetColumns()

        If CurrentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
            With colPopulationTarget
                .Name = "PopulationTarget"
                .HeaderText = strHeaderText(IndexPopulationTarget)
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            With colPopulationScore
                .Name = "PopulationScore"
                .HeaderText = strHeaderText(IndexPopulationScore)
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .ReadOnly = True
            End With

            Columns.Add(colPopulationTarget)
            Columns.Add(colPopulationScore)
        End If

        Reload()
    End Sub

    Private Sub LoadTargetColumns()
        Dim intIndex As Integer

        For Each selResponseClass As ResponseClass In CurrentIndicator.ResponseClasses
            Dim colResponseClass As New DataGridViewCheckBoxColumn
            With colResponseClass
                intIndex = CurrentIndicator.ResponseClasses.IndexOf(selResponseClass)
                .Name = String.Format("Class{0}", intIndex)
                .HeaderText = strHeaderText(intIndex)
                .MinimumWidth = 100
                .FillWeight = 1
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End With
            Me.Columns.Add(colResponseClass)
        Next
    End Sub

    Private Sub SetColumnHeadersText()
        Dim selResponseClass As ResponseClass
        Dim intTargetIndex As Integer

        'header text of target columns
        ReDim strHeaderText(CurrentIndicator.ResponseClasses.Count + 1)
        For intTargetIndex = 0 To CurrentIndicator.ResponseClasses.Count - 1
            selResponseClass = CurrentIndicator.ResponseClasses(intTargetIndex)
            strHeaderText(intTargetIndex) = String.Format("{0}{1}{1}{2}", selResponseClass.ClassName, vbCrLf, selResponseClass.Value)
        Next

        strHeaderText(IndexPopulationTarget) = LANG_PopulationTargetText

        If Me.TargetGroup IsNot Nothing Then
            Dim strTypeName As String = TargetGroup.TypeName.ToLower.Substring(0, 5)

            strHeaderText(IndexPopulationScore) = String.Format("{0} ({1} {2}.)", LANG_ScoreValueTargetGroup, TargetGroupNumber, strTypeName)
        Else
            strHeaderText(IndexPopulationScore) = LANG_ScoreValueTargetGroup
        End If
    End Sub

    Private Sub SetRowHeadersText()
        Dim strHeaderText As String = String.Empty
        Dim strDate As String
        Dim strScore As String = String.Empty
        Dim intTargetIndex, intRowIndex As Integer
        Dim strBaselineScore As String = 0.ToString

        CurrentIndicator.CalculateScores()

        'header text of baseline column
        strBaselineScore = CurrentIndicator.GetBaselineFormattedScore

        'set header text
        Dim strBaseline As String = String.Format("{0}{1}{2}: {3}", LANG_Baseline, vbCrLf, LANG_ScoringValue, strBaselineScore)

        If RowCount > 0 Then
            Rows(0).HeaderCell.Value = strBaseline
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

            intRowIndex = intTargetIndex + 1
            If intRowIndex <= RowCount - 1 Then
                Rows(intRowIndex).HeaderCell.Value = strHeaderText
            End If
        Next
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
        Me.RowCount = CurrentIndicator.Targets.Count + 1
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
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selItem As Object)
        Dim intColIndex, intRowIndex As Integer

        Select Case selItem.GetType
            Case GetType(Baseline)
                intRowIndex = 0
                intColIndex = IndexPopulationTarget
            Case GetType(PopulationTarget)
                Dim selPopulationTarget As PopulationTarget = TryCast(selItem, PopulationTarget)

                If selPopulationTarget IsNot Nothing Then
                    intRowIndex = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget) + 1
                    intColIndex = IndexPopulationTarget
                End If
        End Select

        If intColIndex >= 0 And intColIndex < ColumnCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Private Sub MoveCurrentCell()
        If CurrentCell.ColumnIndex = IndexPopulationScore Then
            CurrentCell = Me(IndexPopulationScore - 1, CurrentCell.RowIndex)
            CurrentCell.Selected = True
        Else
            Dim objCurrentCell As DataGridViewCell = CurrentCell
            CurrentCell = Nothing
            CurrentCell = objCurrentCell
        End If
        Invalidate()
    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)

        If e.ColumnIndex < IndexPopulationTarget Then
            e.Value = OnCellValueNeeded_Targets(e.ColumnIndex, e.RowIndex)
        ElseIf e.ColumnIndex = IndexPopulationTarget Then
            e.Value = OnCellValueNeeded_PopulationTargets(e.RowIndex)
        ElseIf e.ColumnIndex = IndexPopulationScore Then
            e.Value = OnCellValueNeeded_PopulationScores(e.RowIndex)
        End If
    End Sub

    Private Function OnCellValueNeeded_Targets(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer) As Boolean
        Dim boolValue As Boolean

        If intRowIndex = 0 Then
            boolValue = CurrentIndicator.Baseline.BooleanValues(intColumnIndex).Value
        Else
            Dim intTargetIndex As Integer = intRowIndex - 1
            Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)

            boolValue = selTarget.BooleanValues(intColumnIndex).Value
        End If

        Return boolValue
    End Function

    Private Function OnCellValueNeeded_PopulationTargets(ByVal intRowIndex As Integer) As String
        Dim strValue As String

        With CurrentIndicator
            If intRowIndex = 0 Then
                strValue = DisplayAsUnit(.Baseline.PopulationPercentage, .ValuesDetail.NrDecimals, "%")
            Else
                Dim intTargetIndex As Integer = intRowIndex - 1
                Dim selPopulationTarget As PopulationTarget = CurrentIndicator.PopulationTargets(intTargetIndex)

                strValue = DisplayAsUnit(selPopulationTarget.Percentage, .ValuesDetail.NrDecimals, "%")
            End If
        End With
        Return strValue
    End Function

    Private Function OnCellValueNeeded_PopulationScores(ByVal intRowIndex As Integer) As String
        Dim strValue As String

        With CurrentIndicator
            If intRowIndex = 0 Then
                strValue = .GetPopulationBaselineFormattedScore()
            Else
                Dim intTargetIndex As Integer = intRowIndex - 1
                strValue = .GetPopulationTargetFormattedScore(intTargetIndex)
            End If
        End With

        Return strValue
    End Function

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex

        If intColumnIndex < IndexPopulationTarget Then
            Dim boolSelected As Boolean
            If Boolean.TryParse(e.Value, boolSelected) Then
                If intRowIndex = 0 Then
                    OnCellValuePushed_Baseline(intColumnIndex, boolSelected)
                Else
                    OnCellValuePushed_Targets(intColumnIndex, intRowIndex, boolSelected)
                End If
            End If
        ElseIf intColumnIndex = IndexPopulationTarget Then
            Dim dblValue As Double
            If Double.TryParse(e.Value, dblValue) Then _
                OnCellValuePushed_PopulationTargets(intRowIndex, dblValue)
        End If
    End Sub

    Private Sub OnCellValuePushed_Baseline(ByVal intColumnIndex As Integer, ByVal boolSelected As Boolean)
        Dim boolSum As Boolean = CurrentIndicator.AddClassValuesForTotal
        Dim objOldValues As BooleanValues

        With CurrentIndicator.Baseline
            Using copier As New ObjectCopy
                objOldValues = copier.CopyCollection(.BooleanValues)
            End Using

            .BooleanValues.SetValue(intColumnIndex, boolSelected, boolSum)
            UndoRedo.BooleanValueChecked(CurrentIndicator.Baseline, intColumnIndex, objOldValues, .BooleanValues)
        End With

        Invalidate()
        SetRowHeadersText()
    End Sub

    Private Sub OnCellValuePushed_Targets(ByVal intColumnIndex As Integer, ByVal intRowIndex As Integer, ByVal boolSelected As Boolean)
        Dim boolSum As Boolean = CurrentIndicator.AddClassValuesForTotal
        Dim intTargetIndex As Integer = intRowIndex - 1
        Dim selTarget As Target = CurrentIndicator.Targets(intTargetIndex)
        Dim objOldValues As BooleanValues

        Using copier As New ObjectCopy
            objOldValues = copier.CopyCollection(selTarget.BooleanValues)
        End Using

        selTarget.BooleanValues.SetValue(intColumnIndex, boolSelected, boolSum)
        UndoRedo.BooleanValueChecked(selTarget, intColumnIndex, objOldValues, selTarget.BooleanValues)

        SetFutureTargets(selTarget, intRowIndex, boolSelected)
        
        Invalidate()
        SetColumnHeadersText()
    End Sub

    Public Sub SetFutureTargets(ByVal selTarget As Target, ByVal intBooleanValueIndex As Integer, ByVal boolSelected As Boolean)
        Dim boolSum As Boolean = CurrentIndicator.AddClassValuesForTotal

        If boolSum = False Then
            If boolSelected = True Then

                'set targets below to the same value
                Dim intIndex As Integer = CurrentIndicator.Targets.IndexOf(selTarget) + 1
                If intIndex <= CurrentIndicator.Targets.Count - 1 Then
                    For i = intIndex To CurrentIndicator.Targets.Count - 1
                        For j = 0 To selTarget.BooleanValues.Count - 1
                            CurrentIndicator.Targets(i).BooleanValues(j).Value = selTarget.BooleanValues(j).Value
                        Next
                        'With CurrentIndicator.Targets(i)
                        '    .BooleanValues.SetValue(intBooleanValueIndex, boolSelected, boolSum)
                        'End With
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub OnCellValuePushed_PopulationTargets(ByVal intRowIndex As Integer, ByVal dblValue As Double)
        If intRowIndex = 0 Then
            CurrentIndicator.Baseline.PopulationPercentage = dblValue
        Else
            Dim intTargetIndex As Integer = intRowIndex - 1
            Dim selPopulationTarget As PopulationTarget = CurrentIndicator.PopulationTargets(intTargetIndex)

            UndoRedo.UndoBuffer_Initialise(selPopulationTarget, "Percentage", selPopulationTarget.Percentage)
            selPopulationTarget.Percentage = dblValue
            UndoRedo.ValueChanged(dblValue)
        End If

        Invalidate()
        SetColumnHeadersText()
    End Sub
#End Region

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim intRowIndex As Integer = e.RowIndex
        Dim intColumnIndex As Integer = e.ColumnIndex
        Dim rCell As Rectangle = e.CellBounds

        If intColumnIndex = IndexPopulationTarget Then
            'If intRowIndex >= 0 Then
            e.Graphics.FillRectangle(brBlue, rCell)
            e.PaintContent(rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
            'End If
        ElseIf intColumnIndex = IndexPopulationScore Then
            'If intRowIndex >= 0 Then
            e.Graphics.FillRectangle(SystemBrushes.ControlDark, rCell)
            e.PaintContent(rCell)
            e.Paint(rCell, DataGridViewPaintParts.Border)

            e.Handled = True
            'End If
        End If
    End Sub
#End Region
End Class
