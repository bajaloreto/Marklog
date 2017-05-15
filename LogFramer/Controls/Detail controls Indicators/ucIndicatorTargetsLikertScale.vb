Public Class IndicatorTargetsLikertScale
    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

    Public Event CurrentDataGridViewChanged(ByVal sender As Object, ByVal e As CurrentDataGridViewChangedEventArgs)

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

    Public ReadOnly Property CurrentTargetDatagridview
        Get
            Dim selTabPage As TabPage = TabControlTargets.SelectedTab

            Select Case CurrentIndicator.QuestionType
                Case Indicator.QuestionTypes.LikertScale
                    Dim dgvLikert As DataGridViewTargetsScaleLikert = selTabPage.Controls("dgvTargetsScaleLikert")

                    If dgvLikert IsNot Nothing Then
                        Return dgvLikert
                    Else
                        Return Nothing
                    End If
                Case Indicator.QuestionTypes.FrequencyLikert
                    Dim dgvFrequencyLikert As DataGridViewTargetsFrequencyLikert = selTabPage.Controls("dgvTargetsFrequencyLikert")

                    If dgvFrequencyLikert IsNot Nothing Then
                        Return dgvFrequencyLikert
                    Else
                        Return Nothing
                    End If
            End Select

            Return Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Overridable Sub LoadScales()
        Dim intTargetIndex As Integer
        Dim selTarget As Target

        CurrentIndicator.CalculateScores()

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                Select Case .QuestionType
                    Case Indicator.QuestionTypes.LikertScale
                        Dim BaselineLikertScale As New DataGridViewTargetsScaleLikert(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, -1)
                        With BaselineLikertScale
                            .Name = "dgvTargetsScaleLikert"
                            .Dock = DockStyle.Fill

                            AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                            AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                            AddHandler .Enter, AddressOf DataGridView_Enter
                        End With
                        
                        TabPageBaseline.Controls.Add(BaselineLikertScale)

                        For intTargetIndex = 0 To .Targets.Count - 1
                            selTarget = .Targets(intTargetIndex)

                            Dim NewTabPage As New TabPage()
                            Dim NewLikertScale As New DataGridViewTargetsScaleLikert(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, intTargetIndex)
                            With NewLikertScale
                                .Name = "dgvTargetsScaleLikert"
                                .Dock = DockStyle.Fill

                                AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                                AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                                AddHandler .Enter, AddressOf DataGridView_Enter
                            End With

                            NewTabPage.Controls.Add(NewLikertScale)
                            TabControlTargets.TabPages.Add(NewTabPage)
                        Next
                    Case Indicator.QuestionTypes.FrequencyLikert
                        Dim BaselineFrequencyLikert As New DataGridViewTargetsFrequencyLikert(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, -1)
                        With BaselineFrequencyLikert
                            .Name = "dgvTargetsFrequencyLikert"
                            .Dock = DockStyle.Fill

                            AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                            AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                            AddHandler .Enter, AddressOf DataGridView_Enter
                        End With

                        TabPageBaseline.Controls.Add(BaselineFrequencyLikert)

                        For intTargetIndex = 0 To .Targets.Count - 1
                            selTarget = .Targets(intTargetIndex)

                            Dim NewTabPage As New TabPage()
                            Dim TargetFrequencyLikert As New DataGridViewTargetsFrequencyLikert(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, intTargetIndex)
                            With TargetFrequencyLikert
                                .Name = "dgvTargetsFrequencyLikert"
                                .Dock = DockStyle.Fill

                                AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                                AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                                AddHandler .Enter, AddressOf DataGridView_Enter
                            End With

                            NewTabPage.Controls.Add(TargetFrequencyLikert)
                            TabControlTargets.TabPages.Add(NewTabPage)
                        Next
                    Case Indicator.QuestionTypes.MaxDiff
                        Dim BaselineMaxDiff As New DataGridViewTargetsMaxDiff(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, -1)
                        With BaselineMaxDiff
                            .Name = "dgvTargetsMaxDiff"
                            .Dock = DockStyle.Fill

                            AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                            AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                            AddHandler .Enter, AddressOf DataGridView_Enter
                        End With
                        TabPageBaseline.Controls.Add(BaselineMaxDiff)

                        For intTargetIndex = 0 To .Targets.Count - 1
                            selTarget = .Targets(intTargetIndex)

                            Dim NewTabPage As New TabPage()
                            Dim TargetMaxDiff As New DataGridViewTargetsMaxDiff(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, intTargetIndex)

                            With TargetMaxDiff
                                .Name = "dgvTargetsMaxDiff"
                                .Dock = DockStyle.Fill

                                AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                                AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                                AddHandler .Enter, AddressOf DataGridView_Enter
                            End With
                            NewTabPage.Controls.Add(TargetMaxDiff)
                            TabControlTargets.TabPages.Add(NewTabPage)
                        Next

                End Select

            End With
            SetTabHeaders()
        End If
    End Sub

    Public Sub SetTabHeaders()
        Dim strBaselineScore As String = 0.ToString
        Dim strBaseline, strHeaderText As String
        Dim strDate As String
        Dim strScore As String = String.Empty
        Dim selTargetDeadline As TargetDeadline

        CurrentIndicator.CalculateScores()

        'header text of baseline column
        strBaselineScore = CurrentIndicator.GetBaselineFormattedScore

        'set header text
        strBaseline = String.Format("{0} - {1}: {2}", LANG_Baseline, LANG_ScoringValue, strBaselineScore)

        TabPageBaseline.Text = strBaseline

        For intTargetIndex = 0 To CurrentIndicator.Targets.Count - 1
            selTargetDeadline = TargetDeadlinesSection.TargetDeadlines(intTargetIndex)

            'deadline
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
            strHeaderText = String.Format("{0} {1} - {2}", LANG_Target, strDate, strScore)

            TabControlTargets.TabPages(intTargetIndex + 1).Text = strHeaderText
        Next
    End Sub

    Public Sub SetFocusOnItem(ByVal selItem As Object, Optional ByVal intIndex As Integer = 0, Optional ByVal intLength As Integer = 0)
        Dim intRowIndex, intColIndex As Integer
        Dim intTabPage As Integer

        Select Case selItem.GetType
            Case GetType(Statement)
                Dim selStatement As Statement = DirectCast(selItem, Statement)
                intTabPage = 0
                intRowIndex = CurrentIndicator.Statements.IndexOf(selStatement)

                TabControlTargets.SelectedIndex = intTabPage
                Select Case CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.LikertScale
                        Dim dgvTargets As DataGridViewTargetsScaleLikert

                        dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsScaleLikert")

                        If dgvTargets IsNot Nothing Then
                            With dgvTargets
                                .SetFocusOnItem(selItem)
                                .CurrentCell = dgvTargets(intColIndex, intRowIndex)
                                .BeginEdit(False)

                                Dim ctlEdit As RichTextEditingControlLogframe = CType(.EditingControl, RichTextEditingControlLogframe)
                                ctlEdit.Select(intIndex, intLength)
                            End With
                        End If
                    Case Indicator.QuestionTypes.FrequencyLikert
                        Dim dgvTargets As DataGridViewTargetsFrequencyLikert

                        dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsFrequencyLikert")

                        If dgvTargets IsNot Nothing Then
                            With dgvTargets
                                .SetFocusOnItem(selItem)
                                .CurrentCell = dgvTargets(intColIndex, intRowIndex)
                                .BeginEdit(False)

                                Dim ctlEdit As RichTextEditingControlLogframe = CType(.EditingControl, RichTextEditingControlLogframe)
                                ctlEdit.Select(intIndex, intLength)
                            End With
                        End If
                End Select

            Case GetType(DoubleValue)
                Dim dgvTargets As DataGridViewBaseClass
                Dim selValue As DoubleValue = DirectCast(selItem, DoubleValue)
                Dim selMatrixRow As DoubleValuesMatrixRow
                Dim intTargetIndex As Integer

                selMatrixRow = CurrentIndicator.Baseline.DoubleValuesMatrix.GetParentMatrixRowOfValue(selValue)
                If selMatrixRow IsNot Nothing Then
                    intTabPage = 0
                    intRowIndex = CurrentIndicator.Baseline.DoubleValuesMatrix.IndexOf(selMatrixRow)
                    intColIndex = selMatrixRow.DoubleValues.IndexOf(selValue) + 1
                Else
                    For Each selTarget As Target In CurrentIndicator.Targets
                        selMatrixRow = selTarget.DoubleValuesMatrix.GetParentMatrixRowOfValue(selValue)
                        If selMatrixRow IsNot Nothing Then
                            intTabPage = intTargetIndex + 1
                            intRowIndex = selTarget.DoubleValuesMatrix.IndexOf(selMatrixRow)
                            intColIndex = selMatrixRow.DoubleValues.IndexOf(selValue) + 1

                            Exit For
                        End If
                        intTargetIndex += 1
                    Next
                End If
                TabControlTargets.SelectedIndex = intTabPage
                dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsFrequencyLikert")

                If dgvTargets IsNot Nothing Then
                    dgvTargets.CurrentCell = dgvTargets(intColIndex, intRowIndex)
                End If
            Case GetType(PopulationTarget)
                Dim selPopulationTarget As PopulationTarget = DirectCast(selItem, PopulationTarget)

                Select Case CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.LikertScale
                        Dim dgvTargets As DataGridViewTargetsScaleLikert

                        intTabPage = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget) + 1

                        TabControlTargets.SelectedIndex = intTabPage
                        dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsScaleLikert")

                        If dgvTargets IsNot Nothing Then
                            dgvTargets.SetFocusOnItem(selItem)
                        End If
                    Case Indicator.QuestionTypes.FrequencyLikert
                        Dim dgvTargets As DataGridViewTargetsFrequencyLikert

                        intTabPage = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget) + 1

                        TabControlTargets.SelectedIndex = intTabPage
                        dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsFrequencyLikert")

                        If dgvTargets IsNot Nothing Then
                            dgvTargets.SetFocusOnItem(selItem)
                        End If
                End Select
        End Select
    End Sub

    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        If intTextSelectionIndex > 0 Then
            intTextSelectionIndex = frmProject.TextSelectionValues.SelectAll
        End If

        Dim selTabPage As TabPage = TabControlTargets.SelectedTab

        Select Case CurrentIndicator.QuestionType
            Case Indicator.QuestionTypes.LikertScale
                For i = 0 To TabControlTargets.TabCount - 1
                    Dim dgvLikert As DataGridViewTargetsScaleLikert = TabControlTargets.TabPages(i).Controls("dgvTargetsScaleLikert")

                    If dgvLikert IsNot Nothing Then
                        With dgvLikert
                            .TextSelectionIndex = intTextSelectionIndex

                            .ResetAllImages()
                            .ReloadImages()
                            .Invalidate()
                        End With
                    End If
                Next
            Case Indicator.QuestionTypes.FrequencyLikert
                For i = 0 To TabControlTargets.TabCount - 1
                    Dim dgvFrequencyLikert As DataGridViewTargetsFrequencyLikert = selTabPage.Controls("dgvTargetsFrequencyLikert")

                    If dgvFrequencyLikert IsNot Nothing Then
                        With dgvFrequencyLikert
                            .TextSelectionIndex = intTextSelectionIndex

                            .ResetAllImages()
                            .ReloadImages()
                            .Invalidate()
                        End With
                    End If
                Next
        End Select
    End Sub
#End Region

#Region "Events"
    Private Sub DataGridView_EditingControlShowing(sender As Object, e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
        End With
    End Sub

    Private Sub DataGridView_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
        RaiseEvent CurrentDataGridViewChanged(Me, New CurrentDataGridViewChangedEventArgs(sender))
    End Sub

    Private Sub TabControlTargets_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles TabControlTargets.SelectedIndexChanged
        Dim selTabPage As TabPage = TabControlTargets.SelectedTab

        With CurrentIndicator
            Select Case .QuestionType
                Case Indicator.QuestionTypes.LikertScale
                    Dim dgvLikert As DataGridViewTargetsScaleLikert = selTabPage.Controls("dgvTargetsScaleLikert")

                    If dgvLikert IsNot Nothing Then
                        dgvLikert.Reload()
                    End If
                Case Indicator.QuestionTypes.FrequencyLikert
                    Dim dgvFrequencyLikert As DataGridViewTargetsFrequencyLikert = selTabPage.Controls("dgvTargetsFrequencyLikert")

                    If dgvFrequencyLikert IsNot Nothing Then
                        dgvFrequencyLikert.Reload()
                    End If
                Case Indicator.QuestionTypes.MaxDiff
                    Dim dgvMaxDiff As DataGridViewTargetsMaxDiff = selTabPage.Controls("dgvTargetsMaxDiff")

                    If dgvMaxDiff IsNot Nothing Then
                        dgvMaxDiff.Reload()
                    End If
            End Select
        End With
    End Sub
#End Region
End Class
