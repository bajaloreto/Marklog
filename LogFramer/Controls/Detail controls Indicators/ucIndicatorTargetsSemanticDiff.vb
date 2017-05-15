Public Class IndicatorTargetsSemanticDiff
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
            Dim dgvSemanticDiff As DataGridViewTargetsSemanticDiff = selTabPage.Controls("dgvTargetsSemanticDiff")

            If dgvSemanticDiff IsNot Nothing Then
                Return dgvSemanticDiff
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Overridable Sub LoadScales()
        Dim intTargetIndex As Integer
        Dim selTarget As Target

        CurrentIndicator.CalculateScores()

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                Dim BaselineSemanticDiff As New DataGridViewTargetsSemanticDiff(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, -1)
                With BaselineSemanticDiff
                    .Dock = DockStyle.Fill
                    .Name = "dgvTargetsSemanticDiff"

                    AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                    AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                    AddHandler .Enter, AddressOf DataGridView_Enter
                End With

                TabPageBaseline.Controls.Add(BaselineSemanticDiff)

                For intTargetIndex = 0 To .Targets.Count - 1
                    selTarget = .Targets(intTargetIndex)

                    Dim NewTabPage As New TabPage()
                    Dim TargetSemanticDiff As New DataGridViewTargetsSemanticDiff(Me.CurrentIndicator, TargetDeadlinesSection.TargetDeadlines, Me.TargetGroup, intTargetIndex)
                    With TargetSemanticDiff
                        .Dock = DockStyle.Fill
                        .Name = "dgvTargetsSemanticDiff"

                        AddHandler .ScoreUpdated, AddressOf SetTabHeaders
                        AddHandler .EditingControlShowing, AddressOf DataGridView_EditingControlShowing
                        AddHandler .Enter, AddressOf DataGridView_Enter
                    End With

                    NewTabPage.Controls.Add(TargetSemanticDiff)
                    TabControlTargets.TabPages.Add(NewTabPage)
                Next
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

    Public Sub SetFocusOnItem(ByVal selItem As Object, Optional ByVal strPropertyName As String = "", Optional ByVal intIndex As Integer = 0, Optional ByVal intLength As Integer = 0)
        Dim intRowIndex, intColIndex As Integer
        Dim intTabPage As Integer
        Dim dgvTargets As DataGridViewTargetsSemanticDiff

        Select Case selItem.GetType
            Case GetType(Statement)
                Dim selStatement As Statement = DirectCast(selItem, Statement)
                intTabPage = 0
                intRowIndex = CurrentIndicator.Statements.IndexOf(selStatement)

                TabControlTargets.SelectedIndex = intTabPage
                dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsSemanticDiff")

                If dgvTargets IsNot Nothing Then
                    With dgvTargets
                        If strPropertyName = "SecondLabel" Then intColIndex = .Columns.Count - 1
                        .CurrentCell = dgvTargets(intColIndex, intRowIndex)
                        .BeginEdit(False)

                        Dim ctlEdit As RichTextEditingControlLogframe = CType(.EditingControl, RichTextEditingControlLogframe)
                        ctlEdit.Select(intIndex, intLength)
                    End With
                End If

            Case GetType(Baseline)
                intTabPage = 0
                intColIndex = 1
                intRowIndex = CurrentIndicator.Statements.Count + 1

                TabControlTargets.SelectedIndex = intTabPage
                dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsSemanticDiff")

                If dgvTargets IsNot Nothing Then
                    dgvTargets.CurrentCell = dgvTargets(intColIndex, intRowIndex)
                End If
            Case GetType(PopulationTarget)
                Dim selPopulationTarget As PopulationTarget = DirectCast(selItem, PopulationTarget)

                intTabPage = CurrentIndicator.PopulationTargets.IndexOf(selPopulationTarget) + 1
                intColIndex = 1
                intRowIndex = CurrentIndicator.Statements.Count + 1

                TabControlTargets.SelectedIndex = intTabPage
                dgvTargets = TabControlTargets.SelectedTab.Controls("dgvTargetsSemanticDiff")

                If dgvTargets IsNot Nothing Then
                    dgvTargets.CurrentCell = dgvTargets(intColIndex, intRowIndex)
                End If
        End Select
    End Sub

    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        If intTextSelectionIndex > 0 Then
            intTextSelectionIndex = frmProject.TextSelectionValues.SelectAll
        End If

        Dim selTabPage As TabPage = TabControlTargets.SelectedTab

        For i = 0 To TabControlTargets.TabCount - 1
            Dim dgvSemanticDiff As DataGridViewTargetsSemanticDiff = TabControlTargets.TabPages(i).Controls("dgvTargetsSemanticDiff")

            If dgvSemanticDiff IsNot Nothing Then
                With dgvSemanticDiff
                    .TextSelectionIndex = intTextSelectionIndex

                    .ResetAllImages()
                    .ReloadImages()
                    .Invalidate()
                End With
            End If
        Next
            
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
        Dim dgvSemanticDiff As DataGridViewTargetsSemanticDiff = selTabPage.Controls("dgvTargetsSemanticDiff")

        If dgvSemanticDiff IsNot Nothing Then
            dgvSemanticDiff.Reload()
        End If
    End Sub
#End Region
End Class
