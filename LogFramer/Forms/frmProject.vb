Imports System.Text.RegularExpressions

Public Class frmProject
    Friend WithEvents UndoList As New classUndoList
    Friend WithEvents RedoList As New classUndoList
    Friend WithEvents NewDetailLogframe As New UserControl
    Friend WithEvents CurrentDetailLogframe As UserControl
    Friend WithEvents NewDetailPlanning As New UserControl
    Friend WithEvents CurrentDetailPlanning As UserControl
    Friend WithEvents NewDetailBudget As New DetailBudgetItem
    Friend WithEvents CurrentDetailBudget As DetailBudgetItem
    Private DetailExchangeRates As DetailExchangeRates

    Private lfLogFrame As New LogFrame
    Private strDocName, strDocPath As String

    Private objCurrentDataGridView As DataGridView = dgvLogframe
    Private intTextSelectionIndex As Integer

#Region "Enumerations"
    Public Enum TextSelectionValues
        SelectNothing = 0
        SelectAll = 1
        SelectLogframe = 2
        SelectStructs = 3
        SelectGoals = 4
        SelectPurposes = 5
        SelectOutputs = 6
        SelectActivities = 7
        SelectIndicators = 4
        SelectVerificationSources = 5
        SelectResources = 6
        SelectResourceBudgets = 7
        SelectBudgetItems = 8
        SelectAssumptions = 9
        SelectStatements = 10
    End Enum
#End Region

#Region "Properties"
    Public Property ProjectLogframe As LogFrame
        Get
            Return lfLogFrame
        End Get
        Set(ByVal value As LogFrame)
            lfLogFrame = value
        End Set
    End Property

    Public Property DocName As String
        Get
            Return strDocName
        End Get
        Set(ByVal value As String)
            strDocName = value
        End Set
    End Property

    Public Property DocPath As String
        Get
            Return strDocPath
        End Get
        Set(ByVal value As String)
            strDocPath = value
        End Set
    End Property

    Public Property CurrentDataGridView() As DataGridViewBaseClassRichText
        Get
            Return objCurrentDataGridView
        End Get
        Set(ByVal value As DataGridViewBaseClassRichText)
            objCurrentDataGridView = value
        End Set
    End Property

    Public ReadOnly Property CurrentDataGridViewBudgetYear As DataGridViewBudgetYear
        Get
            Dim dgvBudgetYear As DataGridViewBudgetYear = Nothing

            If TabControlBudget.TabPages.Count > 0 AndAlso TabControlBudget.SelectedTab IsNot Nothing Then
                Dim SplitContainerBudgetYear As SplitContainer = TabControlBudget.SelectedTab.Controls("SplitContainerBudgetYear")

                If SplitContainerBudgetYear IsNot Nothing Then
                    dgvBudgetYear = SplitContainerBudgetYear.Panel1.Controls("dgvBudgetYear")
                End If
            End If

            Return dgvBudgetYear
        End Get
    End Property

    Public Property TextSelectionIndex() As Integer
        Get
            Return intTextSelectionIndex
        End Get
        Set(ByVal value As Integer)
            intTextSelectionIndex = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        MyBase.OnLoad(e)

        Me.ControlBox = False
        Me.WindowState = FormWindowState.Maximized
        Me.BringToFront()
    End Sub

    Public Function Initialize_PanelsLogframe() As String
        'temporary for Logframer 2.x
        'remove for later versions

        Me.SplitContainerMain.Panel2Collapsed = True
        With TabControlProject.TabPages
            .Remove(TabPageMonitoring)
            .Remove(TabPageActions)
        End With
        With Me.ProjectInfo
            .TabControlProjectInfo.TabPages.Remove(.TabPageDonors)
        End With

        'details window logframe
        With Me.SplitContainerLogFrame
            .Panel1MinSize = 200
            .Panel2MinSize = 242

            If My.Settings.setShowDetailsLogframe = True And NewDetailLogframe IsNot Nothing Then
                .SplitterDistance = .Height - NewDetailLogframe.Height
                .Panel2Collapsed = False

                Return LANG_HideDetails
            Else
                .Panel2Collapsed = True
                My.Settings.setShowDetailsLogframe = False
                Return LANG_ShowDetails
            End If
        End With
    End Function

    Public Function Initialize_PanelsPlanning() As String
        With Me.SplitContainerPlanning
            .Panel1MinSize = 200
            .Panel2MinSize = 242

            If My.Settings.setShowDetailsPlanning = True And NewDetailPlanning IsNot Nothing Then
                .SplitterDistance = .Height - NewDetailPlanning.Height
                .Panel2Collapsed = False

                Return LANG_HideDetails
            Else
                .Panel2Collapsed = True
                My.Settings.setShowDetailsPlanning = False
                Return LANG_ShowDetails
            End If
        End With
    End Function

    Public Function Initialize_PanelsBudget() As String
        Dim strShowDetails As String = String.Empty

        With Me.SplitContainerExchangeRates
            .Panel1MinSize = 300
            .Panel2MinSize = 300

            If My.Settings.setShowExchangeRates = True And DetailExchangeRates IsNot Nothing Then
                .SplitterDistance = .Width - DetailExchangeRates.Width
                .Panel2Collapsed = False
            Else
                .Panel2Collapsed = True
            End If
        End With

        For i = 0 To TabControlBudget.TabPages.Count - 1
            Dim SplitContainerBudget As SplitContainer = GetSplitContainerBudget(i)

            With SplitContainerBudget
                .Panel1MinSize = 200
                .Panel2MinSize = 242

                If My.Settings.setShowDetailsBudget = True And NewDetailBudget IsNot Nothing Then
                    .SplitterDistance = .Height - NewDetailBudget.Height
                    .Panel2Collapsed = False

                    strShowDetails = LANG_HideDetails
                Else
                    .Panel2Collapsed = True
                    strShowDetails = LANG_ShowDetails
                End If
            End With
        Next

        Return strShowDetails
    End Function

    Public Sub ClearUndo()
        Me.UndoList.Clear()
        Me.RedoList.Clear()
    End Sub

    Public Function IsDefaultProjectName() As Boolean
        Dim strPattern As String = LANG_Project & "\s?[0-9]+" 'translation of "Project" + optional space + one or more numbers
        If Regex.IsMatch(Me.Name, strPattern) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub AddNewProject()
        With ProjectLogframe
            .StartDate = New Date(Now.Year, Now.Month, Now.Day)
            .EndDate = .StartDate
            .Budget.BudgetYears.Add(New BudgetYear(.StartDate))
        End With

        ProjectInitialise()

        dgvLogframe.Reload()
        dgvPlanning.Reload()
        If CurrentDataGridViewBudgetYear IsNot Nothing Then
            CurrentDataGridViewBudgetYear.Reload()
        End If
    End Sub

    Public Sub ProjectInitialise()
        ProjectInfo.LoadItems()

        ProjectInitialise_dgvLogframe()
        ProjectInitialise_dgvPlanning()
        ProjectInitialise_Budget()

        With TabControlQuestionnaires
            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                    Dim NewTabPage As TabPage
                    Dim dgvQuestionnaire As New DataGridViewQuestionnaire

                    With dgvQuestionnaire
                        .Dock = DockStyle.Fill
                        .ShowRegistrationOption = Indicator.RegistrationOptions.BeneficiaryLevel
                        .ShowTargetGroup = selTargetGroup.Guid
                        .InitialiseColumns()
                    End With

                    .TabPages.Add(selTargetGroup.Guid.ToString, selTargetGroup.Name)
                    NewTabPage = .TabPages(.TabPages.Count - 1)
                    NewTabPage.Controls.Add(dgvQuestionnaire)

                Next
            Next
        End With
    End Sub
#End Region

#Region "Tab controls"
    Private Sub TabControlProject_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlProject.SelectedIndexChanged
        If TabControlProject.SelectedIndex <> CurrentMainTab Then
            Select Case TabControlProject.SelectedTab.Name
                Case TabPageLogframe.Name
                    dgvLogframe.Reload()
                Case TabPagePlanning.Name
                    dgvPlanning.Reload()
                Case TabPageBudget.Name
                    If CurrentDataGridViewBudgetYear IsNot Nothing Then
                        CurrentDataGridViewBudgetYear.Reload()
                    End If
                    If My.Settings.setShowExchangeRates = True And (Me.DetailExchangeRates Is Nothing OrElse DetailExchangeRates.Budget Is Nothing) Then
                        DetailExchangeRates = New DetailExchangeRates(ProjectLogframe.Budget)
                        DetailExchangeRates.Dock = DockStyle.Fill

                        SplitContainerExchangeRates.Panel2.Controls.Add(DetailExchangeRates)
                    End If
            End Select
        End If
        CurrentMainTab = TabControlProject.SelectedIndex

        frmParent.SetRibbonLayOutTabsVisibility()
    End Sub

    Private Sub TabControlQuestionnaires_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlQuestionnaires.SelectedIndexChanged
        If TabControlQuestionnaires.SelectedIndex <> CurrentQuestionnaireTab Then
            If TabControlQuestionnaires.TabPages(TabControlQuestionnaires.SelectedIndex).Controls.Count > 0 Then
                Dim ctlDatagridview As Control = TabControlQuestionnaires.TabPages(TabControlQuestionnaires.SelectedIndex).Controls(0)
                Dim dgvQuestionnaire As DataGridViewQuestionnaire = TryCast(ctlDatagridview, DataGridViewQuestionnaire)

                If dgvQuestionnaire IsNot Nothing Then
                    dgvQuestionnaire.Reload()
                End If
            End If
        End If
        CurrentQuestionnaireTab = TabControlQuestionnaires.SelectedIndex
    End Sub
#End Region

#Region "Undo & Redo"
    Private Sub UndoList_UndoItemAdded() Handles UndoList.UndoItemAdded
        frmParent.ReloadSplitUndoRedoButtons()
    End Sub

    Private Sub UndoList_UndoItemRemoved() Handles UndoList.UndoItemRemoved
        frmParent.ReloadSplitUndoRedoButtons()
    End Sub

    Private Sub RedoList_UndoItemAdded() Handles RedoList.UndoItemAdded
        frmParent.ReloadSplitUndoRedoButtons()
    End Sub

    Private Sub RedoList_UndoItemRemoved() Handles RedoList.UndoItemRemoved
        frmParent.ReloadSplitUndoRedoButtons()
    End Sub
#End Region
End Class