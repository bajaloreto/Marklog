Imports System.Windows.Forms
Imports System.Reflection

Public Class DialogFind
    Private intCurrentProjectTab, intCurrentBudgetTab As Integer
    Dim intStartProjectTab, intStartBudgetTab As Integer
    Private strFindText As String
    Private strReplaceText As String
    Private intMatchIndex As Integer = -1
    Private intMatchLength As Integer
    Private objMatchItem As Object
    Private objMatchChildItem As Object
    Private strMatchProperty As String

    Private objCurrentControl As Control
    Private objCurrentItem, objStartItem As Object
    Private boolStart As Boolean
    Private lstOpenDialogs As New List(Of Form)

    Private RowIndex As Integer = 1
    Private ColIndex As Integer = 1
    Private StartIndex As Integer, StartIndex2 As Integer
    Private PropInfoIndex As Integer, PropInfoIndex2 As Integer
    Private KeyMomentIndex As Integer
    Private TargetGroupIndex As Integer, TargetGroupInformationIndex As Integer
    Private ResponseIndex As Integer
    Private BudgetItemReferenceIndex As Integer
    Private boolMainChecked As Boolean, boolTargetGroupChecked As Boolean
    Private boolSearchAllPanes As Boolean
    Private boolReplaceAll As Boolean
    Private intItemsReplaced As Integer

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        Me.TopMost = True
        lblMessage.Text = String.Empty
        lblMessageReplace.Text = String.Empty

        Me.tbFind.Focus()
    End Sub

#Region "Properties"
    Public Property FindText As String
        Get
            Return strFindText
        End Get
        Set(ByVal value As String)
            strFindText = value
        End Set
    End Property

    Public Property ReplaceText As String
        Get
            Return strReplaceText
        End Get
        Set(ByVal value As String)
            strReplaceText = value
        End Set
    End Property

    Public Property MatchIndex As Integer
        Get
            Return intMatchIndex
        End Get
        Set(ByVal value As Integer)
            intMatchIndex = value
        End Set
    End Property

    Public Property MatchLength As Integer
        Get
            Return intMatchLength
        End Get
        Set(ByVal value As Integer)
            intMatchLength = value
        End Set
    End Property

    Public Property MatchItem As Object
        Get
            Return objMatchItem
        End Get
        Set(ByVal value As Object)
            objMatchItem = value
        End Set
    End Property

    Public Property MatchChildItem As Object
        Get
            Return objMatchChildItem
        End Get
        Set(ByVal value As Object)
            objMatchChildItem = value
        End Set
    End Property

    Public Property MatchProperty As String
        Get
            Return strMatchProperty
        End Get
        Set(ByVal value As String)
            strMatchProperty = value
        End Set
    End Property

    Private ReadOnly Property MatchCase As Boolean
        Get
            Return Me.chkMatchCase.Checked
        End Get
    End Property

    Private Property SearchAllPanes As Boolean
        Get
            Return boolSearchAllPanes
        End Get
        Set(ByVal value As Boolean)
            boolSearchAllPanes = value
        End Set
    End Property

    Public Property ReplaceAll As Boolean
        Get
            Return boolReplaceAll
        End Get
        Set(ByVal value As Boolean)
            boolReplaceAll = value
        End Set
    End Property

    Public Property ItemsReplaced As Integer
        Get
            Return intItemsReplaced
        End Get
        Set(ByVal value As Integer)
            intItemsReplaced = value
        End Set
    End Property
#End Region

#Region "Events"
    Private Sub btnFindFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindFirst.Click
        Dim strFind As String = Trim(tbFind.Text)

        If String.IsNullOrEmpty(strFind) Then
            lblMessage.Text = LANG_FindIndicateWordsToFind
            Exit Sub
        End If

        SetStartLocation()
        FindReplace(strFind, String.Empty, True)
    End Sub

    Private Sub btnFindNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindNext.Click
        Dim strFind As String = Trim(tbFind.Text)

        If String.IsNullOrEmpty(strFind) Then
            lblMessage.Text = LANG_FindIndicateWordsToFind
            Exit Sub
        End If

        SetStartLocation()
        FindReplace(strFind)
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub chkSearchAllPanes_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSearchAllPanes.CheckedChanged
        Me.SearchAllPanes = chkSearchAllPanes.Checked
    End Sub

    Private Sub btnReplaceNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReplaceNext.Click
        Dim strFind As String = Trim(tbFindReplace.Text)
        Dim strReplace As String = Trim(tbReplace.Text)

        If String.IsNullOrEmpty(strFind) Then
            lblMessageReplace.Text = LANG_FindIndicateWordsToFind
            Exit Sub
        End If
        If String.IsNullOrEmpty(strReplace) Then
            lblMessageReplace.Text = LANG_ReplaceIndicateWordsToReplace
            Exit Sub
        End If

        SetStartLocation()
        FindReplace(strFind, strReplace)

        'With CurrentProjectForm.dgvLogframe
        '    RowIndex = .CurrentCell.RowIndex
        '    ColIndex = .CurrentCell.ColumnIndex
        'End With
        'FindText(strFind, strReplace)
    End Sub

    Private Sub btnReplaceAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReplaceAll.Click
        Dim strFind As String = Trim(tbFindReplace.Text)
        Dim strReplace As String = Trim(tbReplace.Text)

        If String.IsNullOrEmpty(strFind) Then
            lblMessageReplace.Text = LANG_FindIndicateWordsToFind
            Exit Sub
        End If
        If String.IsNullOrEmpty(strReplace) Then
            lblMessageReplace.Text = LANG_ReplaceIndicateWordsToReplace
            Exit Sub
        End If

        SetStartLocation()
        FindReplace(strFind, strReplace, False, True)

        'ReplaceAll(strFind, strReplace)
    End Sub

    Private Sub btnReplaceClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReplaceClose.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub chkReplaceDetails_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkReplaceDetails.CheckedChanged
        If chkReplaceDetails.Checked = True Then
            boolSearchAllPanes = True
            My.Settings.setShowDetailsLogframe = True
            CurrentProjectForm.SetTypeOfDetailWindowLogframe()
        Else
            boolSearchAllPanes = False
        End If
    End Sub

    Private Sub tbFind_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbFind.TextChanged
        tbFindReplace.Text = tbFind.Text
    End Sub

    Private Sub tbFindReplace_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbFindReplace.TextChanged
        tbFind.Text = tbFindReplace.Text
    End Sub
#End Region

#Region "General methods"
    Private Sub SetStartLocation()
        ItemsReplaced = 0

        lblMessage.Text = String.Empty
        lblMessageReplace.Text = String.Empty

        If lstOpenDialogs.Count > 0 Then
            For Each selDialog As Form In lstOpenDialogs
                selDialog.Close()
            Next
            lstOpenDialogs.Clear()
        End If

        With CurrentProjectForm
            intCurrentProjectTab = .TabControlProject.SelectedIndex
            intCurrentBudgetTab = .TabControlBudget.SelectedIndex
            objCurrentControl = CurrentControl
            'objCurrentItem = GetCurrentItem()

            If intCurrentBudgetTab < 0 Then intCurrentBudgetTab = 0

            intStartProjectTab = intCurrentProjectTab
            intStartBudgetTab = intCurrentBudgetTab
            objStartItem = objCurrentItem

            boolStart = True

        End With
    End Sub

    Private Sub FindReplace(ByVal strFind As String, Optional ByVal strReplace As String = "", Optional ByVal boolFromStart As Boolean = False, _
                            Optional ByVal replaceall As Boolean = False)
        Dim boolFound As Boolean
        Dim CurrentObject As Object = GetCurrentObject(boolFromStart)

        Me.FindText = strFind
        Me.ReplaceText = strReplace
        Me.ReplaceAll = replaceall

        With CurrentProjectForm
            If .TabControlProject.SelectedIndex <> intCurrentProjectTab Then
                .TabControlProject.SelectedIndex = intCurrentProjectTab
            End If
            If intCurrentProjectTab = 3 And .TabControlBudget.SelectedIndex <> intCurrentBudgetTab Then
                .TabControlBudget.SelectedIndex = intCurrentBudgetTab
            End If
        End With

        If CurrentObject IsNot Nothing Then
            Using objFindReplace As New FindReplace(CurrentLogFrame)
                InitialiseFindReplace(objFindReplace)

                With objFindReplace
                    Select Case intCurrentProjectTab
                        Case 0 'Project info
                            boolFound = .FindReplaceInProjectInfo(strFind, strReplace, MatchCase, CurrentObject)
                        Case 1 'Logframe
                            boolFound = .FindReplaceInLogframe(strFind, strReplace, MatchCase, CurrentObject)
                        Case 2 'Planning
                            boolFound = .FindReplaceInPlanning(strFind, strReplace, MatchCase, CurrentObject)
                        Case 3 'Budget
                            boolFound = .FindReplaceInBudgetYear(strFind, strReplace, MatchCase, CurrentObject, intCurrentBudgetTab)
                    End Select

                    Me.ItemsReplaced += .ItemsReplaced

                    If boolFound = True Then
                        Me.MatchItem = .MatchItem
                        Me.MatchChildItem = .MatchChildItem
                        Me.MatchIndex = .MatchIndex
                        Me.MatchLength = .MatchLength
                        Me.MatchProperty = .MatchProperty

                        ShowInControl()
                    Else
                        If Me.ReplaceAll = True Then
                            If SearchAllPanes = False Then
                                ReplaceAll_Finalise()
                            Else
                                FindReplace_MoveToNextTab()
                            End If

                        Else
                            If SearchAllPanes = False Then
                                FindReplace_NotFound()
                            Else
                                FindReplace_MoveToNextTab()
                            End If
                        End If
                    End If
                End With
            End Using
        Else
            If SearchAllPanes = False Then
                FindReplace_NotFound()
            Else
                FindReplace_MoveToNextTab()
            End If
        End If
    End Sub

    Private Sub InitialiseFindReplace(ByVal objFindReplace As FindReplace)
        With objFindReplace
            If Me.MatchItem IsNot Nothing Then
                'previous match is starting point for search action
                .StartItem = Me.MatchItem
                .MatchItem = Me.MatchItem
                .MatchChildItem = Me.MatchChildItem
                .MatchProperty = Me.MatchProperty
                .MatchIndex = Me.MatchIndex
                .MatchLength = Me.MatchLength
            End If
            .ReplaceAll = Me.ReplaceAll
            .SearchAllPanes = Me.SearchAllPanes

            If My.Settings.setShowResourcesBudget = True Then
                .SearchResources = True
            Else
                .SearchResources = False
            End If
        End With
    End Sub

    Private Sub FindReplace_NotFound()
        Select Case intCurrentProjectTab
            Case 0 'Project info
                lblMessage.Text = LANG_FindNotInProjectInfo
                lblMessageReplace.Text = LANG_FindNotInProjectInfo
            Case 1 'Logframe
                lblMessage.Text = LANG_FindReachedBottomOfLogframe
                lblMessageReplace.Text = LANG_FindReachedBottomOfLogframe
            Case 2 'Planning
                lblMessage.Text = LANG_FindReachedBottomOfPlanning
                lblMessageReplace.Text = LANG_FindReachedBottomOfPlanning
            Case 3 'Budget
                lblMessage.Text = LANG_FindReachedBottomOfBudgetYear
                lblMessageReplace.Text = LANG_FindReachedBottomOfBudgetYear
        End Select
    End Sub

    Private Sub FindReplace_MoveToNextTab()
        If intCurrentProjectTab < 3 Then
            If intCurrentProjectTab = 2 Then intCurrentBudgetTab = 0
            intCurrentProjectTab += 1
        ElseIf intCurrentProjectTab = 3 Then
            If intCurrentBudgetTab < CurrentLogFrame.Budget.BudgetYears.Count - 1 Then
                intCurrentBudgetTab += 1
            Else
                intCurrentProjectTab += 1
            End If
        ElseIf intCurrentProjectTab > 3 Then
            intCurrentProjectTab = 0
            intCurrentBudgetTab = 0
        End If

        MatchItem = Nothing
        MatchChildItem = Nothing
        MatchProperty = String.Empty
        MatchIndex = -1
        MatchLength = 0

        If intCurrentProjectTab <> intStartProjectTab Then
            FindReplace(Me.FindText, Me.ReplaceText, True, Me.ReplaceAll)

            Exit Sub
        Else
            If intCurrentProjectTab = 3 And intCurrentBudgetTab <> intStartBudgetTab Then
                FindReplace(Me.FindText, Me.ReplaceText, True, Me.ReplaceAll)

                Exit Sub
            End If
        End If

        lblMessage.Text = LANG_FindNotInProject

        If Me.ReplaceAll = True Then
            ReplaceAll_Finalise()
        Else
            lblMessageReplace.Text = LANG_FindNotInProject
        End If
    End Sub

    Private Function GetCurrentObject(ByVal boolFromStart As Boolean) As Object
        Dim CurrentObject As Object = Nothing

        With CurrentProjectForm
            Select Case intCurrentProjectTab
                Case 0 'Project info
                    If CurrentControl Is Nothing Then CurrentProjectForm.ProjectInfo.tbProjectTitle.Focus()
                    CurrentObject = GetCurrentItem()
                    If CurrentObject Is Nothing Then CurrentObject = CurrentLogFrame
                Case 1 'Logframe
                    If boolFromStart = False Then
                        CurrentObject = .dgvLogframe.CurrentItem(False)
                    End If

                    If CurrentObject Is Nothing Then
                        CurrentObject = .dgvLogframe.GetFirstItem
                    End If
                Case 2 'Planning
                    If boolFromStart = False Then
                        CurrentObject = .dgvPlanning.CurrentItem(False)
                    End If

                    If CurrentObject Is Nothing Then
                        CurrentObject = .dgvPlanning.GetFirstItem
                    End If
                Case 3 'Budget
                    Dim dgvBudgetYear As DataGridViewBudgetYear = .GetDataGridViewBudgetYear(intCurrentBudgetTab) '.CurrentDataGridViewBudgetYear

                    If dgvBudgetYear IsNot Nothing Then
                        If boolFromStart = False Then
                            CurrentObject = dgvBudgetYear.CurrentItem(False)
                        End If

                        If CurrentObject Is Nothing Then
                            CurrentObject = dgvBudgetYear.GetFirstItem
                        End If
                    End If
            End Select
        End With

        Return CurrentObject
    End Function
#End Region

#Region "Show found/replaced text in its control"
    Private Sub ShowInControl()
        Select Case intCurrentProjectTab
            Case 0 'Project info
                ShowInControl_ProjectInfo()
            Case 1 'Logframe
                ShowInControl_Logframe()
            Case 2 'Planning
                ShowInControl_Planning()
            Case 3 'Budget
                ShowInControl_BudgetYear()
        End Select
    End Sub

    Public Function GetControlByPropertyName(ByVal objPane As Object) As Control
        Dim selControl As Control = Nothing
        Dim boolFound As Boolean

        If objPane Is Nothing Then Return Nothing

        For Each selControl In objPane.Controls
            Select Case selControl.GetType
                Case GetType(TabControl)
                    selControl.Focus()

                    selControl = GetControlByPropertyName_TabControl(selControl)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                Case GetType(SplitContainer)
                    Dim objSplitContainer As SplitContainer = CType(selControl, SplitContainer)
                    objSplitContainer.Focus()
                    selControl = GetControlByPropertyName(objSplitContainer.Panel1)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                    selControl = GetControlByPropertyName(objSplitContainer.Panel2)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                Case GetType(AssumptionAssumptions), GetType(AssumptionDependencies), GetType(AssumptionRisks), _
                    GetType(IndicatorFormula), GetType(IndicatorFrequencyLikert), GetType(IndicatorMaxDiff), GetType(IndicatorMixedSubIndicators), GetType(IndicatorMultipleOption), _
                    GetType(IndicatorOpenEnded), GetType(IndicatorRanking), GetType(IndicatorRatio), GetType(IndicatorScales), GetType(IndicatorTargetsLikertScale), _
                    GetType(IndicatorTargetsSemanticDiff), GetType(IndicatorValues), GetType(IndicatorYesNo), _
                    GetType(GroupBox), GetType(TableLayoutPanel), GetType(Panel)

                    selControl.Focus()
                    selControl = GetControlByPropertyName(selControl)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                Case Else
                    If selControl.GetType IsNot GetType(Label) AndAlso GetControlByPropertyName_VerifyNameAndType(selControl.Name) Then
                        If String.IsNullOrEmpty(Me.ReplaceText) = False Then
                            If selControl.DataBindings.Count > 0 Then
                                selControl.DataBindings(0).ReadValue()
                            End If
                        End If
                        boolFound = True
                        Exit For
                    End If
            End Select
        Next

        If boolFound = True Then
            Return selControl
        Else
            Return Nothing
        End If

    End Function

    Private Function GetControlByPropertyName_VerifyNameAndType(ByVal strName As String) As Boolean
        Dim intIndex As Integer
        Dim intExpectedLength As Integer
        Dim strType As String

        If strName.Contains(Me.MatchProperty) Then
            intIndex = strName.IndexOf(Me.MatchProperty)
            intExpectedLength = intIndex + MatchProperty.Length

            If strName.Length > intExpectedLength + 1 Then '+1: allow for indexed controls of same property; for instance ntbNrDecimals1 and ntbNrDecimals2
                Return False
            Else
                strType = strName.Substring(0, intIndex)

                Select Case strType
                    Case "tb", "dtb", "ntb", "rtb", "cmb", "dtp"
                        Return True
                End Select
            End If
        End If

        Return False
    End Function

    Private Function GetControlByPropertyName_TabControl(ByVal selTabControl As TabControl) As Control
        Dim selControl As Control = Nothing
        Dim boolFound As Boolean

        For Each selPage As TabPage In selTabControl.TabPages
            selPage.Focus()

            selControl = GetControlByPropertyName(selPage)
            If selControl IsNot Nothing Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = True Then
            Return selControl
        Else
            Return Nothing
        End If

    End Function
#End Region

#Region "Show in project info"
    Private Sub ShowInControl_ProjectInfo()
        Dim selControl As Control = Nothing
        Dim strProperty As String = Me.MatchProperty

        If MatchChildItem Is Nothing Then
            selControl = GetControlByPropertyName(CurrentProjectForm.ProjectInfo)

            If selControl IsNot Nothing Then
                MarkTextInControl(selControl)
            End If
        Else
            ShowInControl_ChildItems()
        End If
    End Sub
#End Region

#Region "Show in logframe"
    Private Sub ShowInControl_Logframe()
        Dim selControl As Control = Nothing
        Dim strProperty As String = Me.MatchProperty

        With CurrentProjectForm.dgvLogframe
            Select Case MatchItem.GetType
                Case GetType(Goal), GetType(Purpose), GetType(Output), GetType(Activity), GetType(Indicator), GetType(Resource), GetType(VerificationSource), GetType(Assumption)
                    .EnsureSectionIsVisible(MatchItem.Section)

                    If MatchItem.Section = LogFrame.SectionTypes.ActivitiesSection Then
                        .EnsureIndicatorsResourcesAreVisible(MatchItem)
                    End If

                    If MatchItem.GetType = GetType(Resource) And MatchProperty = "TotalCostAmount" Then
                        .SetFocusOnItem(MatchItem, True)
                        strProperty = String.Empty
                    Else
                        .SetFocusOnItem(MatchItem)
                    End If


                    If String.IsNullOrEmpty(strProperty) Then
                        .HighlightTextInCell(MatchIndex, MatchLength)
                    Else
                        If My.Settings.setShowDetailsLogframe = False Then
                            CurrentProjectForm.ShowDetailsLogframe()
                        End If
                        CurrentProjectForm.SetTypeOfDetailWindowLogframe()

                        If MatchChildItem Is Nothing Then
                            selControl = ShowInControl_Logframe_DetailPanel()

                            If selControl IsNot Nothing Then
                                Select Case CurrentProjectForm.CurrentDetailLogframe.GetType
                                    Case GetType(DetailIndicator)
                                        Dim objDetailIndicator As DetailIndicator = TryCast(CurrentProjectForm.CurrentDetailLogframe, DetailIndicator)

                                        If objDetailIndicator IsNot Nothing Then
                                            objDetailIndicator.TabControlIndicator.SelectedIndex = 0

                                        End If
                                End Select
                                MarkTextInControl(selControl)
                            End If
                        Else
                            ShowInControl_ChildItems()
                        End If
                    End If
            End Select

        End With
    End Sub

    Private Function ShowInControl_Logframe_DetailPanel() As Control
        Dim selControl As Control = Nothing

        Select Case MatchItem.GetType
            Case GetType(Goal), GetType(Indicator), GetType(VerificationSource), GetType(Assumption)
                selControl = GetControlByPropertyName(CurrentProjectForm.CurrentDetailLogframe)
            Case GetType(Resource)
                selControl = CType(CurrentProjectForm.CurrentDetailLogframe, DetailResource).dgvBudgetItemReferences
        End Select

        Return selControl
    End Function

    Private Sub MarkTextInControl(ByVal selControl As Control)
        selControl.Focus()

        Select Case selControl.GetType
            Case GetType(TextBoxLF), GetType(TextBox)
                Dim ctlTextBox As TextBox = CType(selControl, TextBox)
                ctlTextBox.Select(MatchIndex, MatchLength)
            Case GetType(NumericTextBoxLF)
                selControl.Select()
            Case GetType(DataGridViewBudgetItemReferences)
                CType(selControl, DataGridViewBudgetItemReferences).SetFocusOnItem(MatchChildItem, MatchProperty)
            Case Else
                If selControl.CanSelect Then selControl.Select()
        End Select
    End Sub

    Private Sub ShowInControl_ChildItems()
        Select Case MatchItem.GetType
            Case GetType(LogFrame)
                ShowInControl_ChildItems_Logframe()
            Case GetType(Purpose)
                ShowInControl_ChildItems_Purpose()
            Case GetType(Output)
                ShowInControl_ChildItems_Output()
            Case GetType(Activity)
                ShowInControl_ChildItems_Activity()
            Case GetType(Indicator)
                ShowInControl_ChildItems_Indicator()
        End Select
    End Sub

    Private Sub ShowInControl_ChildItems_Logframe()
        Dim objDetailProjectInfo As DetailProjectInfo = CType(CurrentProjectForm.ProjectInfo, DetailProjectInfo)
        Dim selLogframe As LogFrame = CType(MatchItem, LogFrame)

        If objDetailProjectInfo IsNot Nothing And selLogframe IsNot Nothing Then
            Select Case MatchChildItem.GetType
                Case GetType(TargetGroup)
                    Dim selTargetGroup As TargetGroup = DirectCast(MatchChildItem, TargetGroup)

                    objDetailProjectInfo.lvTargetGroups.Focus()
                    objDetailProjectInfo.lvTargetGroups.SetFocusOnItem(selTargetGroup)

                    Dim dialogTargetGroup As New DialogTargetGroup(selTargetGroup)
                    lstOpenDialogs.Add(dialogTargetGroup)
                    dialogTargetGroup.Show(Me)

                    Dim selControl As Control = GetControlByPropertyName(dialogTargetGroup)
                    If selControl IsNot Nothing Then
                        MarkTextInControl(selControl)
                    End If

                    dialogTargetGroup.TopMost = True
                Case GetType(TargetGroupInformation)
                    Dim selTargetGroupInfo As TargetGroupInformation = DirectCast(MatchChildItem, TargetGroupInformation)
                    Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(selTargetGroupInfo.ParentTargetGroupGuid)

                    objDetailProjectInfo.lvTargetGroups.Focus()

                    If selTargetGroup IsNot Nothing Then
                        objDetailProjectInfo.lvTargetGroups.SetFocusOnItem(selTargetGroup)

                        Dim dialogTargetGroup As New DialogTargetGroup(selTargetGroup)
                        lstOpenDialogs.Add(dialogTargetGroup)
                        dialogTargetGroup.Show(Me)

                        Dim dialogTargetGroupInfo As New DialogTargetGroupInformation(selTargetGroupInfo)
                        lstOpenDialogs.Add(dialogTargetGroupInfo)
                        dialogTargetGroupInfo.Show(dialogTargetGroup)

                        Dim selControl As Control = GetControlByPropertyName(dialogTargetGroupInfo)
                        If selControl IsNot Nothing Then
                            MarkTextInControl(selControl)
                        End If

                        dialogTargetGroupInfo.TopMost = True
                    End If
            End Select
        End If
    End Sub

    Private Sub ShowInControl_ChildItems_Purpose()
        Dim objDetailPurpose As DetailPurpose = CType(CurrentProjectForm.CurrentDetailLogframe, DetailPurpose)
        Dim selPurpose As Purpose = CType(MatchItem, Purpose)

        If objDetailPurpose IsNot Nothing And selPurpose IsNot Nothing Then
            Select Case MatchChildItem.GetType
                Case GetType(TargetGroup)
                    Dim selTargetGroup As TargetGroup = DirectCast(MatchChildItem, TargetGroup)

                    objDetailPurpose.lvDetailTargetGroups.Focus()
                    objDetailPurpose.lvDetailTargetGroups.SetFocusOnItem(selTargetGroup)

                    Dim dialogTargetGroup As New DialogTargetGroup(selTargetGroup)
                    lstOpenDialogs.Add(dialogTargetGroup)
                    dialogTargetGroup.Show(Me)

                    Dim selControl As Control = GetControlByPropertyName(dialogTargetGroup)
                    If selControl IsNot Nothing Then
                        MarkTextInControl(selControl)
                    End If

                    dialogTargetGroup.TopMost = True
                Case GetType(TargetGroupInformation)
                    Dim selTargetGroupInfo As TargetGroupInformation = DirectCast(MatchChildItem, TargetGroupInformation)
                    Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(selTargetGroupInfo.ParentTargetGroupGuid)

                    objDetailPurpose.lvDetailTargetGroups.Focus()

                    If selTargetGroup IsNot Nothing Then
                        objDetailPurpose.lvDetailTargetGroups.SetFocusOnItem(selTargetGroup)

                        Dim dialogTargetGroup As New DialogTargetGroup(selTargetGroup)
                        lstOpenDialogs.Add(dialogTargetGroup)
                        dialogTargetGroup.Show(Me)

                        Dim dialogTargetGroupInfo As New DialogTargetGroupInformation(selTargetGroupInfo)
                        lstOpenDialogs.Add(dialogTargetGroupInfo)
                        dialogTargetGroupInfo.Show(dialogTargetGroup)

                        Dim selControl As Control = GetControlByPropertyName(dialogTargetGroupInfo)
                        If selControl IsNot Nothing Then
                            MarkTextInControl(selControl)
                        End If

                        dialogTargetGroupInfo.TopMost = True
                    End If
            End Select
        End If
    End Sub

    Private Sub ShowInControl_ChildItems_Output()
        Dim objDetailOutput As DetailOutput = CType(CurrentProjectForm.CurrentDetailLogframe, DetailOutput)
        Dim selOutput As Output = CType(MatchItem, Output)

        If objDetailOutput IsNot Nothing And selOutput IsNot Nothing Then
            Select Case MatchChildItem.GetType
                Case GetType(KeyMoment)
                    Dim selKeyMoment As KeyMoment = DirectCast(MatchChildItem, KeyMoment)

                    objDetailOutput.lvDetailKeyMoments.Focus()
                    objDetailOutput.lvDetailKeyMoments.SetFocusOnItem(selKeyMoment)

                    Dim dialogKeyMoment As New DialogKeyMoment(selKeyMoment)
                    lstOpenDialogs.Add(dialogKeyMoment)
                    dialogKeyMoment.Show(Me)

                    Dim selControl As Control = GetControlByPropertyName(dialogKeyMoment)
                    If selControl IsNot Nothing Then
                        MarkTextInControl(selControl)
                    End If

                    dialogKeyMoment.TopMost = True
            End Select
        End If
    End Sub

    Private Sub ShowInControl_ChildItems_Activity()
        Dim objDetailActivity As DetailActivity = Nothing
        Dim selActivity As Activity = CType(MatchItem, Activity)
        Dim selControl As Control = Nothing

        Select Case intCurrentProjectTab
            Case 1 'Logframe
                objDetailActivity = CType(CurrentProjectForm.CurrentDetailLogframe, DetailActivity)
            Case 2 'Planning
                objDetailActivity = CType(CurrentProjectForm.CurrentDetailPlanning, DetailActivity)
        End Select

        If objDetailActivity IsNot Nothing And selActivity IsNot Nothing Then
            With objDetailActivity
                Select Case MatchProperty
                    Case "StartDate", "Period", "Duration"
                        .TabControlActivity.SelectedTab = .TabPageDuration

                        selControl = GetControlByPropertyName(.TabPageDuration)
                    Case "Organisation", "Location"
                        .TabControlActivity.SelectedTab = .TabPageOrganisation

                        selControl = GetControlByPropertyName(.TabPageOrganisation)
                    Case "Preparation", "FollowUp"
                        .TabControlActivity.SelectedTab = .TabPagePreparation

                        selControl = GetControlByPropertyName(.TabPagePreparation)
                    Case "RepeatEvery", "RepeatTimes"
                        .TabControlActivity.SelectedTab = .TabPageRepeat

                        selControl = GetControlByPropertyName(.TabPageRepeat)
                End Select
            End With
        End If

        If selControl IsNot Nothing Then
            MarkTextInControl(selControl)
        End If
    End Sub

    Private Sub ShowInControl_ChildItems_Indicator()
        Dim objDetailIndicator As DetailIndicator = CType(CurrentProjectForm.CurrentDetailLogframe, DetailIndicator)
        Dim selIndicator As Indicator = CType(MatchItem, Indicator)
        Dim boolShowScore, boolShowPopulationTarget As Boolean

        If objDetailIndicator IsNot Nothing And selIndicator IsNot Nothing Then
            Select Case MatchChildItem.GetType
                Case GetType(ResponseClass)
                    ShowInControl_ChildItems_ResponseClasses(objDetailIndicator)
                Case GetType(Statement)
                    ShowInControl_ChildItems_Statements(objDetailIndicator)
                Case GetType(Baseline), GetType(Target)
                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                            If MatchProperty = "Score" Then boolShowScore = True
                            ShowInControl_ChildItems_Targets(objDetailIndicator, boolShowScore, False)
                        Case Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.Ratio
                            If MatchProperty = "Score" Then boolShowScore = True
                            ShowInControl_ChildItems_Targets(objDetailIndicator, boolShowScore, False)
                        Case Indicator.QuestionTypes.SemanticDiff, Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale, _
                            Indicator.QuestionTypes.LikertTypeScale, Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert
                            ShowInControl_ChildItems_Targets(objDetailIndicator, False, False)
                    End Select
                Case GetType(DoubleValue)
                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Ranking, Indicator.QuestionTypes.FrequencyLikert
                            ShowInControl_ChildItems_Targets(objDetailIndicator, False, False)
                    End Select
                Case GetType(PopulationTarget)
                    If MatchProperty = "Percentage" Then boolShowPopulationTarget = True
                    ShowInControl_ChildItems_Targets(objDetailIndicator, False, boolShowPopulationTarget)
            End Select
        End If
    End Sub

    Private Sub ShowInControl_ChildItems_ResponseClasses(ByVal objDetailIndicator As DetailIndicator)
        Dim selIndicator As Indicator = CType(MatchItem, Indicator)

        objDetailIndicator.TabControlIndicator.SelectedIndex = 0

        Select Case selIndicator.QuestionType
            Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, _
                Indicator.QuestionTypes.LikertTypeScale, Indicator.QuestionTypes.SemanticDiff, Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert

                Dim dgvResponseClasses As DataGridViewResponseClasses = CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewResponseClasses)

                If dgvResponseClasses IsNot Nothing Then
                    With dgvResponseClasses
                        If MatchProperty = "Value" Then
                            .SetFocusOnItem(MatchChildItem, True)
                            .BeginEdit(True)
                        Else
                            .SetFocusOnItem(MatchChildItem, False)
                            .BeginEdit(False)

                            Dim ctlEdit As TextBox = CType(.EditingControl, TextBox)
                            ctlEdit.Select(MatchIndex, MatchLength)
                        End If

                    End With
                End If
            Case Indicator.QuestionTypes.Ranking
                Dim dgvResponseClasses As DataGridViewResponseClasses = CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewResponseClasses)

                If dgvResponseClasses IsNot Nothing Then
                    With dgvResponseClasses
                        .SetFocusOnItem(MatchChildItem, False)
                        .BeginEdit(False)

                        Dim ctlEdit As TextBox = CType(.EditingControl, TextBox)
                        ctlEdit.Select(MatchIndex, MatchLength)
                    End With
                End If
            Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                Dim dgvStatements As DataGridViewStatementsScales = CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsScales)

                If dgvStatements IsNot Nothing Then
                    With dgvStatements
                        .Focus()
                        .SetFocusOnItem(MatchChildItem)
                        .BeginEdit(True)
                    End With
                End If
        End Select
    End Sub

    Private Sub ShowInControl_ChildItems_Statements(ByVal objDetailIndicator As DetailIndicator)
        Dim selIndicator As Indicator = CType(MatchItem, Indicator)
        Dim selStatement As Statement = CType(MatchChildItem, Statement)

        If selStatement Is Nothing Then Exit Sub

        Select Case selIndicator.QuestionType
            Case Indicator.QuestionTypes.Ratio
                objDetailIndicator.TabControlIndicator.SelectedIndex = 0

                Dim ctlRatio As IndicatorRatio = CType(objDetailIndicator.PanelQuestionType.Controls(0), IndicatorRatio)
                Dim intIndex As Integer = selIndicator.Statements.IndexOf(selStatement)

                If ctlRatio IsNot Nothing Then
                    Select Case intIndex
                        Case 0
                            ctlRatio.rtbFirstStatement.Focus()
                            ctlRatio.rtbFirstStatement.Select(MatchIndex, MatchLength)
                        Case 1
                            ctlRatio.rtbSecondStatement.Focus()
                            ctlRatio.rtbSecondStatement.Select(MatchIndex, MatchLength)
                    End Select

                End If
            Case Indicator.QuestionTypes.Formula
                objDetailIndicator.TabControlIndicator.SelectedIndex = 0
                Dim dgvStatements As DataGridViewStatementsFormula = CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsFormula)

                If dgvStatements IsNot Nothing Then
                    With dgvStatements
                        .Focus()
                        .SetFocusOnItem(MatchChildItem, MatchProperty, MatchIndex, MatchLength)
                    End With
                End If
            Case Indicator.QuestionTypes.SemanticDiff
                objDetailIndicator.TabControlIndicator.SelectedIndex = 2
                Dim ctlSemanticDiff As IndicatorTargetsSemanticDiff = CType(objDetailIndicator.PanelTargets.Controls(0), IndicatorTargetsSemanticDiff)

                If ctlSemanticDiff IsNot Nothing Then
                    ctlSemanticDiff.SetFocusOnItem(MatchChildItem, MatchProperty, MatchIndex, MatchLength)
                End If
            Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                objDetailIndicator.TabControlIndicator.SelectedIndex = 0
                Dim dgvStatements As DataGridViewStatementsScales = CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsScales)

                If dgvStatements IsNot Nothing Then
                    With dgvStatements
                        .Focus()
                        .SetFocusOnItem(MatchChildItem)
                        .BeginEdit(False)

                        Dim ctlEdit As RichTextEditingControlLogframe = CType(.EditingControl, RichTextEditingControlLogframe)
                        ctlEdit.Select(MatchIndex, MatchLength)
                    End With
                End If
            Case Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert
                objDetailIndicator.TabControlIndicator.SelectedIndex = 2
                Dim ctlLikert As IndicatorTargetsLikertScale = CType(objDetailIndicator.PanelTargets.Controls(0), IndicatorTargetsLikertScale)

                If ctlLikert IsNot Nothing Then
                    ctlLikert.SetFocusOnItem(MatchChildItem, MatchIndex, MatchLength)
                End If
        End Select
    End Sub

    Private Sub ShowInControl_ChildItems_Targets(ByVal objDetailIndicator As DetailIndicator, ByVal boolShowScore As Boolean, ByVal boolShowPopulationTarget As Boolean)
        Dim selIndicator As Indicator = CType(MatchItem, Indicator)

        objDetailIndicator.TabControlIndicator.SelectedIndex = 2

        Select Case selIndicator.QuestionType
            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                Dim dgvTargets As DataGridViewTargetsValues = CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsValues)

                If dgvTargets IsNot Nothing Then
                    dgvTargets.SetFocusOnItem(MatchChildItem, boolShowScore, boolShowPopulationTarget)

                    If MatchProperty = "Formula" Then
                        Dim intTargetIndex As Integer
                        Dim selTarget As Target = TryCast(MatchChildItem, Target)

                        If selTarget IsNot Nothing Then intTargetIndex = selIndicator.Targets.IndexOf(selTarget)
                        dgvTargets.ShowDialog_Targets(intTargetIndex)
                    Else
                        dgvTargets.BeginEdit(True)
                    End If
                    dgvTargets.BeginEdit(True)
                End If
            Case Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.Ratio
                Dim dgvTargets As DataGridViewTargetsFormula = CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsFormula)

                If dgvTargets IsNot Nothing Then
                    dgvTargets.SetFocusOnItem(MatchChildItem, boolShowScore)
                    dgvTargets.BeginEdit(True)
                End If
            Case Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions
                Dim dgvTargets As DataGridViewTargetsClasses = CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsClasses)

                If dgvTargets IsNot Nothing Then
                    dgvTargets.SetFocusOnItem(MatchChildItem)
                    dgvTargets.BeginEdit(True)
                End If
            Case Indicator.QuestionTypes.Ranking
                Dim dgvTargets As DataGridViewTargetsRanking = CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsRanking)

                If dgvTargets IsNot Nothing Then
                    dgvTargets.SetFocusOnItem(MatchChildItem)
                    'dgvTargets.BeginEdit(True)
                End If
            Case Indicator.QuestionTypes.LikertTypeScale
                Dim dgvTargets As DataGridViewTargetsScaleLikertType = CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScaleLikertType)

                If dgvTargets IsNot Nothing Then
                    dgvTargets.SetFocusOnItem(MatchChildItem)
                    dgvTargets.BeginEdit(True)
                End If
            Case Indicator.QuestionTypes.SemanticDiff
                Dim ctlSemanticDiff As IndicatorTargetsSemanticDiff = CType(objDetailIndicator.PanelTargets.Controls(0), IndicatorTargetsSemanticDiff)

                If ctlSemanticDiff IsNot Nothing Then
                    ctlSemanticDiff.SetFocusOnItem(MatchChildItem)
                End If
            Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                Dim dgvTargets As DataGridViewTargetsScales = CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScales)

                If dgvTargets IsNot Nothing Then
                    dgvTargets.SetFocusOnItem(MatchChildItem)
                    dgvTargets.BeginEdit(True)
                End If

            Case Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert
                Dim ctlFrequencyLikert As IndicatorTargetsLikertScale = CType(objDetailIndicator.PanelTargets.Controls(0), IndicatorTargetsLikertScale)

                If ctlFrequencyLikert IsNot Nothing Then
                    ctlFrequencyLikert.SetFocusOnItem(MatchChildItem)
                End If
        End Select
    End Sub

    Private Sub ReplaceAll_Finalise()
        With CurrentProjectForm
            If Me.ReplaceAll = True Then

                With .ProjectInfo
                    For Each selControl As Control In .Controls
                        If selControl.DataBindings.Count > 0 Then
                            selControl.DataBindings(0).ReadValue()
                        End If
                    Next
                End With
                .dgvLogframe.Reload()
                .dgvPlanning.Reload()
                .ReloadAllDataGridViewBudgetYears()
            Else
                Select Case intCurrentProjectTab
                    Case 0 'Project info
                        With .ProjectInfo
                            For Each selControl As Control In .Controls
                                If selControl.DataBindings.Count > 0 Then
                                    selControl.DataBindings(0).ReadValue()
                                End If
                            Next
                        End With
                    Case 1 'Logframe
                        CurrentProjectForm.dgvLogframe.Reload()
                    Case 2 'Planning
                        .dgvPlanning.Reload()
                    Case 3 'Budget
                        .ReloadAllDataGridViewBudgetYears()
                End Select
            End If
        End With

        lblMessageReplace.Text = String.Format(LANG_ReplacedItems, Me.ItemsReplaced)
    End Sub
#End Region

#Region "Show in planning"
    Private Sub ShowInControl_Planning()
        Dim selControl As Control = Nothing
        Dim strProperty As String = Me.MatchProperty

        With CurrentProjectForm.dgvPlanning
            .SetFocusOnItem(MatchItem)

            If MatchItem.GetType Is GetType(KeyMoment) And strProperty = "Description" Then
                .HighlightTextInCell(MatchIndex, MatchLength)
            ElseIf MatchItem.GetType Is GetType(Activity) And String.IsNullOrEmpty(strProperty) Then
                .HighlightTextInCell(MatchIndex, MatchLength)
            Else
                If My.Settings.setShowDetailsPlanning = False Then
                    CurrentProjectForm.ShowDetailsPlanning()
                End If
                CurrentProjectForm.SetTypeOfDetailWindowPlanning()

                If MatchChildItem Is Nothing Then
                    selControl = GetControlByPropertyName(CurrentProjectForm.CurrentDetailPlanning)

                    If selControl IsNot Nothing Then
                        MarkTextInControl(selControl)
                    End If
                Else
                    ShowInControl_ChildItems()
                End If
            End If

        End With
    End Sub
#End Region

#Region "Show in budget year"
    Private Sub ShowInControl_BudgetYear()
        Dim selControl As Control = Nothing
        Dim strProperty As String = Me.MatchProperty

        With CurrentProjectForm.CurrentDataGridViewBudgetYear
            .SetFocusOnItem(MatchItem, MatchProperty)

            'If String.IsNullOrEmpty(strProperty) Then
            .HighlightTextInCell(MatchIndex, MatchLength)
            'Else
            'If My.Settings.setShowDetailsBudget = False Then
            '    CurrentProjectForm.ShowDetailsBudget()
            'End If
            'CurrentProjectForm.SetTypeOfDetailWindowBudget(intCurrentBudgetTab)

            'If MatchChildItem Is Nothing Then
            '    selControl = GetControlByPropertyName(CurrentProjectForm.CurrentDetailPlanning)

            '    If selControl IsNot Nothing Then
            '        MarkTextInControl(selControl)
            '    End If
            'Else
            '    ShowInControl_ChildItems()
            'End If
            'End If

        End With
    End Sub
#End Region
End Class
