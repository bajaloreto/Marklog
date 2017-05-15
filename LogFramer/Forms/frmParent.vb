Public Class frmParent
    Friend WithEvents ColorPicker As New ucColorPaletteDialog
    Friend WithEvents SplitButtonTextColor As New ToolStripSplitButton
    Friend WithEvents TimerAutoSave As New Windows.Forms.Timer

    Private CurrentProjectButton As RibbonButton
    Private ControlKeyDown As Boolean

#Region "Initialise"
    Private Sub frmParent_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        'If there is text on the Windows clipboard (copied from another application), copy it to the Logframer text clipboard
        TextClipboard.CheckClipboardContentInList()
    End Sub

    Private Sub frmParent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Maximized

        'screenshot mode

        'Me.WindowState = FormWindowState.Normal
        'Me.Size = New Size(1200, 800)

        'Enable the use of shortcut keys (see 'Shortcut keys' region in this document
        Me.KeyPreview = True

        'remove in version 3.0
        Me.RibbonPanelCloud.Visible = False

        'pen for selection rectangles
        penSelection.DashStyle = Drawing2D.DashStyle.Dot

        LoadTranslations()
        LoadFontLibrary()
        LoadFontSizes()

        Initialize_FontSystem()
        Initialize_DefaultProject()
        Initialize_Panels()
        Initialize_Clipboard()

        UpdateRecentFilesList()
    End Sub

    Private Sub Initialize_DefaultProject()
        'Default project when no other projects are opened (at start-up) 
        'or when all existing projects are closed

        AddNewProject()
        CurrentProjectButton = RibbonButtonProject1
        AddHandler RibbonButtonProject1.Click, AddressOf RibbonButtonProject_Click
        RibbonButtonProject1.Checked = True
    End Sub

    Private Sub Initialize_FontSystem()
        'Fonts
        CurrentText.Font = My.Settings.setDefaultFont
        Dim objFontCollection As System.Drawing.Text.InstalledFontCollection = New System.Drawing.Text.InstalledFontCollection
        Dim i As Integer

        With RibbonComboBoxFontName
            For i = 0 To FontLibrary.Count - 1
                .DropDownItems.Add(New RibbonButton(FontLibrary(i)))
            Next
        End With

        'Fontsizes
        With RibbonComboBoxFontSize
            For i = 0 To FontSizes.Count - 1
                .DropDownItems.Add(New RibbonButton(FontSizes(i)))
            Next
        End With
    End Sub

    Private Sub Initialize_Panels()
        With CurrentProjectForm
            RibbonButtonDetailsWindow.Text = .Initialize_PanelsLogframe()
            RibbonButtonPlanningDetailsWindow.Text = .Initialize_PanelsPlanning()
            RibbonButtonBudgetDetailsWindow.Text = .Initialize_PanelsBudget()
        End With
    End Sub

    Private Sub Initialize_Clipboard()
        For Each selForm As frmProject In Me.MdiChildren
            With selForm.SplitContainerUtilities
                If .Panel1Collapsed = False Then .Panel1Collapsed = True
            End With
        Next
    End Sub

    Private Sub Initialize_ExchangeRates()
        For Each selForm As frmProject In Me.MdiChildren
            With selForm.SplitContainerExchangeRates
                If .Panel2Collapsed = False Then .Panel2Collapsed = True
            End With
        Next
    End Sub
#End Region

#Region "Quick access bar"
    Private Sub RibbonButtonQuickSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonQuickSave.Click
        Me.SaveProject(False)
    End Sub

    Private Sub RibbonButtonUndo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonUndo.Click
        UndoRedo.Undo(0)
    End Sub

    Private Sub RibbonButtonUndo_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.RibbonItemEventArgs) Handles RibbonButtonUndo.DropDownItemClicked
        Dim intIndex As Integer = RibbonButtonUndo.DropDownItems.IndexOf(e.Item)
        UndoRedo.Undo(intIndex)
    End Sub

    Public Sub ReloadSplitUndoRedoButtons()

        With RibbonButtonUndo
            .DropDownItems.Clear()
            For Each selItem As UndoListItem In CurrentUndoList
                Dim NewButton As New RibbonButton(selItem.Description)

                NewButton.MaxSizeMode = RibbonElementSizeMode.Medium
                NewButton.MinSizeMode = RibbonElementSizeMode.Medium
                .DropDownItems.Add(NewButton)
            Next
        End With
        With RibbonButtonRedo
            .DropDownItems.Clear()
            For Each selItem As UndoListItem In CurrentRedoList
                .DropDownItems.Add(New RibbonButton(selItem.Description))
            Next
        End With

        StatusLabelGeneral.Text = LANG_Ready
    End Sub

    Private Sub RibbonButtonRedo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonRedo.Click
        UndoRedo.Redo(0)
    End Sub

    Private Sub RibbonButtonRedo_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.RibbonItemEventArgs) Handles RibbonButtonRedo.DropDownItemClicked
        Dim intIndex As Integer = RibbonButtonRedo.DropDownItems.IndexOf(e.Item)
        UndoRedo.Redo(intIndex)
    End Sub
#End Region

#Region "Shortcut keys"
    Private Sub frmParent_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ControlKeyDown = False Then
            If e.Control And e.KeyCode = Keys.D1 Then
                '<Ctrl><1> Hide first column
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowProjectLogic()
                        Case GetType(DataGridViewPlanning)

                        Case GetType(DataGridViewBudgetYear)

                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D2 Then
                '<Ctrl><2> Hide second column
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowIndicators()
                        Case GetType(DataGridViewPlanning)
                            ShowPlanningDates()
                        Case GetType(DataGridViewBudgetYear)
                            ShowDuration()
                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D3 Then
                '<Ctrl><3> Hide third column
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowVerificationSources()
                        Case GetType(DataGridViewPlanning)

                        Case GetType(DataGridViewBudgetYear)
                            ShowLocalCurrency()
                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D4 Then
                '<Ctrl><4> Hide fourth column
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowAssumptions()
                        Case GetType(DataGridViewPlanning)

                        Case GetType(DataGridViewBudgetYear)

                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D5 Then
                '<Ctrl><5> Hide goals section
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowGoals()
                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D6 Then
                '<Ctrl><6> Hide purposes section
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowPurposes()
                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D7 Then
                '<Ctrl><7> Hide outputs section
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowOutputs()
                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D8 Then
                '<Ctrl><8> Hide activities section
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            ShowActivities()
                    End Select
                End If
                ControlKeyDown = True


            ElseIf e.Control And e.KeyCode = Keys.A Then
                '<Ctrl><A> Copy item or text
                CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectAll)
            ElseIf e.Control And e.KeyCode = Keys.B Then
                '<Ctrl><B> Text bold
                FontBold()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.C Then
                '<Ctrl><C> Copy item or text
                If CurrentClipboardType = DataGridViewClipboard.ContentTypes.Text Then
                    CopyText()
                Else
                    CopyItem()
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.D Then
                '<Ctrl><D> Show Detail pane
                Select Case CurrentMainTab
                    Case 1 'Logframe
                        RibbonButtonDetailsWindow.Text = CurrentProjectForm.ShowDetailsLogframe()
                    Case 2 'Planning
                        RibbonButtonPlanningDetailsWindow.Text = CurrentProjectForm.ShowDetailsPlanning()
                    Case 3 'Budget
                        RibbonButtonBudgetDetailsWindow.Text = CurrentProjectForm.ShowDetailsBudget()
                End Select
            ElseIf e.Control And e.KeyCode = Keys.E Then
                '<Ctrl><E> Hide empty cells
                If CurrentDataGridView IsNot Nothing Then
                    Select Case CurrentDataGridView.GetType
                        Case GetType(DataGridViewLogframe)
                            HideEmptyCellsLogframe()
                        Case GetType(DataGridViewPlanning)
                            HideEmptyCellsPlanning()
                        Case GetType(DataGridViewBudgetYear)
                            HideEmptyCellsBudgetYear()
                    End Select
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.F Then
                '<Ctrl><F> Find
                DisplayFindDialog()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.H Then
                '<Ctrl><H> Replace
                DisplayReplaceDialog()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.I Then
                '<Ctrl><I> Text italic
                FontItalic()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.L Then
                '<Ctrl><L> Align left
                CurrentText.HorizontalAlignment = HorizontalAlignment.Left
                CurrentProjectForm.ChangeParagraphAlignment()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.M Then
                '<Ctrl><M> Align center
                CurrentText.HorizontalAlignment = HorizontalAlignment.Center
                CurrentProjectForm.ChangeParagraphAlignment()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.N Then
                '<Ctrl><N> New project
                NewProject()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.O Then
                '<Ctrl><O> Open project
                OpenProject()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.P Then
                '<Ctrl><P> Print report
                ShortCut_Print()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.Q Then
                '<Ctrl><Q> Quit
                Quit()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.R Then
                '<Ctrl><R> Align right
                CurrentText.HorizontalAlignment = HorizontalAlignment.Right
                CurrentProjectForm.ChangeParagraphAlignment()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.S Then
                '<Ctrl><S> Save
                SaveProject(False)
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.U Then
                '<Ctrl><U> Text underline
                FontUnderline()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.V Then
                '<Ctrl><V> Paste item or text
                If CurrentClipboardType = DataGridViewClipboard.ContentTypes.Text Then
                    e.SuppressKeyPress = True
                    PasteText(0)
                Else
                    PasteItem()
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.W Then
                '<Ctrl><W> Close project
                CloseProject()
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.X Then
                '<Ctrl><X> Cut item or text
                If CurrentClipboardType = DataGridViewClipboard.ContentTypes.Text Then
                    CutText()
                Else
                    CutItem()
                End If
                ControlKeyDown = True
            ElseIf e.Control And e.KeyCode = Keys.U Then
                '<Ctrl><Y> Redo
                UndoRedo.Redo(0)
            ElseIf e.Control And e.KeyCode = Keys.Z Then
                '<Ctrl><Z> Undo
                UndoRedo.Undo(0)

            ElseIf e.KeyCode = Keys.Delete Then
                '<Del> Delete
                e.SuppressKeyPress = True
                RemoveItem()
                ControlKeyDown = True

            ElseIf e.KeyCode = Keys.F1 Then
                '<F1> Help
                Help.ShowHelp(Me, GetHelpPath(), HelpNavigator.TableOfContents)
            End If
        End If
    End Sub

    Private Sub frmParent_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        ControlKeyDown = False
    End Sub

    Private Sub ShortCut_Print()
        With CurrentProjectForm
            Select Case .TabControlProject.SelectedTab.Name
                Case .TabPageLogframe.Name
                    Preview_Logframe()
                Case .TabPagePlanning.Name
                    Preview_Planning()
                Case .TabPageBudget.Name
                    Preview_Budget()
                Case .TabPageProject.Name
                    With .ProjectInfo
                        Select Case .TabControlProjectInfo.SelectedTab.Name
                            Case .TabPageTargetGroups.Name
                                Preview_TargetGroupIdForm()
                            Case .TabPagePartners.Name
                                Preview_PartnerList()
                            Case .TabPageTargetDeadlines.Name
                                Preview_PMF()
                        End Select
                    End With
                Case Else
                    Preview_Logframe()
            End Select
        End With
    End Sub
#End Region

#Region "Ribbon-Visibility"
    Public Sub SetRibbonItemsVisibility()
        RibbonLF.SuspendLayout()
        RibbonLF.SuspendUpdating()

        'reset Enabled properties

        'Ribbon tabs
        RibbonTabText.Visible = True
        RibbonTabItems.Visible = True
        RibbonTabLayOutLogframe.Visible = True
        RibbonTabLayOutPlanning.Visible = True
        RibbonTabLayOutBudget.Visible = True
        RibbonTabCollaborate.Visible = True

        'Ribbon panels
        RibbonPanelTypeface.Enabled = True
        RibbonPanelParagraph.Enabled = True
        RibbonPanelClipboardItems.Enabled = True
        RibbonPanelAddRemove.Enabled = True
        RibbonPanelReferences.Enabled = True

        'Text tab
        RibbonButtonPasteKeepFormatting.Enabled = True
        RibbonButtonPasteMergeFormatting.Enabled = True
        RibbonButtonSelectTextInCell.Enabled = True
        RibbonButtonSelectTextInLogframe.Enabled = True
        RibbonButtonSelectAllText.Enabled = True
        RibbonComboBoxTextSelection.Enabled = True

        'Items tab
        RibbonButtonNewItem.Enabled = True
        RibbonButtonRemoveItem.Enabled = True
        RibbonButtonInsertItem.Enabled = True
        RibbonButtonInsertChild.Enabled = True
        RibbonButtonInsertParent.Enabled = True
        RibbonButtonMoveUp.Enabled = True
        RibbonButtonMoveDown.Enabled = True
        RibbonButtonSectionUp.Enabled = True
        RibbonButtonSectionDown.Enabled = True
        RibbonButtonLevelUp.Enabled = True
        RibbonButtonLevelDown.Enabled = True
        RibbonButtonLink.Enabled = True
        RibbonButtonUnlink.Enabled = True

        'Collaborate tab
        RibbonButtonSkype.Enabled = True

        'disable/hide
        Select Case CurrentControl.GetType
            Case GetType(DataGridViewLogframe)
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertChild.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

            Case GetType(DataGridViewPlanning)
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = True
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertChild.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False
                RibbonButtonLink.Enabled = True
                RibbonButtonUnlink.Enabled = True

            Case GetType(DataGridViewBudgetYear)
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False

                RibbonButtonInsertChild.Enabled = True
                RibbonButtonInsertParent.Enabled = True
                RibbonButtonLevelUp.Enabled = True
                RibbonButtonLevelDown.Enabled = True
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLink.Enabled = False
                RibbonButtonUnlink.Enabled = False

            Case GetType(DataGridViewResponseClasses)
                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

                If RibbonTabItems.Active = False Then RibbonLF.ActiveTab = RibbonTabItems

            Case GetType(DataGridViewStatementsFormula), GetType(DataGridViewStatementsMaxDiff), GetType(DataGridViewStatementsScales), _
                GetType(DataGridViewTargetsScaleLikert), GetType(DataGridViewTargetsFrequencyLikert), GetType(DataGridViewTargetsSemanticDiff)
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

                If RibbonTabText.Active = False And RibbonTabItems.Active = False Then RibbonLF.ActiveTab = RibbonTabItems

            Case GetType(DataGridViewTargetsClasses), GetType(DataGridViewTargetsFormula), GetType(DataGridViewTargetsMaxDiff), GetType(DataGridViewTargetsRanking), _
            GetType(DataGridViewTargetsScaleLikertType), GetType(DataGridViewTargetsScales), GetType(DataGridViewTargetsValues)

                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelAddRemove.Enabled = False
                RibbonPanelClipboardItems.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertItem.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonMoveUp.Enabled = False
                RibbonButtonMoveDown.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

                If RibbonTabItems.Active = False Then RibbonLF.ActiveTab = RibbonTabItems

            Case GetType(DataGridViewBudgetItemReferences)
                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False

                RibbonButtonInsertChild.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLink.Enabled = False
                RibbonButtonUnlink.Enabled = False
            Case GetType(ListViewAddresses), GetType(ListViewPartners)
                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertItem.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonMoveUp.Enabled = False
                RibbonButtonMoveDown.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False
                RibbonButtonSkype.Enabled = False

                If RibbonTabText.Active Or RibbonTabLayOutLogframe.Active Then RibbonLF.ActiveTab = RibbonTabCollaborate
            Case GetType(ListViewContacts)
                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertItem.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonMoveUp.Enabled = False
                RibbonButtonMoveDown.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

                If RibbonTabText.Active Or RibbonTabLayOutLogframe.Active Then RibbonLF.ActiveTab = RibbonTabCollaborate
            Case GetType(ListViewBudgetItemReferences), GetType(ListViewKeyMoments), GetType(ListViewProcesses), GetType(ListViewSubIndicatorsBase), _
                GetType(ListViewSubIndicatorsBeneficiary), GetType(ListViewSubIndicators), GetType(ListViewTargetGroups), GetType(ListViewTargetGroupInformations)

                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonInsertItem.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonMoveUp.Enabled = False
                RibbonButtonMoveDown.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

                If RibbonTabText.Active Or RibbonTabLayOutLogframe.Active Then RibbonLF.ActiveTab = RibbonTabItems
            Case GetType(ListViewTargetDeadlines), GetType(ListViewRiskMonitoringDeadlines)
                RibbonTabText.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                Dim boolShow As Boolean

                Select Case CurrentControl.GetType
                    Case GetType(ListViewTargetDeadlines)
                        If CType(CurrentControl, ListViewTargetDeadlines).TargetDeadlinesSection.Repetition = TargetDeadlinesSection.RepetitionOptions.UserSelect Then
                            boolShow = True
                        End If
                    Case GetType(ListViewRiskMonitoringDeadlines)
                        If CType(CurrentControl, ListViewRiskMonitoringDeadlines).RiskMonitoring.Repetition = RiskMonitoring.RepetitionOptions.UserSelect Then
                            boolShow = True
                        End If
                End Select
                If boolShow = True Then
                    RibbonTabItems.Visible = True

                    RibbonButtonInsertItem.Enabled = False
                    RibbonButtonInsertParent.Enabled = False
                    RibbonButtonInsertChild.Enabled = False
                    RibbonButtonMoveUp.Enabled = False
                    RibbonButtonMoveDown.Enabled = False
                    RibbonButtonSectionUp.Enabled = False
                    RibbonButtonSectionDown.Enabled = False
                    RibbonButtonLevelUp.Enabled = False
                    RibbonButtonLevelDown.Enabled = False

                    RibbonLF.ActiveTab = RibbonTabItems
                Else
                    RibbonLF.ActiveTab = RibbonTabFile
                    RibbonTabItems.Visible = False
                End If
                
            Case GetType(NumericTextBox), GetType(NumericTextBoxLF), _
                GetType(TextBox), GetType(TextBoxLF)

                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelTypeface.Enabled = False
                RibbonPanelParagraph.Enabled = False
                RibbonPanelClipboardItems.Enabled = False
                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonPasteKeepFormatting.Enabled = False
                RibbonButtonPasteMergeFormatting.Enabled = False
                RibbonButtonSelectTextInCell.Enabled = False
                RibbonButtonSelectTextInLogframe.Enabled = False
                RibbonButtonSelectAllText.Enabled = False
                RibbonComboBoxTextSelection.Enabled = False

                RibbonButtonNewItem.Enabled = False
                RibbonButtonInsertItem.Enabled = False
                RibbonButtonInsertParent.Enabled = False
                RibbonButtonInsertChild.Enabled = False
                RibbonButtonMoveUp.Enabled = False
                RibbonButtonMoveDown.Enabled = False
                RibbonButtonSectionUp.Enabled = False
                RibbonButtonSectionDown.Enabled = False
                RibbonButtonLevelUp.Enabled = False
                RibbonButtonLevelDown.Enabled = False

                If RibbonTabText.Active = False Then RibbonLF.ActiveTab = RibbonTabText
            Case GetType(CheckBox), GetType(CheckBoxLF), GetType(RadioButton), _
            GetType(ComboBox), GetType(ComboBoxSelectGuid), GetType(ComboBoxSelectIndex), GetType(ComboBoxSelectValue), GetType(ComboBoxText), _
            GetType(DateTextBox), GetType(DateTimePickerLF)

                'RibbonTabText.Visible = False
                'RibbonTabItems.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                If RibbonTabFile.Active = False Then RibbonLF.ActiveTab = RibbonTabFile
            Case GetType(RichTextBox)
                RibbonTabItems.Visible = False
                RibbonTabLayOutLogframe.Visible = False
                RibbonTabLayOutPlanning.Visible = False
                RibbonTabLayOutBudget.Visible = False
                RibbonTabCollaborate.Visible = False

                RibbonPanelClipboardItems.Enabled = False
                RibbonPanelLink.Enabled = False
                RibbonPanelReferences.Enabled = False

                RibbonButtonSelectTextInCell.Enabled = False
                RibbonButtonSelectTextInLogframe.Enabled = False
                RibbonButtonSelectAllText.Enabled = False
                RibbonComboBoxTextSelection.Enabled = False

                If RibbonTabText.Active = False Then RibbonLF.ActiveTab = RibbonTabText

            Case Else

        End Select
        RibbonLF.Invalidate()
        RibbonLF.ResumeLayout()
        RibbonLF.ResumeUpdating()
    End Sub

    Public Sub SetRibbonLayOutTabsVisibility()
        RibbonLF.SuspendLayout()
        RibbonLF.SuspendUpdating()

        With CurrentProjectForm
            Select Case .TabControlProject.SelectedTab.Name
                Case .TabPageLogframe.Name
                    RibbonTabLayOutLogframe.Visible = True
                    RibbonTabLayOutPlanning.Visible = False
                    RibbonTabLayOutBudget.Visible = False
                    RibbonTabCollaborate.Visible = False

                    RibbonPanelLink.Visible = False
                    RibbonPanelReferences.Visible = False

                    RibbonComboBoxTextSelection.Visible = True

                    If RibbonTabLayOutPlanning.Active = True Or RibbonTabLayOutBudget.Active = True Then RibbonLF.ActiveTab = RibbonTabLayOutLogframe
                Case .TabPagePlanning.Name
                    RibbonTabLayOutLogframe.Visible = False
                    RibbonTabLayOutPlanning.Visible = True
                    RibbonTabLayOutBudget.Visible = False
                    RibbonTabCollaborate.Visible = False

                    RibbonPanelLink.Visible = True
                    RibbonPanelReferences.Visible = False

                    RibbonComboBoxTextSelection.Visible = False

                    If RibbonTabLayOutLogframe.Active = True Or RibbonTabLayOutBudget.Active = True Then RibbonLF.ActiveTab = RibbonTabLayOutPlanning
                Case .TabPageBudget.Name
                    RibbonTabLayOutLogframe.Visible = False
                    RibbonTabLayOutPlanning.Visible = False
                    RibbonTabLayOutBudget.Visible = True
                    RibbonTabCollaborate.Visible = False

                    RibbonPanelLink.Visible = False
                    RibbonPanelReferences.Visible = True

                    RibbonComboBoxTextSelection.Visible = False

                    If RibbonTabLayOutLogframe.Active = True Or RibbonTabLayOutPlanning.Active = True Then RibbonLF.ActiveTab = RibbonTabLayOutBudget
                Case Else
                    RibbonTabLayOutLogframe.Visible = False
                    RibbonTabLayOutPlanning.Visible = False
                    RibbonTabLayOutBudget.Visible = False
                    RibbonTabCollaborate.Visible = False

                    RibbonPanelLink.Visible = False
                    RibbonPanelReferences.Visible = False

                    RibbonComboBoxTextSelection.Visible = False

                    If RibbonTabLayOutLogframe.Active = True Or RibbonTabLayOutPlanning.Active = True Or RibbonTabLayOutBudget.Active = True Then RibbonLF.ActiveTab = RibbonTabItems
            End Select
        End With
        RibbonLF.Invalidate()
        RibbonLF.ResumeLayout()
        RibbonLF.ResumeUpdating()
    End Sub
#End Region

#Region "Settings"
    Private Sub RibbonButtonSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSettings.Click
        Dim dialogSettings As New DialogSettings

        If dialogSettings.ShowDialog() = vbOK Then
            SetProjectLogicScheme()
        End If
    End Sub

    Private Sub RibbonButtonLanguage_DropDownShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLanguage.DropDownShowing
        Select Case UserLanguage
            Case "fr"
                RibbonButtonLanguageFrench.Checked = True
            Case "nl"
                RibbonButtonLanguageDutch.Checked = True
            Case Else
                RibbonButtonLanguageEnglish.Checked = True
        End Select
    End Sub

    Private Sub RibbonButtonLanguageEnglish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLanguageEnglish.Click
        ChangeLanguage("en-GB")
    End Sub

    Private Sub RibbonButtonLanguageFrench_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLanguageFrench.Click
        ChangeLanguage("fr")
    End Sub

    Private Sub RibbonButtonLanguageDutch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLanguageDutch.Click
        ChangeLanguage("nl")
    End Sub

    Private Sub ChangeLanguage(ByVal strLanguage As String)
        ChangeLanguage_Variables(strLanguage)

        Application.Restart()
    End Sub

    Private Sub RibbonButtonOnlineHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonOnlineHelp.Click
        Using objInternet As New classInternet
            objInternet.ExecuteFile(My.Settings.setHelpOnline, ERR_CannotOpenWebsite)
        End Using
    End Sub

    Private Sub RibbonButtonOnlineHelpLogframer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonOnlineHelpLogframer.Click
        Using objInternet As New classInternet
            objInternet.ExecuteFile(My.Settings.setHelpOnline, ERR_CannotOpenWebsite)
        End Using
    End Sub

    Private Sub RibbonButtonOnlineHelpFacilidev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonOnlineHelpFacilidev.Click
        Using objInternet As New classInternet
            objInternet.ExecuteFile(My.Settings.setHelpOnlineFacilidev, ERR_CannotOpenWebsite)
        End Using
    End Sub

    Private Sub RibbonButtonHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonHelp.Click
        Dim strHelpPath As String = GetHelpPath()

        If String.IsNullOrEmpty(strHelpPath) = False Then
            Help.ShowHelp(Me, strHelpPath, HelpNavigator.TableOfContents)
        End If
    End Sub
#End Region
End Class