Partial Public Class frmParent
#Region "Cut - copy - paste"
    Private Sub RibbonTabItems_ActiveChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonTabItems.ActiveChanged
        SetClipboardType(DataGridViewClipboard.ContentTypes.Items)
    End Sub

    Private Sub RibbonButtonCutItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonCutItems.Click
        CutItem()
    End Sub

    Private Sub RibbonButtonCopyItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonCopyItems.Click
        CopyItem()
    End Sub

    Private Sub RibbonButtonPasteItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteItems.Click
        PasteItem()
    End Sub

    Private Sub RibbonButtonPasteNoVertLogic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteNoVertLogic.Click
        PasteItem(Nothing, ClipboardItems.PasteOptions.PasteNoVert)
    End Sub

    Private Sub RibbonButtonPasteNoDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteNoDetail.Click
        PasteItem(Nothing, ClipboardItems.PasteOptions.PasteNoDetails)
    End Sub

    Private Sub RibbonButtonPasteNoDependencies_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteNoDependencies.Click
        PasteItem(Nothing, ClipboardItems.PasteOptions.PasteNoDependencies)
    End Sub

    Private Sub RibbonPanelClipboardItems_ButtonMoreClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonPanelClipboardItems.ButtonMoreClick
        ClipboardShowHide(DataGridViewClipboard.ContentTypes.Items)
    End Sub

    Public Sub CutItem()
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.CutItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvPlanning"
                        .dgvPlanning.CutItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).CutItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).CutItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).CutItems(False)

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).CutItems(False)
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).CutItems(False)
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).CutItems(False)

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).CutItems(False)
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).CutItems(False)
                    Case "lvAddresses", "lvBudgetItemReferences", "lvContacts", "lvEmails", "lvDetailKeyMoments", "lvDetailProcesses", "lvDetailTargetGroups", _
                        "lvPartners", "lvSubIndicators", "lvTargetDeadlines", "lvTargets", "lvTargetGroups", "lvTelephoneNumbers", "lvWebsites"
                        CType(CurrentControl, ListViewSortable).CutItems()
                End Select
            End With
        End If
    End Sub

    Public Sub CopyItem()
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.CopyItems()
                    Case "dgvPlanning"
                        .dgvPlanning.CopyItems()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).CopyItems()
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).CopyItems()
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).CopyItems()

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).CopyItems()
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).CopyItems()
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).CopyItems()

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).CopyItems()
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).CopyItems()
                    Case "lvAddresses", "lvBudgetItemReferences", "lvContacts", "lvEmails", "lvDetailKeyMoments", "lvDetailProcesses", "lvDetailTargetGroups", _
                        "lvPartners", "lvSubIndicators", "lvTargetDeadlines", "lvTargets", "lvTargetGroups", "lvTelephoneNumbers", "lvWebsites"
                        CType(CurrentControl, ListViewSortable).CopyItems()
                End Select
            End With
        End If
    End Sub

    Public Sub PasteItem(Optional ByVal datCopyGroup As Date = Nothing, Optional ByVal intPasteOption As Integer = 0)
        Dim objPasteItems As ClipboardItems = ItemClipboard.GetCopyGroup(datCopyGroup)

        If objPasteItems IsNot Nothing AndAlso objPasteItems.Count > 0 Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.PasteItems(objPasteItems, intPasteOption)
                    Case "dgvPlanning"
                        .dgvPlanning.PasteItems(objPasteItems, intPasteOption)
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).PasteItems(objPasteItems, intPasteOption)
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).PasteItems(objPasteItems, intPasteOption)
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).PasteItems(objPasteItems, intPasteOption)

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).PasteItems(objPasteItems, intPasteOption)
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).PasteItems(objPasteItems, intPasteOption)
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).PasteItems(objPasteItems, intPasteOption)

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).PasteItems(objPasteItems, intPasteOption)
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).PasteItems(objPasteItems, intPasteOption)
                    Case "lvAddresses", "lvBudgetItemReferences", "lvContacts", "lvEmails", "lvDetailKeyMoments", "lvDetailProcesses", "lvDetailTargetGroups", _
                        "lvPartners", "lvSubIndicators", "lvTargetDeadlines", "lvTargets", "lvTargetGroups", "lvTelephoneNumbers", "lvWebsites"
                        CType(CurrentControl, ListViewSortable).PasteItems(objPasteItems)
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonReferenceCopyPreviousYear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonReferenceCopyPreviousYear.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvBudgetYear"
                        .CopyReferenceValues(-1)
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonReferenceCopyNextYear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonReferenceCopyNextYear.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvBudgetYear"
                        .CopyReferenceValues(1)
                End Select
            End With
        End If
    End Sub
#End Region

#Region "New - insert - edit - remove"
    Private Sub RibbonButtonNewItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonNewItem.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.AddNewItem()
                    Case "dgvPlanning"
                        .dgvPlanning.AddNewItem()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).AddNewItem()
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).AddNewItem()
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).AddNewItem()

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).AddNewItem()
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).AddNewItem()
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).AddNewItem()

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).AddNewItem()
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).AddNewItem()
                    Case "lvBudgetItemReferences", "lvContacts", "lvDetailKeyMoments", "lvDetailProcesses", "lvDetailTargetGroups", "lvPartners", "lvSubIndicators", _
                        "lvTargetDeadlines", "lvTargets", "lvTargetGroups"
                        CType(CurrentControl, ListViewSortable).NewItem()
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonInsertItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonInsertItem.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.InsertItem()
                    Case "dgvPlanning"
                        .dgvPlanning.InsertItem()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).InsertItem()
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).InsertItem()
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).InsertItem()

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).InsertItem()
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).InsertItem()
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).InsertItem()

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).InsertItem()
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).InsertItem()
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonInsertParent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonInsertParent.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.InsertParentItem()
                    Case "dgvPlanning"
                        .dgvPlanning.InsertParentItem()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).InsertParentItem()
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonInsertChild_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonInsertChild.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.InsertChildItem()
                    Case "dgvPlanning"
                        .dgvPlanning.InsertChildItem()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).InsertChildItem()
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonRemoveItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonRemoveItem.Click
        RemoveItem()
    End Sub

    Public Sub RemoveItem()
        If CurrentControl IsNot Nothing Then

            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.RemoveItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvPlanning"
                        .dgvPlanning.RemoveItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).RemoveItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).RemoveItems(My.Settings.setWarnLinkedObjectDelete)
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).RemoveItems(False)

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).RemoveItems(False)
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).RemoveItems(False)
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).RemoveItems(False)

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).RemoveItems(False)
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).RemoveItems(False)
                    Case "lvBudgetItemReferences", "lvContacts", "lvDetailKeyMoments", "lvDetailProcesses", "lvDetailTargetGroups", "lvPartners", "lvSubIndicators", _
                        "lvTargetDeadlines", "lvTargets", "lvTargetGroups"
                        CType(CurrentControl, ListViewSortable).RemoveItem()
                    Case Else
                        Select Case CurrentControl.GetType
                            Case GetType(TextBoxLF)
                                Dim selTextBox As TextBoxLF = CType(CurrentControl, TextBoxLF)

                                If selTextBox.SelectedText.Length > 0 Then
                                    selTextBox.SelectedText = selTextBox.SelectedText.Remove(0)
                                Else
                                    selTextBox.Clear()
                                End If

                                UndoRedo.TextRemoved(selTextBox.Text)
                        End Select
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonEditItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonEditItem.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        With .dgvLogframe
                            If .CurrentCell IsNot Nothing Then .BeginEdit(False)
                        End With
                    Case "dgvPlanning"
                        With .dgvPlanning
                            If .CurrentCell IsNot Nothing Then .BeginEdit(False)
                        End With
                    Case "dgvBudgetYear"
                        Dim dgvBudgetYear As DataGridViewBudgetYear = TryCast(CurrentControl, DataGridViewBudgetYear)
                        If dgvBudgetYear IsNot Nothing Then
                            If dgvBudgetYear.CurrentCell IsNot Nothing Then dgvBudgetYear.BeginEdit(False)
                        End If
                    Case "dgvBudgetItemReferences"
                        Dim dgvBudgetItemReferences As DataGridViewBudgetItemReferences = TryCast(CurrentControl, DataGridViewBudgetItemReferences)
                        If dgvBudgetItemReferences IsNot Nothing Then
                            If dgvBudgetItemReferences.CurrentCell IsNot Nothing Then dgvBudgetItemReferences.BeginEdit(False)
                        End If
                    Case "dgvResponseClasses"
                        Dim dgvResponseClasses As DataGridViewResponseClasses = TryCast(CurrentControl, DataGridViewResponseClasses)
                        If dgvResponseClasses IsNot Nothing Then
                            If dgvResponseClasses.CurrentCell IsNot Nothing Then dgvResponseClasses.BeginEdit(False)
                        End If

                    Case "dgvStatementsFormula"
                        Dim dgvStatements As DataGridViewStatementsFormula = TryCast(CurrentControl, DataGridViewStatementsFormula)
                        If dgvStatements IsNot Nothing Then
                            If dgvStatements.CurrentCell IsNot Nothing Then dgvStatements.BeginEdit(False)
                        End If
                    Case "dgvStatementsMaxDiff"
                        Dim dgvStatements As DataGridViewStatementsMaxDiff = TryCast(CurrentControl, DataGridViewStatementsMaxDiff)
                        If dgvStatements IsNot Nothing Then
                            If dgvStatements.CurrentCell IsNot Nothing Then dgvStatements.BeginEdit(False)
                        End If
                    Case "dgvStatementsScales"
                        Dim dgvStatements As DataGridViewStatementsScales = TryCast(CurrentControl, DataGridViewStatementsScales)
                        If dgvStatements IsNot Nothing Then
                            If dgvStatements.CurrentCell IsNot Nothing Then dgvStatements.BeginEdit(False)
                        End If

                    Case "dgvTargetsScaleLikert"
                        Dim dgvTargets As DataGridViewTargetsScaleLikert = TryCast(CurrentControl, DataGridViewTargetsScaleLikert)
                        If dgvTargets IsNot Nothing Then
                            If dgvTargets.CurrentCell IsNot Nothing Then dgvTargets.BeginEdit(False)
                        End If
                    Case "dgvTargetsSemanticDiff"
                        Dim dgvTargets As DataGridViewTargetsSemanticDiff = TryCast(CurrentControl, DataGridViewTargetsSemanticDiff)
                        If dgvTargets IsNot Nothing Then
                            If dgvTargets.CurrentCell IsNot Nothing Then dgvTargets.BeginEdit(False)
                        End If
                    Case "dgvTargets"
                        Select Case CurrentControl.GetType
                            Case GetType(DataGridViewTargetsValues)
                                Dim dgvTargets As DataGridViewTargetsValues = TryCast(CurrentControl, DataGridViewTargetsValues)
                                If dgvTargets IsNot Nothing Then dgvTargets.Edit()
                            Case GetType(DataGridViewTargetsClasses)
                                Dim dgvTargets As DataGridViewTargetsClasses = TryCast(CurrentControl, DataGridViewTargetsClasses)
                                If dgvTargets IsNot Nothing Then dgvTargets.Edit()
                            Case GetType(DataGridViewTargetsClasses)
                                Dim dgvTargets As DataGridViewTargetsClasses = TryCast(CurrentControl, DataGridViewTargetsClasses)
                                If dgvTargets IsNot Nothing Then dgvTargets.Edit()
                        End Select

                    Case "lvBudgetItemReferences", "lvContacts", "lvDetailKeyMoments", "lvDetailProcesses", "lvDetailTargetGroups", "lvPartners", "lvSubIndicators", _
                        "lvTargetDeadlines", "lvTargets", "lvTargetGroups"
                        CType(CurrentControl, ListViewSortable).EditItem()
                    Case Else
                        Select Case CurrentControl.GetType
                            Case GetType(TextBoxLF)
                                CType(CurrentControl, TextBoxLF).SelectAll()

                        End Select
                End Select
            End With
        End If
    End Sub
#End Region

#Region "Move"
    Private Sub RibbonButtonMoveUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonMoveUp.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.MoveItem(-1)
                    Case "dgvPlanning"
                        .dgvPlanning.MoveItem(-1)
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).MoveItem(-1)
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).MoveItem(-1)
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).MoveItem(-1)

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).MoveItem(-1)
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).MoveItem(-1)
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).MoveItem(-1)

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).MoveItem(-1)
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).MoveItem(-1)
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonMoveDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonMoveDown.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.MoveItem(1)
                    Case "dgvPlanning"
                        .dgvPlanning.MoveItem(1)
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).MoveItem(1)
                    Case "dgvBudgetItemReferences"
                        CType(CurrentControl, DataGridViewBudgetItemReferences).MoveItem(1)
                    Case "dgvResponseClasses"
                        CType(CurrentControl, DataGridViewResponseClasses).MoveItem(1)

                    Case "dgvStatementsFormula"
                        CType(CurrentControl, DataGridViewStatementsFormula).MoveItem(1)
                    Case "dgvStatementsMaxDiff"
                        CType(CurrentControl, DataGridViewStatementsMaxDiff).MoveItem(1)
                    Case "dgvStatementsScales"
                        CType(CurrentControl, DataGridViewStatementsScales).MoveItem(1)

                    Case "dgvTargetsScaleLikert"
                        CType(CurrentControl, DataGridViewTargetsScaleLikert).MoveItem(1)
                    Case "dgvTargetsSemanticDiff"
                        CType(CurrentControl, DataGridViewTargetsSemanticDiff).MoveItem(1)
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonSectionUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSectionUp.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.ChangeSection(-1)
                    Case "dgvPlanning"
                        .dgvPlanning.ChangeSection(-1)
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonSectionDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSectionDown.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.ChangeSection(1)
                    Case "dgvPlanning"
                        .dgvPlanning.ChangeSection(1)
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonLevelUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLevelUp.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.LevelUp()
                    Case "dgvPlanning"
                        .dgvPlanning.LevelUp()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).LevelUp()
                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonLevelDown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLevelDown.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvLogframe"
                        .dgvLogframe.LevelDown()
                    Case "dgvPlanning"
                        .dgvPlanning.LevelDown()
                    Case "dgvBudgetYear"
                        CType(CurrentControl, DataGridViewBudgetYear).LevelDown()
                End Select
            End With
        End If
    End Sub
#End Region

#Region "Links"
    Private Sub RibbonButtonLink_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLink.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvPlanning"
                        .dgvPlanning.DragLink = RibbonButtonLink.Checked

                End Select
            End With
        End If
    End Sub

    Private Sub RibbonButtonUnlink_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonUnlink.Click
        If CurrentControl IsNot Nothing Then
            With CurrentProjectForm
                Select Case CurrentControl.Name
                    Case "dgvPlanning"
                        .dgvPlanning.Unlink = RibbonButtonUnlink.Checked

                End Select
            End With
        End If
    End Sub
#End Region

End Class
