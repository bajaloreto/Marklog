Partial Public Class frmProject
    Private Sub ProjectInitialise_dgvPlanning()
        With dgvPlanning
            .ShowDatesColumns = My.Settings.setShowDatesColumns
            .ShowActivityLinks = My.Settings.setPlanningActivityLinks
            .ShowKeyMomentLinks = My.Settings.setPlanningKeyMomentLinks
            .ElementsView = My.Settings.setPlanningElementsView

            .InitialiseColumns()
        End With
    End Sub

    Private Sub dgvPlanning_DragLinkReleased() Handles dgvPlanning.DragLinkReleased
        CType(MdiParent, frmParent).RibbonButtonLink.Checked = False
    End Sub

    Private Sub dgvPlanning_UnlinkReleased() Handles dgvPlanning.UnlinkReleased
        CType(MdiParent, frmParent).RibbonButtonUnlink.Checked = False
    End Sub

    Private Sub dgvPlanning_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPlanning.Enter
        If Me.CurrentDataGridView IsNot dgvPlanning Then Me.CurrentDataGridView = dgvPlanning
    End Sub

    Private Sub dgvPlanning_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPlanning.CellClick
        If Me.CurrentDataGridView IsNot dgvPlanning Then Me.CurrentDataGridView = dgvPlanning
        If Me.TextSelectionIndex <> TextSelectionValues.SelectNothing Then Me.TextSelectionIndex = TextSelectionValues.SelectNothing
    End Sub

    Private Sub dgvPlanning_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvPlanning.MouseDown
        Dim hit As DataGridView.HitTestInfo = dgvPlanning.HitTest(e.X, e.Y)

        If hit.ColumnIndex > 0 And hit.RowIndex > 0 Then
            Dim selGridRow As PlanningGridRow = dgvPlanning.Grid(hit.RowIndex)
            Dim strObjects As String = String.Empty
            Dim strMsg As String

            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.KeyMoment
                    strObjects = LANG_KeyMoments
                Case PlanningGridRow.RowTypes.Activity
                    strObjects = My.Settings.setStruct4.ToLower
            End Select
            If e.Button = MouseButtons.Left Then
                strMsg = String.Format(LANG_DragToSelectMultiple, strObjects)

                frmParent.StatusLabelGeneral.Text = strMsg
            ElseIf e.Button = MouseButtons.Right Then
                With frmParent
                    If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
                End With
                strMsg = String.Format(LANG_DragToMoveSelected, strObjects)

                frmParent.StatusLabelGeneral.Text = strMsg
            End If

        End If
    End Sub

    Private Sub dgvPlanning_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvPlanning.MouseUp
        frmParent.StatusLabelGeneral.Text = String.Empty

        If dgvPlanning.IsCurrentCellInEditMode = False Then
            With frmParent
                If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
            End With
        End If

        'Show details window
        SetTypeOfDetailWindowPlanning()
    End Sub

    Private Sub dgvPlanning_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvPlanning.EditingControlShowing
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
        End With
    End Sub

    Private Sub dgvPlanning_Reloaded() Handles dgvPlanning.Reloaded
        If dgvPlanning.CurrentCell Is Nothing OrElse dgvPlanning.CurrentCellAddress = New Point(0, 0) Then
            dgvPlanning.CurrentCell = dgvPlanning(1, 0)
            dgvPlanning(1, 0).Selected = True
        End If

        SetTypeOfDetailWindowPlanning()
    End Sub

    Private Sub dgvPlanning_PlanningObjectSelected(ByVal sender As Object, ByVal e As PlanningObjectSelectedEventArgs) Handles dgvPlanning.PlanningObjectSelected
        Dim objPlanningObject As Object = e.PlanningObject

        With frmParent
            If objPlanningObject Is Nothing Then
                .RibbonButtonInsertChild.Enabled = False
                .RibbonButtonInsertParent.Enabled = False
                .RibbonButtonLevelUp.Enabled = False
                .RibbonButtonLevelDown.Enabled = False
                .RibbonButtonSectionUp.Enabled = False
                .RibbonButtonSectionDown.Enabled = False
            Else
                Select Case objPlanningObject.GetType
                    Case GetType(Activity)
                        Dim selActivity As Activity = DirectCast(objPlanningObject, Activity)
                        Dim ParentOutput As Output
                        Dim ParentActivity As Activity
                        Dim ParentActivities As Activities = Nothing

                        .RibbonButtonInsertChild.Enabled = True
                        .RibbonButtonMoveUp.Enabled = True
                        .RibbonButtonMoveDown.Enabled = True

                        If selActivity.ParentOutputGuid = Guid.Empty Then
                            .RibbonButtonLevelUp.Enabled = True
                            .RibbonButtonInsertParent.Enabled = True
                        Else

                            .RibbonButtonLevelUp.Enabled = False
                            .RibbonButtonInsertParent.Enabled = False
                        End If

                        If selActivity.ParentActivityGuid <> Guid.Empty Then
                            ParentActivity = CurrentLogFrame.GetActivityByGuid(selActivity.ParentActivityGuid)
                            If ParentActivity IsNot Nothing Then ParentActivities = ParentActivity.Activities
                        ElseIf selActivity.ParentOutputGuid <> Guid.Empty Then
                            ParentOutput = CurrentLogFrame.GetOutputByGuid(selActivity.ParentOutputGuid)
                            If ParentOutput IsNot Nothing Then ParentActivities = ParentOutput.Activities
                        End If

                        If ParentActivities IsNot Nothing Then
                            If ParentActivities.IndexOf(selActivity) > 0 Then
                                .RibbonButtonLevelDown.Enabled = True
                            Else
                                .RibbonButtonLevelDown.Enabled = False
                            End If
                        End If
                    Case Else
                        .RibbonButtonMoveUp.Enabled = False
                        .RibbonButtonMoveDown.Enabled = False
                        .RibbonButtonInsertChild.Enabled = False
                        .RibbonButtonInsertParent.Enabled = False
                        .RibbonButtonLevelUp.Enabled = False
                        .RibbonButtonLevelDown.Enabled = False

                End Select
            End If
        End With
    End Sub

    Private Sub dgvPlanning_NoTextSelected() Handles dgvPlanning.NoTextSelected
        Dim objParentForm As frmParent = MdiParent
        objParentForm.SetTextSelectionToNothing()
    End Sub
End Class
