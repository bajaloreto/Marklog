Partial Public Class frmProject
    Private Sub ProjectInitialise_dgvLogframe()
        With dgvLogframe
            .Logframe = Me.ProjectLogframe
            .ShowIndicatorColumn = My.Settings.setShowIndicatorColumn
            .ShowVerificationSourceColumn = My.Settings.setShowVerificationSourceColumn
            .ShowAssumptionColumn = My.Settings.setShowAssumptionColumn
            .ShowResourcesBudget = My.Settings.setShowResourcesBudget

            .InitialiseColumns()
        End With
    End Sub

    Private Sub dgvLogframe_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLogframe.Enter
        If Me.CurrentDataGridView IsNot dgvLogframe Then Me.CurrentDataGridView = dgvLogframe
    End Sub

    Private Sub dgvLogframe_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvLogframe.CellClick
        If Me.CurrentDataGridView IsNot dgvLogframe Then Me.CurrentDataGridView = dgvLogframe
        If Me.TextSelectionIndex <> TextSelectionValues.SelectNothing Then Me.TextSelectionIndex = TextSelectionValues.SelectNothing
    End Sub

    Private Sub dgvLogframe_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLogframe.MouseDown
        Dim hit As DataGridView.HitTestInfo = dgvLogframe.HitTest(e.X, e.Y)
        If hit.ColumnIndex > 0 And hit.RowIndex > 0 Then
            Dim selGridRow As LogframeRow = dgvLogframe.Grid(hit.RowIndex)
            Dim strStructs As String = ""
            Dim strMsg As String

            Select Case selGridRow.Section
                Case LogFrame.SectionTypes.GoalsSection
                    strStructs = My.Settings.setStruct1.ToLower
                Case LogFrame.SectionTypes.PurposesSection
                    strStructs = My.Settings.setStruct2.ToLower
                Case LogFrame.SectionTypes.OutputsSection
                    strStructs = My.Settings.setStruct3.ToLower
                Case LogFrame.SectionTypes.ActivitiesSection
                    strStructs = My.Settings.setStruct4.ToLower
            End Select
            If e.Button = MouseButtons.Left Then
                strMsg = String.Format(LANG_DragToSelectMultiple, strStructs)

                frmParent.StatusLabelGeneral.Text = strMsg
            ElseIf e.Button = MouseButtons.Right Then
                With frmParent
                    If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
                End With
                strMsg = String.Format(LANG_DragToMoveSelected, strStructs)

                frmParent.StatusLabelGeneral.Text = strMsg
            End If

        End If
    End Sub

    Private Sub dgvLogframe_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLogframe.MouseUp
        frmParent.StatusLabelGeneral.Text = String.Empty

        If dgvLogframe.IsCurrentCellInEditMode = False Then
            With frmParent
                If .RibbonTabItems.Active = False Then .RibbonLF.ActiveTab = .RibbonTabItems
            End With
        End If
        'Show details window
        SetTypeOfDetailWindowLogframe()
    End Sub

    Private Sub dgvLogframe_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvLogframe.EditingControlShowing
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
        End With
    End Sub

    Private Sub dgvLogframe_Reloaded() Handles dgvLogframe.Reloaded
        If dgvLogframe.CurrentCell Is Nothing Then
            dgvLogframe.CurrentCell = dgvLogframe(1, 1)
            dgvLogframe(1, 1).Selected = True
        End If
        SetTypeOfDetailWindowLogframe()
    End Sub

    Private Sub dgvLogframe_LogframeObjectSelected(ByVal sender As Object, ByVal e As LogframeObjectSelectedEventArgs) Handles dgvLogframe.LogframeObjectSelected
        Dim objLogframeObject As LogframeObject = e.LogframeObject

        With frmParent
            If objLogframeObject Is Nothing Then
                .RibbonButtonInsertChild.Enabled = False
                .RibbonButtonInsertParent.Enabled = False
                .RibbonButtonLevelUp.Enabled = False
                .RibbonButtonLevelDown.Enabled = False
                .RibbonButtonSectionUp.Enabled = False
                .RibbonButtonSectionDown.Enabled = False
            Else
                Select Case objLogframeObject.GetType
                    Case GetType(Activity)
                        Dim selActivity As Activity = DirectCast(objLogframeObject, Activity)
                        Dim ParentOutput As Output
                        Dim ParentActivity As Activity
                        Dim ParentActivities As Activities = Nothing

                        .RibbonButtonInsertChild.Enabled = True
                        .RibbonButtonInsertParent.Enabled = True

                        If selActivity.ParentOutputGuid = Guid.Empty Then
                            .RibbonButtonLevelUp.Enabled = True
                            '.RibbonButtonInsertParent.Enabled = True
                        Else
                            .RibbonButtonLevelUp.Enabled = False
                            ' .RibbonButtonInsertParent.Enabled = False
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

                        .RibbonButtonSectionUp.Enabled = True
                        .RibbonButtonSectionDown.Enabled = False
                    Case GetType(Indicator)
                        Dim selIndicator As Indicator = DirectCast(objLogframeObject, Indicator)
                        Dim ParentStruct As Struct
                        Dim ParentIndicator As Indicator
                        Dim ParentIndicators As Indicators = Nothing

                        .RibbonButtonInsertChild.Enabled = True
                        .RibbonButtonInsertParent.Enabled = True

                        If selIndicator.ParentStructGuid = Guid.Empty Then
                            .RibbonButtonLevelUp.Enabled = True
                        Else
                            .RibbonButtonLevelUp.Enabled = False
                        End If

                        If selIndicator.ParentIndicatorGuid <> Guid.Empty Then
                            ParentIndicator = CurrentLogFrame.GetIndicatorByGuid(selIndicator.ParentIndicatorGuid)
                            If ParentIndicator IsNot Nothing Then ParentIndicators = ParentIndicator.Indicators
                        ElseIf selIndicator.ParentStructGuid <> Guid.Empty Then
                            ParentStruct = CurrentLogFrame.GetStructByGuid(selIndicator.ParentStructGuid)
                            If ParentStruct IsNot Nothing Then ParentIndicators = ParentStruct.Indicators
                        End If

                        If ParentIndicators IsNot Nothing Then
                            If ParentIndicators.IndexOf(selIndicator) > 0 Then
                                .RibbonButtonLevelDown.Enabled = True
                            Else
                                .RibbonButtonLevelDown.Enabled = False
                            End If
                        End If

                        .RibbonButtonSectionUp.Enabled = False
                        .RibbonButtonSectionDown.Enabled = False
                    Case Else
                        .RibbonButtonInsertChild.Enabled = False
                        .RibbonButtonInsertParent.Enabled = False
                        .RibbonButtonLevelUp.Enabled = False
                        .RibbonButtonLevelDown.Enabled = False

                        Dim selStruct As Struct = TryCast(objLogframeObject, Struct)

                        If selStruct Is Nothing Then
                            .RibbonButtonSectionUp.Enabled = False
                            .RibbonButtonSectionDown.Enabled = False
                        Else
                            .RibbonButtonSectionUp.Enabled = True
                            .RibbonButtonSectionDown.Enabled = True
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub dgvLogframe_ShowAssumptionColumnChanged() Handles dgvLogframe.ShowAssumptionColumnChanged
        Dim objParentForm As frmParent = MdiParent
        objParentForm.ShowAssumptions()
    End Sub

    Private Sub dgvLogframe_ShowIndicatorColumnChanged() Handles dgvLogframe.ShowIndicatorColumnChanged
        Dim objParentForm As frmParent = MdiParent
        objParentForm.ShowIndicators()
    End Sub

    Private Sub dgvLogframe_ShowVerificationSourceColumnChanged() Handles dgvLogframe.ShowVerificationSourceColumnChanged
        Dim objParentForm As frmParent = MdiParent
        objParentForm.ShowVerificationSources()
    End Sub

    Private Sub dgvLogframe_ShowResourcesBudgetChanged() Handles dgvLogframe.ShowResourcesBudgetChanged
        Dim objParentForm As frmParent = MdiParent
        objParentForm.ShowResourcesBudget()
    End Sub

    Private Sub dgvLogframe_NoTextSelected() Handles dgvLogframe.NoTextSelected
        Dim objParentForm As frmParent = MdiParent
        objParentForm.SetTextSelectionToNothing()
    End Sub
End Class
