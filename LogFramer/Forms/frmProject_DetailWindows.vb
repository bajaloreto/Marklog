Partial Public Class frmProject

#Region "Detailwindow logframe"
    Private Sub Clipboard_CloseButtonClicked() Handles Clipboard.CloseButtonClicked
        CType(MdiParent, frmParent).ClipboardShowHide(0)
    End Sub

    Public Function ShowDetailsLogframe() As String
        With Me.SplitContainerLogFrame
            If .Panel2Collapsed = True Then
                My.Settings.setShowDetailsLogframe = True

                If NewDetailLogframe IsNot Nothing Then .SplitterDistance = .Height - NewDetailLogframe.Height
                .Panel2Collapsed = False

                SetTypeOfDetailWindowLogframe()

                Return LANG_HideDetails
            Else
                My.Settings.setShowDetailsLogframe = False

                .Panel2Collapsed = True
                Return LANG_ShowDetails
            End If
        End With
    End Function

    Public Sub SetTypeOfDetailWindowLogframe()
        If My.Settings.setShowDetailsLogframe = True And dgvLogframe.CurrentCell IsNot Nothing Then
            Dim strColName As String
            Dim selGridRow As LogframeRow
            With dgvLogframe
                strColName = .Columns(.CurrentCell.ColumnIndex).Name
                selGridRow = .Grid(.CurrentRow.Index)
            End With

            If strColName.Contains("Struct") Then
                If selGridRow.Section = LogFrame.SectionTypes.GoalsSection Then
                    Me.NewDetailLogframe = New DetailGoal
                ElseIf selGridRow.Section = LogFrame.SectionTypes.PurposesSection Then
                    Dim objPurpose As Purpose = TryCast(selGridRow.Struct, Purpose)
                    If Not (objPurpose Is Nothing) Then
                        Me.NewDetailLogframe = New DetailPurpose(objPurpose)
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                ElseIf selGridRow.Section = LogFrame.SectionTypes.OutputsSection Then
                    Dim objOutput As Output = TryCast(selGridRow.Struct, Output)
                    If Not (objOutput Is Nothing) Then
                        Me.NewDetailLogframe = New DetailOutput(objOutput)
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                ElseIf selGridRow.Section = LogFrame.SectionTypes.ActivitiesSection Then
                    Dim objActivity As Activity = TryCast(selGridRow.Struct, Activity)
                    If Not (objActivity Is Nothing) Then
                        Me.NewDetailLogframe = New DetailActivity(objActivity)
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                Else
                    Me.NewDetailLogframe = Nothing
                End If
            ElseIf strColName.Contains("Ind") Then
                If dgvLogframe.IsResourceBudgetRow(selGridRow) = False Then
                    Dim objInd As Indicator = selGridRow.Indicator
                    If Not (objInd Is Nothing) Then
                        Me.NewDetailLogframe = New DetailIndicator(objInd)
                        CType(NewDetailLogframe, DetailIndicator).SelectText(Me.TextSelectionIndex)
                        AddHandler CType(NewDetailLogframe, DetailIndicator).CurrentDataGridViewChanged, AddressOf SetCurrentDataGridView
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                Else
                    Dim objRsc As Resource = selGridRow.Resource
                    If Not (objRsc Is Nothing) Then
                        Me.NewDetailLogframe = New DetailResource(objRsc)
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                End If
            ElseIf strColName.Contains("Ver") Then
                If dgvLogframe.IsResourceBudgetRow(selGridRow) = False Then
                    Dim objVer As VerificationSource = selGridRow.VerificationSource
                    If Not (objVer Is Nothing) Then
                        Me.NewDetailLogframe = New DetailVerificationSource(objVer)
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                Else
                    Dim objRsc As Resource = selGridRow.Resource
                    If Not (objRsc Is Nothing) Then
                        Me.NewDetailLogframe = New DetailResource(objRsc)
                    Else
                        Me.NewDetailLogframe = Nothing
                    End If
                End If
            ElseIf strColName.Contains("Asm") Then
                Dim objAsm As Assumption = selGridRow.Assumption
                If Not (objAsm Is Nothing) Then
                    Me.NewDetailLogframe = New DetailAssumption(objAsm)
                Else
                    Me.NewDetailLogframe = Nothing
                End If
            Else
                Me.NewDetailLogframe = Nothing
            End If
            ChangeDetailLogframe(Me.NewDetailLogframe)
        End If
    End Sub

    Private Sub ChangeDetailLogframe(ByVal NewDetail As UserControl)

        With Me.SplitContainerLogFrame
            If Me.CurrentDetailLogframe IsNot Nothing Then
                .Panel2.Controls.Remove(Me.CurrentDetailLogframe)
            End If
            If NewDetail IsNot Nothing Then
                NewDetail.Dock = DockStyle.Fill
                .Panel2.Controls.Add(NewDetail)
                NewDetail.Refresh()
            End If
        End With
        Me.CurrentDetailLogframe = NewDetail

    End Sub

    Public Sub ReloadDetailWindowLogframe()
        If My.Settings.setShowDetailsLogframe = True And Me.NewDetailLogframe IsNot Nothing Then
            If Me.NewDetailLogframe.GetType Is GetType(DetailGoal) Then
                CType(Me.NewDetailLogframe, DetailGoal).lvPartners.LoadItems()
            ElseIf Me.NewDetailLogframe.GetType Is GetType(DetailPurpose) Then
                CType(Me.NewDetailLogframe, DetailPurpose).lvDetailTargetGroups.LoadItems()
            ElseIf Me.NewDetailLogframe.GetType Is GetType(DetailOutput) Then
                CType(Me.NewDetailLogframe, DetailOutput).lvDetailKeyMoments.LoadItems()
            ElseIf Me.NewDetailLogframe.GetType Is GetType(DetailActivity) Then

            ElseIf Me.NewDetailLogframe.GetType Is GetType(DetailIndicator) Then

            ElseIf Me.NewDetailLogframe.GetType Is GetType(DetailVerificationSource) Then
                SetTypeOfDetailWindowLogframe()
            ElseIf Me.NewDetailLogframe.GetType Is GetType(DetailResource) Then
                CType(Me.NewDetailLogframe, DetailResource).dgvBudgetItemReferences.Reload()
            End If
        End If
    End Sub

    Private Sub SetCurrentDataGridView(ByVal sender As Object, ByVal e As CurrentDataGridViewChangedEventArgs)
        If Me.CurrentDataGridView IsNot e.DataGridView Then Me.CurrentDataGridView = e.DataGridView
    End Sub
#End Region

#Region "Detailwindow Planning"
    Public Function ShowDetailsPlanning() As String
        With Me.SplitContainerPlanning
            If .Panel2Collapsed = True Then
                My.Settings.setShowDetailsPlanning = True

                If NewDetailPlanning IsNot Nothing Then .SplitterDistance = .Height - NewDetailPlanning.Height
                .Panel2Collapsed = False

                SetTypeOfDetailWindowPlanning()

                Return LANG_HideDetails
            Else
                My.Settings.setShowDetailsPlanning = False

                .Panel2Collapsed = True
                Return LANG_ShowDetails
            End If
        End With
    End Function

    Public Sub SetTypeOfDetailWindowPlanning()
        If My.Settings.setShowDetailsPlanning = True And dgvPlanning.CurrentCell IsNot Nothing Then
            Dim selGridRow As PlanningGridRow = dgvPlanning.Grid(dgvPlanning.CurrentRow.Index)

            If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                If selGridRow.KeyMoment IsNot Nothing Then
                    Me.NewDetailPlanning = New DetailKeyMoment(selGridRow.KeyMoment)
                Else
                    Me.NewDetailPlanning = Nothing
                End If
            ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
                If selGridRow.Struct IsNot Nothing Then
                    Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
                    If selActivity IsNot Nothing Then
                        Me.NewDetailPlanning = New DetailActivity(selActivity)
                    Else
                        Me.NewDetailPlanning = Nothing
                    End If
                Else
                    Me.NewDetailPlanning = Nothing
                End If

            Else
                Me.NewDetailPlanning = Nothing
            End If
        Else
            Me.NewDetailPlanning = Nothing
        End If
        ChangeDetailPlanning(Me.NewDetailPlanning)
    End Sub

    Private Sub ChangeDetailPlanning(ByVal NewDetail As UserControl)

        With Me.SplitContainerPlanning
            If Me.CurrentDetailPlanning IsNot Nothing Then
                .Panel2.Controls.Remove(Me.CurrentDetailPlanning)
            End If
            If NewDetail IsNot Nothing Then
                .Panel2.Controls.Add(NewDetail)
                NewDetail.Refresh()
            End If
        End With
        Me.CurrentDetailPlanning = NewDetail
    End Sub
#End Region

#Region "Detailwindow Budget"
    Public Function GetSplitContainerBudget(ByVal intIndex As Integer) As SplitContainer
        Dim SplitContainerBudgetYear As SplitContainer = Nothing

        If TabControlBudget.TabPages.Count > intIndex Then
            SplitContainerBudgetYear = TabControlBudget.TabPages(intIndex).Controls("SplitContainerBudgetYear")
        End If

        Return SplitContainerBudgetYear
    End Function

    Public Function ShowDetailsBudget() As String
        Dim strShowDetails As String = String.Empty

        For i = 0 To TabControlBudget.TabPages.Count - 1
            Dim SplitContainerBudget As SplitContainer = GetSplitContainerBudget(i)

            With SplitContainerBudget
                If .Panel2Collapsed = True Then
                    My.Settings.setShowDetailsBudget = True

                    If NewDetailBudget IsNot Nothing Then .SplitterDistance = .Height - NewDetailBudget.Height
                    .Panel2Collapsed = False

                    SetTypeOfDetailWindowBudget(i)

                    strShowDetails = LANG_HideDetails
                Else
                    My.Settings.setShowDetailsBudget = False

                    .Panel2Collapsed = True
                    strShowDetails = LANG_ShowDetails
                End If
            End With
        Next

        Return strShowDetails
    End Function

    Public Sub SetTypeOfDetailWindowBudget(ByVal intBudgetYearIndex As Integer)
        NewDetailBudget = Nothing

        If My.Settings.setShowDetailsBudget = True Then
            Dim dgvBudgetYear As DataGridViewBudgetYear = GetDataGridViewBudgetYear(intBudgetYearIndex)

            If dgvBudgetYear IsNot Nothing AndAlso dgvBudgetYear.CurrentCell IsNot Nothing Then
                Dim selGridRow As BudgetGridRow = dgvBudgetYear.Grid(dgvBudgetYear.CurrentRow.Index)
                Dim selBudgetItem As BudgetItem = selGridRow.BudgetItem

                NewDetailBudget = New DetailBudgetItem(selGridRow.BudgetItem, dgvBudgetYear.GetBudgetItemReferences)
            End If
        End If
        ChangeDetailBudget(intBudgetYearIndex, Me.NewDetailBudget)
    End Sub

    Private Sub ChangeDetailBudget(ByVal intIndex As Integer, ByVal NewDetail As UserControl)
        Dim SplitContainerBudget As SplitContainer = GetSplitContainerBudget(intIndex)

        With SplitContainerBudget
            If SplitContainerBudget IsNot Nothing Then
                .Panel2.Controls.Clear()
                If NewDetail IsNot Nothing Then
                    .Panel2.Controls.Add(NewDetail)
                    NewDetail.Refresh()
                End If
            End If
        End With

        Me.CurrentDetailBudget = NewDetail
    End Sub

    Private Sub NewDetailBudget_BudgetUpdateParentTotalsNeeded(ByVal sender As Object, ByVal e As UpdateParentTotalsEventArgs) Handles NewDetailBudget.BudgetUpdateParentTotalsNeeded
        With Me.ProjectLogframe.Budget
            .UpdateParentTotals(e.ChildItem)
            .ChangeRatioBudgetHeader(e.ChildItem)
        End With

        Dim dgvBudgetYear As DataGridViewBudgetYear = GetCurrentDataGridViewBudgetYear()
        If dgvBudgetYear IsNot Nothing Then
            dgvBudgetYear.Reload()
        End If
    End Sub
#End Region

End Class
