Partial Public Class classUndoRedo
#Region "Reload"
    Private Sub Reload(ByVal selUndoListItem As UndoListItem)
        With CurrentProjectForm
            Select Case .TabControlProject.SelectedIndex
                Case 0 'Project information
                    Reload_ProjectInfo(selUndoListItem)
                Case 1 'Logframe
                    Reload_Logframe(selUndoListItem)
                Case 2 'Planning
                    Reload_Planning(selUndoListItem)
                Case 3 'Budget
                    Reload_Budget(selUndoListItem)
            End Select
        End With
    End Sub

    Private Sub ReadValueOfControl(ByVal objPane As Object, ByVal strPropertyName As String)
        Dim selControl As Control

        selControl = GetControlByPropertyName(objPane, strPropertyName)

        If selControl IsNot Nothing Then
            If selControl.DataBindings.Count > 0 Then
                selControl.DataBindings(0).ReadValue()
            End If
        End If
    End Sub

    Public Function GetControlByPropertyName(ByVal objPane As Object, ByVal strPropertyName As String) As Control
        Dim selControl As Control = Nothing
        Dim boolFound As Boolean

        If objPane Is Nothing Then Return Nothing
        If strPropertyName.Contains(".") Then strPropertyName = GetNestedPropertyName(strPropertyName)

        For Each selControl In objPane.Controls
            Select Case selControl.GetType
                Case GetType(TabControl)
                    selControl.Focus()

                    selControl = GetControlByPropertyName_TabControl(selControl, strPropertyName)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                Case GetType(SplitContainer)
                    Dim objSplitContainer As SplitContainer = CType(selControl, SplitContainer)
                    objSplitContainer.Focus()
                    selControl = GetControlByPropertyName(objSplitContainer.Panel1, strPropertyName)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                    selControl = GetControlByPropertyName(objSplitContainer.Panel2, strPropertyName)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                Case GetType(AssumptionAssumptions), GetType(AssumptionDependencies), GetType(AssumptionRisks), _
                    GetType(BudgetItemRatio), _
                    GetType(IndicatorFormula), GetType(IndicatorFrequencyLikert), GetType(IndicatorMaxDiff), GetType(IndicatorMixedSubIndicators), GetType(IndicatorMultipleOption), _
                    GetType(IndicatorOpenEnded), GetType(IndicatorRanking), GetType(IndicatorRatio), GetType(IndicatorScales), GetType(IndicatorTargetsLikertScale), _
                    GetType(IndicatorTargetsSemanticDiff), GetType(IndicatorValues), GetType(IndicatorYesNo), _
                    GetType(GroupBox), GetType(TableLayoutPanel), GetType(Panel)

                    selControl.Focus()
                    selControl = GetControlByPropertyName(selControl, strPropertyName)
                    If selControl IsNot Nothing Then
                        boolFound = True
                        Exit For
                    End If
                Case Else
                    If selControl.GetType IsNot GetType(Label) AndAlso GetControlByPropertyName_VerifyNameAndType(selControl.Name, strPropertyName) Then
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

    Public Function GetNestedPropertyName(ByVal strPropertyName As String) As String
        Dim strElements() As String = strPropertyName.Split(".")
        Dim strElement As String = strElements(strElements.Count - 1)

        Return strElement
    End Function

    Private Function GetControlByPropertyName_VerifyNameAndType(ByVal strControlName As String, ByVal strPropertyName As String) As Boolean
        Dim intIndex As Integer
        Dim intExpectedLength As Integer
        Dim strType As String

        If strControlName.Contains(strPropertyName) Then
            intIndex = strControlName.IndexOf(strPropertyName)
            intExpectedLength = intIndex + strPropertyName.Length

            If strControlName.Length > intExpectedLength + 1 Then '+1: allow for indexed controls of same property; for instance ntbNrDecimals1 and ntbNrDecimals2
                Return False
            Else
                strType = strControlName.Substring(0, intIndex)

                Select Case strType
                    Case "tb", "dtb", "ntb", "rtb", "cmb", "dtp"
                        Return True
                End Select
            End If
        End If

        Return False
    End Function

    Private Function GetControlByPropertyName_TabControl(ByVal selTabControl As TabControl, ByVal strPropertyName As String) As Control
        Dim selControl As Control = Nothing
        Dim boolFound As Boolean

        For Each selPage As TabPage In selTabControl.TabPages
            selPage.Focus()

            selControl = GetControlByPropertyName(selPage, strPropertyName)
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

#Region "Reload Project information"
    Private Sub Reload_ProjectInfo(ByVal selUndoListItem As UndoListItem)
        With CurrentProjectForm.ProjectInfo
            Select Case selUndoListItem.Item.GetType
                Case GetType(Address)
                    .lvPartners.LoadItems()
                    .lvContacts.LoadItems()
                Case GetType(Contact)
                    .lvContacts.LoadItems()
                Case GetType(Email)
                    .lvPartners.LoadItems()
                    .lvContacts.LoadItems()
                Case GetType(LogFrame)
                    Dim selControl As Control
                    Dim strPropertyName As String = selUndoListItem.PropertyName

                    selControl = GetControlByPropertyName(CurrentProjectForm.ProjectInfo, strPropertyName)

                    If selControl IsNot Nothing Then
                        If selControl.DataBindings.Count > 0 Then
                            selControl.DataBindings(0).ReadValue()
                        Else
                            If strPropertyName = "StartDate" Or strPropertyName = "EndDate" Then .ModifyEndDate()
                        End If
                    End If
                Case GetType(ProjectPartner)
                    .lvPartners.LoadItems()
                Case GetType(TargetGroup), GetType(TargetGroupInformation)
                    .lvTargetGroups.LoadTargetGroups(True)
                Case GetType(TelephoneNumber)
                    .lvPartners.LoadItems()
                    .lvContacts.LoadItems()
                Case GetType(Website)
                    .lvPartners.LoadItems()
                    .lvContacts.LoadItems()
            End Select
        End With
    End Sub
#End Region

#Region "Reload logframe"
    Private Sub Reload_Logframe(ByVal selUndoListItem As UndoListItem)

        If String.IsNullOrEmpty(selUndoListItem.PropertyName) Then
            Reload_Logframe_ListViews(selUndoListItem)
        Else
            With CurrentProjectForm
                Select Case selUndoListItem.PropertyName
                    Case "RTF"
                        .dgvLogframe.Reload()
                    Case Else
                        Select Case selUndoListItem.Item.GetType
                            Case GetType(LogFrame)
                                Reload_Logframe_Logframe(selUndoListItem)
                            Case GetType(ProjectPartner), GetType(Organisation), GetType(Address), GetType(Email), GetType(TelephoneNumber), GetType(Website)
                                If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailGoal) Then
                                    CType(.CurrentDetailLogframe, DetailGoal).lvPartners.LoadItems()
                                End If
                            Case GetType(TargetGroup)
                                If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailPurpose) Then
                                    CType(.CurrentDetailLogframe, DetailPurpose).lvDetailTargetGroups.LoadTargetGroups(False)
                                End If
                            Case GetType(KeyMoment)
                                If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailOutput) Then
                                    CType(.CurrentDetailLogframe, DetailOutput).lvDetailKeyMoments.LoadItems()
                                End If
                            Case GetType(ActivityDetail)
                                Reload_Logframe_ActivityDetail(selUndoListItem)
                            Case GetType(Indicator), GetType(AudioVisualDetail), GetType(OpenEndedDetail), GetType(ScalesDetail), GetType(ValuesDetail)
                                Reload_Logframe_IndicatorDetail(selUndoListItem)
                            Case GetType(ValueRange)
                                Reload_Logframe_ValueRange(selUndoListItem)
                            Case GetType(Statement)
                                Reload_Logframe_Statement(selUndoListItem)
                            Case GetType(ResponseClass)
                                Reload_Logframe_ResponseClass(selUndoListItem)
                            Case GetType(Baseline), GetType(Target), GetType(PopulationTarget), GetType(BooleanValue), GetType(BooleanValuesMatrixRow), _
                                GetType(DoubleValue), GetType(DoubleValuesMatrixRow)
                                Reload_Logframe_Target(selUndoListItem)
                            Case GetType(VerificationSource)
                                Reload_Logframe_VerificationSource(selUndoListItem)
                            Case GetType(BudgetItemReference)
                                Reload_Logframe_BudgetItemReference(selUndoListItem)
                            Case GetType(Assumption)
                                Reload_Logframe_Assumption(selUndoListItem)
                        End Select
                End Select
            End With
        End If
    End Sub

    Private Sub Reload_Logframe_ListViews(ByVal selUndoListItem As UndoListItem)
        With CurrentProjectForm
            Select Case selUndoListItem.Item.GetType

                Case GetType(ProjectPartner), GetType(Organisation), GetType(Address), GetType(Email), GetType(TelephoneNumber), GetType(Website)
                    If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailGoal) Then
                        CType(.CurrentDetailLogframe, DetailGoal).lvPartners.LoadItems()
                    End If
                Case GetType(TargetGroup)
                    If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailPurpose) Then
                        CType(.CurrentDetailLogframe, DetailPurpose).lvDetailTargetGroups.LoadTargetGroups(False)
                    End If
                Case GetType(KeyMoment)
                    If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailOutput) Then
                        CType(.CurrentDetailLogframe, DetailOutput).lvDetailKeyMoments.LoadItems()
                    End If
                Case GetType(Activity), GetType(ActivityDetail)
                    If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailActivity) Then
                        CType(.CurrentDetailLogframe, DetailActivity).lvDetailProcesses.LoadItems()
                    End If
                    .dgvLogframe.Reload()
                Case Else
                    .dgvLogframe.Reload()
            End Select
        End With
    End Sub

    Private Sub Reload_Logframe_Logframe(ByVal selUndoListItem As UndoListItem)
        Dim selControl As Control
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailGoal) Then
                Dim objDetailGoal As DetailGoal = CType(.CurrentDetailLogframe, DetailGoal)
                selControl = GetControlByPropertyName(objDetailGoal.TabPageProjectInfo, strPropertyName)

                If selControl IsNot Nothing Then
                    If selControl.DataBindings.Count > 0 Then
                        selControl.DataBindings(0).ReadValue()
                    Else
                        If strPropertyName = "StartDate" Or strPropertyName = "EndDate" Then objDetailGoal.ModifyEndDate()
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_ActivityDetail(ByVal selUndoListItem As UndoListItem)
        Dim selControl As Control
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailActivity) Then
                Dim objDetailActivity As DetailActivity = CType(.CurrentDetailLogframe, DetailActivity)

                selControl = GetControlByPropertyName(objDetailActivity, strPropertyName)

                If selControl IsNot Nothing Then
                    If selControl.DataBindings.Count > 0 Then
                        selControl.DataBindings(0).ReadValue()
                    Else

                    End If
                    Select Case strPropertyName
                        Case "Period", "PeriodUnit", "PeriodDirection", "GuidReferenceMoment", "StartDate", "Duration", "DurationUnit", _
                            "DurationUntilEnd", "Preparation", "PreparationUnit", "PreparationFromStart", "FollowUp", "FollowUpUnit", _
                            "FollowUpUntilEnd"
                            objDetailActivity.UpdatePeriods()
                        Case "Repeat", "RepeatUnit", "RepeatTimes", "RepeatUntilEnd"
                            objDetailActivity.LoadRepeatDates()

                    End Select
                End If
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_IndicatorDetail(ByVal selUndoListItem As UndoListItem)
        Dim selControl As Control
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailIndicator) Then
                Dim objDetailIndicator As DetailIndicator = CType(.CurrentDetailLogframe, DetailIndicator)

                ReadValueOfControl(objDetailIndicator, strPropertyName)

                selControl = objDetailIndicator.PanelQuestionType.Controls(0)

                If selControl IsNot Nothing Then
                    Select Case strPropertyName
                        Case "ScoringSystem"
                            Select Case selControl.GetType
                                Case GetType(IndicatorValues)
                                    CType(selControl, IndicatorValues).SetScoringSystem()
                                Case GetType(IndicatorRatio)
                                    CType(selControl, IndicatorRatio).SetScoringSystem()
                                Case GetType(IndicatorFormula)
                                    CType(selControl, IndicatorFormula).SetScoringSystem()
                                Case GetType(IndicatorYesNo)
                                    CType(selControl, IndicatorYesNo).SetScoringSystem()
                                Case GetType(IndicatorMultipleOption)
                                    CType(selControl, IndicatorMultipleOption).SetScoringSystem()
                                Case GetType(IndicatorRanking)
                                    CType(selControl, IndicatorRanking).SetScoringSystem()
                                Case GetType(IndicatorScales)
                                    CType(selControl, IndicatorScales).SetScoringSystem()
                                Case GetType(IndicatorFrequencyLikert)
                                    CType(selControl, IndicatorFrequencyLikert).SetScoringSystem()
                            End Select
                        Case "TargetSystem"
                            Select Case selControl.GetType
                                Case GetType(IndicatorValues)
                                    CType(selControl, IndicatorValues).SetTargetSystem()
                                Case GetType(IndicatorRatio)
                                    CType(selControl, IndicatorRatio).SetTargetSystem()
                                Case GetType(IndicatorFormula)
                                    CType(selControl, IndicatorRatio).SetTargetSystem()
                            End Select
                    End Select
                End If

                Select Case objDetailIndicator.CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.Formula
                        'changes in valuesdetail: unit, nrDecimals...
                        CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsFormula).Reload()
                End Select
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_ValueRange(ByVal selUndoListItem As UndoListItem)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailIndicator) Then
                Dim objDetailIndicator As DetailIndicator = CType(.CurrentDetailLogframe, DetailIndicator)

                Select Case objDetailIndicator.CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.Formula
                        CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsFormula).Reload()
                End Select
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_Statement(ByVal selUndoListItem As UndoListItem)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailIndicator) Then
                Dim objDetailIndicator As DetailIndicator = CType(.CurrentDetailLogframe, DetailIndicator)

                Select Case objDetailIndicator.CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.MaxDiff
                        If TryCast(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsMaxDiff) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsMaxDiff).Reload()
                        End If
                    Case Indicator.QuestionTypes.Ratio
                        Select Case strPropertyName
                            Case "FirstLabel"
                                ReadValueOfControl(objDetailIndicator, "FirstStatement")
                                ReadValueOfControl(objDetailIndicator, "SecondStatement")
                            Case Else
                                ReadValueOfControl(objDetailIndicator, strPropertyName)
                        End Select
                    Case Indicator.QuestionTypes.Formula
                        If TryCast(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsFormula) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsFormula).Reload()
                        End If
                    Case Indicator.QuestionTypes.SemanticDiff
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsSemanticDiff) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsSemanticDiff).Reload()
                        End If
                    Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                        If TryCast(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsScales) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsScales).Reload()
                        End If
                    Case Indicator.QuestionTypes.LikertScale
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScaleLikert) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScaleLikert).Reload()
                        End If
                    Case Indicator.QuestionTypes.FrequencyLikert
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsFrequencyLikert) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsFrequencyLikert).Reload()
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_ResponseClass(ByVal selUndoListItem As UndoListItem)
        Dim selControl As Control
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailIndicator) Then
                Dim objDetailIndicator As DetailIndicator = CType(.CurrentDetailLogframe, DetailIndicator)
                selControl = objDetailIndicator.PanelQuestionType.Controls(0)

                Select Case objDetailIndicator.CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.MaxDiff
                        If selControl IsNot Nothing AndAlso selControl.GetType Is GetType(IndicatorMaxDiff) Then
                            CType(selControl, IndicatorMaxDiff).ReadScoreValues()
                        End If
                    Case Indicator.QuestionTypes.YesNo
                        If selControl IsNot Nothing AndAlso selControl.GetType Is GetType(IndicatorYesNo) Then
                            CType(selControl, IndicatorYesNo).ReadScoreValues()
                        End If
                    Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.Ranking, _
                        Indicator.QuestionTypes.LikertTypeScale, Indicator.QuestionTypes.SemanticDiff, Indicator.QuestionTypes.LikertScale, _
                        Indicator.QuestionTypes.FrequencyLikert

                        CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewResponseClasses).Invalidate()
                    Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                        If TryCast(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsScales) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentStatementsDataGridView, DataGridViewStatementsScales).Reload()
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_Target(ByVal selUndoListItem As UndoListItem)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailIndicator) Then
                Dim objDetailIndicator As DetailIndicator = CType(.CurrentDetailLogframe, DetailIndicator)

                Select Case objDetailIndicator.CurrentIndicator.QuestionType
                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                        CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsValues).Reload()
                    Case Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                        CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsFormula).Reload()
                    Case Indicator.QuestionTypes.YesNo, Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions

                        With CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsClasses)
                            If selUndoListItem.GetType = GetType(Target) Then
                                Dim selBooleanValue As BooleanValue = CType(selUndoListItem.OldValue, BooleanValues)(selUndoListItem.OldIndex)

                                .SetFutureTargets(selUndoListItem.Item, selUndoListItem.OldIndex, selBooleanValue.Value)
                            End If

                            .Reload()
                        End With
                    Case Indicator.QuestionTypes.Ranking
                        CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsRanking).Reload()
                    Case Indicator.QuestionTypes.LikertTypeScale

                        With CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScaleLikertType)
                            If selUndoListItem.GetType = GetType(Target) Then
                                Dim selBooleanValue As BooleanValue = CType(selUndoListItem.OldValue, BooleanValues)(selUndoListItem.OldIndex)

                                .SetFutureTargets(selUndoListItem.Item, selUndoListItem.OldIndex, selBooleanValue.Value)
                            End If

                            .Reload()
                        End With
                    Case Indicator.QuestionTypes.SemanticDiff
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsSemanticDiff) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsSemanticDiff).Reload()
                        End If
                    Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScales) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScales).Reload()
                        End If
                    Case Indicator.QuestionTypes.LikertScale
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScaleLikert) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsScaleLikert).Reload()
                        End If
                    Case Indicator.QuestionTypes.FrequencyLikert
                        If TryCast(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsFrequencyLikert) IsNot Nothing Then
                            CType(objDetailIndicator.CurrentTargetsDataGridView, DataGridViewTargetsFrequencyLikert).Reload()
                        End If
                End Select
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_VerificationSource(ByVal selUndoListItem As UndoListItem)
        Dim objDetailVerificationSource As DetailVerificationSource
        Dim selVerificationSource As VerificationSource = CType(selUndoListItem.Item, VerificationSource)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailVerificationSource) Then
                objDetailVerificationSource = CType(.CurrentDetailLogframe, DetailVerificationSource)

                ReadValueOfControl(objDetailVerificationSource, strPropertyName)
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_BudgetItemReference(ByVal selUndoListItem As UndoListItem)
        Dim objDetailResource As DetailResource
        Dim selBudgetItemReference As BudgetItemReference = CType(selUndoListItem.Item, BudgetItemReference)

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailResource) Then
                objDetailResource = CType(.CurrentDetailLogframe, DetailResource)

                Select Case selUndoListItem.PropertyName
                    Case "ReferenceBudgetItem"
                        objDetailResource.dgvBudgetItemReferences.CalculateTotalCost(selBudgetItemReference)
                    Case "Percentage"
                        objDetailResource.dgvBudgetItemReferences.CalculateTotalCost(selBudgetItemReference)
                    Case "TotalCost"
                        objDetailResource.dgvBudgetItemReferences.CalculatePercentage(selBudgetItemReference)
                End Select

                objDetailResource.dgvBudgetItemReferences.Reload()
            End If
        End With
    End Sub

    Private Sub Reload_Logframe_Assumption(ByVal selUndoListItem As UndoListItem)
        Dim objDetailAssumption As DetailAssumption
        Dim selAssumption As Assumption = CType(selUndoListItem.Item, Assumption)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            If .CurrentDetailLogframe IsNot Nothing AndAlso .CurrentDetailLogframe.GetType = GetType(DetailAssumption) Then
                objDetailAssumption = CType(.CurrentDetailLogframe, DetailAssumption)

                ReadValueOfControl(objDetailAssumption, strPropertyName)

                If strPropertyName = "RaidType" Then
                    objDetailAssumption.ChangeUserControls(selAssumption.RaidType)
                End If
            End If
        End With
    End Sub
#End Region

#Region "Reload Planning"
    Private Sub Reload_Planning(ByVal selUndoListItem As UndoListItem)
        If String.IsNullOrEmpty(selUndoListItem.PropertyName) = False Then
            With CurrentProjectForm
                Select Case selUndoListItem.PropertyName
                    Case "RTF"
                        .dgvPlanning.Reload()
                    Case Else
                        Select Case selUndoListItem.Item.GetType
                            Case GetType(KeyMoment)
                                Reload_Planning_KeyMoment(selUndoListItem)
                            Case GetType(ActivityDetail)
                                Reload_Planning_ActivityDetail(selUndoListItem)
                        End Select
                End Select
            End With
        End If
    End Sub

    Private Sub Reload_Planning_KeyMoment(ByVal selUndoListItem As UndoListItem)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            Select Case strPropertyName
                Case "Description", "KeyMoment", "Period", "PeriodUnit"
                    .dgvPlanning.Reload()
            End Select

            If .CurrentDetailPlanning IsNot Nothing AndAlso .CurrentDetailPlanning.GetType = GetType(DetailKeyMoment) Then
                Dim objDetailKeyMoment As DetailKeyMoment = CType(.CurrentDetailPlanning, DetailKeyMoment)

                ReadValueOfControl(objDetailKeyMoment, strPropertyName)

                objDetailKeyMoment.Invalidate()
            End If
        End With
    End Sub

    Private Sub Reload_Planning_ActivityDetail(ByVal selUndoListItem As UndoListItem)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            Select Case strPropertyName
                Case "StartDate", "Period", "PeriodUnit", "Duration", "DurationUnit"
                    .dgvPlanning.Reload()
            End Select

            If .CurrentDetailPlanning IsNot Nothing AndAlso .CurrentDetailPlanning.GetType = GetType(DetailActivity) Then
                Dim objDetailActivity As DetailActivity = CType(.CurrentDetailPlanning, DetailActivity)

                ReadValueOfControl(objDetailActivity, strPropertyName)

                Select Case strPropertyName
                    Case "Period", "PeriodUnit", "PeriodDirection", "GuidReferenceMoment", "StartDate", "Duration", "DurationUnit", _
                        "DurationUntilEnd", "Preparation", "PreparationUnit", "PreparationFromStart", "FollowUp", "FollowUpUnit", _
                        "FollowUpUntilEnd"
                        objDetailActivity.UpdatePeriods()
                    Case "Repeat", "RepeatUnit", "RepeatTimes", "RepeatUntilEnd"
                        objDetailActivity.LoadRepeatDates()

                End Select

                objDetailActivity.Invalidate()
            End If
        End With
    End Sub
#End Region

#Region "Reload Budget"
    Private Sub Reload_Budget(ByVal selUndoListItem As UndoListItem)
        Select Case selUndoListItem.Item.GetType
            Case GetType(BudgetItem)
                Reload_Budget_BudgetItem(selUndoListItem)
            Case GetType(BudgetItemReference)
                Reload_Budget_BudgetItemReference(selUndoListItem)
        End Select
        
    End Sub

    Private Sub Reload_Budget_BudgetItem(ByVal selUndoListItem As UndoListItem)
        Dim selBudgetItem As BudgetItem = CType(selUndoListItem.Item, BudgetItem)
        Dim strPropertyName As String = selUndoListItem.PropertyName

        With CurrentProjectForm
            .ReloadAllDataGridViewBudgetYears()

            If String.IsNullOrEmpty(strPropertyName) = False Then
                Select Case strPropertyName
                    Case "Type", "CalculationGuidRatio", "Ratio"
                        If .CurrentDetailBudget IsNot Nothing AndAlso .CurrentDetailBudget.GetType = GetType(DetailBudgetItem) Then
                            Dim objDetailBudgetItem As DetailBudgetItem = CType(.CurrentDetailBudget, DetailBudgetItem)

                            ReadValueOfControl(objDetailBudgetItem, strPropertyName)

                            If strPropertyName = "Type" Then objDetailBudgetItem.ChangeUserControls(objDetailBudgetItem.CurrentBudgetItem.Type)
                        End If
                End Select
            End If

        End With
    End Sub

    Private Sub Reload_Budget_BudgetItemReference(ByVal selUndoListItem As UndoListItem)
        With CurrentProjectForm
            .ReloadAllDataGridViewBudgetYears()

            If .CurrentDetailBudget IsNot Nothing AndAlso .CurrentDetailBudget.GetType = GetType(DetailBudgetItem) Then
                Dim objDetailBudgetItem As DetailBudgetItem = CType(.CurrentDetailBudget, DetailBudgetItem)

                objDetailBudgetItem.lvBudgetItemReferences.LoadItems()
            End If
        End With
    End Sub
#End Region
End Class
