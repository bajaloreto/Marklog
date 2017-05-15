Partial Public Class frmProject

    Private Structure TextSelection
        Public SelectGoals As Boolean
        Public SelectPurposes As Boolean
        Public SelectOutputs As Boolean
        Public SelectActivities As Boolean
        Public SelectIndicators As Boolean
        Public SelectVerificationSources As Boolean
        Public SelectStatements As Boolean
        Public SelectResources As Boolean
        Public SelectAssumptions As Boolean
        Public SelectBudgetItems As Boolean

        Public Sub SetSelectors(ByVal intTextSelectionIndex As Integer)
            SelectAssumptions = False
            SelectIndicators = False
            SelectResources = False
            SelectStatements = False
            SelectVerificationSources = False

            Select Case intTextSelectionIndex
                Case TextSelectionValues.SelectAll
                    SelectGoals = True
                    SelectPurposes = True
                    SelectOutputs = True
                    SelectActivities = True
                    SelectIndicators = True
                    SelectVerificationSources = True
                    SelectResources = True
                    SelectAssumptions = True
                    SelectStatements = True
                Case TextSelectionValues.SelectLogframe
                    SelectGoals = True
                    SelectPurposes = True
                    SelectOutputs = True
                    SelectActivities = True
                    SelectIndicators = True
                    SelectVerificationSources = True
                    SelectResources = True
                    SelectAssumptions = True
                Case TextSelectionValues.SelectStructs
                    SelectGoals = True
                    SelectPurposes = True
                    SelectOutputs = True
                    SelectActivities = True
                Case TextSelectionValues.SelectGoals
                    SelectGoals = True
                Case TextSelectionValues.SelectPurposes
                    SelectPurposes = True
                Case TextSelectionValues.SelectOutputs
                    SelectOutputs = True
                Case TextSelectionValues.SelectActivities
                    SelectActivities = True
                Case TextSelectionValues.SelectIndicators
                    SelectIndicators = True
                Case TextSelectionValues.SelectResources
                    SelectResources = True
                Case TextSelectionValues.SelectAssumptions
                    SelectAssumptions = True
                Case TextSelectionValues.SelectVerificationSources
                    SelectVerificationSources = True
                Case TextSelectionValues.SelectStatements
                    SelectStatements = True
                Case TextSelectionValues.SelectBudgetItems
                    SelectBudgetItems = True
            End Select
        End Sub

    End Structure

#Region "Select text"
    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex

        Select Case TabControlProject.SelectedTab.Name
            Case TabPageLogframe.Name
                Select Case TextSelectionIndex
                    Case TextSelectionValues.SelectNothing, TextSelectionValues.SelectAll
                        dgvLogframe.SelectText(Me.TextSelectionIndex)

                        If CurrentDetailLogframe IsNot Nothing AndAlso CurrentDetailLogframe.GetType Is GetType(DetailIndicator) Then
                            CType(CurrentDetailLogframe, DetailIndicator).SelectText(intTextSelectionIndex)
                        End If
                    Case TextSelectionValues.SelectStatements
                        If CurrentDetailLogframe IsNot Nothing AndAlso CurrentDetailLogframe.GetType Is GetType(DetailIndicator) Then
                            CType(CurrentDetailLogframe, DetailIndicator).SelectText(intTextSelectionIndex)
                        End If
                    Case Else
                        dgvLogframe.SelectText(Me.TextSelectionIndex)

                        If CurrentDetailLogframe IsNot Nothing AndAlso CurrentDetailLogframe.GetType Is GetType(DetailIndicator) Then
                            CType(CurrentDetailLogframe, DetailIndicator).SelectText(TextSelectionValues.SelectNothing)
                        End If
                End Select
            Case TabPagePlanning.Name
                If intTextSelectionIndex = TextSelectionValues.SelectAll Then
                    Me.TextSelectionIndex = TextSelectionValues.SelectActivities
                Else
                    Me.TextSelectionIndex = TextSelectionValues.SelectNothing
                End If

                dgvPlanning.SelectText(Me.TextSelectionIndex)
            Case TabPageBudget.Name
                If intTextSelectionIndex = TextSelectionValues.SelectAll Then
                    Me.TextSelectionIndex = TextSelectionValues.SelectBudgetItems
                Else
                    Me.TextSelectionIndex = TextSelectionValues.SelectNothing
                End If

                CurrentDataGridViewBudgetYear.SelectText(Me.TextSelectionIndex)
        End Select
    End Sub

    Private Sub SelectText_Reload()
        Select Case CurrentMainTab
            Case 1 'Logframe
                dgvLogframe.Reload()
                If CurrentDetailLogframe IsNot Nothing Then CurrentDetailLogframe.Refresh()
            Case 2 'Planning
                dgvPlanning.Reload()
            Case 3 'Budget
                For i = 0 To CurrentBudget.BudgetYears.Count - 1
                    GetDataGridViewBudgetYear(i).Reload()
                Next
                'CurrentDataGridViewBudgetYear.Reload()
        End Select
    End Sub
#End Region

#Region "Font"
    Public Sub ChangeFont()
        If Me.TextSelectionIndex = TextSelectionValues.SelectNothing Then
            'change selected text in editing control

            Dim ctl As RichTextEditingControlLogframe = CurrentEditingControl
            If ctl IsNot Nothing Then
                ctl.ChangeFont()
                UndoRedo.SetChangedText(ctl.Rtf, ctl.SelectedText)
            End If
        Else
            'change text of all items of a certain category
            Dim objTextSelection As New TextSelection
            objTextSelection.SetSelectors(Me.TextSelectionIndex)

            With objTextSelection
                Using objRtf As New RichTextManager
                    For Each selGoal As Goal In CurrentLogFrame.Goals
                        If .SelectGoals = True Then selGoal.RTF = objRtf.ChangeFont(selGoal.RTF)
                        ChangeFont_Children(objRtf, selGoal)
                    Next
                    For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                        If .SelectPurposes = True Then selPurpose.RTF = objRtf.ChangeFont(selPurpose.RTF)
                        ChangeFont_Children(objRtf, selPurpose)

                        For Each selOutput As Output In selPurpose.Outputs
                            If .SelectOutputs = True Then selOutput.RTF = objRtf.ChangeFont(selOutput.RTF)
                            ChangeFont_Children(objRtf, selOutput)

                            ChangeFont_Activities(objRtf, selOutput.Activities, .SelectActivities)
                        Next
                    Next
                    For Each selBudgetYear As BudgetYear In CurrentBudget.BudgetYears
                        ChangeFont_BudgetItems(objRtf, selBudgetYear.BudgetItems, .SelectBudgetItems)
                    Next
                End Using
            End With

            SelectText_Reload()
        End If
    End Sub

    Private Sub ChangeFont_Activities(ByVal objRtf As RichTextManager, ByVal selActivities As Activities, ByVal boolSelectActivities As Boolean)
        For Each selActivity As Activity In selActivities
            If boolSelectActivities = True Then selActivity.RTF = objRtf.ChangeFont(selActivity.RTF)
            ChangeFont_Children(objRtf, selActivity)

            If selActivity.Activities.Count > 0 Then
                ChangeFont_Activities(objRtf, selActivity.Activities, boolSelectActivities)
            End If
        Next
    End Sub

    Private Sub ChangeFont_BudgetItems(ByVal objRtf As RichTextManager, ByVal selBudgetItems As BudgetItems, ByVal boolSelectBudgetItems As Boolean)
        For Each selBudgetItem As BudgetItem In selBudgetItems
            If boolSelectBudgetItems = True Then selBudgetItem.RTF = objRtf.ChangeFont(selBudgetItem.RTF)

            If selBudgetItem.BudgetItems.Count > 0 Then
                ChangeFont_BudgetItems(objRtf, selBudgetItem.BudgetItems, boolSelectBudgetItems)
            End If
        Next
    End Sub

    Private Sub ChangeFont_Children(ByVal objRtf As RichTextManager, ByVal selStruct As Struct)
        Dim selActivity As Activity = TryCast(selStruct, Activity)
        Dim objTextSelection As New TextSelection
        objTextSelection.SetSelectors(Me.TextSelectionIndex)

        With objTextSelection
            ChangeFont_Children_Indicators(objRtf, objTextSelection, selStruct.Indicators)

            If selActivity IsNot Nothing Then
                For Each selResource As Resource In selActivity.Resources
                    If .SelectResources = True Then selResource.RTF = objRtf.ChangeFont(selResource.RTF)
                Next
            End If

            If .SelectAssumptions = True Then
                For Each selAssumption As Assumption In selStruct.Assumptions
                    selAssumption.RTF = objRtf.ChangeFont(selAssumption.RTF)
                Next
            End If
        End With
    End Sub

    Private Sub ChangeFont_Children_Indicators(ByVal objRtf As RichTextManager, ByVal objTextSelection As TextSelection, ByVal selIndicators As Indicators)
        With objTextSelection
            For Each selIndicator As Indicator In selIndicators
                If .SelectIndicators = True Then selIndicator.RTF = objRtf.ChangeFont(selIndicator.RTF)
                If .SelectVerificationSources = True Then
                    For Each selVerificationSource As VerificationSource In selIndicator.VerificationSources
                        selVerificationSource.RTF = objRtf.ChangeFont(selVerificationSource.RTF)
                    Next
                End If
                If .SelectStatements = True Then
                    For Each selStatement As Statement In selIndicator.Statements
                        selStatement.FirstLabel = objRtf.ChangeFont(selStatement.FirstLabel)
                        selStatement.SecondLabel = objRtf.ChangeFont(selStatement.SecondLabel)
                    Next
                End If

                If selIndicator.Indicators.Count > 0 Then
                    ChangeFont_Children_Indicators(objRtf, objTextSelection, selIndicator.Indicators)
                End If
            Next
        End With
    End Sub

    Public Sub ChangeFontCase(ByVal intCase As Integer)
        If Me.TextSelectionIndex = TextSelectionValues.SelectNothing Then
            Dim ctl As RichTextEditingControlLogframe = CurrentEditingControl
            If ctl IsNot Nothing Then ctl.ChangeFontCase(intCase)

            UndoRedo.SetChangedText(ctl.Rtf, ctl.SelectedText)
        Else
            'change text of all items of a certain category
            Dim objTextSelection As New TextSelection
            objTextSelection.SetSelectors(Me.TextSelectionIndex)

            With objTextSelection
                Using objRtf As New RichTextManager
                    For Each selGoal As Goal In CurrentLogFrame.Goals
                        If .SelectGoals = True Then selGoal.RTF = objRtf.ChangeFontCase(selGoal.RTF, intCase)
                        ChangeFontCase_Children(objRtf, selGoal, intCase)
                    Next
                    For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                        If .SelectPurposes = True Then selPurpose.RTF = objRtf.ChangeFontCase(selPurpose.RTF, intCase)
                        ChangeFontCase_Children(objRtf, selPurpose, intCase)
                        For Each selOutput As Output In selPurpose.Outputs
                            If .SelectOutputs = True Then selOutput.RTF = objRtf.ChangeFontCase(selOutput.RTF, intCase)
                            ChangeFontCase_Children(objRtf, selOutput, intCase)

                            ChangeFontCase_Activities(objRtf, selOutput.Activities, .SelectActivities, intCase)
                        Next
                    Next
                    For Each selBudgetYear As BudgetYear In CurrentBudget.BudgetYears
                        ChangeFontCase_BudgetItems(objRtf, selBudgetYear.BudgetItems, .SelectBudgetItems, intCase)
                    Next
                End Using
            End With

            SelectText_Reload()
        End If
    End Sub

    Private Sub ChangeFontCase_Activities(ByVal objRtf As RichTextManager, ByVal selActivities As Activities, ByVal boolSelectActivities As Boolean, ByVal intCase As Integer)
        For Each selActivity As Activity In selActivities
            If boolSelectActivities = True Then selActivity.RTF = objRtf.ChangeFontCase(selActivity.RTF, intCase)
            ChangeFontCase_Children(objRtf, selActivity, intCase)

            If selActivity.Activities.Count > 0 Then
                ChangeFontCase_Activities(objRtf, selActivity.Activities, boolSelectActivities, intCase)
            End If
        Next
    End Sub

    Private Sub ChangeFontCase_BudgetItems(ByVal objRtf As RichTextManager, ByVal selBudgetItems As BudgetItems, ByVal boolSelectBudgetItems As Boolean, ByVal intCase As Integer)
        For Each selBudgetItem As BudgetItem In selBudgetItems
            If boolSelectBudgetItems = True Then selBudgetItem.RTF = objRtf.ChangeFontCase(selBudgetItem.RTF, intCase)

            If selBudgetItem.BudgetItems.Count > 0 Then
                ChangeFontCase_BudgetItems(objRtf, selBudgetItem.BudgetItems, boolSelectBudgetItems, intCase)
            End If
        Next
    End Sub

    Private Sub ChangeFontCase_Children(ByVal objRtf As RichTextManager, ByVal selStruct As Struct, ByVal intCase As Integer)
        Dim selActivity As Activity = TryCast(selStruct, Activity)
        Dim objTextSelection As New TextSelection

        objTextSelection.SetSelectors(Me.TextSelectionIndex)

        With objTextSelection
            ChangeFontCase_Children_Indicators(objRtf, objTextSelection, selStruct.Indicators, intCase)

            If selActivity IsNot Nothing Then
                For Each selResource As Resource In selActivity.Resources
                    If .SelectResources = True Then selResource.RTF = objRtf.ChangeFontCase(selResource.RTF, intCase)
                Next
            End If

            If .SelectAssumptions = True Then
                For Each selAssumption As Assumption In selStruct.Assumptions
                    selAssumption.RTF = objRtf.ChangeFontCase(selAssumption.RTF, intCase)
                Next
            End If
        End With
    End Sub

    Private Sub ChangeFontCase_Children_Indicators(ByVal objRtf As RichTextManager, ByVal objTextSelection As TextSelection, ByVal selIndicators As Indicators, ByVal intCase As Integer)
        With objTextSelection
            For Each selIndicator As Indicator In selIndicators
                If .SelectIndicators = True Then selIndicator.RTF = objRtf.ChangeFontCase(selIndicator.RTF, intCase)
                If .SelectVerificationSources = True Then
                    For Each selVerifiationSource As VerificationSource In selIndicator.VerificationSources
                        selVerifiationSource.RTF = objRtf.ChangeFontCase(selVerifiationSource.RTF, intCase)
                    Next
                End If
                If .SelectStatements = True Then
                    For Each selStatement As Statement In selIndicator.Statements
                        selStatement.FirstLabel = objRtf.ChangeFontCase(selStatement.FirstLabel, intCase)
                        selStatement.SecondLabel = objRtf.ChangeFontCase(selStatement.SecondLabel, intCase)
                    Next
                End If

                If selIndicator.Indicators.Count > 0 Then
                    ChangeFontCase_Children_Indicators(objRtf, objTextSelection, selIndicator.Indicators, intCase)
                End If
            Next
        End With
    End Sub

    Public Sub ChangeFontOffSet(ByVal intDirection As Integer)
        If Me.TextSelectionIndex = TextSelectionValues.SelectNothing Then
            Dim ctl As RichTextEditingControlLogframe = CurrentEditingControl
            ctl.ChangeFont_CharOffset(intDirection)

            UndoRedo.SetChangedText(ctl.Rtf, ctl.SelectedText)
        End If
    End Sub
#End Region

#Region "Colour"
    Public Sub ChangeFontColor(ByVal selColor As Color, ByVal boolMarker As Boolean)
        If Me.TextSelectionIndex = TextSelectionValues.SelectNothing Then
            Dim ctl As RichTextEditingControlLogframe = CurrentEditingControl
            If boolMarker = False Then
                ctl.ChangeTextColor(selColor)
            Else
                ctl.ChangeMarkerColor(selColor)
            End If

            UndoRedo.SetChangedText(ctl.Rtf, ctl.SelectedText)
        Else
            'change text of all items of a certain category
            Dim objTextSelection As New TextSelection
            objTextSelection.SetSelectors(Me.TextSelectionIndex)

            With objTextSelection
                Using objRtf As New RichTextManager
                    For Each selGoal As Goal In CurrentLogFrame.Goals
                        If .SelectGoals = True Then
                            If boolMarker = False Then
                                selGoal.RTF = objRtf.ChangeFontColor(selGoal.RTF, selColor)
                            Else
                                selGoal.RTF = objRtf.ChangeMarkerColor(selGoal.RTF, selColor)
                            End If
                        End If
                        ChangeFontColor_Children(objRtf, selGoal, selColor, boolMarker)
                    Next
                    For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                        If .SelectPurposes = True Then
                            If boolMarker = False Then
                                selPurpose.RTF = objRtf.ChangeFontColor(selPurpose.RTF, selColor)
                            Else
                                selPurpose.RTF = objRtf.ChangeMarkerColor(selPurpose.RTF, selColor)
                            End If
                        End If
                        ChangeFontColor_Children(objRtf, selPurpose, selColor, boolMarker)
                        For Each selOutput As Output In selPurpose.Outputs
                            If .SelectOutputs = True Then
                                If boolMarker = False Then
                                    selOutput.RTF = objRtf.ChangeFontColor(selOutput.RTF, selColor)
                                Else
                                    selOutput.RTF = objRtf.ChangeMarkerColor(selOutput.RTF, selColor)
                                End If
                            End If
                            ChangeFontColor_Children(objRtf, selOutput, selColor, boolMarker)
                            ChangeFontColor_Activities(objRtf, selOutput.Activities, .SelectActivities, selColor, boolMarker)
                        Next
                    Next
                    For Each selBudgetYear As BudgetYear In CurrentBudget.BudgetYears
                        ChangeFontColor_BudgetItems(objRtf, selBudgetYear.BudgetItems, .SelectBudgetItems, selColor, boolMarker)
                    Next
                End Using
            End With

            SelectText_Reload()
        End If
    End Sub

    Private Sub ChangeFontColor_Activities(ByVal objRtf As RichTextManager, ByVal selActivities As Activities, ByVal boolSelectActivities As Boolean, _
                                           ByVal selColor As Color, ByVal boolMarker As Boolean)
        For Each selActivity As Activity In selActivities
            If boolSelectActivities = True Then
                If boolMarker = False Then
                    selActivity.RTF = objRtf.ChangeFontColor(selActivity.RTF, selColor)
                Else
                    selActivity.RTF = objRtf.ChangeMarkerColor(selActivity.RTF, selColor)
                End If
            End If
            ChangeFontColor_Children(objRtf, selActivity, selColor, boolMarker)

            If selActivity.Activities.Count > 0 Then
                ChangeFontColor_Activities(objRtf, selActivity.Activities, boolSelectActivities, selColor, boolMarker)
            End If
        Next
    End Sub

    Private Sub ChangeFontColor_BudgetItems(ByVal objRtf As RichTextManager, ByVal selBudgetItems As BudgetItems, ByVal boolSelectBudgetItems As Boolean, _
                                            ByVal selColor As Color, ByVal boolMarker As Boolean)
        For Each selBudgetItem As BudgetItem In selBudgetItems
            If boolSelectBudgetItems = True Then
                If boolMarker = False Then
                    selBudgetItem.RTF = objRtf.ChangeFontColor(selBudgetItem.RTF, selColor)
                Else
                    selBudgetItem.RTF = objRtf.ChangeMarkerColor(selBudgetItem.RTF, selColor)
                End If
            End If

            If selBudgetItem.BudgetItems.Count > 0 Then
                ChangeFontColor_BudgetItems(objRtf, selBudgetItem.BudgetItems, boolSelectBudgetItems, selColor, boolMarker)
            End If
        Next
    End Sub

    Private Sub ChangeFontColor_Children(ByVal objRtf As RichTextManager, ByVal selStruct As Struct, ByVal selColor As Color, ByVal boolMarker As Boolean)
        Dim selActivity As Activity = TryCast(selStruct, Activity)
        Dim objTextSelection As New TextSelection

        objTextSelection.SetSelectors(Me.TextSelectionIndex)

        With objTextSelection
            For Each selIndicator As Indicator In selStruct.Indicators

            Next

            If selActivity IsNot Nothing Then
                For Each selResource As Resource In selActivity.Resources
                    If .SelectResources = True Then
                        If boolMarker = False Then
                            selResource.RTF = objRtf.ChangeFontColor(selResource.RTF, selColor)
                        Else
                            selResource.RTF = objRtf.ChangeMarkerColor(selResource.RTF, selColor)
                        End If
                    End If
                Next
            End If

            If .SelectAssumptions = True Then
                For Each selAssumption As Assumption In selStruct.Assumptions
                    If boolMarker = False Then
                        selAssumption.RTF = objRtf.ChangeFontColor(selAssumption.RTF, selColor)
                    Else
                        selAssumption.RTF = objRtf.ChangeMarkerColor(selAssumption.RTF, selColor)
                    End If
                Next
            End If
        End With
    End Sub

    Private Sub ChangeFontColor_Children_Indicators(ByVal objRtf As RichTextManager, ByVal objTextSelection As TextSelection, ByVal selIndicators As Indicators, _
                                                    ByVal selColor As Color, ByVal boolMarker As Boolean)
        With objTextSelection
            For Each selIndicator As Indicator In selIndicators
                If .SelectIndicators = True Then
                    If boolMarker = False Then
                        selIndicator.RTF = objRtf.ChangeFontColor(selIndicator.RTF, selColor)
                    Else
                        selIndicator.RTF = objRtf.ChangeMarkerColor(selIndicator.RTF, selColor)
                    End If
                End If
                If .SelectVerificationSources = True Then
                    For Each selVerificationSource As VerificationSource In selIndicator.VerificationSources
                        If boolMarker = False Then
                            selVerificationSource.RTF = objRtf.ChangeFontColor(selVerificationSource.RTF, selColor)
                        Else
                            selVerificationSource.RTF = objRtf.ChangeMarkerColor(selVerificationSource.RTF, selColor)
                        End If
                    Next
                End If
                If .SelectStatements = True Then
                    For Each selStatement As Statement In selIndicator.Statements
                        If boolMarker = False Then
                            selStatement.FirstLabel = objRtf.ChangeFontColor(selStatement.FirstLabel, selColor)
                            selStatement.SecondLabel = objRtf.ChangeFontColor(selStatement.SecondLabel, selColor)
                        Else
                            selStatement.FirstLabel = objRtf.ChangeMarkerColor(selStatement.FirstLabel, selColor)
                            selStatement.SecondLabel = objRtf.ChangeMarkerColor(selStatement.SecondLabel, selColor)
                        End If
                    Next
                End If

                If selIndicator.Indicators.Count > 0 Then
                    ChangeFontColor_Children_Indicators(objRtf, objTextSelection, selIndicator.Indicators, selColor, boolMarker)
                End If
            Next
        End With
    End Sub
#End Region

#Region "Paragraph"
    Public Sub ChangeParagraphAlignment()
        If Me.TextSelectionIndex = TextSelectionValues.SelectNothing Then
            Dim ctl As RichTextEditingControlLogframe = CurrentEditingControl
            Dim strOldText As String = ctl.Rtf
            ctl.ChangeTextAlignment(CurrentText.HorizontalAlignment)

            Select Case CurrentText.HorizontalAlignment
                Case HorizontalAlignment.Left
                    UndoRedo.ParagraphAlignLeft(strOldText, ctl.Rtf)
                Case HorizontalAlignment.Center
                    UndoRedo.ParagraphAlignCenter(strOldText, ctl.Rtf)
                Case HorizontalAlignment.Right
                    UndoRedo.ParagraphAlignRight(strOldText, ctl.Rtf)
            End Select
        Else
            'change text of all items of a certain category
            Dim objTextSelection As New TextSelection
            objTextSelection.SetSelectors(Me.TextSelectionIndex)

            With objTextSelection
                Using objRtf As New RichTextManager
                    Dim intAlignment As Integer = CurrentText.HorizontalAlignment

                    For Each selGoal As Goal In CurrentLogFrame.Goals
                        If .SelectGoals = True Then selGoal.RTF = objRtf.ChangeAlignment(selGoal.RTF, intAlignment)
                        ChangeParagraphAlignment_Children(objRtf, selGoal, intAlignment)
                    Next
                    For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                        If .SelectPurposes = True Then selPurpose.RTF = objRtf.ChangeAlignment(selPurpose.RTF, intAlignment)
                        ChangeParagraphAlignment_Children(objRtf, selPurpose, intAlignment)
                        For Each selOutput As Output In selPurpose.Outputs
                            If .SelectOutputs = True Then selOutput.RTF = objRtf.ChangeAlignment(selOutput.RTF, intAlignment)

                            ChangeParagraphAlignment_Children(objRtf, selOutput, intAlignment)
                            ChangeParagraphAlignment_Activities(objRtf, selOutput.Activities, .SelectActivities, intAlignment)
                        Next
                    Next
                    For Each selBudgetYear As BudgetYear In CurrentBudget.BudgetYears
                        ChangeParagraphAlignment_BudgetItems(objRtf, selBudgetYear.BudgetItems, .SelectBudgetItems, intAlignment)
                    Next
                End Using
            End With

            SelectText_Reload()
        End If
    End Sub

    Private Sub ChangeParagraphAlignment_Activities(ByVal objRtf As RichTextManager, ByVal selActivities As Activities, ByVal boolSelectActivities As Boolean, _
                                                    ByVal intAlignment As Integer)
        For Each selActivity As Activity In selActivities
            If boolSelectActivities = True Then selActivity.RTF = objRtf.ChangeAlignment(selActivity.RTF, intAlignment)
            ChangeParagraphAlignment_Children(objRtf, selActivity, intAlignment)

            If selActivity.Activities.Count > 0 Then
                ChangeParagraphAlignment_Activities(objRtf, selActivity.Activities, boolSelectActivities, intAlignment)
            End If
        Next
    End Sub

    Private Sub ChangeParagraphAlignment_BudgetItems(ByVal objRtf As RichTextManager, ByVal selBudgetItems As BudgetItems, ByVal boolSelectBudgetItems As Boolean, _
                                                     ByVal intAlignment As Integer)
        For Each selBudgetItem As BudgetItem In selBudgetItems
            If boolSelectBudgetItems = True Then selBudgetItem.RTF = objRtf.ChangeAlignment(selBudgetItem.RTF, intAlignment)

            If selBudgetItem.BudgetItems.Count > 0 Then
                ChangeParagraphAlignment_BudgetItems(objRtf, selBudgetItem.BudgetItems, boolSelectBudgetItems, intAlignment)
            End If
        Next
    End Sub

    Private Sub ChangeParagraphAlignment_Children(ByVal objRtf As RichTextManager, ByVal selStruct As Struct, ByVal intAlignment As Integer)
        Dim selActivity As Activity = TryCast(selStruct, Activity)
        Dim objTextSelection As New TextSelection

        objTextSelection.SetSelectors(Me.TextSelectionIndex)

        With objTextSelection
            ChangeParagraphAlignment_Children_Indicators(objRtf, objTextSelection, selStruct.Indicators, intAlignment)

            If selActivity IsNot Nothing Then
                For Each selResource As Resource In selActivity.Resources
                    If .SelectResources = True Then selResource.RTF = objRtf.ChangeAlignment(selResource.RTF, intAlignment)
                Next
            End If

            If .SelectAssumptions = True Then
                For Each selAssumption As Assumption In selStruct.Assumptions
                    selAssumption.RTF = objRtf.ChangeAlignment(selAssumption.RTF, intAlignment)
                Next
            End If
        End With
    End Sub

    Private Sub ChangeParagraphAlignment_Children_Indicators(ByVal objRtf As RichTextManager, ByVal objTextSelection As TextSelection, ByVal selIndicators As Indicators, ByVal intAlignment As Integer)
        With objTextSelection
            For Each selIndicator As Indicator In selIndicators
                If .SelectIndicators = True Then selIndicator.RTF = objRtf.ChangeAlignment(selIndicator.RTF, intAlignment)
                If .SelectVerificationSources = True Then
                    For Each selVerificationSource As VerificationSource In selIndicator.VerificationSources
                        selVerificationSource.RTF = objRtf.ChangeAlignment(selVerificationSource.RTF, intAlignment)
                    Next
                End If
                If .SelectStatements = True Then
                    For Each selStatement As Statement In selIndicator.Statements
                        selStatement.FirstLabel = objRtf.ChangeAlignment(selStatement.FirstLabel, intAlignment)
                        selStatement.SecondLabel = objRtf.ChangeAlignment(selStatement.SecondLabel, intAlignment)
                    Next
                End If

                If selIndicator.Indicators.Count > 0 Then
                    ChangeParagraphAlignment_Children_Indicators(objRtf, objTextSelection, selIndicator.Indicators, intAlignment)
                End If
            Next
        End With
    End Sub

    Public Sub ChangeParagraphLeftIndent(ByVal intIncrement As Integer)
        If Me.TextSelectionIndex = TextSelectionValues.SelectNothing Then
            Dim boolEditMode As Boolean = True
            If CurrentDataGridView.IsCurrentCellInEditMode = False Then boolEditMode = False

            Dim ctl As RichTextEditingControlLogframe = CurrentEditingControl
            Dim strOldText As String = ctl.Rtf

            ctl.ChangeLeftIndent(intIncrement)
            UndoRedo.ParagraphLeftIndentChanged(strOldText, ctl.Rtf)
        Else
            'change text of all items of a certain category
            Dim objTextSelection As New TextSelection
            objTextSelection.SetSelectors(Me.TextSelectionIndex)

            With objTextSelection
                Using objRtf As New RichTextManager
                    For Each selGoal As Goal In CurrentLogFrame.Goals
                        If .SelectGoals = True Then selGoal.RTF = objRtf.ChangeLeftIndent(selGoal.RTF, intIncrement)
                        ChangeParagraphLeftIndent_Children(objRtf, selGoal, intIncrement)
                    Next
                    For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                        If .SelectPurposes = True Then selPurpose.RTF = objRtf.ChangeLeftIndent(selPurpose.RTF, intIncrement)
                        ChangeParagraphLeftIndent_Children(objRtf, selPurpose, intIncrement)
                        For Each selOutput As Output In selPurpose.Outputs
                            If .SelectOutputs = True Then selOutput.RTF = objRtf.ChangeLeftIndent(selOutput.RTF, intIncrement)

                            ChangeParagraphLeftIndent_Children(objRtf, selOutput, intIncrement)
                            ChangeParagraphLeftIndent_Activities(objRtf, selOutput.Activities, .SelectActivities, intIncrement)
                        Next
                    Next
                    For Each selBudgetYear As BudgetYear In CurrentBudget.BudgetYears
                        ChangeParagraphLeftIndent_BudgetItems(objRtf, selBudgetYear.BudgetItems, .SelectBudgetItems, intIncrement)
                    Next
                End Using
            End With

            SelectText_Reload()
        End If
    End Sub

    Private Sub ChangeParagraphLeftIndent_Activities(ByVal objRtf As RichTextManager, ByVal selActivities As Activities, ByVal boolSelectActivities As Boolean, _
                                                     ByVal intIncrement As Integer)
        For Each selActivity As Activity In selActivities
            If boolSelectActivities = True Then selActivity.RTF = objRtf.ChangeLeftIndent(selActivity.RTF, intIncrement)
            ChangeParagraphLeftIndent_Children(objRtf, selActivity, intIncrement)

            If selActivity.Activities.Count > 0 Then
                ChangeParagraphLeftIndent_Activities(objRtf, selActivity.Activities, boolSelectActivities, intIncrement)
            End If
        Next
    End Sub

    Private Sub ChangeParagraphLeftIndent_BudgetItems(ByVal objRtf As RichTextManager, ByVal selBudgetItems As BudgetItems, ByVal boolSelectBudgetItems As Boolean, _
                                                     ByVal intIncrement As Integer)
        For Each selBudgetItem As BudgetItem In selBudgetItems
            If boolSelectBudgetItems = True Then selBudgetItem.RTF = objRtf.ChangeLeftIndent(selBudgetItem.RTF, intIncrement)

            If selBudgetItem.BudgetItems.Count > 0 Then
                ChangeParagraphLeftIndent_BudgetItems(objRtf, selBudgetItem.BudgetItems, boolSelectBudgetItems, intIncrement)
            End If
        Next
    End Sub

    Private Sub ChangeParagraphLeftIndent_Children(ByVal objRtf As RichTextManager, ByVal selStruct As Struct, ByVal intIncrement As Integer)
        Dim selActivity As Activity = TryCast(selStruct, Activity)
        Dim objTextSelection As New TextSelection

        objTextSelection.SetSelectors(Me.TextSelectionIndex)

        With objTextSelection
            For Each selInd As Indicator In selStruct.Indicators

            Next

            If selActivity IsNot Nothing Then
                For Each selResource As Resource In selActivity.Resources
                    If .SelectResources = True Then selResource.RTF = objRtf.ChangeLeftIndent(selResource.RTF, intIncrement)
                Next
            End If

            If .SelectAssumptions = True Then
                For Each selAssumption As Assumption In selStruct.Assumptions
                    selAssumption.RTF = objRtf.ChangeLeftIndent(selAssumption.RTF, intIncrement)
                Next
            End If
        End With
    End Sub

    Private Sub ChangeParagraphLeftIndent_Children_Indicators(ByVal objRtf As RichTextManager, ByVal objTextSelection As TextSelection, ByVal selIndicators As Indicators, ByVal intIncrement As Integer)
        With objTextSelection
            For Each selIndicator As Indicator In selIndicators
                If .SelectIndicators = True Then selIndicator.RTF = objRtf.ChangeLeftIndent(selIndicator.RTF, intIncrement)
                If .SelectVerificationSources = True Then
                    For Each selVerificationSource As VerificationSource In selIndicator.VerificationSources
                        selVerificationSource.RTF = objRtf.ChangeLeftIndent(selVerificationSource.RTF, intIncrement)
                    Next
                End If
                If .SelectStatements = True Then
                    For Each selStatement As Statement In selIndicator.Statements
                        selStatement.FirstLabel = objRtf.ChangeLeftIndent(selStatement.FirstLabel, intIncrement)
                        selStatement.SecondLabel = objRtf.ChangeLeftIndent(selStatement.SecondLabel, intIncrement)
                    Next
                End If

                If selIndicator.Indicators.Count > 0 Then
                    ChangeParagraphLeftIndent_Children_Indicators(objRtf, objTextSelection, selIndicator.Indicators, intIncrement)
                End If
            Next
        End With
    End Sub
#End Region
End Class
