Partial Public Class classUndoRedo
#Region "Undo actions"
    Public Sub Undo(ByVal UndoItemIndex As Integer)
        If UndoItemIndex >= 0 And UndoItemIndex <= CurrentUndoList.Count - 1 Then
            Dim intItems As Integer = UndoItemIndex

            Do
                Dim selUndoListItem As UndoListItem = CurrentUndoList(0)

                Select Case selUndoListItem.ActionIndex
                    Case Actions.TextChanged, Actions.TextRemoved, Actions.TextCut, Actions.TextPasted, _
                        Actions.TextFontChanged, Actions.TextFontSizeChanged, _
                        Actions.TextCaseUpper, Actions.TextCaseLower, Actions.TextCaseSentence, Actions.TextCaseFirstLetter, _
                        Actions.TextBold, Actions.TextItalic, Actions.TextUnderline, Actions.TextStrikeOut, Actions.TextSubScript, Actions.TextSuperScript, _
                        Actions.TextFrontColor, Actions.TextBackColor, _
                        Actions.ParagraphAlignLeft, Actions.ParagraphAlignCenter, Actions.ParagraphAlignRight, Actions.ParagraphLeftIndentChanged, _
                        Actions.ValueChanged, Actions.OptionChanged, Actions.OptionChecked, Actions.OptionUnchecked, _
                        Actions.AmountChanged, _
                        Actions.DateChanged, Actions.DateRelativeChanged, Actions.DurationFromStartOfProject, Actions.DurationUntilEndOfProject

                        Undo_Values(selUndoListItem)

                    Case Actions.BooleanValueChecked, Actions.BooleanValueUnchecked
                        Undo_BooleanValues(selUndoListItem)

                    Case Actions.DoubleValueChanged
                        Undo_DoubleValues(selUndoListItem)

                    Case Actions.ItemInserted, Actions.ItemPasted, Actions.ItemChildInserted
                        Undo_ItemInsert(selUndoListItem)

                    Case Actions.ItemRemoved, Actions.ItemCut
                        Undo_ItemRemove(selUndoListItem)

                    Case Actions.ItemRemovedNotVertical, Actions.ItemCutNotVertical
                        Undo_ItemRemoveNotVertical(selUndoListItem)

                    Case Actions.ItemLevelDown, Actions.ItemParentInserted
                        Undo_ItemLevelDown(selUndoListItem)

                    Case Actions.ItemLevelUp, Actions.ItemParentChanged
                        Undo_ItemLevelUp(selUndoListItem)

                    Case Actions.ItemMoveDown, Actions.ItemMoveUp, Actions.ItemSectionDown, Actions.ItemSectionUp
                        Undo_ItemMove(selUndoListItem)

                End Select

                CurrentRedoList.Insert(0, selUndoListItem)

                CurrentUndoList.RemoveAt(0)
                If CurrentUndoList.Count = 0 Then Exit Do
                intItems -= 1
            Loop While intItems >= 0

        End If
    End Sub

    Private Sub Undo_Values(ByVal selUndoListItem As UndoListItem)
        If String.IsNullOrEmpty(selUndoListItem.PropertyName) = False Then
            If SetPropertyValue(selUndoListItem.Item, selUndoListItem.PropertyName, selUndoListItem.OldValue) = True Then

                Select Case selUndoListItem.ActionIndex
                    Case Actions.TextChanged, Actions.ValueChanged, Actions.AmountChanged
                        Undo_Recalculate(selUndoListItem)
                End Select

                Reload(selUndoListItem)
            End If
        End If
    End Sub

    Private Sub Undo_BooleanValues(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            Dim objBooleanValues As BooleanValues = CType(.OldValue, BooleanValues)

            Select Case .Item.GetType
                Case GetType(Baseline)
                    CType(.Item, Baseline).BooleanValues = objBooleanValues
                Case GetType(Target)
                    CType(.Item, Target).BooleanValues = objBooleanValues
                Case GetType(BooleanValuesMatrixRow)
                    CType(.Item, BooleanValuesMatrixRow).BooleanValues = objBooleanValues
            End Select

            Reload(selUndoListItem)
        End With
    End Sub

    Private Sub Undo_DoubleValues(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            Dim objDoubleValues As DoubleValues = CType(.OldValue, DoubleValues)

            Select Case .Item.GetType
                Case GetType(Baseline)
                    CType(.Item, Baseline).DoubleValues = objDoubleValues
                Case GetType(Target)
                    CType(.Item, Target).DoubleValues = objDoubleValues
                Case GetType(DoubleValuesMatrixRow)
                    CType(.Item, DoubleValuesMatrixRow).DoubleValues = objDoubleValues
            End Select

            Reload(selUndoListItem)
        End With
    End Sub

    Private Sub Undo_ItemInsert(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            .NewParentItem.Remove(.Item)
            Undo_RecalculateChildTotals(.NewParentItem)
            Undo_RemoveReferrers(.Item)
        End With

        Reload(selUndoListItem)
    End Sub

    Private Sub Undo_ItemRemove(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            .OldParentItem.Insert(.OldIndex, .Item)
            Undo_AddReferrers(.Item)
            Undo_RecalculateChildTotals(.OldParentItem)
        End With

        Reload(selUndoListItem)
    End Sub

    Private Sub Undo_ItemRemoveNotVertical(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            Dim objParentStructs As Structs = .OldParentItem
            Dim objStruct As Struct = .OldParentItem(.OldIndex)
            Dim selItem As Struct = .Item
            objStruct.RTF = selItem.RTF
            objStruct.Indicators = selItem.Indicators
            objStruct.Assumptions = selItem.Assumptions
            If (objStruct.GetType) Is GetType(Activity) Then CType(objStruct, Activity).Resources = CType(selItem, Activity).Resources
            If (objStruct.GetType) Is GetType(Purpose) Then CType(objStruct, Purpose).TargetGroups = CType(selItem, Purpose).TargetGroups
            If (objStruct.GetType) Is GetType(Output) Then CType(objStruct, Output).KeyMoments = CType(selItem, Output).KeyMoments

            Undo_RecalculateChildTotals(.OldParentItem)
        End With

        Reload(selUndoListItem)
    End Sub

    Private Sub Undo_ItemLevelDown(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            Undo_LevelUpReferrers(.Item)

            .NewParentItem.Remove(.Item)
            Undo_RecalculateChildTotals(.NewParentItem)

            .OldParentItem.Insert(.OldIndex, .Item)
            Undo_RecalculateParentTotals(.Item)
        End With

        Reload(selUndoListItem)
    End Sub

    Private Sub Undo_ItemLevelUp(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            .NewParentItem.Remove(.Item)
            Undo_RecalculateChildTotals(.NewParentItem)

            .OldParentItem.Insert(.OldIndex, .Item)
            Undo_RecalculateParentTotals(.Item)

            Undo_LevelDownReferrers(.Item)
        End With

        Reload(selUndoListItem)
    End Sub

    Private Sub Undo_ItemMove(ByVal selUndoListItem As UndoListItem)
        With selUndoListItem
            .NewParentItem.Remove(.Item)
            Undo_RecalculateChildTotals(.NewParentItem)

            Dim objExistingItem As LogframeObject = .OldParentItem(.OldIndex)

            'Necessary for section change (change from one type of struct to another)
            If objExistingItem Is Nothing Then
                .OldParentItem.Insert(.OldIndex, .OldItem)
            Else
                If objExistingItem.Guid <> CType(.OldItem, LogframeObject).Guid Then
                    .OldParentItem.Insert(.OldIndex, .OldItem)
                Else
                    'Purpose/output that has been moved down but still has a placeholder
                    objExistingItem = .OldItem
                End If
            End If

            'Necessary for budget items
            Undo_RecalculateParentTotals(.OldItem)
            Undo_MoveReferrers(.OldItem)
        End With

        Reload(selUndoListItem)
    End Sub
#End Region

#Region "Redo"
    Public Sub Redo(ByVal RedoItemIndex As Integer)
        If RedoItemIndex >= 0 And RedoItemIndex <= CurrentRedoList.Count - 1 Then
            Dim intItems As Integer = RedoItemIndex

            Do
                Dim selRedoListItem As UndoListItem = CurrentRedoList(0)

                Select Case selRedoListItem.ActionIndex
                    Case Actions.TextChanged, Actions.TextRemoved, Actions.TextCut, Actions.TextPasted, _
                        Actions.TextFontChanged, Actions.TextFontSizeChanged, _
                        Actions.TextCaseUpper, Actions.TextCaseLower, Actions.TextCaseSentence, Actions.TextCaseFirstLetter, _
                        Actions.TextBold, Actions.TextItalic, Actions.TextUnderline, Actions.TextStrikeOut, Actions.TextSubScript, Actions.TextSuperScript, _
                        Actions.TextFrontColor, Actions.TextBackColor, _
                        Actions.ParagraphAlignLeft, Actions.ParagraphAlignCenter, Actions.ParagraphAlignRight, Actions.ParagraphLeftIndentChanged, _
                        Actions.ValueChanged, Actions.OptionChanged, Actions.OptionChecked, Actions.OptionUnchecked, _
                        Actions.AmountChanged, _
                        Actions.DateChanged, Actions.DateRelativeChanged, Actions.DurationFromStartOfProject, Actions.DurationUntilEndOfProject

                        Redo_Values(selRedoListItem)

                    Case Actions.BooleanValueChecked, Actions.BooleanValueUnchecked
                        Redo_BooleanValues(selRedoListItem)

                    Case Actions.DoubleValueChanged
                        Redo_DoubleValues(selRedoListItem)

                    Case Actions.ItemInserted, Actions.ItemPasted, Actions.ItemChildInserted
                        Redo_ItemInsert(selRedoListItem)

                    Case Actions.ItemRemoved, Actions.ItemCut
                        Redo_ItemRemove(selRedoListItem)

                    Case Actions.ItemRemovedNotVertical, Actions.ItemCutNotVertical
                        Redo_ItemRemoveNotVertical(selRedoListItem)

                    Case Actions.ItemLevelDown, Actions.ItemParentInserted
                        Redo_ItemLevelDown(selRedoListItem)

                    Case Actions.ItemLevelUp, Actions.ItemParentChanged
                        Redo_ItemLevelUp(selRedoListItem)

                    Case Actions.ItemMoveDown, Actions.ItemMoveUp, Actions.ItemSectionDown, Actions.ItemSectionUp
                        Redo_ItemMove(selRedoListItem)

                End Select

                CurrentUndoList.Insert(0, selRedoListItem)

                CurrentRedoList.RemoveAt(0)
                If CurrentRedoList.Count = 0 Then Exit Do
                intItems -= 1
            Loop While intItems >= 0

        End If
    End Sub

    Private Sub Redo_Values(ByVal selRedoListItem As UndoListItem)
        If String.IsNullOrEmpty(selRedoListItem.PropertyName) = False Then
            If SetPropertyValue(selRedoListItem.Item, selRedoListItem.PropertyName, selRedoListItem.NewValue) = True Then

                Select Case selRedoListItem.ActionIndex
                    Case Actions.TextChanged, Actions.ValueChanged, Actions.AmountChanged
                        Undo_Recalculate(selRedoListItem)
                End Select

                Reload(selRedoListItem)
            End If
        End If
    End Sub

    Private Sub Redo_BooleanValues(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            Dim objBooleanValues As BooleanValues = CType(.NewValue, BooleanValues)

            Select Case .Item.GetType
                Case GetType(Baseline)
                    CType(.Item, Baseline).BooleanValues = objBooleanValues
                Case GetType(Target)
                    CType(.Item, Target).BooleanValues = objBooleanValues
                Case GetType(BooleanValuesMatrixRow)
                    CType(.Item, BooleanValuesMatrixRow).BooleanValues = objBooleanValues
            End Select

            Reload(selRedoListItem)
        End With
    End Sub

    Private Sub Redo_DoubleValues(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            Dim objDoubleValues As DoubleValues = CType(.NewValue, DoubleValues)

            Select Case .Item.GetType
                Case GetType(Baseline)
                    CType(.Item, Baseline).DoubleValues = objDoubleValues
                Case GetType(Target)
                    CType(.Item, Target).DoubleValues = objDoubleValues
                Case GetType(DoubleValuesMatrixRow)
                    CType(.Item, DoubleValuesMatrixRow).DoubleValues = objDoubleValues
            End Select

            Reload(selRedoListItem)
        End With
    End Sub

    Private Sub Redo_ItemInsert(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            .NewParentItem.Insert(.NewIndex, .Item)
            Undo_RecalculateChildTotals(.NewParentItem)
            Undo_RemoveReferrers(.Item)
        End With

        Reload(selRedoListItem)
    End Sub

    Private Sub Redo_ItemRemove(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            .OldParentItem.Remove(.Item)
            Undo_AddReferrers(.Item)
            Undo_RecalculateChildTotals(.OldParentItem)
        End With

        Reload(selRedoListItem)
    End Sub

    Private Sub Redo_ItemRemoveNotVertical(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            Dim objParentStructs As Structs = .OldParentItem
            Dim objStruct As Struct = .OldParentItem(.OldIndex)
            Dim selItem As Struct = .Item

            objStruct.RTF = String.Empty
            objStruct.Indicators.Clear()
            objStruct.Assumptions.Clear()
            If (objStruct.GetType) Is GetType(Activity) Then CType(objStruct, Activity).Resources.Clear()
            If (objStruct.GetType) Is GetType(Purpose) Then CType(objStruct, Purpose).TargetGroups.Clear()
            If (objStruct.GetType) Is GetType(Output) Then CType(objStruct, Output).KeyMoments.Clear()

            Undo_RecalculateChildTotals(.OldParentItem)
        End With

        Reload(selRedoListItem)
    End Sub

    Private Sub Redo_ItemLevelDown(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            .NewParentItem.Insert(0, .Item)
            Undo_RecalculateChildTotals(.NewParentItem)

            .OldParentItem.remove(.Item)
            Undo_RecalculateParentTotals(.Item)

            Undo_LevelDownReferrers(.Item)
        End With

        Reload(selRedoListItem)
    End Sub

    Private Sub Redo_ItemLevelUp(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            Undo_LevelUpReferrers(.Item)

            .NewParentItem.Insert(.NewIndex, .Item)
            Undo_RecalculateChildTotals(.NewParentItem)

            .OldParentItem.Remove(.Item)
            Undo_RecalculateParentTotals(.Item)
        End With

        Reload(selRedoListItem)
    End Sub

    Private Sub Redo_ItemMove(ByVal selRedoListItem As UndoListItem)
        With selRedoListItem
            .OldParentItem.Remove(.OldItem)
            Undo_RecalculateChildTotals(.OldParentItem)

            .NewParentItem.Insert(.NewIndex, .Item)
            Undo_RecalculateChildTotals(.NewParentItem)
            Undo_MoveReferrers(.Item)
        End With

        Reload(selRedoListItem)
    End Sub
#End Region

#Region "Recalculate totals"
    Private Sub Undo_Recalculate(ByVal selUndoListItem As UndoListItem)
        Select Case selUndoListItem.Item.GetType
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = CType(selUndoListItem.Item, BudgetItem)

                Select Case selUndoListItem.PropertyName
                    Case "Duration", "Number", "UnitCost", "CurrencyCode", "ExchangeRate"
                        If selBudgetItem IsNot Nothing Then
                            selBudgetItem.SetTotalCost()
                        End If

                        With CurrentBudget
                            .UpdateParentTotals(selBudgetItem)
                            .UpdateMultiYearBudget()
                        End With
                    Case "CalculationGuidRatio", "Ratio"
                        If selBudgetItem IsNot Nothing Then
                            selBudgetItem.CalculateRatio()
                        End If

                        With CurrentBudget
                            .UpdateParentTotals(selBudgetItem)
                            .ChangeRatioBudgetHeader(selBudgetItem)
                            .UpdateMultiYearBudget()
                        End With
                End Select
        End Select
    End Sub

    Private Sub Undo_RecalculateChildTotals(ByVal selParentItem As Object)
        Select Case selParentItem.GetType
            Case GetType(BudgetItem)
                With CurrentBudget
                    .UpdateChildTotals(selParentItem)
                    .UpdateMultiYearBudget()
                End With
        End Select

    End Sub

    Private Sub Undo_RecalculateParentTotals(ByVal selChildItem As Object)
        Select Case selChildItem.GetType
            Case GetType(BudgetItem)
                With CurrentBudget
                    .UpdateParentTotals(selChildItem)
                    .UpdateMultiYearBudget()
                End With
        End Select

    End Sub
#End Region

#Region "Budget - items referring to Totals page"
    Private Sub Undo_AddReferrers(ByVal selItem As Object)
        Select Case selItem.GetType
            Case GetType(BudgetItem)
                With CurrentBudget
                    If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                        .InsertBudgetHeader(CType(selItem, BudgetItem))
                    End If
                End With
        End Select
    End Sub

    Private Sub Undo_RemoveReferrers(ByVal selItem As Object)
        Select Case selItem.GetType
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = CType(selItem, BudgetItem)
                Dim objGuid As Guid

                If selBudgetItem.ParentBudgetItemGuid <> Guid.Empty Then
                    objGuid = selBudgetItem.ParentBudgetItemGuid
                Else
                    objGuid = selBudgetItem.ParentBudgetYearGuid
                End If

                With CurrentBudget
                    If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                        .RemoveBudgetHeader(selBudgetItem.Guid, objGuid)
                    End If
                End With
        End Select
    End Sub

    Private Sub Undo_MoveReferrers(ByVal selItem As Object)
        Select Case selItem.GetType
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = CType(selItem, BudgetItem)

                With CurrentBudget
                    If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                        .MoveBudgetHeader(selBudgetItem)
                    End If
                End With
        End Select
    End Sub

    Private Sub Undo_LevelUpReferrers(ByVal selItem As Object)
        Select Case selItem.GetType
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = CType(selItem, BudgetItem)

                With CurrentBudget
                    If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                        .LevelUpHeader(selBudgetItem)
                    End If
                End With
        End Select
    End Sub

    Private Sub Undo_LevelDownReferrers(ByVal selItem As Object)
        Select Case selItem.GetType
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = CType(selItem, BudgetItem)

                With CurrentBudget
                    If .MultiYearBudget = True And .BudgetYears.Count > 0 Then
                        .LevelDownHeader(selBudgetItem)
                    End If
                End With
        End Select
    End Sub
#End Region



End Class
