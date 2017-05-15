Partial Public Class classUndoRedo

#Region "Text actions"
    Public Sub SetChangedText(ByVal NewText As String, ByVal SelectedText As String)
        UndoBuffer.NewValue = NewText
        UndoBuffer.SelectedText = SelectedText
        UndoBufferToList()
    End Sub

    Public Sub TextChanged(ByVal NewText As String)
        If UndoBuffer.OldValue IsNot Nothing And UndoBuffer.ActionIndex <= Actions.TextChanged Then
            Dim strOldText As String = Trim(UndoBuffer.OldValue.ToString)
            NewText = Trim(NewText)
            If String.Equals(strOldText, NewText) = False Then
                UndoBuffer.NewValue = NewText
                UndoBuffer.ActionIndex = Actions.TextChanged
                UndoBufferToList()
            End If
        End If
    End Sub

    Public Sub TextRemoved(ByVal NewText As String)
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.TextRemoved
        UndoBufferToList()
    End Sub

    Public Sub TextCut(ByVal NewText As String)
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.TextCut
        UndoBufferToList()
    End Sub

    Public Sub TextPasted(ByVal NewText As String)
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.TextPasted
        UndoBufferToList()
    End Sub

    Public Sub ParagraphAlignLeft(ByVal OldText As String, ByVal NewText As String)
        UndoBuffer.OldValue = OldText
        UndoBuffer.PropertyName = "RTF"
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.ParagraphAlignLeft
        UndoBufferToList()
    End Sub

    Public Sub ParagraphAlignCenter(ByVal OldText As String, ByVal NewText As String)
        UndoBuffer.OldValue = OldText
        UndoBuffer.PropertyName = "RTF"
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.ParagraphAlignCenter
        UndoBufferToList()
    End Sub

    Public Sub ParagraphAlignRight(ByVal OldText As String, ByVal NewText As String)
        UndoBuffer.OldValue = OldText
        UndoBuffer.PropertyName = "RTF"
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.ParagraphAlignRight
        UndoBufferToList()
    End Sub

    Public Sub ParagraphLeftIndentChanged(ByVal OldText As String, ByVal NewText As String)
        UndoBuffer.OldValue = OldText
        UndoBuffer.PropertyName = "RTF"
        UndoBuffer.NewValue = NewText
        UndoBuffer.ActionIndex = Actions.ParagraphLeftIndentChanged
        UndoBufferToList()
    End Sub
#End Region

#Region "Value actions"
    Public Sub ValueChanged(ByVal NewValue As Object)
        If NewValue <> UndoBuffer.OldValue Then
            UndoBuffer.NewValue = NewValue
            UndoBuffer.ActionIndex = Actions.ValueChanged
            UndoBufferToList()
        End If
    End Sub

    Public Sub ValueChanged(ByVal OldValue As Object, ByVal NewValue As Object)
        If NewValue <> OldValue Then
            UndoBuffer.OldValue = OldValue
            UndoBuffer.NewValue = NewValue
            UndoBuffer.ActionIndex = Actions.ValueChanged
            UndoBufferToList()
        End If
    End Sub

    Public Sub DoubleValueChanged(ByVal ParentItem As Object, ByVal intDoubleValueIndex As Integer, ByVal objOldValues As DoubleValues, ByVal objNewValues As DoubleValues)
        With UndoBuffer
            .Item = ParentItem
            .OldIndex = intDoubleValueIndex
            .PropertyName = "Value"
            .ActionIndex = Actions.DoubleValueChanged
            .OldValue = objOldValues
            .NewValue = objNewValues
        End With

        UndoBufferToList()
    End Sub

    Public Sub OptionChanged(ByVal NewOption As Object)
        UndoBuffer.NewValue = NewOption
        UndoBuffer.ActionIndex = Actions.OptionChanged
        UndoBufferToList()
    End Sub

    Public Sub OptionChanged(ByVal selItem As Object, ByVal strPropertyName As String, ByVal OldOption As Object, ByVal NewOption As Object)
        UndoBuffer.Item = selItem
        UndoBuffer.PropertyName = strPropertyName
        UndoBuffer.OldValue = OldOption
        UndoBuffer.NewValue = NewOption
        UndoBuffer.ActionIndex = Actions.OptionChanged
        UndoBufferToList()
    End Sub

    Public Sub OptionChecked(ByVal boolNewValue As Boolean)
        Dim boolOldValue As Boolean
        Dim intActionIndex As Integer

        If boolNewValue = True Then
            boolOldValue = False
            intActionIndex = Actions.OptionChecked
        Else
            boolOldValue = True
            intActionIndex = Actions.OptionUnchecked
        End If

        With UndoBuffer
            .ActionIndex = intActionIndex
            .OldItem = .Item
            .OldValue = boolOldValue
            .NewValue = boolNewValue
        End With

        UndoBufferToList()
    End Sub

    Public Sub BooleanValueChecked(ByVal ParentItem As Object, ByVal intBooleanValueIndex As Integer, ByVal objOldValues As BooleanValues, ByVal objNewValues As BooleanValues)
        With UndoBuffer
            .Item = ParentItem
            .OldIndex = intBooleanValueIndex
            .PropertyName = "Value"
            .ActionIndex = Actions.BooleanValueChecked
            .OldValue = objOldValues
            .NewValue = objNewValues
        End With

        UndoBufferToList()
    End Sub

    Public Sub AmountChanged(ByVal NewValue As Currency)
        If NewValue IsNot UndoBuffer.OldValue Then
            UndoBuffer.NewValue = NewValue
            UndoBuffer.ActionIndex = Actions.AmountChanged
            UndoBufferToList()
        End If
    End Sub
#End Region

#Region "Date actions"
    Public Sub DateChanged(ByVal NewDate As Date)
        UndoBuffer.NewValue = NewDate
        UndoBuffer.ActionIndex = Actions.DateChanged
        UndoBufferToList()
    End Sub

    Public Sub DateRelativeChanged(ByVal selItem As Object, ByVal strPropertyName As String, ByVal objOldValue As Object, ByVal objNewValue As Object)
        With UndoBuffer
            .ActionIndex = Actions.DateRelativeChanged
            .Item = selItem
            .OldItem = .Item
            .PropertyName = strPropertyName
            .OldValue = objOldValue
            .NewValue = objNewValue
        End With

        UndoBufferToList()
    End Sub

    Public Sub DurationUntilEndOfProject(ByVal boolNewValue As Boolean)
        Dim boolOldValue As Boolean

        If boolNewValue = True Then boolOldValue = False Else boolOldValue = True

        With UndoBuffer
            .ActionIndex = Actions.DurationUntilEndOfProject
            .OldItem = .Item
            .OldValue = boolOldValue
            .NewValue = boolNewValue
        End With

        UndoBufferToList()
    End Sub

    Public Sub DurationFromStartOfProject(ByVal boolNewValue As Boolean)
        Dim boolOldValue As Boolean

        If boolNewValue = True Then boolOldValue = False Else boolOldValue = True

        With UndoBuffer
            .ActionIndex = Actions.DurationFromStartOfProject
            .OldItem = .Item
            .OldValue = boolOldValue
            .NewValue = boolNewValue
        End With

        UndoBufferToList()
    End Sub
#End Region

#Region "Item actions"
    Public Sub ItemInserted(ByVal NewItem As Object, ByVal colParent As Object)
        Dim intIndex As Integer = colParent.IndexOf(NewItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemInserted
            .Item = NewItem
            .OldIndex = -1
            .NewIndex = intIndex
            .OldParentItem = Nothing
            .NewParentItem = colParent
        End With

        UndoBufferToList()
        VerifyCreationBeforeModification(NewItem)
    End Sub

    Private Sub VerifyCreationBeforeModification(ByVal selItem As Object)
        'if the item is created via a dialog window, its properties may change before the item is actually commited to the collection
        Dim selUndoListItem As UndoListItem
        Dim intLastIndex As Integer
        If CurrentUndoList.Count > 1 Then
            selUndoListItem = CurrentUndoList(0)
            If selUndoListItem.Item Is selItem And selUndoListItem.ActionIndex = Actions.ItemInserted Then
                For i = 1 To CurrentUndoList.Count - 1
                    If CurrentUndoList(i).Item Is selItem Then
                        intLastIndex = i
                    End If
                Next
            End If

            If intLastIndex > 0 Then
                CurrentUndoList.Remove(selUndoListItem)
                CurrentUndoList.Insert(intLastIndex, selUndoListItem)
                ReloadSplitUndoRedoButtons()
            End If
        End If
    End Sub

    Public Sub ItemRemoved(ByVal RemovedItem As Object, ByVal colParent As Object)
        Dim intIndex As Integer = colParent.IndexOf(RemovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemRemoved
            .Item = RemovedItem
            .OldIndex = intIndex
            .NewIndex = -1
            .OldParentItem = colParent
            .NewParentItem = Nothing
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemRemovedNotVertical(ByVal RemovedItem As Object, ByVal colParent As Object)
        Dim intIndex As Integer = colParent.IndexOf(RemovedItem)
        Dim objCopy As Object

        Using copier As New ObjectCopy
            objCopy = copier.CopyObject(RemovedItem)
        End Using

        With UndoBuffer
            .ActionIndex = Actions.ItemRemovedNotVertical
            .Item = objCopy
            .OldIndex = intIndex
            .NewIndex = -1
            .OldParentItem = colParent
            .NewParentItem = Nothing
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemCut(ByVal CutItem As Object, ByVal colParent As Object)
        Dim intIndex As Integer = colParent.IndexOf(CutItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemCut
            .Item = CutItem
            .OldIndex = intIndex
            .NewIndex = -1
            .OldParentItem = colParent
            .NewParentItem = Nothing
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemCutNotVertical(ByVal CutItem As Object, ByVal colParent As Object)
        Dim intIndex As Integer = colParent.IndexOf(CutItem)
        Dim objCopy As Object

        Using copier As New ObjectCopy
            objCopy = copier.CopyObject(CutItem)
        End Using

        With UndoBuffer
            .ActionIndex = Actions.ItemCutNotVertical
            .Item = objCopy
            .OldIndex = intIndex
            .NewIndex = -1
            .OldParentItem = colParent
            .NewParentItem = Nothing
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemPasted(ByVal PastedItem As Object, ByVal colParent As Object)
        Dim intIndex As Integer = colParent.IndexOf(PastedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemPasted
            .Item = PastedItem
            .OldIndex = -1
            .NewIndex = intIndex
            .OldParentItem = Nothing
            .NewParentItem = colParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemParentChanged(ByVal MovedItem As Object, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.Count

        With UndoBuffer
            .ActionIndex = Actions.ItemParentChanged
            .Item = MovedItem
            .NewIndex = intNewIndex
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemParentInserted(ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemParentInserted
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemChildInserted(ByVal InsertedItem As Object, ByVal colParent As Object)
        Dim intNewIndex As Integer = colParent.IndexOf(InsertedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemChildInserted
            .Item = InsertedItem
            .OldIndex = -1
            .NewIndex = intNewIndex
            .OldParentItem = Nothing
            .NewParentItem = colParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemMovedUp(ByVal OldItem As Object, ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemMoveUp
            .OldItem = OldItem
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemMovedDown(ByVal OldItem As Object, ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemMoveDown
            .OldItem = OldItem
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemSectionUp(ByVal OldItem As Object, ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemMoveUp
            .OldItem = OldItem
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemSectionDown(ByVal OldItem As Object, ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemMoveDown
            .OldItem = OldItem
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemLevelUp(ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemLevelUp
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub

    Public Sub ItemLevelDown(ByVal MovedItem As Object, ByVal colOldParent As Object, ByVal intOldIndex As Integer, ByVal colNewParent As Object)
        Dim intNewIndex As Integer = colNewParent.IndexOf(MovedItem)

        With UndoBuffer
            .ActionIndex = Actions.ItemLevelDown
            .Item = MovedItem
            .OldIndex = intOldIndex
            .NewIndex = intNewIndex
            .OldParentItem = colOldParent
            .NewParentItem = colNewParent
        End With

        UndoBufferToList()
    End Sub
#End Region

End Class
