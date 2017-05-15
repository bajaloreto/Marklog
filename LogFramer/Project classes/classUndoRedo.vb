Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class classUndoRedo
    Friend UndoBuffer As New UndoListItem

    Private Const CONST_MaxDescriptionLength = 32
    Private intFind As Integer

    Public Enum Actions
        NotSet
        TextChanged
        TextRemoved
        TextFontChanged
        TextFontSizeChanged
        TextBold
        TextItalic
        TextUnderline
        TextStrikeOut
        TextSuperScript
        TextSubScript
        TextFrontColor
        TextBackColor
        TextCaseUpper
        TextCaseLower
        TextCaseSentence
        TextCaseFirstLetter
        TextCut
        TextPasted

        ParagraphAlignLeft
        ParagraphAlignCenter
        ParagraphAlignRight
        ParagraphLeftIndentChanged

        ValueChanged
        DoubleValueChanged
        OptionChanged
        OptionChecked
        OptionUnchecked
        BooleanValueChecked
        BooleanValueUnchecked
        AmountChanged

        DateChanged
        DateRelativeChanged
        DurationUntilEndOfProject
        DurationFromStartOfProject

        ItemInserted
        ItemRemoved
        ItemRemovedNotVertical
        ItemCut
        ItemCutNotVertical
        ItemPasted
        ItemParentInserted
        ItemParentChanged
        ItemChildInserted
        ItemMoveUp
        ItemMoveDown
        ItemSectionUp
        ItemSectionDown
        ItemLevelUp
        ItemLevelDown
    End Enum

#Region "General methods"
    Public Sub New()

    End Sub

    Private Function GetChangesInText(ByVal strOldText As String, ByVal strNewText As String) As String
        strOldText = RemovePunctuation(strOldText)
        strNewText = RemovePunctuation(strNewText)

        Dim strChanges As String = " " & strNewText.ToLower & " "
        Dim strRemoved As String = String.Empty
        Dim strOldWords() As String = strOldText.ToLower.Split(" "c)
        'Dim strNewWords() As String = strNewText.ToLower.Split(" "c)
        Dim selWord As String
        Dim intIndex As Integer

        If String.IsNullOrEmpty(strOldText) = False Then
            For i = 0 To strOldWords.Count - 1
                selWord = " " & strOldWords(i)

                intIndex = strChanges.IndexOf(selWord)
                If intIndex >= 0 And strChanges.Length > intIndex + 1 Then
                    strChanges = strChanges.Remove(intIndex + 1, selWord.Length)
                Else
                    strRemoved &= selWord & " "
                End If
            Next
        End If

        strChanges = Trim(strChanges)
        strRemoved = Trim(strRemoved)

        If strChanges.Length > CONST_MaxDescriptionLength Then strChanges = strChanges.Substring(0, CONST_MaxDescriptionLength) & "..."
        If strRemoved.Length > CONST_MaxDescriptionLength Then strRemoved = strRemoved.Substring(0, CONST_MaxDescriptionLength) & "..."

        If String.IsNullOrEmpty(strChanges) = False Then
            Return strChanges
        Else
            If UndoBuffer.ActionIndex = Actions.TextChanged Then UndoBuffer.ActionIndex = Actions.TextRemoved
            Return strRemoved
        End If
    End Function

    Private Function RemovePunctuation(ByVal strText As String) As String
        Dim strPattern As String = "[.,?!()'""?]+"
        strText = strText.ToLower

        strText = Regex.Replace(strText, strPattern, String.Empty)

        Return strText
    End Function

    Public Function GetPropertyValue(ByVal selItem As Object, ByVal strPropertyName As String) As Object
        Dim objValue As Object = Nothing

        If strPropertyName.Contains(".") Then
            objValue = GetNestedPropertyValue(selItem, strPropertyName)
        Else
            Dim objType As Type = selItem.GetType
            Dim propInfo As PropertyInfo = objType.GetProperty(strPropertyName)

            If propInfo IsNot Nothing Then
                objValue = propInfo.GetValue(selItem, Nothing)
            End If
        End If

        Return objValue
    End Function

    Public Function GetNestedPropertyValue(ByVal objValue As Object, ByVal strPropertyName As String) As Object
        Dim strElements() As String = strPropertyName.Split(".")
        Dim strElement As String
        Dim objType As Type
        Dim objPropInfo As PropertyInfo

        For Each strElement In strElements
            If objValue Is Nothing Then Return Nothing

            objType = objValue.GetType()
            objPropInfo = objType.GetProperty(strElement)
            If objPropInfo Is Nothing Then Return Nothing

            objValue = objPropInfo.GetValue(objValue, Nothing)
        Next

        Return objValue
    End Function

    Private Function SetPropertyValue(ByVal selItem As Object, ByVal strPropertyName As String, ByVal objValue As Object) As Boolean
        If selItem Is Nothing Or String.IsNullOrEmpty(strPropertyName) Then Return False

        If strPropertyName.Contains(".") Then
            Return SetNestedPropertyValue(selItem, strPropertyName, objValue)
        Else
            Dim objType As Type = selItem.GetType
            Dim propInfo As PropertyInfo = objType.GetProperty(strPropertyName)

            If propInfo IsNot Nothing Then
                propInfo.SetValue(selItem, objValue, Nothing)

                Return True
            End If
        End If

        Return False
    End Function

    Public Function SetNestedPropertyValue(ByVal selItem As Object, ByVal strPropertyName As String, ByVal objValue As Object) As Boolean
        Dim strElements() As String = strPropertyName.Split(".")
        Dim strElement As String
        Dim objType As Type
        Dim objPropInfo As PropertyInfo

        For i = 0 To strElements.Count - 1
            strElement = strElements(i)
            If selItem Is Nothing Then Return False

            objType = selItem.GetType()
            objPropInfo = objType.GetProperty(strElement)
            If objPropInfo Is Nothing Then Return False

            If i < strElements.Count - 1 Then
                selItem = objPropInfo.GetValue(selItem, Nothing)
            Else
                objPropInfo.SetValue(selItem, objValue, Nothing)
            End If
        Next

        Return True
    End Function
#End Region

#Region "UndoBuffer"
    Public Sub UndoBuffer_Initialise(ByVal selItem As Object, Optional ByVal strPropertyName As String = "")
        Dim objParentCollection As Object = CurrentLogFrame.GetParentCollection(selItem)
        Dim intIndex As Integer
        Dim objValue As Object = Nothing

        If selItem IsNot Nothing Then 'AndAlso selItem IsNot UndoBuffer.Item Then
            If objParentCollection IsNot Nothing Then intIndex = objParentCollection.IndexOf(selItem)
            
            'Debug.Print(selItem.ToString & vbTab & strPropertyName)

            UndoBuffer = New UndoListItem(objParentCollection, intIndex, selItem)

            If String.IsNullOrEmpty(strPropertyName) = False Then
                objValue = GetPropertyValue(selItem, strPropertyName)

                UndoBuffer.PropertyName = strPropertyName
                UndoBuffer.OldValue = objValue
            End If
        Else
            UndoBuffer = New UndoListItem()
        End If
    End Sub

    Public Sub UndoBuffer_Initialise(ByVal selItem As Object, ByVal strPropertyName As String, ByVal objValue As Object)
        Dim objParentCollection As Object = CurrentLogFrame.GetParentCollection(selItem)
        Dim intIndex As Integer

        If selItem IsNot Nothing Then
            If objParentCollection IsNot Nothing Then intIndex = objParentCollection.IndexOf(selItem)

            UndoBuffer = New UndoListItem(objParentCollection, intIndex, selItem)

            UndoBuffer.PropertyName = strPropertyName
            UndoBuffer.OldValue = objValue
        End If
    End Sub

    Private Function UndoBuffer_Initialise_GetParentCollection(ByVal selItem As Object) As Object
        Dim objParentCollection As Object = Nothing

        If selItem IsNot Nothing Then 'AndAlso selItem IsNot UndoBuffer.Item Then
            Select Case selItem.GetType
                Case GetType(ActivityDetail), GetType(LogFrame), GetType(AudioVisualDetail), GetType(OpenEndedDetail), GetType(ScalesDetail), GetType(ValuesDetail)
                    'do nothing
                Case Else
                    'get item's parent collection and index of item
                    objParentCollection = CurrentLogFrame.GetParentCollection(selItem)
            End Select
        End If

        Return objParentCollection
    End Function

    Public Sub UndoBuffer_SetAction(ByVal actionindex As Integer)
        Dim selItem As Object = GetCurrentItem()

        If selItem Is Nothing Then Exit Sub

        If UndoBuffer IsNot Nothing Then
            If UndoBuffer.Item Is selItem And UndoBuffer.ActionIndex = Actions.NotSet Then
                UndoBuffer.ActionIndex = actionindex
            End If
        End If
    End Sub

    Private Sub UndoBuffer_SetDescription()
        Dim strDescription As String = String.Empty

        If UndoBuffer IsNot Nothing Then
            strDescription = UndoBuffer.GetActionText

            Select Case UndoBuffer.ActionIndex
                Case Actions.TextChanged, Actions.TextRemoved, Actions.TextCut, Actions.TextPasted
                    Dim strOldText As String = String.Empty, strNewText As String = String.Empty
                    Dim strChanges As String

                    If UndoBuffer.NewValue Is Nothing Then Exit Select
                    If UndoBuffer.OldValue IsNot Nothing Then strOldText = UndoBuffer.OldValue
                    strNewText = UndoBuffer.NewValue.ToString

                    If IsRichText(strOldText) Then strOldText = RichTextToText(strOldText)
                    If IsRichText(strNewText) Then strNewText = RichTextToText(strNewText)

                    strChanges = GetChangesInText(strOldText, strNewText)

                    strDescription = UndoBuffer.GetActionText 'action index is changed by GetChangesInText if text has been removed instead of added

                    strDescription = String.Format("{0}: {1}", strDescription, strChanges)
                Case Actions.TextFontChanged, Actions.TextFontSizeChanged, _
                    Actions.TextCaseUpper, Actions.TextCaseLower, Actions.TextCaseSentence, Actions.TextCaseFirstLetter, _
                    Actions.TextBold, Actions.TextItalic, Actions.TextUnderline, Actions.TextStrikeOut, Actions.TextSubScript, Actions.TextSuperScript, _
                    Actions.TextFrontColor, Actions.TextBackColor
                    Dim strSelectedText As String = UndoBuffer.SelectedText

                    If strSelectedText.Length > CONST_MaxDescriptionLength Then strSelectedText = strSelectedText.Substring(0, CONST_MaxDescriptionLength) & "..."

                    strDescription = String.Format("{0}: {1}", strDescription, strSelectedText)
                Case Actions.ParagraphAlignLeft, Actions.ParagraphAlignCenter, Actions.ParagraphAlignRight, Actions.ParagraphLeftIndentChanged
                    'action text only
                Case Actions.ValueChanged
                    Dim dblOldValue, dblNewValue As Double

                    If UndoBuffer.NewValue Is Nothing Then Exit Select
                    If UndoBuffer.OldValue IsNot Nothing Then dblOldValue = ParseDouble(UndoBuffer.OldValue)
                    dblNewValue = ParseDouble(UndoBuffer.NewValue)

                    strDescription = String.Format("{0}: {1} --> {2}", strDescription, dblOldValue, dblNewValue)
                Case Actions.OptionChanged
                    If IsNumeric(UndoBuffer.OldValue) Then
                        strDescription = UndoBuffer_SetDescription_OptionChanged(strDescription)
                    End If
                Case Actions.OptionChecked, Actions.OptionUnchecked

                Case Actions.DateChanged
                    Dim datOldValue, datNewValue As Date
                    Dim strOldValue, strNewValue As String

                    If UndoBuffer.NewValue Is Nothing Then Exit Select
                    If UndoBuffer.OldValue IsNot Nothing Then datOldValue = UndoBuffer.OldValue
                    datNewValue = UndoBuffer.NewValue

                    strOldValue = datOldValue.ToShortDateString
                    strNewValue = datNewValue.ToShortDateString

                    strDescription = String.Format("{0}: {1} --> {2}", strDescription, strOldValue, strNewValue)
                Case Actions.ItemInserted, Actions.ItemRemoved, Actions.ItemRemovedNotVertical, Actions.ItemCut, Actions.ItemCutNotVertical, Actions.ItemPasted
                    Dim objType As Type = UndoBuffer.Item.GetType
                    Dim propInfo As PropertyInfo = objType.GetProperty("ItemName")
                    Dim strItemName As String

                    If propInfo IsNot Nothing Then
                        strItemName = propInfo.GetValue(UndoBuffer.Item, Nothing)

                        strDescription = String.Format("{0}: {1}", strDescription, strItemName)
                    End If
            End Select

            UndoBuffer.Description = strDescription
        End If
    End Sub

    Private Function UndoBuffer_SetDescription_OptionChanged(ByVal strDescription As String) As String
        Dim intOldValue, intNewValue As Integer
        Dim strOldOption As String = String.Empty, strNewOption As String = String.Empty
        Dim strPropertyName As String = GetNestedPropertyName(UndoBuffer.PropertyName)

        If UndoBuffer.NewValue Is Nothing Then Return String.Empty
        If UndoBuffer.OldValue IsNot Nothing Then intOldValue = ParseInteger(UndoBuffer.OldValue)
        intNewValue = ParseInteger(UndoBuffer.NewValue)

        Select Case UndoBuffer.Item.GetType
            Case GetType(ActivityDetail), GetType(KeyMoment)
                Select Case strPropertyName
                    Case "PeriodUnit", "DurationUnit", "PreparationUnit", "FollowUpUnit", "RepeatUnit"
                        strOldOption = LIST_DurationUnits(intOldValue)
                        strNewOption = LIST_DurationUnits(intNewValue)
                    Case "PeriodDirection"
                        strOldOption = LIST_DirectionValues(intOldValue)
                        strNewOption = LIST_DirectionValues(intNewValue)
                    Case "Type"
                        intFind = intOldValue
                        strOldOption = LIST_ActivityTypes.Find(AddressOf UndoBuffer_SetDescription_GetIdValuePairText).Value
                        intFind = intNewValue
                        strNewOption = LIST_ActivityTypes.Find(AddressOf UndoBuffer_SetDescription_GetIdValuePairText).Value
                End Select
            Case GetType(Address)
                Select Case strPropertyName
                    Case "Type"
                        strOldOption = LIST_AddressTypes(intOldValue)
                        strNewOption = LIST_AddressTypes(intNewValue)
                End Select
            Case GetType(Assumption)
                Select Case strPropertyName
                    Case "RaidType"
                        intFind = intOldValue
                        strOldOption = LIST_RaidTypes.Find(AddressOf UndoBuffer_SetDescription_GetIdValuePairText).Value
                        intFind = intNewValue
                        strNewOption = LIST_RaidTypes.Find(AddressOf UndoBuffer_SetDescription_GetIdValuePairText).Value
                    Case "ImportanceLevel"
                        strOldOption = LIST_ImportanceLevels(intOldValue)
                        strNewOption = LIST_ImportanceLevels(intNewValue)
                    Case "RiskCategory"
                        strOldOption = LIST_RiskCategories(intOldValue)
                        strNewOption = LIST_RiskCategories(intNewValue)
                    Case "Likelihood"
                        strOldOption = LIST_Likelihoods(intOldValue)
                        strNewOption = LIST_Likelihoods(intNewValue)
                    Case "RiskImpact"
                        strOldOption = LIST_RiskImpacts(intOldValue)
                        strNewOption = LIST_RiskImpacts(intNewValue)
                    Case "RiskResponse"
                        strOldOption = LIST_RiskResponses(intOldValue)
                        strNewOption = LIST_RiskResponses(intNewValue)
                End Select
            Case GetType(BudgetItem)
                Select Case strPropertyName
                    Case "Type"
                        strOldOption = LIST_BudgetItemTypes(intOldValue)
                        strNewOption = LIST_BudgetItemTypes(intNewValue)
                End Select
            Case GetType(Contact)
                Select Case strPropertyName
                    Case "Gender"
                        strOldOption = LIST_Gender(intOldValue)
                        strNewOption = LIST_Gender(intNewValue)
                End Select
            Case GetType(Email)
                Select Case strPropertyName
                    Case "Type"
                        strOldOption = LIST_EmailTypes(intOldValue)
                        strNewOption = LIST_EmailTypes(intNewValue)
                End Select
            Case GetType(Indicator)
                Select Case strPropertyName
                    Case "QuestionType"
                        strOldOption = LIST_QuestionTypes(intOldValue)
                        strNewOption = LIST_QuestionTypes(intNewValue)
                    Case "Registration"
                        strOldOption = LIST_RegistrationOptions(intOldValue)
                        strNewOption = LIST_RegistrationOptions(intNewValue)
                    Case "ScoringSystem"
                        strOldOption = LIST_ScoringSystems(intOldValue)
                        strNewOption = LIST_ScoringSystems(intNewValue)
                    Case "TargetSystem"
                        strOldOption = LIST_TargetSystems(intOldValue)
                        strNewOption = LIST_TargetSystems(intNewValue)
                    Case "TargetGroupGuid"
                        Dim ListTargetGroups As Dictionary(Of Guid, String) = CurrentLogFrame.GetTargetGroupsList

                        If UndoBuffer.OldValue IsNot Nothing Then
                            strOldOption = ListTargetGroups.Item(UndoBuffer.OldValue)
                        End If
                        If UndoBuffer.NewValue IsNot Nothing Then
                            strNewOption = ListTargetGroups.Item(UndoBuffer.NewValue)
                        End If
                    Case "AggregateHorizontal", "AggregateVertical"
                        strOldOption = LIST_AggregationOptions(intOldValue)
                        strNewOption = LIST_AggregationOptions(intNewValue)
                End Select
            Case GetType(LogFrame)
                Select Case strPropertyName
                    Case "DurationUnit"
                        strOldOption = LIST_DurationRange(intOldValue)
                        strNewOption = LIST_DurationRange(intNewValue)
                End Select
            Case GetType(OpenEndedDetail)
                Select Case strPropertyName
                    Case "WhiteSpace"
                        strOldOption = LIST_WhiteSpaceValues(intOldValue)
                        strNewOption = LIST_WhiteSpaceValues(intNewValue)
                End Select
            Case GetType(ProjectPartner)
                Select Case strPropertyName
                    Case "Type"
                        strOldOption = LIST_OrganisationTypes(intOldValue)
                        strNewOption = LIST_OrganisationTypes(intNewValue)
                    Case "Role"
                        strOldOption = LIST_ProjectPartnerRoleNames(intOldValue)
                        strNewOption = LIST_ProjectPartnerRoleNames(intNewValue)
                End Select
            Case GetType(TargetGroup)
                Select Case strPropertyName
                    Case "Type"
                        strOldOption = LIST_TargetGroupTypes(intOldValue)
                        strNewOption = LIST_TargetGroupTypes(intNewValue)
                End Select
            Case GetType(TargetGroupInformation)
                Select Case strPropertyName
                    Case "PropertyType"
                        strOldOption = LIST_TargetGroupInformationTypes(intOldValue)
                        strNewOption = LIST_TargetGroupInformationTypes(intNewValue)
                End Select
            Case GetType(TelephoneNumber)
                Select Case strPropertyName
                    Case "Type"
                        strOldOption = LIST_TelephoneNumberTypes(intOldValue)
                        strNewOption = LIST_TelephoneNumberTypes(intNewValue)
                End Select
        End Select

        strDescription = String.Format("{0}: {1} --> {2}", strDescription, strOldOption, strNewOption)

        Return strDescription
    End Function

    Private Function UndoBuffer_SetDescription_GetIdValuePairText(ByVal selItem As IdValuePair) As Boolean
        If selItem.Id = intFind Then Return True Else Return False
    End Function

    Private Sub UndoBufferToList()
        If UndoBuffer IsNot Nothing AndAlso UndoBuffer.ActionIndex <> Actions.NotSet Then
            UndoBuffer_SetDescription()

            CurrentUndoList.Insert(0, UndoBuffer)

            UndoBuffer = New UndoListItem()
            ReloadSplitUndoRedoButtons()
        End If
    End Sub

    Private Sub ReloadSplitUndoRedoButtons()
        frmParent.ReloadSplitUndoRedoButtons()
    End Sub
#End Region

#Region "Changes to current control"
    Public Sub CurrentControlChanged(ByVal strPropertyName As String)
        Dim selItem As Object = GetCurrentItem()

        If selItem IsNot Nothing Then UndoBuffer_Initialise(selItem, strPropertyName)
    End Sub

    Public Sub CurrentControlValidated(ByVal selControl As Control)
        Select Case UndoBuffer.ActionIndex
            Case Actions.TextChanged, Actions.ValueChanged
                If selControl.DataBindings.Count > 0 Then
                    Dim selDataBinding As Binding = CurrentControl.DataBindings(0)
                    Dim objBindingSource As BindingSource = selDataBinding.DataSource

                    Dim selItem As Object = objBindingSource.DataSource 'selDataBinding.DataSource
                    Dim strPropertyName As String = selDataBinding.BindingMemberInfo.BindingMember

                    If selItem Is UndoBuffer.Item And strPropertyName = UndoBuffer.PropertyName Then
                        Dim objValue As Object = GetPropertyValue(selItem, strPropertyName)

                        If objValue IsNot Nothing Then
                            UndoBuffer.NewValue = objValue

                            UndoBufferToList()
                        End If
                    End If
                End If

        End Select
    End Sub
#End Region

End Class
