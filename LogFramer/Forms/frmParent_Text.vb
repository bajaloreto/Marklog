Partial Public Class frmParent
    Private Sub RibbonTabText_ActiveChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonTabText.ActiveChanged
        CurrentText.InitializeFont(My.Settings.setDefaultFont)
        SetClipboardType(DataGridViewClipboard.ContentTypes.Text)
    End Sub

#Region "Ribbon-Text-Clipboard"
    Private Sub RibbonButtonCutText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonCutText.Click
        CutText()
    End Sub

    Private Sub RibbonButtonCopyText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonCopyText.Click
        CopyText()
    End Sub

    Private Sub RibbonButtonPasteText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteText.Click
        PasteText(0)
    End Sub

    Private Sub RibbonButtonPasteKeepFormatting_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteKeepFormatting.Click
        PasteText(0, ClipboardTextItems.PasteOptions.KeepFormatting)
    End Sub

    Private Sub RibbonButtonPasteMergeFormatting_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteMergeFormatting.Click
        PasteText(0, ClipboardTextItems.PasteOptions.MergeFormatting)
    End Sub

    Private Sub RibbonButtonPasteTextOnly_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonPasteTextOnly.Click
        PasteText(0, ClipboardTextItems.PasteOptions.TextOnly)
    End Sub

    Public Sub CutText()
        Select Case CurrentControl.GetType
            Case GetType(DataGridViewLogframe), GetType(DataGridViewPlanning), GetType(DataGridViewBudgetYear), _
                GetType(DataGridViewStatementsFormula), GetType(DataGridViewStatementsMaxDiff), GetType(DataGridViewStatementsScales), _
                GetType(DataGridViewTargetsFormula), GetType(DataGridViewTargetsFrequencyLikert), GetType(DataGridViewTargetsMaxDiff), _
                GetType(DataGridViewTargetsScaleLikert), GetType(DataGridViewTargetsScales), GetType(DataGridViewTargetsSemanticDiff)

                If CurrentEditingControl IsNot Nothing Then
                    If CurrentEditingControl.SelectionLength > 0 Then
                        TextClipboard.CutRichText(CurrentEditingControl.SelectedText, CurrentEditingControl.SelectedRtf)
                        CurrentEditingControl.Cut()

                        UndoRedo.TextCut(CurrentEditingControl.Rtf)
                    End If
                ElseIf CurrentDataGridView IsNot Nothing Then
                    With CurrentDataGridView
                        If .CurrentCell.Selected = True Then
                            Dim strCellValue As String = .CurrentCell.Value.ToString()
                            Clipboard.SetText(strCellValue)

                            Select Case .CurrentCell.GetType
                                Case GetType(RichTextCellLogframe)
                                    TextClipboard.CutRichText(strCellValue)
                                Case Else
                                    TextClipboard.CutText(strCellValue)
                            End Select

                            UndoRedo.TextCut(strCellValue)

                            'erase text/value when cutting
                            If .CurrentCell.ValueType Is GetType(String) Then
                                .CurrentCell.Value = String.Empty
                            Else
                                .CurrentCell.Value = 0
                            End If
                        End If
                    End With
                End If
            Case Else
                If CurrentControl IsNot Nothing Then
                    Select Case CurrentControl.GetType
                        Case GetType(TextBox), GetType(TextBoxLF), GetType(NumericTextBoxLF), GetType(NumericBoundTextBoxLF)
                            Dim selTextBox As TextBox = CType(CurrentControl, TextBox)
                            selTextBox.Cut()

                            UndoRedo.TextCut(selTextBox.Text)
                        Case GetType(ComboBoxText)
                            Clipboard.SetText(CType(CurrentControl, ComboBoxText).SelectedText)
                    End Select
                    Dim strCut As String = Clipboard.GetText()
                    TextClipboard.CutText(strCut)

                End If
        End Select
    End Sub

    Public Sub CopyText()
        Select Case CurrentControl.GetType
            Case GetType(DataGridViewLogframe), GetType(DataGridViewPlanning), GetType(DataGridViewBudgetYear), _
                GetType(DataGridViewStatementsFormula), GetType(DataGridViewStatementsMaxDiff), GetType(DataGridViewStatementsScales), _
                GetType(DataGridViewTargetsFormula), GetType(DataGridViewTargetsFrequencyLikert), GetType(DataGridViewTargetsMaxDiff), _
                GetType(DataGridViewTargetsScaleLikert), GetType(DataGridViewTargetsScales), GetType(DataGridViewTargetsSemanticDiff)

                If CurrentEditingControl IsNot Nothing Then
                    If CurrentEditingControl.SelectionLength > 0 Then
                        TextClipboard.CopyRichText(CurrentEditingControl.SelectedText, CurrentEditingControl.SelectedRtf)
                        CurrentEditingControl.Copy()
                    End If
                ElseIf CurrentDataGridView IsNot Nothing Then
                    With CurrentDataGridView
                        If .CurrentCell.Selected = True Then
                            Dim strCellValue As String = .CurrentCell.Value.ToString()
                            Clipboard.SetText(strCellValue)

                            Select Case .CurrentCell.GetType
                                Case GetType(RichTextCellLogframe)
                                    TextClipboard.CopyRichText(strCellValue)
                                Case Else
                                    TextClipboard.CopyText(strCellValue)
                            End Select
                        End If
                    End With
                End If
            Case Else
                If CurrentControl IsNot Nothing Then
                    Select Case CurrentControl.GetType
                        Case GetType(TextBox), GetType(TextBoxLF), GetType(NumericTextBoxLF), GetType(NumericBoundTextBoxLF)
                            CType(CurrentControl, TextBox).Copy()
                        Case GetType(ComboBoxText)
                            Clipboard.SetText(CType(CurrentControl, ComboBoxText).SelectedText)
                    End Select
                    Dim strCopy As String = Clipboard.GetText()
                    TextClipboard.CopyText(strCopy)
                End If
        End Select
    End Sub

    Public Sub PasteText(ByVal intIndex As Integer, Optional ByVal intPasteOption As Integer = 0)
        Dim selClipboardItem As ClipboardTextItem = TextClipboard.Item(intIndex)

        Dim strText As String = selClipboardItem.Text
        Dim strRtf As String = selClipboardItem.RTF

        Select Case CurrentControl.GetType
            Case GetType(DataGridViewLogframe), GetType(DataGridViewPlanning), GetType(DataGridViewBudgetYear), _
                GetType(DataGridViewStatementsFormula), GetType(DataGridViewStatementsMaxDiff), GetType(DataGridViewStatementsScales), _
                GetType(DataGridViewTargetsFormula), GetType(DataGridViewTargetsFrequencyLikert), GetType(DataGridViewTargetsMaxDiff), _
                GetType(DataGridViewTargetsScaleLikert), GetType(DataGridViewTargetsScales), GetType(DataGridViewTargetsSemanticDiff)

                If CurrentDataGridView.EditingControl IsNot Nothing Then
                    PasteTextInActiveDatagridviewCell(intPasteOption, strText, strRtf)
                ElseIf CurrentDataGridView IsNot Nothing Then
                    PasteTextInSelectedDatagridviewCell(intPasteOption, strText, strRtf)
                End If
            Case Else
                PastTextInControl(strText)
        End Select
    End Sub

    Private Sub PasteTextInActiveDatagridviewCell(ByVal intPasteOption As Integer, ByVal strText As String, ByVal strRtf As String)
        If CurrentEditingControl IsNot Nothing Then
            With CurrentEditingControl
                .Select(CurrentText.SelectionStart, CurrentText.SelectionLength)
                If String.IsNullOrEmpty(strRtf) Then
                    .SelectedText = strText
                Else
                    Select Case intPasteOption
                        Case ClipboardTextItems.PasteOptions.KeepFormatting
                            .SelectedRtf = strRtf
                        Case ClipboardTextItems.PasteOptions.MergeFormatting
                            Dim selFont As Font = .SelectionFont
                            Dim selStyle As FontStyle
                            Using objRichTextManager As New RichTextManager
                                With objRichTextManager
                                    .Rtf = strRtf
                                    For i = 0 To .TextLength - 1
                                        .Select(i, 1)
                                        selStyle = .SelectionFont.Style
                                        .SelectionFont = New Font(selFont.FontFamily, selFont.Size, selStyle)
                                    Next
                                    strRtf = .Rtf
                                End With
                            End Using
                            .SelectedRtf = strRtf
                        Case ClipboardTextItems.PasteOptions.TextOnly
                            .SelectedText = strText
                    End Select

                End If

                UndoRedo.TextPasted(CurrentEditingControl.Rtf)
            End With
        End If
    End Sub

    Private Sub PasteTextInSelectedDatagridviewCell(ByVal intPasteOption As Integer, ByVal strText As String, ByVal strRtf As String)
        With CurrentDataGridView
            .BeginEdit(True)

            'CurrentEditingControl.Clear()
            PasteTextInActiveDatagridviewCell(intPasteOption, strText, strRtf)

            .EndEdit()
        End With
    End Sub

    Private Sub PastTextInControl(ByVal strText As String)
        If CurrentControl IsNot Nothing Then
            Select Case CurrentControl.GetType
                Case GetType(TextBox), GetType(TextBoxLF), GetType(NumericTextBoxLF), GetType(NumericBoundTextBoxLF)
                    Dim selTextBox As TextBox = CType(CurrentControl, TextBox)
                    selTextBox.Paste(strText)

                    UndoRedo.TextPasted(selTextBox.Text)
                Case GetType(ComboBoxText)
                    CType(CurrentControl, ComboBoxText).SelectedText = strText
            End Select
        End If
    End Sub

    Private Sub RibbonPanelClipboardText_ButtonMoreClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonPanelClipboardText.ButtonMoreClick
        ClipboardShowHide(DataGridViewClipboard.ContentTypes.Text)
    End Sub

    Private Sub SetClipboardType(ByVal intContentType As Integer)
        CurrentClipboardType = intContentType

        For Each selForm As frmProject In Me.MdiChildren
            With selForm.SplitContainerUtilities
                If .Panel1Collapsed = False Then
                    selForm.Clipboard.dgClipboard.LoadColumns(intContentType)
                End If
            End With
        Next
    End Sub

    Public Sub ClipboardShowHide(ByVal intContentType As Integer)
        For Each selForm As frmProject In Me.MdiChildren
            With selForm.SplitContainerUtilities
                If .Panel1Collapsed Then
                    .Panel1Collapsed = False
                    selForm.Clipboard.dgClipboard.LoadColumns(intContentType)
                Else
                    .Panel1Collapsed = True
                End If
            End With
        Next
    End Sub
#End Region

#Region "Ribbon-Text-Typeface"
    Private Sub RibbonComboBoxFontName_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.RibbonItemEventArgs) Handles RibbonComboBoxFontName.DropDownItemClicked
        Dim selButton As RibbonButton = CType(e.Item, RibbonButton)
        CurrentText.FontComplex = False
        CurrentText.ChangeFontName(selButton.Text)

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFontChanged

        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonComboBoxFontSize_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.RibbonItemEventArgs) Handles RibbonComboBoxFontSize.DropDownItemClicked
        Dim selButton As RibbonButton = CType(e.Item, RibbonButton)
        Dim intSize As Integer
        If Integer.TryParse(selButton.Text, intSize) Then
            CurrentText.FontSizeComplex = False
            CurrentText.ChangeFontSize(intSize)

            UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFontSizeChanged

            CurrentProjectForm.ChangeFont()
        End If
    End Sub

    Private Sub RibbonButtonFontLarger_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontLarger.Click
        Dim intSize As Integer = CurrentText.Font.Size

        Select Case CurrentText.Font.Size
            Case 8 To 11 '8-12
                intSize += 1
            Case 12 To 26 '14-28
                intSize += 2
            Case 28
                intSize = 36
            Case 36
                intSize = 48
            Case 48
                intSize = 72
        End Select

        CurrentText.ChangeFontSize(intSize)

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFontSizeChanged

        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonButtonFontSmaller_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontSmaller.Click
        Dim intSize As Integer = CurrentText.Font.Size

        Select Case CurrentText.Font.Size
            Case 9 To 12 '8-12
                intSize -= 1
            Case 14 To 28 '14-28
                intSize -= 2
            Case 36
                intSize = 28
            Case 48
                intSize = 36
            Case 72
                intSize = 48
        End Select

        CurrentText.ChangeFontSize(intSize)

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFontSizeChanged

        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonButtonUppercase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonUppercase.Click
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextCaseUpper
        CurrentProjectForm.ChangeFontCase(RichTextManager.TextCases.Uppercase)
    End Sub

    Private Sub RibbonButtonLowercase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLowercase.Click
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextCaseLower
        CurrentProjectForm.ChangeFontCase(RichTextManager.TextCases.Lowercase)
    End Sub

    Private Sub RibbonButtonSentence_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSentence.Click
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextCaseSentence
        CurrentProjectForm.ChangeFontCase(RichTextManager.TextCases.Sentence)
    End Sub

    Private Sub RibbonButtonFirstLetterUppercase_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFirstLetterUppercase.Click
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextCaseFirstLetter
        CurrentProjectForm.ChangeFontCase(RichTextManager.TextCases.FirstLetter)
    End Sub

    Private Sub RibbonButtonFontBold_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontBold.Click
        FontBold()
    End Sub

    Private Sub FontBold()
        Dim boolBold As Boolean

        With CurrentText
            boolBold = .Font.Bold

            If boolBold = False Then boolBold = True Else boolBold = False

            .ChangeStyle(boolBold, .Font.Italic, .Font.Underline, .Font.Strikeout)
        End With

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextBold
        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonButtonFontItalic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontItalic.Click
        FontItalic()
    End Sub

    Private Sub FontItalic()
        Dim boolItalic As Boolean

        With CurrentText
            boolItalic = .Font.Italic

            If boolItalic = False Then boolItalic = True Else boolItalic = False

            .ChangeStyle(.Font.Bold, boolItalic, .Font.Underline, .Font.Strikeout)
        End With

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextItalic
        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonButtonFontUnderline_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontUnderline.Click
        FontUnderline()
    End Sub

    Private Sub FontUnderline()
        Dim boolUnderline As Boolean

        With CurrentText
            boolUnderline = .Font.Underline

            If boolUnderline = False Then boolUnderline = True Else boolUnderline = False

            .ChangeStyle(.Font.Bold, .Font.Italic, boolUnderline, .Font.Strikeout)
        End With

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextUnderline
        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonButtonFontStrikeout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontStrikeout.Click
        With CurrentText
            .ChangeStyle(.Font.Bold, .Font.Italic, .Font.Underline, RibbonButtonFontStrikeout.Checked)
        End With
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextStrikeOut
        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonButtonFontSubscript_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontSubscript.Click
        CurrentText.ChangeCharOffset(classSelectedText.TextOffSetValues.SubScript)

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextSubScript

        CurrentProjectForm.ChangeFontOffSet(classSelectedText.TextOffSetValues.SubScript)
    End Sub

    Private Sub RibbonButtonFontSuperscript_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontSuperscript.Click
        CurrentText.ChangeCharOffset(classSelectedText.TextOffSetValues.SuperScript)

        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextSuperScript

        CurrentProjectForm.ChangeFontOffSet(classSelectedText.TextOffSetValues.SuperScript)
    End Sub

    Private Sub RibbonButtonFontColor_DropDownItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.RibbonItemEventArgs) Handles RibbonButtonFontColor.DropDownItemClicked
        RibbonButtonFontColor.CloseDropDown()
    End Sub

    Private Sub RibbonButtonFontColor_DropDownShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFontColor.DropDownShowing
        ColorPicker = New ucColorPaletteDialog
        Dim p As Point = New Point(RibbonButtonFontColor.Bounds.Left, RibbonButtonFontColor.Bounds.Bottom)

        p = PointToScreen(p)
        With ColorPicker
            .Type = ucColorPaletteDialog.PaletteType.normal_colors
            .ColorPaletteDialog(p.X, p.Y)

            If CurrentEditingControl IsNot Nothing Then .ColorText = CurrentEditingControl.SelectionColor
            .ShowDialog()
            If CurrentEditingControl IsNot Nothing And .DialogResult = Windows.Forms.DialogResult.OK Then
                UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFrontColor
                CurrentProjectForm.ChangeFontColor(.ColorChoice, False)
            End If
        End With
    End Sub

    Private Sub RibbonButtonFontColor_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RibbonButtonFontColor.MouseDown
        RibbonButtonFontColor.CloseDropDown()
    End Sub

    Private Sub RibbonButtonMarkerColor_DropDownShowing(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonMarkerColor.DropDownShowing
        ColorPicker = New ucColorPaletteDialog
        Dim p As Point = New Point(RibbonButtonMarkerColor.Bounds.Left, RibbonButtonMarkerColor.Bounds.Bottom)

        p = PointToScreen(p)
        With ColorPicker
            .Type = ucColorPaletteDialog.PaletteType.marker_colors
            .ColorPaletteDialog(p.X, p.Y)

            If CurrentEditingControl IsNot Nothing Then .ColorMarker = CurrentEditingControl.SelectionBackColor
            .ShowDialog()
            If CurrentEditingControl IsNot Nothing And .DialogResult = Windows.Forms.DialogResult.OK Then
                UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextBackColor
                CurrentProjectForm.ChangeFontColor(.ColorChoice, True)
            End If
        End With
    End Sub

    Private Sub RibbonButtonMarkerColor_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RibbonButtonMarkerColor.MouseDown
        RibbonButtonMarkerColor.CloseDropDown()
    End Sub

    Private Sub RibbonButtonEraseFormating_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonEraseFormating.Click
        With CurrentText
            .FontComplex = False
            .FontSizeComplex = False
            .Font = My.Settings.setDefaultFont
        End With
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFontChanged
        CurrentProjectForm.ChangeFont()
    End Sub

    Private Sub RibbonPanelTypeface_ButtonMoreClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonPanelTypeface.ButtonMoreClick
        Using objFontDialog As New FontDialog

            With CurrentText
                objFontDialog.Font = .Font
                objFontDialog.ShowColor = False
                '.Color = CurrentText.FontColor

                If objFontDialog.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                    .Font = objFontDialog.Font
                    .FontComplex = False
                    .FontSizeComplex = False
                    UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.TextFontChanged
                    CurrentProjectForm.ChangeFont()
                End If
            End With
        End Using
    End Sub
#End Region

#Region "Ribbon-Text-Paragraph"
    Private Sub RibbonButtonAlignLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonAlignLeft.Click
        CurrentText.HorizontalAlignment = HorizontalAlignment.Left
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.ParagraphAlignLeft
        CurrentProjectForm.ChangeParagraphAlignment()
    End Sub

    Private Sub RibbonButtonAlignCentre_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonAlignCentre.Click
        CurrentText.HorizontalAlignment = HorizontalAlignment.Center
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.ParagraphAlignCenter
        CurrentProjectForm.ChangeParagraphAlignment()
    End Sub

    Private Sub RibbonButtonAlignRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonAlignRight.Click
        CurrentText.HorizontalAlignment = HorizontalAlignment.Right
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.ParagraphAlignRight
        CurrentProjectForm.ChangeParagraphAlignment()
    End Sub

    Private Sub RibbonButtonLeftIndentIncrease_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLeftIndentIncrease.Click
        Dim intIndent As Integer = My.Settings.setIndentIncrement
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.ParagraphLeftIndentChanged
        CurrentProjectForm.ChangeParagraphLeftIndent(intIndent)
    End Sub

    Private Sub RibbonButtonLeftIndentDecrease_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonLeftIndentDecrease.Click
        Dim intIndent As Integer = My.Settings.setIndentIncrement * -1
        UndoRedo.UndoBuffer.ActionIndex = classUndoRedo.Actions.ParagraphLeftIndentChanged
        CurrentProjectForm.ChangeParagraphLeftIndent(intIndent)
    End Sub
#End Region

#Region "Ribbon-Text-Edit"
    Private Sub RibbonButtonFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonFind.Click
        DisplayFindDialog()
    End Sub

    Private Sub DisplayFindDialog()
        Dim FindDialog As New DialogFind
        With FindDialog
            .TabControlFindReplace.SelectTab(0)
            .ActiveControl = .tbFind
            .Show()
        End With
    End Sub

    Private Sub RibbonButtonReplace_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonReplace.Click
        DisplayReplaceDialog()
    End Sub

    Private Sub DisplayReplaceDialog()
        Dim FindDialog As New DialogFind
        With FindDialog
            .TabControlFindReplace.SelectTab(1)
            .ActiveControl = .tbFindReplace
            .Show()
        End With
    End Sub

    Private Sub RibbonButtonSelectAllText_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectAllText.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectAll)
    End Sub

    Private Sub RibbonButtonSelectAssumptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectAssumptions.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectAssumptions)
    End Sub

    Private Sub RibbonButtonSelectBudget_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectBudget.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectResourceBudgets)
    End Sub

    Private Sub RibbonButtonSelectBudgetItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectBudgetItems.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectBudgetItems)
    End Sub

    Private Sub RibbonButtonSelectIndicators_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectIndicators.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectIndicators)
    End Sub

    Private Sub RibbonButtonSelectLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectLogframe.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectLogframe)
    End Sub

    Private Sub RibbonButtonSelectNothing_Click(sender As Object, e As System.EventArgs) Handles RibbonButtonSelectNothing.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectNothing)
    End Sub

    Private Sub RibbonButtonSelectProjectLogic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectProjectLogic.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectStructs)
    End Sub

    Private Sub RibbonButtonSelectResources_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectResources.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectResources)
    End Sub

    Private Sub RibbonButtonSelectStatements_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectStatements.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectStatements)
    End Sub

    Private Sub RibbonButtonSelectTextInCell_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectTextInCell.Click
        If CurrentEditingControl IsNot Nothing Then
            CurrentEditingControl.SelectAll()
        End If
    End Sub

    Private Sub RibbonButtonSelectTextInLogframe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectTextInLogframe.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectLogframe)
    End Sub

    Private Sub RibbonButtonSelectVerificationSources_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSelectVerificationSources.Click
        CurrentProjectForm.SelectText(frmProject.TextSelectionValues.SelectVerificationSources)
    End Sub
#End Region

End Class
