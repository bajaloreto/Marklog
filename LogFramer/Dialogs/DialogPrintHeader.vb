Imports System.Windows.Forms

Public Class DialogPrintHeader
    Private objRichTextBox As RichTextBox = Nothing
    Public FontNameModifiedByUser As Boolean
    Public FontSizeModifiedByUser As Boolean

#Region "Properties"
    Public Enum TextOffSetValues
        SuperScript = 1
        SubScript = -1
    End Enum

    Private Property CurrentRichTextBox As RichTextBox
        Get
            Return objRichTextBox
        End Get
        Set(ByVal value As RichTextBox)
            objRichTextBox = value
        End Set
    End Property
#End Region

#Region "Events"
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        'Load values of comboboxes
        'Fonts
        Dim fnt As System.Drawing.Text.InstalledFontCollection = New System.Drawing.Text.InstalledFontCollection
        Dim fFamily As FontFamily
        Dim i As Integer

        With cmbFonts
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            For Each fFamily In fnt.Families
                .Items.Add(fFamily.Name)
            Next
            .SelectedItem = My.Settings.setDefaultFont.FontFamily.Name
        End With

        'Fontsizes
        With cmbFontSize
            .DropDownStyle = ComboBoxStyle.DropDown
            For i = 8 To 12
                .Items.Add(i.ToString)
            Next
            For i = 14 To 28 Step 2
                .Items.Add(i.ToString)
            Next
            .Items.Add("36")
            .Items.Add("48")
            .Items.Add("72")
            .SelectedItem = My.Settings.setDefaultFont.SizeInPoints.ToString
        End With

        rtbLeft.Font = My.Settings.setDefaultFont
        rtbMiddle.Font = My.Settings.setDefaultFont
        If rtbMiddle.Text = "" Then
            rtbMiddle.SelectAll()
            rtbMiddle.SelectionAlignment = HorizontalAlignment.Center
        End If
        rtbRight.Font = My.Settings.setDefaultFont
        If rtbRight.Text = "" Then
            rtbRight.SelectAll()
            rtbRight.SelectionAlignment = HorizontalAlignment.Right
        End If
    End Sub

    Private Sub cmbFonts_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFonts.Enter
        If MouseButtons = Windows.Forms.MouseButtons.Left Then FontNameModifiedByUser = True

    End Sub

    Private Sub cmbFonts_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFonts.Leave
        FontNameModifiedByUser = False
    End Sub

    Private Sub cmbFonts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFonts.SelectedIndexChanged
        If FontNameModifiedByUser = True Then
            ChangeFont()

            FontNameModifiedByUser = False
        End If
    End Sub

    Private Sub cmbFontSize_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFontSize.Enter
        If MouseButtons = Windows.Forms.MouseButtons.Left Then _
            FontSizeModifiedByUser = True
    End Sub

    Private Sub cmbFontSize_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFontSize.Leave
        FontSizeModifiedByUser = False
    End Sub

    Private Sub cmbFontSize_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFontSize.SelectedIndexChanged
        If FontSizeModifiedByUser = True Then
            ChangeFont()

            FontSizeModifiedByUser = False
        End If
    End Sub

    Private Sub buttonTextBold_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextBold.CheckStateChanged
        ChangeFont()
    End Sub

    Private Sub buttonTextItalics_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextItalics.CheckStateChanged
        ChangeFont()
    End Sub

    Private Sub buttonTextUnderlined_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextUnderlined.CheckStateChanged
        ChangeFont()
    End Sub

    Private Sub buttonTextStrikeThrough_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextStrikeThrough.CheckStateChanged
        ChangeFont()
    End Sub

    Private Sub buttonTextSuperscript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonTextSuperscript.Click
        TextOffset(TextOffSetValues.SuperScript)
    End Sub

    Private Sub buttonTextSubscript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonTextSubscript.Click
        'CurrentUndoList.ChangeTextOperation(UndoListItemOld.Actions.SubScript)
        TextOffset(TextOffSetValues.SubScript)
    End Sub

    Private Sub ChangeFontColor(ByVal selColor As Color, ByVal boolMarker As Boolean)
        If boolMarker = False Then
            ChangeTextColor(selColor)
        Else
            ChangeMarkerColor(selColor)
        End If
    End Sub

    Private Sub buttonTextColor_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonTextColor.ButtonClick
        ChangeFontColor(buttonTextColor.ForeColor, False)
    End Sub

    Private Sub buttonTextColor_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextColor.DropDownOpening
        Dim cpColorPicker As New ucColorPaletteDialog
        Dim p As Point = New Point(buttonTextColor.Bounds.Left, tspFormatting.Bounds.Top + buttonTextColor.Bounds.Top + buttonTextColor.Bounds.Height)

        p = PointToScreen(p)
        cpColorPicker.Type = ucColorPaletteDialog.PaletteType.normal_colors
        cpColorPicker.ColorPaletteDialog(p.X, p.Y)

        cpColorPicker.ColorText = CurrentRichTextBox.SelectionColor
        cpColorPicker.ShowDialog()
        If cpColorPicker.DialogResult = Windows.Forms.DialogResult.OK Then
            ChangeFontColor(cpColorPicker.ColorChoice, False)
        End If
    End Sub

    Private Sub buttonTextBackGround_ButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextBackGround.ButtonClick
        ChangeFontColor(buttonTextBackGround.BackColor, True)
    End Sub

    Private Sub buttonTextBackGround_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonTextBackGround.DropDownOpening
        Dim cpColorPicker As New ucColorPaletteDialog
        Dim p As Point = New Point(buttonTextBackGround.Bounds.Left, tspFormatting.Bounds.Top + buttonTextBackGround.Bounds.Top + buttonTextBackGround.Bounds.Height)

        p = PointToScreen(p)
        cpColorPicker.Type = ucColorPaletteDialog.PaletteType.marker_colors
        cpColorPicker.ColorPaletteDialog(p.X, p.Y)

        cpColorPicker.ColorMarker = CurrentRichTextBox.SelectionBackColor
        cpColorPicker.ShowDialog()
        If cpColorPicker.DialogResult = Windows.Forms.DialogResult.OK Then
            ChangeFontColor(cpColorPicker.ColorChoice, True)
        End If
    End Sub

    Private Sub buttonParAlignLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonParAlignLeft.Click
        TextAlignment(HorizontalAlignment.Left)
    End Sub

    Private Sub buttonParAlignCentre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonParAlignCentre.Click
        TextAlignment(HorizontalAlignment.Center)
    End Sub

    Private Sub buttonParAlignRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonParAlignRight.Click
        TextAlignment(HorizontalAlignment.Right)
    End Sub

    Public Sub ButtonSelect(ByVal selFont As Font)

        FontNameModifiedByUser = False
        FontSizeModifiedByUser = False
        Try
            If cmbFonts.FindString(selFont.FontFamily.Name) >= 0 Then
                cmbFonts.SelectedItem = selFont.FontFamily.Name.ToString
            End If

            Dim sngSize As Single = selFont.SizeInPoints
            Dim strSize As String = sngSize.ToString

            If cmbFontSize.Items.Contains(strSize) = False Then
                Dim selSize As Single
                For i = 0 To cmbFontSize.Items.Count - 1
                    Single.TryParse(cmbFontSize.Items(i), selSize)
                    If selSize > sngSize Then
                        cmbFontSize.Items.Insert(i, strSize)
                        Exit For
                    End If
                Next
            End If

            cmbFontSize.SelectedItem = strSize

            If selFont.Bold = True Then
                buttonTextBold.Checked = True
            Else
                buttonTextBold.Checked = False
            End If
            If selFont.Italic = True Then
                buttonTextItalics.Checked = True
            Else
                buttonTextItalics.Checked = False
            End If
            If selFont.Underline = True Then
                buttonTextUnderlined.Checked = True
            Else
                buttonTextUnderlined.Checked = False
            End If
            If selFont.Strikeout = True Then
                buttonTextStrikeThrough.Checked = True
            Else
                buttonTextStrikeThrough.Checked = False
            End If

            With CurrentRichTextBox
                If .SelectionCharOffset > 0 Then buttonTextSuperscript.Checked = True Else _
                    buttonTextSuperscript.Checked = False
                If .SelectionCharOffset < 0 Then buttonTextSubscript.Checked = True Else _
                    buttonTextSubscript.Checked = False
                If .SelectionAlignment = HorizontalAlignment.Left Then _
                    buttonParAlignLeft.Checked = True Else buttonParAlignLeft.Checked = False
                If .SelectionAlignment = HorizontalAlignment.Center Then _
                    buttonParAlignCentre.Checked = True Else buttonParAlignCentre.Checked = False
                If .SelectionAlignment = HorizontalAlignment.Right Then _
                    buttonParAlignRight.Checked = True Else buttonParAlignRight.Checked = False
            End With
        Catch ex As NullReferenceException
            '
        End Try

    End Sub

    Private Sub ButtonSelect_FindSizeInItems(ByVal strSize As String, ByVal sngSize As Single)
        Dim PrevSize As Single
        Dim i As Integer, index As Integer

        With cmbFontSize
            For i = 0 To .Items.Count - 1
                PrevSize = Single.Parse(.Items(i))
                If PrevSize > sngSize Then
                    If i > 0 Then index = i - 1
                    .Items.Insert(index, strSize)
                    Exit For
                End If
            Next
        End With
    End Sub

    Private Sub rtbLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbLeft.Click
        ButtonSelect(rtbLeft.SelectionFont)
    End Sub

    Private Sub rtbLeft_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbLeft.Enter
        If CurrentRichTextBox IsNot rtbLeft Then CurrentRichTextBox = rtbLeft
    End Sub

    Private Sub rtbMiddle_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbMiddle.Click
        ButtonSelect(rtbMiddle.SelectionFont)
    End Sub

    Private Sub rtbMiddle_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbMiddle.Enter
        If CurrentRichTextBox IsNot rtbMiddle Then CurrentRichTextBox = rtbMiddle
    End Sub

    Private Sub rtbRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbRight.Click
        ButtonSelect(rtbRight.SelectionFont)
    End Sub

    Private Sub rtbRight_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbRight.Enter
        If CurrentRichTextBox IsNot rtbRight Then CurrentRichTextBox = rtbRight
    End Sub

    Private Sub buttonAddPageNumber_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonAddPageNumber.Click
        If CurrentRichTextBox IsNot Nothing Then
            CurrentRichTextBox.SelectedText = "{&Page}"
        End If
    End Sub

    Private Sub buttonAddTotalPages_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonAddTotalPages.Click
        If CurrentRichTextBox IsNot Nothing Then
            CurrentRichTextBox.SelectedText = "{&Pages}"
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
#End Region

#Region "Change text"
    Public Sub ChangeFont()
        If CurrentRichTextBox Is Nothing Then Exit Sub

        Dim strFontFamily As String = cmbFonts.SelectedItem.ToString
        Dim sngFontSize As Single
        Dim boolBold As Boolean = buttonTextBold.Checked
        Dim boolItalic As Boolean = buttonTextItalics.Checked
        Dim boolUnderline As Boolean = buttonTextUnderlined.Checked
        Dim boolStrikeout As Boolean = buttonTextStrikeThrough.Checked
        Dim fstNewFontStyle As FontStyle
        Dim intSelStart As Integer, intSelStop As Integer
        Dim intSelLen As Integer
        Dim i As Integer

        If Single.TryParse(cmbFontSize.SelectedItem, sngFontSize) = False Then sngFontSize = My.Settings.setDefaultFont.SizeInPoints

        intSelStart = CurrentRichTextBox.SelectionStart
        intSelLen = CurrentRichTextBox.SelectionLength
        intSelStop = intSelStart + intSelLen

        For i = intSelStart To intSelStop - 1
            fstNewFontStyle = FontStyle.Regular
            CurrentRichTextBox.Select(i, 1)
            If strFontFamily = "" Then strFontFamily = CurrentRichTextBox.SelectionFont.Name
            If sngFontSize > 0 Then
                If CurrentRichTextBox.SelectionCharOffset <> 0 Then sngFontSize *= 0.6
            Else
                sngFontSize = CurrentRichTextBox.SelectionFont.SizeInPoints
            End If

            If boolBold = True Then fstNewFontStyle = FontStyle.Bold
            If boolItalic = True Then fstNewFontStyle += FontStyle.Italic
            If boolUnderline = True Then fstNewFontStyle += FontStyle.Underline
            If boolStrikeout = True Then fstNewFontStyle += FontStyle.Strikeout
            Try
                CurrentRichTextBox.SelectionFont = New Font(strFontFamily, sngFontSize, fstNewFontStyle, GraphicsUnit.Point)
            Catch ex2 As ArgumentException
                Try
                    CurrentRichTextBox.SelectionFont = New Font(strFontFamily, sngFontSize, CurrentRichTextBox.SelectionFont.Style, GraphicsUnit.Point)
                Catch ex3 As ArgumentException
                    Dim strMsg As String

                    Select Case UserLanguage
                        Case "fr"
                            strMsg = "Cette police ne supporte pas le style (gras, italique, souligné ou barré) " & _
                                 "souhaité." & vbLf & vbLf & _
                                 "Veuillez choisir une autre police dans la liste."
                        Case Else
                            strMsg = "This font does not support the style (bold, italic, underline or strikethrough) " & _
                                "you requested." & vbLf & vbLf & _
                                "Please choose another font from the list."
                    End Select
                    MsgBox(strMsg)
                    CurrentRichTextBox.Select(intSelStart, intSelLen)
                    Exit Sub
                End Try
            End Try

        Next
        CurrentRichTextBox.Select(intSelStart, intSelLen)

        ButtonSelect(CurrentRichTextBox.SelectionFont)
    End Sub

    Public Sub TextOffset(ByVal intUpDown As Integer)
        Dim sngSize As Single
        Dim intStart As Integer, intLen As Integer

        With CurrentRichTextBox
            If .SelectionCharOffset = 0 Then
                .SelectionCharOffset = .SelectionFont.SizeInPoints * intUpDown * 0.4
                sngSize = .SelectionFont.SizeInPoints * 0.6
                .SelectionFont = New Font(.SelectionFont.Name, sngSize, .SelectionFont.Style, GraphicsUnit.Point)
            Else
                intStart = .SelectionStart
                intLen = .SelectionLength

                If intStart > 0 Then
                    .SelectionStart = intStart - 1
                    .SelectionLength = 0
                End If
                sngSize = .SelectionFont.SizeInPoints
                .SelectionStart = intStart
                .SelectionLength = intLen
                .SelectionCharOffset = 0

                .SelectionFont = New Font(.SelectionFont.Name, sngSize, .SelectionFont.Style, GraphicsUnit.Point)
            End If
        End With

        ButtonSelect(CurrentRichTextBox.SelectionFont)

    End Sub

    Public Sub ChangeTextColor(ByVal color As Color)
        CurrentRichTextBox.SelectionColor = color
    End Sub

    Public Sub ChangeMarkerColor(ByVal color As Color)
        CurrentRichTextBox.SelectionBackColor = color
    End Sub

    Public Sub TextAlignment(ByVal intAlignment As Integer)
        CurrentRichTextBox.SelectionAlignment = intAlignment
    End Sub
#End Region
End Class
