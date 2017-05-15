Imports System.Globalization

Module modMain
    Friend WithEvents CurrentText As New classSelectedText
    Friend WithEvents TextClipboard As New ClipboardTextItems
    Friend WithEvents ItemClipboard As New ClipboardItems
    Friend WithEvents UndoRedo As New classUndoRedo

    Private intSelectedMainTab, intSelectedQuestionnaireTab, intSelectedActivitiesTab, intSelectedIndicatorsTab As Integer

    Public Const CONST_AllResponses As Integer = -1
    Public Const CONST_Padding As Integer = 2
    Public Const CONST_HorizontalPadding As Integer = 4
    Public Const CONST_VerticalPadding As Integer = 4
    Public Const CONST_StructIndentMultiplier As Integer = 10
    Public Const CONST_QuestionTypeTitle As Integer = -10

    Public Const CONST_Equals As String = "="
    Public Const CONST_SmallerThan As String = "<"
    Public Const CONST_SmallerThanOrEqual As String = "<="
    Public Const CONST_LargerThan As String = ">"
    Public Const CONST_LargerThanOrEqual As String = ">="

    'planning
    Public Const CONST_PlanningColumnIndex As Integer = 4
    Public Const CONST_MinRowHeight As Integer = 24
    Public Const CONST_BarHeight As Integer = 16
    Public Const CONST_PreparationHeight As Integer = 8
    Public Const CONST_FollowUpHeight As Integer = 8
    Public Const CONST_DistanceToActivity = 10
    Public Const CONST_DistanceToKeyMoment = 6
    Public Const CONST_SelectRadius = 5

    Public MeasureUnitsUser As New List(Of StructuredComboBoxItem)
    Public ResponseTypes(13) As StructuredComboBoxItem
    Public UserLanguage As String = My.Settings.setLanguage
    Public FontLibrary As New List(Of String)
    Public FontSizes As New List(Of String)

    Public penSelection As New Pen(Brushes.Red, 2)
    Public penBlack1 As Pen = New Pen(Color.Black, 1)
    Public penBlack2 As Pen = New Pen(Color.Black, 2)
    Public penDarkBlue1 As New Pen(Brushes.MediumBlue, 1)
    Public penLightBlue2 As New Pen(Brushes.DeepSkyBlue, 2)
    Public penBorder As Pen = New Pen(SystemColors.ControlDarkDark, 1)
    Public brBlue As New SolidBrush(Color.FromArgb(100, Color.LightBlue))

    Private frmCurrentProjectForm As frmProject
    Private ctlCurrentControl As Control
    Private tsTotal As New TimeSpan
    Private strFindFont, strFindFontSize As String
    Private intCurrentClipboardType As Integer

    Public Enum DurationUnits
        NotDefined = 0
        Minute = 1
        Hour = 2
        Day = 3
        Week = 4
        Month = 5
        Trimester = 6
        Semester = 7
        Year = 8
    End Enum

    Public Sub Main()

    End Sub

    Public Property MyPrintersettings As System.Drawing.Printing.PrinterSettings
        Get
            If My.Settings.setMyPrinterSettings Is Nothing Then My.Settings.setMyPrinterSettings = New System.Drawing.Printing.PrinterSettings
            Return My.Settings.setMyPrinterSettings
        End Get
        Set(ByVal value As System.Drawing.Printing.PrinterSettings)
            My.Settings.setMyPrinterSettings = value
        End Set
    End Property

#Region "Current selection"
    Public Property CurrentProjectForm As frmProject
        Get
            Return frmCurrentProjectForm
        End Get
        Set(ByVal value As frmProject)
            frmCurrentProjectForm = value
        End Set
    End Property

    Public Property CurrentLogFrame() As LogFrame
        Get
            If CurrentProjectForm IsNot Nothing Then
                Return CurrentProjectForm.ProjectLogframe
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As LogFrame)
            If CurrentProjectForm IsNot Nothing Then
                CurrentProjectForm.ProjectLogframe = value
            End If
        End Set
    End Property

    Public Property CurrentBudget() As Budget
        Get
            If CurrentLogFrame IsNot Nothing Then
                Return CurrentLogFrame.Budget
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Budget)
            If CurrentLogFrame IsNot Nothing Then
                CurrentLogFrame.Budget = value
            End If
        End Set
    End Property

    Public ReadOnly Property CurrentUndoList As classUndoList
        Get
            If CurrentProjectForm IsNot Nothing Then
                Return CurrentProjectForm.UndoList
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property CurrentRedoList As classUndoList
        Get
            If CurrentProjectForm IsNot Nothing Then
                Return CurrentProjectForm.RedoList
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property CurrentClipboardType As Integer
        Get
            Return intCurrentClipboardType
        End Get
        Set(ByVal value As Integer)
            intCurrentClipboardType = value
        End Set
    End Property

    Public ReadOnly Property CurrentDataGridView() As DataGridViewBaseClassRichText
        Get
            Return CurrentProjectForm.CurrentDataGridView
        End Get
    End Property

    Public ReadOnly Property CurrentEditingControl() As RichTextEditingControlLogframe
        Get
            If CurrentDataGridView IsNot Nothing Then
                Dim objEditingControl As RichTextEditingControlLogframe = TryCast(CurrentDataGridView.EditingControl, RichTextEditingControlLogframe)
                If objEditingControl Is Nothing Then
                    CurrentDataGridView.BeginEdit(False)

                    If CurrentDataGridView.EditingControl IsNot Nothing AndAlso CurrentDataGridView.EditingControl.GetType = GetType(RichTextEditingControlLogframe) Then
                        objEditingControl = CType(CurrentDataGridView.EditingControl, RichTextEditingControlLogframe)
                        If objEditingControl IsNot Nothing Then objEditingControl.Select(CurrentText.SelectionStart, CurrentText.SelectionLength)
                    End If
                    
                End If
                Return objEditingControl
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property CurrentControl As Control
        Get
            Return ctlCurrentControl
        End Get
        Set(ByVal value As Control)
            ctlCurrentControl = value
            'Debug.Print(ctlCurrentControl.Name)

            'NOTE: lay-out tabs are also set in frmParent through SetVisibilityLayOutTabs method
            frmParent.SetRibbonItemsVisibility()

            'set undo buffer with original value
            Dim strPropertyName As String = String.Empty
            Dim objValue As Object = Nothing

            If value.DataBindings.Count > 0 Then
                Dim selDataBinding As Binding = CurrentControl.DataBindings(0)
                strPropertyName = selDataBinding.BindingMemberInfo.BindingMember

            End If
            UndoRedo.CurrentControlChanged(strPropertyName)

        End Set
    End Property

    Public Property CurrentMainTab As Integer
        Get
            Return intSelectedMainTab
        End Get
        Set(ByVal value As Integer)
            intSelectedMainTab = value
            Select Case intSelectedMainTab
                Case 1 'Logframe
                    CurrentProjectForm.dgvLogframe.Focus()
                    CurrentControl = CurrentProjectForm.dgvLogframe
                Case 2 'Planning
                    CurrentProjectForm.dgvPlanning.Focus()
                    CurrentControl = CurrentProjectForm.dgvPlanning
                Case 3 'Budget
                    Dim dgvBudget As DataGridViewBudgetYear = CurrentProjectForm.GetCurrentDataGridViewBudgetYear()
                    If dgvBudget IsNot Nothing Then
                        dgvBudget.Focus()
                        CurrentControl = dgvBudget
                    End If
            End Select
        End Set
    End Property

    Public Property CurrentQuestionnaireTab As Integer
        Get
            Return intSelectedQuestionnaireTab
        End Get
        Set(ByVal value As Integer)
            intSelectedQuestionnaireTab = value
        End Set
    End Property

    Public Property CurrentActivitiesTab As Integer
        Get
            Return intSelectedActivitiesTab
        End Get
        Set(ByVal value As Integer)
            intSelectedActivitiesTab = value
        End Set
    End Property

    Public Property CurrentIndicatorsTab As Integer
        Get
            Return intSelectedIndicatorsTab
        End Get
        Set(ByVal value As Integer)
            intSelectedIndicatorsTab = value
        End Set
    End Property

    Public Function GetCurrentItem() As Object
        Dim objItem As Object = Nothing

        If CurrentControl IsNot Nothing Then
            Select Case CurrentControl.GetType
                Case GetType(DataGridViewLogframe)
                    objItem = CType(CurrentControl, DataGridViewLogframe).CurrentItem(False)
                Case GetType(DataGridViewPlanning)
                    objItem = CType(CurrentControl, DataGridViewPlanning).CurrentItem(False)
                Case GetType(DataGridViewBudgetYear)
                    objItem = CType(CurrentControl, DataGridViewBudgetYear).CurrentItem(False)
                Case GetType(DataGridViewBudgetItemReferences)
                    objItem = CType(CurrentControl, DataGridViewBudgetItemReferences).CurrentItem(False)
                Case GetType(DataGridViewResponseClasses)
                    objItem = CType(CurrentControl, DataGridViewResponseClasses).CurrentItem(False)
                Case GetType(DataGridViewStatementsFormula)
                    objItem = CType(CurrentControl, DataGridViewStatementsFormula).CurrentItem(False)
                Case GetType(DataGridViewStatementsMaxDiff)
                    objItem = CType(CurrentControl, DataGridViewStatementsMaxDiff).CurrentItem(False)
                Case GetType(DataGridViewStatementsScales)
                    objItem = CType(CurrentControl, DataGridViewStatementsScales).CurrentItem(False)
                Case GetType(DataGridViewTargetsClasses)
                    objItem = CType(CurrentControl, DataGridViewTargetsClasses).CurrentItem(False)
                Case GetType(DataGridViewTargetsFormula)
                    objItem = CType(CurrentControl, DataGridViewTargetsFormula).CurrentItem(False)
                Case GetType(DataGridViewTargetsFrequencyLikert)
                    objItem = CType(CurrentControl, DataGridViewTargetsFrequencyLikert).CurrentItem(False)
                Case GetType(DataGridViewTargetsMaxDiff)
                    objItem = CType(CurrentControl, DataGridViewTargetsMaxDiff).CurrentItem(False)
                Case GetType(DataGridViewTargetsScaleLikert)
                    objItem = CType(CurrentControl, DataGridViewTargetsScaleLikert).CurrentItem(False)
                Case GetType(DataGridViewTargetsScales)
                    objItem = CType(CurrentControl, DataGridViewTargetsScales).CurrentItem(False)
                Case GetType(DataGridViewTargetsSemanticDiff)
                    objItem = CType(CurrentControl, DataGridViewTargetsSemanticDiff).CurrentItem(False)

                Case GetType(ListViewAddresses)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewAddresses).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetAddressByGuid(selGuid)
                    End If
                Case GetType(ListViewBudgetItemReferences)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewBudgetItemReferences).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetBudgetItemByGuid(selGuid) 'GetBudgetItemReferenceByGuid ?
                    End If
                Case GetType(ListViewContacts)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewContacts).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetContactByGuid(selGuid)
                    End If
                Case GetType(ListViewEmails)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewEmails).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetEmailByGuid(selGuid)
                    End If
                Case GetType(ListViewKeyMoments)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewKeyMoments).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetKeyMomentByGuid(selGuid)
                    End If
                Case GetType(ListViewPartners)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewPartners).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetProjectPartnerByGuid(selGuid)
                    End If
                Case GetType(ListViewProcesses)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewProcesses).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetActivityByGuid(selGuid)
                    End If
                Case GetType(ListViewTargetGroups)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewTargetGroups).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetTargetGroupByGuid(selGuid)
                    End If
                Case GetType(ListViewTelephoneNumbers)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewTelephoneNumbers).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetTelephoneNumberByGuid(selGuid)
                    End If
                Case GetType(ListViewWebsites)
                    Dim selGuid As Guid = CType(CurrentControl, ListViewWebsites).SelectedGuid

                    If selGuid <> Nothing And selGuid <> Guid.Empty Then
                        objItem = CurrentLogFrame.GetWebsiteByGuid(selGuid)
                    End If

                Case Else
                    If CurrentControl.DataBindings.Count > 0 Then
                        Dim selDataBinding As Binding = CurrentControl.DataBindings(0)
                        Dim strProperty As String = selDataBinding.PropertyName
                        Dim selSource As BindingSource = selDataBinding.DataSource
                        objItem = selSource.DataSource
                    End If
            End Select
        End If

        Return objItem
    End Function
#End Region

#Region "Properties of last selected text (font & paragraph)"
    Private Sub CurrentText_FontNameChanged() Handles CurrentText.FontNameChanged
        With frmParent.RibbonComboBoxFontName
            If CurrentText.FontComplex = False Then
                strFindFont = CurrentText.Font.Name
                Dim intIndex As Integer = FontLibrary.FindIndex(AddressOf FindFontName)
                If intIndex >= 0 And intIndex <= .DropDownItems.Count - 1 Then _
                    .SelectedItem = .DropDownItems(intIndex)
            Else
                .SelectedValue = String.Empty
            End If
        End With
    End Sub

    Private Function FindFontName(ByVal strFont As String) As Boolean
        If strFont = strFindFont Then Return True Else Return False
    End Function

    Private Sub CurrentText_FontSizeChanged() Handles CurrentText.FontSizeChanged
        With frmParent.RibbonComboBoxFontSize
            If CurrentText.FontSizeComplex = False Then
                strFindFontSize = CurrentText.Font.Size.ToString("N0")
                Dim intIndex As Integer = FontSizes.FindIndex(AddressOf FindFontSize)
                If intIndex >= 0 And intIndex <= .DropDownItems.Count - 1 Then _
                    .SelectedItem = .DropDownItems(intIndex)
            Else
                .SelectedValue = String.Empty
            End If
        End With
    End Sub

    Private Function FindFontSize(ByVal strFontSize As String) As Boolean
        If strFontSize = strFindFontSize Then Return True Else Return False
    End Function

    Private Sub CurrentText_FontBoldChanged() Handles CurrentText.FontBoldChanged
        With frmParent.RibbonButtonFontBold
            If CurrentText.BoldComplex = False Then
                .Checked = CurrentText.Font.Bold
            Else
                .Checked = False
            End If
        End With
    End Sub

    Private Sub CurrentText_FontItalicChanged() Handles CurrentText.FontItalicChanged
        With frmParent.RibbonButtonFontItalic
            If CurrentText.ItalicComplex = False Then
                .Checked = CurrentText.Font.Italic
            Else
                .Checked = False
            End If
        End With
    End Sub

    Private Sub CurrentText_FontUnderlineChanged() Handles CurrentText.FontUnderlineChanged
        With frmParent.RibbonButtonFontUnderline
            If CurrentText.UnderlineComplex = False Then
                .Checked = CurrentText.Font.Underline
            Else
                .Checked = False
            End If
        End With
    End Sub

    Private Sub CurrentText_FontStrikeoutChanged() Handles CurrentText.FontStrikeoutChanged
        With frmParent.RibbonButtonFontStrikeout
            If CurrentText.StrikeoutComplex = False Then
                .Checked = CurrentText.Font.Strikeout
            Else
                .Checked = False
            End If
        End With
    End Sub

    Private Sub CurrentText_FontCharOffSetChanged() Handles CurrentText.FontCharOffSetChanged
        With frmParent
            If CurrentText.CharOffSet = 0 Then
                .RibbonButtonFontSubscript.Checked = False
                .RibbonButtonFontSuperscript.Checked = False
            ElseIf CurrentText.CharOffSet > 0 Then
                .RibbonButtonFontSubscript.Checked = False
                .RibbonButtonFontSuperscript.Checked = True
            ElseIf CurrentText.CharOffSet < 0 Then
                .RibbonButtonFontSubscript.Checked = True
                .RibbonButtonFontSuperscript.Checked = False
            End If
        End With
    End Sub

    Private Sub CurrentText_ParagraphAlignmentChanged() Handles CurrentText.ParagraphAlignmentChanged
        With frmParent
            Select Case CurrentText.HorizontalAlignment
                Case HorizontalAlignment.Left
                    .RibbonButtonAlignLeft.Checked = True
                    .RibbonButtonAlignCentre.Checked = False
                    .RibbonButtonAlignRight.Checked = False
                Case HorizontalAlignment.Center
                    .RibbonButtonAlignLeft.Checked = False
                    .RibbonButtonAlignCentre.Checked = True
                    .RibbonButtonAlignRight.Checked = False
                Case HorizontalAlignment.Right
                    .RibbonButtonAlignLeft.Checked = False
                    .RibbonButtonAlignCentre.Checked = False
                    .RibbonButtonAlignRight.Checked = True
            End Select
        End With
    End Sub
#End Region

#Region "Other text and font methods"
    Public Sub LoadFontLibrary()
        Dim objFontCollection As New System.Drawing.Text.InstalledFontCollection
        For Each objFontFamily As FontFamily In objFontCollection.Families
            FontLibrary.Add(objFontFamily.Name)
        Next
    End Sub

    Public Sub LoadFontSizes()
        For i = 8 To 12
            FontSizes.Add(i.ToString)
        Next
        For i = 14 To 28 Step 2
            FontSizes.Add(i.ToString)
        Next
        FontSizes.Add(36)
        FontSizes.Add(48)
        FontSizes.Add(72)
    End Sub

    Public Function IsRichText(ByVal strText As String) As Boolean
        Return strText.StartsWith("{\rtf1")
    End Function

    Public Function TextToRichText(ByVal strText As String, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal underline As Boolean = False) As String
        Dim strRtf As String

        Using rtfPainter As New RichTextManager
            With rtfPainter
                .Text = strText


                If bold = True Or italic = True Or underline = True Then
                    Dim objStyle As New FontStyle
                    .SelectAll()
                    If bold = True Then objStyle = FontStyle.Bold
                    If italic = True Then objStyle += FontStyle.Italic
                    If underline = True Then objStyle = FontStyle.Underline

                    .SelectionFont = New Font(.SelectionFont, objStyle)
                End If
                strRtf = .Rtf
            End With
        End Using

        Return strRtf
    End Function

    Public Function RichTextToText(ByVal strRtf As String) As String
        Dim strText As String

        Using rtfPainter As New RichTextManager
            rtfPainter.Rtf = strRtf
            strText = rtfPainter.Text
        End Using

        Return strText
    End Function
#End Region

#Region "Clipboard methods"
    Private Sub TextClipboard_ListUpdated() Handles TextClipboard.ListUpdated
        Dim selForm As frmProject = CType(frmParent.ActiveMdiChild, frmProject)
        If selForm IsNot Nothing Then
            With selForm.Clipboard.dgClipboard
                .LoadColumns(DataGridViewClipboard.ContentTypes.Text)
            End With
        End If
    End Sub

    Private Sub ItemClipboard_ListUpdated() Handles ItemClipboard.ListUpdated
        Dim selForm As frmProject = CType(frmParent.ActiveMdiChild, frmProject)
        If selForm IsNot Nothing Then
            With selForm.Clipboard.dgClipboard
                .LoadColumns(DataGridViewClipboard.ContentTypes.Items)
            End With
        End If
    End Sub
#End Region

#Region "Parse decimal numbers independently of global settings"
    Public Function ParseInteger(ByVal strNumber As String) As Integer
        Dim intValue As Integer
        Dim style As NumberStyles = NumberStyles.Float Or NumberStyles.AllowThousands
        Dim culture As CultureInfo = CultureInfo.CurrentCulture

        If Integer.TryParse(strNumber, style, culture, intValue) = False Then
            culture = CultureInfo.InvariantCulture

            Integer.TryParse(strNumber, style, culture, intValue)
        End If
        Return intValue
    End Function

    Public Function ParseSingle(ByVal strNumber As String) As Single
        Dim sngValue As Single
        Dim style As NumberStyles = NumberStyles.Float Or NumberStyles.AllowThousands
        Dim culture As CultureInfo = CultureInfo.CurrentCulture

        If Single.TryParse(strNumber, style, culture, sngValue) = False Then
            culture = CultureInfo.InvariantCulture

            Single.TryParse(strNumber, style, culture, sngValue)
        End If
        Return sngValue
    End Function

    Public Function ParseDouble(ByVal strNumber) As Double
        Dim dblValue As Double
        Dim style As NumberStyles = NumberStyles.Float Or NumberStyles.AllowThousands
        Dim culture As CultureInfo = CultureInfo.CurrentCulture

        If Double.TryParse(strNumber, style, culture, dblValue) = False Then
            culture = CultureInfo.InvariantCulture

            Double.TryParse(strNumber, style, culture, dblValue)
            'If Double.TryParse(strNumber, style, culture, dblValue) = False Then
            '    RaiseEvent ErrorParsingDecimal()
            'End If
        End If
        Return dblValue
    End Function

    Public Function DisplayAsUnit(ByVal dblValue As Double, ByVal intNrDecimals As Integer, ByVal strUnit As String) As String
        Dim strPrecision As String = "#,##0."
        If intNrDecimals > 0 Then
            For i = 1 To intNrDecimals
                strPrecision &= "0"
            Next
        End If
        If String.IsNullOrEmpty(strUnit) = False Then strPrecision = String.Format("{0} '{1}'", strPrecision, strUnit)

        Return dblValue.ToString(strPrecision)
    End Function
#End Region

#Region "Colour functions"
    Public Function LighterColor(ByVal OldColor As Color) As Color
        Dim intRed As Integer = OldColor.R + 50
        Dim intGreen As Integer = OldColor.G + 50
        Dim intBlue As Integer = OldColor.B + 50
        If intRed > 255 Then intRed = 255
        If intGreen > 255 Then intGreen = 255
        If intBlue > 255 Then intBlue = 255

        Dim NewColor As Color = Color.FromArgb(255, intRed, intGreen, intBlue)
        Return NewColor
    End Function

    Public Function LighterColor(ByVal OldColor As Color, ByVal intLevels As Integer) As Color
        Dim intRed As Integer = OldColor.R + (25 * intLevels)
        Dim intGreen As Integer = OldColor.G + (25 * intLevels)
        Dim intBlue As Integer = OldColor.B + (25 * intLevels)
        If intRed > 255 Then intRed = 255
        If intGreen > 255 Then intGreen = 255
        If intBlue > 255 Then intBlue = 255

        Dim NewColor As Color = Color.FromArgb(255, intRed, intGreen, intBlue)

        Return NewColor
    End Function

    Private Function LighterBrush(ByVal OldBrush As SolidBrush) As SolidBrush
        Dim NewColor As Color = LighterColor(OldBrush.Color)

        Dim NewBrush As New SolidBrush(NewColor)
        Return NewBrush
    End Function

    Public Function DarkerColor(ByVal OldColor As Color) As Color
        Dim intRed As Integer = OldColor.R - 50
        Dim intGreen As Integer = OldColor.G - 50
        Dim intBlue As Integer = OldColor.B - 50
        If intRed < 0 Then intRed = 0
        If intGreen < 0 Then intGreen = 0
        If intBlue < 0 Then intBlue = 0

        Dim NewColor As Color = Color.FromArgb(255, intRed, intGreen, intBlue)
        Return NewColor
    End Function

    Public Function DarkerColor(ByVal OldColor As Color, ByVal intLevels As Integer) As Color
        Dim intRed As Integer = OldColor.R - (25 * intLevels)
        Dim intGreen As Integer = OldColor.G - (25 * intLevels)
        Dim intBlue As Integer = OldColor.B - (25 * intLevels)
        If intRed < 0 Then intRed = 0
        If intGreen < 0 Then intGreen = 0
        If intBlue < 0 Then intBlue = 0

        Dim NewColor As Color = Color.FromArgb(255, intRed, intGreen, intBlue)

        Return NewColor
    End Function

    Private Function DarkerBrush(ByVal OldBrush As SolidBrush) As SolidBrush
        Dim NewColor As Color = DarkerColor(OldBrush.Color)

        Dim NewBrush As New SolidBrush(NewColor)
        Return NewBrush
    End Function
#End Region

    Public Sub ReadStopWatch(ByVal strProcessName As String, ByVal objStopWatch As Stopwatch)
        'Get the elapsed time as a TimeSpan value.
        Dim ts As TimeSpan = objStopWatch.Elapsed
        tsTotal = tsTotal.Add(ts)
        Dim elapsedTime As String = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", _
                                                  ts.Hours, ts.Minutes, ts.Seconds, _
                                                  ts.Milliseconds / 10)
        Dim totalTime As String = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", _
                                                  tsTotal.Hours, tsTotal.Minutes, tsTotal.Seconds, tsTotal.Milliseconds / 10)
        'Console.WriteLine(elapsedTime)
        'Console.ReadLine()
    End Sub

    

    
End Module
