Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization

Public Class DataGridViewQuestionnaire
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New QuestionnaireGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents SelectionRectangle As New SelectionRectangle

    Public Event Reloaded()
    Public Event IndicatorSelected(ByVal sender As Object, ByVal e As IndicatorSelectedEventArgs)

    Private Const CONST_MinRowHeight As Integer = 24

    Private objTargetDeadlines As TargetDeadlines
    Private colSortNumber As New DataGridViewTextBoxColumn
    Private colRTF As New RichTextColumnLogframe
    Private colRange As New DataGridViewTextBoxColumn
    Private boolColumnWidthChanged As Boolean

    Private RowModifyIndex As Integer = -1
    Private EditRow As QuestionnaireGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private SelectedGridRows As New List(Of QuestionnaireGridRow)
    Private PasteRow As QuestionnaireGridRow
    Private DragIndicator As Indicator
    Private sngScale As Single
    Private ClickPoint As Point
    Private pStartInsertLine, pEndInsertLine As New Point
    Private pStartInsertLineOld, pEndInsertLineOld As New Point

    Private intShowRegistrationOption As Integer
    Private objShowTargetGroup As New Guid
    Private boolHideEmptyCells As Boolean
    Private boolShowGoals As Boolean = True
    Private boolShowPurposes As Boolean = True
    Private boolShowOutputs As Boolean = True
    Private boolShowActivities As Boolean = True

#Region "Properties"
    Public Property TargetDeadlines As TargetDeadlines
        Get
            Return objTargetDeadlines
        End Get
        Set(value As TargetDeadlines)
            objTargetDeadlines = value
        End Set
    End Property

    Public Property ShowRegistrationOption As Integer
        Get
            Return intShowRegistrationOption
        End Get
        Set(value As Integer)
            intShowRegistrationOption = value
        End Set
    End Property

    Public Property ShowTargetGroup As Guid
        Get
            Return objShowTargetGroup
        End Get
        Set(value As Guid)
            objShowTargetGroup = value
        End Set
    End Property

    Public Property HideEmptyCells As Boolean
        Get
            Return boolHideEmptyCells
        End Get
        Set(ByVal value As Boolean)
            boolHideEmptyCells = value
        End Set
    End Property

    Public Property ShowGoals() As Boolean
        Get
            Return boolShowGoals
        End Get
        Set(ByVal value As Boolean)
            boolShowGoals = value
        End Set
    End Property

    Public Property ShowPurposes() As Boolean
        Get
            Return boolShowPurposes
        End Get
        Set(ByVal value As Boolean)
            boolShowPurposes = value
        End Set
    End Property

    Public Property ShowOutputs() As Boolean
        Get
            Return boolShowOutputs
        End Get
        Set(ByVal value As Boolean)
            boolShowOutputs = value
        End Set
    End Property

    Public Property ShowActivities() As Boolean
        Get
            Return boolShowActivities
        End Get
        Set(ByVal value As Boolean)
            boolShowActivities = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim selPlanningRow As QuestionnaireGridRow = Me.Grid(Me.CurrentCell.RowIndex)
                If selPlanningRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
                    Dim selIndicator As Indicator = selPlanningRow.Indicator
                    If selIndicator IsNot Nothing Then
                        If OnlyIfTextShows = True And String.IsNullOrEmpty(selIndicator.RTF) = False Then
                            selObject = selIndicator
                        Else
                            selObject = selIndicator
                        End If
                    End If

                End If
            End If
            Return selObject
        End Get
    End Property
#End Region 'Properties

#Region "Initialise"
    Public Sub New()
        DoubleBuffered = True
        VirtualMode = True
        AutoGenerateColumns = False
        AllowUserToAddRows = False

        ShowCellToolTips = False
        MultiSelect = True
        RowHeadersVisible = False
        BackgroundColor = Color.White
        DefaultCellStyle.Padding = New Padding(2)

        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = False
    End Sub

    Public Sub InitialiseColumns()
        Columns.Clear()

        LoadColumns()
    End Sub

    Private Sub LoadColumns()
        Me.Columns.Clear()


        With colSortNumber
            .Name = "SortNumber"
            .HeaderText = LANG_Number
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .Frozen = True
        End With
        Me.Columns.Add(colSortNumber)

        With colRTF
            .Name = "RTF"
            .HeaderText = LANG_Description
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .Frozen = True
            .MinimumWidth = 200
            .FillWeight = 100
        End With
        Me.Columns.Add(colRTF)

        With colRange
            .Name = "Range"
            .HeaderText = LANG_Range
            .MinimumWidth = 100
            .FillWeight = 20
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        End With
        Columns.Add(colRange)

        'LoadTargetColumns()
    End Sub

    Private Sub LoadTargetColumns()
        Me.TargetDeadlines = CurrentLogFrame.GetTargetDeadlines(2)

        If Columns.Count > 3 Then
            For i = 3 To Columns.Count - 1
                Columns.RemoveAt(3)
            Next
        End If

        For Each selDeadline As TargetDeadline In Me.TargetDeadlines
            Dim colTarget As New DataGridViewTextBoxColumn
            With colTarget
                .Name = selDeadline.Deadline
                .HeaderText = String.Format("{0} - {1}", LANG_Target, selDeadline.Deadline.ToShortDateString)
                .MinimumWidth = 100
                .FillWeight = 20
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            End With
            Columns.Add(colTarget)
        Next
    End Sub

    'Public Sub ShowColumns()
    '    colStartDate.Visible = Me.ShowDatesColumns
    '    colEndDate.Visible = Me.ShowDatesColumns

    '    Invalidate()
    'End Sub

    Public Sub Reload()
        Me.SuspendLayout()
        LoadTargetColumns()
        LoadSections()

        Me.RowCount = Me.Grid.Count
        Me.RowModifyIndex = -1
        AutoSizeSortColumn()

        ResetAllImages()
        ReloadImages()
        SectionTitles_Protect()
        Me.Invalidate()
        Me.ResumeLayout()

        RaiseEvent Reloaded()
    End Sub

    Private Sub AutoSizeSortColumn()
        Dim intWidth As Integer = 0

        colSortNumber.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        intWidth = colSortNumber.Width
        colSortNumber.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colSortNumber.Width = intWidth
    End Sub

    Public Sub SectionTitles_Protect()
        For Each selRow As DataGridViewRow In Me.Rows
            If Grid(selRow.Index).RowType <> QuestionnaireGridRow.RowTypes.Indicator Then
                'make title bar of each section and repeated objectives or outputs inaccessible
                For Each selCell As DataGridViewCell In selRow.Cells
                    selCell.ReadOnly = True
                Next
            Else
                'make numbers inaccessible
                selRow.Cells(0).ReadOnly = True
                For i = 1 To selRow.Cells.Count - 1
                    selRow.Cells(i).ReadOnly = False
                Next
            End If
        Next
    End Sub
#End Region

#Region "Cell images"
    Private Sub ResetAllImages()
        For Each selRow As QuestionnaireGridRow In Me.Grid
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As QuestionnaireGridRow)
        selRow.StructImageDirty = True
        selRow.StructHeight = 0
    End Sub

    Private Sub ReloadImages()
        For Each selRow As QuestionnaireGridRow In Me.Grid
            If selRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
                ReloadImages_Indicator(selRow)
            End If
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Indicator(ByVal selRow As QuestionnaireGridRow)
        Dim intColumnWidth As Integer
        Dim selIndicator As Indicator = selRow.Indicator
        If selIndicator Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.IndicatorImageDirty = True Then
                intColumnWidth = colRTF.Width

                If String.IsNullOrEmpty(selIndicator.Text) Then
                    selRow.IndicatorImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, Indicator.ItemName, selRow.SortNumber, False)
                Else
                    selRow.IndicatorImage = .RichTextWithPaddingToBitmap(intColumnWidth, selIndicator.RTF, False, selRow.Indent)
                End If
                selRow.IndicatorHeight = selRow.IndicatorImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Load Indicators"
    Private Sub LoadSections()
        Me.Grid.Clear()
        If CurrentLogFrame IsNot Nothing Then
            If ShowGoals = False And ShowPurposes = False And ShowOutputs = False And ShowActivities = False Then
                ShowPurposes = True
                frmParent.RibbonButtonShowPurposes.Checked = True
            End If

            'Section: Goals
            If ShowGoals = True Then LoadSections_SectionGoals()

            'Section: Purpose(s)
            If ShowPurposes = True Then LoadSections_SectionPurposes()

            'Section: Outputs
            If ShowOutputs = True Then LoadSections_SectionOutputs()

            'Section: Activities
            If ShowActivities = True Then LoadSections_SectionActivities()
        End If

        Me.RowCount = Me.Grid.Count
    End Sub

    Private Sub LoadSections_SectionGoals()
        Dim rowGoalSection As QuestionnaireGridRow = LoadSections_CreateSectionTitle(LogFrame.SectionTypes.GoalsSection)
        Me.Grid.Add(rowGoalSection)

        For Each selGoal As Goal In CurrentLogFrame.Goals
            Dim strGoalSort As String = CurrentLogFrame.GetStructSortNumber(selGoal)
            Dim rowGoal As New QuestionnaireGridRow
            With rowGoal
                .SortNumber = strGoalSort
                .Struct = selGoal
                .RowType = QuestionnaireGridRow.RowTypes.Goal
            End With
            Me.Grid.Add(rowGoal)

            LoadSections_Indicators(selGoal.Indicators, strGoalSort)
        Next
    End Sub

    Private Sub LoadSections_SectionPurposes()
        Dim rowPurposeSection As QuestionnaireGridRow = LoadSections_CreateSectionTitle(LogFrame.SectionTypes.PurposesSection)
        Me.Grid.Add(rowPurposeSection)

        For Each selPurpose As Purpose In CurrentLogFrame.Purposes
            Dim strPurposeSort As String = CurrentLogFrame.GetStructSortNumber(selPurpose)
            Dim rowPurpose As New QuestionnaireGridRow
            With rowPurpose
                .SortNumber = strPurposeSort
                .Struct = selPurpose
                .RowType = QuestionnaireGridRow.RowTypes.Purpose
            End With
            Me.Grid.Add(rowPurpose)

            LoadSections_Indicators(selPurpose.Indicators, strPurposeSort)
        Next
    End Sub

    Private Sub LoadSections_SectionOutputs()
        Dim rowOutputSection As QuestionnaireGridRow = LoadSections_CreateSectionTitle(LogFrame.SectionTypes.OutputsSection)
        Me.Grid.Add(rowOutputSection)

        For Each selPurpose As Purpose In CurrentLogFrame.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                Dim strOutputSort As String = CurrentLogFrame.GetStructSortNumber(selOutput)
                Dim rowOutput As New QuestionnaireGridRow
                With rowOutput
                    .SortNumber = strOutputSort
                    .Struct = selOutput
                    .RowType = QuestionnaireGridRow.RowTypes.Output
                End With
                Me.Grid.Add(rowOutput)

                LoadSections_Indicators(selOutput.Indicators, strOutputSort)
            Next
        Next
    End Sub

    Private Sub LoadSections_SectionActivities()
        Dim rowActivitySection As QuestionnaireGridRow = LoadSections_CreateSectionTitle(LogFrame.SectionTypes.ActivitiesSection)
        Me.Grid.Add(rowActivitySection)

        For Each selPurpose As Purpose In CurrentLogFrame.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                For Each selActivity As Activity In selOutput.Activities
                    Dim strActivitySort As String = CurrentLogFrame.GetStructSortNumber(selActivity)
                    Dim rowActivity As New QuestionnaireGridRow
                    With rowActivity
                        .SortNumber = strActivitySort
                        .Struct = selActivity
                        .RowType = QuestionnaireGridRow.RowTypes.Activity
                    End With
                    Me.Grid.Add(rowActivity)

                    LoadSections_Indicators(selActivity.Indicators, strActivitySort)
                Next
            Next
        Next
    End Sub

    Private Function LoadSections_CreateSectionTitle(ByVal intSection As Integer) As QuestionnaireGridRow
        Dim TitleRow As New QuestionnaireGridRow
        Dim strStruct As String = String.Empty
        Dim objStruct As Struct = Nothing

        Select Case intSection
            Case LogFrame.SectionTypes.GoalsSection
                strStruct = CurrentLogFrame.StructNamePlural(0)
                objStruct = New Goal(LoadSections_SectionTitleSetFont(strStruct))
            Case LogFrame.SectionTypes.PurposesSection
                strStruct = CurrentLogFrame.StructNamePlural(1)
                objStruct = New Purpose(LoadSections_SectionTitleSetFont(strStruct))
            Case LogFrame.SectionTypes.OutputsSection
                strStruct = CurrentLogFrame.StructNamePlural(2)
                objStruct = New Output(LoadSections_SectionTitleSetFont(strStruct))
            Case LogFrame.SectionTypes.ActivitiesSection
                strStruct = CurrentLogFrame.StructNamePlural(3)
                objStruct = New Activity(LoadSections_SectionTitleSetFont(strStruct))
        End Select

        With TitleRow
            .RowType = QuestionnaireGridRow.RowTypes.Section
            .Section = intSection
            .Struct = objStruct
        End With

        Return TitleRow
    End Function

    Private Function LoadSections_SectionTitleSetFont(ByVal strText As String) As String
        Using objRtf As New RichTextManager
            objRtf.Text = strText
            objRtf.SelectAll()
            objRtf.SelectionFont = CurrentLogFrame.SectionTitleFont
            objRtf.SelectionColor = CurrentLogFrame.SectionTitleFontColor
            objRtf.SelectionAlignment = HorizontalAlignment.Center
            strText = objRtf.Rtf
        End Using
        Return strText
    End Function

    Private Sub LoadSections_Indicators(ByVal objIndicators As Indicators, ByVal strParentSort As String, Optional ByVal intLevel As Integer = 0)
        Dim intIndex As Integer

        For Each selIndicator As Indicator In objIndicators
            intIndex = objIndicators.IndexOf(selIndicator)

            If selIndicator.Registration <= ShowRegistrationOption Then
                If ShowTargetGroup <> Guid.Empty Then
                    If selIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel AndAlso selIndicator.TargetGroupGuid = ShowTargetGroup Then
                        LoadSections_Indicators_Add(selIndicator, strParentSort, intIndex, intLevel)
                    End If
                Else
                    LoadSections_Indicators_Add(selIndicator, strParentSort, intIndex, intLevel)
                End If
            End If
        Next
    End Sub

    Private Sub LoadSections_Indicators_Add(ByVal selIndicator As Indicator, ByVal strParentSort As String, ByVal intIndex As Integer, ByVal intLevel As Integer)
        Dim rowIndicator As New QuestionnaireGridRow
        Dim strIndicatorSort As String

        If CurrentLogFrame.SortNumberRepeatParent = True Then strIndicatorSort = strParentSort Else strIndicatorSort = String.Empty
        strIndicatorSort = CurrentLogFrame.CreateSortNumber(intIndex, strParentSort)

        With rowIndicator
            .SortNumber = strIndicatorSort
            .Indicator = selIndicator
            .Indent = intLevel
            .RowType = QuestionnaireGridRow.RowTypes.Indicator
        End With
        Me.Grid.Add(rowIndicator)

        If selIndicator.Indicators.Count > 0 Then
            LoadSections_Indicators(selIndicator.Indicators, strIndicatorSort, intLevel + 1)
        End If
    End Sub

    Private Function GetRowIndexByReferenceGuid(ByVal selGuid As Guid) As Integer
        Dim intRowIndex As Integer = -1

        For Each selGridRow As QuestionnaireGridRow In Me.Grid
            Select Case selGridRow.RowType
                Case QuestionnaireGridRow.RowTypes.Indicator
                    Dim selIndicator As Indicator = selGridRow.Indicator
                    If selIndicator IsNot Nothing AndAlso selIndicator.Guid = selGuid Then
                        intRowIndex = Grid.IndexOf(selGridRow)
                        Exit For
                    End If
            End Select
        Next

        Return intRowIndex
    End Function
#End Region 'load Indicators/key moments

#Region "Events"
    Private Sub RichTextEditingControl_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs) Handles RichTextEditingControl.ContentsResized
        Dim intRequiredHeight As Integer = e.NewRectangle.Height + RichTextEditingControl.Margin.Vertical + SystemInformation.VerticalResizeBorderThickness
        If CurrentRow.Height < intRequiredHeight Then CurrentRow.Height = intRequiredHeight
    End Sub

    Private Sub RichTextEditingControl_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextEditingControl.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim pClick As New Point(CurrentCell.ContentBounds.X, CurrentCell.ContentBounds.Y)
            pClick.X += e.X + CONST_Padding
            pClick.Y += e.Y + CONST_Padding

            Dim dragSize As Size = SystemInformation.DragSize
            DragBoxFromMouseDown = New Rectangle(New Point(Me.Location.X + e.X - (dragSize.Width / 2), _
                                                           Me.Location.Y + e.Y - (dragSize.Height / 2)), dragSize)
            If SelectionRectangle.Rectangle.Contains(pClick.X, pClick.Y) Then DragMultipleCells = True Else DragMultipleCells = False
        Else
            DragBoxFromMouseDown = Rectangle.Empty
        End If
    End Sub

    Private Sub RichTextEditingControl_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextEditingControl.TextChanged
        If RichTextEditingControl.Text.Length > 0 Then
            Me.RowModifyIndex = Me.CurrentRow.Index
        Else
            Me.RowModifyIndex = -1
        End If
    End Sub

    Protected Overrides Sub OnEditingControlShowing(ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        MyBase.OnEditingControlShowing(e)

        Dim selGridRow As QuestionnaireGridRow = Me.Grid(CurrentRow.Index)

        'If selGridRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then _
        '    CurrentUndoList.InitialiseUndoBuffer(selGridRow.Indicator, selGridRow.SortNumber)
    End Sub
#End Region

#Region "General methods"
    Public Sub AddNewItem()
        If CurrentCell Is Nothing Then Exit Sub

        Dim intRowIndex As Integer
        Dim intColIndex = 1
        Dim strColName As String = Me.Columns(intColIndex).Name
        Dim selGridRow As QuestionnaireGridRow

        For intRowIndex = CurrentCell.RowIndex To RowCount - 1
            selGridRow = Grid(intRowIndex)
            Select Case selGridRow.RowType
                Case QuestionnaireGridRow.RowTypes.Indicator
                    If selGridRow.Indicator Is Nothing Then Exit For
            End Select
        Next

        CurrentCell = Me(intColIndex, intRowIndex)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selObject As Object = Me.CurrentItem(False)
        Dim intIndex As Integer

        If selObject Is Nothing Then Exit Sub

        Select Case selObject.GetType
            Case GetType(Indicator)
                Dim selIndicator As Indicator = DirectCast(selObject, Indicator)
                Dim selIndicators As Indicators = CurrentLogFrame.GetParentCollection(selIndicator)

                If selIndicators IsNot Nothing Then
                    intIndex = selIndicators.IndexOf(selIndicator)

                    Dim NewIndicator As New Indicator
                    selIndicators.Insert(intIndex, NewIndicator) 'IntoProcess
                    selObject = NewIndicator
                End If
        End Select

        Reload()

        SetFocusOnItem(selObject)

        Me.BeginEdit(False)
    End Sub

    Public Sub InsertParentItem()
        Dim selIndicator As Indicator
        Dim ParentIndicators As Indicators = Nothing

        selIndicator = TryCast(Me.CurrentItem(False), Indicator)

        If selIndicator IsNot Nothing Then
            ParentIndicators = CurrentLogFrame.GetParentCollection(selIndicator)

            If ParentIndicators IsNot Nothing Then
                Dim intIndex As Integer = ParentIndicators.IndexOf(selIndicator)
                Dim NewParent As New Indicator

                ParentIndicators.Remove(selIndicator)
                'CurrentUndoList.CutOperation(selIndicator, ParentIndicators, intIndex, , True)

                ParentIndicators.Insert(intIndex, NewParent)
                'CurrentUndoList.InsertNewOperation(NewParent, ParentIndicators, , , True)

                NewParent.Indicators.Add(selIndicator)
                'CurrentUndoList.PasteOperation(selIndicator, NewParent.Indicators, 0, , True)

                Reload()
                SetFocusOnItem(NewParent)
                Me.BeginEdit(False)
            End If
        End If
    End Sub

    Public Sub InsertChildItem()
        Dim selIndicator As Indicator = TryCast(Me.CurrentItem(False), Indicator)

        If selIndicator IsNot Nothing Then
            Dim ChildIndicator As New Indicator

            selIndicator.Indicators.Insert(0, ChildIndicator)
            'CurrentUndoList.InsertNewOperation(ChildIndicator, selIndicator.Indicators, , , True)
            Reload()
            SetFocusOnItem(ChildIndicator)
            Me.BeginEdit(False)
        End If
    End Sub

    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selGridRow As QuestionnaireGridRow = Grid(CurrentCell.RowIndex)
        Dim selItem As Object = Me.CurrentItem(False)

        If selItem Is Nothing Then Exit Sub

        If selGridRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
            If intDirection < 0 Then
                selItem = MoveIndicator_ToPreviousParent(selItem)
            Else
                selItem = MoveIndicator_ToNextParent(selItem)
            End If
        End If

        Reload()
        SetFocusOnItem(selItem)
    End Sub

    Private Function MoveIndicator_ToPreviousParent(ByVal selIndicator As Indicator) As Indicator
        Dim objIndicators As Indicators = CurrentLogFrame.GetParentCollection(selIndicator)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim PreviousIndicator As Indicator = Me.Grid.GetPreviousIndicator(intRowIndex)
        Dim intIndex As Integer

        If PreviousIndicator Is Nothing Then Return selIndicator

        Select Case PreviousIndicator.GetType
            Case GetType(Indicator)
                Dim objPreviousIndicators As Indicators = CurrentLogFrame.GetParentCollection(PreviousIndicator)
                intIndex = objPreviousIndicators.IndexOf(PreviousIndicator)
                If intIndex = objPreviousIndicators.Count - 1 Then intIndex += 1
                objIndicators.Remove(selIndicator)
                objPreviousIndicators.Insert(intIndex, selIndicator)

            Case GetType(Output)
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToPreviousParent(selIndicator)

            Case GetType(Purpose)
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToPreviousParent(selIndicator)
        End Select

        Return selIndicator
    End Function

    Private Function MoveIndicator_ToNextParent(ByVal selIndicator As Indicator) As Indicator
        Dim objIndicators As Indicators = CurrentLogFrame.GetParentCollection(selIndicator)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As QuestionnaireGridRow = Me.Grid(intRowIndex)
        Dim NextIndicator As Indicator = Me.Grid.GetNextIndicator(intRowIndex)
        Dim intIndex As Integer

        If NextIndicator Is Nothing Then Return selIndicator

        Select Case NextIndicator.GetType
            Case GetType(Indicator)
                If CurrentLogFrame.IsParentLineage(NextIndicator, selIndicator) Then
                    intRowIndex += 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    MoveIndicator_ToNextParent(selIndicator)
                Else
                    Dim objNextIndicators As Indicators
                    If CType(NextIndicator, Indicator).Indicators.Count > 0 Then
                        'if the next Indicator has sub-Indicators, insert as first sub-Indicator
                        objNextIndicators = CType(NextIndicator, Indicator).Indicators
                        intIndex = 0
                    Else
                        'insert before the next Indicator (at the same level)
                        objNextIndicators = CurrentLogFrame.GetParentCollection(NextIndicator)
                        intIndex = objNextIndicators.IndexOf(NextIndicator)
                    End If
                    objIndicators.Remove(selIndicator)
                    objNextIndicators.Insert(intIndex, selIndicator)
                End If

            Case GetType(Output)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToNextParent(selIndicator)

            Case GetType(Purpose)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToNextParent(selIndicator)
        End Select

        Return selIndicator
    End Function

    Public Sub ChangeSection(ByVal intDirection As Integer)
        If CurrentCell Is Nothing Then Exit Sub

        Dim selItem As Object = Nothing
        Dim selGridRow As QuestionnaireGridRow = Grid(CurrentCell.RowIndex)


        Select Case selGridRow.RowType
            Case QuestionnaireGridRow.RowTypes.Indicator
                Dim selIndicator As Indicator = selGridRow.Indicator
                If selIndicator IsNot Nothing Then _
                    selItem = ChangeSection_Indicator(intDirection, selIndicator)
        End Select

        Reload()
        If selItem IsNot Nothing Then SetFocusOnItem(selItem)
        CurrentRow.Selected = True
    End Sub

    Public Function ChangeSection_Indicator(ByVal intDirection As Integer, ByVal selIndicator As Indicator) As Object
        'Dim intIndex As Integer, intParentIndex As Integer
        'Dim RootIndicator As Indicator = CurrentLogFrame.GetRootIndicator(selIndicator)
        'Dim selIndicators As Indicators = CurrentLogFrame.GetParentCollection(selIndicator)
        'Dim selParentOutput As Output = CurrentLogFrame.GetOutputByGuid(RootIndicator.ParentOutputGuid)
        'Dim selParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selParentOutput.ParentPurposeGuid)

        'intIndex = selParentOutput.Indicators.IndexOf(RootIndicator)
        'intParentIndex = selParentPurpose.Outputs.IndexOf(selParentOutput)

        'If intDirection < 0 Then
        '    intParentIndex -= 1
        '    If intParentIndex >= 0 Then
        '        selIndicators.Remove(selIndicator)
        '        selParentOutput = selParentPurpose.Outputs(intParentIndex)
        '        If intIndex > selParentOutput.Indicators.Count Then intIndex = selParentOutput.Indicators.Count
        '        If intIndex < 0 Then intIndex = 0
        '        selParentOutput.Indicators.Insert(intIndex, selIndicator)
        '    Else
        '        Dim intPurposeIndex As Integer = CurrentLogFrame.Purposes.IndexOf(selParentPurpose)
        '        intPurposeIndex -= 1
        '        If intPurposeIndex >= 0 Then
        '            selParentPurpose = CurrentLogFrame.Purposes(intPurposeIndex)
        '            If selParentPurpose.Outputs.Count > 0 Then
        '                selIndicators.Remove(selIndicator)
        '                selParentOutput = selParentPurpose.Outputs(selParentPurpose.Outputs.Count - 1)
        '                If intIndex > selParentOutput.Indicators.Count Then intIndex = selParentOutput.Indicators.Count
        '                If intIndex < 0 Then intIndex = 0
        '                selParentOutput.Indicators.Insert(intIndex, selIndicator)
        '            End If
        '        End If
        '    End If
        'Else
        '    intParentIndex += 1
        '    If intParentIndex <= selParentPurpose.Outputs.Count - 1 Then
        '        selIndicators.Remove(selIndicator)
        '        selParentOutput = selParentPurpose.Outputs(intParentIndex)
        '        If intIndex > selParentOutput.Indicators.Count Then intIndex = selParentOutput.Indicators.Count
        '        If intIndex < 0 Then intIndex = 0
        '        selParentOutput.Indicators.Insert(intIndex, selIndicator)
        '    Else
        '        Dim intPurposeIndex As Integer = CurrentLogFrame.Purposes.IndexOf(selParentPurpose)
        '        intPurposeIndex += 1
        '        If intPurposeIndex <= CurrentLogFrame.Purposes.Count - 1 Then
        '            selParentPurpose = CurrentLogFrame.Purposes(intPurposeIndex)
        '            If selParentPurpose.Outputs.Count > 0 Then
        '                selIndicators.Remove(selIndicator)
        '                selParentOutput = selParentPurpose.Outputs(selParentPurpose.Outputs.Count - 1)
        '                If intIndex > selParentOutput.Indicators.Count Then intIndex = selParentOutput.Indicators.Count
        '                If intIndex < 0 Then intIndex = 0
        '                selParentOutput.Indicators.Insert(intIndex, selIndicator)
        '            End If
        '        End If
        '    End If
        'End If

        Return Nothing 'Return selIndicator
    End Function

    Public Sub LevelUp()
        Dim selIndicator As Indicator = TryCast(Me.CurrentItem(False), Indicator)
        If selIndicator Is Nothing Then Exit Sub

        Dim ParentIndicator As Indicator = TryCast(CurrentLogFrame.GetParent(selIndicator), Indicator)
        If ParentIndicator Is Nothing Then Exit Sub

        Dim objParentIndicators As Indicators = CurrentLogFrame.GetParentCollection(ParentIndicator)
        Dim intIndex As Integer = objParentIndicators.IndexOf(ParentIndicator)

        intIndex += 1
        ParentIndicator.Indicators.Remove(selIndicator)
        objParentIndicators.Insert(intIndex, selIndicator)

        Reload()
        SetFocusOnItem(selIndicator)
    End Sub

    Public Sub LevelDown()
        Dim selIndicator As Indicator = TryCast(Me.CurrentItem(False), Indicator)
        If selIndicator Is Nothing Then Exit Sub

        Dim objIndicators As Indicators = CurrentLogFrame.GetParentCollection(selIndicator)
        Dim intIndex As Integer = objIndicators.IndexOf(selIndicator)
        If intIndex = 0 Then Exit Sub

        intIndex -= 1
        objIndicators.Remove(selIndicator)
        Dim PreviousIndicator As Indicator = objIndicators(intIndex)
        PreviousIndicator.Indicators.Add(selIndicator)

        Reload()
        SetFocusOnItem(selIndicator)
    End Sub

    Private Sub SetFocusOnItem(ByVal selItem As Object)
        Dim selGridRow As QuestionnaireGridRow
        Dim intColIndex As Integer = 1
        Dim intRowIndex As Integer

        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            Select Case selGridRow.RowType
                Case QuestionnaireGridRow.RowTypes.Indicator
                    If selGridRow.Indicator Is selItem Then
                        Exit For
                    End If
            End Select
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub
#End Region

#Region "Mouse actions"
    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)

        If CurrentCell.IsInEditMode = False Then
            MoveCurrentCell()
            InvalidateSelectionRectangle()
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Tab Then

            Dim x As Integer = CurrentCell.ColumnIndex
            Dim y As Integer = CurrentRow.Index
            Dim boolShift As Boolean = e.Shift

            If y = Me.RowCount - 1 Then
                CurrentCell = Me(x, y - 1)
            Else
                CurrentCell = Me(x, y + 1)
            End If

            CurrentCell = Me(x, y)

            MoveCurrentCell()
            CurrentCell.Selected = True

            InvalidateSelectionRectangle()
        End If
        MyBase.OnKeyUp(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If hit.ColumnIndex > 0 And hit.RowIndex >= 0 Then
            Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
            Dim selGridRow As QuestionnaireGridRow = Me.Grid(hit.RowIndex)

            If e.Button = MouseButtons.Right Then
                If hit.Type = DataGridViewHitTestType.Cell Then
                    ' Create a rectangle using the DragSize, with the mouse position being at the center of the rectangle.
                    OnMouseDown_SetDragRectangle(e.Location, selCell)
                End If
            Else
                If selCell.ColumnIndex = 1 Then
                    'determine where exactly in the text the user clicked
                    OnMouseDown_SetClickPoint(e.Location)
                End If

                DragBoxFromMouseDown = Rectangle.Empty
            End If
        End If
        MyBase.OnMouseDown(e)
        Invalidate()
    End Sub

    Private Sub OnMouseDown_SetDragRectangle(ByVal MouseLocation As Point, ByVal selCell As DataGridViewCell)
        Dim dragSize As Size = SystemInformation.DragSize
        DragBoxFromMouseDown = New Rectangle(New Point(MouseLocation.X - (dragSize.Width / 2), MouseLocation.Y - (dragSize.Height / 2)), dragSize)

        If SelectionRectangle.Rectangle.Contains(MouseLocation) = False Then Me.CurrentCell = selCell
    End Sub

    Private Sub OnMouseDown_SetClickPoint(ByVal MouseLocation As Point)
        If SelectionRectangle.Rectangle.Contains(MouseLocation) = True Then
            ClickPoint = MouseLocation
        Else
            ClickPoint = Nothing
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim ptMouseLocation As New Point(e.X, e.Y)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If ((e.Button And MouseButtons.Right) = MouseButtons.Right) Then
            'drag items
            OnMouseMove_DragItems(e.Location)
        End If

        MyBase.OnMouseMove(e)

        InvalidateSelectionRectangle()
    End Sub

    Private Sub OnMouseMove_DragItems(ByVal MouseLocation As Point)
        Dim p1 As New Point(Me.Location.X, Me.Location.Y)
        Dim p2 As New Point(Me.Location.X, Me.Location.Y + Me.Height)

        'if cursor nears the rim of the datagridview while dragging, scroll up or down
        If MouseLocation.Y < p1.Y + 50 And FirstDisplayedScrollingRowIndex > 0 Then _
            FirstDisplayedScrollingRowIndex -= 1
        If MouseLocation.Y > p2.Y - 50 And Rows(Me.RowCount - 1).Displayed = False Then _
            FirstDisplayedScrollingRowIndex += 1

        'select cursor
        If Control.ModifierKeys = Keys.Control Then
            Cursor.Current = Cursors.HSplit
            DragOperator = DragOperatorValues.Copy
        Else
            Cursor.Current = Cursors.Hand
            DragOperator = DragOperatorValues.Move
        End If

        Drag(MouseLocation)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim ptMouseLocation As New Point(e.X, e.Y)

        DragReleased = True

        InvalidateSelectionRectangle()
        If e.Button = MouseButtons.Right Then
            If Control.ModifierKeys = Keys.Control Then
                DragOperator = DragOperatorValues.Copy
            Else
                DragOperator = DragOperatorValues.Move
            End If



        End If
        DragBoxFromMouseDown = Rectangle.Empty
        DragIndicator = Nothing

        MyBase.OnMouseUp(e)
        Invalidate()
    End Sub

    Public Sub Drag(ByVal MouseLocation As Point)
        ' If the mouse moves outside the rectangle, start the drag.
        If (Rectangle.op_Inequality(DragBoxFromMouseDown, Rectangle.Empty) And _
            Not DragBoxFromMouseDown.Contains(MouseLocation)) Then

            If DragReleased = True Then
                If DragOperator = DragOperatorValues.Copy Then
                    CopyItems()
                Else
                    CutItems(False)
                End If
                DragReleased = False
            End If

            Dim hit As DataGridView.HitTestInfo = HitTest(MouseLocation.X, MouseLocation.Y)
            If hit.Type = DataGridViewHitTestType.Cell Then
                Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
                Dim intLastColIndex As Integer = 1

                'If Me.ShowDatesColumns = True Then intLastColIndex = 3

                SelectionRectangle.FirstRowIndex = selCell.RowIndex
                SelectionRectangle.LastRowIndex = selCell.RowIndex
                If SelectionRectangle.LastRowIndex > Me.RowCount - 1 Then SelectionRectangle.LastRowIndex = Me.RowCount - 1

                'draw insert/copy indicator line
                Dim rSelStart As Rectangle = GetCellDisplayRectangle(0, SelectionRectangle.FirstRowIndex, False)
                Dim rSelEnd As Rectangle = GetCellDisplayRectangle(intLastColIndex, SelectionRectangle.LastRowIndex, False)

                Dim intVertDivider As Integer = rSelStart.Top + ((rSelEnd.Bottom - rSelStart.Top) / 2)
                pStartInsertLine.X = rSelStart.Left
                pEndInsertLine.X = rSelEnd.Right - 1
                If MouseLocation.Y <= intVertDivider Then pStartInsertLine.Y = rSelStart.Top Else pStartInsertLine.Y = rSelEnd.Bottom
                pEndInsertLine.Y = pStartInsertLine.Y

                Dim graph As Graphics = CreateGraphics()
                If Not (pStartInsertLine = Nothing Or pEndInsertLine = Nothing) Then
                    If pStartInsertLine <> pStartInsertLineOld Then
                        Invalidate()
                        Update()

                        Dim penGreen2 As New Pen(Color.Green, 2)
                        graph.DrawLine(penGreen2, pStartInsertLine, pEndInsertLine)
                        Dim triangleStart(2) As Point
                        triangleStart(0) = New Point(pStartInsertLine.X, pStartInsertLine.Y - 6)
                        triangleStart(1) = New Point(pStartInsertLine.X, pStartInsertLine.Y + 6)
                        triangleStart(2) = New Point(pStartInsertLine.X + 6, pStartInsertLine.Y)
                        graph.FillPolygon(Brushes.Green, triangleStart)
                        Dim triangleEnd(2) As Point
                        triangleEnd(0) = New Point(pEndInsertLine.X, pEndInsertLine.Y - 6)
                        triangleEnd(1) = New Point(pEndInsertLine.X, pEndInsertLine.Y + 6)
                        triangleEnd(2) = New Point(pEndInsertLine.X - 8, pEndInsertLine.Y)
                        graph.FillPolygon(Brushes.Green, triangleEnd)
                        pStartInsertLineOld = pStartInsertLine
                        pEndInsertLineOld = pEndInsertLine
                    End If
                End If

            Else
                pStartInsertLine = Nothing
                pEndInsertLine = Nothing
            End If
        End If
    End Sub

    Public Sub Drop(ByVal mouseX As Integer, ByVal mouseY As Integer)
        Dim hit As DataGridView.HitTestInfo = HitTest(pStartInsertLine.X, pStartInsertLine.Y)

        If hit.Type = DataGridViewHitTestType.Cell Then
            Dim CopyGroup As ClipboardItems = ItemClipboard.GetCopyGroup()
            Dim selCell As DataGridViewCell

            If hit.RowIndex < Me.RowCount - 1 Then
                selCell = Me(hit.ColumnIndex, hit.RowIndex)
            Else
                selCell = Me(hit.ColumnIndex, Me.RowCount - 2)
            End If

            PasteItems(CopyGroup, ClipboardItems.PasteOptions.PasteAll, selCell)
        End If
    End Sub

    Private Sub MoveCurrentCell()
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim CurrentGridRow As QuestionnaireGridRow = Me.Grid(CurrentCell.RowIndex)
        Dim selIndicator As Indicator = CurrentGridRow.Indicator
        Dim strColumnName As String = Me.Columns(intColumnIndex).Name

        If CurrentGridRow.RowType <> QuestionnaireGridRow.RowTypes.Indicator Then
            intRowIndex += 1
            Me.CurrentCell = Me(intColumnIndex, intRowIndex)
            MoveCurrentCell()
        End If

        Select Case strColumnName
            Case "SortNumber"
                intColumnIndex += 1
                Me.CurrentCell = Me(intColumnIndex, intRowIndex)
                MoveCurrentCell()
            Case "RTF"
                Me.CurrentCell = Me(intColumnIndex, intRowIndex)

                If ClickPoint.IsEmpty = False Then
                    Me.BeginEdit(False)

                    Dim rCell As Rectangle = Me.GetCellDisplayRectangle(intColumnIndex, intRowIndex, False)
                    ClickPoint.X -= rCell.X
                    ClickPoint.Y -= rCell.Y

                    If RichTextEditingControl IsNot Nothing Then
                        With RichTextEditingControl
                            Dim intCharIndex As Integer = .GetCharIndexFromPosition(ClickPoint)
                            .Select(intCharIndex, 0)
                            .SetCurrentText()

                            ClickPoint = Nothing
                        End With
                    End If
                End If
        End Select

        RaiseEvent IndicatorSelected(Me, New IndicatorSelectedEventArgs(selIndicator))
    End Sub

    Protected Overrides Sub OnColumnWidthChanged(ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
        If Me.IsCurrentCellInEditMode = False Then
            MyBase.OnColumnWidthChanged(e)

            boolColumnWidthChanged = True
        End If
    End Sub

    Protected Overrides Sub OnScroll(ByVal e As System.Windows.Forms.ScrollEventArgs)
        MyBase.OnScroll(e)

        InvalidateSelectionRectangle()
    End Sub
#End Region

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)
        CopyItems()

        RemoveItems(ShowWarning)
    End Sub

    Public Overrides Sub CopyItems()
        Dim selRow As QuestionnaireGridRow
        Dim strSort As String
        Dim CopyGroup As Date = Now()

        With SelectionRectangle
            For i = .FirstRowIndex To .LastRowIndex
                selRow = Me.Grid(i)
                If selRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
                    If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Indicator Then
                        strSort = Indicator.ItemName & " " & selRow.SortNumber
                        Dim NewItem As New ClipboardItem(CopyGroup, selRow.Indicator, strSort)
                        ItemClipboard.Insert(0, NewItem)
                    End If
                End If
            Next
        End With
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
        If PasteCell Is Nothing Then PasteCell = CurrentCell
        If PasteCell Is Nothing Then Exit Sub

        Dim intColumnIndex As Integer = PasteCell.ColumnIndex
        Dim intRowIndex As Integer = PasteCell.RowIndex

        PasteRow = Grid(intRowIndex)
        Dim selItem As ClipboardItem

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Indicator)
                    PasteItems_Indicator(selItem, intPasteOption)
                Case Else
                    PasteItems_Other(selItem, intPasteOption)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(intColumnIndex, intRowIndex)
    End Sub

    Private Sub PasteIndicator(ByVal NewIndicator As Indicator)
        Dim ParentIndicators As Indicators
        Dim PasteRowIndicator As Indicator = PasteRow.Indicator
        Dim intIndex As Integer
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRowIndicator IsNot Nothing Then
            ParentIndicators = CurrentLogFrame.GetParentCollection(PasteRowIndicator)
            intIndex = ParentIndicators.IndexOf(PasteRowIndicator)
        Else
            Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))
            If PreviousStruct Is Nothing Then Exit Sub
            ParentIndicators = PreviousStruct.Indicators

            intIndex = ParentIndicators.Count
        End If
        ParentIndicators.Insert(intIndex, NewIndicator)

        PasteRow.Indicator = NewIndicator

        'CurrentUndoList.PasteOperation(NewIndicator, ParentIndicators, intIndex, strSortNumber, True)
    End Sub

    Private Sub PasteItems_Indicator(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selIndicator As Indicator = DirectCast(selItem.Item, Indicator)

        If PasteRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
            Dim NewIndicator As New Indicator

            Using copier As New ObjectCopy
                NewIndicator = copier.CopyObject(selIndicator)
            End Using

            PasteIndicator(NewIndicator)
        End If
    End Sub

    Private Sub PasteItems_Other(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selObject As LogframeObject = TryCast(selItem.Item, LogframeObject)
        Dim strText As String = String.Empty, strRtf As String = String.Empty

        If selObject IsNot Nothing Then
            strText = selObject.Text
            strRtf = selObject.RTF
        Else
            strText = selItem.ToString
        End If
        Dim strSortNumber As String = PasteRow.SortNumber

        If PasteRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
            Dim NewIndicator As New Indicator

            If String.IsNullOrEmpty(strRtf) Then
                NewIndicator.SetText(strText)
            Else
                NewIndicator.RTF = strRtf
            End If

            PasteIndicator(NewIndicator)
        End If
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim strSourceColName As String
        Dim boolShift As Boolean
        Dim boolRemoveAll As Boolean
        Dim intRowIndex, intColumnIndex As Integer
        Dim strSortNumber As String = String.Empty
        Dim objIndicator As Indicator
        Dim objIndicators As Indicators

        intRowIndex = CurrentCell.RowIndex
        intColumnIndex = CurrentCell.ColumnIndex
        If Control.ModifierKeys = Keys.Shift Then boolShift = True

        'copy cells to delete
        strSourceColName = Columns(SelectionRectangle.FirstColumnIndex).Name
        For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
            SelectedGridRows.Add(Me.Grid(i))
        Next

        If ShowWarning = True Then
            Dim boolCancel As Boolean = RemoveItems_Warning(strSourceColName)
            If boolCancel = True Then Exit Sub
        End If

        For Each selGridRow As QuestionnaireGridRow In SelectedGridRows
            strSortNumber = selGridRow.SortNumber

            If selGridRow.RowType = QuestionnaireGridRow.RowTypes.Indicator And selGridRow.Indicator IsNot Nothing Then
                objIndicator = selGridRow.Indicator
                objIndicators = CurrentLogFrame.GetParentCollection(objIndicator)

                If objIndicators IsNot Nothing AndAlso objIndicators.Contains(objIndicator) Then
                    If objIndicator.Indicators.Count > 0 Then
                        If boolShift = True Then boolRemoveAll = True Else boolRemoveAll = False
                    Else
                        boolRemoveAll = True
                    End If

                    If boolRemoveAll = True Then
                        'CurrentUndoList.DeleteOperation(objIndicator, objIndicators, 0, strSortNumber, True)
                        objIndicators.Remove(objIndicator)
                    Else
                        Using copier As New ObjectCopy
                            'CurrentUndoList.DeleteOperation(copier.CopyObject(objIndicator), objIndicators, 0, strSortNumber, True)
                            '.ActionIndex = UndoListItemOld.Actions.DeleteNotVertical
                        End Using

                        With objIndicator
                            .RTF = String.Empty
                            .Indicators.Clear()
                            .VerificationSources.Clear()
                            .Statements.Clear()
                        End With

                        'CurrentUndoList.UndoBuffer.NewValue = objIndicator
                        'CurrentUndoList.AddToUndoList()
                    End If
                End If
            End If
        Next

        SelectedGridRows.Clear()

        ClearSelection()
        CurrentCell = Me(intColumnIndex, intRowIndex)
        Me.Reload()
    End Sub

    Private Function RemoveItems_Warning(ByVal strSourceColName As String) As Boolean
        Dim intNrIndicators As Integer
        Dim intNrInd As Integer, intNrResp As Integer
        Dim strMsg As String, strTitle As String = String.Empty
        Dim strMsgKeyMoment As String = String.Empty, strMsgIndicator As String = String.Empty

        For Each selGridRow As QuestionnaireGridRow In SelectedGridRows
            Select Case selGridRow.RowType
                Case QuestionnaireGridRow.RowTypes.Indicator
                    If selGridRow.Indicator IsNot Nothing Then
                        intNrInd += selGridRow.Indicator.Indicators.Count
                        intNrResp += selGridRow.Indicator.Statements.Count
                        intNrIndicators += 1
                    End If
            End Select
        Next

        If intNrIndicators > 0 Then
            strMsgIndicator = String.Format("{0} {1}", intNrIndicators, LANG_Indicators.ToLower)

            If intNrInd > 0 Or intNrResp > 0 Then
                strMsgIndicator &= " with:" & vbLf

                If intNrInd > 0 Then
                    strMsgIndicator &= String.Format("   - {0} {1}", intNrInd, LANG_Indicators.ToLower) & vbLf
                End If
                If intNrResp > 0 Then
                    strMsgIndicator &= String.Format("   - {0} {1}", intNrResp, LANG_Statements.ToLower) & vbLf
                End If
            End If

            strMsgIndicator &= vbLf
        End If

        strTitle = LANG_Remove
        strMsg = String.Format(LANG_RemoveQuestionnaireItems, strMsgKeyMoment, strMsgIndicator)

        Dim wdDeleteChildren As New DialogWarning(strMsg, strTitle)
        wdDeleteChildren.Type = DialogWarning.WarningDialogTypes.wdDeleteChildren
        wdDeleteChildren.ShowDialog()

        If wdDeleteChildren.DialogResult = Windows.Forms.DialogResult.No Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region 'remove items

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim selGridRow As QuestionnaireGridRow = Me.Grid(RowIndex)

        Dim intRowHeight As Integer = CONST_MinRowHeight

        'If selGridRow.RowType = QuestionnaireGridRow.RowTypes.Purpose And My.Settings.setRepeatPurposes = False Then
        '    selRow.Visible = False
        'ElseIf selGridRow.RowType = QuestionnaireGridRow.RowTypes.Output And My.Settings.setRepeatOutputs = False Then
        '    selRow.Visible = False
        'Else
        selRow.Visible = True
        intRowHeight = CalculateRowHeight(RowIndex)
        'End If

        If intRowHeight > CONST_MinRowHeight Then selRow.Height = intRowHeight Else selRow.Height = CONST_MinRowHeight
    End Sub

    Public Sub ResetRowHeights()
        For Each selRow As DataGridViewRow In Me.Rows

            SetRowHeight(selRow.Index)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selGridRow As QuestionnaireGridRow = Me.Grid(RowIndex)

        Select Case selGridRow.RowType
            Case QuestionnaireGridRow.RowTypes.Indicator
                If selGridRow.IndicatorHeight > intRowHeight Then intRowHeight = selGridRow.IndicatorHeight
            Case Else

        End Select


        Return intRowHeight
    End Function
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)
        Dim RowTmp As QuestionnaireGridRow = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name


        If e.RowIndex = RowCount - 1 Then
            Return
        End If

        ' Store a reference to the planning grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), QuestionnaireGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        Dim selIndicator As Indicator = RowTmp.Indicator
        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "SortNumber"
                e.Value = RowTmp.SortNumber
            Case "RTF"
                If RowTmp.IsIndicator Then
                    If RowTmp.Indicator IsNot Nothing Then e.Value = RowTmp.Indicator.RTF
                Else
                    If RowTmp.Struct IsNot Nothing Then e.Value = RowTmp.Struct.RTF
                End If
            Case "Range"
                If RowTmp.IsIndicator Then
                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.PercentageValue
                            If selIndicator.Statements.Count > 0 AndAlso selIndicator.Statements(0).ValuesDetail IsNot Nothing Then
                                Dim selValuesDetail As ValuesDetail = selIndicator.Statements(0).ValuesDetail
                                Dim selRange As ValueRange = selIndicator.Statements(0).ValuesDetail.ValueRange

                                If selRange IsNot Nothing Then e.Value = selValuesDetail.DisplayField(selRange)
                            End If
                    End Select
                End If
            Case Else
                If RowTmp.IsIndicator Then
                    Dim datColumn As Date
                    If Date.TryParse(strColName, datColumn) = False Then Exit Select

                    Select Case selIndicator.QuestionType
                        Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.PercentageValue
                            If selIndicator.Statements.Count > 0 AndAlso selIndicator.Statements(0).ValuesDetail IsNot Nothing Then
                                Dim selValuesDetail As ValuesDetail = selIndicator.Statements(0).ValuesDetail

                                'For Each selTarget As Target In selValuesDetail.Targets
                                '    If selTarget.Deadline.Date = datColumn.Date Then
                                '        e.Value = selValuesDetail.DisplayField(selTarget)
                                '    End If
                                'Next
                            End If
                    End Select
                End If
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As QuestionnaireGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim selRowIndex As Integer = e.RowIndex

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new planning grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As QuestionnaireGridRow = CType(Me.Grid(selRowIndex), QuestionnaireGridRow)

                Me.EditRow = New QuestionnaireGridRow
                With EditRow
                    .Indent = CurrentGridRow.Indent
                    .Indicator = CurrentGridRow.Indicator
                    .IndicatorHeight = CurrentGridRow.IndicatorHeight
                    .IndicatorImage = CurrentGridRow.IndicatorImage
                    .IndicatorImageDirty = CurrentGridRow.IndicatorImageDirty
                    .RowType = CurrentGridRow.RowType
                    .SortNumber = CurrentGridRow.SortNumber
                    .Struct = CurrentGridRow.Struct
                    .StructHeight = CurrentGridRow.StructHeight
                    .StructImage = CurrentGridRow.StructImage
                    .StructImageDirty = CurrentGridRow.StructImageDirty
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        ' Set the appropriate objRowEdit property to the cell value entered.
        strColName = Me.Columns(e.ColumnIndex).Name
        Select Case strColName
            Case "RTF"
                Dim ctlRTF As RichTextEditingControlLogframe = CType(Me.EditingControl, RichTextEditingControlLogframe)

                'necessary for pasting into non-active cell
                If ctlRTF Is Nothing Then
                    BeginEdit(False)
                    ctlRTF = CType(Me.EditingControl, RichTextEditingControlLogframe)
                End If

                If RowTmp.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
                    strCellValue = TryCast(ctlRTF.Rtf, String)
                End If
        End Select

        If RowTmp.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
            If RowTmp.Indicator Is Nothing Then RowTmp.Indicator = InitialiseIndicator(RowTmp, selRowIndex)
        End If

        Select Case strColName
            Case "RTF"
                If RowTmp.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
                    RowTmp.Struct.RTF = strCellValue
                    RowTmp.StructImageDirty = True
                End If
        End Select

        'CurrentUndoList.AddToUndoList()
    End Sub

    Private Function InitialiseIndicator(ByVal selGridRow As QuestionnaireGridRow, ByVal intRowIndex As Integer) As Indicator
        Dim NewIndicator As Indicator = Nothing

        Dim PreviousIndicator As Indicator = Grid.GetPreviousIndicator(intRowIndex)
        Dim ParentPurpose As Indicator = Nothing
        Dim ParentStruct As Struct = Nothing
        NewIndicator = New Indicator

        If PreviousIndicator Is Nothing Then
            ParentStruct = Grid.GetPreviousStruct(intRowIndex)
            If ParentStruct IsNot Nothing Then ParentStruct.Indicators.Add(NewIndicator)

            Return NewIndicator
        End If

        If PreviousIndicator.ParentStructGuid = Guid.Empty And PreviousIndicator.ParentIndicatorGuid = Guid.Empty Then
            'If CurrentLogFrame.Purposes.Count = 0 Then
            '    ParentPurpose = New Purpose
            '    CurrentLogFrame.Purposes.Add(ParentPurpose)
            'Else
            '    ParentPurpose = CurrentLogFrame.Purposes(0)
            'End If
            'If ParentPurpose.Outputs.Count = 0 Then
            '    ParentStruct = New Output
            '    ParentPurpose.Outputs.Add(ParentStruct)
            'Else
            '    ParentStruct = ParentPurpose.Outputs(0)
            'End If

            'ParentStruct.Indicators.Add(NewIndicator)
        Else
            If PreviousIndicator.ParentStructGuid <> Guid.Empty Then
                ParentStruct = CurrentLogFrame.GetStructByGuid(PreviousIndicator.ParentStructGuid)
                If ParentStruct IsNot Nothing Then ParentStruct.Indicators.Add(NewIndicator)
            ElseIf PreviousIndicator.ParentIndicatorGuid <> Guid.Empty Then
                Dim ParentIndicator As Indicator
                ParentIndicator = CurrentLogFrame.GetIndicatorByGuid(PreviousIndicator.ParentIndicatorGuid)
                If ParentIndicator IsNot Nothing Then ParentIndicator.Indicators.Add(NewIndicator)
            End If
        End If

        Return NewIndicator
    End Function

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New QuestionnaireGridRow
        Else
            ' If the user has canceled the edit of an existing row, 
            ' release the corresponding logframe row.
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        End If
        Me.Reload()
        MyBase.OnCancelRowEdit(e)

    End Sub

    Protected Overrides Sub OnNewRowNeeded(ByVal e As System.Windows.Forms.DataGridViewRowEventArgs)
        Me.EditRow = New QuestionnaireGridRow()
        Me.EditRowFlag = Me.Rows.Count - 1
    End Sub

    Protected Overrides Sub OnRowDirtyStateNeeded(ByVal e As System.Windows.Forms.QuestionEventArgs)
        If Not rowScopeCommit Then

            ' In cell-level commit scope, indicate whether the value
            ' of the current cell has been modified.
            e.Response = Me.IsCurrentCellDirty

        End If
    End Sub

    Protected Overrides Sub OnRowValidated(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ' Save row changes if any were made and release the edited 
        ' planning grid row if there is one.
        If e.RowIndex >= Me.Grid.Count AndAlso e.RowIndex <> Me.Rows.Count - 1 Then

            ' Add the new planning grid row to grid.
            Me.Grid.Add(Me.EditRow)

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then

            ' Save the modified planning grid row in grid.
            Me.Grid(e.RowIndex) = Me.EditRow
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        MyBase.OnRowValidated(e)
    End Sub
#End Region 'Virtual mode

#Region "Custom painting - general"
    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If Me.Grid.Count > 0 Then
            Dim CellGraphics As Graphics = e.Graphics
            Dim rCell As Rectangle = e.CellBounds
            Dim intIndex As Integer = e.RowIndex
            Dim selGridRow As QuestionnaireGridRow = Grid(intIndex)

            If intIndex >= 0 Then
                If selGridRow.RowType = QuestionnaireGridRow.RowTypes.Indicator Then
                    e.PaintBackground(rCell, False)
                    If e.ColumnIndex = 1 Then
                        CellPainting_RichText(selGridRow, rCell, CellGraphics)
                    Else
                        CellPainting_NormalText(e.FormattedValue, False, False, rCell, CellGraphics)
                    End If
                    e.Paint(rCell, DataGridViewPaintParts.Border)
                    e.Handled = True
                End If
            End If
        End If
    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As QuestionnaireGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow.StructImage IsNot Nothing Then
            rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StructImage.Width, selGridRow.StructImage.Height)
            CellGraphics.DrawImage(selGridRow.StructImage, rImage)
        ElseIf selGridRow.IndicatorImage IsNot Nothing Then
            rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.IndicatorImage.Width, selGridRow.IndicatorImage.Height)
            CellGraphics.DrawImage(selGridRow.IndicatorImage, rImage)
        End If
    End Sub

    Private Sub CellPainting_NormalText(ByVal strText As String, ByVal boolKeyMoment As Boolean, ByVal boolAlignRight As Boolean, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim fntText As Font = DefaultCellStyle.Font
        Dim brText As New SolidBrush(Color.Black)
        Dim objFormat As New StringFormat

        rCell.X += CONST_Padding
        rCell.Y += CONST_Padding
        rCell.Width -= CONST_HorizontalPadding
        rCell.Height -= CONST_VerticalPadding

        If boolKeyMoment = True Then
            brText = New SolidBrush(Color.DarkBlue)
            fntText = New Font(Me.DefaultCellStyle.Font, FontStyle.Bold)
        End If
        If boolAlignRight = True Then
            objFormat.Alignment = StringAlignment.Far
        Else
            objFormat.Alignment = StringAlignment.Near
        End If

        CellGraphics.DrawString(strText, fntText, brText, rCell, objFormat)
    End Sub

    Private Function LighterBrush(ByVal OldBrush As SolidBrush) As SolidBrush
        Dim intRed As Integer = OldBrush.Color.R + 50
        Dim intGreen As Integer = OldBrush.Color.G + 50
        Dim intBlue As Integer = OldBrush.Color.B + 50
        If intRed > 255 Then intRed = 255
        If intGreen > 255 Then intGreen = 255
        If intBlue > 255 Then intBlue = 255

        Dim NewBrush As New SolidBrush(Color.FromArgb(255, intRed, intGreen, intBlue))
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

    Public Function MakeGradientBrush(ByVal BaseColor As Color, ByVal intX As Integer, ByVal intY As Integer, ByVal intHeight As Integer) As LinearGradientBrush
        Dim colLight As Color = BaseColor
        Dim colDark As Color = DarkerColor(colLight)
        If intX < 0 Then intX = 0

        Dim NewBrush As New LinearGradientBrush(New Point(intX, intY), New Point(intX, intY + intHeight), colLight, colDark)
        Return NewBrush
    End Function
#End Region

#Region "Custom paint - Post paint"
    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim RowGraphics As Graphics = e.Graphics
        Dim selGridRow As QuestionnaireGridRow = Grid(e.RowIndex)

        If selGridRow IsNot Nothing Then
            Select Case selGridRow.RowType
                Case QuestionnaireGridRow.RowTypes.Indicator
                    DrawSelectionRectangle(RowGraphics)
                Case Else
                    'repeated purposes and outputs
                    RowPostPaint_PaintStructs(RowGraphics, e.RowBounds, selGridRow, e.RowIndex)
            End Select
        End If
    End Sub

    Private Sub RowPostPaint_PaintStructs(ByVal RowGraphics As Graphics, ByVal rRowBounds As Rectangle, ByVal selGridRow As QuestionnaireGridRow, ByVal intRowIndex As Integer)
        'draw repeated Purposes/Outputs
        Dim rStructSort As Rectangle, rStructRTF As Rectangle
        Dim boolRepeat As Boolean = True
        Dim brBackGround As Brush = Brushes.LightSteelBlue
        Dim intX = 1, intY As Integer
        Dim fntRepeat As Font = New Font(CurrentLogFrame.SortNumberFont, FontStyle.Bold)
        Dim brRepeat As SolidBrush = New SolidBrush(Color.Black)
        Dim objStringFormat As New StringFormat()

        objStringFormat.Alignment = StringAlignment.Near
        objStringFormat.LineAlignment = StringAlignment.Near

        If Me.RowHeadersVisible = True Then intX = Me.RowHeadersWidth
        intY = rRowBounds.Top

        Dim rRow As New Rectangle(intX, intY, Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - HorizontalScrollingOffset, rRowBounds.Height)

        'If selGridRow.RowType = QuestionnaireGridRow.RowTypes.RepeatPurpose Then brBackGround = Brushes.CornflowerBlue Else brBackGround = Brushes.LightSteelBlue
        RowGraphics.FillRectangle(brBackGround, rRow)

        'draw number
        rStructSort = GetCellDisplayRectangle(colSortNumber.Index, intRowIndex, False)
        rStructSort.X += 2

        RowGraphics.DrawString(selGridRow.SortNumber, fntRepeat, brRepeat, rStructSort, objStringFormat)

        'draw text
        rStructRTF = GetCellDisplayRectangle(colRTF.Index, intRowIndex, False)
        Dim strText As String = selGridRow.Struct.Text
        Dim rRepeat As New Rectangle(rStructRTF.X, rStructRTF.Y, _
                                     Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - rStructSort.Width - HorizontalScrollingOffset + 1, rRowBounds.Height)
        Dim intHeight As Integer = RowGraphics.MeasureString(strText, CurrentLogFrame.SectionTitleFont, rRepeat.Width, objStringFormat).Height
        rRepeat.Height = intHeight
        Me.Rows(intRowIndex).Height = intHeight
        RowGraphics.DrawString(strText, CurrentLogFrame.SectionTitleFont, brRepeat, rRepeat, objStringFormat)
        RowGraphics.DrawRectangle(penBorder, rRow)
    End Sub
#End Region 'Custom painting

#Region "Selection rectangle"
    Private Sub InvalidateSelectionRectangle()
        With SelectionRectangle
            Dim rSelection As New Rectangle(.Rectangle.X - 2, .Rectangle.Y - 2, .Rectangle.Width + 4, .Rectangle.Height + 4)
            Me.Invalidate(rSelection)
        End With
    End Sub

    Private Sub DrawSelectionRectangle(ByVal graph As Graphics)
        SetSelectionRectangle()

        With SelectionRectangle
            Dim IndexLastDisplayedRow As Integer = Me.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            If .FirstRowIndex > IndexLastDisplayedRow And .LastRowIndex > IndexLastDisplayedRow Then Exit Sub

            Dim IndexFirstDisplayedRow As Integer = Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
            If .FirstRowIndex < IndexFirstDisplayedRow And .LastRowIndex < IndexFirstDisplayedRow Then Exit Sub


            graph.DrawRectangle(penSelection, .Rectangle)
        End With
    End Sub

    Public Sub SetSelectionRectangle()
        Dim rTopLeftCell As New Rectangle
        Dim rBottomRightCell As New Rectangle
        Dim FirstRowIndexSelection As Integer

        If SelectedCells.Count > 0 Then
            SetSelectionRectangleGridArea()

            With SelectionRectangle
                If .FirstRowIndex < 0 Or .LastRowIndex < 0 Then Exit Sub
                FirstRowIndexSelection = .FirstRowIndex

                rTopLeftCell = GetCellDisplayRectangle(.FirstColumnIndex, .FirstRowIndex, False)


                If Me.Rows(.LastRowIndex).Displayed = False And .LastRowIndex >= Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed) Then
                    .LastRowIndex = Me.Rows.GetLastRow(DataGridViewElementStates.Displayed)
                End If
                rBottomRightCell = GetCellDisplayRectangle(.LastColumnIndex, .LastRowIndex, True)

                Dim intLeft As Integer = GetColumnDisplayRectangle(.FirstColumnIndex, True).X
                Dim intTop As Integer = rTopLeftCell.Y
                Dim intFirstColumnWidth As Integer = GetColumnDisplayRectangle(.FirstColumnIndex, True).Width
                Dim intFirstRowHeight As Integer = rTopLeftCell.Height
                Dim intRight As Integer = GetColumnDisplayRectangle(.LastColumnIndex, True).X
                Dim intBottom As Integer = rBottomRightCell.Y
                Dim intLastColumnWidth As Integer = GetColumnDisplayRectangle(.LastColumnIndex, True).Width
                Dim intLastRowHeight As Integer = rBottomRightCell.Height

                If .FirstRowIndex = .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                    .X = intLeft
                    .Y = intTop
                    .Width = intFirstColumnWidth
                    .Height = intFirstRowHeight
                ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                    .X = intLeft
                    .Y = intTop
                    .Width = intFirstColumnWidth
                    .Height = intBottom + intLastRowHeight - intTop
                ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                    .X = intRight
                    .Y = intBottom
                    .Width = intLastColumnWidth
                    .Height = intTop + intFirstRowHeight - intBottom
                ElseIf .FirstRowIndex = .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                    .X = intLeft
                    .Y = intTop
                    .Width = intRight + intLastColumnWidth - intLeft
                    .Height = intFirstRowHeight
                ElseIf .FirstRowIndex = .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                    .X = intRight
                    .Y = intBottom
                    .Width = intLeft + intFirstColumnWidth - intRight
                    .Height = intLastRowHeight

                ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                    .X = intLeft
                    .Y = intTop
                    .Width = intRight + intLastColumnWidth - intLeft
                    .Height = intBottom + intLastRowHeight - intTop
                ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                    .X = intLeft
                    .Y = intBottom
                    .Width = intRight + intLastColumnWidth - intLeft
                    .Height = intTop + intFirstRowHeight - intBottom
                ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                    .X = intRight
                    .Y = intTop
                    .Width = intLeft + intFirstColumnWidth - intRight
                    .Height = intBottom + intLastRowHeight - intTop
                ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                    .X = intRight
                    .Y = intBottom
                    .Width = intLeft + intFirstColumnWidth - intRight
                    .Height = intTop + intFirstRowHeight - intBottom
                End If

                .X += 1
                .Y += 1
                .Width -= 3
                .Height -= 3

                If .Rectangle.Bottom = -1 And .LastRowIndex >= Me.FirstDisplayedScrollingRowIndex Then .Height = Me.Bottom - .Y - 1
            End With
        End If
    End Sub

    Public Sub SetSelectionRectangleGridArea()
        Dim Vdir As Integer
        Dim intSelSize As Integer = SelectedCells.Count - 1
        If intSelSize < 0 Then intSelSize = 0

        If SelectedCells(0).RowIndex >= SelectedCells(intSelSize).RowIndex Then Vdir = 1 Else Vdir = -1

        SelectionRectangle.FirstRowIndex = CurrentCell.RowIndex
        SelectionRectangle.FirstColumnIndex = CurrentCell.ColumnIndex
        SelectionRectangle.LastRowIndex = CurrentCell.RowIndex
        SelectionRectangle.LastColumnIndex = CurrentCell.ColumnIndex

        If SelectedCells.Count > 1 Then
            For i = 0 To SelectedCells.Count - 1
                If SelectedCells(i).RowIndex < SelectionRectangle.FirstRowIndex Then SelectionRectangle.FirstRowIndex = SelectedCells(i).RowIndex
                If SelectedCells(i).RowIndex > SelectionRectangle.LastRowIndex Then SelectionRectangle.LastRowIndex = SelectedCells(i).RowIndex
                SelectionRectangle.FirstColumnIndex = 0
                SelectionRectangle.LastColumnIndex = Columns.Count - 1
            Next

            SetSelectionRectangleGridArea_Modify(Vdir)
        End If
    End Sub

    Public Sub SetSelectionRectangleGridArea_Modify(ByVal Vdir As Integer)
        Dim intRowType As Integer
        Dim boolRest As Boolean
        Dim i As Integer

        With SelectionRectangle
            'do not select title rows
            If Me.Grid(.FirstRowIndex).RowType <> QuestionnaireGridRow.RowTypes.Indicator Then
                .FirstRowIndex += 1
                If .LastRowIndex > .FirstRowIndex Then .LastRowIndex = .FirstRowIndex

                SetSelectionRectangleGridArea_Modify(Vdir)
            End If

            'if selection is larger than section, limit selection to that section
            If .LastRowIndex > .FirstRowIndex Then

                If Vdir > 0 Then
                    intRowType = Me.Grid(.FirstRowIndex).RowType
                    For i = .FirstRowIndex + 1 To .LastRowIndex
                        If Me.Grid(i).RowType <> intRowType And boolRest = False Then
                            boolRest = True
                            .LastRowIndex = i - 1
                        End If
                    Next
                Else
                    intRowType = Me.Grid(.LastRowIndex).RowType
                    For i = .LastRowIndex - 1 To .FirstRowIndex Step -1
                        If Me.Grid(i).RowType <> intRowType And boolRest = False Then
                            boolRest = True
                            .FirstRowIndex = i + 1
                        End If
                    Next
                End If
            End If
            If .LastRowIndex < 0 Then .LastRowIndex = 0
        End With
    End Sub
#End Region



End Class

Public Class IndicatorSelectedEventArgs
    Inherits EventArgs

    Public Property Indicator As Indicator

    Public Sub New(ByVal objIndicator As Indicator)
        MyBase.New()

        Me.Indicator = objIndicator
    End Sub
End Class


