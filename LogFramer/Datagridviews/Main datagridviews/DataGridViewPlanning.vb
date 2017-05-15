Public Class DataGridViewPlanning
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New PlanningGridRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe
    Friend WithEvents SelectionRectangle As New SelectionRectangle

    Public Event DragLinkReleased()
    Public Event PlanningObjectSelected(ByVal sender As Object, ByVal e As PlanningObjectSelectedEventArgs)
    Public Event Reloaded()
    Public Event UnlinkReleased()
    Public Event NoTextSelected()

    Private colSortNumber As New DataGridViewTextBoxColumn
    Private colRTF As New RichTextColumnLogframe
    Private colStartDate As New DataGridViewTextBoxColumn
    Private colEndDate As New DataGridViewTextBoxColumn
    Private colPlanning As New DataGridViewTextBoxColumn
    Private colHidden As New DataGridViewTextBoxColumn
    Private boolColumnWidthChanged As Boolean

    Private RowModifyIndex As Integer = -1
    Private EditRow As PlanningGridRow = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = True
    Private SelectedGridRows As New List(Of PlanningGridRow)
    Private PasteRow As PlanningGridRow
    Private sngScale As Single
    Private ClickPoint As Point
    Private pStartInsertLine, pEndInsertLine As New Point
    Private pStartInsertLineOld, pEndInsertLineOld As New Point

    Private DragActivity, RefActivity As Activity
    Private DragKeyMoment, RefKeyMoment As KeyMoment
    Private rDragRectangle As Rectangle
    Private rDragActivity As Rectangle
    Private rDragActivityStart As Rectangle
    Private rDragActivityEnd As Rectangle
    Private rDragPreparationStart As Rectangle
    Private rDragFollowUpEnd As Rectangle
    Private rDragKeyMoment As Rectangle
    Private boolDragActivity As Boolean
    Private boolDragActivityStart As Boolean
    Private boolDragActivityEnd As Boolean
    Private boolDragPreparationStart As Boolean
    Private boolDragFollowUpEnd As Boolean
    Private boolDragKeyMoment As Boolean
    Private datDragMoment, datInitialDragMoment As Date

    Private LinkActivity As Activity
    Private LinkKeyMoment As KeyMoment
    Private LinkRowIndex As Integer
    Private rLinkActivity As Rectangle
    Private rLinkKeyMoment As Rectangle
    Private boolDragLink, boolUnlink As Boolean
    Private boolLinkActivityStart As Boolean
    Private boolLinkActivityEnd As Boolean
    Private boolLinkKeyMoment As Boolean
    Private boolLinkHoverOverActivity, boolLinkHoverOverKeyMoment As Boolean
    Private ptLinkStart, ptLinkEnd As Point

    Private datStartDate As New Date
    Private datEndDate As New Date
    Private datPeriodFrom As New Date
    Private datPeriodEnd As New Date
    Private intPeriodView As Integer = PeriodViews.Month
    Private intElementsView As Integer
    Private boolShowDatesColumns As Boolean
    Private boolHideEmptyCells As Boolean
    Private boolShowActivityLinks As Boolean
    Private boolShowKeyMomentLinks As Boolean
    Private boolScroll As Boolean
    Private intHorizontalScrollingOffset As Integer

#Region "Enumerations"
    Public Enum TextSelectionValues
        SelectNothing = 0
        SelectAll = 1
        SelectActivities = 7
    End Enum
#End Region

#Region "Properties"
    Public Enum PeriodViews
        Day = 0
        Week = 1
        Month = 2
        Trimester = 3
        Semester = 4
        Year = 5
    End Enum

    Public Enum ShowElements
        ShowBoth = 0
        ShowActivities = 1
        ShowKeyMoments = 2
    End Enum

    Public Property PeriodView() As Integer
        Get
            Return intPeriodView
        End Get
        Set(ByVal value As Integer)
            intPeriodView = value
            If CurrentLogFrame IsNot Nothing Then
                Dim datExecutionStart As Date = CurrentLogFrame.GetExecutionStart(True)
                Dim datExecutionEnd As Date = CurrentLogFrame.GetExecutionEnd(True)

                If datExecutionStart > Date.MinValue And datExecutionEnd > Date.MinValue Then
                    Select Case intPeriodView
                        Case PeriodViews.Day
                            Me.PeriodStart = datExecutionStart.AddDays(-14)
                            Me.PeriodEnd = datExecutionEnd.AddDays(14)
                        Case PeriodViews.Week
                            Me.PeriodStart = datExecutionStart.AddMonths(-1)
                            Me.PeriodEnd = datExecutionEnd.AddMonths(1)
                        Case PeriodViews.Month
                            Me.PeriodStart = datExecutionStart.AddMonths(-3)
                            Me.PeriodEnd = datExecutionEnd.AddMonths(3)
                        Case PeriodViews.Trimester
                            Me.PeriodStart = datExecutionStart.AddMonths(-6)
                            Me.PeriodEnd = datExecutionEnd.AddMonths(6)
                        Case PeriodViews.Semester
                            Me.PeriodStart = datExecutionStart.AddMonths(-6)
                            Me.PeriodEnd = datExecutionEnd.AddMonths(6)
                        Case PeriodViews.Year
                            Me.PeriodStart = datExecutionStart.AddYears(-1)
                            Me.PeriodEnd = datExecutionEnd.AddYears(1)
                    End Select
                End If
            End If

        End Set
    End Property

    Public Property ElementsView() As Integer
        Get
            Return intElementsView
        End Get
        Set(ByVal value As Integer)
            intElementsView = value
        End Set
    End Property

    Public Property PeriodStart() As Date
        Get
            Return datPeriodFrom
        End Get
        Set(ByVal value As Date)
            datPeriodFrom = value
        End Set
    End Property

    Public Property PeriodEnd() As Date
        Get
            Return datPeriodEnd
        End Get
        Set(ByVal value As Date)
            datPeriodEnd = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return datStartDate
        End Get
        Set(ByVal value As Date)
            datStartDate = value
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Return datEndDate
        End Get
        Set(ByVal value As Date)
            datEndDate = value
        End Set
    End Property

    Public Property ShowDatesColumns As Boolean
        Get
            Return boolShowDatesColumns
        End Get
        Set(ByVal value As Boolean)
            boolShowDatesColumns = value
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

    Public Property ShowActivityLinks As Boolean
        Get
            Return boolShowActivityLinks
        End Get
        Set(ByVal value As Boolean)
            boolShowActivityLinks = value
        End Set
    End Property

    Public Property ShowKeyMomentLinks As Boolean
        Get
            Return boolShowKeyMomentLinks
        End Get
        Set(ByVal value As Boolean)
            boolShowKeyMomentLinks = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim selPlanningRow As PlanningGridRow = Me.Grid(Me.CurrentCell.RowIndex)
                Select Case selPlanningRow.RowType
                    Case PlanningGridRow.RowTypes.KeyMoment
                        If selPlanningRow.KeyMoment IsNot Nothing Then
                            If OnlyIfTextShows = True And String.IsNullOrEmpty(selPlanningRow.KeyMoment.Description) = False Then
                                selObject = selPlanningRow.KeyMoment
                            Else
                                selObject = selPlanningRow.KeyMoment
                            End If
                        End If
                    Case PlanningGridRow.RowTypes.Activity
                        Dim selActivity As Activity = TryCast(selPlanningRow.Struct, Activity)
                        If selActivity IsNot Nothing Then
                            If OnlyIfTextShows = True And String.IsNullOrEmpty(selActivity.RTF) = False Then
                                selObject = selActivity
                            Else
                                selObject = selActivity
                            End If
                        End If

                End Select
            End If
            Return selObject
        End Get
    End Property

    Public Property DragLink As Boolean
        Get
            Return boolDragLink
        End Get
        Set(ByVal value As Boolean)
            boolDragLink = value
        End Set
    End Property

    Public Property Unlink As Boolean
        Get
            Return boolUnlink
        End Get
        Set(ByVal value As Boolean)
            boolUnlink = value
        End Set
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
        InitialisePeriod()

        Me.PeriodView = My.Settings.setPlanningPeriodView

        LoadColumns()
    End Sub

    Public Sub InitialisePeriod()
        Me.StartDate = CurrentLogFrame.StartDate
        Me.EndDate = CurrentLogFrame.EndDate
        Me.PeriodStart = CurrentLogFrame.GetExecutionStart
        Me.PeriodEnd = CurrentLogFrame.GetExecutionEnd

        Me.PeriodView = My.Settings.setPlanningPeriodView
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
            .MinimumWidth = 250
        End With
        Me.Columns.Add(colRTF)

        With colStartDate
            .Name = "StartDate"
            .HeaderText = LANG_StartDate
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .Frozen = True
            .DefaultCellStyle.Format = "d"
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colStartDate)

        With colEndDate
            .Name = "EndDate"
            .HeaderText = LANG_EndDate
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .DefaultCellStyle.Format = "d"
            .Frozen = True
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopRight
            .HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With
        Me.Columns.Add(colEndDate)

        With colPlanning
            .Name = "Planning"
            .HeaderText = LANG_Planning
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .Frozen = False
            .ReadOnly = True
        End With
        Me.Columns.Add(colPlanning)

        With colHidden
            .Name = "Hidden"
            .HeaderText = ""
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .Frozen = False
            .Width = 0
        End With
        Me.Columns.Add(colHidden)

        ShowColumns()
    End Sub

    Public Sub ShowColumns()
        colStartDate.Visible = Me.ShowDatesColumns
        colEndDate.Visible = Me.ShowDatesColumns

        Invalidate()
    End Sub

    Public Sub Reload()
        intHorizontalScrollingOffset = Me.HorizontalScrollingOffset

        If CurrentLogFrame.GetExecutionStart < Me.PeriodStart Or CurrentLogFrame.GetExecutionEnd > Me.PeriodEnd Then
            'reset period start and end
            Me.PeriodView = My.Settings.setPlanningPeriodView
        End If

        Me.SuspendLayout()
        LoadEvents()

        Me.RowCount = Me.Grid.Count
        Me.RowModifyIndex = -1
        AutoSizeSortColumn()

        ResetAllImages()
        ReloadImages()
        SectionTitles_Protect()
        Me.HorizontalScrollingOffset = intHorizontalScrollingOffset
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
            Select Case Me.Grid(selRow.Index).RowType
                Case PlanningGridRow.RowTypes.RepeatPurpose, PlanningGridRow.RowTypes.RepeatOutput
                    'make title bar of each section and repeated objectives or outputs inaccessible
                    For Each selCell As DataGridViewCell In selRow.Cells
                        selCell.ReadOnly = True
                    Next
                Case Else
                    'make numbers inaccessible
                    selRow.Cells(0).ReadOnly = True
                    For i = 1 To CONST_PlanningColumnIndex - 1
                        selRow.Cells(i).ReadOnly = False
                    Next
                    selRow.Cells(CONST_PlanningColumnIndex).ReadOnly = True
            End Select
        Next
    End Sub
#End Region

#Region "Load activities/key moments"
    Private Sub LoadEvents()
        Me.Grid.Clear()

        If CurrentLogFrame IsNot Nothing Then
            If CurrentLogFrame.Purposes.Count = 0 Then
                If Me.HideEmptyCells = False Then
                    If Me.ElementsView <> ShowElements.ShowActivities Then _
                        Me.Grid.Add(New PlanningGridRow(PlanningGridRow.RowTypes.KeyMoment))
                    If Me.ElementsView <> ShowElements.ShowKeyMoments Then _
                        Me.Grid.Add(New PlanningGridRow(PlanningGridRow.RowTypes.Activity))
                End If
            Else
                For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                    If CurrentLogFrame.Purposes.Count > 1 Then
                        Dim strPurposeSort As String = CurrentLogFrame.GetStructSortNumber(selPurpose)
                        Dim rowPurpose As New PlanningGridRow
                        With rowPurpose
                            .SortNumber = strPurposeSort
                            .Struct = selPurpose
                            .RowType = PlanningGridRow.RowTypes.RepeatPurpose
                        End With
                        Me.Grid.Add(rowPurpose)
                    End If

                    If selPurpose.Outputs.Count = 0 Then
                        If Me.HideEmptyCells = False Then
                            If Me.ElementsView <> ShowElements.ShowActivities Then _
                                Me.Grid.Add(New PlanningGridRow(PlanningGridRow.RowTypes.KeyMoment))
                            If Me.ElementsView <> ShowElements.ShowKeyMoments Then _
                                Me.Grid.Add(New PlanningGridRow(PlanningGridRow.RowTypes.Activity))
                        End If
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            If selPurpose.Outputs.Count > 1 Then
                                Dim strOutputSort As String = CurrentLogFrame.GetStructSortNumber(selOutput)
                                Dim rowOutput As New PlanningGridRow
                                With rowOutput
                                    .SortNumber = strOutputSort
                                    .Struct = selOutput
                                    .RowType = PlanningGridRow.RowTypes.RepeatOutput
                                End With
                                Me.Grid.Add(rowOutput)
                            End If
                            LoadEvents_Output(selOutput)
                        Next
                    End If
                Next
            End If
        Else
            Exit Sub
        End If

        Me.RowCount = Me.Grid.Count
        LoadEvents_LinksIndices()
        SetScale()
    End Sub

    Private Sub LoadEvents_Output(ByVal selOutput As Output)

        If selOutput IsNot Nothing Then
            If Me.ElementsView <> ShowElements.ShowActivities Then
                Dim intNr As Integer = 1
                Dim strSortNr As String = "#" & CurrentLogFrame.GetStructSortNumber(selOutput) & CurrentLogFrame.SortNumberDivider

                'Load key moments of this output
                For Each selKeyMoment As KeyMoment In selOutput.KeyMoments.Sort
                    Dim NewGridItem As New PlanningGridRow
                    With NewGridItem

                        .SortNumber = strSortNr & intNr.ToString
                        .KeyMoment = selKeyMoment
                        .RowType = PlanningGridRow.RowTypes.KeyMoment
                    End With
                    Me.Grid.Add(NewGridItem)
                    intNr += 1
                Next
                If Me.HideEmptyCells = False Then
                    Me.Grid.Add(New PlanningGridRow(PlanningGridRow.RowTypes.KeyMoment))
                End If
            End If
            If Me.ElementsView <> ShowElements.ShowKeyMoments Then
                'Load activities

                LoadEvents_Activities(selOutput.Activities, 0)

            End If
        End If
    End Sub

    Private Sub LoadEvents_Activities(ByVal selActivities As Activities, ByVal intLevel As Integer)
        Dim boolLastHasSubActivities As Boolean

        For Each selActivity As Activity In selActivities
            Dim NewGridItem As New PlanningGridRow
            boolLastHasSubActivities = False

            With NewGridItem
                .SortNumber = CurrentLogFrame.GetStructSortNumber(selActivity)
                .Indent = intLevel
                .Struct = selActivity
                .RowType = PlanningGridRow.RowTypes.Activity
            End With

            Me.Grid.Add(NewGridItem)

            If selActivity.Activities.Count > 0 Then
                LoadEvents_Activities(selActivity.Activities, intLevel + 1)
                boolLastHasSubActivities = True
            End If

        Next
        If Me.HideEmptyCells = False And boolLastHasSubActivities = False Then
            Me.Grid.Add(New PlanningGridRow(PlanningGridRow.RowTypes.Activity))
        End If
    End Sub

    Private Sub LoadEvents_LinksIndices()
        For Each selGridRow As PlanningGridRow In Me.Grid
            If selGridRow.Struct IsNot Nothing Then
                Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
                If selActivity IsNot Nothing AndAlso selActivity.IsProcess = False Then
                    Dim selActivityDetail As ActivityDetail = selActivity.ActivityDetail
                    Dim intRowIndex As Integer = Grid.IndexOf(selGridRow)

                    If selActivityDetail.Relative = True Then
                        Dim intSourceRowIndex As Integer

                        'If selActivityDetail.PeriodDirection = ActivityDetail.DirectionValues.After Then
                        intSourceRowIndex = LoadEvents_GetRowIndexByReferenceGuid(selActivityDetail.GuidReferenceMoment)

                        If intSourceRowIndex >= 0 Then
                            Dim objSourceRow As PlanningGridRow = Me.Grid(intSourceRowIndex)

                            objSourceRow.OutgoingLinksIndices.Add(intRowIndex)
                            selGridRow.IncomingLinkIndices.Add(intSourceRowIndex)

                        End If
                        'Else
                        '    Dim intTargetRowIndex As Integer
                        '    intTargetRowIndex = LoadEvents_GetRowIndexByReferenceGuid(selActivityDetail.GuidReferenceMoment)

                        '    If intTargetRowIndex >= 0 Then
                        '        Dim objTargetRow As PlanningGridRow = Me.Grid(intTargetRowIndex)

                        '        selGridRow.OutgoingLinksIndices.Add(intTargetRowIndex)
                        '        objTargetRow.IncomingLinkIndices.Add(intRowIndex)

                        '    End If
                        'End If
                    End If
                End If
            ElseIf selGridRow.KeyMoment IsNot Nothing Then
                Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
                Dim intRowIndex As Integer = Grid.IndexOf(selGridRow)

                If selKeyMoment.Relative = True Then
                    Dim intSourceRowIndex As Integer

                    'If selKeyMoment.PeriodDirection = KeyMoment.DirectionValues.After Then
                    intSourceRowIndex = LoadEvents_GetRowIndexByReferenceGuid(selKeyMoment.GuidReferenceMoment)

                    If intSourceRowIndex >= 0 Then
                        Dim objSourceRow As PlanningGridRow = Me.Grid(intSourceRowIndex)

                        objSourceRow.OutgoingLinksIndices.Add(intRowIndex)
                        selGridRow.IncomingLinkIndices.Add(intSourceRowIndex)

                    End If
                    'Else
                    '    Dim intTargetRowIndex As Integer
                    '    intTargetRowIndex = LoadEvents_GetRowIndexByReferenceGuid(selKeyMoment.GuidReferenceMoment)

                    '    If intTargetRowIndex >= 0 Then
                    '        Dim objTargetRow As PlanningGridRow = Me.Grid(intTargetRowIndex)

                    '        selGridRow.OutgoingLinksIndices.Add(intTargetRowIndex)
                    '        objTargetRow.IncomingLinkIndices.Add(intRowIndex)

                    '    End If
                    'End If
                End If
            End If
        Next
    End Sub

    Private Function LoadEvents_GetRowIndexByReferenceGuid(ByVal selGuid As Guid) As Integer
        Dim intRowIndex As Integer = -1

        For Each selGridRow As PlanningGridRow In Me.Grid
            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.KeyMoment
                    Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
                    If selKeyMoment IsNot Nothing AndAlso selKeyMoment.Guid = selGuid Then
                        intRowIndex = Grid.IndexOf(selGridRow)
                        Exit For
                    End If
                Case PlanningGridRow.RowTypes.Activity
                    If selGridRow.Struct IsNot Nothing Then
                        Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
                        If selActivity IsNot Nothing AndAlso selActivity.Guid = selGuid Then
                            intRowIndex = Grid.IndexOf(selGridRow)
                            Exit For
                        End If
                    End If
            End Select
        Next

        Return intRowIndex
    End Function
#End Region 'load activities/key moments

#Region "Cell images"
    Private Sub ResetAllImages()
        For Each selRow As PlanningGridRow In Me.Grid
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As PlanningGridRow)
        selRow.StructImageDirty = True
        selRow.StructHeight = 0
    End Sub

    Private Sub ReloadImages()
        For Each selRow As PlanningGridRow In Me.Grid
            If selRow.RowType = PlanningGridRow.RowTypes.Activity Then
                ReloadImages_Activity(selRow)
            End If
        Next

        ResetRowHeights()
    End Sub

    Private Sub ReloadImages_Activity(ByVal selRow As PlanningGridRow)
        Dim intColumnWidth As Integer
        Dim selActivity As Activity = TryCast(selRow.Struct, Activity)
        Dim boolSelected As Boolean

        If selActivity Is Nothing Then Exit Sub

        With RichTextManager
            If selRow.StructImageDirty = True Then
                intColumnWidth = colRTF.Width

                If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectActivities Then
                    boolSelected = True
                Else
                    boolSelected = False
                End If

                If String.IsNullOrEmpty(selActivity.Text) Then
                    selRow.StructImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, Activity.ItemName, selRow.SortNumber, boolSelected)
                Else
                    selRow.StructImage = .RichTextWithPaddingToBitmap(intColumnWidth, selActivity.RTF, boolSelected, selRow.Indent)
                End If
                selRow.StructHeight = selRow.StructImage.Height
            End If
        End With
    End Sub

    Private Sub SetScale()
        If Me.RowCount > 0 Then

            Me.AutoResizeColumn(0)
            Me.AutoResizeColumn(1)
            Dim intMaxWidth As Integer = Me.Width / 4
            If colRTF.Width > intMaxWidth Then colRTF.Width = intMaxWidth
            Dim intNow, intMove As Integer
            Dim spanPeriod As TimeSpan = datPeriodEnd - datPeriodFrom
            If spanPeriod = TimeSpan.Zero Then Exit Sub

            Select Case PeriodView
                Case PeriodViews.Day
                    sngScale = 30
                Case PeriodViews.Week
                    sngScale = 10
                Case PeriodViews.Month
                    sngScale = 1
                Case PeriodViews.Year
                    sngScale = 1 / spanPeriod.Days * (Me.Width - colSortNumber.Width - colRTF.Width - colStartDate.Width - colEndDate.Width)
            End Select

            Dim intWidth As Integer = spanPeriod.Days * sngScale
            If intWidth > 65000 Then intWidth = 65000
            Me.Columns(CONST_PlanningColumnIndex).Width = intWidth
            intNow = Now.Subtract(datPeriodFrom).Days
            intNow *= sngScale

            Me.FirstDisplayedScrollingColumnIndex = CONST_PlanningColumnIndex 'ColumnCount - 1
            Dim intMaxMove As Integer = Me.HorizontalScrollingOffset
            intMove = CInt(intNow / Me.Columns(CONST_PlanningColumnIndex).Width * intMaxMove)
            If intMove > intMaxMove Then intMove = intMaxMove
            If Me.RowCount > 0 And intMove > 0 Then
                Try
                    Me.HorizontalScrollingOffset = intMove
                Catch ex As ArgumentOutOfRangeException

                End Try
            End If
            Columns(ColumnCount - 1).Visible = False
        End If
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim selGridRow As PlanningGridRow = Me.Grid(RowIndex)

        Dim intRowHeight As Integer = CONST_MinRowHeight

        If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose And My.Settings.setRepeatPurposes = False Then
            selRow.Visible = False
        ElseIf selGridRow.RowType = GridRow.RowTypes.RepeatOutput And My.Settings.setRepeatOutputs = False Then
            selRow.Visible = False
        Else
            selRow.Visible = True
            intRowHeight = CalculateRowHeight(RowIndex)
        End If

        If intRowHeight > CONST_MinRowHeight Then selRow.Height = intRowHeight Else selRow.Height = CONST_MinRowHeight
    End Sub

    Public Sub ResetRowHeights()
        For Each selRow As DataGridViewRow In Me.Rows

            SetRowHeight(selRow.Index)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selGridRow As PlanningGridRow = Me.Grid(RowIndex)

        If selGridRow.StructHeight > intRowHeight Then intRowHeight = selGridRow.StructHeight

        Return intRowHeight
    End Function
#End Region

#Region "Events"
    Private Sub RichTextEditingControl_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs) Handles RichTextEditingControl.ContentsResized
        Dim intRequiredHeight As Integer = e.NewRectangle.Height + RichTextEditingControl.Margin.Vertical + SystemInformation.VerticalResizeBorderThickness

        If CurrentRow.Height < intRequiredHeight Then
            CurrentRow.Height = intRequiredHeight
            SetSelectionRectangle()
            InvalidateSelectionRectangle()
        End If
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

            Dim selItem As Object = GetCurrentItem()

            If selItem IsNot Nothing Then
                Select Case selItem.GetType
                    Case GetType(Activity)
                        UndoRedo.UndoBuffer_Initialise(selItem, "RTF")
                    Case GetType(KeyMoment)
                        UndoRedo.UndoBuffer_Initialise(selItem, "Description")
                End Select
            End If
        End If
    End Sub

    Private Sub RichTextEditingControl_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextEditingControl.TextChanged
        If RichTextEditingControl.Text.Length > 0 Then
            Me.RowModifyIndex = Me.CurrentRow.Index
        Else
            Me.RowModifyIndex = -1
        End If

        UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.TextChanged)
    End Sub

    Private Sub RichTextEditingControl_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles RichTextEditingControl.Validating
        If RichTextEditingControl.Text.Length > 0 Then
            Dim selGridRow As PlanningGridRow = Me.Grid(CurrentRow.Index)

            If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                With RichTextEditingControl
                    .SelectAll()
                    .SelectionColor = Color.DarkBlue
                    .SelectionFont = New Font(Me.DefaultCellStyle.Font, FontStyle.Bold)
                End With
            End If
        End If
    End Sub

    Private Sub RichTextEditingControl_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles RichTextEditingControl.Validated
        UndoRedo.TextChanged(RichTextEditingControl.Rtf)
    End Sub

    Protected Overrides Sub OnEditingControlShowing(ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        MyBase.OnEditingControlShowing(e)

        Dim selItem As Object = GetCurrentItem()

        If Me.TextSelectionIndex <> TextSelectionValues.SelectNothing Then
            Me.TextSelectionIndex = TextSelectionValues.SelectNothing
            ResetAllImages()
            ReloadImages()
            Invalidate()
            RaiseEvent NoTextSelected()
        End If

        Me.RichTextEditingControl = TryCast(e.Control, RichTextEditingControlLogframe)

        If selItem IsNot Nothing Then
            With UndoRedo
                Select Case selItem.GetType
                    Case GetType(Activity)
                        UndoRedo.UndoBuffer_Initialise(selItem, "RTF")
                    Case GetType(KeyMoment)
                        UndoRedo.UndoBuffer_Initialise(selItem, "Description")
                End Select
            End With
        End If
    End Sub

    Private Sub DataGridViewPlanning_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Me.Scroll
        Invalidate()
    End Sub
#End Region

    
End Class

Public Class PlanningObjectSelectedEventArgs
    Inherits EventArgs

    Public Property PlanningObject As Object

    Public Sub New(ByVal objPlanningObject As Object)
        MyBase.New()

        Me.PlanningObject = objPlanningObject
    End Sub
End Class


