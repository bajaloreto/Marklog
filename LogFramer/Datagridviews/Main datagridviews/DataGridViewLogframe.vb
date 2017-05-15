Public Class DataGridViewLogframe
    Inherits DataGridViewBaseClassRichText

    Friend WithEvents Grid As New LogframeRows
    Friend WithEvents RichTextManager As New RichTextManager
    Friend WithEvents SelectionRectangle As New SelectionRectangle
    Friend WithEvents RichTextEditingControl As New RichTextEditingControlLogframe

    Public Event Reloaded()
    Public Event LogframeObjectSelected(ByVal sender As Object, ByVal e As LogframeObjectSelectedEventArgs)
    Public Event NoTextSelected()
    Public Event ShowIndicatorColumnChanged()
    Public Event ShowVerificationSourceColumnChanged()
    Public Event ShowAssumptionColumnChanged()
    Public Event ShowResourcesBudgetChanged()

    Private objLogframe As LogFrame

    Private colStructSortColumn As New DataGridViewTextBoxColumn
    Private colStructRtfColumn As New RichTextColumnLogframe()
    Private colIndicatorSortColumn As New DataGridViewTextBoxColumn
    Private colIndicatorRtfColumn As New RichTextColumnLogframe()
    Private colVerificationSourceSortColumn As New DataGridViewTextBoxColumn
    Private colVerificationSourceRtfColumn As New RichTextColumnLogframe()
    Private colAssumptionSortColumn As New DataGridViewTextBoxColumn
    Private colAssumptionRtfColumn As New RichTextColumnLogframe()

    Private intRowModifyIndex As Integer = -1
    Private EditRowFlag As Integer = -1
    Private EditRow As LogframeRow
    Private rowScopeCommit As Boolean = True
    Private SelectedGridRows As New List(Of LogframeRow)
    Private PasteRow As LogframeRow
    Private boolReloading As Boolean
    Private ClickPoint As Point = Nothing
    Private pStartInsertLine As New Point, pEndInsertLine As New Point
    Private pStartInsertLineOld As New Point, pEndInsertLineOld As New Point
    Private OldSelectionRectangle As New SelectionRectangle
    Private boolColumnWidthChanged As Boolean
    Private PreviousCellLocation As New Point

    Private boolHideEmptyCells As Boolean
    Private boolTextSelected As Boolean
    Private boolShowIndColumn As Boolean = True
    Private boolShowVerColumn As Boolean = True
    Private boolShowAsmColumn As Boolean = True
    Private boolShowGoals As Boolean = True
    Private boolShowPurposes As Boolean = True
    Private boolShowOutputs As Boolean = True
    Private boolShowActivities As Boolean = True
    Private boolShowResourcesBudget As Boolean = True

#Region "Enumerations"
    Public Enum TextSelectionValues
        SelectNothing = 0
        SelectAll = 1
        SelectLogframe = 2
        SelectStructs = 3
        SelectIndicators = 4
        SelectVerificationSources = 5
        SelectResources = 6
        SelectResourceBudgets = 7
        SelectAssumptions = 9
    End Enum
#End Region

#Region "Properties"
    Public Property Logframe As LogFrame
        Get
            Return objLogframe
        End Get
        Set(ByVal value As LogFrame)
            objLogframe = value
        End Set
    End Property

    Public Property StructSortColumn As DataGridViewTextBoxColumn
        Get
            Return colStructSortColumn
        End Get
        Set(ByVal value As DataGridViewTextBoxColumn)
            colStructSortColumn = value
        End Set
    End Property

    Public Property StructRtfColumn As RichTextColumnLogframe
        Get
            Return colStructRtfColumn
        End Get
        Set(ByVal value As RichTextColumnLogframe)
            colStructRtfColumn = value
        End Set
    End Property

    Public Property IndicatorSortColumn As DataGridViewTextBoxColumn
        Get
            Return colIndicatorSortColumn
        End Get
        Set(ByVal value As DataGridViewTextBoxColumn)
            colIndicatorSortColumn = value
        End Set
    End Property

    Public Property IndicatorRtfColumn As RichTextColumnLogframe
        Get
            Return colIndicatorRtfColumn
        End Get
        Set(ByVal value As RichTextColumnLogframe)
            colIndicatorRtfColumn = value
        End Set
    End Property

    Public Property VerificationSourceSortColumn As DataGridViewTextBoxColumn
        Get
            Return colVerificationSourceSortColumn
        End Get
        Set(ByVal value As DataGridViewTextBoxColumn)
            colVerificationSourceSortColumn = value
        End Set
    End Property

    Public Property VerificationSourceRtfColumn As RichTextColumnLogframe
        Get
            Return colVerificationSourceRtfColumn
        End Get
        Set(ByVal value As RichTextColumnLogframe)
            colVerificationSourceRtfColumn = value
        End Set
    End Property

    Public Property AssumptionSortColumn As DataGridViewTextBoxColumn
        Get
            Return colAssumptionSortColumn
        End Get
        Set(ByVal value As DataGridViewTextBoxColumn)
            colAssumptionSortColumn = value
        End Set
    End Property

    Public Property AssumptionRtfColumn As RichTextColumnLogframe
        Get
            Return colAssumptionRtfColumn
        End Get
        Set(ByVal value As RichTextColumnLogframe)
            colAssumptionRtfColumn = value
        End Set
    End Property

    Public Property RowModifyIndex As Integer
        Get
            Return intRowModifyIndex
        End Get
        Set(ByVal value As Integer)
            intRowModifyIndex = value
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

    Public Property ShowIndicatorColumn() As Boolean
        Get
            Return boolShowIndColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowIndColumn = value
        End Set
    End Property

    Public Property ShowVerificationSourceColumn() As Boolean
        Get
            Return boolShowVerColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowVerColumn = value
        End Set
    End Property

    Public Property ShowAssumptionColumn() As Boolean
        Get
            Return boolShowAsmColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowAsmColumn = value
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

    Public Property ShowResourcesBudget() As Boolean
        Get
            Return boolShowResourcesBudget
        End Get
        Set(ByVal value As Boolean)
            boolShowResourcesBudget = value
        End Set
    End Property

    Public ReadOnly Property StructRtfColumnWidth As Integer
        Get
            Return Columns(1).Width
        End Get
    End Property

    Public ReadOnly Property IndRtfColumnWidth As Integer
        Get
            If ShowIndicatorColumn = True Then
                Return Columns(3).Width
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property VerRtfColumnWidth As Integer
        Get
            If ShowVerificationSourceColumn = True Then
                Return Columns(5).Width
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property AsmRtfColumnWidth As Integer
        Get
            If ShowAssumptionColumn = True Then
                Return Columns(7).Width
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property CurrentRtfColumn As String
        Get
            Dim strColumnName As String = String.Empty
            If Me.CurrentCell IsNot Nothing Then
                strColumnName = Me.Columns(Me.CurrentCell.ColumnIndex).Name
                If strColumnName.Contains("Sort") Then strColumnName.Replace("Sort", "RTF")
            End If

            Return strColumnName
        End Get
    End Property

    Public ReadOnly Property CurrentSection As Integer
        Get
            Dim intSection As Integer = -1
            If Me.CurrentCell IsNot Nothing Then
                Dim selLogframeRow As LogframeRow = Me.Grid(Me.CurrentCell.RowIndex)
                intSection = selLogframeRow.Section
            End If
            Return intSection
        End Get
    End Property

    Public Overrides ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object
        Get
            Dim selObject As Object = Nothing
            If Me.CurrentCell IsNot Nothing Then
                Dim selLogframeRow As LogframeRow = Me.Grid(Me.CurrentCell.RowIndex)
                Select Case CurrentRtfColumn
                    Case "StructRTF"
                        If OnlyIfTextShows = True And String.IsNullOrEmpty(selLogframeRow.StructRtf) = False Then
                            selObject = selLogframeRow.Struct
                        Else
                            selObject = selLogframeRow.Struct
                        End If
                    Case "IndRTF"
                        If IsResourceBudgetRow(selLogframeRow) = False Then
                            If OnlyIfTextShows = True And String.IsNullOrEmpty(selLogframeRow.IndicatorRtf) = False Then
                                selObject = selLogframeRow.Indicator
                            Else
                                selObject = selLogframeRow.Indicator
                            End If
                        Else
                            If OnlyIfTextShows = True And String.IsNullOrEmpty(selLogframeRow.ResourceRtf) = False Then
                                selObject = selLogframeRow.Resource
                            Else
                                selObject = selLogframeRow.Resource
                            End If
                        End If
                    Case "VerRTF"
                        If IsResourceBudgetRow(selLogframeRow) = False Then
                            If OnlyIfTextShows = True And String.IsNullOrEmpty(selLogframeRow.VerificationSourceRtf) = False Then
                                selObject = selLogframeRow.VerificationSource
                            Else
                                selObject = selLogframeRow.VerificationSource
                            End If
                        Else
                            If OnlyIfTextShows = True And String.IsNullOrEmpty(selLogframeRow.ResourceRtf) = False Then
                                selObject = selLogframeRow.Resource 'TotalCostAmount
                            Else
                                selObject = selLogframeRow.Resource 'TotalCostAmount
                            End If
                        End If
                    Case "AsmRTF"
                        If OnlyIfTextShows = True And String.IsNullOrEmpty(selLogframeRow.AssumptionRtf) = False Then
                            selObject = selLogframeRow.Assumption
                        Else
                            selObject = selLogframeRow.Assumption
                        End If
                End Select
            End If
            Return selObject
        End Get
    End Property
#End Region 'Properties

#Region "Initialise"
    Public Sub New()
        Initialise()
    End Sub

    Public Sub New(ByVal selLogframe As LogFrame)
        Me.objLogframe = selLogframe
        Initialise()
    End Sub

    Private Sub Initialise()
        Me.DoubleBuffered = True

        VirtualMode = True
        AutoGenerateColumns = False
        AllowUserToAddRows = False

        Me.ColumnHeadersVisible = True
        Me.RowHeadersVisible = False
        Me.ScrollBars = Windows.Forms.ScrollBars.Vertical
        DefaultCellStyle.Padding = New Padding(2)
        ShowCellToolTips = False
        MultiSelect = True

        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = False
    End Sub

    Public Sub InitialiseColumns()
        Columns.Clear()

        'Add project logic (always shown)
        Columns.Add(StructSortColumn)
        Columns.Add(StructRtfColumn)

        With StructSortColumn
            .Name = "StructSort"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            With .DefaultCellStyle
                .Alignment = DataGridViewContentAlignment.TopLeft
                .Font = Me.Logframe.SortNumberFont
            End With
            .ReadOnly = True
        End With
        With StructRtfColumn
            .Name = "StructRTF"
            .Visible = True
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .FillWeight = 1
        End With

        'Add Indicators
        Columns.Add(IndicatorSortColumn)
        Columns.Add(IndicatorRtfColumn)

        With IndicatorSortColumn
            .Name = "IndSort"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            With .DefaultCellStyle
                .Alignment = DataGridViewContentAlignment.TopLeft
                .Font = Me.Logframe.SortNumberFont
            End With
            .ReadOnly = True
        End With
        With IndicatorRtfColumn
            .Name = "IndRTF"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .FillWeight = 1
        End With

        'Add Means of verification
        Columns.Add(VerificationSourceSortColumn)
        Columns.Add(VerificationSourceRtfColumn)

        With VerificationSourceSortColumn
            .Name = "VerSort"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            With .DefaultCellStyle
                .Alignment = DataGridViewContentAlignment.TopLeft
                .Font = Me.Logframe.SortNumberFont
            End With
            .ReadOnly = True
        End With
        With VerificationSourceRtfColumn
            .Name = "VerRTF"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .FillWeight = 1
        End With

        'Add Assumptions
        Columns.Add(AssumptionSortColumn)
        Columns.Add(AssumptionRtfColumn)

        With AssumptionSortColumn
            .Name = "AsmSort"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            With .DefaultCellStyle
                .Alignment = DataGridViewContentAlignment.TopLeft
                .Font = Me.Logframe.SortNumberFont
            End With
            .ReadOnly = True
        End With
        With AssumptionRtfColumn
            .Name = "AsmRTF"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .FillWeight = 1
        End With

        ShowColumns()
    End Sub

    Public Sub ShowColumns()
        IndicatorSortColumn.Visible = Me.ShowIndicatorColumn
        IndicatorRtfColumn.Visible = Me.ShowIndicatorColumn
        VerificationSourceSortColumn.Visible = Me.ShowVerificationSourceColumn
        VerificationSourceRtfColumn.Visible = Me.ShowVerificationSourceColumn
        AssumptionSortColumn.Visible = Me.ShowAssumptionColumn
        AssumptionRtfColumn.Visible = Me.ShowAssumptionColumn

        Invalidate()
    End Sub

    Public Sub Reload()
        If boolReloading = True Then Exit Sub

        boolReloading = True

        Me.SuspendLayout()
        LoadSections()

        Me.RowCount = Me.Grid.Count
        Me.RowModifyIndex = -1
        AutoSizeSortColumns()

        ResetAllImages()
        ReloadImages()
        SectionTitles_Protect()
        Me.Invalidate()
        Me.ResumeLayout()
        boolReloading = False

        RaiseEvent Reloaded()
    End Sub

    Private Sub AutoSizeSortColumns()
        Dim selcolumn As DataGridViewColumn
        Dim intWidth As Integer = 0

        For i = 0 To Me.ColumnCount - 1 Step 2
            selcolumn = Me.Columns(i)

            selcolumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            intWidth = selcolumn.Width
            selcolumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            selcolumn.Width = intWidth
        Next
    End Sub
#End Region

#Region "Load sections"
    Public Sub LoadSections()
        Me.Grid.Clear()

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
    End Sub

    Private Sub LoadSections_SectionGoals()
        Dim rowSectionTitle As LogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.GoalsSection)
        rowSectionTitle.RowType = LogframeRow.RowTypes.Section

        Me.Grid.Add(rowSectionTitle)

        Dim intGoalRowIndex As Integer
        Dim strGoalSort As String

        If Me.Logframe.Goals.Count > 0 Then
            For Each selGoal As Goal In Me.Logframe.Goals
                intGoalRowIndex = Me.Grid.Count
                strGoalSort = Me.Logframe.CreateSortNumber(Logframe.Goals.IndexOf(selGoal))

                Dim rowGoal As New LogframeRow
                With rowGoal
                    .Section = Logframe.SectionTypes.GoalsSection
                    .StructSort = strGoalSort
                    .Struct = selGoal
                End With
                Me.Grid.Add(rowGoal)
                If selGoal.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.GoalsSection, selGoal, intGoalRowIndex, strGoalSort)
                If selGoal.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.GoalsSection, selGoal, intGoalRowIndex, strGoalSort)
            Next
        End If
        If Me.HideEmptyCells = False Then
            Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.GoalsSection))
        End If
    End Sub

    Private Sub LoadSections_SectionPurposes()
        Dim rowSectionTitle As LogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.PurposesSection)
        rowSectionTitle.RowType = LogframeRow.RowTypes.Section
        Me.Grid.Add(rowSectionTitle)

        Dim intPurposeRowIndex As Integer
        Dim strPurposeSort As String

        If Me.Logframe.Purposes.Count > 0 Then
            For Each selPurpose As Purpose In Me.Logframe.Purposes
                intPurposeRowIndex = Me.Grid.Count
                strPurposeSort = Logframe.CreateSortNumber(Logframe.Purposes.IndexOf(selPurpose))

                Dim rowPurpose As New LogframeRow
                With rowPurpose
                    .Section = Logframe.SectionTypes.PurposesSection
                    .StructSort = strPurposeSort
                    .Struct = selPurpose
                End With
                Me.Grid.Add(rowPurpose)
                If selPurpose.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.PurposesSection, selPurpose, intPurposeRowIndex, strPurposeSort)
                If selPurpose.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.PurposesSection, selPurpose, intPurposeRowIndex, strPurposeSort)
            Next
        End If
        If Me.HideEmptyCells = False Then
            Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.PurposesSection))
        End If
    End Sub

    Private Sub LoadSections_SectionOutputs()
        Dim rowSectionTitle As LogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.OutputsSection)
        rowSectionTitle.RowType = LogframeRow.RowTypes.Section
        Me.Grid.Add(rowSectionTitle)

        Dim intOutputRowIndex As Integer
        Dim strPurposeSort As String, strOutputSort As String

        For Each selPurpose As Purpose In Me.Logframe.Purposes
            strPurposeSort = Logframe.CreateSortNumber(Logframe.Purposes.IndexOf(selPurpose))

            If Me.Logframe.Purposes.Count > 1 Then LoadSections_RepeatPurposes(selPurpose, strPurposeSort)

            For Each selOutput As Output In selPurpose.Outputs
                intOutputRowIndex = Me.Grid.Count
                strOutputSort = Logframe.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)

                Dim rowOutput As New LogframeRow
                With rowOutput
                    .Section = Logframe.SectionTypes.OutputsSection
                    .StructSort = strOutputSort
                    .Struct = selOutput
                End With
                Me.Grid.Add(rowOutput)
                If selOutput.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.OutputsSection, selOutput, intOutputRowIndex, strOutputSort)
                If selOutput.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.OutputsSection, selOutput, intOutputRowIndex, strOutputSort)
            Next
            If Me.HideEmptyCells = False Then
                Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.OutputsSection))
            End If
        Next
        If Me.Logframe.Purposes.Count = 0 And Me.HideEmptyCells = False Then
            Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.OutputsSection))
        End If
    End Sub

    Private Sub LoadSections_SectionActivities()
        Dim rowSectionTitle As LogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.ActivitiesSection)
        rowSectionTitle.RowType = LogframeRow.RowTypes.Section
        Me.Grid.Add(rowSectionTitle)

        Dim strPurposeSort As String, strOutputSort As String

        For Each selPurpose As Purpose In Me.Logframe.Purposes
            strPurposeSort = Logframe.CreateSortNumber(Logframe.Purposes.IndexOf(selPurpose))
            If Me.Logframe.Purposes.Count > 1 Then LoadSections_RepeatPurposes(selPurpose, strPurposeSort)

            For Each selOutput As Output In selPurpose.Outputs
                strOutputSort = Logframe.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)
                If selPurpose.Outputs.Count > 1 Then LoadSections_RepeatOutputs(selOutput, strOutputSort)

                LoadSections_SectionActivities_Activities(selOutput.Activities, strOutputSort, 0)
            Next
            If selPurpose.Outputs.Count = 0 And Me.HideEmptyCells = False Then
                Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.ActivitiesSection))
            End If
        Next
        If Me.Logframe.Purposes.Count = 0 And Me.HideEmptyCells = False Then
            Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.ActivitiesSection))
        End If
    End Sub

    Private Sub LoadSections_SectionActivities_Activities(ByVal selActivities As Activities, ByVal strParentSort As String, ByVal intLevel As Integer)
        Dim intActivityRowIndex As Integer
        Dim strActivitySort As String
        Dim boolLastHasSubActivities As Boolean

        For Each selActivity As Activity In selActivities
            boolLastHasSubActivities = False
            intActivityRowIndex = Me.Grid.Count
            strActivitySort = Logframe.CreateSortNumber(selActivities.IndexOf(selActivity), strParentSort)

            Dim rowActivity As New LogframeRow
            With rowActivity
                .Section = Logframe.SectionTypes.ActivitiesSection
                .StructSort = strActivitySort
                .Struct = selActivity
                .StructIndent = intLevel
            End With
            Me.Grid.Add(rowActivity)

            If IsResourceBudgetRow(rowActivity) = False Then
                If selActivity.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.ActivitiesSection, selActivity, intActivityRowIndex, strActivitySort)
            Else
                If selActivity.Resources.Count > 0 Then LoadSections_Resources(Logframe.SectionTypes.ActivitiesSection, selActivity, intActivityRowIndex, strActivitySort)
            End If
            If selActivity.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.ActivitiesSection, selActivity, intActivityRowIndex, strActivitySort)
            If selActivity.Activities.Count > 0 Then
                LoadSections_SectionActivities_Activities(selActivity.Activities, strActivitySort, intLevel + 1)
                boolLastHasSubActivities = True
            End If
        Next

        If Me.HideEmptyCells = False And boolLastHasSubActivities = False Then
            Me.Grid.Add(New LogframeRow(Logframe.SectionTypes.ActivitiesSection))
        End If
    End Sub

    Private Sub LoadSections_Indicators(ByVal intSection As Integer, ByVal objParentStruct As Struct, ByVal intStructRowIndex As Integer, ByVal strStructSort As String)
        If ShowIndicatorColumn = True Then
            LoadSections_SectionIndicators_Indicators(intSection, objParentStruct.Indicators, intStructRowIndex, strStructSort, 0, True)
        End If
    End Sub

    Private Sub LoadSections_SectionIndicators_Indicators(ByVal intSection As Integer, ByVal selIndicators As Indicators, ByVal intIndicatorRowIndex As Integer, ByVal strParentSort As String, _
                                                          ByVal intLevel As Integer, ByVal boolFirstOfStruct As Boolean)
        Dim strIndicatorSort As String
        Dim boolLastHasSubIndicators As Boolean
        Dim rowIndicator As LogframeRow
        Dim selIndicator As Indicator = Nothing

        For Each selIndicator In selIndicators
            boolLastHasSubIndicators = False


            If Me.Logframe.SortNumberRepeatParent = True Then strIndicatorSort = strParentSort Else strIndicatorSort = String.Empty
            strIndicatorSort = Logframe.CreateSortNumber(selIndicators.IndexOf(selIndicator), strParentSort)

            If boolFirstOfStruct = True Then
                rowIndicator = Grid(intIndicatorRowIndex)
                boolFirstOfStruct = False
            Else
                rowIndicator = New LogframeRow(intSection)
                Me.Grid.Add(rowIndicator)
                'intIndicatorRowIndex = Grid.Count
            End If

            With rowIndicator
                .IndicatorSort = strIndicatorSort
                .Indicator = selIndicator
                .IndicatorIndent = intLevel
            End With

            If selIndicator.VerificationSources.Count > 0 Then

                LoadSections_VerificationSources(intSection, selIndicator, intIndicatorRowIndex, strIndicatorSort)

            End If
            intIndicatorRowIndex = Me.Grid.Count
            If selIndicator.Indicators.Count > 0 Then
                LoadSections_SectionIndicators_Indicators(intSection, selIndicator.Indicators, intIndicatorRowIndex, strIndicatorSort, intLevel + 1, False)
                boolLastHasSubIndicators = True
            End If
            intIndicatorRowIndex = Me.Grid.Count
        Next

        If Me.HideEmptyCells = False And boolLastHasSubIndicators = False Then
            Me.Grid.Add(New LogframeRow(intSection))
        End If
    End Sub

    Private Sub LoadSections_VerificationSources(ByVal intSection As Integer, ByVal objParentIndicator As Indicator, ByVal intVerificationSourceRowIndex As Integer, ByVal strIndicatorSort As String)
        Dim strVerificationSourceSort As String
        Dim rowVerificationSource As LogframeRow

        If ShowVerificationSourceColumn = True Then
            For Each selVerificationSource As VerificationSource In objParentIndicator.VerificationSources
                If Me.Logframe.SortNumberRepeatParent = True Then strVerificationSourceSort = strIndicatorSort Else strVerificationSourceSort = String.Empty
                strVerificationSourceSort = Logframe.CreateSortNumber(objParentIndicator.VerificationSources.IndexOf(selVerificationSource), strIndicatorSort)

                If intVerificationSourceRowIndex > Grid.Count - 1 Then
                    rowVerificationSource = New LogframeRow(intSection)
                    Me.Grid.Add(rowVerificationSource)
                Else
                    rowVerificationSource = Grid(intVerificationSourceRowIndex)
                End If

                With rowVerificationSource
                    .VerificationSourceSort = strVerificationSourceSort
                    .VerificationSource = selVerificationSource
                End With
                intVerificationSourceRowIndex = Me.Grid.Count
            Next
            If Me.HideEmptyCells = False Then
                Me.Grid.Add(New LogframeRow(intSection))
            End If
        End If
    End Sub

    Private Sub LoadSections_Resources(ByVal intSection As Integer, ByVal objParentStruct As Struct, ByVal intStructRowIndex As Integer, ByVal strStructSort As String)
        If ShowIndicatorColumn = True Then
            Dim intResourceRowIndex As Integer = intStructRowIndex
            Dim strResourceSort As String
            Dim rowResource As LogframeRow
            Dim selActivity As Activity = TryCast(objParentStruct, Activity)

            If selActivity Is Nothing Then Exit Sub

            For Each selResource As Resource In selActivity.Resources
                If Me.Logframe.SortNumberRepeatParent = True Then strResourceSort = strStructSort Else strResourceSort = String.Empty
                strResourceSort = Logframe.CreateSortNumber(selActivity.Resources.IndexOf(selResource), strStructSort)

                If intResourceRowIndex > Grid.Count - 1 Then
                    rowResource = New LogframeRow(intSection)
                    Me.Grid.Add(rowResource)
                Else
                    rowResource = Grid(intResourceRowIndex)
                End If

                With rowResource
                    .ResourceSort = strResourceSort
                    .Resource = selResource
                    .TotalCostAmount = selResource.TotalCostAmount
                End With

                intResourceRowIndex = Me.Grid.Count
            Next
            If Me.HideEmptyCells = False Then
                Me.Grid.Add(New LogframeRow(intSection))
            End If
        End If
    End Sub

    Private Sub LoadSections_Assumptions(ByVal intSection As Integer, ByVal objParentStruct As Struct, ByVal intStructRowIndex As Integer, ByVal strStructSort As String)
        Dim intAssumptionRowIndex As Integer = intStructRowIndex
        Dim strAssumptionSort As String
        Dim rowAssumption As LogframeRow

        If ShowAssumptionColumn = True Then
            For Each selAssumption As Assumption In objParentStruct.Assumptions
                If Me.Logframe.SortNumberRepeatParent = True Then strAssumptionSort = strStructSort Else strAssumptionSort = String.Empty
                strAssumptionSort = Logframe.CreateSortNumber(objParentStruct.Assumptions.IndexOf(selAssumption), strStructSort)

                If intAssumptionRowIndex > Grid.Count - 1 Then
                    rowAssumption = New LogframeRow(intSection)
                    Me.Grid.Add(rowAssumption)
                Else
                    rowAssumption = Grid(intAssumptionRowIndex)
                End If

                With rowAssumption
                    .AssumptionSort = strAssumptionSort
                    .Assumption = selAssumption
                End With
                intAssumptionRowIndex += 1
            Next
            If Me.HideEmptyCells = False Then
                If intSection = Logframe.SectionTypes.ActivitiesSection And Me.ShowResourcesBudget = True Then
                    If GetRowCountOfStruct_Assumptions(objParentStruct) > GetRowCountOfStruct_Resources(objParentStruct) Then _
                        Me.Grid.Add(New LogframeRow(intSection))
                Else
                    If GetRowCountOfStruct_Assumptions(objParentStruct) > GetRowCountOfStruct_Indicators(objParentStruct) Then _
                        Me.Grid.Add(New LogframeRow(intSection))
                End If
            End If
        End If
    End Sub

    Private Function LoadSections_CreateSectionTitle(ByVal intSection As Integer) As LogframeRow
        Dim TitleRow As New LogframeRow
        Dim strStruct As String = String.Empty
        Dim objStruct As Struct = Nothing

        Select Case intSection
            Case Logframe.SectionTypes.GoalsSection
                strStruct = Logframe.StructNamePlural(0)
                objStruct = New Goal(LoadSections_SectionTitleSetFont(strStruct))
            Case Logframe.SectionTypes.PurposesSection
                strStruct = Logframe.StructNamePlural(1)
                objStruct = New Purpose(LoadSections_SectionTitleSetFont(strStruct))
            Case Logframe.SectionTypes.OutputsSection
                strStruct = Logframe.StructNamePlural(2)
                objStruct = New Output(LoadSections_SectionTitleSetFont(strStruct))
            Case Logframe.SectionTypes.ActivitiesSection
                strStruct = Logframe.StructNamePlural(3)
                objStruct = New Activity(LoadSections_SectionTitleSetFont(strStruct))
        End Select

        Dim objIndicator As New Indicator(LoadSections_SectionTitleSetFont(Logframe.IndicatorName))
        Dim objVerificationSource As New VerificationSource(LoadSections_SectionTitleSetFont(Logframe.VerificationSourceName))
        Dim objAssumption As New Assumption(LoadSections_SectionTitleSetFont(Logframe.AssumptionName))

        With TitleRow
            .Section = intSection
            .Struct = objStruct
            .Indicator = objIndicator
            .VerificationSource = objVerificationSource
            .Assumption = objAssumption

            If intSection = Logframe.SectionTypes.ActivitiesSection Then
                Dim objResource As New Resource(LoadSections_SectionTitleSetFont(Logframe.ResourceName))
                .Resource = objResource
            End If
        End With

        Return TitleRow
    End Function

    Private Function LoadSections_SectionTitleSetFont(ByVal strText As String) As String
        Using objRtf As New RichTextManager
            objRtf.Text = strText
            objRtf.SelectAll()
            objRtf.SelectionFont = Me.Logframe.SectionTitleFont
            objRtf.SelectionColor = Me.Logframe.SectionTitleFontColor
            objRtf.SelectionAlignment = HorizontalAlignment.Center
            strText = objRtf.Rtf
        End Using
        Return strText
    End Function

    Private Sub LoadSections_RepeatPurposes(ByVal selPurpose As Purpose, ByVal strPurposeSort As String)
        Dim rowPurpose As New LogframeRow(Logframe.SectionTypes.OutputsSection)
        With rowPurpose
            .StructSort = strPurposeSort
            .Struct = selPurpose
            .RowType = LogframeRow.RowTypes.RepeatPurpose
        End With
        Me.Grid.Add(rowPurpose)
    End Sub

    Private Sub LoadSections_RepeatOutputs(ByVal selOutput As Output, ByVal strOutputSort As String)
        Dim rowOutput As New LogframeRow(Logframe.SectionTypes.ActivitiesSection)
        With rowOutput
            .StructSort = strOutputSort
            .Struct = selOutput
            .RowType = LogframeRow.RowTypes.RepeatOutput
        End With
        Me.Grid.Add(rowOutput)
    End Sub
#End Region

#Region "Cell images"
    Private Sub ResetAllImages()
        For Each selRow As LogframeRow In Me.Grid
            ResetRowImages(selRow)
        Next
    End Sub

    Private Sub ResetRowImages(ByVal selRow As LogframeRow)
        selRow.StructImageDirty = True
        selRow.IndicatorImageDirty = True
        selRow.VerificationSourceImageDirty = True
        selRow.ResourceImageDirty = True
        selRow.AssumptionImageDirty = True

        selRow.StructHeight = 0
        selRow.IndicatorHeight = 0
        selRow.VerificationSourceHeight = 0
        selRow.ResourceHeight = 0
        selRow.AssumptionHeight = 0
    End Sub

    Private Sub ReloadImages()
        For Each selRow As LogframeRow In Me.Grid
            If selRow.RowType = LogframeRow.RowTypes.Normal Then
                ReloadImages_Normal(selRow)
            End If
        Next

        ResetRowHeights()
        For Each selRow As LogframeRow In Me.Grid
            If selRow.RowType = LogframeRow.RowTypes.Normal Then
                ReloadImages_MergedIndicators(selRow)
            End If
        Next

        For Each selRow As LogframeRow In Me.Grid
            If selRow.RowType = LogframeRow.RowTypes.Normal Then
                ReloadImages_MergedStructs(selRow)
            End If
        Next

    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As LogframeRow)
        Dim intColumnWidth As Integer
        Dim boolSelected As Boolean

        With RichTextManager
            If selRow.StructImageDirty = True And selRow.Struct IsNot Nothing Then
                intColumnWidth = StructRtfColumn.Width

                Select Case Me.TextSelectionIndex
                    Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe, TextSelectionValues.SelectStructs
                        boolSelected = True
                    Case Else
                        boolSelected = False
                End Select
                
                If String.IsNullOrEmpty(selRow.Struct.Text) Then
                    selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, selRow.Struct.GetItemName(selRow.Section), selRow.StructSort, boolSelected)
                Else
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.StructRtf, boolSelected, selRow.StructIndent)
                End If
            End If

            If ShowIndicatorColumn = True Then
                If selRow.IndicatorImageDirty = True And selRow.Indicator IsNot Nothing Then
                    intColumnWidth = IndicatorRtfColumn.Width

                    Select Case Me.TextSelectionIndex
                        Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe, TextSelectionValues.SelectIndicators
                            boolSelected = True
                        Case Else
                            boolSelected = False
                    End Select

                    If String.IsNullOrEmpty(selRow.Indicator.Text) Then
                        selRow.Indicator.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, Indicator.ItemName, selRow.IndicatorSort, boolSelected)
                    Else
                        selRow.Indicator.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.IndicatorRtf, boolSelected, selRow.IndicatorIndent)
                    End If
                End If

                If ShowVerificationSourceColumn = True Then
                    If selRow.VerificationSourceImageDirty = True And selRow.VerificationSource IsNot Nothing Then
                        intColumnWidth = VerificationSourceRtfColumn.Width

                        Select Case Me.TextSelectionIndex
                            Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe, TextSelectionValues.SelectVerificationSources
                                boolSelected = True
                            Case Else
                                boolSelected = False
                        End Select

                        selRow.VerificationSource.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.VerificationSourceRtf, boolSelected)
                        selRow.VerificationSourceHeight = selRow.VerificationSource.CellImage.Height
                        selRow.VerificationSourceImageDirty = False
                    End If
                End If

                If selRow.ResourceImageDirty = True And selRow.Resource IsNot Nothing Then
                    intColumnWidth = IndicatorRtfColumn.Width

                    Select Case Me.TextSelectionIndex
                        Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe, TextSelectionValues.SelectResources
                            boolSelected = True
                        Case Else
                            boolSelected = False
                    End Select

                    If String.IsNullOrEmpty(selRow.Resource.Text) Then
                        selRow.Resource.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, Resource.ItemName, selRow.ResourceSort, boolSelected)
                    Else
                        selRow.Resource.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.ResourceRtf, boolSelected)
                    End If
                    selRow.ResourceHeight = selRow.Resource.CellImage.Height
                    selRow.ResourceImageDirty = False
                End If
            End If

            If ShowAssumptionColumn = True Then
                If selRow.AssumptionImageDirty = True And selRow.Assumption IsNot Nothing Then
                    intColumnWidth = AssumptionRtfColumn.Width

                    Select Case Me.TextSelectionIndex
                        Case TextSelectionValues.SelectAll, TextSelectionValues.SelectLogframe, TextSelectionValues.SelectAssumptions
                            boolSelected = True
                        Case Else
                            boolSelected = False
                    End Select

                    selRow.Assumption.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.AssumptionRtf, boolSelected)
                    selRow.AssumptionHeight = selRow.Assumption.CellImage.Height
                    selRow.AssumptionImageDirty = False
                End If
            End If
        End With
    End Sub

    Private Sub ReloadImages_MergedIndicators(ByVal selRow As LogframeRow)
        Dim intRowIndex As Integer = Me.Grid.IndexOf(selRow)

        If ShowIndicatorColumn = True Then
            If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.CellImage IsNot Nothing Then
                Dim selIndicator As Indicator = selRow.Indicator
                Dim intRowCount As Integer = GetRowCountOfIndicator(selIndicator)

                If intRowCount <= 1 Then
                    selRow.IndicatorImage = selIndicator.CellImage
                    selRow.IndicatorHeight = selRow.Indicator.CellImage.Height
                    SetRowHeight(intRowIndex)
                Else
                    Dim bmpSource As Image = selRow.Indicator.CellImage
                    Dim intTop As Integer = 0

                    For i = 0 To intRowCount - 1
                        If intRowIndex + i > Me.Grid.Count - 1 Then Exit For
                        Dim dgvRow As DataGridViewRow = Rows(intRowIndex + i)
                        Dim objGridRow As LogframeRow = Grid(intRowIndex + i)
                        Dim intAvailableHeight As Integer = dgvRow.Height

                        If i = intRowCount - 1 Then
                            Dim intNeededHeight As Integer = bmpSource.Height - intTop
                            If intNeededHeight > intAvailableHeight Then intAvailableHeight = intNeededHeight
                        End If

                        Dim rCell As New Rectangle(0, intTop, colIndicatorRtfColumn.Width, intAvailableHeight)
                        Dim bmpIndicator As New Bitmap(rCell.Width, rCell.Height)
                        Dim gIndicator As Graphics = Graphics.FromImage(bmpIndicator)

                        gIndicator.DrawImage(bmpSource, 0, 0, rCell, GraphicsUnit.Pixel)
                        gIndicator.Dispose()
                        objGridRow.IndicatorImage = bmpIndicator
                        objGridRow.IndicatorImageDirty = False
                        objGridRow.IndicatorHeight = bmpIndicator.Height
                        intTop += objGridRow.IndicatorHeight

                        SetRowHeight(intRowIndex + i)
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub ReloadImages_MergedStructs(ByVal selRow As LogframeRow)
        Dim intRowIndex As Integer = Me.Grid.IndexOf(selRow)

        If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
            Dim selStruct As Struct = selRow.Struct
            Dim intRowCount As Integer = GetRowCountOfStruct(selStruct)

            If intRowCount <= 1 Then
                selRow.StructImage = selStruct.CellImage
                selRow.StructHeight = selRow.Struct.CellImage.Height
                SetRowHeight(intRowIndex)
            Else
                Dim bmpSource As Image = selRow.Struct.CellImage
                Dim intTop As Integer = 0
                For i = 0 To intRowCount - 1
                    If intRowIndex + i > Me.Grid.Count - 1 Then Exit For
                    Dim dgvRow As DataGridViewRow = Rows(intRowIndex + i)
                    Dim objGridRow As LogframeRow = Grid(intRowIndex + i)
                    Dim intAvailableHeight As Integer = dgvRow.Height

                    If i = intRowCount - 1 Then
                        Dim intNeededHeight As Integer = bmpSource.Height - intTop
                        If intNeededHeight > intAvailableHeight Then intAvailableHeight = intNeededHeight
                    End If

                    Dim rCell As New Rectangle(0, intTop, colStructRtfColumn.Width, intAvailableHeight)
                    Dim bmpStruct As New Bitmap(rCell.Width, rCell.Height)
                    Dim gStruct As Graphics = Graphics.FromImage(bmpStruct)

                    gStruct.DrawImage(bmpSource, 0, 0, rCell, GraphicsUnit.Pixel)
                    gStruct.Dispose()
                    objGridRow.StructImage = bmpStruct
                    objGridRow.StructImageDirty = False
                    objGridRow.StructHeight = bmpStruct.Height
                    intTop += objGridRow.StructHeight

                    SetRowHeight(intRowIndex + i)
                Next
            End If

        End If
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selRow As DataGridViewRow = Rows(RowIndex)
        Dim selGridRow As LogframeRow = Me.Grid(RowIndex)

        Dim intRowHeight As Integer
        If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose And My.Settings.setRepeatPurposes = False Then
            selRow.Visible = False
        ElseIf selGridRow.RowType = GridRow.RowTypes.RepeatOutput And My.Settings.setRepeatOutputs = False Then
            selRow.Visible = False
        Else
            selRow.Visible = True
            intRowHeight = CalculateRowHeight(RowIndex)
        End If

        If intRowHeight > 0 Then selRow.Height = intRowHeight Else selRow.Height = NewCellHeight()
    End Sub

    Public Sub ResetRowHeights()

        For Each selRow As DataGridViewRow In Me.Rows

            SetRowHeight(selRow.Index)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selGridRow As LogframeRow = Me.Grid(RowIndex)

        If selGridRow.StructHeight > intRowHeight Then intRowHeight = selGridRow.StructHeight

        If ShowIndicatorColumn = True Then
            If IsResourceBudgetRow(selGridRow) = False Then
                If selGridRow.IndicatorHeight > intRowHeight Then intRowHeight = selGridRow.IndicatorHeight

                If ShowVerificationSourceColumn = True Then
                    If selGridRow.VerificationSourceHeight > intRowHeight Then intRowHeight = selGridRow.VerificationSourceHeight
                End If
            Else
                If selGridRow.ResourceHeight > intRowHeight Then intRowHeight = selGridRow.ResourceHeight
            End If
        End If

        If ShowAssumptionColumn = True Then
            If selGridRow.AssumptionHeight > intRowHeight Then intRowHeight = selGridRow.AssumptionHeight
        End If

        Return intRowHeight
    End Function

    Private Function CalculateRowHeight_EditRow(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selGridRow As LogframeRow = Me.Grid(RowIndex)

        intRowHeight = Rows(RowIndex).Height

        If selGridRow.Struct IsNot Nothing AndAlso selGridRow.Struct.CellImage IsNot Nothing Then
            If selGridRow.Struct.CellImage.Height > intRowHeight Then intRowHeight = selGridRow.Struct.CellImage.Height
        End If

        If ShowIndicatorColumn = True Then
            If IsResourceBudgetRow(selGridRow) = False Then
                If selGridRow.Indicator IsNot Nothing AndAlso selGridRow.Indicator.CellImage IsNot Nothing Then
                    If selGridRow.Indicator.CellImage.Height > intRowHeight Then intRowHeight = selGridRow.Indicator.CellImage.Height
                End If

                If ShowVerificationSourceColumn = True Then
                    If selGridRow.VerificationSourceHeight > intRowHeight Then intRowHeight = selGridRow.VerificationSourceHeight
                End If
            Else
                If selGridRow.ResourceHeight > intRowHeight Then intRowHeight = selGridRow.ResourceHeight
            End If
        End If

        If ShowAssumptionColumn = True Then
            If selGridRow.AssumptionHeight > intRowHeight Then intRowHeight = selGridRow.AssumptionHeight
        End If

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

            Dim selItem As LogframeObject = GetCurrentItem()

            If selItem IsNot Nothing Then
                With UndoRedo
                    UndoRedo.UndoBuffer_Initialise(selItem, "RTF")
                End With
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

    Private Sub RichTextEditingControl_Validated(sender As Object, e As System.EventArgs) Handles RichTextEditingControl.Validated
        UndoRedo.TextChanged(RichTextEditingControl.Rtf)
    End Sub

    Protected Overrides Sub OnCurrentCellChanged(e As System.EventArgs)
        MyBase.OnCurrentCellChanged(e)

        Dim selItem As Object = GetCurrentItem()

        UndoRedo.UndoBuffer_Initialise(selItem)
    End Sub

    Protected Overrides Sub OnEditingControlShowing(ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)
        MyBase.OnEditingControlShowing(e)

        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim strColName As String = Me.Columns(Me.CurrentCell.ColumnIndex).Name
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim selObject As Object = Nothing
        Dim strSortNumber As String = String.Empty
        Dim intRowCountStruct, intRowCountIndicator, intNextRows As Integer

        If Me.TextSelectionIndex <> TextSelectionValues.SelectNothing Then
            Me.TextSelectionIndex = TextSelectionValues.SelectNothing
            ResetAllImages()
            ReloadImages()
            Invalidate()
            RaiseEvent NoTextSelected()
        End If

        Me.RichTextEditingControl = TryCast(e.Control, RichTextEditingControlLogframe)
        CurrentRow.Height = CalculateRowHeight_EditRow(Me.CurrentRow.Index)

        If selGridRow.StructImage IsNot Nothing Then
            selGridRow.StructImageDirty = True
            ReloadImages_Normal(selGridRow)
            ReloadImages_MergedStructs(selGridRow)
            If selGridRow.Struct IsNot Nothing Then intRowCountStruct = Me.GetRowCountOfStruct(selGridRow.Struct)
        End If
        If selGridRow.IndicatorImage IsNot Nothing Then
            selGridRow.IndicatorImageDirty = True
            ReloadImages_Normal(selGridRow)
            ReloadImages_MergedIndicators(selGridRow)
            If selGridRow.Indicator IsNot Nothing Then intRowCountIndicator = Me.GetRowCountOfIndicator(selGridRow.Indicator)
        End If
        intNextRows = intRowCountStruct
        If intRowCountIndicator > intNextRows Then intNextRows = intRowCountIndicator
        intNextRows = intRowIndex + intNextRows - 1

        For i = intRowIndex To intNextRows
            InvalidateRow(i)
        Next

        Dim selItem As LogframeObject = GetCurrentItem()

        If selItem IsNot Nothing Then
            UndoRedo.UndoBuffer_Initialise(selItem, "RTF")
        End If
    End Sub

    Private Sub DataGridViewLogframe_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Me.Scroll
        Invalidate()
    End Sub
#End Region

#Region "Manage rows"
    Public Sub SectionTitles_Protect()
        For Each selRow As DataGridViewRow In Me.Rows
            Select Case Me.Grid(selRow.Index).RowType
                Case GridRow.RowTypes.Section, GridRow.RowTypes.RepeatPurpose, GridRow.RowTypes.RepeatOutput
                    'make title bar of each section and repeated objectives or outputs inaccessible
                    For Each selCell As DataGridViewCell In selRow.Cells
                        selCell.ReadOnly = True
                    Next
                Case Else
                    'make numbers inaccessible
                    For Each selCell In selRow.Cells
                        Select Case Me.Columns(selCell.ColumnIndex).Name
                            Case "StructSort", "IndSort", "VerSort", "AsmSort"
                                selCell.ReadOnly = True
                            Case Else
                                selCell.ReadOnly = False

                        End Select
                    Next
            End Select
            'SetRowHeight(selRow.Index)
        Next
    End Sub

    Private Sub MoveCurrentCell()
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim intColumnIndex As Integer = CurrentCell.ColumnIndex
        Dim CurrentGridRow As LogframeRow = Me.Grid(CurrentCell.RowIndex)
        Dim strColumnName As String = Me.Columns(intColumnIndex).Name
        Dim objLogframeObject As LogframeObject = Nothing

        If CurrentGridRow.RowType <> LogframeRow.RowTypes.Normal Then
            'when user clicks in title row, move down
            intRowIndex += 1
            Me.CurrentCell = Me(intColumnIndex, intRowIndex)
            MoveCurrentCell()
        Else
            'make sure cells on same row are validated first
            If PreviousCellLocation.Y = intRowIndex Then
                Dim objCurrentCell As DataGridViewCell = Me(intColumnIndex, intRowIndex)
                Me.CurrentCell = Nothing
                Me.CurrentCell = objCurrentCell
                CurrentCell.Selected = False
            End If
        End If

        Select Case strColumnName
            Case "StructSort", "IndSort", "VerSort", "AsmSort"
                'when user clicks in number column, move to the right
                intColumnIndex += 1
                strColumnName = Me.Columns(intColumnIndex).Name
        End Select

        Select Case strColumnName
            Case "StructRTF"
                intRowIndex = MoveCurrentCell_Structs(CurrentGridRow, intRowIndex)

                If Grid(intRowIndex).Struct IsNot Nothing Then objLogframeObject = Grid(intRowIndex).Struct
            Case "IndRTF"
                If IsResourceBudgetRow(CurrentGridRow) = False Then
                    intRowIndex = MoveCurrentCell_Indicators(CurrentGridRow, intRowIndex)
                    If Grid(intRowIndex).Indicator IsNot Nothing Then objLogframeObject = Grid(intRowIndex).Indicator
                Else
                    intRowIndex = MoveCurrentCell_Resources(CurrentGridRow, intRowIndex)
                    If Grid(intRowIndex).Resource IsNot Nothing Then objLogframeObject = Grid(intRowIndex).Resource
                End If
            Case "VerRTF"
                If IsResourceBudgetRow(CurrentGridRow) = False Then
                    intRowIndex = MoveCurrentCell_VerificationSources(CurrentGridRow, intRowIndex)
                    If Grid(intRowIndex).VerificationSource IsNot Nothing Then objLogframeObject = Grid(intRowIndex).VerificationSource
                Else
                    intRowIndex = MoveCurrentCell_Resources(CurrentGridRow, intRowIndex)
                    If Grid(intRowIndex).Resource IsNot Nothing Then objLogframeObject = Grid(intRowIndex).Resource
                End If
            Case "AsmRTF"
                intRowIndex = MoveCurrentCell_Assumptions(CurrentGridRow, intRowIndex)
                If Grid(intRowIndex).Assumption IsNot Nothing Then objLogframeObject = Grid(intRowIndex).Assumption
        End Select

        Me.CurrentCell = Me(intColumnIndex, intRowIndex)
        PreviousCellLocation = New Point(intColumnIndex, intRowIndex)

        'If ClickPoint.IsEmpty = False Then
        If ClickPoint <> Nothing Then
            Me.BeginEdit(False)
            'Dim ctl As RichTextEditingControlLogframe = CType(Me.EditingControl, RTFeditingControlLogFrame)
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

        RaiseEvent LogframeObjectSelected(Me, New LogframeObjectSelectedEventArgs(objLogframeObject))
    End Sub

    Private Function MoveCurrentCell_Structs(ByVal CurrentGridRow As LogframeRow, ByVal intRowIndex As Integer) As Integer
        Dim selGridRow As LogframeRow

        If CurrentGridRow.Struct Is Nothing And CurrentGridRow.StructImage IsNot Nothing Then
            'clicked in a merged struct
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)
                If selGridRow.RowType <> LogframeRow.RowTypes.Normal Then
                    intRowIndex += 1
                    Exit Do
                End If
            Loop Until selGridRow.Struct IsNot Nothing
        End If

        Return intRowIndex
    End Function

    Private Function MoveCurrentCell_Indicators(ByVal CurrentGridRow As LogframeRow, ByVal intRowIndex As Integer) As Integer
        Dim selGridRow As LogframeRow

        If CurrentGridRow.Indicator Is Nothing And CurrentGridRow.IndicatorImage IsNot Nothing Then
            'clicked in a merged indicator
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)
                If CurrentGridRow.Struct IsNot Nothing Then
                    Exit Do
                ElseIf selGridRow.RowType <> LogframeRow.RowTypes.Normal Then
                    intRowIndex += 1
                    Exit Do
                End If
            Loop Until selGridRow.Indicator IsNot Nothing

        ElseIf CurrentGridRow.Indicator Is Nothing And CurrentGridRow.Struct Is Nothing And CurrentGridRow.StructImage IsNot Nothing Then
            'more assumptions than indicators, move to top indicator row
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)

                If selGridRow.RowType <> LogframeRow.RowTypes.Normal Or selGridRow.IndicatorImage IsNot Nothing Then
                    intRowIndex += 1
                    Exit Do
                ElseIf selGridRow.Struct IsNot Nothing Then
                    Exit Do
                End If
            Loop Until selGridRow.Indicator IsNot Nothing
        End If

        Return intRowIndex
    End Function

    Private Function MoveCurrentCell_Resources(ByVal CurrentGridRow As LogframeRow, ByVal intRowIndex As Integer) As Integer
        Dim selGridRow As LogframeRow

        If CurrentGridRow.Resource Is Nothing And CurrentGridRow.ResourceImage IsNot Nothing Then
            'clicked in a merged Resource
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)
                If CurrentGridRow.Struct IsNot Nothing Then
                    Exit Do
                ElseIf selGridRow.RowType <> LogframeRow.RowTypes.Normal Then
                    intRowIndex += 1
                    Exit Do
                End If
            Loop Until selGridRow.Resource IsNot Nothing

        ElseIf CurrentGridRow.Resource Is Nothing And CurrentGridRow.Struct Is Nothing And CurrentGridRow.StructImage IsNot Nothing Then
            'more assumptions than Resources, move to top Resource row
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)

                If selGridRow.RowType <> LogframeRow.RowTypes.Normal Or selGridRow.ResourceImage IsNot Nothing Then
                    intRowIndex += 1
                    Exit Do
                ElseIf selGridRow.Struct IsNot Nothing Then
                    Exit Do
                End If
            Loop Until selGridRow.Indicator IsNot Nothing
        End If

        Return intRowIndex
    End Function

    Private Function MoveCurrentCell_VerificationSources(ByVal CurrentGridRow As LogframeRow, ByVal intRowIndex As Integer) As Integer
        Dim selGridRow As LogframeRow

        If CurrentGridRow.VerificationSource Is Nothing And CurrentGridRow.IndicatorImage Is Nothing And CurrentGridRow.Struct Is Nothing And CurrentGridRow.StructImage IsNot Nothing Then
            'more assumptions than indicators/verification sources, move to top verification source row
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)
                If selGridRow.RowType <> LogframeRow.RowTypes.Normal Or selGridRow.VerificationSource IsNot Nothing Then
                    intRowIndex += 1
                    Exit Do
                ElseIf selGridRow.IndicatorImage IsNot Nothing Then 'selGridRow.Indicator Is Nothing And
                    intRowIndex += 1
                    Exit Do
                ElseIf selGridRow.Indicator IsNot Nothing Or selGridRow.Struct IsNot Nothing Then
                    Exit Do
                End If
            Loop Until selGridRow.VerificationSource IsNot Nothing
        End If

        Return intRowIndex
    End Function

    Private Function MoveCurrentCell_Assumptions(ByVal CurrentGridRow As LogframeRow, ByVal intRowIndex As Integer) As Integer
        Dim selGridRow As LogframeRow

        If CurrentGridRow.Assumption Is Nothing And CurrentGridRow.Struct Is Nothing And CurrentGridRow.StructImage IsNot Nothing Then
            'more indicators/verification sources than indicators, move to top assumption row
            Do
                intRowIndex -= 1
                selGridRow = Me.Grid(intRowIndex)
                If selGridRow.RowType <> LogframeRow.RowTypes.Normal Or selGridRow.Assumption IsNot Nothing Then
                    intRowIndex += 1
                    Exit Do
                ElseIf selGridRow.Struct IsNot Nothing Then
                    Exit Do
                End If
            Loop Until selGridRow.Assumption IsNot Nothing
        End If

        Return intRowIndex
    End Function
#End Region

#Region "Get row counts"
    Private Function GetRowCountOfStruct(ByVal objStruct As Struct) As Integer
        Dim NrRows As Integer
        Dim intRowCount As Integer
        Dim intRowCountAsm As Integer = GetRowCountOfStruct_Assumptions(objStruct)

        If objStruct IsNot Nothing Then
            If objStruct.ResourceBudget = False Then
                intRowCount = GetRowCountOfStruct_Indicators(objStruct)
            Else
                intRowCount = GetRowCountOfStruct_Resources(objStruct)
            End If
            If intRowCount >= intRowCountAsm Then NrRows = intRowCount Else NrRows = intRowCountAsm
        End If

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Indicators(ByVal objStruct As Struct) As Integer
        Dim NrRows As Integer
        'Dim boolDropLast As Boolean

        If ShowIndicatorColumn = True Then
            If objStruct IsNot Nothing Then
                NrRows = GetRowCountOfStruct_Indicators_Count(objStruct.Indicators)

                'If objStruct.Indicators.Count > 0 Then
                '    If objStruct.Indicators(objStruct.Indicators.Count - 1).Indicators.Count > 0 Then boolDropLast = True
                '    If NrRows > 0 And HideEmptyCells = False And boolDropLast = False Then NrRows += 1 '
                'End If

            End If

        End If
        If NrRows = 0 Then NrRows = 1

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Indicators_Count(ByVal selIndicators As Indicators) As Integer
        Dim NrRows As Integer
        Dim boolLastHasSubIndicators As Boolean

        For Each selIndicator As Indicator In selIndicators
            boolLastHasSubIndicators = False

            NrRows += GetRowCountOfIndicator(selIndicator)
            If selIndicator.Indicators.Count > 0 Then
                NrRows += GetRowCountOfStruct_Indicators_Count(selIndicator.Indicators)
                boolLastHasSubIndicators = True
                'NrRows += 1
            End If
        Next

        If Me.HideEmptyCells = False And boolLastHasSubIndicators = False Then
            NrRows += 1
        End If

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Resources(ByVal objActivity As Activity) As Integer
        Dim NmbrRows As Integer

        If ShowIndicatorColumn = True Then
            If objActivity IsNot Nothing Then
                NmbrRows = objActivity.Resources.Count
                If NmbrRows > 0 And HideEmptyCells = False Then NmbrRows += 1
            End If

        End If
        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function

    Private Function GetRowCountOfStruct_Assumptions(ByVal objStruct As Struct) As Integer
        Dim NmbrRows As Integer

        If ShowAssumptionColumn = True Then
            If objStruct IsNot Nothing Then
                NmbrRows = objStruct.Assumptions.Count
                If NmbrRows > 0 And HideEmptyCells = False Then NmbrRows += 1
            End If

        End If
        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function

    Private Function GetRowCountOfIndicator(ByVal selIndicator As Indicator) As Integer
        Dim NmbrRows As Integer

        If ShowVerificationSourceColumn = True Then
            If selIndicator IsNot Nothing AndAlso selIndicator.VerificationSources.Count > 0 Then
                NmbrRows = selIndicator.VerificationSources.Count
                If HideEmptyCells = False Then NmbrRows += 1
            End If

        End If
        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function
#End Region

End Class

Public Class LogframeObjectSelectedEventArgs
    Inherits EventArgs

    Public Property LogframeObject As LogframeObject

    Public Sub New(ByVal objLogframeObject As LogframeObject)
        MyBase.New()

        Me.LogframeObject = objLogframeObject
    End Sub
End Class
