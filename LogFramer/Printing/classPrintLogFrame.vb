Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports System.Reflection

Public Class PrintLogFrame
    Inherits ReportBaseClass

    Private objLogframe As LogFrame
    Private objPrintList As New PrintLogframeRows
    Private objClippedRow As PrintLogframeRow = Nothing

    Private intStructSortWidth, intStructRtfWidth, intIndicatorSortWidth, intIndicatorRtfWidth, intVerificationSourceSortWidth, intVerificationSourceRtfWidth, intAssumptionSortWidth, intAssumptionRtfWidth As Integer
    Private boolColumnsWidthSet As Boolean

    Private bmStructOverFlow As Bitmap, bmIndOverFlow As Bitmap
    Private strClippedTextTop, strClippedTextBottom As String
    Private boolShowIndicatorColumn As Boolean
    Private boolShowVerificationSourceColumn As Boolean
    Private boolShowAssumptionColumn As Boolean
    Private boolShowGoals As Boolean = True
    Private boolShowPurposes As Boolean = True
    Private boolShowOutputs As Boolean = True
    Private boolShowActivities As Boolean = True
    Private boolShowResourcesBudget As Boolean = True

    Public Event LinePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, _
                   ByVal structrtfwidth As Integer, ByVal indrtfwidth As Integer, _
                   ByVal verrtfwidth As Integer, ByVal asmrtfwidth As Integer, _
                   ByVal boolShowIndColumn As Boolean, ByVal boolShowVerColumn As Boolean, ByVal boolShowAsmColumn As Boolean, _
                   ByVal boolShowGoals As Boolean, boolShowPurposes As Boolean, boolShowOutputs As Boolean, boolShowActivities As Boolean, boolShowResourcesBudget As Boolean)
        Me.Logframe = logframe
        Me.ReportSetup = logframe.ReportSetupLogframe
        Me.ShowIndicatorColumn = boolShowIndColumn
        Me.ShowVerificationSourceColumn = boolShowVerColumn
        Me.ShowAssumptionColumn = boolShowAsmColumn
        Me.ShowGoals = boolShowGoals
        Me.ShowPurposes = boolShowPurposes
        Me.ShowOutputs = boolShowOutputs
        Me.ShowActivities = boolShowActivities
        Me.ShowResourcesBudget = boolShowResourcesBudget
        Me.StructRTFWidth = structrtfwidth
        Me.IndicatorRTFWidth = indrtfwidth
        Me.VerificationSourceRTFWidth = verrtfwidth
        Me.AssumptionRTFWidth = asmrtfwidth
    End Sub

#Region "Properties"
    Public Property Logframe As LogFrame
        Get
            Return objLogframe
        End Get
        Set(ByVal value As LogFrame)
            objLogframe = value
        End Set
    End Property

    Public Property PrintList() As PrintLogframeRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintLogframeRows)
            objPrintList = value
        End Set
    End Property

    Private Property StructSortWidth() As Integer
        Get
            Return intStructSortWidth
        End Get
        Set(ByVal value As Integer)
            intStructSortWidth = value
        End Set
    End Property

    Private Property StructRTFWidth() As Integer
        Get
            Return intStructRtfWidth
        End Get
        Set(ByVal value As Integer)
            intStructRtfWidth = value
        End Set
    End Property

    Private Property IndicatorSortWidth() As Integer
        Get
            If ShowIndicatorColumn = True Then
                Return intIndicatorSortWidth
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intIndicatorSortWidth = value
        End Set
    End Property

    Private Property IndicatorRTFWidth() As Integer
        Get
            If ShowIndicatorColumn = True Then
                Return intIndicatorRtfWidth
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intIndicatorRtfWidth = value
        End Set
    End Property

    Private Property VerificationSourceSortWidth() As Integer
        Get
            If ShowVerificationSourceColumn = True Then
                Return intVerificationSourceSortWidth
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intVerificationSourceSortWidth = value
        End Set
    End Property

    Private Property VerificationSourceRTFWidth() As Integer
        Get
            If ShowVerificationSourceColumn = True Then
                Return intVerificationSourceRtfWidth
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intVerificationSourceRtfWidth = value
        End Set
    End Property

    Private Property AssumptionSortWidth() As Integer
        Get
            If ShowAssumptionColumn = True Then
                Return intAssumptionSortWidth
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intAssumptionSortWidth = value
        End Set
    End Property

    Private Property AssumptionRTFWidth() As Integer
        Get
            If ShowAssumptionColumn = True Then
                Return intAssumptionRtfWidth
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intAssumptionRtfWidth = value
        End Set
    End Property

    Private ReadOnly Property TotalSortWidth() As Integer
        Get
            Return Me.StructSortWidth + Me.IndicatorSortWidth + Me.VerificationSourceSortWidth + Me.AssumptionSortWidth
        End Get
    End Property

    Private ReadOnly Property TotalRTFWidth() As Integer
        Get
            Return Me.StructRTFWidth + Me.IndicatorRTFWidth + Me.VerificationSourceRTFWidth + Me.AssumptionRTFWidth
        End Get
    End Property

    Private ReadOnly Property TotalWidth() As Integer
        Get
            Return Me.TotalSortWidth + Me.TotalRTFWidth
        End Get
    End Property

    Public Property ShowIndicatorColumn() As Boolean
        Get
            Return boolShowIndicatorColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowIndicatorColumn = value
        End Set
    End Property

    Public Property ShowVerificationSourceColumn() As Boolean
        Get
            Return boolShowVerificationSourceColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowVerificationSourceColumn = value
        End Set
    End Property

    Public Property ShowAssumptionColumn() As Boolean
        Get
            Return boolShowAssumptionColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowAssumptionColumn = value
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

    Public Property ClippedRow As PrintLogframeRow
        Get
            Return objClippedRow
        End Get
        Set(value As PrintLogframeRow)
            objClippedRow = value
        End Set
    End Property
#End Region

#Region "Load sections"
    Public Sub LoadSections()
        Me.PrintList.Clear()

        If ShowGoals = False And ShowPurposes = False And ShowOutputs = False And ShowActivities = False Then
            ShowPurposes = True
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
        Dim rowSectionTitle As PrintLogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.GoalsSection)
        rowSectionTitle.RowType = PrintLogframeRow.RowTypes.Section

        Me.PrintList.Add(rowSectionTitle)

        Dim intGoalRowIndex As Integer
        Dim strGoalSort As String

        If Me.Logframe.Goals.Count > 0 Then
            For Each selGoal As Goal In Me.Logframe.Goals
                intGoalRowIndex = Me.PrintList.Count
                strGoalSort = Logframe.CreateSortNumber(Logframe.Goals.IndexOf(selGoal))

                Dim rowGoal As New PrintLogframeRow
                With rowGoal
                    .Section = Logframe.SectionTypes.GoalsSection
                    .StructSort = strGoalSort
                    .Struct = New Goal(selGoal.RTF)
                    .RowCountStruct = GetRowCountOfStruct(selGoal)
                End With
                Me.PrintList.Add(rowGoal)
                If selGoal.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.GoalsSection, selGoal, intGoalRowIndex, strGoalSort)
                If selGoal.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.GoalsSection, selGoal, intGoalRowIndex, strGoalSort)
            Next
        End If
    End Sub

    Private Sub LoadSections_SectionPurposes()
        Dim rowSectionTitle As PrintLogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.PurposesSection)
        rowSectionTitle.RowType = PrintLogframeRow.RowTypes.Section
        Me.PrintList.Add(rowSectionTitle)

        Dim intPurposeRowIndex As Integer
        Dim strPurposeSort As String

        If Me.Logframe.Purposes.Count > 0 Then
            For Each selPurpose As Purpose In Me.Logframe.Purposes
                intPurposeRowIndex = Me.PrintList.Count
                strPurposeSort = Logframe.CreateSortNumber(Logframe.Purposes.IndexOf(selPurpose))

                Dim rowPurpose As New PrintLogframeRow
                With rowPurpose
                    .Section = Logframe.SectionTypes.PurposesSection
                    .StructSort = strPurposeSort
                    .Struct = New Purpose(selPurpose.RTF)
                    .RowCountStruct = GetRowCountOfStruct(selPurpose)
                End With
                Me.PrintList.Add(rowPurpose)
                If selPurpose.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.PurposesSection, selPurpose, intPurposeRowIndex, strPurposeSort)
                If selPurpose.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.PurposesSection, selPurpose, intPurposeRowIndex, strPurposeSort)
            Next
        End If
    End Sub

    Private Sub LoadSections_SectionOutputs()
        Dim rowSectionTitle As PrintLogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.OutputsSection)
        rowSectionTitle.RowType = PrintLogframeRow.RowTypes.Section
        Me.PrintList.Add(rowSectionTitle)

        Dim intOutputRowIndex As Integer
        Dim strPurposeSort As String, strOutputSort As String

        For Each selPurpose As Purpose In Me.Logframe.Purposes
            strPurposeSort = Logframe.CreateSortNumber(Logframe.Purposes.IndexOf(selPurpose))

            If Me.Logframe.Purposes.Count > 1 Then LoadSections_RepeatPurposes(selPurpose, strPurposeSort)

            For Each selOutput As Output In selPurpose.Outputs
                intOutputRowIndex = Me.PrintList.Count
                strOutputSort = Logframe.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)

                Dim rowOutput As New PrintLogframeRow
                With rowOutput
                    .Section = Logframe.SectionTypes.OutputsSection
                    .StructSort = strOutputSort
                    .Struct = New Output(selOutput.RTF)
                    .RowCountStruct = GetRowCountOfStruct(selOutput)
                End With
                Me.PrintList.Add(rowOutput)
                If selOutput.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.OutputsSection, selOutput, intOutputRowIndex, strOutputSort)
                If selOutput.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.OutputsSection, selOutput, intOutputRowIndex, strOutputSort)
            Next
        Next
    End Sub

    Private Sub LoadSections_SectionActivities()
        Dim rowSectionTitle As PrintLogframeRow = LoadSections_CreateSectionTitle(Logframe.SectionTypes.ActivitiesSection)
        rowSectionTitle.RowType = PrintLogframeRow.RowTypes.Section
        Me.PrintList.Add(rowSectionTitle)

        Dim strPurposeSort As String, strOutputSort As String

        For Each selPurpose As Purpose In Me.Logframe.Purposes
            strPurposeSort = Logframe.CreateSortNumber(Logframe.Purposes.IndexOf(selPurpose))
            If Me.Logframe.Purposes.Count > 1 Then LoadSections_RepeatPurposes(selPurpose, strPurposeSort)

            For Each selOutput As Output In selPurpose.Outputs
                strOutputSort = Logframe.CreateSortNumber(selPurpose.Outputs.IndexOf(selOutput), strPurposeSort)
                If selPurpose.Outputs.Count > 1 Then LoadSections_RepeatOutputs(selOutput, strOutputSort)

                LoadSections_SectionActivities_Activities(selOutput.Activities, strOutputSort, 0)
            Next
        Next
    End Sub

    Private Sub LoadSections_SectionActivities_Activities(ByVal selActivities As Activities, ByVal strParentSort As String, ByVal intLevel As Integer)
        Dim intActivityRowIndex As Integer
        Dim strActivitySort As String

        For Each selActivity As Activity In selActivities
            intActivityRowIndex = Me.PrintList.Count
            strActivitySort = Logframe.CreateSortNumber(selActivities.IndexOf(selActivity), strParentSort)

            Dim rowActivity As New PrintLogframeRow
            With rowActivity
                .Section = Logframe.SectionTypes.ActivitiesSection
                .StructSort = strActivitySort
                .Struct = New Activity(selActivity.RTF)
                .StructIndent = intLevel
                .RowCountStruct = GetRowCountOfStruct(selActivity)
            End With
            Me.PrintList.Add(rowActivity)

            If IsResourceBudgetRow(rowActivity) = False Then
                If selActivity.Indicators.Count > 0 Then LoadSections_Indicators(Logframe.SectionTypes.ActivitiesSection, selActivity, intActivityRowIndex, strActivitySort)
            Else
                If selActivity.Resources.Count > 0 Then LoadSections_Resources(Logframe.SectionTypes.ActivitiesSection, selActivity, intActivityRowIndex, strActivitySort)
            End If
            If selActivity.Assumptions.Count > 0 Then LoadSections_Assumptions(Logframe.SectionTypes.ActivitiesSection, selActivity, intActivityRowIndex, strActivitySort)
            If selActivity.Activities.Count > 0 Then LoadSections_SectionActivities_Activities(selActivity.Activities, strActivitySort, intLevel + 1)
        Next
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
        Dim rowIndicator As PrintLogframeRow
        Dim selIndicator As Indicator = Nothing

        For Each selIndicator In selIndicators
            boolLastHasSubIndicators = False


            If Me.Logframe.SortNumberRepeatParent = True Then strIndicatorSort = strParentSort Else strIndicatorSort = String.Empty
            strIndicatorSort = Logframe.CreateSortNumber(selIndicators.IndexOf(selIndicator), strParentSort)

            If boolFirstOfStruct = True Then
                rowIndicator = PrintList(intIndicatorRowIndex)
                boolFirstOfStruct = False
            Else
                rowIndicator = New PrintLogframeRow(intSection)
                Me.PrintList.Add(rowIndicator)
            End If

            With rowIndicator
                .IndicatorSort = strIndicatorSort
                .Indicator = New Indicator(selIndicator.RTF)
                .RowCountIndicator = GetRowCountOfIndicator(selIndicator)
                .IndicatorIndent = intLevel
            End With

            If selIndicator.VerificationSources.Count > 0 Then
                LoadSections_VerificationSources(intSection, selIndicator, intIndicatorRowIndex, strIndicatorSort)
            End If

            intIndicatorRowIndex = Me.PrintList.Count
            If selIndicator.Indicators.Count > 0 Then
                LoadSections_SectionIndicators_Indicators(intSection, selIndicator.Indicators, intIndicatorRowIndex, strIndicatorSort, intLevel + 1, False)
                boolLastHasSubIndicators = True
            End If
            intIndicatorRowIndex = Me.PrintList.Count
        Next
    End Sub

    Private Sub LoadSections_VerificationSources(ByVal intSection As Integer, ByVal objParentIndicator As Indicator, ByVal intVerificationSourceRowIndex As Integer, ByVal strIndicatorSort As String)
        Dim strVerificationSourceSort As String
        Dim rowVerificationSource As PrintLogframeRow

        If ShowVerificationSourceColumn = True Then
            For Each selVerificationSource As VerificationSource In objParentIndicator.VerificationSources
                If Me.Logframe.SortNumberRepeatParent = True Then strVerificationSourceSort = strIndicatorSort Else strVerificationSourceSort = String.Empty
                strVerificationSourceSort = Logframe.CreateSortNumber(objParentIndicator.VerificationSources.IndexOf(selVerificationSource), strIndicatorSort)

                If intVerificationSourceRowIndex > PrintList.Count - 1 Then
                    rowVerificationSource = New PrintLogframeRow(intSection)
                    Me.PrintList.Add(rowVerificationSource)
                Else
                    rowVerificationSource = PrintList(intVerificationSourceRowIndex)
                End If

                With rowVerificationSource
                    .VerificationSourceSort = strVerificationSourceSort
                    .VerificationSource = New VerificationSource(selVerificationSource.RTF)
                End With
                intVerificationSourceRowIndex = Me.PrintList.Count
            Next
        End If
    End Sub

    Private Sub LoadSections_Resources(ByVal intSection As Integer, ByVal objParentStruct As Struct, ByVal intStructRowIndex As Integer, ByVal strStructSort As String)
        Dim intResourceRowIndex As Integer = intStructRowIndex
        Dim strResourceSort As String
        Dim rowResource As PrintLogframeRow
        Dim selActivity As Activity = TryCast(objParentStruct, Activity)

        If selActivity Is Nothing Then Exit Sub

        If ShowIndicatorColumn = True Then
            For Each selResource As Resource In selActivity.Resources
                If Me.Logframe.SortNumberRepeatParent = True Then strResourceSort = strStructSort Else strResourceSort = String.Empty
                strResourceSort = Logframe.CreateSortNumber(selActivity.Resources.IndexOf(selResource), strStructSort)

                If intResourceRowIndex > PrintList.Count - 1 Then
                    rowResource = New PrintLogframeRow(intSection)
                    Me.PrintList.Add(rowResource)
                Else
                    rowResource = PrintList(intResourceRowIndex)
                End If

                With rowResource
                    .ResourceSort = strResourceSort
                    .Resource = New Resource(selResource.RTF)
                    .TotalCostAmount = selResource.TotalCostAmount
                End With

                intResourceRowIndex = Me.PrintList.Count
            Next
        End If
    End Sub

    Private Sub LoadSections_Assumptions(ByVal intSection As Integer, ByVal objParentStruct As Struct, ByVal intStructRowIndex As Integer, ByVal strStructSort As String)
        Dim intAssumptionRowIndex As Integer = intStructRowIndex
        Dim strAssumptionSort As String
        Dim rowAssumption As PrintLogframeRow

        If ShowAssumptionColumn = True Then
            For Each selAssumption As Assumption In objParentStruct.Assumptions
                If Me.Logframe.SortNumberRepeatParent = True Then strAssumptionSort = strStructSort Else strAssumptionSort = String.Empty
                strAssumptionSort = Logframe.CreateSortNumber(objParentStruct.Assumptions.IndexOf(selAssumption), strStructSort)

                If intAssumptionRowIndex > PrintList.Count - 1 Then
                    rowAssumption = New PrintLogframeRow(intSection)
                    Me.PrintList.Add(rowAssumption)
                Else
                    rowAssumption = PrintList(intAssumptionRowIndex)
                End If

                With rowAssumption
                    .AssumptionSort = strAssumptionSort
                    .Assumption = New Assumption(selAssumption.RTF)
                End With
                intAssumptionRowIndex += 1
            Next
        End If
    End Sub

    Private Function LoadSections_CreateSectionTitle(ByVal intSection As Integer) As PrintLogframeRow
        Dim TitleRow As New PrintLogframeRow
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
            If ShowIndicatorColumn = True Then .Indicator = objIndicator
            If ShowVerificationSourceColumn = True Then .VerificationSource = objVerificationSource
            If ShowAssumptionColumn = True Then .Assumption = objAssumption

            If intSection = Logframe.SectionTypes.ActivitiesSection And ShowIndicatorColumn = True Then
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
        Dim rowPurpose As New PrintLogframeRow(Logframe.SectionTypes.OutputsSection)
        With rowPurpose
            .StructSort = strPurposeSort
            .Struct = selPurpose
            .RowType = PrintLogframeRow.RowTypes.RepeatPurpose
        End With
        Me.PrintList.Add(rowPurpose)
    End Sub

    Private Sub LoadSections_RepeatOutputs(ByVal selOutput As Output, ByVal strOutputSort As String)
        Dim rowOutput As New PrintLogframeRow(Logframe.SectionTypes.ActivitiesSection)
        With rowOutput
            .StructSort = strOutputSort
            .Struct = selOutput
            .RowType = PrintLogframeRow.RowTypes.RepeatOutput
        End With
        Me.PrintList.Add(rowOutput)
    End Sub
#End Region

#Region "Set column widths"
    Private Function GetStructSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintLogframeRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.StructSort) = False AndAlso selRow.StructSort.Length > strSort.Length Then strSort = selRow.StructSort
        Next
        If strSort <> String.Empty Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + 4
        End If
        Return intWidth
    End Function

    Private Function GetIndicatorSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintLogframeRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.IndicatorSort) = False AndAlso selRow.IndicatorSort.Length > strSort.Length Then strSort = selRow.IndicatorSort
            If String.IsNullOrEmpty(selRow.ResourceSort) = False AndAlso selRow.ResourceSort.Length > strSort.Length Then strSort = selRow.ResourceSort
        Next
        If strSort <> String.Empty Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + 4
        End If
        Return intWidth
    End Function

    Private Function GetVerificationSourceSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintLogframeRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.VerificationSourceSort) = False AndAlso selRow.VerificationSourceSort.Length > strSort.Length Then strSort = selRow.VerificationSourceSort
        Next
        If strSort <> String.Empty Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + 4
        End If
        Return intWidth
    End Function

    Private Function GetAssumptionSortColumnWidth() As Integer
        Dim intWidth As Integer
        Dim strSort As String = String.Empty

        For Each selRow As PrintLogframeRow In Me.PrintList
            If String.IsNullOrEmpty(selRow.AssumptionSort) = False AndAlso selRow.AssumptionSort.Length > strSort.Length Then strSort = selRow.AssumptionSort
        Next
        If strSort <> String.Empty Then
            intWidth = PageGraph.MeasureString(strSort, CurrentLogFrame.SortNumberFont).Width + 4
        End If
        Return intWidth
    End Function

    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer
        Dim sngDivider As Single

        If PageGraph IsNot Nothing Then
            Me.StructSortWidth = GetStructSortColumnWidth()
            If ShowIndicatorColumn = True Then IndicatorSortWidth = GetIndicatorSortColumnWidth() Else IndicatorSortWidth = 0
            If ShowVerificationSourceColumn = True Then Me.VerificationSourceSortWidth = GetVerificationSourceSortColumnWidth() Else VerificationSourceSortWidth = 0
            If ShowAssumptionColumn = True Then Me.AssumptionSortWidth = GetAssumptionSortColumnWidth() Else AssumptionSortWidth = 0

            intAvailableWidth = PrintWidth - Me.TotalSortWidth
            sngDivider = intAvailableWidth / Me.TotalRTFWidth

            Me.StructRTFWidth *= sngDivider
            Me.IndicatorRTFWidth *= sngDivider
            Me.VerificationSourceRTFWidth *= sngDivider
            Me.AssumptionRTFWidth *= sngDivider
        End If
    End Sub
#End Region

#Region "Cell images"
    Private Sub ReloadImages()
        For Each selRow As PrintLogframeRow In Me.PrintList
            If selRow.RowType = PrintLogframeRow.RowTypes.Normal Then
                ReloadImages_Normal(selRow)
            End If
        Next

        ResetRowHeights()
        For Each selRow As PrintLogframeRow In Me.PrintList
            If selRow.RowType = PrintLogframeRow.RowTypes.Normal Then
                ResetRowHeights_MergedIndicators(selRow)
            End If
        Next

        For Each selRow As PrintLogframeRow In Me.PrintList
            If selRow.RowType = PrintLogframeRow.RowTypes.Normal Then
                ResetRowHeights_MergedStructs(selRow)
            End If
        Next

    End Sub

    Private Sub ReloadImages_Normal(ByVal selRow As PrintLogframeRow)
        Dim intColumnWidth As Integer

        With RichTextManager
            If selRow.Struct IsNot Nothing Then
                intColumnWidth = StructRTFWidth

                If String.IsNullOrEmpty(selRow.Struct.Text) Then
                    selRow.Struct.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, selRow.Struct.GetItemName(selRow.Section), selRow.StructSort, False)
                Else
                    selRow.Struct.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.StructRtf, False, selRow.StructIndent)
                End If
            End If
            If ShowIndicatorColumn = True And selRow.Indicator IsNot Nothing Then
                intColumnWidth = IndicatorRTFWidth

                If String.IsNullOrEmpty(selRow.Indicator.Text) Then
                    selRow.Indicator.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, Indicator.ItemName, selRow.IndicatorSort, False)
                Else
                    selRow.Indicator.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.IndicatorRtf, False, selRow.IndicatorIndent)
                End If
            End If
            If ShowVerificationSourceColumn = True And selRow.VerificationSource IsNot Nothing Then
                intColumnWidth = VerificationSourceRTFWidth

                selRow.VerificationSource.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.VerificationSourceRtf, False)
                selRow.VerificationSourceHeight = selRow.VerificationSource.CellImage.Height
            End If
            If ShowIndicatorColumn = True And selRow.Resource IsNot Nothing Then
                intColumnWidth = IndicatorRTFWidth

                If String.IsNullOrEmpty(selRow.Resource.Text) Then
                    selRow.Resource.CellImage = .EmptyTextWithPaddingToBitmap(intColumnWidth, Resource.ItemName, selRow.ResourceSort, False)
                Else
                    selRow.Resource.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.ResourceRtf, False)
                End If
                selRow.ResourceHeight = selRow.Resource.CellImage.Height
            End If
            If ShowAssumptionColumn = True And selRow.Assumption IsNot Nothing Then
                intColumnWidth = AssumptionRTFWidth

                selRow.Assumption.CellImage = .RichTextWithPaddingToBitmap(intColumnWidth, selRow.AssumptionRtf, False)
                selRow.AssumptionHeight = selRow.Assumption.CellImage.Height
            End If
        End With
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintLogframeRow = Me.PrintList(RowIndex)

        Dim intRowHeight As Integer
        If selPrintListRow.RowType = PrintLogframeRow.RowTypes.Section Then
            intRowHeight = CalculateSectionRowHeight(selPrintListRow)
        ElseIf selPrintListRow.RowType = PrintLogframeRow.RowTypes.RepeatPurpose And My.Settings.setRepeatPurposes = True Then
            intRowHeight = CalculateRepeatsRowHeight(selPrintListRow)
        ElseIf selPrintListRow.RowType = PrintLogframeRow.RowTypes.RepeatOutput And My.Settings.setRepeatOutputs = True Then
            intRowHeight = CalculateRepeatsRowHeight(selPrintListRow)
        Else
            intRowHeight = CalculateRowHeight(RowIndex)
        End If

        If intRowHeight > 0 Then selPrintListRow.RowHeight = intRowHeight Else selPrintListRow.RowHeight = NewCellHeight()
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim intRowHeight As Integer
        Dim selPrintListRow As PrintLogframeRow = Me.PrintList(RowIndex)

        If selPrintListRow.AssumptionHeight > intRowHeight Then intRowHeight = selPrintListRow.AssumptionHeight

        If IsResourceBudgetRow(selPrintListRow) = False Then
            'If selPrintListRow.IndicatorHeight > intRowHeight Then intRowHeight = selPrintListRow.IndicatorHeight
            If String.IsNullOrEmpty(selPrintListRow.VerificationSourceRtf) = False Then
                If selPrintListRow.VerificationSourceHeight > intRowHeight Then intRowHeight = selPrintListRow.VerificationSourceHeight
            Else
                If String.IsNullOrEmpty(selPrintListRow.IndicatorRtf) = False Then
                    If selPrintListRow.IndicatorHeight > intRowHeight Then intRowHeight = selPrintListRow.IndicatorHeight
                Else
                    If selPrintListRow.StructHeight > intRowHeight Then intRowHeight = selPrintListRow.StructHeight
                End If
            End If
        Else
            If String.IsNullOrEmpty(selPrintListRow.ResourceRtf) = False Then
                If selPrintListRow.ResourceHeight > intRowHeight Then intRowHeight = selPrintListRow.ResourceHeight
            Else
                If selPrintListRow.StructHeight > intRowHeight Then intRowHeight = selPrintListRow.StructHeight
            End If

        End If

        Return intRowHeight
    End Function

    Private Function CalculateSectionRowHeight(ByVal selRow As PrintLogframeRow) As Integer
        Dim intRowHeight As Integer
        Dim objStringFormat As New StringFormat()
        Dim rIndicator, rVerificationSource, rAssumption As Rectangle

        objStringFormat = StringFormat.GenericDefault
        objStringFormat.Alignment = StringAlignment.Center
        objStringFormat.LineAlignment = StringAlignment.Center

        If PageGraph IsNot Nothing Then
            Dim rStruct As New Rectangle(LeftMargin, LastRowY, StructSortWidth + StructRTFWidth, 0)
            rStruct.Height = PageGraph.MeasureString(selRow.Struct.Text, Logframe.SectionTitleFont, rStruct.Width, objStringFormat).Height

            If ShowIndicatorColumn = True Then
                rIndicator = New Rectangle(rStruct.Right + 1, LastRowY, IndicatorSortWidth + IndicatorRTFWidth, 0)
                rIndicator.Height = PageGraph.MeasureString(selRow.Indicator.Text, Logframe.SectionTitleFont, rIndicator.Width, objStringFormat).Height
            End If

            If ShowVerificationSourceColumn = True Then
                rVerificationSource = New Rectangle(rIndicator.Right + 1, LastRowY, VerificationSourceSortWidth + VerificationSourceRTFWidth, 0)
                rVerificationSource.Height = PageGraph.MeasureString(selRow.VerificationSource.Text, Logframe.SectionTitleFont, rVerificationSource.Width, objStringFormat).Height
            End If

            If ShowAssumptionColumn = True Then
                rAssumption = New Rectangle(rVerificationSource.Right + 1, LastRowY, AssumptionSortWidth + AssumptionRTFWidth, 0)
                rAssumption.Height = PageGraph.MeasureString(selRow.Assumption.Text, Logframe.SectionTitleFont, rAssumption.Width, objStringFormat).Height
            End If

            If rStruct.Height > intRowHeight Then intRowHeight = rStruct.Height
            If ShowIndicatorColumn = True And rIndicator.Height > intRowHeight Then intRowHeight = rIndicator.Height
            If ShowVerificationSourceColumn = True And rVerificationSource.Height > intRowHeight Then intRowHeight = rVerificationSource.Height
            If ShowAssumptionColumn = True And rAssumption.Height > intRowHeight Then intRowHeight = rAssumption.Height
        End If

        Return intRowHeight
    End Function

    Private Function CalculateRepeatsRowHeight(ByVal selRow As PrintLogframeRow) As Integer
        Dim intRowHeight As Integer
        Dim objStringFormat As New StringFormat()
        Dim rRow As New Rectangle(LeftMargin, LastRowY, TotalWidth, selRow.RowHeight)

        objStringFormat = StringFormat.GenericDefault
        objStringFormat.Alignment = StringAlignment.Near
        objStringFormat.LineAlignment = StringAlignment.Near

        If PageGraph IsNot Nothing Then
            Dim rStruct As New Rectangle(LeftMargin + StructSortWidth, LastRowY, TotalWidth - StructSortWidth, 0)
            intRowHeight = PageGraph.MeasureString(selRow.Struct.Text, Logframe.SectionTitleFont, rStruct.Width, objStringFormat).Height
        End If

        Return intRowHeight
    End Function

    Private Sub ResetRowHeights_MergedIndicators(ByVal selRow As PrintLogframeRow)
        Dim intRowIndex As Integer = Me.PrintList.IndexOf(selRow)

        If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.CellImage IsNot Nothing Then
            Dim selIndicator As Indicator = selRow.Indicator
            Dim intRowCount As Integer = selRow.RowCountIndicator
            If ClippedRow IsNot Nothing Then intRowCount += 1

            If intRowCount <= 1 Then
                selRow.IndicatorHeight = selRow.Indicator.CellImage.Height
                If selRow.RowHeight < selRow.IndicatorHeight Then selRow.RowHeight = selRow.IndicatorHeight
            Else
                Dim intIndicatorHeight As Integer = selRow.Indicator.CellImage.Height
                Dim intTop As Integer = 0
                For i = 0 To intRowCount - 1
                    If intRowIndex + i > Me.PrintList.Count - 1 Then Exit For

                    Dim objPrintListRow As PrintLogframeRow = PrintList(intRowIndex + i)
                    Dim intAvailableHeight As Integer = objPrintListRow.RowHeight

                    If i = intRowCount - 1 Then
                        Dim intNeededHeight As Integer = intIndicatorHeight - intTop
                        If intNeededHeight > intAvailableHeight Then objPrintListRow.RowHeight = intNeededHeight
                    End If

                    If intAvailableHeight > 0 Then
                        objPrintListRow.IndicatorHeight = intAvailableHeight 'bmpIndicator.Height
                        intTop += objPrintListRow.RowHeight
                    Else
                        objPrintListRow.IndicatorHeight = 0
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ResetRowHeights_MergedStructs(ByVal selRow As PrintLogframeRow)
        Dim intRowIndex As Integer = Me.PrintList.IndexOf(selRow)

        If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
            Dim selStruct As Struct = selRow.Struct
            Dim intRowCount As Integer = selRow.RowCountStruct
            If ClippedRow IsNot Nothing Then intRowCount += 1

            If intRowCount <= 1 Then
                selRow.StructHeight = selRow.Struct.CellImage.Height
                If selRow.RowHeight < selRow.StructHeight Then selRow.RowHeight = selRow.StructHeight
            Else
                Dim intStructHeight As Integer = selRow.Struct.CellImage.Height
                Dim intTop As Integer = 0
                For i = 0 To intRowCount - 1
                    If intRowIndex + i > Me.PrintList.Count - 1 Then Exit For

                    Dim objPrintListRow As PrintLogframeRow = PrintList(intRowIndex + i)
                    Dim intAvailableHeight As Integer = objPrintListRow.RowHeight

                    If i = intRowCount - 1 Then
                        Dim intNeededHeight As Integer = intStructHeight - intTop
                        If intNeededHeight > intAvailableHeight Then objPrintListRow.RowHeight = intNeededHeight
                    End If

                    If intAvailableHeight > 0 Then
                        objPrintListRow.StructHeight = intAvailableHeight
                        intTop += objPrintListRow.RowHeight
                    Else
                        objPrintListRow.StructHeight = 0
                    End If
                Next
            End If
        End If
    End Sub
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
            End If
        Next

        Return NrRows
    End Function

    Private Function GetRowCountOfStruct_Resources(ByVal objActivity As Activity) As Integer
        Dim NmbrRows As Integer

        If ShowIndicatorColumn = True Then
            If objActivity IsNot Nothing Then NmbrRows = objActivity.Resources.Count
        End If
        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function

    Private Function GetRowCountOfStruct_Assumptions(ByVal objStruct As Struct) As Integer
        Dim NmbrRows As Integer

        If ShowAssumptionColumn = True Then
            If objStruct IsNot Nothing Then NmbrRows = objStruct.Assumptions.Count
        End If
        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function

    Private Function GetRowCountOfIndicator(ByVal selIndicator As Indicator) As Integer
        Dim NmbrRows As Integer

        If ShowVerificationSourceColumn = True Then
            If selIndicator IsNot Nothing AndAlso selIndicator.VerificationSources.Count > 0 Then
                NmbrRows = selIndicator.VerificationSources.Count
            End If
        End If
        If NmbrRows = 0 Then NmbrRows = 1

        Return NmbrRows
    End Function
#End Region

#Region "General methods"
    Public Function IsResourceBudgetRow(ByVal selRow As PrintLogframeRow) As Boolean
        If selRow.Section = Logframe.SectionTypes.ActivitiesSection And Me.ShowResourcesBudget = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal

        For Each selRow As PrintLogframeRow In PrintList
            intTotalHeight += selRow.RowHeight
        Next

        decPages = intTotalHeight / Me.ContentHeight
        Return Math.Ceiling(decPages)
    End Function
#End Region

#Region "Print page"
    Protected Overrides Sub OnBeginPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnBeginPrint(e)

        boolColumnsWidthSet = False
        LastRowY = ContentTop
        objClippedRow = Nothing
        strClippedTextTop = String.Empty
        strClippedTextBottom = String.Empty

        LoadSections()
    End Sub

    Protected Overrides Sub OnQueryPageSettings(ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs)
        MyBase.OnQueryPageSettings(e)
    End Sub

    Protected Overrides Sub OnEndPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        LastRowY = ContentTop
        objClippedRow = Nothing
        strClippedTextTop = String.Empty
        strClippedTextBottom = String.Empty

        MyBase.OnEndPrint(e)
    End Sub

    Protected Overrides Sub OnPrintPage(ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        MyBase.OnPrintPage(e)
        PageGraph = e.Graphics

        If boolColumnsWidthSet = False Then
            SetColumnsWidth()
            ReloadImages()
            Me.TotalPages = GetTotalPages()
            boolColumnsWidthSet = True
        End If

        Dim selRow As PrintLogframeRow
        Dim intMinHeight As Integer

        'Print Header
        PrintHeader()

        Do While RowIndex <= PrintList.Count - 1
            selRow = PrintList(RowIndex)

            If LastRowY + selRow.RowHeight < Me.ContentBottom Then
                PrintPage_PrintRow(selRow)

                If selRow.ClippedRow = True Then
                    ClippedRow = Nothing
                End If
            Else
                'set the minimum height to the line spacing of the first line of the row
                intMinHeight = PrintPage_GetMinHeight(selRow)
                intMinHeight += LastRowY

                If intMinHeight < Me.ContentBottom Then
                    PrintPage_PrintRow(selRow)

                    If ClippedRow IsNot Nothing Then
                        PrintList.Insert(RowIndex, ClippedRow)
                        ReloadImages_Normal(ClippedRow)
                        SetRowHeight(RowIndex)
                    End If
                    PrintFooter()

                    LastRowY = ContentTop
                    PageNumber += 1
                    e.HasMorePages = True
                    Exit Do
                Else
                    'else go to a new page and print the line there
                    PrintFooter()

                    LastRowY = ContentTop
                    PageNumber += 1
                    e.HasMorePages = True
                    Exit Do
                End If
                End If

            If RowIndex > PrintList.Count - 1 Then
                PrintFooter()
                PageGraph.DrawLine(penBlack2, LeftMargin, LastRowY, LeftMargin + ContentWidth, LastRowY)
                e.HasMorePages = False
            End If
        Loop
        If PrintList.Count = 0 Then RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(0, 0))
    End Sub

    Private Sub PrintPage_PrintRow(ByVal selRow As PrintLogframeRow)
        'if there is enough space to print the line, do so...
        If selRow.RowType = PrintLogframeRow.RowTypes.Section Then
            PrintPage_PrintSections(selRow)
        ElseIf selRow.RowType = PrintLogframeRow.RowTypes.RepeatPurpose Or selRow.RowType = PrintLogframeRow.RowTypes.RepeatOutput Then
            PrintPage_PrintRepeats(selRow)
        Else
            PrintPage_PrintStruct(selRow)
            If ShowIndicatorColumn = True Then
                If IsResourceBudgetRow(selRow) = False Then
                    PrintPage_PrintIndicator(selRow)
                Else
                    PrintPage_PrintResource(selRow)
                End If
            End If
            If ShowVerificationSourceColumn = True Then
                If IsResourceBudgetRow(selRow) = False Then
                    PrintPage_PrintVerificationSource(selRow)
                Else
                    PrintPage_PrintResourceBudget(selRow)
                End If
            End If
            If ShowAssumptionColumn = True Then PrintPage_PrintAssumption(selRow)
        End If

        RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(RowIndex, PrintList.Count))
        RowIndex += 1
        LastRowY += selRow.RowHeight
    End Sub

    Private Sub PrintPage_PrintSortNumber(ByVal strValue As String, ByVal rCell As Rectangle)
        Dim formatCells As New StringFormat()
        Dim brText As SolidBrush = New SolidBrush(Color.Black)

        If PageGraph IsNot Nothing Then
            rCell.X += CONST_HorizontalPadding
            rCell.Y += CONST_VerticalPadding
            formatCells.Alignment = StringAlignment.Near
            formatCells.LineAlignment = StringAlignment.Near
            PageGraph.DrawString(strValue, CurrentLogFrame.SortNumberFont, brText, rCell, formatCells)
        End If
    End Sub

    Private Sub PrintPage_PrintSections(ByVal selRow As PrintLogframeRow)
        Dim objStringFormat As New StringFormat()
        Dim brFont As Brush = New SolidBrush(Logframe.SectionTitleFontColor)

        objStringFormat = StringFormat.GenericDefault
        objStringFormat.Alignment = StringAlignment.Center
        objStringFormat.LineAlignment = StringAlignment.Center

        Dim rStruct As New Rectangle(LeftMargin, LastRowY, StructSortWidth + StructRTFWidth, selRow.RowHeight)
        Dim rIndicator As New Rectangle(rStruct.Right, LastRowY, IndicatorSortWidth + IndicatorRTFWidth, selRow.RowHeight)
        Dim rVerificationSource As New Rectangle(rIndicator.Right, LastRowY, VerificationSourceSortWidth + VerificationSourceRTFWidth, selRow.RowHeight)
        Dim rAssumption As New Rectangle(rVerificationSource.Right, LastRowY, AssumptionSortWidth + AssumptionRTFWidth, selRow.RowHeight)

        Dim rRow As New Rectangle(LeftMargin, LastRowY, TotalWidth, selRow.RowHeight)
        Dim brSection As LinearGradientBrush = New LinearGradientBrush(rRow, Color.LightGray, Color.Gray, LinearGradientMode.Vertical)

        If PageGraph IsNot Nothing Then
            PageGraph.FillRectangle(brSection, rRow)

            PageGraph.DrawString(selRow.Struct.Text, Logframe.SectionTitleFont, brFont, rStruct, objStringFormat)
            If ShowIndicatorColumn = True Then
                If IsResourceBudgetRow(selRow) = False Then
                    PageGraph.DrawString(selRow.Indicator.Text, Logframe.SectionTitleFont, brFont, rIndicator, objStringFormat)
                Else
                    PageGraph.DrawString(selRow.Resource.Text, Logframe.SectionTitleFont, brFont, rIndicator, objStringFormat)
                End If
            End If
            If ShowVerificationSourceColumn = True Then
                If IsResourceBudgetRow(selRow) = False Then
                    PageGraph.DrawString(selRow.VerificationSource.Text, Logframe.SectionTitleFont, brFont, rVerificationSource, objStringFormat)
                Else
                    PageGraph.DrawString(Resource.BudgetName, Logframe.SectionTitleFont, brFont, rVerificationSource, objStringFormat)
                End If
            End If
            If ShowAssumptionColumn = True Then PageGraph.DrawString(selRow.Assumption.Text, Logframe.SectionTitleFont, brFont, rAssumption, objStringFormat)
        End If

        PrintPage_PrintLines(rStruct, True)
        If ShowIndicatorColumn = True Then PrintPage_PrintLines(rIndicator, False)
        If ShowVerificationSourceColumn = True Then PrintPage_PrintLines(rVerificationSource, False)
        If ShowAssumptionColumn = True Then PrintPage_PrintLines(rAssumption, False)

        PageGraph.DrawLine(penBlack2, rRow.Left, rRow.Top, rRow.Right, rRow.Top)
    End Sub

    Private Sub PrintPage_PrintRepeats(ByVal selRow As PrintLogframeRow)
        Dim rCell As Rectangle
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim brBackGround As Brush
        'Dim fntRepeat As Font = New Font(Me.Logframe.SortNumberFont, FontStyle.Bold)
        Dim brRepeat As SolidBrush = New SolidBrush(Color.Black)
        Dim objStringFormat As New StringFormat()
        Dim rRow As New Rectangle(LeftMargin, LastRowY, TotalWidth, selRow.RowHeight)

        objStringFormat.Alignment = StringAlignment.Near
        objStringFormat.LineAlignment = StringAlignment.Near

        If selRow.RowType = PrintLogframeRow.RowTypes.RepeatPurpose Then brBackGround = Brushes.CornflowerBlue Else brBackGround = Brushes.LightSteelBlue

        If PageGraph IsNot Nothing Then
            PageGraph.FillRectangle(brBackGround, rRow)

            'draw sort number
            intX = LeftMargin
            intY = LastRowY
            rCell = New Rectangle(intX, intY, StructSortWidth, selRow.RowHeight)
            PrintPage_PrintSortNumber(selRow.StructSort, rCell)

            'draw text
            intX = Me.LeftMargin + Me.StructSortWidth
            intColWidth = TotalWidth - Me.StructSortWidth
            rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

            PageGraph.DrawString(selRow.Struct.Text, Logframe.SectionTitleFont, brRepeat, rCell, objStringFormat)
            PageGraph.DrawLine(penBlack2, rRow.Left, rRow.Top, rRow.Left, rRow.Bottom)
            PageGraph.DrawLine(penBlack2, rRow.Right, rRow.Top, rRow.Right, rRow.Bottom)
            PageGraph.DrawLine(penBlack1, rRow.Left, rRow.Top, rRow.Right, rRow.Top)
        End If
    End Sub

    Private Sub PrintPage_PrintLines(ByVal rCell As Rectangle, ByVal boolIsStruct As Boolean)
        If PageGraph IsNot Nothing Then
            'lines between columns
            If rCell.Bottom > Me.ContentBottom Then
                If boolIsStruct = True Then _
                    PageGraph.DrawLine(penBlack2, Me.LeftMargin, rCell.Top, Me.LeftMargin, Me.ContentBottom)
                PageGraph.DrawLine(penBlack2, rCell.Right, rCell.Top, rCell.Right, Me.ContentBottom)
            Else
                If boolIsStruct = True Then _
                    PageGraph.DrawLine(penBlack2, Me.LeftMargin, rCell.Top, Me.LeftMargin, rCell.Bottom + 1)
                PageGraph.DrawLine(penBlack2, rCell.Right, rCell.Top, rCell.Right, rCell.Bottom)
            End If
        End If
    End Sub

    Private Function PrintPage_GetMinHeight(ByVal selRow As PrintLogframeRow)
        Dim intMinHeight As Integer
        Dim boolAssumption As Boolean

        With RichTextManager
            If String.IsNullOrEmpty(selRow.StructRtf) = False Then
                intMinHeight = .GetFirstLineSpacing(StructRTFWidth, selRow.StructRtf)
            Else
                If IsResourceBudgetRow(selRow) = False Then
                    If String.IsNullOrEmpty(selRow.IndicatorRtf) = False Then
                        intMinHeight = .GetFirstLineSpacing(IndicatorRTFWidth, selRow.IndicatorRtf)
                    ElseIf String.IsNullOrEmpty(selRow.VerificationSourceRtf) = False Then
                        intMinHeight = .GetFirstLineSpacing(VerificationSourceRTFWidth, selRow.VerificationSourceRtf)
                    Else
                        boolAssumption = True
                    End If
                Else
                    If String.IsNullOrEmpty(selRow.ResourceRtf) = False Then
                        .Rtf = selRow.ResourceRtf
                        .Width = IndicatorRTFWidth
                        intMinHeight = .GetFirstLineSpacing(IndicatorRTFWidth, selRow.ResourceRtf)
                    Else
                        boolAssumption = True
                    End If
                End If
            End If

            If boolAssumption = True And String.IsNullOrEmpty(selRow.AssumptionRtf) = False Then
                intMinHeight = .GetFirstLineSpacing(AssumptionRTFWidth, selRow.AssumptionRtf)
            Else
                intMinHeight = NewCellHeight()
            End If
        End With
        Return intMinHeight
    End Function
#End Region

#Region "Clipped text"
    Private Sub PrintClippedText(ByVal strRtf As String, ByVal rImage As Rectangle, intColumnWidth As Integer)
        With RichTextManager
            Dim bmClip As Bitmap = .PrintClippedRichText(intColumnWidth, strRtf, ContentBottom - rImage.Y)
            strClippedTextTop = .ClippedTextTop
            strClippedTextBottom = .ClippedTextBottom

            rImage.X += 2
            rImage.Y += 2
            rImage.Height = bmClip.Height
            PageGraph.DrawImage(bmClip, rImage)
        End With
    End Sub
#End Region

#Region "Print struct"
    Private Sub PrintPage_PrintStruct(ByVal selRow As PrintLogframeRow)
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim rCell, rImage As Rectangle

        'StructSort
        intX = LeftMargin
        intY = LastRowY

        rCell = New Rectangle(intX, intY, StructSortWidth, selRow.RowHeight)

        PrintPage_PrintSortNumber(selRow.StructSort, rCell)

        'StructRTF
        intX = Me.LeftMargin + Me.StructSortWidth
        intColWidth = Me.StructRTFWidth

        rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

        If PageGraph IsNot Nothing Then
            If selRow.Struct IsNot Nothing AndAlso selRow.Struct.CellImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selRow.Struct.CellImage.Width, selRow.Struct.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Struct.CellImage, rImage)
                Else
                    PrintClippedText(selRow.StructRtf, rImage, StructRTFWidth)
                    selRow.StructRtf = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintLogframeRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    Select Case selRow.Section
                        Case Logframe.SectionTypes.GoalsSection
                            ClippedRow.Struct = New Goal(strClippedTextBottom)
                        Case Logframe.SectionTypes.PurposesSection
                            ClippedRow.Struct = New Purpose(strClippedTextBottom)
                        Case Logframe.SectionTypes.OutputsSection
                            ClippedRow.Struct = New Output(strClippedTextBottom)
                        Case Logframe.SectionTypes.ActivitiesSection
                            ClippedRow.Struct = New Activity(strClippedTextBottom)
                            ClippedRow.StructIndent = selRow.StructIndent
                    End Select
                End If
            End If
            If String.IsNullOrEmpty(selRow.StructRtf) = False Then _
                PageGraph.DrawLine(penBlack1, Me.ContentLeft, intY, Me.ContentRight, intY)
        End If

        PrintPage_PrintLines(rCell, True)
    End Sub
#End Region

#Region "Print indicators"
    Private Sub PrintPage_PrintIndicator(ByVal selRow As PrintLogframeRow)
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim rCell, rImage As Rectangle

        'IndicatorSort
        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth
        intY = LastRowY
        rCell = New Rectangle(intX, intY, IndicatorSortWidth, selRow.RowHeight)

        PrintPage_PrintSortNumber(selRow.IndicatorSort, rCell)

        'IndRTF
        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth
        intX += Me.IndicatorSortWidth
        intColWidth = Me.IndicatorRTFWidth

        rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

        If PageGraph IsNot Nothing Then
            If selRow.Indicator IsNot Nothing AndAlso selRow.Indicator.CellImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selRow.Indicator.CellImage.Width, selRow.Indicator.CellImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.Indicator.CellImage, rImage)
                Else
                    PrintClippedText(selRow.IndicatorRtf, rImage, IndicatorRTFWidth)
                    selRow.IndicatorRtf = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintLogframeRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.Indicator = New Indicator(strClippedTextBottom)
                End If
            End If
            If String.IsNullOrEmpty(selRow.IndicatorRtf) = False And selRow.ClippedRow = False Then _
                PageGraph.DrawLine(penBlack1, intX - IndicatorSortWidth, intY, intX + IndicatorRTFWidth + VerificationSourceSortWidth + VerificationSourceRTFWidth, intY)
        End If

        PrintPage_PrintLines(rCell, False)
    End Sub
#End Region

#Region "Print verification sources"
    Private Sub PrintPage_PrintVerificationSource(ByVal selRow As PrintLogframeRow)
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim rCell, rImage As Rectangle

        'VerificationSourceSort
        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth + _
                Me.IndicatorSortWidth + Me.IndicatorRTFWidth
        intY = LastRowY
        If IsResourceBudgetRow(selRow) = False Then
            rCell = New Rectangle(intX, intY, VerificationSourceSortWidth, selRow.RowHeight)
            PrintPage_PrintSortNumber(selRow.VerificationSourceSort, rCell)
        End If

        'VerRTF
        intX += Me.VerificationSourceSortWidth
        intColWidth = Me.VerificationSourceRTFWidth

        rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

        If PageGraph IsNot Nothing Then
            If selRow.VerificationSourceImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selRow.VerificationSourceImage.Width, selRow.VerificationSourceImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.VerificationSourceImage, rImage)
                Else
                    PrintClippedText(selRow.VerificationSourceRtf, rImage, VerificationSourceRTFWidth)
                    selRow.VerificationSourceRtf = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintLogframeRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    ClippedRow.VerificationSource = New VerificationSource(strClippedTextBottom)
                End If
            End If

            If String.IsNullOrEmpty(selRow.VerificationSourceRtf) = False And selRow.ClippedRow = False Then _
                PageGraph.DrawLine(penBlack1, intX - VerificationSourceSortWidth, intY, intX + VerificationSourceRTFWidth, intY)
        End If

        PrintPage_PrintLines(rCell, False)
    End Sub
#End Region

#Region "Print Resources"
    Private Sub PrintPage_PrintResource(ByVal selRow As PrintLogframeRow)
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim rCell, rImage As Rectangle

        'ResourceSort
        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth
        intY = LastRowY
        rCell = New Rectangle(intX, intY, IndicatorSortWidth, selRow.RowHeight)

        PrintPage_PrintSortNumber(selRow.ResourceSort, rCell)

        'IndRTF
        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth
        intX += Me.IndicatorSortWidth
        intColWidth = Me.IndicatorRTFWidth

        rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

        If PageGraph IsNot Nothing Then
            If selRow.ResourceImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selRow.ResourceImage.Width, selRow.ResourceImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.ResourceImage, rImage)
                Else
                    PrintClippedText(selRow.ResourceRtf, rImage, IndicatorRTFWidth)
                    selRow.ResourceRtf = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintLogframeRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    objClippedRow.Resource = New Resource(strClippedTextBottom)
                End If
            End If
            If String.IsNullOrEmpty(selRow.ResourceRtf) = False And selRow.ClippedRow = False Then _
                PageGraph.DrawLine(penBlack1, intX - IndicatorSortWidth, intY, intX + IndicatorRTFWidth + VerificationSourceSortWidth + VerificationSourceRTFWidth, intY)
        End If

        PrintPage_PrintLines(rCell, False)
    End Sub
#End Region

#Region "Print resource amount"
    Private Sub PrintPage_PrintResourceBudget(ByVal selRow As PrintLogframeRow)
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim formatCells As New StringFormat()

        Dim rCell, rImage As Rectangle

        formatCells = StringFormat.GenericTypographic
        formatCells.Alignment = StringAlignment.Far

        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth + _
                Me.IndicatorSortWidth + Me.IndicatorRTFWidth
        intY = LastRowY
        intColWidth = Me.VerificationSourceSortWidth + Me.VerificationSourceRTFWidth

        rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

        If selRow.ClippedRow = True And selRow.TotalCostAmount = 0 Then
            PrintPage_PrintLines(rCell, False)
            Exit Sub
        End If

        If selRow.Resource IsNot Nothing Then
            Dim brText As New SolidBrush(selRow.Resource.GetFontColor)
            Dim selFont As Font = selRow.Resource.GetFont
            Dim strBudget As String = String.Format("{0} {1}", selRow.TotalCostAmount.ToString("N2"), CurrentLogFrame.CurrencyCode)

            rImage = rCell
            rImage.Y += 2
            rImage.Width -= 4
            If PageGraph IsNot Nothing Then
                PageGraph.DrawString(strBudget, selFont, brText, rImage, formatCells)
                PageGraph.DrawLine(penBlack1, intX - VerificationSourceSortWidth, intY, intX + VerificationSourceRTFWidth, intY)
            End If
        End If

        PrintPage_PrintLines(rCell, False)
    End Sub
#End Region

#Region "Print assumptions"
    Private Sub PrintPage_PrintAssumption(ByVal selRow As PrintLogframeRow)
        Dim intX, intY As Integer
        Dim intColWidth As Integer
        Dim rCell, rImage As Rectangle

        'AssumptionSort
        intX = Me.LeftMargin + Me.StructSortWidth + Me.StructRTFWidth + _
                Me.IndicatorSortWidth + Me.IndicatorRTFWidth + Me.VerificationSourceSortWidth + Me.VerificationSourceRTFWidth
        intY = LastRowY
        rCell = New Rectangle(intX, intY, AssumptionSortWidth, selRow.RowHeight)
        PrintPage_PrintSortNumber(selRow.AssumptionSort, rCell)

        'AsmRTF
        intX += Me.AssumptionSortWidth
        intColWidth = Me.AssumptionRTFWidth

        rCell = New Rectangle(intX, intY, intColWidth, selRow.RowHeight)

        If PageGraph IsNot Nothing Then
            If selRow.AssumptionImage IsNot Nothing Then
                rImage = New Rectangle(rCell.X, rCell.Y, selRow.AssumptionImage.Width, selRow.AssumptionImage.Height)

                If rImage.Bottom <= ContentBottom Then
                    PageGraph.DrawImage(selRow.AssumptionImage, rImage)
                Else
                    PrintClippedText(selRow.AssumptionRtf, rImage, AssumptionRTFWidth)
                    selRow.AssumptionRtf = strClippedTextTop

                    If ClippedRow Is Nothing Then
                        ClippedRow = New PrintLogframeRow(selRow.Section)
                        ClippedRow.ClippedRow = True
                    End If
                    objClippedRow.Assumption = New Assumption(strClippedTextBottom)
                End If
            End If
            If String.IsNullOrEmpty(selRow.AssumptionRtf) = False And selRow.ClippedRow = False Then _
                PageGraph.DrawLine(penBlack1, intX - AssumptionSortWidth, intY, intX + AssumptionRTFWidth, intY)
        End If

        PrintPage_PrintLines(rCell, False)
    End Sub
#End Region 'Print assumptions
End Class



