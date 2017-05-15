Public Class DetailIndicator
    Friend WithEvents lvSubIndicators As ListViewSubIndicatorsBase
    Friend WithEvents IndicatorBindingSource As New BindingSource

    Public Event CurrentDataGridViewChanged(ByVal sender As Object, ByVal e As CurrentDataGridViewChangedEventArgs)

    Private indCurrentIndicator As Indicator
    Private ctlCurrentStatementDataGridView As DataGridViewBaseClass
    Private ctlCurrentTargetsDataGridView As DataGridView
    Private objHiddenTab As TabPage
    Private strHiddenPageName As String

    Private boolTextSelected As Boolean
    Private intCurrentQuestionType As Integer

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()

        Me.CurrentIndicator = indicator
        Me.Dock = DockStyle.Fill
        LoadItems()

        TabControlIndicator.SelectedIndex = CurrentIndicatorsTab
    End Sub

#Region "Properties"
    Public Property CurrentIndicator() As Indicator
        Get
            Return indCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            indCurrentIndicator = value
        End Set
    End Property

    Public Property TextSelected() As Boolean
        Get
            Return boolTextSelected
        End Get
        Set(ByVal value As Boolean)
            boolTextSelected = value
            Me.Refresh()
        End Set
    End Property

    Private Property CurrentQuestionType As Integer
        Get
            Return intCurrentQuestionType
        End Get
        Set(ByVal value As Integer)
            intCurrentQuestionType = value
        End Set
    End Property

    Public Property CurrentStatementsDataGridView As DataGridViewBaseClass
        Get
            Return ctlCurrentStatementDataGridView
        End Get
        Set(ByVal value As DataGridViewBaseClass)
            ctlCurrentStatementDataGridView = value
        End Set
    End Property

    Public Property CurrentTargetsDataGridView As DataGridView
        Get
            Return ctlCurrentTargetsDataGridView
        End Get
        Set(ByVal value As DataGridView)
            ctlCurrentTargetsDataGridView = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub LoadItems()
        SuspendLayout()

        chkAdvanced.Checked = My.Settings.setShowAdvancedIndicatorOptions

        If CurrentIndicator IsNot Nothing Then
            CurrentQuestionType = CurrentIndicator.QuestionType

            LoadItems_VerifyTargets()
            strHiddenPageName = TabPageTargets.Name

            Dim ListTargetGroups As Dictionary(Of Guid, String) = CurrentLogFrame.GetTargetGroupsList
            IndicatorBindingSource.DataSource = CurrentIndicator

            tbIndicator.DataBindings.Add("Text", IndicatorBindingSource, "Text")

            Dim QuestionTypes As List(Of StructuredComboBoxItem) = LoadQuestionTypesComboList()
            With cmbQuestionType
                .DataSource = QuestionTypes
                .ValueMember = "Type"
                .DisplayMember = "Description"
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataBindings.Add("SelectedValue", IndicatorBindingSource, "QuestionType")
            End With

            With cmbRegistration
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_RegistrationOptions.ToArray)
                .DataBindings.Add("SelectedIndex", IndicatorBindingSource, "Registration")
            End With

            With cmbTargetGroupGuid
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .DataSource = New BindingSource(ListTargetGroups, Nothing)
                .ValueMember = "Key"
                .DisplayMember = "Value"
                .DataBindings.Add("SelectedValue", IndicatorBindingSource, "TargetGroupGuid")
            End With

            With cmbAggregateHorizontal
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_AggregationOptions.ToArray)
                .DataBindings.Add("SelectedIndex", IndicatorBindingSource, "AggregateHorizontal")
            End With

            With cmbAggregateVertical
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList
                .Items.AddRange(LIST_AggregationOptions.ToArray)
                .Items.RemoveAt(.Items.Count - 1)
                .Items.RemoveAt(.Items.Count - 1)
                .DataBindings.Add("SelectedIndex", IndicatorBindingSource, "AggregateVertical")
            End With

            ntbWeightingFactorChildren.DataBindings.Add("Text", IndicatorBindingSource, "WeightingFactorChildren")
            gbWeightingChildren.Visible = My.Settings.setShowAdvancedIndicatorOptions

            'set sub forms
            ChangeUserControls(CurrentIndicator.QuestionType)

        End If
        ResumeLayout()

    End Sub

    Private Sub LoadItems_VerifyTargets()
        Dim intSection As Integer = CurrentIndicator.Section

        If intSection = 0 Then
            intSection = CurrentProjectForm.dgvLogframe.CurrentSection
            CurrentIndicator.Section = intSection
        End If

        Dim objTargetDeadlines As TargetDeadlines = CurrentLogFrame.GetTargetDeadlines(intSection)
        If objTargetDeadlines IsNot Nothing Then
            objTargetDeadlines.ConformWithDeadlines(CurrentIndicator.Targets)
            objTargetDeadlines.ConformWithDeadlines(CurrentIndicator.PopulationTargets)
        End If
    End Sub
#End Region

#Region "Events"
    Private Sub chkAdvanced_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAdvanced.CheckedChanged
        If chkAdvanced.Checked <> My.Settings.setShowAdvancedIndicatorOptions Then
            My.Settings.setShowAdvancedIndicatorOptions = chkAdvanced.Checked
            ChangeUserControls(CurrentIndicator.QuestionType)

            gbWeightingChildren.Visible = My.Settings.setShowAdvancedIndicatorOptions
        End If
    End Sub

    Private Sub cmbQuestionType_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbQuestionType.SelectedValueChanged
        If cmbQuestionType.DataBindings.Count < 1 Then Exit Sub

        If cmbQuestionType.SelectedItem Is Nothing Then Exit Sub

        Dim intQuestionType As Integer = CType(cmbQuestionType.SelectedItem, StructuredComboBoxItem).Type

        With CurrentIndicator
            If intQuestionType <> CurrentQuestionType Then
                If intQuestionType = CONST_QuestionTypeTitle Then
                    Dim intIndex As Integer = cmbQuestionType.SelectedIndex + 1
                    Dim NextItem As StructuredComboBoxItem = TryCast(cmbQuestionType.Items(intIndex), StructuredComboBoxItem)

                    If NextItem IsNot Nothing Then
                        intQuestionType = NextItem.Type
                    Else
                        intQuestionType = 0
                    End If
                    CurrentIndicator.QuestionType = intQuestionType
                    cmbQuestionType.DataBindings(0).ReadValue()
                Else
                    cmbQuestionType.DataBindings(0).WriteValue()
                End If

                .SetQuestionType_Settings()
                .SetQuestionType_Statements()
                .SetQuestionType_ResponseClasses()
                .SetQuestionType_ResetTargets()

                'if indicator has sub-indicators, set them to the same type
                If .Indicators.Count > 0 Then
                    .Indicators.UpdateQuestionTypeOfChildren(intQuestionType)
                End If

                ChangeUserControls(intQuestionType)
                CurrentQuestionType = intQuestionType
            End If
        End With
    End Sub

    Private Sub cmbRegistration_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbRegistration.SelectedIndexChanged
        cmbRegistration.DataBindings(0).WriteValue()

        If CurrentIndicator.Registration <> Indicator.RegistrationOptions.ProgrammeLevel Then
            Select Case CurrentQuestionType
                Case Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.PercentageValue
                    CurrentIndicator.AggregateHorizontal = Indicator.AggregationOptions.Average
            End Select
        End If

        If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
            Dim intSelRegistration As Integer = cmbRegistration.SelectedIndex
            Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(CurrentIndicator.ParentIndicatorGuid)

            If intSelRegistration < ParentIndicator.Registration Then
                cmbRegistration.SelectedIndex = ParentIndicator.Registration
            End If
        End If

        AggregateHorizontalVisibility()
    End Sub

    Private Sub AggregateHorizontalVisibility()
        If CurrentIndicator.Registration = Indicator.RegistrationOptions.ProgrammeLevel Then
            lblAggregateHorizontal.Visible = False
            cmbAggregateHorizontal.Visible = False
        Else
            Select Case CurrentQuestionType
                Case Indicator.QuestionTypes.Image, Indicator.QuestionTypes.OpenEnded
                    lblAggregateHorizontal.Visible = False
                    cmbAggregateHorizontal.Visible = False
                Case Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.PercentageValue
                    lblAggregateHorizontal.Visible = True
                    cmbAggregateHorizontal.Visible = True
                    lblAggregateHorizontal.Enabled = False
                    cmbAggregateHorizontal.Enabled = False
                Case Else
                    lblAggregateHorizontal.Visible = True
                    cmbAggregateHorizontal.Visible = True
            End Select

        End If
    End Sub

    Private Sub cmbTargetGroupGuid_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTargetGroupGuid.SelectedValueChanged
        If CurrentIndicator.ParentIndicatorGuid <> Guid.Empty Then
            Dim SelTargetGroupGuid As Guid
            Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(CurrentIndicator.ParentIndicatorGuid)
            SelTargetGroupGuid = CType(cmbTargetGroupGuid.SelectedItem, KeyValuePair(Of Guid, String)).Key

            If ParentIndicator IsNot Nothing AndAlso ParentIndicator.TargetGroupGuid <> Guid.Empty Then
                If SelTargetGroupGuid <> ParentIndicator.TargetGroupGuid Then
                    CurrentIndicator.TargetGroupGuid = ParentIndicator.TargetGroupGuid

                    If cmbTargetGroupGuid.DataBindings.Count > 0 Then cmbTargetGroupGuid.DataBindings(0).ReadValue()
                End If
            End If
        End If
    End Sub

    Private Sub cmbAggregateVertical_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbAggregateVertical.SelectedIndexChanged
        cmbAggregateVertical.DataBindings(0).WriteValue()
        lvSubIndicators.LoadItems()
    End Sub

    Private Sub ntbWeightingFactorChildren_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbWeightingFactorChildren.Validated
        lvSubIndicators.LoadItems()
    End Sub

    Private Sub TabControlIndicator_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControlIndicator.SelectedIndexChanged
        If TabControlIndicator.SelectedTab Is TabPageTargets Then
            ReloadTargets()
        End If
        CurrentIndicatorsTab = TabControlIndicator.SelectedIndex
    End Sub
#End Region

#Region "Add custom controls"
    Public Sub ChangeUserControls(ByVal intQuestionType As Integer)
        'Change detail windows of tab pages
        PanelQuestionType.Controls.Clear()
        ReloadTargets()

        Select Case intQuestionType
            Case Indicator.QuestionTypes.OpenEnded
                If CurrentIndicator.Indicators.Count = 0 Then
                    Dim ctlOpenEnded As New IndicatorOpenEnded(CurrentIndicator)

                    ctlOpenEnded.Dock = DockStyle.Fill
                    PanelQuestionType.Controls.Add(ctlOpenEnded)
                End If

            Case Indicator.QuestionTypes.MaxDiff
                If CurrentIndicator.Indicators.Count = 0 Then
                    Dim ctlMaxDiff As New IndicatorMaxDiff(CurrentIndicator)
                    CurrentStatementsDataGridView = ctlMaxDiff.dgvStatementsMaxDiff

                    AddHandler ctlMaxDiff.CurrentDataGridViewChanged, AddressOf CurrentDataGridViewNeedsUpdate

                    ctlMaxDiff.Dock = DockStyle.Fill
                    PanelQuestionType.Controls.Add(ctlMaxDiff)
                End If
            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                Dim ctlValues As New IndicatorValues(CurrentIndicator)

                AddHandler ctlValues.cmbUnit.Validated, AddressOf UpdateUnitsChildIndicators
                AddHandler ctlValues.rbtnScoringPercentage.CheckedChanged, AddressOf UpdateUnitsChildIndicators
                AddHandler ctlValues.TargetSystemUpdated, AddressOf TargetSystemUpdated
                AddHandler ctlValues.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlValues.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlValues)

                If intQuestionType = Indicator.QuestionTypes.PercentageValue And CurrentIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                    cmbAggregateVertical.Enabled = False
                End If

            Case Indicator.QuestionTypes.Ratio
                Dim ctlRatio As New IndicatorRatio(CurrentIndicator)

                AddHandler ctlRatio.rbtnScoringPercentage.CheckedChanged, AddressOf UpdateUnitsChildIndicators
                AddHandler ctlRatio.TargetSystemUpdated, AddressOf TargetSystemUpdated
                AddHandler ctlRatio.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlRatio.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlRatio)

            Case Indicator.QuestionTypes.Formula
                Dim ctlFormula As New IndicatorFormula(CurrentIndicator)
                CurrentStatementsDataGridView = ctlFormula.dgvStatementsFormula

                AddHandler ctlFormula.CurrentDataGridViewChanged, AddressOf CurrentDataGridViewNeedsUpdate
                AddHandler ctlFormula.ScoreSystemUpdated, AddressOf ScoreSystemUpdated
                AddHandler ctlFormula.TargetSystemUpdated, AddressOf TargetSystemUpdated
                AddHandler ctlFormula.StatementsUpdated, AddressOf StatementsUpdated

                ctlFormula.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlFormula)

            Case Indicator.QuestionTypes.YesNo
                Dim ctlYesNo As New IndicatorYesNo(CurrentIndicator)

                AddHandler ctlYesNo.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlYesNo.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlYesNo)

            Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, _
                Indicator.QuestionTypes.LikertTypeScale, Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.SemanticDiff
                Dim ctlMultipleOptions As New IndicatorMultipleOption(CurrentIndicator)

                AddHandler ctlMultipleOptions.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlMultipleOptions.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlMultipleOptions)
                CurrentStatementsDataGridView = ctlMultipleOptions.dgvResponseClasses
            Case Indicator.QuestionTypes.Ranking
                Dim ctlRanking As New IndicatorRanking(CurrentIndicator)

                AddHandler ctlRanking.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlRanking.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlRanking)
                CurrentStatementsDataGridView = ctlRanking.dgvResponseClasses
            Case Indicator.QuestionTypes.FrequencyLikert
                Dim ctlFrequencyLikert As New IndicatorFrequencyLikert(CurrentIndicator)

                AddHandler ctlFrequencyLikert.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlFrequencyLikert.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlFrequencyLikert)
                CurrentStatementsDataGridView = ctlFrequencyLikert.dgvResponseClasses
            Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                Dim ctlScales As New IndicatorScales(CurrentIndicator)
                CurrentStatementsDataGridView = ctlScales.dgvStatementsScales

                AddHandler ctlScales.CurrentDataGridViewChanged, AddressOf CurrentDataGridViewNeedsUpdate
                AddHandler ctlScales.ScoreSystemUpdated, AddressOf ScoreSystemUpdated
                AddHandler ctlScales.tbAgreeText.Validated, AddressOf ReloadTargets
                AddHandler ctlScales.tbDisagreeText.Validated, AddressOf ReloadTargets

                ctlScales.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlScales)

            Case Indicator.QuestionTypes.MixedSubIndicators
                Dim ctlMixedSubIndicators As New IndicatorMixedSubIndicators(CurrentIndicator)

                AddHandler ctlMixedSubIndicators.ScoreSystemUpdated, AddressOf ScoreSystemUpdated

                ctlMixedSubIndicators.Dock = DockStyle.Fill
                PanelQuestionType.Controls.Add(ctlMixedSubIndicators)
        End Select

        ChangeUserControls_SubIndicators()
    End Sub

    Private Sub ChangeUserControls_SubIndicators()
        Dim CurrentTargetDeadlinesSection As TargetDeadlinesSection = CurrentLogFrame.GetTargetDeadlinesSection(CurrentIndicator.Section)

        gbSubIndicators.Controls.Clear()
        Select Case CurrentIndicator.QuestionType
            Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Image, Indicator.QuestionTypes.ImageWithTargets, Indicator.QuestionTypes.Ranking
                lvSubIndicators = New ListViewSubIndicatorsBase(CurrentIndicator)

                gbTotals.Visible = False
                gbWeightingChildren.Visible = False
                gbSubIndicators.Dock = DockStyle.Fill
            Case Else
                Select Case CurrentIndicator.Registration
                    Case Indicator.RegistrationOptions.BeneficiaryLevel
                        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(CurrentIndicator.TargetGroupGuid)
                        lvSubIndicators = New ListViewSubIndicatorsBeneficiary(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    Case Else
                        lvSubIndicators = New ListViewSubIndicators(CurrentIndicator, CurrentTargetDeadlinesSection)
                End Select

                gbTotals.Visible = True
                gbWeightingChildren.Visible = True
        End Select

        lvSubIndicators.Name = "lvSubIndicators"
        lvSubIndicators.Dock = DockStyle.Fill

        AddHandler lvSubIndicators.IndicatorModified, AddressOf lvSubIndicators_IndicatorModified

        gbSubIndicators.Controls.Add(lvSubIndicators)
        lvSubIndicators.LoadItems()
    End Sub

    Private Sub CurrentDataGridViewNeedsUpdate(ByVal sender As Object, ByVal e As CurrentDataGridViewChangedEventArgs)
        Dim selDataGridView As DataGridView = CType(e.DataGridView, DataGridView)

        'set current targets datagridview for questiontypes that have multiple targets datagridviews (for different years)
        Select Case selDataGridView.GetType
            Case GetType(DataGridViewTargetsSemanticDiff), GetType(DataGridViewTargetsScaleLikert), GetType(DataGridViewTargetsFrequencyLikert)
                Me.CurrentTargetsDataGridView = selDataGridView
        End Select

        RaiseEvent CurrentDataGridViewChanged(sender, e)
    End Sub

    Private Sub ReloadTargets()
        TabPageTargets.Text = LANG_Targets

        If CurrentIndicator.Indicators.Count = 0 Then
            'show/hide targets tab
            Select Case CurrentIndicator.QuestionType
                Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff
                    ReloadTargets_Hide()
                    Exit Sub
                Case Else
                    If TabControlIndicator.TabPages.ContainsKey(strHiddenPageName) = False Then _
                        TabControlIndicator.TabPages.Insert(2, objHiddenTab)
            End Select

            Dim CurrentTargetDeadlinesSection As TargetDeadlinesSection = CurrentLogFrame.GetTargetDeadlinesSection(CurrentIndicator.Section)
            Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(CurrentIndicator.TargetGroupGuid)

            PanelTargets.Controls.Clear()

            Select Case CurrentIndicator.QuestionType

                Case Indicator.QuestionTypes.Image
                    TabPageTargets.Text = LANG_Baseline

                    Dim ctlImage As New ucImage(CurrentIndicator.Baseline.AudioVisualDetail)
                    ctlImage.Dock = DockStyle.Fill

                    PanelTargets.Controls.Add(ctlImage)
                Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                    Dim dgvTargets As New DataGridViewTargetsValues(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    With dgvTargets
                        .Name = "dgvTargets"
                        .Dock = DockStyle.Fill
                        .Reload()
                    End With
                    CurrentTargetsDataGridView = dgvTargets
                    PanelTargets.Controls.Add(dgvTargets)
                Case Indicator.QuestionTypes.Formula, Indicator.QuestionTypes.Ratio
                    Dim dgvTargets As New DataGridViewTargetsFormula(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    With dgvTargets
                        .Name = "dgvTargets"
                        .Dock = DockStyle.Fill
                        .Reload()
                    End With
                    CurrentTargetsDataGridView = dgvTargets
                    PanelTargets.Controls.Add(dgvTargets)
                Case Indicator.QuestionTypes.MultipleChoice, Indicator.QuestionTypes.MultipleOptions, Indicator.QuestionTypes.YesNo
                    Dim dgvTargets As New DataGridViewTargetsClasses(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    With dgvTargets
                        .Name = "dgvTargets"
                        .Dock = DockStyle.Fill
                        .Reload()
                    End With
                    CurrentTargetsDataGridView = dgvTargets
                    PanelTargets.Controls.Add(dgvTargets)

                Case Indicator.QuestionTypes.Ranking
                    Dim dgvTargets As New DataGridViewTargetsRanking(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    With dgvTargets
                        .Name = "dgvTargets"
                        .Dock = DockStyle.Fill
                        .Reload()
                    End With
                    CurrentTargetsDataGridView = dgvTargets
                    PanelTargets.Controls.Add(dgvTargets)

                Case Indicator.QuestionTypes.LikertTypeScale
                    Dim dgvTargets As New DataGridViewTargetsScaleLikertType(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    With dgvTargets
                        .Name = "dgvTargets"
                        .Dock = DockStyle.Fill
                        .Reload()
                    End With
                    CurrentTargetsDataGridView = dgvTargets
                    PanelTargets.Controls.Add(dgvTargets)

                Case Indicator.QuestionTypes.SemanticDiff
                    Dim ctlSemantifDiffTargets As New IndicatorTargetsSemanticDiff

                    With ctlSemantifDiffTargets
                        .CurrentIndicator = Me.CurrentIndicator
                        .TargetDeadlinesSection = CurrentTargetDeadlinesSection
                        .TargetGroup = selTargetGroup
                        .Dock = DockStyle.Fill
                        AddHandler .CurrentDataGridViewChanged, AddressOf CurrentDataGridViewNeedsUpdate
                    End With

                    PanelTargets.Controls.Add(ctlSemantifDiffTargets)
                    ctlSemantifDiffTargets.LoadScales()
                    CurrentTargetsDataGridView = ctlSemantifDiffTargets.CurrentTargetDatagridview

                Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale
                    Dim dgvTargets As New DataGridViewTargetsScales(CurrentIndicator, CurrentTargetDeadlinesSection, selTargetGroup)
                    With dgvTargets
                        .Name = "dgvTargets"
                        .Dock = DockStyle.Fill
                        .Reload()
                    End With
                    CurrentTargetsDataGridView = dgvTargets
                    PanelTargets.Controls.Add(dgvTargets)

                Case Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.FrequencyLikert
                    Dim ctlLikertScaleTargets As New IndicatorTargetsLikertScale

                    With ctlLikertScaleTargets
                        .CurrentIndicator = Me.CurrentIndicator
                        .TargetDeadlinesSection = CurrentTargetDeadlinesSection
                        .TargetGroup = selTargetGroup
                        .Dock = DockStyle.Fill
                        AddHandler .CurrentDataGridViewChanged, AddressOf CurrentDataGridViewNeedsUpdate
                    End With

                    PanelTargets.Controls.Add(ctlLikertScaleTargets)
                    ctlLikertScaleTargets.LoadScales()
                    CurrentTargetsDataGridView = ctlLikertScaleTargets.CurrentTargetDatagridview

                Case Indicator.QuestionTypes.ImageWithTargets
                    Dim ctlIndicatorTargetsImages As New IndicatorTargetsImages

                    With ctlIndicatorTargetsImages
                        .CurrentIndicator = Me.CurrentIndicator
                        .TargetDeadlinesSection = CurrentTargetDeadlinesSection
                        .TargetGroup = selTargetGroup
                        .Dock = DockStyle.Fill
                    End With

                    PanelTargets.Controls.Add(ctlIndicatorTargetsImages)
                    ctlIndicatorTargetsImages.LoadImages()

                Case Indicator.QuestionTypes.MixedSubIndicators
                    objHiddenTab = TabPageTargets
                    TabControlIndicator.TabPages.Remove(TabPageTargets)
                    If CurrentIndicatorsTab > TabControlIndicator.TabPages.Count - 1 Then
                        CurrentIndicatorsTab -= 1
                        TabControlIndicator.SelectedIndex = CurrentIndicatorsTab
                    End If
            End Select
        Else
            ReloadTargets_Hide()
        End If
    End Sub

    Private Sub ReloadTargets_Hide()
        objHiddenTab = TabPageTargets
        TabControlIndicator.TabPages.Remove(TabPageTargets)
        If CurrentIndicatorsTab > TabControlIndicator.TabPages.Count - 1 Then
            CurrentIndicatorsTab -= 1
            TabControlIndicator.SelectedIndex = CurrentIndicatorsTab
        End If
    End Sub

    Private Sub UpdateUnitsChildIndicators()
        If CurrentIndicator.ValuesDetail IsNot Nothing Then
            With CurrentIndicator.ValuesDetail
                CurrentIndicator.UpdateUnitsOfChildren(.ValueName, .NrDecimals, .Unit)
            End With

            lvSubIndicators.LoadItems()
        End If
    End Sub

    Private Sub TargetSystemUpdated()
        ReloadTargets()
    End Sub

    Private Sub ScoreSystemUpdated()
        ReloadTargets()
    End Sub

    Private Sub StatementsUpdated()
        ReloadTargets()
    End Sub

    Private Sub lvSubIndicators_IndicatorModified()
        CurrentProjectForm.dgvLogframe.Reload()
    End Sub

    Private Sub EditingControlShowing()
        With frmParent
            If .RibbonTabText.Active = False Then .RibbonLF.ActiveTab = .RibbonTabText
        End With
    End Sub
#End Region

#Region "Select text"
    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        If intTextSelectionIndex > 0 Then
            intTextSelectionIndex = frmProject.TextSelectionValues.SelectAll
        End If

        If CurrentStatementsDataGridView IsNot Nothing Then
            Select Case CurrentStatementsDataGridView.GetType
                Case GetType(DataGridViewStatementsFormula)
                    Dim dgvStatements As DataGridViewStatementsFormula = DirectCast(CurrentStatementsDataGridView, DataGridViewStatementsFormula)

                    With dgvStatements
                        .TextSelectionIndex = intTextSelectionIndex

                        .ResetAllImages()
                        .ReloadImages()
                        .Invalidate()
                    End With
                Case GetType(DataGridViewStatementsMaxDiff)
                    Dim dgvStatements As DataGridViewStatementsMaxDiff = DirectCast(CurrentStatementsDataGridView, DataGridViewStatementsMaxDiff)

                    With dgvStatements
                        .TextSelectionIndex = intTextSelectionIndex

                        .ResetAllImages()
                        .ReloadImages()
                        .Invalidate()
                    End With
                Case GetType(DataGridViewStatementsScales)
                    Dim dgvStatements As DataGridViewStatementsScales = DirectCast(CurrentStatementsDataGridView, DataGridViewStatementsScales)

                    With dgvStatements
                        .TextSelectionIndex = intTextSelectionIndex

                        .ResetAllImages()
                        .ReloadImages()
                        .Invalidate()
                    End With
            End Select
        End If

        If CurrentTargetsDataGridView IsNot Nothing Then
            Select Case CurrentTargetsDataGridView.GetType
                Case GetType(DataGridViewTargetsSemanticDiff)
                    Dim dgvStatements As DataGridViewTargetsSemanticDiff = DirectCast(CurrentTargetsDataGridView, DataGridViewTargetsSemanticDiff)

                    With dgvStatements
                        .TextSelectionIndex = intTextSelectionIndex

                        .ResetAllImages()
                        .ReloadImages()
                        .Invalidate()
                    End With
                Case GetType(DataGridViewTargetsScaleLikert)
                    Dim ctlLikertScaleTargets As IndicatorTargetsLikertScale = PanelTargets.Controls(0)

                    If ctlLikertScaleTargets IsNot Nothing Then
                        ctlLikertScaleTargets.SelectText(intTextSelectionIndex)
                    End If
                Case GetType(DataGridViewTargetsFrequencyLikert)
                    Dim ctlLikertScaleTargets As IndicatorTargetsLikertScale = PanelTargets.Controls(0)

                    If ctlLikertScaleTargets IsNot Nothing Then
                        ctlLikertScaleTargets.SelectText(intTextSelectionIndex)
                    End If
            End Select
        End If
    End Sub
#End Region
    
End Class
