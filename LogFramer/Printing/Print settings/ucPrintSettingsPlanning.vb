Imports System.Drawing.Printing

Public Class PrintSettingsPlanning
    Private objOutputGuid As Guid
    Private intPeriodView As Integer, intPlanningElements As Integer
    Private datStartDate, datEndDate, datPeriodStart, datPeriodEnd As Date
    Private boolShowActivityLinks, boolShowKeyMomentLinks, boolRepeatRowHeaders, boolShowDatesColumns As Boolean
    Private boolLoad As Boolean

    Public Event PrintPlanningSetupChanged(ByVal sender As Object, ByVal e As PrintPlanningSetupChangedEventArgs)

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

    Public Property OutputGuid() As Guid
        Get
            Return objOutputGuid
        End Get
        Set(ByVal value As Guid)
            objOutputGuid = value
            My.Settings.setPrintPlanningOutputGuid = value
        End Set
    End Property

    Public Property PlanningElements() As Integer
        Get
            Return intPlanningElements
        End Get
        Set(ByVal value As Integer)
            intPlanningElements = value
            My.Settings.setPrintPlanningElements = value
        End Set
    End Property

    Public Property PeriodView() As Integer
        Get
            Return intPeriodView
        End Get
        Set(ByVal value As Integer)
            intPeriodView = value

            If CurrentLogFrame IsNot Nothing Then
                Me.PeriodStart = CurrentLogFrame.GetExecutionStart(True)
                Me.PeriodEnd = CurrentLogFrame.GetExecutionEnd(True)
            End If
        End Set
    End Property

    Public Property PeriodStart() As Date
        Get
            Return datPeriodStart
        End Get
        Set(ByVal value As Date)
            If value <> datPeriodStart Then
                datPeriodStart = value
                My.Settings.setPrintPlanningPeriodFrom = datPeriodStart
                If boolLoad = False Then RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                        ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
            End If
        End Set
    End Property

    Public Property PeriodEnd()
        Get
            Return datPeriodEnd
        End Get
        Set(ByVal value)
            If value <> datPeriodEnd Then
                datPeriodEnd = value
                My.Settings.setPrintPlanningPeriodUntil = datPeriodEnd
                If boolLoad = False Then RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                        ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
            End If
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

    Public Property RepeatRowHeaders As Boolean
        Get
            Return boolRepeatRowHeaders
        End Get
        Set(ByVal value As Boolean)
            boolRepeatRowHeaders = value
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

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        SetPeriod()

        With cmbSelectOutput
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataSource = CurrentLogFrame.OutputsList
            .ValueMember = "Id"
            .DisplayMember = "Value"
        End With

        With Me.cmbShowElements
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.AddRange(LIST_ShowElements)
            .SelectedIndex = My.Settings.setPlanningElementsView
        End With

        With Me.cmbSelectPeriodView
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.AddRange(LIST_PeriodViews)
            .SelectedIndex = My.Settings.setPlanningPeriodView
        End With

        chkShowDatesColumns.Checked = My.Settings.setPrintPlanningShowDatesColumns
        chkShowActivityLinks.Checked = My.Settings.setPrintPlanningShowActivityLinks
        chkShowKeyMomentLinks.Checked = My.Settings.setPrintPlanningShowKeyMomentLinks
        chkRepeatRowHeaders.Checked = My.Settings.setPrintPlanningRepeatRowHeaders

        boolLoad = False
    End Sub

    Public Sub SetPeriod()
        If CurrentLogFrame.StartDate <> Date.MinValue Then
            Me.StartDate = CurrentLogFrame.StartDate
            Me.EndDate = CurrentLogFrame.EndDate
            Me.PeriodStart = CurrentLogFrame.GetExecutionStart
            Me.PeriodEnd = CurrentLogFrame.GetExecutionEnd
        Else
            Me.StartDate = New Date(Now.Year, 1, 1)
            Me.EndDate = New Date(Now.Year, 12, 31)
            Me.PeriodStart = New Date(Now.Year, 1, 1)
            Me.PeriodEnd = New Date(Now.Year, 12, 31)
        End If
        dtpPeriodFrom.Value = Me.PeriodStart
        dtpPeriodUntil.Value = Me.PeriodEnd
    End Sub

    Private Sub cmbSelectOutput_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectOutput.SelectedValueChanged
        Me.OutputGuid = CType(cmbSelectOutput.SelectedItem, IdValuePair).Id
        If Me.OutputGuid = CurrentLogFrame.Guid Then Me.OutputGuid = Nothing
        If boolLoad = False Then RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub cmbShowElements_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbShowElements.SelectedIndexChanged
        Me.PlanningElements = cmbShowElements.SelectedIndex
        If boolLoad = False Then RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub cmbSelectView_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectPeriodView.SelectedIndexChanged
        Me.PeriodView = cmbSelectPeriodView.SelectedIndex
        If boolLoad = False Then RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub chkShowDatesColumns_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkShowDatesColumns.CheckedChanged
        If boolLoad = True Then Exit Sub
        Me.ShowDatesColumns = chkShowDatesColumns.Checked
        My.Settings.setPrintPlanningShowDatesColumns = Me.ShowDatesColumns
        RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                                                                                        ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub chkActivityShowLinks_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowActivityLinks.CheckedChanged
        If boolLoad = True Then Exit Sub
        Me.ShowActivityLinks = chkShowActivityLinks.Checked
        My.Settings.setPrintPlanningShowActivityLinks = Me.ShowActivityLinks
        RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                                                                                        ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub chkKeyMomentShowLinks_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowKeyMomentLinks.CheckedChanged
        If boolLoad = True Then Exit Sub
        Me.ShowKeyMomentLinks = chkShowKeyMomentLinks.Checked
        My.Settings.setPrintPlanningShowKeyMomentLinks = Me.ShowKeyMomentLinks
        RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                                                                                        ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub chkRepeatRowHeaders_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRepeatRowHeaders.CheckedChanged
        If boolLoad = True Then Exit Sub
        Me.RepeatRowHeaders = chkRepeatRowHeaders.Checked
        My.Settings.setPrintPlanningRepeatRowHeaders = Me.RepeatRowHeaders
        RaiseEvent PrintPlanningSetupChanged(Me, New PrintPlanningSetupChangedEventArgs(OutputGuid, PeriodView, PlanningElements, PeriodStart, PeriodEnd, _
                                                                                        ShowActivityLinks, ShowKeyMomentLinks, RepeatRowHeaders, ShowDatesColumns))
    End Sub

    Private Sub dtpPeriodFrom_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriodFrom.CloseUp
        PeriodStart = New Date(dtpPeriodFrom.Value.Year, dtpPeriodFrom.Value.Month, dtpPeriodFrom.Value.Day, 0, 0, 0)
    End Sub

    Private Sub dtpPeriodFrom_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriodFrom.Validated
        PeriodStart = New Date(dtpPeriodFrom.Value.Year, dtpPeriodFrom.Value.Month, dtpPeriodFrom.Value.Day, 0, 0, 0)
    End Sub

    Private Sub dtpPeriodUntil_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriodUntil.CloseUp
        PeriodEnd = New Date(dtpPeriodUntil.Value.Year, dtpPeriodUntil.Value.Month, dtpPeriodUntil.Value.Day, 0, 0, 0)
    End Sub

    Private Sub dtpPeriodUntil_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriodUntil.Validated
        PeriodEnd = New Date(dtpPeriodUntil.Value.Year, dtpPeriodUntil.Value.Month, dtpPeriodUntil.Value.Day, 0, 0, 0)
    End Sub
End Class

Public Class PrintPlanningSetupChangedEventArgs
    Inherits EventArgs

    Public Property OutputGuid As Guid
    Public Property PeriodView As Integer
    Public Property PlanningElements As Integer
    Public Property PeriodFrom As Date
    Public Property PeriodUntil As Date
    Public Property ShowActivityLinks As Boolean
    Public Property ShowKeyMomentLinks As Boolean
    Public Property RepeatRowHeaders As Boolean
    Public Property ShowDatesColumns As Boolean

    Public Sub New(ByVal objOutputGuid As Guid, ByVal intPeriodView As Integer, ByVal intPlanningElements As Integer, ByVal datPeriodFrom As Date, ByVal datPeriodUntil As Date, _
                   boolShowActivityLinks As Boolean, ByVal boolShowKeyMomentLinks As Boolean, boolRepeatRowHeaders As Boolean, boolShowDatesColumns As Boolean)
        MyBase.New()

        OutputGuid = objOutputGuid
        PeriodView = intPeriodView
        PlanningElements = intPlanningElements
        PeriodFrom = datPeriodFrom
        PeriodUntil = datPeriodUntil
        ShowActivityLinks = boolShowActivityLinks
        ShowKeyMomentLinks = boolShowKeyMomentLinks
        RepeatRowHeaders = boolRepeatRowHeaders
        ShowDatesColumns = boolShowDatesColumns
    End Sub
End Class
