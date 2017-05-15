Imports System.Xml.Serialization
Imports System.Web.Script.Serialization
Imports System.Drawing.Printing

<XmlRoot()> _
Public Class LogFrame

#Region "Enumerations"
    Public Enum SectionTypes
        GoalsSection = 1
        PurposesSection = 2
        OutputsSection = 3
        ActivitiesSection = 4
    End Enum

    Public Enum ObjectTypes
        NotSet = -1
        LogFrame = 1
        ProjectPartner = 2
        Contact = 3
        Organisation = 4

        Struct = 10
        Goal = 11
        Purpose = 12
        Output = 13
        Activity = 14
        PurposeHidden = 15
        OutputHidden = 16

        Indicator = 20
        Statement = 21
        BooleanResponse = 22
        IntegerResponse = 23
        ValueRange = 24
        ResponseClass = 25
        Target = 26
        Baseline = 27

        VerificationSource = 30
        Resource = 40
        ResourceBudget = 41
        Budget = 50
        BudgetYear = 51
        BudgetItem = 52
        BudgetItemReference = 53
        Assumption = 60

        Address = 70
        TelephoneNumber = 71
        Email = 72
        Website = 73
        TargetGroup = 80
        TargetGroupInformation = 81
        KeyMoment = 90
    End Enum

    Public Enum DurationUnits
        Days = 0
        Weeks = 1
        Months = 2
        Years = 3
    End Enum
#End Region

#Region "Collections"
    Private objGoals As New Goals
    Private objPurposes As New Purposes
    Private objProjectPartners As New ProjectPartners
    Private objKeyMoments As New KeyMoments
    Private objBudget As New Budget

    <ScriptIgnore()> _
    Public Property Goals As Goals
        Get
            Return objGoals
        End Get
        Set(ByVal value As Goals)
            objGoals = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property Purposes As Purposes
        Get
            Return objPurposes
        End Get
        Set(ByVal value As Purposes)
            objPurposes = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ProjectPartners As ProjectPartners
        Get
            Return objProjectPartners
        End Get
        Set(ByVal value As ProjectPartners)
            objProjectPartners = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property Keymoments() As KeyMoments
        Get
            Return objKeyMoments
        End Get
        Set(ByVal value As KeyMoments)
            objKeyMoments = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property Budget As Budget
        Get
            Return objBudget
        End Get
        Set(value As Budget)
            objBudget = value
        End Set
    End Property
#End Region

#Region "Properties"
    Public LogFramerVersion As String = "2.0"

    Private intIdLogframe As Integer = -1
    Private strProjectTitle As String
    Private strShortTitle As String
    Private strCode As String
    Private sngDuration As Single
    Private intDurationUnit As Integer
    Private datStartDate As Date, datEndDate As Date
    Private strCurrencyCode As String
    Private strSortNumberDivider As String = "-"
    Private boolSortNumberRepeatParent As Boolean = True
    Private boolSortNumberTerminateWithDivider As Boolean
    Private intSchemeIndex As Integer
    Private strStructName(3) As String
    Private strStructNamePlural(3) As String
    Private strIndicatorName, strVerificationSourceName As String
    Private strResourceName, strBudgetName As String
    Private strAssumptionName As String
    Private objTargetDeadlinesGoals As New TargetDeadlinesSection(SectionTypes.GoalsSection)
    Private objTargetDeadlinesPurposes As New TargetDeadlinesSection(SectionTypes.PurposesSection)
    Private objTargetDeadlinesOutputs As New TargetDeadlinesSection(SectionTypes.OutputsSection)
    Private objTargetDeadlinesActivities As New TargetDeadlinesSection(SectionTypes.ActivitiesSection)
    Private objRiskMonitoring As New RiskMonitoring
    Private objGuid As Guid

    Public Property idLogframe As Integer
        Get
            Return intIdLogframe
        End Get
        Set(ByVal value As Integer)
            intIdLogframe = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ProjectTitle() As String
        Get
            Return strProjectTitle
        End Get
        Set(ByVal value As String)
            strProjectTitle = value
        End Set
    End Property

    Public Property ShortTitle() As String
        Get
            Return strShortTitle
        End Get
        Set(value As String)
            strShortTitle = value
        End Set
    End Property

    Public Property Code As String
        Get
            Return strCode
        End Get
        Set(value As String)
            strCode = value
        End Set
    End Property

    Public Property Duration() As Single
        Get
            Return sngDuration
        End Get
        Set(ByVal value As Single)
            sngDuration = value
        End Set
    End Property

    Public Property DurationUnit() As Integer
        Get
            Return intDurationUnit
        End Get
        Set(ByVal value As Integer)
            intDurationUnit = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Dim selDate As Date
            For Each selMoment As KeyMoment In Me.Keymoments
                If selMoment.Type = KeyMoment.Types.ProjectStart Then
                    selDate = selMoment.KeyMoment
                    Exit For
                End If
            Next
            selDate = DateSerial(selDate.Year, selDate.Month, selDate.Day)
            Return selDate
        End Get
        Set(ByVal value As Date)
            Dim boolFound As Boolean

            value = DateSerial(value.Year, value.Month, value.Day)

            For Each selMoment As KeyMoment In Me.Keymoments
                If selMoment.Type = KeyMoment.Types.ProjectStart Then
                    selMoment.KeyMoment = value
                    selMoment.Description = LANG_ProjectStart
                    selMoment.Relative = False
                    boolFound = True
                    Exit For
                End If
            Next
            If boolFound = False Then
                Dim NewMoment As New KeyMoment(value, LANG_ProjectStart, KeyMoment.Types.ProjectStart)
                NewMoment.Relative = False
                Me.Keymoments.Add(NewMoment)
            End If
        End Set
    End Property

    Public Property EndDate() As Date
        Get
            Dim selDate As Date
            For Each selMoment As KeyMoment In Me.Keymoments
                If selMoment.Type = KeyMoment.Types.ProjectEnd Then
                    selDate = selMoment.KeyMoment
                    Exit For
                End If
            Next
            selDate = DateSerial(selDate.Year, selDate.Month, selDate.Day)
            Return selDate
        End Get
        Set(ByVal value As Date)
            Dim boolFound As Boolean

            value = DateSerial(value.Year, value.Month, value.Day)

            For Each selMoment As KeyMoment In Me.Keymoments
                If selMoment.Type = KeyMoment.Types.ProjectEnd Then
                    selMoment.KeyMoment = value
                    selMoment.Description = LANG_ProjectEnd
                    selMoment.Relative = False
                    boolFound = True
                    Exit For
                End If
            Next
            If boolFound = False Then
                Dim NewMoment As New KeyMoment(value, LANG_ProjectEnd, KeyMoment.Types.ProjectEnd)
                NewMoment.Relative = False
                Me.Keymoments.Add(NewMoment)
            End If
        End Set
    End Property

    Public Property CurrencyCode As String
        Get
            If String.IsNullOrEmpty(strCurrencyCode) Then strCurrencyCode = My.Settings.setDefaultCurrency
            Return strCurrencyCode
        End Get
        Set(ByVal value As String)
            strCurrencyCode = value.ToUpper.Substring(0, 3)
        End Set
    End Property

    Public Property SortNumberDivider As String
        Get
            Return strSortNumberDivider
        End Get
        Set(ByVal value As String)
            strSortNumberDivider = value
        End Set
    End Property

    Public Property SortNumberRepeatParent As Boolean
        Get
            Return boolSortNumberRepeatParent
        End Get
        Set(ByVal value As Boolean)
            boolSortNumberRepeatParent = value
        End Set
    End Property

    Public Property SortNumberTerminateWithDivider As Boolean
        Get
            Return boolSortNumberTerminateWithDivider
        End Get
        Set(ByVal value As Boolean)
            boolSortNumberTerminateWithDivider = value
        End Set
    End Property

    Public Property SchemeIndex As Integer
        Get
            Return intSchemeIndex
        End Get
        Set(value As Integer)
            intSchemeIndex = value
        End Set
    End Property

    Public Property StructName() As String()
        Get
            Return strStructName
        End Get
        Set(ByVal value As String())
            strStructName = value
        End Set
    End Property

    Public Property StructNamePlural() As String()
        Get
            Return strStructNamePlural
        End Get
        Set(ByVal value As String())
            strStructNamePlural = value
        End Set
    End Property

    Public Property IndicatorName As String
        Get
            If String.IsNullOrEmpty(strIndicatorName) Then strIndicatorName = LANG_Indicators
            Return strIndicatorName
        End Get
        Set(ByVal value As String)
            strIndicatorName = value
        End Set
    End Property

    Public Property VerificationSourceName As String
        Get
            If String.IsNullOrEmpty(strVerificationSourceName) Then strVerificationSourceName = LANG_VerificationSources
            Return strVerificationSourceName
        End Get
        Set(ByVal value As String)
            strVerificationSourceName = value
        End Set
    End Property

    Public Property ResourceName As String
        Get
            If String.IsNullOrEmpty(strResourceName) Then strResourceName = LANG_Resources
            Return strResourceName
        End Get
        Set(ByVal value As String)
            strResourceName = value
        End Set
    End Property

    Public Property BudgetName As String
        Get
            If String.IsNullOrEmpty(strBudgetName) Then strBudgetName = LANG_Budget
            Return strBudgetName
        End Get
        Set(ByVal value As String)
            strBudgetName = value
        End Set
    End Property

    Public Property AssumptionName As String
        Get
            If String.IsNullOrEmpty(strAssumptionName) Then strAssumptionName = LANG_Assumptions
            Return strAssumptionName
        End Get
        Set(ByVal value As String)
            strAssumptionName = value
        End Set
    End Property

    Public Property TargetDeadlinesGoals As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesGoals
        End Get
        Set(value As TargetDeadlinesSection)
            objTargetDeadlinesGoals = value
            objTargetDeadlinesGoals.Section = LogFrame.SectionTypes.GoalsSection
        End Set
    End Property

    Public Property TargetDeadlinesPurposes As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesPurposes
        End Get
        Set(value As TargetDeadlinesSection)
            objTargetDeadlinesPurposes = value
            objTargetDeadlinesPurposes.Section = LogFrame.SectionTypes.PurposesSection
        End Set
    End Property

    Public Property TargetDeadlinesOutputs As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesOutputs
        End Get
        Set(value As TargetDeadlinesSection)
            objTargetDeadlinesOutputs = value
            objTargetDeadlinesOutputs.Section = LogFrame.SectionTypes.OutputsSection
        End Set
    End Property

    Public Property TargetDeadlinesActivities As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesActivities
        End Get
        Set(value As TargetDeadlinesSection)
            objTargetDeadlinesActivities = value
            objTargetDeadlinesActivities.Section = LogFrame.SectionTypes.ActivitiesSection
        End Set
    End Property

    Public Property RiskMonitoring As RiskMonitoring
        Get
            Return objRiskMonitoring
        End Get
        Set(ByVal value As RiskMonitoring)
            objRiskMonitoring = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property IsEmpty() As Boolean
        Get
            Dim boolIsEmpty As Boolean

            If Me.Goals.Count = 0 And Me.Purposes.Count = 0 Then
                boolIsEmpty = True
            Else
                boolIsEmpty = False
            End If
            Return boolIsEmpty
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property StructCount As Integer
        Get
            Dim intStructCount As Integer = Me.Goals.Count

            intStructCount += Me.Purposes.Count
            For Each selPurpose As Purpose In Me.Purposes
                intStructCount += selPurpose.Outputs.Count
                For Each selOutput As Output In selPurpose.Outputs
                    intStructCount += selOutput.Activities.Count
                Next
            Next

            Return intStructCount
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Project
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Projects
        End Get
    End Property
#End Region

#Region "Lists"
    Private lstUnits As New List(Of String)
    Private SearchUnit As String

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property TargetGroupNamesList() As String()
        Get
            Dim alNames As New List(Of String)
            alNames.Add(String.Empty)
            For Each selPurpose As Purpose In Me.Purposes
                For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                    alNames.Add(selTargetGroup.Name)
                Next
            Next
            Dim strNames(alNames.Count - 1) As String
            strNames = alNames.ToArray
            Return strNames
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property PartnerNamesList() As String()
        Get
            Dim alNames As New List(Of String)
            alNames.Add(String.Empty)
            For Each selPartner As ProjectPartner In Me.ProjectPartners
                alNames.Add(selPartner.Organisation.Name)
            Next
            Dim strNames(alNames.Count - 1) As String
            strNames = alNames.ToArray
            Return strNames
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property OutputsList() As List(Of IdValuePair)
        Get
            Dim lstOutputs As New List(Of IdValuePair)
            Dim strSortNumber As String
            Dim Language As String = My.Settings.setLanguage

            lstOutputs.Add(New IdValuePair(Me.Guid, LANG_AllOutputs))

            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    strSortNumber = Me.GetStructSortNumber(selOutput)
                    lstOutputs.Add(New IdValuePair(selOutput.Guid, strSortNumber & ". " & selOutput.Text))
                Next
            Next
            Return lstOutputs
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ActivitiesLocationsList() As List(Of String)
        Get
            Dim lstLocations As New List(Of String)
            lstLocations.Add("")
            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    For Each selLocation As String In selOutput.Activities.LocationsList
                        If lstLocations.Contains(selLocation) = False Then _
                                lstLocations.Add(selLocation)
                    Next
                Next
            Next
            Return lstLocations
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property KeyMomentsList() As KeyMoments
        Get
            Dim objAllMoments As New KeyMoments
            objAllMoments.AddRange(Me.Keymoments.Sort)

            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    objAllMoments.AddRange(selOutput.KeyMoments.Sort)
                Next
            Next

            Return objAllMoments
        End Get
    End Property
#End Region

#Region "Lay-out properties"
    Public SectionTitleFontName As String
    Public SectionTitleFontSize As Single
    Public SectionTitleFontBold As Boolean
    Public SectionTitleFontItalics As Boolean
    Public SectionTitleFontUnderlined As Boolean
    Public SortNumberFontName As String
    Public SortNumberFontSize As Single
    Public SortNumberFontBold As Boolean
    Public SortNumberFontItalics As Boolean
    Public SortNumberFontUnderlined As Boolean
    Public DetailsFontName As String
    Public DetailsFontSize As Single
    Private colSectionTitleColor As Color
    Private colSectionColorTop As Color
    Private colSectionColorBottom As Color

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public SectionTitles As New SectionTitles

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property SectionTitleFont As Font
        Get
            Dim fntSection As Font
            If SectionTitleFontName = String.Empty Then
                fntSection = My.Settings.setDefaultFontSections
            Else
                fntSection = New Font(SectionTitleFontName, SectionTitleFontSize)
                If SectionTitleFontBold = True Then fntSection = New Font(fntSection, FontStyle.Bold)
                If SectionTitleFontItalics = True Then fntSection = New Font(fntSection, FontStyle.Italic)
                If SectionTitleFontUnderlined = True Then fntSection = New Font(fntSection, FontStyle.Underline)
            End If

            Return fntSection
        End Get
    End Property

    Public Sub SetSectionTitleFont(ByVal fntSection As Font, ByVal colTextColor As Color)
        SectionTitleFontName = fntSection.Name
        SectionTitleFontSize = fntSection.SizeInPoints
        SectionTitleFontBold = fntSection.Bold
        SectionTitleFontItalics = fntSection.Italic
        SectionTitleFontUnderlined = fntSection.Underline
        SectionTitleFontColor = colTextColor
    End Sub

    Public Property SectionTitleFontColor As Color
        Get
            If colSectionTitleColor = Nothing Then colSectionTitleColor = My.Settings.setDefaultFontSectionColor
            Return colSectionTitleColor
        End Get
        Set(ByVal value As Color)
            colSectionTitleColor = value
        End Set
    End Property

    Public Property SectionColorTop As Color
        Get
            If colSectionColorTop = Nothing Then colSectionColorTop = My.Settings.setDefaultColorSectionTop
            Return colSectionColorTop
        End Get
        Set(ByVal value As Color)
            colSectionColorTop = value
        End Set
    End Property

    Public Property SectionColorBottom As Color
        Get
            If colSectionColorBottom = Nothing Then colSectionColorBottom = My.Settings.setDefaultColorSectionBottom
            Return colSectionColorBottom
        End Get
        Set(ByVal value As Color)
            colSectionColorBottom = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property SortNumberFont As Font
        Get
            Dim fntSection As Font
            If SortNumberFontName = String.Empty Then
                fntSection = My.Settings.setDefaultFontSortNumbers
            Else
                fntSection = New Font(SortNumberFontName, SortNumberFontSize)
                If SortNumberFontBold = True Then fntSection = New Font(fntSection, FontStyle.Bold)
                If SortNumberFontItalics = True Then fntSection = New Font(fntSection, FontStyle.Italic)
                If SortNumberFontUnderlined = True Then fntSection = New Font(fntSection, FontStyle.Underline)
            End If

            Return fntSection
        End Get
    End Property

    Public Sub SetSortNumberFont(ByVal fntSection As Font)
        SortNumberFontName = fntSection.Name
        SortNumberFontSize = fntSection.SizeInPoints
        SortNumberFontBold = fntSection.Bold
        SortNumberFontItalics = fntSection.Italic
        SortNumberFontUnderlined = fntSection.Underline
    End Sub

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property DetailsFont As Font
        Get
            Dim fntSection As Font
            If DetailsFontName = String.Empty Then
                fntSection = My.Settings.setDefaultFontUI
            Else
                fntSection = New Font(DetailsFontName, DetailsFontSize)
            End If

            Return fntSection
        End Get
    End Property

    Public Sub SetDetailsFont(ByVal fntSection As Font)
        DetailsFontName = fntSection.Name
        DetailsFontSize = fntSection.SizeInPoints
    End Sub
#End Region

#Region "Print properties"
    Private objReportSetupLogframe As New ReportSetup
    Private objReportSetupPlanning As New ReportSetup
    Private objReportSetupBudget As New ReportSetup
    Private objReportSetupPmf As New ReportSetup
    Private objReportSetupIndicators As New ReportSetup
    Private objReportSetupResources As New ReportSetup
    Private objReportSetupRiskRegister As New ReportSetup
    Private objReportSetupAssumptions As New ReportSetup
    Private objReportSetupDependencies As New ReportSetup
    Private objReportSetupPartnerList As New ReportSetup
    Private objReportSetupTargetGroupIdForm As New ReportSetup

    Private boolPrintLogFrameColor As Boolean = True
    Private boolPrintLogFrameLandScape As Boolean = False
    Private margPrintLogFrameMargins As New Margins
    Private psPrintLogFramePaperSize As New PaperSize

    <ScriptIgnore()> _
    Public Property ReportSetupLogframe As ReportSetup
        Get
            Return objReportSetupLogframe
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupLogframe = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupPlanning As ReportSetup
        Get
            Return objReportSetupPlanning
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupPlanning = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupBudget As ReportSetup
        Get
            Return objReportSetupBudget
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupBudget = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupIndicators As ReportSetup
        Get
            Return objReportSetupIndicators
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupIndicators = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupPmf As ReportSetup
        Get
            Return objReportSetupPmf
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupPmf = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupResources As ReportSetup
        Get
            Return objReportSetupResources
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupResources = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupRiskRegister As ReportSetup
        Get
            Return objReportSetupRiskRegister
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupRiskRegister = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupAssumptions As ReportSetup
        Get
            Return objReportSetupAssumptions
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupAssumptions = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupDependencies As ReportSetup
        Get
            Return objReportSetupDependencies
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupDependencies = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupPartnerList As ReportSetup
        Get
            Return objReportSetupPartnerList
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupPartnerList = value
        End Set
    End Property

    <ScriptIgnore()> _
    Public Property ReportSetupTargetGroupIdForm As ReportSetup
        Get
            Return objReportSetupTargetGroupIdForm
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetupTargetGroupIdForm = value
        End Set
    End Property
#End Region 'Print properties

#Region "Initialise"
    Public Sub New()
        If Me.StructNamePlural(0) = String.Empty Then
            Me.StructName = New String() {My.Settings.setStruct1sing, My.Settings.setStruct2sing, My.Settings.setStruct3sing, My.Settings.setStruct4sing}
            Me.StructNamePlural = New String() {My.Settings.setStruct1, My.Settings.setStruct2, My.Settings.setStruct3, My.Settings.setStruct4}
            CreateSectionTitles()
        End If

        'default page settings
        Dim DefaultPrinterSettings As New PrinterSettings
        With DefaultPrinterSettings
            margPrintLogFrameMargins = .DefaultPageSettings.Margins
            psPrintLogFramePaperSize = .DefaultPageSettings.PaperSize
        End With
    End Sub

    Public Sub CreateSectionTitles()
        Me.SectionTitles.Clear()

        Dim SectionGoals As New SectionTitle(SectionTitleSetFont(Me.StructNamePlural(0)))
        Dim SectionIndicator As New Indicator(SectionTitleSetFont(LANG_Indicators))
        SectionGoals.Indicators.Add(SectionIndicator)
        Dim SectionVerificationSource As New VerificationSource(SectionTitleSetFont(LANG_VerificationSources))
        SectionIndicator.VerificationSources.Add(SectionVerificationSource)
        Dim SectionAssumption As New Assumption(SectionTitleSetFont(LANG_Assumptions))
        SectionGoals.Assumptions.Add(SectionAssumption)

        Dim SectionPurposes As New SectionTitle(SectionTitleSetFont(Me.StructNamePlural(1)))
        SectionPurposes.Indicators.Add(SectionIndicator)
        SectionPurposes.Assumptions.Add(SectionAssumption)

        Dim SectionOutputs As New SectionTitle(SectionTitleSetFont(Me.StructNamePlural(2)))
        SectionOutputs.Indicators.Add(SectionIndicator)
        SectionOutputs.Assumptions.Add(SectionAssumption)

        Dim SectionActivities As New SectionTitle(SectionTitleSetFont(Me.StructNamePlural(3)))
        SectionActivities.Indicators.Add(SectionIndicator)
        SectionActivities.Assumptions.Add(SectionAssumption)
        Dim SectionResource As New Resource(SectionTitleSetFont(LANG_Resources))
        SectionActivities.Resources.Add(SectionResource)

        Me.SectionTitles.Add(SectionGoals)
        Me.SectionTitles.Add(SectionPurposes)
        Me.SectionTitles.Add(SectionOutputs)
        Me.SectionTitles.Add(SectionActivities)
    End Sub

    Private Function SectionTitleSetFont(ByVal strText As String) As String
        Using objRtf As New RichTextManager
            objRtf.Text = strText
            objRtf.SelectAll()
            objRtf.SelectionFont = Me.SectionTitleFont
            objRtf.SelectionColor = Me.SectionTitleFontColor
            objRtf.SelectionAlignment = HorizontalAlignment.Center
            strText = objRtf.Rtf
        End Using
        Return strText
    End Function
#End Region

#Region "Add and remove"
    Public Function AddNewStruct(ByVal intSection As Integer) As Struct
        Dim NewStruct As Struct = Nothing
        Select Case intSection
            Case SectionTypes.GoalsSection
                NewStruct = New Goal
                Me.Goals.Add(NewStruct)
            Case SectionTypes.PurposesSection
                NewStruct = New Purpose
                Me.Purposes.Add(NewStruct)
            Case SectionTypes.OutputsSection
                NewStruct = New Output
                If Me.Purposes.Count = 0 Then AddNewStruct(SectionTypes.PurposesSection)
                Dim selPurpose As Purpose = Me.Purposes(Me.Purposes.Count - 1)
                selPurpose.Outputs.Add(NewStruct)
            Case SectionTypes.ActivitiesSection
                NewStruct = New Activity
                If Me.Purposes.Count = 0 Then AddNewStruct(SectionTypes.PurposesSection)
                Dim selPurpose As Purpose = Me.Purposes(Me.Purposes.Count - 1)
                If selPurpose.Outputs.Count = 0 Then AddNewStruct(SectionTypes.OutputsSection)
                Dim selOutput As Output = selPurpose.Outputs(selPurpose.Outputs.Count - 1)
                selOutput.Activities.Add(NewStruct)
        End Select
        Return NewStruct
    End Function

    Public Sub RemoveStruct(ByVal struct As Struct)
        Select Case struct.Section
            Case SectionTypes.GoalsSection
                Me.Goals.Remove(struct)
            Case SectionTypes.PurposesSection
                Me.Purposes.Remove(struct)
            Case SectionTypes.OutputsSection
                Dim ParentPurpose As Purpose = GetParent(CType(struct, Output))
                ParentPurpose.Outputs.Remove(struct)
            Case SectionTypes.ActivitiesSection
                Dim ParentActivities As Activities = GetParentCollection(CType(struct, Activity))
                ParentActivities.Remove(struct)
        End Select
    End Sub
#End Region

#Region "General methods"
    Public Function CreateSortNumber(ByVal intIndex As Integer, Optional ByVal strParentSort As String = "") As String
        If intIndex < 0 Then Return String.Empty

        Dim strSort As String
        Dim strNumber As String = (intIndex + 1).ToString

        If SortNumberTerminateWithDivider = False Then
            If String.IsNullOrEmpty(strParentSort) = False Then
                strParentSort &= CurrentLogFrame.SortNumberDivider
            End If
            strSort = strParentSort & strNumber
        Else
            strSort = strParentSort & strNumber & SortNumberDivider
        End If
        Return strSort
    End Function
#End Region

End Class

Public Class Logframes
    Inherits System.Collections.CollectionBase


    Public Sub Add(ByVal logframe As LogFrame)
        List.Add(logframe)
    End Sub

    Public Function Contains(ByVal logframe As LogFrame) As Boolean
        Return List.Contains(logframe)
    End Function

    Public Function IndexOf(ByVal logframe As LogFrame) As Integer
        Return List.IndexOf(logframe)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal logframe As LogFrame)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of Logframe is not valid, cannot be inserted!")
        ElseIf index = Count Then
            List.Add(logframe)
        Else
            List.Insert(index, logframe)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Logframe
        Get
            Return CType(List.Item(index), Logframe)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal logframe As LogFrame)
        If List.Contains(logframe) Then List.Remove(logframe)
    End Sub

    Public Function GetLogframeByGuid(ByVal objGuid As Guid) As Logframe
        Dim selLogframe As Logframe = Nothing
        For Each objLogframe As Logframe In Me.List
            If objLogframe.Guid = objGuid Then
                selLogframe = objLogframe
                Exit For
            End If
        Next
        Return selLogframe
    End Function

    
End Class
