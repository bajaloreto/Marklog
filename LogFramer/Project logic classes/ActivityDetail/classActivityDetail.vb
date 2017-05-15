Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ActivityDetail
    Private intIdActivityDetail As Integer = -1
    Private intIdStruct As Integer
    Private objGuid As Guid
    Private objReferenceMoment As New Guid
    Private intType As Integer
    Private boolRelative As Boolean 'Absolute date
    Private datStartDate As Date = Now.Date
    Private sngPeriod As Single
    Private intPeriodUnit As Integer
    Private intPeriodDirection As Integer
    Private intDuration As Integer = 1
    Private intDurationUnit As Integer = 3 'Day
    Private boolDurationUntilEnd As Boolean
    Private sngRepeatEvery As Single
    Private intRepeatUnit As Integer
    Private intRepeatTimes As Integer
    Private lstRepeatStartDates As New List(Of Date), lstRepeatEndDates As New List(Of Date)
    Private boolRepeatUntilEnd As Boolean
    Private intPreparation As Integer
    Private intPreparationUnit As Integer
    Private boolPreparationFromStart As Boolean
    Private boolPreparationRepeat As Boolean
    Private intFollowUp As Integer
    Private intFollowUpUnit As Integer
    Private boolFollowUpUntilEnd As Boolean
    Private boolFollowUpRepeat As Boolean
    Private strOrganisation As String
    Private strLocation As String

#Region "Enumerations"
    Public Enum Types As Integer
        Other = 0
        Planning = 10
        ProjectProposal = 11
        DonorSigned = 12
        Report = 13
        Monitoring = 14
        Audit = 15
        Evaluation = 16

        Identification = 20
        Selection = 21
        Registration = 22

        Procurement = 30
        HiringStaff = 31
        Logistics = 32
        Meeting = 33
        Debriefing = 34
        Travel = 35

        Construction = 40
        Research = 41
        Distribution = 42
        Activity = 43
        MedicalTreatment = 44
        HumanitarianAssistance = 45
        PeaceBuilding = 46
        Training = 47
        CapacityDevelopment = 48
        Awareness = 49
    End Enum

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

    Public Enum DirectionValues
        NotDefined = 0
        Before = 1
        After = 2
    End Enum
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub CalculateRepeatDates()
        If RepeatEvery > 0 And RepeatUnit > 0 Then
            Dim NewStartDate As Date = StartDateMainActivity
            Dim NewEndDate As Date = EndDateMainActivity

            RepeatStartDates.Clear()
            RepeatEndDates.Clear()

            If RepeatUntilEnd = False Then
                For i = 1 To RepeatTimes
                    CalculateRepeatDates_Repeat(NewStartDate, NewEndDate)
                Next
            Else
                Do While NewEndDate < CurrentLogFrame.EndDate
                    CalculateRepeatDates_Repeat(NewStartDate, NewEndDate)
                Loop
                If RepeatStartDates.Count > 1 Then
                    RepeatStartDates.RemoveAt(RepeatStartDates.Count - 1)
                    RepeatEndDates.RemoveAt(RepeatEndDates.Count - 1)
                End If
                intRepeatTimes = lstRepeatStartDates.Count - 1
            End If
        End If
    End Sub

    Private Sub CalculateRepeatDates_Repeat(ByRef NewStartDate As Date, ByRef NewEndDate As Date)
        Select Case RepeatUnit
            Case DurationUnits.Day
                NewStartDate = NewStartDate.AddDays(Me.RepeatEvery)
                NewEndDate = NewEndDate.AddDays(Me.RepeatEvery)
            Case DurationUnits.Hour
                NewStartDate = NewStartDate.AddHours(Me.RepeatEvery)
                NewEndDate = NewEndDate.AddHours(Me.RepeatEvery)
            Case DurationUnits.Minute
                NewStartDate = NewStartDate.AddMinutes(Me.RepeatEvery)
                NewEndDate = NewEndDate.AddMinutes(Me.RepeatEvery)
            Case DurationUnits.Month
                NewStartDate = NewStartDate.AddMonths(Me.RepeatEvery)
                NewEndDate = NewEndDate.AddMonths(Me.RepeatEvery)
            Case DurationUnits.Semester
                NewStartDate = NewStartDate.AddMonths(Me.RepeatEvery * 6)
                NewEndDate = NewEndDate.AddMonths(Me.RepeatEvery * 6)
            Case DurationUnits.Trimester
                NewStartDate = NewStartDate.AddMonths(Me.RepeatEvery * 3)
                NewEndDate = NewEndDate.AddMonths(Me.RepeatEvery * 3)
            Case DurationUnits.Week
                NewStartDate = NewStartDate.AddDays(Me.RepeatEvery * 7)
                NewEndDate = NewEndDate.AddDays(Me.RepeatEvery * 7)
            Case DurationUnits.Year
                NewStartDate = NewStartDate.AddYears(Me.RepeatEvery)
                NewEndDate = NewEndDate.AddYears(Me.RepeatEvery)
        End Select
        RepeatStartDates.Add(NewStartDate)
        RepeatEndDates.Add(NewEndDate)
    End Sub
#End Region

#Region "Basic properties"
    Public Property idActivityDetail As Integer
        Get
            Return intIdActivityDetail
        End Get
        Set(ByVal value As Integer)
            intIdActivityDetail = value
        End Set
    End Property

    Public Property idStruct As Integer
        Get
            Return intIdStruct
        End Get
        Set(ByVal value As Integer)
            intIdStruct = value
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

    Public Property Type() As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property TypeName() As String
        Get
            If LIST_ActivityTypes(Type) IsNot Nothing Then
                Return LIST_ActivityTypes(Type).Value
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property Organisation() As String
        Get
            Return strOrganisation
        End Get
        Set(ByVal value As String)
            strOrganisation = value
        End Set
    End Property

    Public Property Location() As String
        Get
            Return strLocation
        End Get
        Set(ByVal value As String)
            strLocation = value
        End Set
    End Property
#End Region

#Region "Start and end dates"
    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property StartDateMainActivity() As Date
        Get
            If Relative = False Then
                Return StartDate
            Else
                Dim RefDate As Date
                Dim selDate As Date
                Dim sngPeriod As Single
                If PeriodDirection = ActivityDetail.DirectionValues.After Then
                    sngPeriod = Period
                    RefDate = CurrentLogFrame.GetReferenceDateByGuid(GuidReferenceMoment)
                ElseIf PeriodDirection = ActivityDetail.DirectionValues.Before Then
                    sngPeriod = Period * -1
                    RefDate = CurrentLogFrame.GetReferenceDateByGuid(GuidReferenceMoment, True)
                    If RefDate = Date.MinValue Then sngPeriod *= -1
                Else
                    Return selDate
                End If

                If RefDate = Date.MinValue And sngPeriod < 0 Then
                    sngPeriod = 0
                End If
                Select Case PeriodUnit
                    Case DurationUnits.Day
                        selDate = RefDate.AddDays(sngPeriod)
                    Case DurationUnits.Hour
                        selDate = RefDate.AddHours(sngPeriod)
                    Case DurationUnits.Minute
                        selDate = RefDate.AddMinutes(sngPeriod)
                    Case DurationUnits.Month
                        selDate = RefDate.AddMonths(sngPeriod)
                    Case DurationUnits.Semester
                        selDate = RefDate.AddMonths(sngPeriod * 6)
                    Case DurationUnits.Trimester
                        selDate = RefDate.AddMonths(sngPeriod * 3)
                    Case DurationUnits.Week
                        selDate = RefDate.AddDays(sngPeriod * 7)
                    Case DurationUnits.Year
                        selDate = RefDate.AddYears(sngPeriod)
                End Select

                'set time to midnight
                selDate = DateSerial(selDate.Year, selDate.Month, selDate.Day)
                Return selDate
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property EndDateMainActivity() As Date
        Get
            Dim selDate As Date = StartDateMainActivity

            If selDate = Date.MinValue And Duration < 0 Then Return selDate

            Select Case DurationUnit
                Case DurationUnits.Day
                    selDate = selDate.AddDays(Duration)
                Case DurationUnits.Hour
                    selDate = selDate.AddHours(Duration)
                Case DurationUnits.Minute
                    selDate = selDate.AddMinutes(Duration)
                Case DurationUnits.Month
                    selDate = selDate.AddMonths(Duration)
                Case DurationUnits.Semester
                    selDate = selDate.AddMonths(Duration * 6)
                Case DurationUnits.Trimester
                    selDate = selDate.AddMonths(Duration * 3)
                Case DurationUnits.Week
                    selDate = selDate.AddDays(Duration * 7)
                Case DurationUnits.Year
                    selDate = selDate.AddYears(Duration)
            End Select
            If DurationUnit <> DurationUnits.NotDefined And selDate > Date.MinValue Then selDate = selDate.AddDays(-1)
            Return selDate
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property StartDateRepeats() As Date
        Get
            Return StartDateMainActivity
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property EndDateRepeats() As Date
        Get
            Dim datEnd As Date = EndDateMainActivity

            If RepeatEndDates.Count > 0 Then
                For Each selDate As Date In RepeatEndDates
                    If selDate > datEnd Then datEnd = selDate
                Next
            End If
            
            Return datEnd
        End Get
    End Property

    Public Property Relative() As Boolean
        Get
            Return boolRelative
        End Get
        Set(ByVal value As Boolean)
            boolRelative = value
            If boolRelative = True Then
                datStartDate = Nothing
            Else
                sngPeriod = 0
                intPeriodUnit = 0
                intPeriodDirection = 0
            End If
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return datStartDate
        End Get
        Set(ByVal value As Date)
            datStartDate = DateSerial(value.Year, value.Month, value.Day)
        End Set
    End Property

    Public Property Period() As Single
        Get
            Return sngPeriod
        End Get
        Set(ByVal value As Single)
            sngPeriod = value
        End Set
    End Property

    Public Property PeriodUnit() As Integer
        Get
            Return intPeriodUnit
        End Get
        Set(ByVal value As Integer)
            intPeriodUnit = value
        End Set
    End Property

    Public Property PeriodDirection() As Integer
        Get
            Return intPeriodDirection
        End Get
        Set(ByVal value As Integer)
            intPeriodDirection = value
        End Set
    End Property

    Public Property GuidReferenceMoment() As Guid
        Get
            Return objReferenceMoment
        End Get
        Set(ByVal value As Guid)
            objReferenceMoment = value
        End Set
    End Property

    Public Property Duration() As Integer
        Get
            If Me.DurationUntilEnd = True Then
                Dim intPeriod As Integer
                Dim datEnd As Date = CurrentLogFrame.EndDate
                Select Case DurationUnit
                    Case DurationUnits.Day
                        intPeriod = DateDiff(DateInterval.Day, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                    Case DurationUnits.Hour
                        intPeriod = DateDiff(DateInterval.Hour, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                    Case DurationUnits.Minute
                        intPeriod = DateDiff(DateInterval.Minute, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                    Case DurationUnits.Month
                        datEnd = datEnd.AddDays(1)
                        intPeriod = DateDiff(DateInterval.Month, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                    Case DurationUnits.Semester
                        datEnd = datEnd.AddDays(1)
                        intPeriod = DateDiff(DateInterval.Year, StartDateMainActivity, datEnd, FirstDayOfWeek.System) * 2
                    Case DurationUnits.Trimester
                        datEnd = datEnd.AddDays(1)
                        intPeriod = DateDiff(DateInterval.Quarter, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                    Case DurationUnits.Week
                        intPeriod = DateDiff(DateInterval.WeekOfYear, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                    Case DurationUnits.Year
                        datEnd = datEnd.AddDays(1)
                        intPeriod = DateDiff(DateInterval.Year, StartDateMainActivity, datEnd, FirstDayOfWeek.System)
                End Select
                intDuration = intPeriod
            End If
            Return intDuration
        End Get
        Set(ByVal value As Integer)
            intDuration = value
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

    Public Property DurationUntilEnd() As Boolean
        Get
            Return boolDurationUntilEnd
        End Get
        Set(ByVal value As Boolean)
            boolDurationUntilEnd = value
            If value = True And intDurationUnit = 0 Then
                intDurationUnit = DurationUnits.Week
            End If
        End Set
    End Property
#End Region 'Start and end dates

#Region "Repeat"
    Public Property RepeatEvery() As Single
        Get
            Return sngRepeatEvery
        End Get
        Set(ByVal value As Single)
            sngRepeatEvery = value
            'CalculateRepeatDates()
        End Set
    End Property

    Public Property RepeatUnit() As Integer
        Get
            Return intRepeatUnit
        End Get
        Set(ByVal value As Integer)
            intRepeatUnit = value
            'CalculateRepeatDates()
        End Set
    End Property

    Public Property RepeatTimes() As Integer
        Get
            Return intRepeatTimes
        End Get
        Set(ByVal value As Integer)
            intRepeatTimes = value
            'CalculateRepeatDates()
        End Set
    End Property

    Public Property RepeatUntilEnd() As Boolean
        Get
            Return boolRepeatUntilEnd
        End Get
        Set(ByVal value As Boolean)
            boolRepeatUntilEnd = value
            'CalculateRepeatDates()
        End Set
    End Property

    Public Property RepeatStartDates() As List(Of Date)
        Get
            Return lstRepeatStartDates
        End Get
        Set(ByVal value As List(Of Date))
            lstRepeatStartDates = value
        End Set
    End Property

    Public Property RepeatEndDates() As List(Of Date)
        Get
            Return lstRepeatEndDates
        End Get
        Set(ByVal value As List(Of Date))
            lstRepeatEndDates = value
        End Set
    End Property
#End Region 'Repeat

#Region "Preparation"
    Public Property Preparation() As Integer
        Get
            If Me.PreparationFromStart = True Then
                Dim intPeriod As Integer
                Dim datProjectStart As Date = CurrentLogFrame.StartDate
                Dim datStartDate As Date = StartDateMainActivity
                Select Case PreparationUnit
                    Case DurationUnits.Day
                        intPeriod = DateDiff(DateInterval.DayOfYear, datProjectStart, StartDateMainActivity)
                    Case DurationUnits.Hour
                        intPeriod = DateDiff(DateInterval.Hour, datProjectStart, StartDateMainActivity)
                    Case DurationUnits.Minute
                        intPeriod = DateDiff(DateInterval.Minute, datProjectStart, StartDateMainActivity)
                    Case DurationUnits.Month
                        intPeriod = DateDiff(DateInterval.Month, datProjectStart, StartDateMainActivity)
                    Case DurationUnits.Semester
                        intPeriod = DateDiff(DateInterval.Year, datProjectStart, StartDateMainActivity) * 2
                    Case DurationUnits.Trimester
                        intPeriod = DateDiff(DateInterval.Quarter, datProjectStart, StartDateMainActivity)
                    Case DurationUnits.Week
                        intPeriod = DateDiff(DateInterval.WeekOfYear, datProjectStart, StartDateMainActivity)
                    Case DurationUnits.Year
                        intPeriod = DateDiff(DateInterval.Year, datProjectStart, StartDateMainActivity)
                End Select
                intPreparation = intPeriod
            End If
            Return intPreparation
        End Get
        Set(ByVal value As Integer)
            intPreparation = value
        End Set
    End Property

    Public Property PreparationUnit() As Integer
        Get
            Return intPreparationUnit
        End Get
        Set(ByVal value As Integer)
            intPreparationUnit = value
        End Set
    End Property

    Public Property PreparationFromStart() As Boolean
        Get
            Return boolPreparationFromStart
        End Get
        Set(ByVal value As Boolean)
            boolPreparationFromStart = value
        End Set
    End Property

    Public Property PreparationRepeat() As Boolean
        Get
            Return boolPreparationRepeat
        End Get
        Set(ByVal value As Boolean)
            boolPreparationRepeat = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property StartDatePreparation() As Date
        Get
            Dim datPrepDate As Date = StartDateMainActivity
            If Me.Preparation > 0 And Me.PreparationUnit > 0 And datPrepDate > Date.MinValue Then
                Dim sngPrepDuration As Single = Me.Preparation * -1
                Select Case PreparationUnit
                    Case DurationUnits.Day
                        datPrepDate = datPrepDate.AddDays(sngPrepDuration)
                    Case DurationUnits.Hour
                        datPrepDate = datPrepDate.AddHours(sngPrepDuration)
                    Case DurationUnits.Minute
                        datPrepDate = datPrepDate.AddMinutes(sngPrepDuration)
                    Case DurationUnits.Month
                        datPrepDate = datPrepDate.AddMonths(sngPrepDuration)
                    Case DurationUnits.Semester
                        datPrepDate = datPrepDate.AddMonths(sngPrepDuration * 6)
                    Case DurationUnits.Trimester
                        datPrepDate = datPrepDate.AddMonths(sngPrepDuration * 3)
                    Case DurationUnits.Week
                        datPrepDate = datPrepDate.AddDays(sngPrepDuration * 7)
                    Case DurationUnits.Year
                        datPrepDate = datPrepDate.AddYears(sngPrepDuration)
                End Select
            End If
            Return datPrepDate
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property PreparationPeriod() As TimeSpan
        Get
            Dim tsPreparation As New TimeSpan
            tsPreparation = Me.StartDateMainActivity.Subtract(StartDatePreparation)

            Return tsPreparation
        End Get
    End Property
#End Region 'Preparation

#Region "Follow up"
    Public Property FollowUp() As Integer
        Get
            If Me.FollowUpUntilEnd = True Then
                Dim intPeriod As Integer
                Dim datEnd As Date = CurrentLogFrame.EndDate
                Select Case FollowUpUnit
                    Case DurationUnits.Day
                        intPeriod = DateDiff(DateInterval.DayOfYear, EndDateMainActivity, datEnd)
                    Case DurationUnits.Hour
                        intPeriod = DateDiff(DateInterval.Hour, EndDateMainActivity, datEnd)
                    Case DurationUnits.Minute
                        intPeriod = DateDiff(DateInterval.Minute, EndDateMainActivity, datEnd)
                    Case DurationUnits.Month
                        intPeriod = DateDiff(DateInterval.Month, EndDateMainActivity, datEnd)
                    Case DurationUnits.Semester
                        intPeriod = DateDiff(DateInterval.Year, EndDateMainActivity, datEnd) * 2
                    Case DurationUnits.Trimester
                        intPeriod = DateDiff(DateInterval.Quarter, EndDateMainActivity, datEnd)
                    Case DurationUnits.Week
                        intPeriod = DateDiff(DateInterval.WeekOfYear, EndDateMainActivity, datEnd)
                    Case DurationUnits.Year
                        intPeriod = DateDiff(DateInterval.Year, EndDateMainActivity, datEnd)
                End Select
                intFollowUp = intPeriod
            End If
            Return intFollowUp
        End Get
        Set(ByVal value As Integer)
            intFollowUp = value
        End Set
    End Property

    Public Property FollowUpUnit() As Integer
        Get
            Return intFollowUpUnit
        End Get
        Set(ByVal value As Integer)
            intFollowUpUnit = value
        End Set
    End Property

    Public Property FollowUpUntilEnd() As Boolean
        Get
            Return boolFollowUpUntilEnd
        End Get
        Set(ByVal value As Boolean)
            boolFollowUpUntilEnd = value
        End Set
    End Property

    Public Property FollowUpRepeat() As Boolean
        Get
            Return boolFollowUpRepeat
        End Get
        Set(ByVal value As Boolean)
            boolFollowUpRepeat = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property EndDateFollowUp() As Date
        Get
            Dim datFollowUpDate As Date = EndDateMainActivity
            If Me.FollowUp > 0 And Me.FollowUpUnit > 0 Then
                Dim intFollowUpDuration As Integer = Me.FollowUp
                Select Case FollowUpUnit
                    Case DurationUnits.Day
                        datFollowUpDate = datFollowUpDate.AddDays(intFollowUpDuration)
                    Case DurationUnits.Hour
                        datFollowUpDate = datFollowUpDate.AddHours(intFollowUpDuration)
                    Case DurationUnits.Minute
                        datFollowUpDate = datFollowUpDate.AddMinutes(intFollowUpDuration)
                    Case DurationUnits.Month
                        datFollowUpDate = datFollowUpDate.AddMonths(intFollowUpDuration)
                    Case DurationUnits.Semester
                        datFollowUpDate = datFollowUpDate.AddMonths(intFollowUpDuration * 6)
                    Case DurationUnits.Trimester
                        datFollowUpDate = datFollowUpDate.AddMonths(intFollowUpDuration * 3)
                    Case DurationUnits.Week
                        datFollowUpDate = datFollowUpDate.AddDays(intFollowUpDuration * 7)
                    Case DurationUnits.Year
                        datFollowUpDate = datFollowUpDate.AddYears(intFollowUpDuration)
                End Select
            End If
            Return datFollowUpDate
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property FollowUpPeriod() As TimeSpan
        Get
            Dim tsFollowUp As New TimeSpan
            tsFollowUp = Me.EndDateFollowUp.Subtract(EndDateMainActivity)

            Return tsFollowUp
        End Get
    End Property
#End Region 'Follow-up"
End Class
