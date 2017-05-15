Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class TargetDeadlinesSection
    <ScriptIgnore()> _
    Public WithEvents TargetDeadlines As New TargetDeadlines

    Private intRepetition As Integer = 1
    Private intSection As Integer = 1 'Goals section
    Private datStartDate As Date
    Private boolRelativeStart As Boolean = True
    Private sngPeriodStart As Single = 1
    Private intPeriodUnitStart As Integer = 8
    Private datEndDate As Date
    Private boolRelativeEnd As Boolean = True
    Private sngPeriodEnd As Single = 1
    Private intPeriodUnitEnd As Integer = 5
    Private objGuid As Guid

    Public Enum RepetitionOptions As Integer
        SingleTarget = 0
        YearlyTarget = 1
        TwiceYear = 2
        QuarterlyTarget = 3
        MonthlyTarget = 4
        UserSelect = 5
    End Enum

#Region "Properties"
    Public Property Repetition As Integer
        Get
            Return intRepetition
        End Get
        Set(ByVal value As Integer)
            intRepetition = value
        End Set
    End Property

    Public Property Section As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property RelativeStart() As Boolean
        Get
            Return boolRelativeStart
        End Get
        Set(ByVal value As Boolean)
            boolRelativeStart = value
            If boolRelativeStart = True Then
                datStartDate = Nothing
            Else
                If datStartDate = Date.MinValue Then datStartDate = Now.Date
                sngPeriodStart = 0
                intPeriodUnitStart = 0
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

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ExactStartDate() As Date
        Get
            If Me.RelativeStart = False Then
                Return Me.StartDate
            Else
                Dim objProjectStartKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment
                If objProjectStartKeyMoment IsNot Nothing Then
                    Dim RefDate As Date = objProjectStartKeyMoment.ExactDateKeyMoment 'CurrentLogFrame.GetReferenceDateByGuid(Me.GuidReferenceMoment)
                    Dim selDate As Date = RefDate
                    Dim sngPeriod As Single = Me.PeriodStart

                    If RefDate = Date.MinValue And sngPeriod < 0 Then Return selDate

                    Select Case Me.PeriodUnitStart
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

                    Return selDate
                Else
                    Me.RelativeStart = False
                    Return Me.StartDate
                End If
            End If
        End Get
    End Property

    Public Property PeriodStart() As Single
        Get
            Return sngPeriodStart
        End Get
        Set(ByVal value As Single)
            sngPeriodStart = value
        End Set
    End Property

    Public Property PeriodUnitStart() As Integer
        Get
            Return intPeriodUnitStart
        End Get
        Set(ByVal value As Integer)
            If value < DurationUnits.Day Then value = DurationUnits.Day
            intPeriodUnitStart = value
        End Set
    End Property

    Public ReadOnly Property PeriodDirectionStart() As Integer
        Get
            Return 2 'Directionvalues.After
        End Get
    End Property

    Public ReadOnly Property GuidReferenceMomentStart() As Guid
        Get
            Return CurrentLogFrame.GetProjectStartKeyMoment.Guid
        End Get
    End Property

    Public Property RelativeEnd() As Boolean
        Get
            Return boolRelativeEnd
        End Get
        Set(ByVal value As Boolean)
            boolRelativeEnd = value
            If boolRelativeEnd = True Then
                datEndDate = Nothing
            Else
                If datEndDate = Date.MinValue Then datEndDate = Now.Date
                sngPeriodEnd = 0
                intPeriodUnitEnd = 0
            End If
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

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ExactEndDate() As Date
        Get
            If Me.RelativeEnd = False Then
                Return Me.EndDate
            Else
                Dim objProjectEndKeyMoment As KeyMoment = CurrentLogFrame.GetProjectEndKeyMoment
                If objProjectEndKeyMoment IsNot Nothing Then
                    Dim RefDate As Date = objProjectEndKeyMoment.ExactDateKeyMoment 'CurrentLogFrame.GetReferenceDateByGuid(Me.GuidReferenceMoment)
                    Dim selDate As Date = RefDate
                    Dim sngPeriod As Single = Me.PeriodEnd


                    If RefDate = Date.MinValue And sngPeriod < 0 Then Return selDate

                    Select Case Me.PeriodUnitEnd
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

                    Return selDate
                Else
                    Me.RelativeEnd = False
                    Return Me.EndDate
                End If
            End If
        End Get
    End Property

    Public Property PeriodEnd() As Single
        Get
            Return sngPeriodEnd
        End Get
        Set(ByVal value As Single)
            sngPeriodEnd = value
        End Set
    End Property

    Public Property PeriodUnitEnd() As Integer
        Get
            Return intPeriodUnitEnd
        End Get
        Set(ByVal value As Integer)
            If value < DurationUnits.Day Then value = DurationUnits.Day
            intPeriodUnitEnd = value
        End Set
    End Property

    Public ReadOnly Property PeriodDirectionEnd() As Integer
        Get
            Return 2 'Directionvalues.After
        End Get
    End Property

    Public ReadOnly Property GuidReferenceMomentEnd() As Guid
        Get
            Return CurrentLogFrame.GetProjectEndKeyMoment.Guid
        End Get
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal section As Integer)
        Me.Section = section
    End Sub

    Public Sub VerifyStartBeforeEnd()
        If ExactEndDate < ExactStartDate Then
            If RelativeStart Then
                RelativeEnd = RelativeStart
                PeriodEnd = PeriodStart
                PeriodUnitEnd = PeriodUnitStart
            Else
                EndDate = StartDate
            End If
        End If
    End Sub

    Public Sub SetTargetDeadlines()

        TargetDeadlines.Clear()

        Select Case intRepetition
            Case RepetitionOptions.SingleTarget
                SetTargetDeadlines_SingleTarget()

            Case RepetitionOptions.YearlyTarget
                SetTargetDeadlines_YearlyTarget()

            Case RepetitionOptions.TwiceYear
                SetTargetDeadlines_TwiceYear()

            Case RepetitionOptions.QuarterlyTarget
                SetTargetDeadlines_QuarterlyTarget()

            Case RepetitionOptions.MonthlyTarget
                SetTargetDeadlines_MonthlyTarget()

        End Select
    End Sub

    Private Sub SetTargetDeadlines_SingleTarget()
        If RelativeEnd Then
            Dim objProjectEndKeyMoment As KeyMoment = CurrentLogFrame.GetProjectEndKeyMoment
            If objProjectEndKeyMoment IsNot Nothing Then
                Dim objGuidProjectEnd As Guid = objProjectEndKeyMoment.Guid
                Dim NewDeadline As New TargetDeadline(objGuidProjectEnd, PeriodEnd, PeriodUnitEnd)

                TargetDeadlines.Add(NewDeadline)
            Else
                TargetDeadlines.Add(New TargetDeadline(EndDate))
            End If
        Else
            TargetDeadlines.Add(New TargetDeadline(EndDate))
        End If
    End Sub

    Private Sub SetTargetDeadlines_YearlyTarget()
        Dim NewDeadline As TargetDeadline
        Dim intYear As Integer

        VerifyStartBeforeEnd()

        If RelativeStart Then
            Dim objProjectStartKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment
            If objProjectStartKeyMoment IsNot Nothing Then
                Dim objGuidProjectStart As Guid = objProjectStartKeyMoment.Guid
                Dim selPeriod As Single

                For i = ExactStartDate.Year To ExactEndDate.Year
                    Select Case Me.PeriodUnitStart
                        Case DurationUnits.Day
                            Dim intDaysInYear As Integer = 365
                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            selPeriod = Me.PeriodStart + (intYear * intDaysInYear)
                        Case DurationUnits.Week
                            Dim intDaysInYear As Integer = 365
                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            selPeriod = Me.PeriodStart + (intYear * Int(intDaysInYear / 7))
                        Case DurationUnits.Month
                            selPeriod = Me.PeriodStart + (intYear * 12)
                        Case DurationUnits.Trimester
                            selPeriod = Me.PeriodStart + (intYear * 4)
                        Case DurationUnits.Semester
                            selPeriod = Me.PeriodStart + (intYear * 2)
                        Case DurationUnits.Year
                            selPeriod = Me.PeriodStart + intYear
                    End Select

                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod, Me.PeriodUnitStart)
                    If NewDeadline.ExactDeadline <= ExactEndDate Then _
                        TargetDeadlines.Add(NewDeadline)
                    intYear += 1
                Next
            End If
        Else
            For i = ExactStartDate.Year To ExactEndDate.Year
                NewDeadline = New TargetDeadline(DateSerial(i, ExactEndDate.Month, Date.DaysInMonth(ExactEndDate.Year, ExactEndDate.Month)))
                If NewDeadline.ExactDeadline <= ExactEndDate Then _
                    TargetDeadlines.Add(NewDeadline)
            Next
        End If
    End Sub

    Private Sub SetTargetDeadlines_TwiceYear()
        Dim NewDeadline As TargetDeadline
        Dim intYear As Integer

        If RelativeStart Then
            Dim objProjectStartKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment
            If objProjectStartKeyMoment IsNot Nothing Then
                Dim objGuidProjectStart As Guid = objProjectStartKeyMoment.Guid
                Dim selPeriod(1) As Single
                Dim selPeriodUnit(1) As Integer

                selPeriodUnit(0) = Me.PeriodUnitStart

                For i = ExactStartDate.Year To ExactEndDate.Year
                    selPeriodUnit(1) = Me.PeriodUnitStart

                    Select Case Me.PeriodUnitStart
                        Case DurationUnits.Day
                            Dim intDaysInYear As Integer = 365
                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            selPeriod(0) = Me.PeriodStart + (intYear * intDaysInYear)
                            selPeriod(1) = Me.PeriodStart + (intYear * intDaysInYear) + Int(intDaysInYear / 2)
                        Case DurationUnits.Week
                            Dim intDaysInYear As Integer = 365
                            Dim intWeeks As Integer

                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            intWeeks = intDaysInYear / 7
                            selPeriod(0) = Me.PeriodStart + (intYear * intWeeks)
                            selPeriod(1) = Me.PeriodStart + (intYear * intWeeks) + (intWeeks / 2)
                        Case DurationUnits.Month
                            selPeriod(0) = Me.PeriodStart + (intYear * 12)
                            selPeriod(1) = Me.PeriodStart + (intYear * 12) + 6
                        Case DurationUnits.Trimester
                            selPeriod(0) = Me.PeriodStart + (intYear * 4)
                            selPeriod(1) = Me.PeriodStart + (intYear * 4) + 2
                        Case DurationUnits.Semester
                            selPeriod(0) = Me.PeriodStart + (intYear * 2)
                            selPeriod(1) = Me.PeriodStart + (intYear * 2) + 1
                        Case DurationUnits.Year
                            selPeriod(0) = Me.PeriodStart + intYear
                            selPeriod(1) = (Me.PeriodStart * 2) + (intYear * 2) + 1
                            selPeriodUnit(1) = DurationUnits.Semester
                    End Select

                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(0), selPeriodUnit(0))
                    TargetDeadlines.Add(NewDeadline)
                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(1), selPeriodUnit(1))
                    TargetDeadlines.Add(NewDeadline)
                    intYear += 1
                Next
            End If
        Else
            For i = ExactStartDate.Year To ExactEndDate.Year
                TargetDeadlines.Add(New TargetDeadline(DateSerial(i, 6, 30)))
                TargetDeadlines.Add(New TargetDeadline(DateSerial(i, 12, 31)))
            Next
        End If
    End Sub

    Private Sub SetTargetDeadlines_QuarterlyTarget()
        Dim NewDeadline As TargetDeadline
        Dim intYear As Integer

        If RelativeStart Then
            Dim objProjectStartKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment
            If objProjectStartKeyMoment IsNot Nothing Then
                Dim objGuidProjectStart As Guid = objProjectStartKeyMoment.Guid
                Dim selPeriod(3) As Single
                Dim selPeriodUnit(3) As Integer

                selPeriodUnit(0) = Me.PeriodUnitStart

                For i = ExactStartDate.Year To ExactEndDate.Year
                    selPeriodUnit(1) = Me.PeriodUnitStart
                    selPeriodUnit(2) = Me.PeriodUnitStart
                    selPeriodUnit(3) = Me.PeriodUnitStart

                    Select Case Me.PeriodUnitStart
                        Case DurationUnits.Day
                            Dim intDaysInYear As Integer = 365
                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            selPeriod(0) = Me.PeriodStart + (intYear * intDaysInYear)
                            selPeriod(1) = Me.PeriodStart + (intYear * intDaysInYear) + Int(intDaysInYear / 4)
                            selPeriod(2) = Me.PeriodStart + (intYear * intDaysInYear) + Int(intDaysInYear / 2)
                            selPeriod(3) = Me.PeriodStart + (intYear * intDaysInYear) + Int(intDaysInYear * 0.75)
                        Case DurationUnits.Week
                            Dim intDaysInYear As Integer = 365
                            Dim intWeeks As Integer

                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            intWeeks = intDaysInYear / 7
                            selPeriod(0) = Me.PeriodStart + (intYear * intWeeks)
                            selPeriod(1) = Me.PeriodStart + (intYear * intWeeks) + Int(intWeeks / 4)
                            selPeriod(2) = Me.PeriodStart + (intYear * intWeeks) + Int(intWeeks / 2)
                            selPeriod(3) = Me.PeriodStart + (intYear * intWeeks) + Int(intWeeks * 0.75)
                        Case DurationUnits.Month
                            selPeriod(0) = Me.PeriodStart + (intYear * 12)
                            selPeriod(1) = Me.PeriodStart + (intYear * 12) + 3
                            selPeriod(2) = Me.PeriodStart + (intYear * 12) + 6
                            selPeriod(3) = Me.PeriodStart + (intYear * 12) + 9
                        Case DurationUnits.Trimester
                            selPeriod(0) = Me.PeriodStart + (intYear * 4)
                            selPeriod(1) = Me.PeriodStart + (intYear * 4) + 1
                            selPeriod(2) = Me.PeriodStart + (intYear * 4) + 2
                            selPeriod(3) = Me.PeriodStart + (intYear * 4) + 3
                        Case DurationUnits.Semester
                            selPeriod(0) = Me.PeriodStart + (intYear * 2)
                            selPeriod(1) = (Me.PeriodStart * 2) + (intYear * 4) + 1
                            selPeriod(2) = (Me.PeriodStart * 2) + (intYear * 4) + 2
                            selPeriod(3) = (Me.PeriodStart * 2) + (intYear * 4) + 3
                            selPeriodUnit(1) = DurationUnits.Trimester
                            selPeriodUnit(2) = DurationUnits.Trimester
                            selPeriodUnit(3) = DurationUnits.Trimester
                        Case DurationUnits.Year
                            selPeriod(0) = Me.PeriodStart + (intYear * 2)
                            selPeriod(1) = (Me.PeriodStart * 4) + (intYear * 4) + 1
                            selPeriod(2) = (Me.PeriodStart * 4) + (intYear * 4) + 2
                            selPeriod(3) = (Me.PeriodStart * 4) + (intYear * 4) + 3
                            selPeriodUnit(1) = DurationUnits.Trimester
                            selPeriodUnit(2) = DurationUnits.Trimester
                            selPeriodUnit(3) = DurationUnits.Trimester
                    End Select

                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(0), selPeriodUnit(0))
                    TargetDeadlines.Add(NewDeadline)
                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(1), selPeriodUnit(1))
                    TargetDeadlines.Add(NewDeadline)
                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(2), selPeriodUnit(2))
                    TargetDeadlines.Add(NewDeadline)
                    NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(3), selPeriodUnit(3))
                    TargetDeadlines.Add(NewDeadline)
                    intYear += 1
                Next
            End If
        Else
            For i = ExactStartDate.Year To ExactEndDate.Year
                TargetDeadlines.Add(New TargetDeadline(DateSerial(i, 3, 31)))
                TargetDeadlines.Add(New TargetDeadline(DateSerial(i, 6, 30)))
                TargetDeadlines.Add(New TargetDeadline(DateSerial(i, 9, 30)))
                TargetDeadlines.Add(New TargetDeadline(DateSerial(i, 12, 31)))
            Next
        End If
    End Sub

    Private Sub SetTargetDeadlines_MonthlyTarget()
        Dim NewDeadline As TargetDeadline
        Dim intYear As Integer

        If RelativeStart Then
            Dim objProjectStartKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment
            If objProjectStartKeyMoment IsNot Nothing Then
                Dim objGuidProjectStart As Guid = objProjectStartKeyMoment.Guid
                Dim selPeriod(12) As Single
                Dim selPeriodUnit(12) As Integer

                selPeriodUnit(1) = Me.PeriodUnitStart

                For i = ExactStartDate.Year To ExactEndDate.Year
                    For j = 2 To 12
                        selPeriodUnit(j) = Me.PeriodUnitStart
                    Next

                    Select Case Me.PeriodUnitStart
                        Case DurationUnits.Day
                            Dim intDaysInYear As Integer = 365
                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            For j = 1 To 12
                                selPeriod(j) = Me.PeriodStart + (intYear * intDaysInYear) + (Int(intDaysInYear / 12) * (j - 1))
                            Next
                        Case DurationUnits.Week
                            Dim intDaysInYear As Integer = 365
                            Dim intWeeks As Integer

                            If Date.IsLeapYear(i) Then intDaysInYear = 366
                            intWeeks = intDaysInYear / 7
                            For j = 1 To 12
                                selPeriod(j) = Me.PeriodStart + (intYear * intWeeks) + (Int(intWeeks / 12) * (j - 1))
                            Next
                        Case DurationUnits.Month
                            For j = 1 To 12
                                selPeriod(j) = Me.PeriodStart + (intYear * 12) + j - 1
                            Next
                        Case DurationUnits.Trimester
                            selPeriod(1) = Me.PeriodStart + (intYear * 4)
                            For j = 2 To 12
                                selPeriod(j) = (Me.PeriodStart * 3) + (intYear * 12) + j - 1
                            Next
                            For j = 2 To 12
                                selPeriodUnit(j) = DurationUnits.Month
                            Next
                        Case DurationUnits.Semester
                            selPeriod(1) = Me.PeriodStart + (intYear * 2)
                            For j = 2 To 12
                                selPeriod(j) = (Me.PeriodStart * 6) + (intYear * 12) + j - 1
                            Next
                            For j = 2 To 12
                                selPeriodUnit(j) = DurationUnits.Month
                            Next
                        Case DurationUnits.Year
                            selPeriod(1) = Me.PeriodStart + intYear
                            For j = 2 To 12
                                selPeriod(j) = (Me.PeriodStart * 12) + (intYear * 12) + j - 1
                            Next
                            For j = 2 To 12
                                selPeriodUnit(j) = DurationUnits.Month
                            Next
                    End Select

                    For j = 1 To 12
                        NewDeadline = New TargetDeadline(objGuidProjectStart, selPeriod(j), selPeriodUnit(j))
                        TargetDeadlines.Add(NewDeadline)
                    Next

                    intYear += 1
                Next
            End If
        Else
            For i = ExactStartDate.Year To ExactEndDate.Year
                For j = 1 To 12
                    TargetDeadlines.Add(New TargetDeadline(DateSerial(i, j, Date.DaysInMonth(i, j))))
                Next
            Next
        End If
    End Sub

    Public Function FormatTargetDeadlineDate(ByVal selTargetDeadline As TargetDeadline) As String
        Dim strDate As String

        Select Case Me.Repetition
            Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
            Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
            Case Else
                strDate = selTargetDeadline.ExactDeadline.ToShortDateString
        End Select

        Return strDate
    End Function

    Public Function GetTargetDeadlineByGuid(ByVal objGuid As Guid) As TargetDeadline
        Return Me.GetTargetDeadlineByGuid(objGuid)
    End Function
#End Region
End Class
