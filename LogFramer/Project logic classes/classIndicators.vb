Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Indicator
    Inherits LogframeObject
    'collections
    <ScriptIgnore()> _
    Public WithEvents Indicators As New Indicators

    <ScriptIgnore()> _
    Public WithEvents VerificationSources As New VerificationSources

    <ScriptIgnore()> _
    Public WithEvents Baseline As New Baseline

    <ScriptIgnore()> _
    Public WithEvents Targets As New Targets

    <ScriptIgnore()> _
    Public WithEvents PopulationTargets As New PopulationTargets

    <ScriptIgnore()> _
    Public WithEvents Statements As New Statements

    <ScriptIgnore()> _
    Public WithEvents ResponseClasses As New ResponseClasses

    <ScriptIgnore()> _
    Public WithEvents OpenEndedDetail As New OpenEndedDetail

    <ScriptIgnore()> _
    Public WithEvents ValuesDetail As New ValuesDetail

    <ScriptIgnore()> _
    Public WithEvents ScalesDetail As New ScalesDetail

    'variables
    Private intIdIndicator As Integer = -1
    Private intIdParent As Integer
    Private objParentStructGuid, objParentIndicatorGuid, objTargetGroupGuid As Guid
    Private intQuestionType As Integer = 50
    Private intScoringSystem, intTargetSystem As Integer
    Private intRegistration As Integer
    Private intAggregateHorizontal, intAggregateVertical As Integer
    Private sngWeightingFactorChildren As Single = 1


#Region "Enumerations"
    Protected Friend Enum QuestionTypes As Integer
        '*** NO TARGETS ***
        'Without targets
        OpenEnded = 0
        MaxDiff = 31
        Image = 60

        'values
        AbsoluteValue = 50
        PercentageValue = 51
        Ratio = 52
        Formula = 53

        'multiple options
        YesNo = 20
        MultipleOptions = 21
        MultipleChoice = 22

        'ordinal questions
        Ranking = 30
        LikertTypeScale = 341
        SemanticDiff = 32

        'Expressing opinion
        Scale = 33
        LikertScale = 34
        CumulativeScale = 35
        FrequencyLikert = 36

        'Other types

        ImageWithTargets = 61

        'combination
        MixedSubIndicators = -1
    End Enum

    Public Enum ScoringSystems As Integer
        Value = 0
        Percentage = 1
        Score = 2
    End Enum

    Public Enum TargetSystems As Integer
        Simple = 0
        ValueRange = 1
        Formula = 2
    End Enum

    Public Enum RegistrationOptions As Integer
        ProgrammeLevel = 0
        TeamLevel = 1
        BeneficiaryLevel = 2
    End Enum

    Public Enum AggregationOptions As Integer
        Sum = 0
        Average = 1
        Minimum = 2
        Maximum = 3
        Spread = 4
        Distribution = 5
    End Enum
#End Region

#Region "Properties"
    Public Property idIndicator As Integer
        Get
            Return intIdIndicator
        End Get
        Set(ByVal value As Integer)
            intIdIndicator = value
        End Set
    End Property

    Public Property idParent As Integer
        Get
            Return intIdParent
        End Get
        Set(ByVal value As Integer)
            intIdParent = value
        End Set
    End Property

    Public Property ParentStructGuid() As Guid
        Get
            Return objParentStructGuid
        End Get
        Set(ByVal value As Guid)
            objParentStructGuid = value
            objParentIndicatorGuid = Guid.Empty
        End Set
    End Property

    Public Property ParentIndicatorGuid As Guid
        Get
            Return objParentIndicatorGuid
        End Get
        Set(ByVal value As Guid)
            objParentIndicatorGuid = value
            objParentStructGuid = Guid.Empty
        End Set
    End Property

    Public Property QuestionType() As Integer
        Get
            Return intQuestionType
        End Get
        Set(ByVal value As Integer)
            If value <> intQuestionType Then
                intQuestionType = value
            End If
        End Set
    End Property

    Public Property ScoringSystem As Integer
        Get
            Return intScoringSystem
        End Get
        Set(ByVal value As Integer)
            intScoringSystem = value

            If Me.Indicators.Count > 0 Then
                Me.Indicators.UpdateScoringSystemOfChildren(intScoringSystem)
            End If
        End Set
    End Property

    Public Property TargetSystem As Integer
        Get
            Return intTargetSystem
        End Get
        Set(ByVal value As Integer)
            intTargetSystem = value

            If intTargetSystem <> TargetSystems.Formula Then
                For Each selTarget As Target In Me.Targets
                    selTarget.Formula = String.Empty
                Next
            End If

            If Me.Indicators.Count > 0 Then
                Me.Indicators.UpdateTargetSystemOfChildren(intTargetSystem)
            End If
        End Set
    End Property

    Public Property Registration As Integer
        Get
            Return intRegistration
        End Get
        Set(ByVal value As Integer)
            intRegistration = value

            If Me.Indicators.Count > 0 Then
                Me.Indicators.UpdateRegistrationOfChildren(intRegistration)
            End If
        End Set
    End Property

    Public Property TargetGroupGuid As Guid
        Get
            Return objTargetGroupGuid
        End Get
        Set(ByVal value As Guid)
            objTargetGroupGuid = value

            If objTargetGroupGuid <> Guid.Empty And Me.Indicators.Count > 0 Then
                Me.Indicators.UpdateTargetGroupGuidOfChildren(objTargetGroupGuid)
            End If
        End Set
    End Property

    Public Property AggregateVertical As Integer
        Get
            Return intAggregateVertical
        End Get
        Set(ByVal value As Integer)
            intAggregateVertical = value
        End Set
    End Property

    Public Property AggregateHorizontal As Integer
        Get
            Return intAggregateHorizontal
        End Get
        Set(ByVal value As Integer)
            intAggregateHorizontal = value
        End Set
    End Property

    Public Property WeightingFactorChildren As Single
        Get
            Return sngWeightingFactorChildren
        End Get
        Set(ByVal value As Single)
            sngWeightingFactorChildren = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property QuestionTypeName As String
        Get
            If LIST_QuestionTypes.ContainsKey(Me.QuestionType) Then
                Return LIST_QuestionTypes(Me.QuestionType)
            Else
                Return String.Empty
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RegistrationOption As String
        Get
            Return LIST_RegistrationOptions(Me.Registration)
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property IsCleanIndicator() As Boolean
        Get
            If Me.VerificationSources.Count = 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Indicator
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Indicators
        End Get
    End Property
#End Region

#Region "Events"
    Private Sub Indicators_IndicatorAdded(ByVal sender As Object, ByVal e As IndicatorAddedEventArgs) Handles Indicators.IndicatorAdded
        Dim selIndicator As Indicator = e.Indicator

        With selIndicator
            .idParent = Me.idIndicator
            .ParentIndicatorGuid = Me.Guid
            .Section = Me.Section
            .ScoringSystem = Me.ScoringSystem
            .Registration = Me.Registration
            .TargetGroupGuid = Me.TargetGroupGuid
            .AggregateVertical = Me.AggregateVertical

            If selIndicator.Targets.Count < Me.Targets.Count Then
                selIndicator.Targets.Clear()
                For i = 0 To Me.Targets.Count - 1
                    selIndicator.Targets.Add(New Target(Me.Targets(i).TargetDeadlineGuid))
                Next
            End If

            If selIndicator.PopulationTargets.Count < Me.PopulationTargets.Count Then
                selIndicator.PopulationTargets.Clear()
                For i = 0 To Me.PopulationTargets.Count - 1
                    selIndicator.PopulationTargets.Add(New PopulationTarget(Me.PopulationTargets(i).TargetDeadlineGuid))
                Next
            End If
        End With

        If Me.QuestionType = QuestionTypes.MaxDiff Then Me.Statements.Clear()
    End Sub

    Private Sub VerificationSources_VerificationSourceAdded(ByVal sender As Object, ByVal e As VerificationSourceAddedEventArgs) Handles VerificationSources.VerificationSourceAdded
        Dim selVerificationSource As VerificationSource = e.VerificationSource

        selVerificationSource.idIndicator = Me.idIndicator
        selVerificationSource.ParentIndicatorGuid = Me.Guid
    End Sub

    Private Sub Statements_StatementAdded(ByVal sender As Object, ByVal e As StatementAddedEventArgs) Handles Statements.StatementAdded
        Dim selStatement As Statement = e.Statement

        selStatement.idIndicator = Me.idIndicator
        selStatement.ParentIndicatorGuid = Me.Guid
    End Sub

    Private Sub ResponseClasses_ResponseClassAdded(ByVal sender As Object, ByVal e As ResponseClassAddedEventArgs) Handles ResponseClasses.ResponseClassAdded
        Dim selResponseClass As ResponseClass = e.ResponseClass

        selResponseClass.idIndicator = Me.idIndicator
        selResponseClass.ParentIndicatorGuid = Me.Guid
    End Sub

    Private Sub Targets_TargetAdded(ByVal sender As Object, ByVal e As TargetAddedEventArgs) Handles Targets.TargetAdded
        Dim selTarget As Target = e.Target

        selTarget.idIndicator = Me.idIndicator
        selTarget.ParentIndicatorGuid = Me.Guid
    End Sub

    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selIndicator As Indicator In Me.Indicators
            selIndicator.ParentIndicatorGuid = Me.Guid
        Next
        For Each selVerificationSource As VerificationSource In Me.VerificationSources
            selVerificationSource.ParentIndicatorGuid = Me.Guid
        Next
        For Each selStatement As Statement In Me.Statements
            selStatement.ParentIndicatorGuid = Me.Guid
        Next
        For Each selResponseClass As ResponseClass In Me.ResponseClasses
            selResponseClass.ParentIndicatorGuid = Me.Guid
        Next
    End Sub
#End Region

#Region "General Methods"
    Public Sub New()
        Me.QuestionType = My.Settings.setDefaultIndicatorType
        Me.Baseline.ParentIndicatorGuid = Me.Guid
    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
        Me.QuestionType = My.Settings.setDefaultIndicatorType
        Me.Baseline.ParentIndicatorGuid = Me.Guid
    End Sub

    Public Function AddClassValuesForTotal() As Boolean
        Select Case QuestionType
            Case QuestionTypes.MultipleOptions, QuestionTypes.Scale, QuestionTypes.CumulativeScale
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Sub SetQuestionType_Settings()
        'if indicator is a child indicator of a different type, set it to 'mixed type'
        If Me.ParentIndicatorGuid <> Guid.Empty Then
            Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(Me.ParentIndicatorGuid)

            If ParentIndicator IsNot Nothing AndAlso ParentIndicator.QuestionType <> intQuestionType Then
                ParentIndicator.QuestionType = QuestionTypes.MixedSubIndicators
            End If
        End If

        'other settings
        Select Case intQuestionType
            Case QuestionTypes.OpenEnded
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple

                ValuesDetail = New ValuesDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.MaxDiff
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple

                With ScalesDetail
                    If String.IsNullOrEmpty(.AgreeText) Then .AgreeText = LANG_BestChoice
                    If String.IsNullOrEmpty(.DisagreeText) Then .DisagreeText = LANG_WorstChoice
                End With

                OpenEndedDetail = New OpenEndedDetail
                ValuesDetail = New ValuesDetail
                'ScalesDetail = New ScalesDetail
            Case QuestionTypes.AbsoluteValue
                ScoringSystem = ScoringSystems.Value
                TargetSystem = TargetSystems.Simple

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.PercentageValue
                ScoringSystem = ScoringSystems.Value
                TargetSystem = TargetSystems.Simple
                Me.ValuesDetail.Unit = "%"

                If Me.ScoringSystem <> ScoringSystems.Score Then Me.AggregateVertical = AggregationOptions.Average

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.Ratio
                ScoringSystem = ScoringSystems.Value
                TargetSystem = TargetSystems.Simple

                If String.IsNullOrEmpty(ValuesDetail.Formula) Then
                    With ValuesDetail
                        .Formula = "A/B"
                        .NrDecimals = 0
                    End With
                End If
            Case QuestionTypes.Formula
                ScoringSystem = ScoringSystems.Value
                TargetSystem = TargetSystems.Simple
                If String.IsNullOrEmpty(ValuesDetail.Formula) Then ValuesDetail.Formula = "A+B"

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.YesNo, QuestionTypes.MultipleOptions, QuestionTypes.MultipleChoice, QuestionTypes.Ranking
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.LikertTypeScale
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.Scale, QuestionTypes.CumulativeScale
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple
                Registration = RegistrationOptions.BeneficiaryLevel

                With ScalesDetail
                    If String.IsNullOrEmpty(.AgreeText) Then .AgreeText = LANG_Agree
                    If String.IsNullOrEmpty(.DisagreeText) Then .DisagreeText = LANG_Disagree
                End With

                OpenEndedDetail = New OpenEndedDetail
                ValuesDetail = New ValuesDetail
            Case QuestionTypes.LikertScale, QuestionTypes.SemanticDiff
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple
                Registration = RegistrationOptions.BeneficiaryLevel

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.FrequencyLikert
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple

                OpenEndedDetail = New OpenEndedDetail
                ScalesDetail = New ScalesDetail
            Case QuestionTypes.MixedSubIndicators
                ScoringSystem = ScoringSystems.Score
                TargetSystem = TargetSystems.Simple

        End Select
    End Sub

    Public Sub SetQuestionType_Statements()
        'set number of statements
        Select Case intQuestionType
            Case QuestionTypes.OpenEnded, QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, _
                QuestionTypes.MultipleChoice, QuestionTypes.MultipleOptions, QuestionTypes.YesNo, _
                QuestionTypes.Ranking

                Statements.Clear()
            Case QuestionTypes.Formula
                If Me.Statements.Count = 0 Then SetQuestionType_AddStatements(2)
            Case QuestionTypes.Ratio
                If Me.Statements.Count = 0 Then SetQuestionType_AddStatements(2)
            Case QuestionTypes.CumulativeScale, QuestionTypes.Scale, QuestionTypes.LikertTypeScale, QuestionTypes.LikertScale, QuestionTypes.MaxDiff
                If Me.Statements.Count = 0 Then SetQuestionType_AddStatements(5)
            Case QuestionTypes.SemanticDiff
                Statements.Clear()
        End Select
    End Sub

    Private Sub SetQuestionType_AddStatements(ByVal intCount As Integer)
        Statements.Clear()
        For i = 1 To intCount
            Dim strStatement As String = String.Format("{0} {1}", LANG_Statement, i)
            Statements.Add(New Statement(TextToRichText(strStatement)))
        Next
    End Sub

    Public Sub SetQuestionType_ResponseClasses()
        ResponseClasses.Clear()

        Select Case intQuestionType
            Case QuestionTypes.YesNo
                ResponseClasses.Add(New ResponseClass(LANG_Yes, 1))
                ResponseClasses.Add(New ResponseClass(LANG_No, 0))
            Case QuestionTypes.MultipleChoice
                For i = 1 To 3
                    ResponseClasses.Add(New ResponseClass(String.Format("{0} {1}", LANG_Option, i), i - 1))
                Next
            Case QuestionTypes.MultipleOptions
                For i = 1 To 3
                    ResponseClasses.Add(New ResponseClass(String.Format("{0} {1}", LANG_Option, i), 1))
                Next
            Case QuestionTypes.Ranking
                For i = 1 To 3
                    ResponseClasses.Add(New ResponseClass(String.Format("{0} {1}", LANG_Option, i), 0))
                Next
            Case QuestionTypes.Scale
                ValuesDetail.NrDecimals = 2

                For i = 1 To 5
                    ResponseClasses.Add(New ResponseClass(LANG_Agree, i))
                Next
            Case QuestionTypes.LikertTypeScale, QuestionTypes.LikertScale
                ResponseClasses.Add(New ResponseClass(LANG_DisagreeStrongly, 1))
                ResponseClasses.Add(New ResponseClass(LANG_Disagree, 2))
                ResponseClasses.Add(New ResponseClass(LANG_Neither, 3))
                ResponseClasses.Add(New ResponseClass(LANG_Agree, 4))
                ResponseClasses.Add(New ResponseClass(LANG_AgreeStrongly, 5))
            Case QuestionTypes.CumulativeScale
                ResponseClasses.Add(New ResponseClass(LANG_Agree, 1))
            Case QuestionTypes.MaxDiff
                ResponseClasses.Add(New ResponseClass(LANG_Agree, 100))
            Case QuestionTypes.SemanticDiff
                ResponseClasses.Add(New ResponseClass(LANG_Extremely, 1))
                ResponseClasses.Add(New ResponseClass(LANG_Quite, 2))
                ResponseClasses.Add(New ResponseClass(LANG_Slightly, 3))
                ResponseClasses.Add(New ResponseClass(LANG_Neutral, 4))
                ResponseClasses.Add(New ResponseClass(LANG_Slightly, 5))
                ResponseClasses.Add(New ResponseClass(LANG_Quite, 6))
                ResponseClasses.Add(New ResponseClass(LANG_Extremely, 7))
        End Select
    End Sub

    Public Sub SetQuestionType_ResetTargets()
        Select Case intQuestionType
            Case QuestionTypes.OpenEnded, QuestionTypes.MaxDiff, QuestionTypes.Image, QuestionTypes.MixedSubIndicators
                Baseline = New Baseline
                Targets.Clear()
                PopulationTargets.Clear()
            Case Else
                For Each selTarget As Target In Me.Targets
                    selTarget.OpMin = CONST_LargerThanOrEqual
                    selTarget.MinValue = 0
                    selTarget.OpMax = String.Empty
                    selTarget.MaxValue = 0
                    selTarget.Score = 0
                Next
                For Each selPopulationTarget As PopulationTarget In Me.PopulationTargets
                    selPopulationTarget.Percentage = 0
                Next
        End Select
    End Sub
#End Region

#Region "Update child indicators"
    Public Sub UpdateUnitsOfChildren(ByVal strValueName, ByVal intNrDecimals, ByVal strUnit)
        Me.ValuesDetail = New ValuesDetail(strValueName, intNrDecimals, strUnit)

        If Me.Indicators.Count > 0 Then
            Me.Indicators.UpdateUnitsOfChildren(strValueName, intNrDecimals, strUnit)
        End If
    End Sub
#End Region

#Region "Baseline methods"

    Public Function GetBaselineTotalValue() As Double
        Dim dblBaseline As Double

        If Me.Indicators.Count = 0 Then
            dblBaseline = Me.Baseline.Value
        Else
            dblBaseline = Indicators.GetBaselineTotalValue(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblBaseline
    End Function

    Public Function GetBaselineTotalScore() As Double
        Dim dblBaseline As Double

        If Me.Indicators.Count = 0 Then
            dblBaseline = Me.Baseline.Score
        Else
            dblBaseline = Indicators.GetBaselineTotalScore(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblBaseline
    End Function

    Public Function GetBaselineFormattedValue(Optional ByVal intResponseIndex As Integer = -1) As String
        Dim strFormattedValue As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim dblBaseline As Double
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                dblBaseline = GetBaselineTotalValue()

                strFormattedValue = DisplayAsUnit(dblBaseline, intNrDecimals, strUnit)
        End Select

        Return strFormattedValue
    End Function

    Public Function GetBaselineFormattedScore() As String
        Dim strFormattedScore As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim dblBaselineScore As Double
        Dim dblMaximumScore As Double
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                Select Case ScoringSystem
                    Case ScoringSystems.Value, ScoringSystems.Percentage
                        strFormattedScore = GetBaselineFormattedValue()
                    Case ScoringSystems.Score
                        dblBaselineScore = GetBaselineTotalScore()
                        strFormattedScore = DisplayAsUnit(dblBaselineScore, intNrDecimals, String.Empty)

                        dblMaximumScore = GetBaselineMaximumScore()
                        strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                        strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
                End Select
            Case Else

                dblBaselineScore = GetBaselineTotalScore()
                strFormattedScore = DisplayAsUnit(dblBaselineScore, intNrDecimals, String.Empty)

                dblMaximumScore = GetBaselineMaximumScore()
                strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
        End Select

        Return strFormattedScore
    End Function

    Public Function GetBaselineMaximumScore() As Double
        Dim dblMaxScore As Double

        If Me.Indicators.Count = 0 Then
            Select Case Me.QuestionType
                Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                    dblMaxScore = Me.Targets.GetHighestScoreOfTargets
                Case QuestionTypes.MultipleChoice, QuestionTypes.MultipleOptions, QuestionTypes.YesNo
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                Case QuestionTypes.Ranking
                    dblMaxScore = ResponseClasses.Count
                Case QuestionTypes.Scale
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                    dblMaxScore /= ResponseClasses.Count
                Case QuestionTypes.CumulativeScale
                    dblMaxScore = ResponseClasses(0).Value * Statements.Count
                Case QuestionTypes.LikertTypeScale, QuestionTypes.FrequencyLikert
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                Case QuestionTypes.SemanticDiff, QuestionTypes.LikertScale
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                    dblMaxScore *= Statements.Count
            End Select

            dblMaxScore *= Me.WeightingFactorChildren
        Else
            dblMaxScore = Me.Indicators.GetBaselineMaximumScore(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblMaxScore
    End Function
#End Region

#Region "Targets methods"
    Public Function GetTargetValuesTotal() As Targets
        Dim objTargetsTotal As Targets
        Dim dblMinValue, dblMaxValue As Double
        Dim strOpMin, strOpMax As String
        Dim intTargetGroupSize As Integer

        If Me.Indicators.Count = 0 Then
            If Me.Registration = RegistrationOptions.BeneficiaryLevel Then
                Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
                If selTargetGroup IsNot Nothing Then intTargetGroupSize = selTargetGroup.Number
            End If

            objTargetsTotal = New Targets
            For Each selTarget As Target In Me.Targets
                strOpMin = selTarget.OpMin
                strOpMax = selTarget.OpMax
                dblMinValue = selTarget.MinValue
                dblMaxValue = selTarget.MaxValue

                dblMinValue *= Me.WeightingFactorChildren
                dblMaxValue *= Me.WeightingFactorChildren

                Dim NewTarget As New Target(selTarget.TargetDeadlineGuid, strOpMin, dblMinValue, strOpMax, dblMaxValue)

                objTargetsTotal.Add(NewTarget)
            Next
        Else
            objTargetsTotal = Me.Indicators.GetTargetValuesTotal(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return objTargetsTotal
    End Function

    Public Function GetTargetScoresTotal() As Targets
        Dim objTargetsTotal As Targets
        Dim dblScoreValue As Double

        If Me.Indicators.Count = 0 Then
            objTargetsTotal = New Targets
            For Each selTarget As Target In Me.Targets
                dblScoreValue = selTarget.Score

                dblScoreValue *= Me.WeightingFactorChildren

                Dim NewTarget As New Target(selTarget.TargetDeadlineGuid, dblScoreValue)
                objTargetsTotal.Add(NewTarget)
            Next
        Else
            objTargetsTotal = Me.Indicators.GetTargetScoresTotal(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return objTargetsTotal
    End Function

    Public Function GetTargetFormattedValue(ByVal intTargetIndex As Integer) As String
        Dim strFormattedValue As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim selTarget As Target
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                selTarget = GetTargetValuesTotal(intTargetIndex)

                If selTarget Is Nothing Then Return strFormattedValue

                Select Case Me.TargetSystem
                    Case TargetSystems.Simple, TargetSystems.Formula
                        strFormattedValue = DisplayAsUnit(selTarget.MinValue, intNrDecimals, strUnit)
                    Case Indicator.TargetSystems.ValueRange
                        strFormattedValue = selTarget.FormatTarget(strValueName, intNrDecimals, strUnit)
                End Select
        End Select

        Return strFormattedValue
    End Function

    Public Function GetTargetFormattedScore(ByVal intTargetIndex As Integer) As String
        Dim strFormattedScore As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim selTarget As Target
        Dim dblMaximumScore As Double
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                selTarget = GetTargetValuesTotal(intTargetIndex)

                Select Case Me.ScoringSystem
                    Case ScoringSystems.Value, ScoringSystems.Percentage
                        strFormattedScore = GetTargetFormattedValue(intTargetIndex)
                    Case ScoringSystems.Score
                        selTarget = GetTargetScoresTotal(intTargetIndex)

                        dblMaximumScore = GetTargetMaximumScore()
                        strFormattedScore = String.Format("{0}/{1}", selTarget.Score, dblMaximumScore)
                End Select
            Case QuestionTypes.Image, QuestionTypes.OpenEnded
                'do nothing
            Case Else

                selTarget = GetTargetScoresTotal(intTargetIndex)

                If selTarget IsNot Nothing Then
                    strFormattedScore = DisplayAsUnit(selTarget.Score, intNrDecimals, String.Empty)
                    dblMaximumScore = GetTargetMaximumScore()
                    strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                    strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
                End If
        End Select

        Return strFormattedScore
    End Function

    Public Function GetTargetMaximumScore(Optional ByVal intTargetIndex As Integer = -1) As Double

        Dim dblMaxScore As Double

        If Me.Indicators.Count = 0 Then
            Select Case Me.QuestionType
                Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                    dblMaxScore = Me.Targets.GetHighestScoreOfTargets
                Case QuestionTypes.MultipleChoice, QuestionTypes.MultipleOptions, QuestionTypes.YesNo
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                Case QuestionTypes.Ranking
                    dblMaxScore = ResponseClasses.Count
                Case QuestionTypes.Scale
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                    dblMaxScore /= ResponseClasses.Count
                Case QuestionTypes.CumulativeScale
                    dblMaxScore = ResponseClasses(0).Value * Statements.Count
                Case QuestionTypes.LikertTypeScale, QuestionTypes.FrequencyLikert
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                Case QuestionTypes.SemanticDiff, QuestionTypes.LikertScale
                    dblMaxScore = ResponseClasses.GetMaximumScore(Me.AddClassValuesForTotal)
                    dblMaxScore *= Statements.Count
            End Select

            dblMaxScore *= Me.WeightingFactorChildren
        Else
            dblMaxScore = Me.Indicators.GetTargetMaximumScore(intTargetIndex, Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblMaxScore
    End Function
#End Region

#Region "Population Targets methods"
    Public Function GetPopulationTotalPercentage() As PopulationTargets
        Dim objPopulationTargetsTotal As PopulationTargets

        If Me.Indicators.Count = 0 Then
            objPopulationTargetsTotal = New PopulationTargets
            For Each selPopulationTarget As PopulationTarget In Me.PopulationTargets
                Dim NewPopulationTarget As New PopulationTarget(selPopulationTarget.TargetDeadlineGuid, selPopulationTarget.Percentage)

                objPopulationTargetsTotal.Add(NewPopulationTarget)
            Next
        Else
            objPopulationTargetsTotal = Me.Indicators.GetPopulationTotalPercentage(Me.AggregateVertical)
        End If

        Return objPopulationTargetsTotal
    End Function

    Public Function GetPopulationBaselineValue() As Double
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblScoreValue As Double

        If Me.Indicators.Count = 0 Then
            dblScoreValue = GetPopulationBaselineMaximumValue() / 100 * Me.Baseline.PopulationPercentage
        Else
            dblScoreValue = Me.Indicators.GetPopulationBaselineValue(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblScoreValue
    End Function

    Public Function GetPopulationBaselineMaximumValue() As Double
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblScoreValue As Double

        If Me.Indicators.Count = 0 Then
            dblScoreValue = Me.Baseline.Value * Me.WeightingFactorChildren

            If selTargetGroup IsNot Nothing Then
                Select Case Me.AggregateHorizontal
                    Case AggregationOptions.Sum
                        dblScoreValue *= selTargetGroup.Number
                End Select

            End If
        Else
            dblScoreValue = Me.Indicators.GetPopulationBaselineMaximumValue(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblScoreValue
    End Function

    Public Function GetPopulationBaselineScore() As Double
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblScoreValue As Double

        If Me.Indicators.Count = 0 Then
            dblScoreValue = Me.Baseline.Score * Me.WeightingFactorChildren

            If selTargetGroup IsNot Nothing Then
                dblScoreValue *= Me.Baseline.PopulationPercentage() / 100
                Select Me.AggregateHorizontal
                    Case AggregationOptions.Sum
                        dblScoreValue *= selTargetGroup.Number
                End Select
            End If
        Else
            dblScoreValue = Me.Indicators.GetPopulationBaselineScore(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblScoreValue
    End Function

    Public Function GetPopulationBaselineMaximumScore() As Double
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblMaxScore As Double

        If Me.Indicators.Count = 0 Then
            dblMaxScore = Me.GetBaselineMaximumScore * Me.WeightingFactorChildren

            If selTargetGroup IsNot Nothing Then
                Select Case Me.AggregateHorizontal
                    Case AggregationOptions.Sum
                        dblMaxScore *= selTargetGroup.Number
                End Select
            End If
        Else
            dblMaxScore = Me.Indicators.GetPopulationBaselineMaximumScore(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblMaxScore
    End Function

    Public Function GetPopulationBaselineFormattedValue() As String
        Dim strFormattedValue As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim dblTotalValue As Double
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                dblTotalValue = GetPopulationBaselineMaximumValue()

                strFormattedValue = DisplayAsUnit(dblTotalValue, intNrDecimals, strUnit)
        End Select

        Return strFormattedValue
    End Function

    Public Function GetPopulationBaselineFormattedScore() As String
        Dim strFormattedScore As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblPopulationTotal As Double
        Dim dblMaximumScore As Double
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                Select Case Me.ScoringSystem
                    Case ScoringSystems.Value, ScoringSystems.Percentage
                        strFormattedScore = GetPopulationBaselineFormattedValue()
                    Case ScoringSystems.Score
                        dblPopulationTotal = GetPopulationBaselineScore()
                        dblMaximumScore = GetPopulationBaselineMaximumScore()

                        strFormattedScore = DisplayAsUnit(dblPopulationTotal, intNrDecimals, String.Empty)
                        strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                        strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
                End Select

            Case Else

                dblPopulationTotal = GetPopulationBaselineScore()
                dblMaximumScore = GetPopulationBaselineMaximumScore()

                strFormattedScore = DisplayAsUnit(dblPopulationTotal, intNrDecimals, String.Empty)
                strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
        End Select

        Return strFormattedScore
    End Function

    Public Function GetPopulationTargetValues() As Targets
        Dim objTargetsTotal As Targets
        Dim selTarget As Target
        Dim selPopulationTarget As PopulationTarget
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblMinValue, dblMaxValue As Double

        If Me.Indicators.Count = 0 Then
            objTargetsTotal = New Targets
            For i = 0 To Me.Targets.Count - 1
                selTarget = Me.Targets(i)
                selPopulationTarget = Me.PopulationTargets(i)

                dblMinValue = selTarget.MinValue * Me.WeightingFactorChildren * selPopulationTarget.Percentage / 100
                dblMaxValue = selTarget.MaxValue * Me.WeightingFactorChildren * selPopulationTarget.Percentage / 100

                If selTargetGroup IsNot Nothing Then
                    Select Case Me.AggregateHorizontal
                        Case AggregationOptions.Sum
                            dblMinValue *= selTargetGroup.Number
                            dblMaxValue *= selTargetGroup.Number
                    End Select
                End If

                Dim NewTarget As New Target(selTarget.TargetDeadlineGuid, selTarget.OpMin, dblMinValue, selTarget.OpMax, dblMaxValue)
                objTargetsTotal.Add(NewTarget)
            Next
        Else
            objTargetsTotal = Me.Indicators.GetPopulationTargetValues(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return objTargetsTotal
    End Function

    Public Function GetPopulationTargetMaximumValues() As Targets
        Dim objTargetsTotal As Targets
        Dim selTarget As Target
        Dim selPopulationTarget As PopulationTarget
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblMinValue, dblMaxValue As Double

        If Me.Indicators.Count = 0 Then
            objTargetsTotal = New Targets
            For i = 0 To Me.Targets.Count - 1
                selTarget = Me.Targets(i)
                selPopulationTarget = Me.PopulationTargets(i)

                dblMinValue = selTarget.MinValue * Me.WeightingFactorChildren
                dblMaxValue = selTarget.MaxValue * Me.WeightingFactorChildren

                If selTargetGroup IsNot Nothing Then
                    Select Case Me.AggregateHorizontal
                        Case AggregationOptions.Sum
                            dblMinValue *= selTargetGroup.Number
                            dblMaxValue *= selTargetGroup.Number
                    End Select
                End If

                Dim NewTarget As New Target(selTarget.TargetDeadlineGuid, selTarget.OpMin, dblMinValue, selTarget.OpMax, dblMaxValue)
                objTargetsTotal.Add(NewTarget)
            Next
        Else
            objTargetsTotal = Me.Indicators.GetPopulationTargetMaximumValues(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return objTargetsTotal
    End Function

    Public Function GetPopulationTargetScores() As Targets
        Dim objTargetsTotal As Targets
        Dim selTarget As Target
        Dim selPopulationTarget As PopulationTarget
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblScoreValue As Double

        If Me.Indicators.Count = 0 Then
            objTargetsTotal = New Targets
            For i = 0 To Me.Targets.Count - 1
                selTarget = Me.Targets(i)
                selPopulationTarget = Me.PopulationTargets(i)

                dblScoreValue = selTarget.Score * Me.WeightingFactorChildren * selPopulationTarget.Percentage / 100
                If selTargetGroup IsNot Nothing Then
                    Select Case Me.AggregateHorizontal
                        Case AggregationOptions.Sum
                            dblScoreValue *= selTargetGroup.Number
                    End Select
                End If

                Dim NewTarget As New Target(selTarget.TargetDeadlineGuid, dblScoreValue)
                objTargetsTotal.Add(NewTarget)
            Next
        Else
            objTargetsTotal = Me.Indicators.GetPopulationTargetScores(Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return objTargetsTotal
    End Function

    Public Function GetPopulationTargetMaximumScore(ByVal intTargetIndex As Integer) As Double
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblMaxScore As Double

        If Me.Indicators.Count = 0 Then
            dblMaxScore = GetTargetMaximumScore(intTargetIndex)
            If selTargetGroup IsNot Nothing Then
                Select Case Me.AggregateHorizontal
                    Case AggregationOptions.Sum
                        dblMaxScore *= selTargetGroup.Number
                End Select
            End If
        Else
            dblMaxScore = Me.Indicators.GetPopulationMaximumScore(intTargetIndex, Me.AggregateVertical, Me.WeightingFactorChildren)
        End If

        Return dblMaxScore
    End Function

    Public Function GetPopulationTargetFormattedValue(ByVal intTargetIndex As Integer) As String
        Dim strFormattedValue As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim objTotalValue, objMaxValue As Target
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                objTotalValue = GetPopulationTargetValues(intTargetIndex)
                objMaxValue = GetPopulationTargetMaximumValues(intTargetIndex)

                If Me.TargetSystem = TargetSystems.Simple Then
                    strFormattedValue = DisplayAsUnit(objTotalValue.MinValue, intNrDecimals, strUnit)
                    strFormattedMaxScore = DisplayAsUnit(objMaxValue.MinValue, intNrDecimals, strUnit)
                Else
                    strFormattedValue = objTotalValue.FormatTarget(strValueName, intNrDecimals, strUnit)
                    strFormattedMaxScore = objMaxValue.FormatTarget(strValueName, intNrDecimals, strUnit)
                End If

                strFormattedValue = String.Format("{0}/{1}", strFormattedValue, strFormattedMaxScore)
        End Select

        Return strFormattedValue
    End Function

    Public Function GetPopulationTargetFormattedScore(ByVal intTargetIndex As Integer) As String
        Dim strFormattedScore As String = String.Empty
        Dim strFormattedMaxScore As String = String.Empty
        Dim selTarget As Target
        Dim selPopulationTarget As PopulationTarget
        Dim selTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(Me.TargetGroupGuid)
        Dim dblPopulationTotal As Double
        Dim dblMaximumScore As Double
        Dim strValueName As String = "x"
        Dim intNrDecimals As Integer
        Dim strUnit As String = String.Empty

        If Me.ValuesDetail IsNot Nothing Then
            With Me.ValuesDetail
                strValueName = .ValueName
                intNrDecimals = .NrDecimals
                strUnit = .Unit
            End With
        End If

        Select Case Me.QuestionType
            Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue, QuestionTypes.Formula, QuestionTypes.Ratio
                selTarget = GetTargetScoresTotal(intTargetIndex)
                selPopulationTarget = GetPopulationTotalPercentage(intTargetIndex)
                Select Case Me.ScoringSystem
                    Case ScoringSystems.Value, ScoringSystems.Percentage
                        strFormattedScore = GetPopulationTargetFormattedValue(intTargetIndex)
                    Case ScoringSystems.Score
                        selTarget = GetPopulationTargetScores(intTargetIndex)

                        If selTarget IsNot Nothing Then
                            dblPopulationTotal = selTarget.Score
                            dblMaximumScore = GetPopulationTargetMaximumScore(intTargetIndex)
                        End If

                        strFormattedScore = DisplayAsUnit(dblPopulationTotal, intNrDecimals, String.Empty)
                        strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                        strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
                End Select

            Case Else
                selTarget = GetPopulationTargetScores(intTargetIndex)

                If selTarget IsNot Nothing Then
                    dblPopulationTotal = selTarget.Score
                    dblMaximumScore = GetPopulationTargetMaximumScore(intTargetIndex)
                End If
                strFormattedScore = DisplayAsUnit(dblPopulationTotal, intNrDecimals, String.Empty)
                strFormattedMaxScore = DisplayAsUnit(dblMaximumScore, intNrDecimals, String.Empty)

                strFormattedScore = String.Format("{0}/{1}", strFormattedScore, strFormattedMaxScore)
        End Select

        Return strFormattedScore
    End Function
#End Region

#Region "Calculate"
    Public Sub CalculateTargetsWithFormula()
        Dim strKey As String
        Dim intNumber As Integer

        If Me.Indicators.Count = 0 And TargetSystem = Indicator.TargetSystems.Formula Then
            Using Eparser As New ExpressionParser
                Select Case Me.QuestionType
                    Case QuestionTypes.AbsoluteValue, QuestionTypes.PercentageValue
                        With Eparser.VariableList
                            .Add("BL", Me.Baseline.Value)
                        End With
                        With Eparser.FormulaList
                            For i = Targets.Count - 1 To 0 Step -1
                                intNumber = i + 1
                                strKey = String.Format("T{0}", intNumber.ToString)
                                .Add(strKey, Targets(i).Formula)
                            Next
                        End With
                End Select

                For Each selTarget As Target In Me.Targets
                    Eparser.Expression = selTarget.Formula
                    'selTarget.OpMin = CONST_LargerThanOrEqual
                    selTarget.MinValue = Eparser.Parse
                    Select Case Me.ScoringSystem
                        Case ScoringSystems.Value
                            selTarget.Score = selTarget.MinValue
                        Case ScoringSystems.Percentage
                            selTarget.Score = 1

                    End Select
                Next
            End Using
        End If
    End Sub

    Private Sub CalculateTargetsWithFormula_SetTargetTotals()
        CalculateTargetsWithFormula()

        If Me.ParentIndicatorGuid <> Guid.Empty Then
            Dim ParentIndicator As Indicator = CurrentLogFrame.GetIndicatorByGuid(Me.ParentIndicatorGuid)

            If ParentIndicator IsNot Nothing Then
                ParentIndicator.Targets = ParentIndicator.Indicators.GetTargetValuesTotal(ParentIndicator.AggregateVertical, ParentIndicator.WeightingFactorChildren)
                ParentIndicator.CalculateTargetsWithFormula_SetTargetTotals()
            End If
        End If
    End Sub

    Public Sub CalculateValues()
        CalculateValues_Baseline()
        CalculateValues_Targets()
    End Sub

    Public Sub CalculateValues_Baseline()

        Select Case Me.QuestionType
            Case QuestionTypes.Formula, QuestionTypes.Ratio
                Baseline.Value = CalculateValues_Formula(Baseline.DoubleValues)
        End Select
    End Sub

    Public Sub CalculateValues_Targets()
        For Each selTarget As Target In Me.Targets
            Select Case Me.QuestionType
                Case QuestionTypes.Formula, QuestionTypes.Ratio
                    selTarget.OpMin = CONST_LargerThanOrEqual
                    selTarget.MinValue = CalculateValues_Formula(selTarget.DoubleValues)
            End Select
        Next
    End Sub

    Private Function CalculateValues_Formula(ByVal objDoubleValues As DoubleValues) As Double
        Dim dblValue As Double
        Dim intCounter, intLeadCounter As Integer
        Dim strFirst As String = String.Empty
        Dim strSecond As String = String.Empty
        Dim strCounter As String

        If Me.Statements.Count > 0 Then
            Using Eparser As New ExpressionParser
                Select Case Me.QuestionType
                    Case QuestionTypes.Formula, QuestionTypes.Ratio
                        For i = 0 To Statements.Count - 1
                            If intCounter > 0 And Decimal.Remainder(intCounter, 26) = 0 Then
                                intCounter = 0

                                strFirst = Chr(intLeadCounter + 65)
                                intLeadCounter += 1
                            End If

                            strSecond = Chr(intCounter + 65)
                            strCounter = strFirst & strSecond

                            Eparser.VariableList.Add(strCounter, objDoubleValues(i).Value)
                            intCounter += 1
                        Next

                        Eparser.Expression = ValuesDetail.Formula
                        dblValue = Eparser.Parse
                End Select
            End Using
        End If

        Return dblValue
    End Function

    Public Sub SetUnitWithFormula()
        Dim strUnit As String = String.Empty
        Dim intCounter, intLeadCounter As Integer
        Dim strFirst As String = String.Empty
        Dim strSecond As String = String.Empty
        Dim strCounter As String

        If Me.Statements.Count > 0 Then
            Using Eparser As New ExpressionParser
                Select Case Me.QuestionType
                    Case QuestionTypes.Formula, QuestionTypes.Ratio
                        For i = 0 To Statements.Count - 1
                            If intCounter > 0 And Decimal.Remainder(intCounter, 26) = 0 Then
                                intCounter = 0

                                strFirst = Chr(intLeadCounter + 65)
                                intLeadCounter += 1
                            End If

                            strSecond = Chr(intCounter + 65)
                            strCounter = strFirst & strSecond

                            Eparser.UnitList.Add(strCounter, Statements(i).ValuesDetail.Unit)
                            intCounter += 1
                        Next

                        Eparser.Expression = ValuesDetail.Formula
                        strUnit = Eparser.ParseUnits
                End Select
            End Using
        End If

        Me.ValuesDetail.Unit = strUnit
    End Sub

    Public Sub CalculateScores()
        CalculateScores_Baseline()
        CalculateScores_Targets()
    End Sub

    Public Sub CalculateScores_Baseline()
        Dim selResponseClass As ResponseClass
        Dim dblScore As Double

        Select Case Me.QuestionType
            Case QuestionTypes.MultipleChoice, QuestionTypes.YesNo
                For intIndex = 0 To Baseline.BooleanValues.Count - 1
                    If Baseline.BooleanValues(intIndex).Value = True Then
                        dblScore = ResponseClasses(intIndex).Value
                        Exit For
                    End If
                Next
                Me.Baseline.Score = dblScore

            Case QuestionTypes.MultipleOptions
                Me.Baseline.Score = CalculateScores_Classes(Baseline.BooleanValues)

            Case QuestionTypes.Ranking
                For intIndex = 0 To ResponseClasses.Count - 1
                    If ResponseClasses(intIndex).Value <> 0 Then
                        dblScore = Baseline.DoubleValues(intIndex).Value
                        Exit For
                    End If
                Next
                Me.Baseline.Score = dblScore

            Case QuestionTypes.Scale
                For i = 0 To Me.Statements.Count - 1
                    selResponseClass = Me.ResponseClasses(i)
                    If i < Me.Baseline.BooleanValues.Count Then
                        If Me.Baseline.BooleanValues(i).Value = True Then
                            dblScore += selResponseClass.Value
                        End If
                    End If
                Next
                dblScore /= Me.ResponseClasses.Count
                Me.Baseline.Score = dblScore

            Case QuestionTypes.CumulativeScale
                Dim intIndex As Integer = -1

                For i = Me.Statements.Count - 1 To 0 Step -1
                    If Me.Baseline.BooleanValues(i).Value = True Then
                        intIndex = i
                        Exit For
                    End If
                Next
                dblScore = CalculateScores_CumulativeScale(intIndex)
                Me.Baseline.Score = dblScore

            Case QuestionTypes.LikertTypeScale
                Me.Baseline.Score = CalculateScores_LikertTypeScale(Baseline.BooleanValues)

            Case QuestionTypes.LikertScale, QuestionTypes.SemanticDiff
                Me.Baseline.Score = CalculateScores_LikertScale(Baseline.BooleanValuesMatrix)
        End Select
    End Sub

    Public Sub CalculateScores_Targets()
        Dim selResponseClass As ResponseClass
        Dim dblScore As Double

        For Each selTarget As Target In Me.Targets
            dblScore = 0

            Select Case Me.QuestionType
                Case QuestionTypes.MultipleChoice, QuestionTypes.YesNo
                    For intIndex = 0 To selTarget.BooleanValues.Count - 1
                        If selTarget.BooleanValues(intIndex).Value = True Then
                            dblScore = ResponseClasses(intIndex).Value
                            Exit For
                        End If
                    Next
                    selTarget.Score = dblScore

                Case QuestionTypes.MultipleOptions
                    selTarget.Score = CalculateScores_Classes(selTarget.BooleanValues)

                Case QuestionTypes.Ranking
                    For intIndex = 0 To ResponseClasses.Count - 1
                        If ResponseClasses(intIndex).Value <> 0 Then
                            dblScore = selTarget.DoubleValues(intIndex).Value
                            Exit For
                        End If
                    Next
                    selTarget.Score = dblScore

                Case QuestionTypes.Scale
                    For i = 0 To Me.ResponseClasses.Count - 1
                        selResponseClass = Me.ResponseClasses(i)

                        If i < selTarget.BooleanValues.Count Then
                            If selTarget.BooleanValues(i).Value = True Then
                                dblScore += selResponseClass.Value
                            End If
                        End If
                    Next
                    dblScore /= Me.ResponseClasses.Count
                    selTarget.Score = dblScore

                Case QuestionTypes.CumulativeScale
                    Dim intIndex As Integer = -1

                    For i = Me.Statements.Count - 1 To 0 Step -1
                        If selTarget.BooleanValues(i).Value = True Then
                            intIndex = i
                            Exit For
                        End If
                    Next
                    selTarget.Score = CalculateScores_CumulativeScale(intIndex)

                Case QuestionTypes.LikertTypeScale
                    selTarget.Score = CalculateScores_LikertTypeScale(selTarget.BooleanValues)

                Case QuestionTypes.LikertScale, QuestionTypes.SemanticDiff
                    selTarget.Score = CalculateScores_LikertScale(selTarget.BooleanValuesMatrix)

            End Select
        Next
    End Sub

    Private Function CalculateScores_Classes(ByVal objBooleanValues As BooleanValues) As Double
        Dim dblScore As Double
        For i = 0 To objBooleanValues.Count - 1
            If objBooleanValues(i).Value = True Then dblScore += ResponseClasses(i).Value
        Next

        Return dblScore
    End Function

    Private Function CalculateScores_CumulativeScale(ByVal intIndex As Integer) As Double
        Dim dblScore As Double
        Dim dblValue As Double = Me.ResponseClasses(0).Value

        If intIndex >= 0 Then
            For i = 0 To intIndex
                dblScore += dblValue
            Next
        End If

        Return dblScore
    End Function

    Private Function CalculateScores_LikertScale(ByVal objMatrix As BooleanValuesMatrix) As Double
        Dim dblScore As Double

        For Each selRow As BooleanValuesMatrixRow In objMatrix
            dblScore += CalculateScores_LikertTypeScale(selRow.BooleanValues)
        Next
        Return dblScore
    End Function

    Private Function CalculateScores_LikertTypeScale(ByVal objBooleanValues As BooleanValues) As Double
        Dim dblScore As Double
        Dim intClassIndex As Integer

        For Each selValue As BooleanValue In objBooleanValues
            If selValue.Value = True Then
                intClassIndex = objBooleanValues.IndexOf(selValue)
                dblScore += ResponseClasses(intClassIndex).Value
                Exit For
            End If
        Next

        Return dblScore
    End Function
#End Region

    
End Class

Public Class Indicators
    Inherits System.Collections.CollectionBase

    Public Event IndicatorAdded(ByVal sender As Object, ByVal e As IndicatorAddedEventArgs)

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As Indicator
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Indicator)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal indicator As Indicator)
        List.Add(indicator)
        RaiseEvent IndicatorAdded(Me, New IndicatorAddedEventArgs(indicator))
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal indicator As Indicator)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, indicator.ItemName))
        Else
            List.Insert(index, indicator)
            RaiseEvent IndicatorAdded(Me, New IndicatorAddedEventArgs(indicator))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, Indicator.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal indicator As Indicator)
        If List.Contains(indicator) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, indicator.ItemName))
        Else
            List.Remove(indicator)
        End If
    End Sub

    Public Function IndexOf(ByVal indicator As Indicator) As Integer
        Return List.IndexOf(indicator)
    End Function

    Public Function Contains(ByVal indicator As Indicator) As Boolean
        Return List.Contains(indicator)
    End Function

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub
#End Region

#Region "Update sub indicators"
    Public Sub UpdateQuestionTypeOfChildren(ByVal intQuestionType As Integer)
        For Each selIndicator As Indicator In Me.List
            If selIndicator.QuestionType <> intQuestionType Then
                With selIndicator
                    .QuestionType = intQuestionType
                    .SetQuestionType_Settings()
                    .SetQuestionType_Statements()
                    .SetQuestionType_ResponseClasses()
                    .SetQuestionType_ResetTargets()

                    If .Indicators.Count > 0 Then
                        .Indicators.UpdateQuestionTypeOfChildren(intQuestionType)
                    End If
                End With
            End If
        Next
    End Sub

    Public Sub UpdateScoringSystemOfChildren(ByVal intScoringSystem As Integer)
        For Each selIndicator As Indicator In Me.List
            If selIndicator.ScoringSystem <> intScoringSystem Then
                selIndicator.ScoringSystem = intScoringSystem 'setting scoring system also fires update of selIndicator's child indicators
            End If
        Next
    End Sub

    Public Sub UpdateTargetSystemOfChildren(ByVal intTargetSystem As Integer)
        For Each selIndicator As Indicator In Me.List
            If selIndicator.TargetSystem <> intTargetSystem Then
                selIndicator.TargetSystem = intTargetSystem 'setting target system also fires update of selIndicator's child indicators
            End If
        Next
    End Sub

    Public Sub UpdateUnitsOfChildren(ByVal strValueName As String, ByVal intNrDecimals As Integer, ByVal strUnit As String)
        For Each selIndicator As Indicator In Me.List
            selIndicator.UpdateUnitsOfChildren(strValueName, intNrDecimals, strUnit)
        Next
    End Sub

    Public Sub UpdateRegistrationOfChildren(ByVal intRegistration As Integer)
        For Each selIndicator As Indicator In Me.List
            If selIndicator.Registration < intRegistration Then
                selIndicator.Registration = intRegistration 'setting registration also fires update of selIndicator's child indicators
            End If
        Next
    End Sub

    Public Sub UpdateTargetGroupGuidOfChildren(ByVal objTargetGroupGuid As Guid)
        For Each selIndicator As Indicator In Me.List
            If selIndicator.TargetGroupGuid <> objTargetGroupGuid Then
                selIndicator.TargetGroupGuid = objTargetGroupGuid 'setting target group guid also fires update of selIndicator's child indicators
            End If
        Next
    End Sub
#End Region

#Region "Targeting - general methods"
    Public Function TargetsValues_MakeTotal(ByVal objTargetsTotal As Targets, ByVal objChildTargetsTotal As Targets, ByVal intAggregateVertical As Integer) As Targets
        Dim selChildTarget As Target
        Dim dblMinValue, dblMaxValue As Double
        Dim strOpMin, strOpMax As String

        For i = 0 To objChildTargetsTotal.Count - 1
            selChildTarget = objChildTargetsTotal(i)

            strOpMin = selChildTarget.OpMin
            dblMinValue = selChildTarget.MinValue
            strOpMax = selChildTarget.OpMax
            dblMaxValue = selChildTarget.MaxValue

            If objTargetsTotal(i) Is Nothing Then
                objTargetsTotal.Add(New Target(selChildTarget.TargetDeadlineGuid, strOpMin, dblMinValue, strOpMax, dblMaxValue))
            Else
                If objTargetsTotal(i).OpMin = CONST_Equals And strOpMin <> CONST_Equals Then objTargetsTotal(i).OpMin = strOpMin
                If objTargetsTotal(i).OpMin = CONST_LargerThanOrEqual And strOpMin = CONST_LargerThan Then objTargetsTotal(i).OpMin = strOpMin
                If objTargetsTotal(i).OpMax = CONST_Equals And strOpMax <> CONST_Equals Then objTargetsTotal(i).OpMax = strOpMax
                If objTargetsTotal(i).OpMax = CONST_SmallerThanOrEqual And strOpMax = CONST_SmallerThan Then objTargetsTotal(i).OpMax = strOpMax

                Select Case intAggregateVertical
                    Case Indicator.AggregationOptions.Sum, Indicator.AggregationOptions.Average
                        If String.IsNullOrEmpty(strOpMin) = False Then objTargetsTotal(i).MinValue += dblMinValue
                        If String.IsNullOrEmpty(strOpMax) = False Then objTargetsTotal(i).MaxValue += dblMaxValue
                    Case Indicator.AggregationOptions.Maximum
                        If String.IsNullOrEmpty(strOpMin) = False And dblMinValue > objTargetsTotal(i).MinValue Then objTargetsTotal(i).MinValue = dblMinValue
                        If String.IsNullOrEmpty(strOpMax) = False And dblMaxValue > objTargetsTotal(i).MaxValue Then objTargetsTotal(i).MaxValue = dblMaxValue
                    Case Indicator.AggregationOptions.Minimum
                        If String.IsNullOrEmpty(strOpMin) = False And dblMinValue < objTargetsTotal(i).MinValue Then objTargetsTotal(i).MinValue = dblMinValue
                        If String.IsNullOrEmpty(strOpMax) = False And dblMaxValue < objTargetsTotal(i).MaxValue Then objTargetsTotal(i).MaxValue = dblMaxValue
                End Select
            End If
        Next

        Return objTargetsTotal
    End Function

    Public Function TargetsValues_Finalise(ByVal objTargetsTotal As Targets, ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        If intAggregateVertical = Indicator.AggregationOptions.Average Then
            Dim intCount As Integer = Me.Count
            For Each selTarget In objTargetsTotal
                If selTarget.MinValue <> 0 Then selTarget.MinValue = selTarget.MinValue / intCount
                If selTarget.MaxValue <> 0 Then selTarget.MaxValue = selTarget.MaxValue / intCount
            Next
        End If
        If sngWeightingValue <> 1 Then
            For Each selTarget In objTargetsTotal
                selTarget.MinValue *= sngWeightingValue
                selTarget.MaxValue *= sngWeightingValue
            Next
        End If

        Return objTargetsTotal
    End Function

    Public Function TargetsScores_MakeTotal(ByVal objTargetsTotal As Targets, ByVal objChildTargetsTotal As Targets, ByVal intAggregateVertical As Integer) As Targets
        Dim selChildTarget As Target
        Dim dblScoreValue As Double

        For i = 0 To objChildTargetsTotal.Count - 1
            selChildTarget = objChildTargetsTotal(i)

            dblScoreValue = selChildTarget.Score

            If objTargetsTotal(i) Is Nothing Then
                objTargetsTotal.Add(New Target(selChildTarget.TargetDeadlineGuid, dblScoreValue))
            Else
                Select Case intAggregateVertical
                    Case Indicator.AggregationOptions.Sum, Indicator.AggregationOptions.Average
                        objTargetsTotal(i).Score += dblScoreValue
                    Case Indicator.AggregationOptions.Maximum
                        If dblScoreValue > objTargetsTotal(i).Score Then objTargetsTotal(i).Score = dblScoreValue
                    Case Indicator.AggregationOptions.Minimum
                        If dblScoreValue < objTargetsTotal(i).Score Then objTargetsTotal(i).Score = dblScoreValue
                End Select
            End If
        Next

        Return objTargetsTotal
    End Function

    Public Function TargetsScores_Finalise(ByVal objTargetsTotal As Targets, ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        If intAggregateVertical = Indicator.AggregationOptions.Average Then
            Dim intCount As Integer = Me.Count
            For Each selTarget As Target In objTargetsTotal
                If selTarget.Score <> 0 Then selTarget.Score = selTarget.Score / intCount
            Next
        End If
        If sngWeightingValue <> 1 Then
            For Each selTarget As Target In objTargetsTotal
                selTarget.Score *= sngWeightingValue
            Next
        End If

        Return objTargetsTotal
    End Function

    Public Function Score_MakeTotal(ByVal dblTotalValue As Double, ByVal dblValueChild As Double, ByVal intAggregateVertical As Integer, ByVal intIndicatorIndex As Integer) As Double
        Select Case intAggregateVertical
            Case Indicator.AggregationOptions.Sum, Indicator.AggregationOptions.Average
                dblTotalValue += dblValueChild
            Case Indicator.AggregationOptions.Maximum
                If dblValueChild > dblTotalValue Then dblTotalValue = dblValueChild
            Case Indicator.AggregationOptions.Minimum
                If intIndicatorIndex = 0 Then dblTotalValue = dblValueChild
                If dblValueChild < dblTotalValue Then dblTotalValue = dblValueChild
        End Select

        Return dblTotalValue

    End Function

    Public Function Score_Finalise(ByVal dblTotalValue As Double, ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        If intAggregateVertical = Indicator.AggregationOptions.Average Then
            Dim intCount As Integer = Me.Count
            If dblTotalValue <> 0 Then dblTotalValue = dblTotalValue / intCount
        End If
        If sngWeightingValue <> 1 Then
            dblTotalValue *= sngWeightingValue
        End If

        Return dblTotalValue
    End Function
#End Region

#Region "Baseline methods"
    Public Function GetBaselineTotalValue(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Single
        Dim dblBaselineValue As Double
        Dim dblBaselineValueChild As Double

        For Each selIndicator As Indicator In Me.List
            dblBaselineValueChild = selIndicator.GetBaselineTotalValue()

            dblBaselineValue = Score_MakeTotal(dblBaselineValue, dblBaselineValueChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblBaselineValue = Score_Finalise(dblBaselineValue, intAggregateVertical, sngWeightingValue)

        Return dblBaselineValue
    End Function

    Public Function GetBaselineTotalScore(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Single
        Dim dblBaselineScore As Double
        Dim dblBaselineScoreChild As Double

        For Each selIndicator As Indicator In Me.List
            dblBaselineScoreChild = selIndicator.GetBaselineTotalScore()

            dblBaselineScore = Score_MakeTotal(dblBaselineScore, dblBaselineScoreChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblBaselineScore = Score_Finalise(dblBaselineScore, intAggregateVertical, sngWeightingValue)

        Return dblBaselineScore
    End Function

    Public Function GetBaselinePopulationValueTotal(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Single
        Dim dblPopulationValue As Double
        Dim dblPopulationValueChild As Double

        For Each selIndicator As Indicator In Me.List
            dblPopulationValueChild = selIndicator.Baseline.PopulationPercentage()

            dblPopulationValue = Score_MakeTotal(dblPopulationValue, dblPopulationValueChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblPopulationValue = Score_Finalise(dblPopulationValue, intAggregateVertical, sngWeightingValue)

        Return dblPopulationValue
    End Function

    Public Function GetBaselineMaximumScore(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblMaxScore As Double
        Dim dblMaxScoreChild As Double

        For Each selIndicator As Indicator In Me.List
            dblMaxScoreChild = selIndicator.GetBaselineMaximumScore()

            dblMaxScore = Score_MakeTotal(dblMaxScore, dblMaxScoreChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblMaxScore = Score_Finalise(dblMaxScore, intAggregateVertical, sngWeightingValue)

        Return dblMaxScore
    End Function
#End Region

#Region "Targets methods"
    Public Function GetTargetValuesTotal(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        Dim objTargetsTotal As New Targets

        For Each selIndicator As Indicator In Me.List
            Dim ChildTargetsTotal As Targets = selIndicator.GetTargetValuesTotal()
            objTargetsTotal = TargetsValues_MakeTotal(objTargetsTotal, ChildTargetsTotal, intAggregateVertical)
        Next

        objTargetsTotal = TargetsValues_Finalise(objTargetsTotal, intAggregateVertical, sngWeightingValue)

        Return objTargetsTotal
    End Function

    Public Function GetTargetScoresTotal(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        Dim objTargetsTotal As New Targets

        For Each selIndicator As Indicator In Me.List
            Dim objChildTargetsTotal As Targets = selIndicator.GetTargetScoresTotal()

            objTargetsTotal = TargetsScores_MakeTotal(objTargetsTotal, objChildTargetsTotal, intAggregateVertical)
        Next

        objTargetsTotal = TargetsScores_Finalise(objTargetsTotal, intAggregateVertical, sngWeightingValue)

        Return objTargetsTotal
    End Function

    Public Function GetTargetMaximumScore(ByVal intTargetIndex As Integer, ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblMaxScore As Double
        Dim dblMaxScoreChild As Double

        For Each selIndicator As Indicator In Me.List
            dblMaxScoreChild = selIndicator.GetTargetMaximumScore(intTargetIndex)

            dblMaxScore = Score_MakeTotal(dblMaxScore, dblMaxScoreChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblMaxScore = Score_Finalise(dblMaxScore, intAggregateVertical, sngWeightingValue)

        Return dblMaxScore
    End Function
#End Region

#Region "Population Targets methods"
    Public Function GetPopulationTotalPercentage(ByVal intAggregateVertical As Integer) As PopulationTargets
        Dim objPopulationTargets As New PopulationTargets
        Dim selChildPopulationTarget As PopulationTarget
        Dim dblPercentage As Double

        Select Case intAggregateVertical
            Case Indicator.AggregationOptions.Sum
                intAggregateVertical = Indicator.AggregationOptions.Average
        End Select

        For Each selIndicator As Indicator In Me.List

            Dim ChildPopulationTargetsTotal As PopulationTargets = selIndicator.GetPopulationTotalPercentage()
            For i = 0 To ChildPopulationTargetsTotal.Count - 1
                selChildPopulationTarget = ChildPopulationTargetsTotal(i)

                dblPercentage = selChildPopulationTarget.Percentage

                If objPopulationTargets(i) Is Nothing Then
                    objPopulationTargets.Add(New PopulationTarget(selChildPopulationTarget.TargetDeadlineGuid, dblPercentage))
                Else

                    Select Case intAggregateVertical
                        Case Indicator.AggregationOptions.Sum, Indicator.AggregationOptions.Average
                            objPopulationTargets(i).Percentage += dblPercentage
                        Case Indicator.AggregationOptions.Maximum
                            If dblPercentage > objPopulationTargets(i).Percentage Then objPopulationTargets(i).Percentage = dblPercentage
                        Case Indicator.AggregationOptions.Minimum
                            If dblPercentage < objPopulationTargets(i).Percentage Then objPopulationTargets(i).Percentage = dblPercentage
                    End Select
                End If
            Next
        Next

        If intAggregateVertical = Indicator.AggregationOptions.Average Then
            Dim intCount As Integer = Me.Count
            For Each selPopulationTarget As PopulationTarget In objPopulationTargets
                selPopulationTarget.Percentage /= intCount
            Next
        End If

        Return objPopulationTargets
    End Function

    Public Function GetPopulationBaselineValue(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblBaselineValue As Double
        Dim dblBaselineValueChild As Double

        For Each selIndicator As Indicator In Me.List
            dblBaselineValueChild = selIndicator.GetPopulationBaselineValue()

            dblBaselineValue = Score_MakeTotal(dblBaselineValue, dblBaselineValueChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblBaselineValue = Score_Finalise(dblBaselineValue, intAggregateVertical, sngWeightingValue)

        Return dblBaselineValue
    End Function

    Public Function GetPopulationBaselineMaximumValue(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblBaselineMaxValue As Double
        Dim dblBaselineMaxValueChild As Double

        For Each selIndicator As Indicator In Me.List
            dblBaselineMaxValueChild = selIndicator.GetPopulationBaselineMaximumValue()

            dblBaselineMaxValue = Score_MakeTotal(dblBaselineMaxValue, dblBaselineMaxValueChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblBaselineMaxValue = Score_Finalise(dblBaselineMaxValue, intAggregateVertical, sngWeightingValue)

        Return dblBaselineMaxValue
    End Function

    Public Function GetPopulationBaselineScore(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblBaselineScore As Double
        Dim dblMaxScoreChild As Double

        For Each selIndicator As Indicator In Me.List
            dblMaxScoreChild = selIndicator.GetPopulationBaselineScore()

            dblBaselineScore = Score_MakeTotal(dblBaselineScore, dblMaxScoreChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblBaselineScore = Score_Finalise(dblBaselineScore, intAggregateVertical, sngWeightingValue)

        Return dblBaselineScore
    End Function

    Public Function GetPopulationBaselineMaximumScore(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblMaxScore As Double
        Dim dblMaxScoreChild As Double

        For Each selIndicator As Indicator In Me.List
            dblMaxScoreChild = selIndicator.GetPopulationBaselineMaximumScore()

            dblMaxScore = Score_MakeTotal(dblMaxScore, dblMaxScoreChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblMaxScore = Score_Finalise(dblMaxScore, intAggregateVertical, sngWeightingValue)

        Return dblMaxScore
    End Function

    Public Function GetPopulationTargetValues(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        Dim objTargetsTotal As New Targets

        For Each selIndicator As Indicator In Me.List
            Dim ChildTargetsTotal As Targets = selIndicator.GetPopulationTargetValues()
            objTargetsTotal = TargetsValues_MakeTotal(objTargetsTotal, ChildTargetsTotal, intAggregateVertical)
        Next

        objTargetsTotal = TargetsValues_Finalise(objTargetsTotal, intAggregateVertical, sngWeightingValue)
        Return objTargetsTotal
    End Function

    Public Function GetPopulationTargetMaximumValues(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        Dim objTargetsTotal As New Targets

        For Each selIndicator As Indicator In Me.List
            Dim ChildTargetsTotal As Targets = selIndicator.GetPopulationTargetMaximumValues()
            objTargetsTotal = TargetsValues_MakeTotal(objTargetsTotal, ChildTargetsTotal, intAggregateVertical)
        Next

        objTargetsTotal = TargetsValues_Finalise(objTargetsTotal, intAggregateVertical, sngWeightingValue)
        Return objTargetsTotal
    End Function

    Public Function GetPopulationTargetScores(ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Targets
        Dim objTargetsTotal As New Targets

        For Each selIndicator As Indicator In Me.List
            Dim ChildTargetsTotal As Targets = selIndicator.GetPopulationTargetScores()
            objTargetsTotal = TargetsScores_MakeTotal(objTargetsTotal, ChildTargetsTotal, intAggregateVertical)
        Next

        objTargetsTotal = TargetsScores_Finalise(objTargetsTotal, intAggregateVertical, sngWeightingValue)
        Return objTargetsTotal
    End Function

    Public Function GetPopulationMaximumScore(ByVal intTargetIndex As Integer, ByVal intAggregateVertical As Integer, ByVal sngWeightingValue As Single) As Double
        Dim dblMaxScore As Double
        Dim dblMaxScoreChild As Double

        For Each selIndicator As Indicator In Me.List
            dblMaxScoreChild = selIndicator.GetPopulationTargetMaximumScore(intTargetIndex)

            dblMaxScore = Score_MakeTotal(dblMaxScore, dblMaxScoreChild, intAggregateVertical, Me.IndexOf(selIndicator))
        Next

        dblMaxScore = Score_Finalise(dblMaxScore, intAggregateVertical, sngWeightingValue)

        Return dblMaxScore
    End Function


#End Region

#Region "Get by GUID"
    Public Function GetIndicatorByGuid(ByVal objGuid As Guid) As Indicator
        Dim selIndicator As Indicator = Nothing
        For Each objIndicator As Indicator In Me.List
            If objIndicator.Guid = objGuid Then
                selIndicator = objIndicator
                Exit For
            ElseIf objIndicator.Indicators.Count > 0 Then
                selIndicator = objIndicator.Indicators.GetIndicatorByGuid(objGuid)
                If selIndicator IsNot Nothing Then Exit For
            End If
        Next
        Return selIndicator
    End Function

    Public Function GetStatementByGuid(ByVal objGuid As Guid) As Statement
        Dim selStatement As Statement = Nothing
        For Each objIndicator As Indicator In Me.List
            selStatement = objIndicator.Statements.GetStatementByGuid(objGuid)
            If selStatement IsNot Nothing Then
                Exit For
            ElseIf objIndicator.Indicators.Count > 0 Then
                selStatement = objIndicator.Indicators.GetStatementByGuid(objGuid)
                If selStatement IsNot Nothing Then Exit For
            End If
        Next
        Return selStatement
    End Function

    Public Function GetTargetByGuid(ByVal objGuid As Guid) As Target
        Dim selTarget As Target = Nothing
        For Each objIndicator As Indicator In Me.List
            selTarget = objIndicator.Targets.GetTargetByGuid(objGuid)
            If selTarget IsNot Nothing Then
                Exit For
            ElseIf objIndicator.Indicators.Count > 0 Then
                selTarget = objIndicator.Indicators.GetTargetByGuid(objGuid)
                If selTarget IsNot Nothing Then Exit For
            End If
        Next
        Return selTarget
    End Function

    'Public Function GetTargetBooleanValueByGuid(ByVal objGuid As Guid) As BooleanValue
    '    Dim selTargetBooleanValue As BooleanValue = Nothing
    '    For Each objIndicator As Indicator In Me.List
    '        For Each selStatement As Response In objIndicator.Responses
    '            selTargetBooleanValue = selStatement.BooleanValues.GetBooleanValueByGuid(objGuid)
    '            If selTargetBooleanValue IsNot Nothing Then Exit For
    '        Next
    '        If selTargetBooleanValue IsNot Nothing Then
    '            Exit For
    '        ElseIf objIndicator.Indicators.Count > 0 Then
    '            selTargetBooleanValue = objIndicator.Indicators.GetTargetBooleanValueByGuid(objGuid)
    '            If selTargetBooleanValue IsNot Nothing Then Exit For
    '        End If
    '    Next
    '    Return selTargetBooleanValue
    'End Function

    'Public Function GetIntegerResponseByGuid(ByVal objGuid As Guid) As IntegerResponse
    '    Dim selIntegerResponse As IntegerResponse = Nothing
    '    For Each objIndicator As Indicator In Me.List
    '        For Each objStatement As Response In objIndicator.Responses
    '            'If objStatement.ValuesDetail IsNot Nothing Then
    '            selIntegerResponse = objStatement.IntegerResponses.GetIntegerResponseByGuid(objGuid)
    '            If selIntegerResponse IsNot Nothing Then Exit For
    '            'End If
    '        Next
    '        If selIntegerResponse IsNot Nothing Then
    '            Exit For
    '        ElseIf objIndicator.Indicators.Count > 0 Then
    '            selIntegerResponse = objIndicator.Indicators.GetIntegerResponseByGuid(objGuid)
    '            If selIntegerResponse IsNot Nothing Then Exit For
    '        End If
    '    Next
    '    Return selIntegerResponse
    'End Function

    Public Function GetResponseClassByGuid(ByVal objGuid As Guid) As ResponseClass
        Dim selResponseClass As ResponseClass = Nothing
        For Each objIndicator As Indicator In Me.List
            selResponseClass = objIndicator.ResponseClasses.GetResponseClassByGuid(objGuid)
            If selResponseClass IsNot Nothing Then
                Exit For
            ElseIf objIndicator.Indicators.Count > 0 Then
                selResponseClass = objIndicator.Indicators.GetResponseClassByGuid(objGuid)
                If selResponseClass IsNot Nothing Then Exit For
            End If
        Next
        Return selResponseClass
    End Function

    Public Function GetVerificationSourceByGuid(ByVal objGuid As Guid) As VerificationSource
        Dim selVerificationSource As VerificationSource = Nothing
        For Each objIndicator As Indicator In Me.List
            selVerificationSource = objIndicator.VerificationSources.GetVerificationSourceByGuid(objGuid)
            If selVerificationSource IsNot Nothing Then
                Exit For
            ElseIf objIndicator.Indicators.Count > 0 Then
                selVerificationSource = objIndicator.Indicators.GetVerificationSourceByGuid(objGuid)
                If selVerificationSource IsNot Nothing Then Exit For
            End If
        Next
        Return selVerificationSource
    End Function
#End Region

End Class

Public Class IndicatorAddedEventArgs
    Inherits EventArgs

    Public Property Indicator As Indicator

    Public Sub New(ByVal objIndicator As Indicator)
        MyBase.New()

        Me.Indicator = objIndicator
    End Sub
End Class
