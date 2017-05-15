Public Class ExportIndicatorRow
    Private strSortNumber As String = String.Empty
    Private objStruct As Struct
    Private objIndicator As Indicator
    Private strRowHeader As String
    Private strUnit As String
    Private boolValueRangeSet As Boolean
    Private sngValueRangeMin, sngValueRangeMax As Single
    Private intNrDecimals As Integer
    Private dblMaximumScore As Double
    Private strScoringLegend As String
    Private objBaselineValue, objBaselineScore As Object
    Private objCells() As Object

    Private intScoringSystem As Integer
    Private strFormula As String

    Private objStatement As Statement
    Private objResponseClasses As ResponseClasses
    Private objTargetDeadline As TargetDeadline
    Private strTargetDeadlineDate As String
    Private intResponseClassesNumber As Integer
    Private boolNoCheckBox As Boolean
    Private intObjectType As Integer

    Public Enum ObjectTypes
        NotSet = -1
        TableHeader = 1
        ColumnHeader = 2
        RowHeader = 3
        WhiteSpace = 4

        Struct = 10
        Goal = 11
        Purpose = 12
        Output = 13
        Activity = 14

        Indicator = 20
        'Statement = 21
        'BooleanResponse = 22
        'IntegerResponse = 23
        'ValueRange = 24
        'ResponseClass = 25
        'Target = 26
        'Baseline = 27

        AnswerAreaOpenEnded = 30
        AnswerAreaMaxDiff = 31
        AnswerAreaValue = 40
        AnswerAreaFormula = 41
        AnswerAreaMultipleOptions = 42
        AnswerAreaLikertType = 43
        AnswerAreaSemanticDiff = 44
        AnswerAreaScale = 45
        AnswerAreaLikertScale = 46
        AnswerAreaCumulativeScale = 47
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal sortnumber As String, ByVal struct As Struct)
        Me.ObjectType = objecttype
        Me.SortNumber = sortnumber
        Me.Struct = struct
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal sortnumber As String, ByVal indicator As Indicator)
        Me.ObjectType = objecttype
        Me.SortNumber = sortnumber
        Me.Indicator = indicator
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal rowheader As String, ByVal unit As String, ByVal valuerangemin As Single, ByVal valuerangemax As Single, ByVal valuerangeset As Boolean, _
                   ByVal nrdecimals As Integer, ByVal scoringsystem As Integer, ByVal scoringlegend As String, ByVal maximumscore As Double, ByVal baselinevalue As Object, ByVal baselinescore As Object, _
                   ByVal cells() As Object)
        Me.ObjectType = objecttype
        Me.RowHeader = rowheader
        Me.Unit = unit
        Me.ValueRangeMin = valuerangemin
        Me.ValueRangeMax = valuerangemax
        Me.ValueRangeSet = valuerangeset
        Me.NrDecimals = nrdecimals
        Me.ScoringSystem = scoringsystem
        Me.ScoringLegend = scoringlegend
        Me.MaximumScore = maximumscore
        Me.BaselineValue = baselinevalue
        Me.BaselineScore = baselinescore
        Me.Cells = cells
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal rowheader As String, ByVal nrdecimals As Integer, ByVal scoringlegend As String, ByVal maximumscore As Double, ByVal baselinescore As Object, ByVal cells() As Object)
        Me.ObjectType = objecttype
        Me.RowHeader = rowheader
        Me.NrDecimals = nrdecimals
        Me.ScoringSystem = Indicator.ScoringSystems.Score
        Me.ScoringLegend = scoringlegend
        Me.MaximumScore = maximumscore
        Me.BaselineScore = baselinescore
        Me.Cells = cells
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal statement As Statement)
        Me.ObjectType = objecttype
        Me.Statement = statement
    End Sub

#Region "Properties"
    Public Property ObjectType As Integer
        Get
            Return intObjectType
        End Get
        Set(ByVal value As Integer)
            intObjectType = value
        End Set
    End Property

    Public Property SortNumber As String
        Get
            Return strSortNumber
        End Get
        Set(ByVal value As String)
            If value = Nothing Then value = String.Empty
            strSortNumber = value
        End Set
    End Property

    Public Property Struct As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
        End Set
    End Property

    Public Property Indicator As Indicator
        Get
            Return objIndicator
        End Get
        Set(ByVal value As Indicator)
            objIndicator = value
        End Set
    End Property

    Public ReadOnly Property QuestionType() As Integer
        Get
            If objIndicator IsNot Nothing Then
                Return objIndicator.QuestionType
            Else
                Return ObjectTypes.NotSet
            End If
        End Get
    End Property

    Public Property Statement() As Statement
        Get
            Return objStatement
        End Get
        Set(ByVal value As Statement)
            objStatement = value
        End Set
    End Property

    Public Property RowHeader As String
        Get
            Return strRowHeader
        End Get
        Set(ByVal value As String)
            strRowHeader = value
        End Set
    End Property

    Public Property Unit As String
        Get
            Return strUnit
        End Get
        Set(ByVal value As String)
            strUnit = value
        End Set
    End Property

    Public Property ValueRangeMin As Single
        Get
            Return sngValueRangeMin
        End Get
        Set(ByVal value As Single)
            sngValueRangeMin = value
        End Set
    End Property

    Public Property ValueRangeMax As Single
        Get
            Return sngValueRangeMax
        End Get
        Set(ByVal value As Single)
            sngValueRangeMax = value
        End Set
    End Property

    Public Property ValueRangeSet As Boolean
        Get
            Return boolValueRangeSet
        End Get
        Set(ByVal value As Boolean)
            boolValueRangeSet = value
        End Set
    End Property

    Public Property NrDecimals As Integer
        Get
            Return intNrDecimals
        End Get
        Set(ByVal value As Integer)
            intNrDecimals = value
        End Set
    End Property

    Public Property ScoringSystem As Integer
        Get
            Return intScoringSystem
        End Get
        Set(ByVal value As Integer)
            intScoringSystem = value
        End Set
    End Property

    Public Property ScoringLegend As String
        Get
            Return strScoringLegend
        End Get
        Set(ByVal value As String)
            strScoringLegend = value
        End Set
    End Property

    Public Property MaximumScore As Double
        Get
            Return dblMaximumScore
        End Get
        Set(ByVal value As Double)
            dblMaximumScore = value
        End Set
    End Property

    Public Property BaselineValue As Object
        Get
            Return objBaselineValue
        End Get
        Set(ByVal value As Object)
            objBaselineValue = value
        End Set
    End Property

    Public Property BaselineScore As Object
        Get
            Return objBaselineScore
        End Get
        Set(ByVal value As Object)
            objBaselineScore = value
        End Set
    End Property

    Public Property Cells As Object()
        Get
            Return objCells
        End Get
        Set(ByVal value As Object())
            objCells = value
        End Set
    End Property

    Public Property Formula As String
        Get
            Return strFormula
        End Get
        Set(ByVal value As String)
            strFormula = value
        End Set
    End Property





    Public Property ResponseClasses() As ResponseClasses
        Get
            Return objResponseClasses
        End Get
        Set(ByVal value As ResponseClasses)
            objResponseClasses = value
        End Set
    End Property

    Public Property ResponseClassesNumber() As Integer
        Get
            Return intResponseClassesNumber
        End Get
        Set(ByVal value As Integer)
            intResponseClassesNumber = value
        End Set
    End Property

    Public Property TargetDeadline As TargetDeadline
        Get
            Return objTargetDeadline
        End Get
        Set(ByVal value As TargetDeadline)
            objTargetDeadline = value
        End Set
    End Property

    Public Property TargetDeadlineDate As String
        Get
            Return strTargetDeadlineDate
        End Get
        Set(ByVal value As String)
            strTargetDeadlineDate = value
        End Set
    End Property




#End Region
End Class

Public Class ExportIndicatorRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportIndicatorRow As ExportIndicatorRow)
        List.Add(ExportIndicatorRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportIndicatorRow As ExportIndicatorRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportIndicatorRow)
        Else
            List.Insert(index, ExportIndicatorRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportIndicatorRow As ExportIndicatorRow)
        If List.Contains(ExportIndicatorRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportIndicatorRow not in list!")
        Else
            List.Remove(ExportIndicatorRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportIndicatorRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportIndicatorRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal IndicatorRow As ExportIndicatorRow) As Integer
        Return List.IndexOf(IndicatorRow)
    End Function

    Public Function Contains(ByVal IndicatorRow As ExportIndicatorRow) As Boolean
        Return List.Contains(IndicatorRow)
    End Function
End Class
