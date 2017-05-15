Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class FindReplace
    Implements IDisposable

    Private objLogframe As LogFrame
    Private intMatchIndex As Integer = -1
    Private intMatchLength As Integer
    Private objMatchItem, objStartItem As Object
    Private objMatchChildItem As Object
    Private strMatchProperty As String
    Private boolMatchCase As Boolean
    Private strFindText As String
    Private strReplaceText As String
    Private boolScanOn, boolScanOnChildItem As Boolean
    Private boolSearchAllPanes As Boolean
    Private boolSearchResources As Boolean
    Private boolSearchPastEnd As Boolean
    Private boolReplaceAll As Boolean
    Private intItemsReplaced As Integer

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame)
        Me.Logframe = logframe
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

    Public Property MatchIndex As Integer
        Get
            Return intMatchIndex
        End Get
        Set(ByVal value As Integer)
            intMatchIndex = value
        End Set
    End Property

    Public Property MatchLength As Integer
        Get
            Return intMatchLength
        End Get
        Set(ByVal value As Integer)
            intMatchLength = value
        End Set
    End Property

    Public Property MatchCase As Boolean
        Get
            Return boolMatchCase
        End Get
        Set(ByVal value As Boolean)
            boolMatchCase = value
        End Set
    End Property

    Public Property MatchItem As Object
        Get
            Return objMatchItem
        End Get
        Set(ByVal value As Object)
            objMatchItem = value
        End Set
    End Property

    Public Property MatchChildItem As Object
        Get
            Return objMatchChildItem
        End Get
        Set(ByVal value As Object)
            objMatchChildItem = value
        End Set
    End Property

    Public Property MatchProperty As String
        Get
            Return strMatchProperty
        End Get
        Set(ByVal value As String)
            strMatchProperty = value
        End Set
    End Property

    Public Property StartItem As Object
        Get
            Return objStartItem
        End Get
        Set(ByVal value As Object)
            objStartItem = value
        End Set
    End Property

    Public Property FindText As String
        Get
            Return strFindText
        End Get
        Set(ByVal value As String)
            strFindText = value
        End Set
    End Property

    Public Property ReplaceText As String
        Get
            Return strReplaceText
        End Get
        Set(ByVal value As String)
            strReplaceText = value
        End Set
    End Property

    Public Property ScanOn As Boolean
        Get
            Return boolScanOn
        End Get
        Set(ByVal value As Boolean)
            boolScanOn = value
        End Set
    End Property

    Public Property ScanOnChildItem As Boolean
        Get
            Return boolScanOnChildItem
        End Get
        Set(ByVal value As Boolean)
            boolScanOnChildItem = value
        End Set
    End Property

    Public Property SearchAllPanes As Boolean
        Get
            Return boolSearchAllPanes
        End Get
        Set(ByVal value As Boolean)
            boolSearchAllPanes = value
        End Set
    End Property

    Public Property SearchResources As Boolean
        Get
            Return boolSearchResources
        End Get
        Set(ByVal value As Boolean)
            boolSearchResources = value
        End Set
    End Property

    Public Property SearchPastEnd As Boolean
        Get
            Return boolSearchPastEnd
        End Get
        Set(ByVal value As Boolean)
            boolSearchPastEnd = value
        End Set
    End Property

    Public Property ReplaceAll As Boolean
        Get
            Return boolReplaceAll
        End Get
        Set(ByVal value As Boolean)
            boolReplaceAll = value
        End Set
    End Property

    Public Property ItemsReplaced As Integer
        Get
            Return intItemsReplaced
        End Get
        Set(ByVal value As Integer)
            intItemsReplaced = value
        End Set
    End Property
#End Region

#Region "Find/replace main"
    Public Function FindReplaceInProjectInfo(ByVal findtext As String, ByVal replacetext As String, ByVal matchcase As Boolean, ByVal startitem As Object) As Boolean
        Dim boolFound As Boolean

        Me.MatchCase = matchcase
        Me.FindText = findtext
        Me.ReplaceText = replacetext

        If startitem Is Me.StartItem Then
            If MatchIndex >= 0 Then MatchIndex += MatchLength
        End If

        Me.StartItem = startitem
        boolScanOn = False

        If FindReplace_ProjectInfo() = True Then
            boolFound = True
        End If

        If boolFound = True And String.IsNullOrEmpty(Me.ReplaceText) = False Then
            Replace()

            ItemsReplaced += 1

            If boolReplaceAll = True Then
                boolFound = FindReplaceInProjectInfo(findtext, replacetext, matchcase, MatchItem)
            End If
        End If

        Return boolFound
    End Function

    Public Function FindReplaceInLogframe(ByVal findtext As String, ByVal replacetext As String, ByVal matchcase As Boolean, ByVal startitem As Object) As Boolean
        Dim boolFound As Boolean

        Me.MatchCase = matchcase
        Me.FindText = findtext
        Me.ReplaceText = replacetext

        If startitem Is Me.StartItem Then
            If MatchIndex >= 0 Then MatchIndex += MatchLength
        End If

        Me.StartItem = startitem
        boolScanOn = False

        If FindReplace_Goals() = True Then
            boolFound = True
        ElseIf FindReplace_Purposes() = True Then
            boolFound = True
        ElseIf FindReplace_Outputs() = True Then
            boolFound = True
        ElseIf FindReplace_ActivitiesInLogframe() = True Then
            boolFound = True
        End If

        If boolFound = True And String.IsNullOrEmpty(Me.ReplaceText) = False Then
            Replace()

            ItemsReplaced += 1

            If boolReplaceAll = True Then
                boolFound = FindReplaceInLogframe(findtext, replacetext, matchcase, MatchItem)
            End If
        End If

        Return boolFound
    End Function

    Public Function FindReplaceInPlanning(ByVal findtext As String, ByVal replacetext As String, ByVal matchcase As Boolean, ByVal startitem As Object) As Boolean
        Dim boolFound As Boolean

        Me.MatchCase = matchcase
        Me.FindText = findtext
        Me.ReplaceText = replacetext

        If startitem Is Me.StartItem Then
            If MatchIndex >= 0 Then MatchIndex += MatchLength
        End If

        Me.StartItem = startitem
        boolScanOn = False

        If FindReplace_Planning() = True Then
            boolFound = True
        End If

        If boolFound = True And String.IsNullOrEmpty(Me.ReplaceText) = False Then
            Replace()

            ItemsReplaced += 1

            If boolReplaceAll = True Then
                boolFound = FindReplaceInPlanning(findtext, replacetext, matchcase, MatchItem)
            End If
        End If

        Return boolFound
    End Function

    Public Function FindReplaceInBudgetYear(ByVal findtext As String, ByVal replacetext As String, ByVal matchcase As Boolean, ByVal startitem As Object, _
                                            ByVal intBudgetYearIndex As Integer) As Boolean
        Dim boolFound As Boolean

        Me.MatchCase = matchcase
        Me.FindText = findtext
        Me.ReplaceText = replacetext

        If startitem Is Me.StartItem Then
            If MatchIndex >= 0 Then MatchIndex += MatchLength
        End If

        Me.StartItem = startitem
        boolScanOn = False

        If FindReplace_BudgetYear(intBudgetYearIndex) = True Then
            boolFound = True
        End If

        If boolFound = True And String.IsNullOrEmpty(Me.ReplaceText) = False Then
            Replace()

            ItemsReplaced += 1

            If boolReplaceAll = True Then
                boolFound = FindReplaceInBudgetYear(findtext, replacetext, matchcase, MatchItem, intBudgetYearIndex)
            End If
        End If

        Return boolFound
    End Function
#End Region

#Region "Search in property values"
    Private Function FindInText(ByVal selItem As Object) As Boolean
        Dim strText As String = selItem.Text
        Dim objMatches As MatchCollection
        Dim intStartIndex As Integer = -1

        If boolMatchCase = False Then
            objMatches = Regex.Matches(strText, Me.FindText, RegexOptions.IgnoreCase)
        Else
            objMatches = Regex.Matches(strText, Me.FindText)
        End If

        If objMatches.Count > 0 Then
            For i = 0 To objMatches.Count - 1
                intStartIndex = objMatches(i).Index

                If intStartIndex > MatchIndex Then
                    MatchIndex = intStartIndex
                    MatchLength = objMatches(i).Length
                    MatchItem = selItem
                    MatchProperty = String.Empty

                    Return True
                End If
            Next
        Else
            MatchIndex = -1
            MatchLength = 0
        End If

        Return False
    End Function

    Private Function FindInItem(ByVal selItem As Object) As Boolean
        Dim objType As Type = selItem.GetType
        Dim propInfo() As PropertyInfo = objType.GetProperties
        Dim PropInfoIndex As Integer
        Dim boolFound As Boolean

        Dim boolIsNumeric As Boolean = IsNumeric(FindText)
        Dim boolIsDate As Boolean = IsDate(FindText)

        If selItem Is Nothing Then Return False

        Do
            Dim selPropInfo As PropertyInfo = propInfo(PropInfoIndex)

            If PropertyCanBeModified(selItem, selPropInfo) = True Then
                Select Case selPropInfo.PropertyType
                    Case GetType(String)
                        boolFound = FindInItem_Text(selItem, selPropInfo)
                        If boolFound = True Then Return True
                    Case GetType(Date)
                        If boolIsDate = True Then
                            boolFound = FindInItem_Date(selItem, selPropInfo)
                            If boolFound = True Then Return True
                        End If
                    Case GetType(Integer), GetType(Single), GetType(Double)
                        If boolIsNumeric = True Then
                            boolFound = FindInItem_Value(selItem, selPropInfo)
                            If boolFound = True Then Return True
                        End If
                    Case GetType(Currency)
                        If boolIsNumeric = True Then
                            boolFound = FindInItem_Currency(selItem, selPropInfo)
                            If boolFound = True Then Return True
                        End If
                End Select
            End If
            PropInfoIndex += 1
            If PropInfoIndex > propInfo.Count - 1 Then
                PropInfoIndex = 0
                boolFound = False

                Exit Do
            End If
        Loop

        Return boolFound
    End Function

    Private Function FindInItem_Text(ByVal selItem As Object, ByVal selPropInfo As PropertyInfo) As Boolean
        Dim objMatches As MatchCollection
        Dim intStartIndex As Integer = -1
        Dim boolScanProperty As Boolean
        Dim strText As String, strPropertyName As String

        strPropertyName = selPropInfo.Name
        If String.IsNullOrEmpty(MatchProperty) Then
            boolScanProperty = True
        ElseIf strPropertyName = MatchProperty Then
            boolScanProperty = True
        End If

        If boolScanProperty = True Then
            Dim objValue As Object = selPropInfo.GetValue(selItem, Nothing)
            If objValue IsNot Nothing Then
                strText = objValue.ToString
            Else
                strText = String.Empty
            End If

            If boolMatchCase = False Then
                objMatches = Regex.Matches(strText, Me.FindText, RegexOptions.IgnoreCase)
            Else
                objMatches = Regex.Matches(strText, Me.FindText)
            End If

            If objMatches.Count > 0 Then
                For i = 0 To objMatches.Count - 1
                    intStartIndex = objMatches(i).Index

                    If intStartIndex > MatchIndex Then
                        MatchIndex = intStartIndex
                        MatchLength = objMatches(i).Length
                        MatchProperty = strPropertyName

                        Return True
                    End If
                Next
            End If
            MatchIndex = -1
            MatchLength = 0
            MatchProperty = String.Empty

        End If

        Return False
    End Function

    Private Function FindInItem_RichText(ByVal strRichText As String) As Boolean
        If String.IsNullOrEmpty(strRichText) Then Return False

        Dim strText As String = RichTextToText(strRichText)
        Dim objMatches As MatchCollection
        Dim intStartIndex As Integer = -1

        If boolMatchCase = False Then
            objMatches = Regex.Matches(strText, Me.FindText, RegexOptions.IgnoreCase)
        Else
            objMatches = Regex.Matches(strText, Me.FindText)
        End If

        If objMatches.Count > 0 Then
            For i = 0 To objMatches.Count - 1
                intStartIndex = objMatches(i).Index

                If intStartIndex > MatchIndex Then
                    MatchIndex = intStartIndex
                    MatchLength = objMatches(i).Length

                    Return True
                End If
            Next
        End If
        MatchIndex = -1
        MatchLength = 0
        MatchProperty = String.Empty

        Return False
    End Function

    Private Function FindInItem_Date(ByVal selItem As Object, ByVal selPropInfo As PropertyInfo) As Boolean
        Dim datFindDate As Date
        Dim selDate As Date
        Dim boolScanProperty As Boolean
        Dim strPropertyName As String

        If Date.TryParse(FindText, datFindDate) = False Then Return False

        strPropertyName = selPropInfo.Name
        If String.IsNullOrEmpty(MatchProperty) Then
            boolScanProperty = True
        ElseIf strPropertyName = MatchProperty Then
            'don't verify the current property but first move to the next one
            MatchProperty = String.Empty
            Return False
        End If

        If boolScanProperty = True Then
            Dim objValue As Object = selPropInfo.GetValue(selItem, Nothing)
            If objValue IsNot Nothing Then
                selDate = objValue

                If selDate.Date = datFindDate.Date Then
                    MatchIndex = 0
                    MatchLength = selDate.ToString.Length
                    MatchProperty = strPropertyName

                    Return True
                End If
            End If

            MatchIndex = -1
            MatchLength = 0
            MatchProperty = String.Empty
        End If

        Return False
    End Function

    Private Function FindInItem_Value(ByVal selItem As Object, ByVal selPropInfo As PropertyInfo) As Boolean
        Dim dblFindValue As Double = ParseDouble(FindText)
        Dim dblValue As Double
        Dim boolScanProperty As Boolean
        Dim strPropertyName As String

        strPropertyName = selPropInfo.Name
        If String.IsNullOrEmpty(MatchProperty) Then
            boolScanProperty = True
        ElseIf strPropertyName = MatchProperty Then
            'don't verify the current property but first move to the next one
            MatchProperty = String.Empty
            Return False
        End If

        If boolScanProperty = True Then
            Dim objValue As Object = selPropInfo.GetValue(selItem, Nothing)
            If objValue IsNot Nothing Then
                dblValue = ParseDouble(objValue)

                If dblValue = dblFindValue Then
                    MatchIndex = 0
                    MatchLength = dblValue.ToString.Length
                    MatchProperty = strPropertyName

                    Return True
                End If
            End If

            MatchIndex = -1
            MatchLength = 0
            MatchProperty = String.Empty
        End If

        Return False
    End Function

    Private Function FindInItem_Currency(ByVal selItem As Object, ByVal selPropInfo As PropertyInfo) As Boolean
        Dim objCurrency As Currency = selPropInfo.GetValue(selItem, Nothing)
        Dim sngFindAmount As Single = ParseSingle(FindText)
        Dim strText As String
        Dim boolScanProperty As Boolean
        Dim strPropertyName As String

        strPropertyName = selPropInfo.Name
        If String.IsNullOrEmpty(MatchProperty) Then
            boolScanProperty = True
        ElseIf strPropertyName = MatchProperty Then
            'don't verify the current property but first move to the next one
            MatchProperty = String.Empty
            Return False
        End If

        If boolScanProperty = True Then
            If sngFindAmount = objCurrency.Amount And selItem IsNot MatchChildItem Then
                strText = objCurrency.Amount.ToString

                MatchIndex = 0
                MatchLength = strText.Length
                MatchProperty = selPropInfo.Name

                Return True
            Else
                MatchIndex = -1
                MatchLength = 0
                MatchProperty = String.Empty
                MatchChildItem = Nothing

                Return False
            End If
        End If
    End Function

    Private Function FindInCollection(ByVal objCollection As Object) As Boolean
        If MatchChildItem Is Nothing Then ScanOnChildItem = True

        For Each selObject As Object In objCollection
            If selObject Is MatchChildItem Then ScanOnChildItem = True

            If ScanOnChildItem = True AndAlso FindInItem(selObject) = True Then
                MatchChildItem = selObject
                Return True

                Exit For
            End If
        Next

        Return False
    End Function

    Private Function FindInMatrix(ByVal objMatrix As DoubleValuesMatrix) As Boolean
        For Each selMatrixRow As DoubleValuesMatrixRow In objMatrix
            If FindInCollection(selMatrixRow.DoubleValues) = True Then
                Return True

                Exit For
            End If
        Next

        Return False
    End Function

    Private Function PropertyCanBeModified(ByVal selItem As Object, ByVal selPropInfo As PropertyInfo) As Boolean
        Dim boolCanBeModified As Boolean

        If selPropInfo.MemberType = MemberTypes.Property And selPropInfo.CanWrite = True Then
            'filter out all properties that can't be modified in certain cases
            Select Case selItem.GetType
                Case GetType(ActivityDetail)
                    Select Case selPropInfo.Name
                        Case "Organisation", "Location", "StartDate", "Period", "Duration", "RepeatEvery", "RepeatTimes", "Preparation", "FollowUp"
                            boolCanBeModified = True
                    End Select
                Case GetType(Assumption)
                    Select Case selPropInfo.Name
                        Case "ResponseStrategy", "Owner", "Impact"
                            boolCanBeModified = True
                    End Select
                Case GetType(AssumptionDetail)
                    Select Case selPropInfo.Name
                        Case "Reason", "HowToValidate"
                            boolCanBeModified = True
                    End Select
                Case GetType(Baseline)
                    Dim selBaseline As Baseline = TryCast(selItem, Baseline)

                    Select Case selPropInfo.Name
                        Case "Score"
                            Dim selIndicator As Indicator = Me.Logframe.GetIndicatorByGuid(selBaseline.ParentIndicatorGuid)

                            If selIndicator IsNot Nothing Then
                                Select Case selIndicator.QuestionType
                                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                                        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                                            boolCanBeModified = True
                                        End If
                                End Select
                            End If
                    End Select
                Case GetType(BudgetItem)
                    Select Case selPropInfo.Name
                        Case "Duration", "DurationUnit", "Number", "NumberUnit", "UnitCost", "CurrencyCode"
                            boolCanBeModified = True
                    End Select
                Case GetType(BudgetItemReference)
                    Select Case selPropInfo.Name
                        Case "Percentage", "TotalCost"
                            boolCanBeModified = True
                    End Select
                Case GetType(DependencyDetail)
                    Select Case selPropInfo.Name
                        Case "Deliverables", "Supplier", "DateExpected", "DateDelivered"
                            boolCanBeModified = True
                    End Select
                Case GetType(Indicator)
                    Dim selIndicator As Indicator = TryCast(selItem, Indicator)
                    Select Case selPropInfo.Name
                        Case "Score"
                            If selIndicator IsNot Nothing Then
                                Select Case selIndicator.QuestionType
                                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                                        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                                            boolCanBeModified = True
                                        End If
                                End Select
                            End If
                        Case "PopulationPercentage"
                            If selIndicator IsNot Nothing Then
                                Select Case selIndicator.QuestionType
                                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula, _
                                        Indicator.QuestionTypes.LikertTypeScale, Indicator.QuestionTypes.SemanticDiff, Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.CumulativeScale

                                        boolCanBeModified = False
                                    Case Else
                                        boolCanBeModified = True
                                End Select
                            Else
                                boolCanBeModified = True
                            End If
                        Case Else
                            boolCanBeModified = True
                    End Select
                Case GetType(LogFrame)
                    Select Case selPropInfo.Name
                        Case "ProjectTitle", "ShortTitle", "Code", "Duration", "StartDate", "EndDate"
                            boolCanBeModified = True
                    End Select
                Case GetType(Resource)
                    Select Case selPropInfo.Name
                        Case "TotalCostAmount"
                            boolCanBeModified = True
                    End Select
                Case GetType(Target)
                    Dim selTarget As Target = TryCast(selItem, Target)

                    Select Case selPropInfo.Name
                        Case "Score"
                            Dim selIndicator As Indicator = Me.Logframe.GetIndicatorByGuid(selTarget.ParentIndicatorGuid)

                            If selIndicator IsNot Nothing Then
                                Select Case selIndicator.QuestionType
                                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                                        If selIndicator.ScoringSystem = Indicator.ScoringSystems.Score Then
                                            boolCanBeModified = True
                                        End If
                                End Select
                            End If
                    End Select
                Case Else
                    boolCanBeModified = True
            End Select
        End If

        Return boolCanBeModified
    End Function

    Private Sub Replace()
        Dim strText As String = String.Empty
        

        If MatchChildItem Is Nothing Then
            'text or value in MatchItem
            If String.IsNullOrEmpty(MatchProperty) Then
                Dim selObject As LogframeObject = CType(MatchItem, LogframeObject)

                strText = Replace_InText(selObject.Text)
                selObject.SetText(strText)
            Else
                If MatchItem.GetType Is GetType(Goal) Then
                    Replace_PropertyValue(Me.Logframe)
                Else
                    Replace_PropertyValue(MatchItem)
                End If

            End If
        Else
            'text or value in MatchChildItem
            Replace_PropertyValue(MatchChildItem)
        End If

        MatchLength = Me.ReplaceText.Length
    End Sub

    Private Function Replace_InText(ByVal strText As String) As String
        strText = strText.Remove(MatchIndex, MatchLength)
        strText = strText.Insert(MatchIndex, Me.ReplaceText)

        Return strText
    End Function

    Private Sub Replace_PropertyValue(ByVal selItem As Object)
        Dim objType As Type
        Dim propInfo As PropertyInfo = Nothing

        objType = selItem.GetType
        propInfo = objType.GetProperty(MatchProperty)

        If propInfo IsNot Nothing Then
            Dim objValue As Object = propInfo.GetValue(selItem, Nothing)

            Select Case objValue.GetType
                Case GetType(String)
                    objValue = Replace_InText(objValue.ToString)
                Case GetType(Integer)
                    objValue = ParseInteger(strReplaceText)
                Case GetType(Single)
                    objValue = ParseSingle(strReplaceText)
                Case GetType(Integer)
                    objValue = ParseDouble(strReplaceText)
                Case GetType(Date)
                    Date.TryParse(strReplaceText, objValue)
            End Select

            propInfo.SetValue(selItem, objValue, Nothing)
        End If
    End Sub

    Private Function Replace_GetText() As String
        Dim strText As String = String.Empty

        Return strText
    End Function

    Private Function Replace_GetPropertyValue() As Object
        Dim objType As Type
        Dim propInfo As PropertyInfo = Nothing

        objType = MatchItem.GetType
        propInfo = objType.GetProperty(MatchProperty)
        Dim objValue As Object = propInfo.GetValue(MatchItem, Nothing)

        Return objValue
        'If objValue IsNot Nothing Then
        '    strText = objValue.ToString
        'Else
        '    strText = String.Empty
        'End If
    End Function
#End Region

#Region "Project info"
    Private Function FindReplace_ProjectInfo() As Boolean
        If MatchChildItem Is Nothing And FindInItem(Me.Logframe) = True Then
            MatchItem = Me.Logframe

            Return (True)
        ElseIf FindReplace_ProjectInfo_TargetGroups() = True Then
            MatchItem = Me.Logframe

            Return (True)
        End If

        Return False
    End Function

    Private Function FindReplace_ProjectInfo_TargetGroups() As Boolean
        For Each selPurpose As Purpose In Logframe.Purposes
            If FindReplace_TargetGroups(selPurpose) = True Then
                Return True
            End If
        Next

        Return False
    End Function
#End Region

#Region "Planning"
    Private Function FindReplace_Planning() As Boolean
        For Each selPurpose As Purpose In Logframe.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                If FindReplace_Planning_KeyMoments(selOutput) = True Then
                    Return True
                ElseIf FindReplace_Planning_Activities(selOutput.Activities) = True Then
                    Return True
                End If
            Next
        Next

        Return False
    End Function

    Private Function FindReplace_Planning_KeyMoments(ByVal selOutput As Output) As Boolean

        For Each selKeyMoment As KeyMoment In selOutput.KeyMoments
            If selKeyMoment Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If FindInItem(selKeyMoment) = True Then
                    MatchItem = selKeyMoment

                    Return True
                End If
            End If
        Next

        Return False
    End Function

    Private Function FindReplace_Planning_Activities(ByVal colActivities As Activities) As Boolean
        For Each selActivity As Activity In colActivities
            If selActivity Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If SearchAllPanes = False Then
                    If FindInText(selActivity) = True Then
                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                Else
                    If MatchChildItem Is Nothing Then
                        If FindReplace_Struct_Normal(selActivity) = True Then Return True
                    Else
                        If FindReplace_Struct_FromStartProperty(selActivity) = True Then Return True
                    End If

                End If
            End If

            If selActivity.Activities.Count > 0 Then
                If FindReplace_Planning_Activities(selActivity.Activities) = True Then Return True
            End If
        Next

        Return False
    End Function
#End Region

#Region "Budget year"
    Private Function FindReplace_BudgetYear(ByVal intBudgetYearIndex As Integer) As Boolean
        Dim selBudgetYear As BudgetYear = CurrentBudget.BudgetYears(intBudgetYearIndex)

        If selBudgetYear IsNot Nothing Then
            If FindReplace_BudgetItems(selBudgetYear.BudgetItems) = True Then
                Return True
            End If
        End If

        Return False
    End Function

    Private Function FindReplace_BudgetItems(ByVal colBudgetItems As BudgetItems) As Boolean
        For Each selBudgetItem As BudgetItem In colBudgetItems
            If selBudgetItem Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If String.IsNullOrEmpty(MatchProperty) AndAlso FindInText(selBudgetItem) = True Then
                    Return True

                ElseIf FindInItem(selBudgetItem) = True Then
                    MatchItem = selBudgetItem

                    Return True
                Else
                    MatchIndex = -1
                    MatchLength = 0
                    MatchChildItem = Nothing

                End If

                'End If
            End If

            If selBudgetItem.BudgetItems.Count > 0 Then
                If FindReplace_BudgetItems(selBudgetItem.BudgetItems) = True Then Return True
            End If
        Next
    End Function
#End Region

#Region "Structs"
    Private Function FindReplace_Structs(ByVal colStructs As Structs) As Boolean
        For Each selStruct As Struct In colStructs
            If selStruct Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If SearchAllPanes = False Then
                    If FindInText(selStruct) = True Then
                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                Else
                    If MatchChildItem Is Nothing Then
                        If FindReplace_Struct_Normal(selStruct) = True Then Return True
                    Else
                        If FindReplace_Struct_FromStartProperty(selStruct) = True Then Return True
                    End If

                End If
            End If

            If selStruct.Indicators.Count > 0 Then
                If FindReplace_Indicators(selStruct.Indicators) = True Then Return True
            End If

            If selStruct.Assumptions.Count > 0 Then
                If FindReplace_Assumptions(selStruct.Assumptions) = True Then Return True
            End If
        Next

        Return False
    End Function

    Private Function FindReplace_Struct_Normal(ByVal selStruct As Struct) As Boolean
        If String.IsNullOrEmpty(MatchProperty) AndAlso FindInText(selStruct) = True Then
            Return True

        Else
            Select Case selStruct.GetType
                Case GetType(Goal)
                    Dim selGoal As Goal = DirectCast(selStruct, Goal)

                    If Me.Logframe.Goals.IndexOf(selGoal) = 0 AndAlso FindInItem(Me.Logframe) = True Then
                        MatchItem = selStruct

                        Return (True)
                    End If
                Case GetType(Purpose)
                    Dim selPurpose As Purpose = DirectCast(selStruct, Purpose)
                    If FindReplace_TargetGroups(selPurpose) = True Then
                        MatchItem = selStruct

                        Return (True)
                    End If
                Case GetType(Output)
                    Dim selOutput As Output = DirectCast(selStruct, Output)
                    If FindReplace_KeyMoments(selOutput) = True Then
                        MatchItem = selStruct

                        Return (True)
                    End If
                Case GetType(Activity)
                    Dim selActivity As Activity = DirectCast(selStruct, Activity)
                    If FindInItem(selActivity.ActivityDetail) = True Then
                        MatchItem = selActivity
                        MatchChildItem = selActivity.ActivityDetail

                        Return (True)
                    End If
            End Select
        End If

        MatchIndex = -1
        MatchLength = 0
        MatchChildItem = Nothing

        Return False
    End Function

    Private Function FindReplace_Struct_FromStartProperty(ByVal selStruct As Struct) As Boolean
        Select Case MatchChildItem.GetType
            Case GetType(TargetGroup)
                If selStruct Is MatchItem AndAlso FindReplace_TargetGroups(selStruct) = True Then
                    MatchItem = selStruct

                    Return True
                End If
            Case GetType(KeyMoment)
                If selStruct Is MatchItem AndAlso FindReplace_KeyMoments(selStruct) = True Then
                    MatchItem = selStruct

                    Return True
                End If
            Case GetType(ActivityDetail)
                Dim selActivity As Activity = DirectCast(selStruct, Activity)
                If selStruct Is MatchItem AndAlso FindInItem(selActivity.ActivityDetail) = True Then
                    MatchItem = selActivity
                    MatchChildItem = selActivity.ActivityDetail

                    Return True
                End If
        End Select

        MatchIndex = -1
        MatchLength = 0
        MatchChildItem = Nothing

        Return False
    End Function
#End Region

#Region "Goals and project information"
    Private Function FindReplace_Goals() As Boolean
        If FindReplace_Structs(Logframe.Goals) = True Then
            Return True
        End If

        Return False
    End Function
#End Region

#Region "Purposes and target groups"
    Private Function FindReplace_Purposes() As Boolean
        If FindReplace_Structs(Logframe.Purposes) = True Then
            Return True
        End If

        Return False
    End Function

    Private Function FindReplace_TargetGroups(ByVal selPurpose As Purpose) As Boolean
        Dim boolScan As Boolean

        If MatchChildItem Is Nothing Then
            boolScan = True
        Else
            If MatchChildItem.GetType IsNot GetType(TargetGroupInformation) Then boolScan = True
        End If

        For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
            If boolScan = True AndAlso FindInItem(selTargetGroup) = True Then
                MatchChildItem = selTargetGroup

                Return True
            ElseIf FindInCollection(selTargetGroup.TargetGroupInformations) = True Then
                MatchItem = selPurpose

                Return True
            End If
        Next

        Return False
    End Function
#End Region

#Region "Outputs and key moments"
    Private Function FindReplace_Outputs() As Boolean
        For Each selPurpose As Purpose In Logframe.Purposes
            If FindReplace_Structs(selPurpose.Outputs) = True Then Return True
        Next
    End Function

    Private Function FindReplace_KeyMoments(ByVal selOutput As Output) As Boolean
        If FindInCollection(selOutput.KeyMoments) = True Then
            Return True
        End If

        Return False
    End Function
#End Region

#Region "Processes, activities and activity details"
    Private Function FindReplace_ActivitiesInLogframe() As Boolean
        'First pass analyses activities, their indicators or resources, their verificationsources or budgettotals and their assumptions
        For Each selPurpose As Purpose In Logframe.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                If FindReplace_Activities(selOutput.Activities) = True Then Return True
            Next
        Next

        If SearchPastEnd = False Then
            SearchPastEnd = True
        Else
            SearchPastEnd = False
            Return False
        End If
        If SearchResources = True Then SearchResources = False Else boolSearchResources = True

        'Second pass only looks in resources or indicators and their budget totals or verification sources. 
        'Activities and assumptions have been analysed already and will be skipped during this pass
        If FindReplace_ActivitiesInLogframe() = True Then
            SearchPastEnd = False
            Return True
        End If

        Return False
    End Function

    Private Function FindReplace_Activities(ByVal colActivities As Activities) As Boolean
        For Each selActivity As Activity In colActivities
            If selActivity Is Me.StartItem Then boolScanOn = True

            If SearchPastEnd = False And ScanOn = True Then
                If SearchAllPanes = False Then
                    If FindInText(selActivity) = True Then
                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                Else
                    If MatchChildItem Is Nothing Then
                        If FindReplace_Struct_Normal(selActivity) = True Then Return True
                    Else
                        If FindReplace_Struct_FromStartProperty(selActivity) = True Then Return True
                    End If

                End If
            End If

            If Me.SearchResources = False Then
                If selActivity.Indicators.Count > 0 Then
                    If FindReplace_Indicators(selActivity.Indicators) = True Then Return True
                End If
            Else
                If selActivity.Resources.Count > 0 Then
                    If FindReplace_Resources(selActivity.Resources) = True Then Return True
                End If
            End If

            If SearchPastEnd = False Then
                If selActivity.Assumptions.Count > 0 Then
                    If FindReplace_Assumptions(selActivity.Assumptions) = True Then Return True
                End If
            End If

            If selActivity.Activities.Count > 0 Then
                If FindReplace_Activities(selActivity.Activities) = True Then Return True
            End If
        Next

        Return False
    End Function
#End Region

#Region "Indicators"
    Private Function FindReplace_Indicators(ByVal colIndicators As Indicators) As Boolean
        For Each selIndicator As Indicator In colIndicators
            If selIndicator Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If SearchAllPanes = False Then
                    If FindInText(selIndicator) = True Then
                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                Else
                    If MatchChildItem Is Nothing Then
                        If FindReplace_Indicator_Normal(selIndicator) = True Then Return True
                    Else
                        If FindReplace_Indicator_FromStartProperty(selIndicator) = True Then Return True
                    End If

                End If
            End If

            If selIndicator.VerificationSources.Count > 0 Then
                If FindReplace_VerificationSources(selIndicator.VerificationSources) = True Then Return True
            End If

            If selIndicator.Indicators.Count > 0 Then
                If FindReplace_Indicators(selIndicator.Indicators) = True Then Return True
            End If
        Next

        Return False
    End Function

    Private Function FindReplace_Indicator_Normal(ByVal selIndicator As Indicator) As Boolean
        If String.IsNullOrEmpty(MatchProperty) AndAlso FindInText(selIndicator) = True Then
            Return True

        ElseIf FindInItem(selIndicator.ValuesDetail) = True Then
            MatchItem = selIndicator

            Return True
        ElseIf FindInItem(selIndicator.ScalesDetail) = True Then
            MatchItem = selIndicator

            Return True
        ElseIf FindInCollection(selIndicator.ResponseClasses) = True Then
            MatchItem = selIndicator

            Return (True)
        ElseIf FindReplace_Statements(selIndicator) = True Then
            MatchItem = selIndicator

            Return (True)
        ElseIf FindReplace_Baseline(selIndicator) = True Then
            MatchItem = selIndicator

            Return True
        ElseIf FindReplace_Targets(selIndicator) = True Then
            MatchItem = selIndicator

            Return (True)
        ElseIf FindInCollection(selIndicator.PopulationTargets) = True Then
            MatchItem = selIndicator

            Return (True)
        Else
            MatchIndex = -1
            MatchLength = 0
            MatchChildItem = Nothing

        End If

        Return False
    End Function

    Private Function FindReplace_Indicator_FromStartProperty(ByVal selIndicator As Indicator) As Boolean
        Select Case MatchChildItem.GetType
            Case GetType(ValuesDetail)
                If FindInItem(selIndicator.ValuesDetail) = True Then
                    MatchItem = selIndicator

                    Return True
                End If
            Case GetType(ScalesDetail)
                If FindInItem(selIndicator.ScalesDetail) = True Then
                    MatchItem = selIndicator

                    Return True
                End If
            Case GetType(ResponseClass)
                If FindInCollection(selIndicator.ResponseClasses) = True Then
                    MatchItem = selIndicator

                    Return (True)
                End If
            Case GetType(Statement)
                If FindReplace_Statements(selIndicator) = True Then
                    MatchItem = selIndicator

                    Return (True)
                End If
            Case GetType(Baseline)
                If FindReplace_Baseline(selIndicator) = True Then
                    MatchItem = selIndicator

                    Return True
                End If
            Case GetType(Target)
                If FindReplace_Targets(selIndicator) = True Then
                    MatchItem = selIndicator

                    Return (True)
                End If
            Case GetType(PopulationTarget)
                If FindInCollection(selIndicator.PopulationTargets) = True Then
                    MatchItem = selIndicator

                    Return (True)
                End If
        End Select

        MatchIndex = -1
        MatchLength = 0
        MatchChildItem = Nothing

        Return False
    End Function

    Private Function FindReplace_Statements(ByVal selIndicator As Indicator) As Boolean
        If selIndicator.QuestionType = Indicator.QuestionTypes.LikertScale Or selIndicator.QuestionType = Indicator.QuestionTypes.FrequencyLikert Then
            'statements can't be modified if these types of indicators have child indicators
            If selIndicator.Indicators.Count > 0 Then
                Return False
            End If
        End If

        If MatchChildItem Is Nothing Then
            ScanOnChildItem = True
            MatchIndex = -1
            MatchLength = 0
            MatchProperty = String.Empty
        End If

        For Each selStatement As Statement In selIndicator.Statements
            If selStatement Is MatchChildItem Then
                If MatchIndex >= 0 Then MatchIndex += MatchLength
                ScanOnChildItem = True
            End If


            If ScanOnChildItem = True Then
                If MatchProperty <> "SecondLabel" AndAlso FindInItem_RichText(selStatement.FirstLabel) = True Then
                    MatchChildItem = selStatement
                    MatchProperty = "FirstLabel"
                    Return True
                ElseIf FindInItem_RichText(selStatement.SecondLabel) = True Then
                    MatchChildItem = selStatement
                    MatchProperty = "SecondLabel"
                    Return True
                Else
                    MatchIndex = -1
                    MatchLength = 0
                    MatchProperty = String.Empty
                End If
            End If
        Next

        Return False
    End Function

    Private Function FindReplace_Baseline(ByVal selIndicator As Indicator) As Boolean
        If FindInItem(selIndicator.Baseline) = True Then
            MatchChildItem = selIndicator.Baseline
            Return True
        ElseIf selIndicator.Baseline.DoubleValues.Count > 0 AndAlso FindInCollection(selIndicator.Baseline.DoubleValues) = True Then
            Return True
        ElseIf selIndicator.Baseline.DoubleValuesMatrix.Count > 0 AndAlso FindInMatrix(selIndicator.Baseline.DoubleValuesMatrix) = True Then
            Return True
        End If

        Return False
    End Function

    Private Function FindReplace_Targets(ByVal selIndicator As Indicator) As Boolean
        If FindInCollection(selIndicator.Targets) = True Then
            Return True
        Else
            For Each selTarget As Target In selIndicator.Targets
                If FindInCollection(selTarget.DoubleValues) = True Then
                    Return True
                End If
            Next
        End If

        Return False
    End Function

    Private Function FindReplace_VerificationSources(ByVal colVerificationSources As VerificationSources) As Boolean
        For Each selVerificationSource As VerificationSource In colVerificationSources
            If selVerificationSource Is Me.StartItem Then boolScanOn = True

            If ScanOn = True And String.IsNullOrEmpty(selVerificationSource.Text) = False Then
                If FindInText(selVerificationSource) = True Then
                    Return True
                ElseIf SearchAllPanes = True AndAlso FindInItem(selVerificationSource) = True Then
                    MatchItem = selVerificationSource

                    Return True
                Else
                    MatchIndex = -1
                    MatchLength = 0
                End If
            End If
        Next

        Return False
    End Function
#End Region

#Region "Resources and budget item references"
    Private Function FindReplace_Resources(ByVal colResources As Resources) As Boolean
        For Each selResource As Resource In colResources
            If selResource Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If SearchAllPanes = False Then
                    If FindInText(selResource) = True Then
                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                Else
                    If String.IsNullOrEmpty(MatchProperty) AndAlso FindInText(selResource) = True Then
                        Return True
                    ElseIf FindInItem(selResource) = True Then
                        MatchItem = selResource

                        Return True
                    ElseIf FindInCollection(selResource.BudgetItemReferences) = True Then
                        MatchItem = selResource
                        'MatchChildItem = selResource.ResourceDetail

                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                End If
            End If
        Next

        Return False
    End Function
#End Region

#Region "Assumptions"
    Private Function FindReplace_Assumptions(ByVal colAssumptions As Assumptions) As Boolean
        For Each selAssumption As Assumption In colAssumptions
            If selAssumption Is Me.StartItem Then boolScanOn = True

            If ScanOn = True Then
                If SearchAllPanes = False Then
                    If FindInText(selAssumption) = True Then
                        Return True
                    Else
                        MatchIndex = -1
                        MatchLength = 0
                    End If
                Else
                    If MatchChildItem Is Nothing Then
                        If FindReplace_Assumption_Normal(selAssumption) = True Then Return True
                    Else
                        If FindReplace_Assumption_FromStartProperty(selAssumption) = True Then Return True
                    End If
                End If
            End If
        Next

        Return False
    End Function

    Private Function FindReplace_Assumption_Normal(ByVal selAssumption As Assumption) As Boolean
        If String.IsNullOrEmpty(MatchProperty) AndAlso FindInText(selAssumption) = True Then
            Return True
        ElseIf FindInItem(selAssumption) = True Then
            MatchItem = selAssumption

            Return True
        ElseIf FindInItem(selAssumption.AssumptionDetail) = True Then
            MatchItem = selAssumption

            Return True
        ElseIf FindInItem(selAssumption.DependencyDetail) = True Then
            MatchItem = selAssumption

            Return True
            'ElseIf FindInItem(selAssumption.RiskDetail) = True Then --> not necessary because no values can be changed in RiskDetail

        Else
            MatchIndex = -1
            MatchLength = 0
            MatchChildItem = Nothing
        End If

        Return False
    End Function

    Private Function FindReplace_Assumption_FromStartProperty(ByVal selAssumption As Assumption) As Boolean
        Select Case MatchChildItem.GetType
            Case GetType(AssumptionDetail)
                If FindInItem(selAssumption.AssumptionDetail) = True Then
                    MatchItem = selAssumption

                    Return True
                End If
            Case GetType(DependencyDetail)
                If FindInItem(selAssumption.DependencyDetail) = True Then
                    MatchItem = selAssumption

                    Return True
                End If
            Case GetType(RiskDetail)
                If FindInItem(selAssumption.RiskDetail) = True Then
                    MatchItem = selAssumption

                    Return (True)
                End If
        End Select

        MatchIndex = -1
        MatchLength = 0
        MatchChildItem = Nothing

        Return False
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
