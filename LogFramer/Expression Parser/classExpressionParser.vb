Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class ExpressionParser
    Implements System.IDisposable

    Private strExpression As String
    'Public LstElements As New ArrayList

    Public Event ParenthesisMissing()
    Public Event DivisionByZero()
    Public Event ErrorParsingDecimal()

    Private Const REGEX_DEC As String = "[0-9]*([\.\,][0-9]+)?"
    Private PowerCharsList As New List(Of Integer)
    Private objVariableList As New Dictionary(Of String, Double)
    Private objFormulaList As New Dictionary(Of String, String)
    Private objUnitList As New Dictionary(Of String, String)

    Public Sub New()

    End Sub

    Public Sub New(ByVal expression As String)
        Me.Expression = expression
    End Sub

#Region "Properties"
    Public Property Expression As String
        Get
            Return strExpression
        End Get
        Set(ByVal value As String)
            If String.IsNullOrEmpty(value) Then
                strExpression = String.Empty
            Else
                strExpression = value.ToUpper
            End If
        End Set
    End Property

    Public Property VariableList As Dictionary(Of String, Double)
        Get
            Return objVariableList
        End Get
        Set(ByVal value As Dictionary(Of String, Double))
            objVariableList = value
        End Set
    End Property

    Public Property FormulaList As Dictionary(Of String, String)
        Get
            Return objFormulaList
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            objFormulaList = value
        End Set
    End Property

    Public Property UnitList As Dictionary(Of String, String)
        Get
            Return objUnitList
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            objUnitList = value
        End Set
    End Property
#End Region

#Region "Parse values"
    Public Function Parse() As Double
        Dim dblSolution As Double
        Dim selExpression As String = Me.Expression

        If String.IsNullOrEmpty(selExpression) Then
            Return 0
        Else
            dblSolution = ParseFormulas(selExpression)

            Return dblSolution
        End If
    End Function

    Private Function ParseFormulas(ByVal selExpression As String) As Double
        Dim dblSolution As Double

        If FormulaList.Count > 0 Then
            For Each selObj As KeyValuePair(Of String, String) In FormulaList
                Dim strVariable As String = selObj.Key
                Dim strFormula As String = String.Format("({0})", selObj.Value)

                Do Until selExpression.Contains(strVariable) = False
                    selExpression = selExpression.Replace(strVariable, strFormula)
                Loop
            Next
        End If

        dblSolution = ParseVariables(selExpression)

        Return dblSolution
    End Function

    Private Function ParseVariables(ByVal selExpression As String) As Double
        Dim dblSolution As Double

        If VariableList.Count > 0 Then
            For Each selObj As KeyValuePair(Of String, Double) In VariableList
                Dim strVariable As String = selObj.Key

                Do Until selExpression.Contains(strVariable) = False
                    selExpression = selExpression.Replace(strVariable, selObj.Value.ToString)
                Loop
            Next
        End If

        dblSolution = ParseBrackets(selExpression)

        Return dblSolution
    End Function

    Private Function ParseBrackets(ByVal selExpression As String) As Double
        Dim intIndex As Integer
        Dim strSelection, strReplace As String
        Dim dblValue As Double

        Do
            intIndex = selExpression.LastIndexOf("(")
            If intIndex = -1 Then
                dblValue = ParsePower(selExpression)
                Exit Do
            Else
                strSelection = selExpression.Substring(intIndex)
                intIndex = strSelection.IndexOf(")")

                If intIndex = -1 Then
                    RaiseEvent ParenthesisMissing()
                    Return 0
                End If

                strReplace = strSelection.Substring(0, intIndex + 1)
                strSelection = strReplace.Substring(1, strReplace.Length - 2)
                'LstElements.Add(strSelection)

                dblValue = ParsePower(strSelection)
                selExpression = selExpression.Replace(strReplace, dblValue.ToString)
            End If
        Loop

        Return dblValue
    End Function

    Private Function ParsePower(ByVal selExpression As String) As Double
        Dim intIndex As Integer
        Dim strSelection As String
        Dim dblValue, dblValue1, dblValue2 As Double
        Dim FindChar As Char = "^"
        Dim pattern As String
        Dim match As Match
        Dim FirstTerm, LastTerm As String

        Do
            pattern = String.Format("{0}\s*\^\s*{1}", REGEX_DEC, REGEX_DEC) 'number with optional decimal followed by zero or more spaces followed by '^' and again zero or more spaces and a number with optional decimal
            match = Regex.Match(selExpression, pattern)

            If match.Success Then
                strSelection = match.Value
                intIndex = strSelection.IndexOf(FindChar)

                FirstTerm = strSelection.Substring(0, intIndex)
                dblValue1 = ParseDouble(FirstTerm)
                LastTerm = strSelection.Substring(intIndex + 1)
                dblValue2 = ParseDouble(LastTerm)

                dblValue = dblValue1 ^ dblValue2

                selExpression = selExpression.Replace(strSelection, dblValue.ToString)
            Else
                dblValue = ParseMultDiv(selExpression)
                Exit Do
            End If

        Loop

        Return dblValue
    End Function

    Private Function ParseMultDiv(ByVal selExpression As String) As Double
        Dim intIndex As Integer
        Dim strSelection As String, strOperator As String = String.Empty
        Dim dblValue, dblValue1, dblValue2 As Double
        Dim FindChar() As Char = New Char() {"*", "/"}
        Dim pattern As String
        Dim match As Match
        Dim FirstTerm, LastTerm As String

        Do
            pattern = String.Format("{0}\s*[\*\/]\s*{1}", REGEX_DEC, REGEX_DEC) 'number with optional decimal followed by zero or more spaces followed by '*' or '/' and again zero or more spaces and a number with optional decimal
            match = Regex.Match(selExpression, pattern)

            If match.Success Then
                strSelection = match.Value
                intIndex = strSelection.IndexOfAny(FindChar)
                strOperator = strSelection.Substring(intIndex, 1)

                FirstTerm = strSelection.Substring(0, intIndex)
                dblValue1 = ParseDouble(FirstTerm)
                LastTerm = strSelection.Substring(intIndex + 1)
                dblValue2 = ParseDouble(LastTerm)

                If strOperator = "*" Then
                    dblValue = dblValue1 * dblValue2
                Else
                    If dblValue2 = 0 Then RaiseEvent DivisionByZero()
                    dblValue = dblValue1 / dblValue2
                End If

                selExpression = selExpression.Replace(strSelection, dblValue.ToString)
            Else
                dblValue = ParsePlusMinus(selExpression)
                Exit Do
            End If

        Loop

        Return dblValue
    End Function

    Private Function ParsePlusMinus(ByVal selExpression As String) As Double
        Dim intIndex As Integer
        Dim strSelection As String, strOperator As String = String.Empty
        Dim dblValue, dblSum As Double
        Dim FindChar() As Char = New Char() {"-", "+"}

        Do
            intIndex = selExpression.IndexOfAny(FindChar)

            If intIndex = -1 Then
                dblValue = ParseDouble(selExpression)

                If strOperator = "+" Then
                    dblSum += dblValue
                ElseIf strOperator = "-" Then
                    dblSum -= dblValue
                Else
                    dblSum = dblValue
                End If
                Exit Do
            Else

                strSelection = selExpression.Substring(0, intIndex)
                dblValue = ParseDouble(strSelection)

                If strOperator = "+" Then
                    dblSum += dblValue
                ElseIf strOperator = "-" Then
                    dblSum -= dblValue
                Else
                    dblSum = dblValue
                End If

                strOperator = selExpression.Substring(intIndex, 1)
                selExpression = selExpression.Substring(intIndex + 1)
            End If

        Loop

        Return dblSum
    End Function

    Public Function ParseDouble(ByVal strNumber) As Double
        Dim dblValue As Double
        Dim style As NumberStyles = NumberStyles.Float
        Dim culture As CultureInfo = CultureInfo.CurrentCulture

        If Double.TryParse(strNumber, style, culture, dblValue) = False Then
            culture = CultureInfo.InvariantCulture
            If Double.TryParse(strNumber, style, culture, dblValue) = False Then
                RaiseEvent ErrorParsingDecimal()
            End If
        End If
        Return dblValue
    End Function
#End Region

#Region "Parse units"
    Public Function ParseUnits() As String
        Dim strUnit As String
        Dim selExpression As String = Me.Expression

        PowerCharsList.Clear()
        PowerCharsList.AddRange({&H2070, &HB9, &HB2, &HB3, &H2074, &H2075, &H2076, &H2077, &H2078, &H2079})

        If String.IsNullOrEmpty(selExpression) Then
            Return String.Empty
        Else
            'Replace variables in the formula by their respective units
            strUnit = ParseUnitVariables(selExpression)

            'Powers are still represented by lowercase numbers -> replace them by uppercase numbers
            strUnit = ParseUnitIntegersToPowers(strUnit)
            Return strUnit
        End If
    End Function

    Private Function ParseUnitVariables(ByVal selExpression As String) As String
        Dim strReplace As String
        Dim pattern As String

        If UnitList.Count > 0 Then
            For Each selObj As KeyValuePair(Of String, String) In UnitList
                Dim strVariable As String = selObj.Key

                

                If selObj.Value IsNot Nothing Then
                    pattern = String.Format("(?<=[\+-\/\*\(\^]|\A)\s*{0}\s*(?=[\+-\/\*\(\^]|\z)", strVariable)
                    'the name of the variable (A, B, C...) must at the start of the string or be preceded by a mathematical operator (- + /...) 
                    'and be directly followed by a mathematical operator or be at the end of the string
                    strReplace = selObj.Value

                    selExpression = Regex.Replace(selExpression, pattern, strReplace)
                Else
                    pattern = String.Format("(([\+-\/\*\(\^])\s*{0})|(^{0}\s*[\+-\/\*\(\^])", strVariable)
                    'if the variable is at the beginning of the string, remove it and the operator that follows as well
                    'if the variable is found later in the string, remove it and the operator that preceeds it

                    selExpression = Regex.Replace(selExpression, pattern, String.Empty)
                End If
            Next
        End If

        'if there are any units expressed as a power (m²), then replace the uppercase power by an integer for processing
        selExpression = ParseUnitSuperScriptPowers(selExpression)


        'First process any brackets in the formula
        selExpression = ParseUnitBrackets(selExpression)

        Return selExpression
    End Function

    Private Function ParseUnitSuperScriptPowers(ByVal selExpression As String) As String
        Dim charPower As Char

        For i = 0 To PowerCharsList.Count - 1
            charPower = ChrW(PowerCharsList(i))

            If selExpression.Contains(charPower) Then
                selExpression = selExpression.Replace(charPower, i.ToString)
            End If
        Next
        Return selExpression
    End Function

    Private Function ParseUnitBrackets(ByVal selExpression As String) As String
        Dim intIndex As Integer
        Dim strSelection, strReplace As String
        Dim strUnit As String

        Do
            intIndex = selExpression.LastIndexOf("(")
            If intIndex = -1 Then
                strUnit = ParseUnitPower(selExpression)
                Exit Do
            Else
                strSelection = selExpression.Substring(intIndex)
                intIndex = strSelection.IndexOf(")")

                If intIndex = -1 Then
                    RaiseEvent ParenthesisMissing()
                    Return "#ERROR - parentheses"
                End If

                strReplace = strSelection.Substring(0, intIndex + 1)
                strSelection = strReplace.Substring(1, strReplace.Length - 2)

                strUnit = ParseUnitPower(strSelection)
                selExpression = selExpression.Replace(strReplace, strUnit.ToString)
            End If
        Loop

        Return strUnit
    End Function

    Public Function ParseUnitPower(ByVal selExpression As String) As String
        Dim pattern As String
        Dim match As Match
        Dim intPower As Integer = 1
        Dim strSelection, strReplace As String
        Dim intIndexPower As Integer
        Dim strFirst, strSecond As String
        Dim strFirstPower As String
        Dim intFirstPower, intSecondPower As Integer

        Do
            pattern = "\w+\s*\^\s*\d+" 'meter ^ 2
            Match = Regex.Match(selExpression, pattern)

            If match.Success Then
                strSelection = match.Value
                intIndexPower = strSelection.IndexOf("^")
                strFirst = strSelection.Substring(0, intIndexPower).Trim
                strSecond = strSelection.Substring(intIndexPower + 1).Trim

                If strFirst.Length > 1 Then
                    strFirstPower = strFirst.Substring(strFirst.Length - 1)

                    If Integer.TryParse(strFirstPower, intFirstPower) Then
                        strFirst = strFirst.Remove(strFirst.Length - 1)
                    Else
                        intFirstPower = 1
                    End If
                End If

                If Integer.TryParse(strSecond, intSecondPower) = False Then
                    intSecondPower = 1
                End If

                strReplace = strFirst.ToLower
                intPower = intFirstPower * intSecondPower
                If intPower > 1 And intPower <= 9 Then strReplace &= intPower.ToString

                selExpression = selExpression.Replace(strSelection, strReplace)
            Else
                selExpression = ParseUnitMultiply(selExpression)
                Exit Do
            End If

        Loop

        Return selExpression
    End Function

    Public Function ParseUnitMultiply(ByVal selExpression As String) As String
        Dim pattern As String
        Dim match As Match
        Dim intPower As Integer = 1
        Dim objTextInfo As TextInfo = New CultureInfo("en-US", False).TextInfo
        Dim strSelection, strReplace As String
        Dim intIndexMultiply As Integer
        Dim strFirst, strSecond As String
        Dim strFirstPower, strSecondPower As String
        Dim intFirstPower, intSecondPower As Integer

        Do
            pattern = "\b\w+\s*\*\s*\w+\b"
            Match = Regex.Match(selExpression, pattern)

            If match.Success Then
                intFirstPower = 1
                intSecondPower = 1
                strSelection = match.Value
                intIndexMultiply = strSelection.IndexOf("*")
                strFirst = strSelection.Substring(0, intIndexMultiply).Trim
                strSecond = strSelection.Substring(intIndexMultiply + 1).Trim

                If strFirst.Length > 1 Then
                    strFirstPower = strFirst.Substring(strFirst.Length - 1)

                    If Integer.TryParse(strFirstPower, intFirstPower) Then
                        strFirst = strFirst.Remove(strFirst.Length - 1)
                    End If
                End If

                If strSecond.Length > 1 Then
                    strSecondPower = strSecond.Substring(strSecond.Length - 1)

                    If Integer.TryParse(strSecondPower, intSecondPower) Then
                        strSecond = strSecond.Remove(strSecond.Length - 1)
                    End If
                End If

                If strFirst.ToLower = strSecond.ToLower Then
                    strReplace = strFirst.ToLower
                    intPower = intFirstPower + intSecondPower
                    If intPower <= 9 Then _
                        strReplace &= intPower.ToString
                Else
                    strReplace = objTextInfo.ToTitleCase(strFirst) & objTextInfo.ToTitleCase(strSecond)
                End If

                selExpression = selExpression.Replace(strSelection, strReplace)
            Else
                selExpression = ParseUnitDivide(selExpression)
                Exit Do
            End If

        Loop

        Return selExpression
    End Function

    Public Function ParseUnitDivide(ByVal selExpression As String) As String
        Dim pattern As String
        Dim match As Match
        Dim intPower As Integer = 1
        Dim objTextInfo As TextInfo = New CultureInfo("en-US", False).TextInfo
        Dim strSelection, strReplace As String
        Dim intIndexDivide As Integer
        Dim strFirst, strSecond As String
        Dim strFirstPower, strSecondPower As String
        Dim intFirstPower, intSecondPower As Integer

        Do
            'pattern = "\b\w+\s*\/\s*\w+\b"
            pattern = "(\w+)\d?\s*\/\s*\1"
            match = Regex.Match(selExpression, pattern)

            If match.Success Then
                intFirstPower = 1
                intSecondPower = 1
                strSelection = match.Value
                intIndexDivide = strSelection.IndexOf("/")
                strFirst = strSelection.Substring(0, intIndexDivide).Trim
                strSecond = strSelection.Substring(intIndexDivide + 1).Trim

                If strFirst.Length > 1 Then
                    strFirstPower = strFirst.Substring(strFirst.Length - 1)

                    If Integer.TryParse(strFirstPower, intFirstPower) Then
                        strFirst = strFirst.Remove(strFirst.Length - 1)
                    End If
                End If

                If strSecond.Length > 1 Then
                    strSecondPower = strSecond.Substring(strSecond.Length - 1)

                    If Integer.TryParse(strSecondPower, intSecondPower) Then
                        strSecond = strSecond.Remove(strSecond.Length - 1)
                    End If
                End If

                If strFirst.ToLower = strSecond.ToLower Then
                    strReplace = strFirst.ToLower
                    intPower = intFirstPower - intSecondPower
                    If intPower = 0 Then
                        strReplace = String.Empty
                    ElseIf intPower > 1 And intPower <= 9 Then
                        strReplace &= intPower.ToString
                    End If

                    selExpression = selExpression.Replace(strSelection, strReplace)
                End If


            Else
                selExpression = ParseUnitPlusMinus(selExpression)
                Exit Do
            End If

        Loop

        Return selExpression
    End Function

    Public Function ParseUnitPlusMinus(ByVal selExpression As String) As String
        Dim pattern As String


        pattern = "(\w+\d?\s*\/\s*\w+\d?)\s*[+-]\s*\1" 'sum of ratios (e.g. kg/ha + kg/ha)
        selExpression = ParseUnitPlusMinus_Verify(selExpression, pattern)
        pattern = "(\w+\d?)\s*[+-]\s*\1" 'sum of single units (e.g. kg + kg; m² + m²; etc.)
        selExpression = ParseUnitPlusMinus_Verify(selExpression, pattern)


        Return selExpression
    End Function

    Private Function ParseUnitPlusMinus_Verify(ByVal selExpression As String, ByVal pattern As String) As String
        Dim match As Match
        Dim strSelection, strReplace As String
        Dim intIndexDivide As Integer
        Dim strFirst, strSecond As String
        Dim FindChar() As Char = New Char() {"-", "+"}

        Do
            Match = Regex.Match(selExpression, pattern)

            If Match.Success Then
                strSelection = Match.Value
                intIndexDivide = strSelection.IndexOfAny(FindChar) 'selExpression.IndexOfAny(FindChar) ---> geeft fout
                strFirst = strSelection.Substring(0, intIndexDivide).Trim
                strFirst = ParseUnitPlusMinus_RemoveSign(strFirst)

                If intIndexDivide <= strSelection.Length - 1 Then
                    strSecond = strSelection.Substring(intIndexDivide).Trim
                    strSecond = ParseUnitPlusMinus_RemoveSign(strSecond)
                Else
                    Exit Do
                End If

                If strFirst.ToLower = strSecond.ToLower Then
                    strReplace = strFirst.ToLower

                    selExpression = selExpression.Replace(strSelection, strReplace)
                End If
                Exit Do
            Else
                Exit Do
            End If

        Loop

        Return selExpression
    End Function

    Private Function ParseUnitPlusMinus_RemoveSign(ByVal strExpression As String) As String
        strExpression = strExpression.Replace("+", "")
        strExpression = strExpression.Replace("-", "")

        Return strExpression
    End Function

    Private Function ParseUnitIntegersToPowers(ByVal selExpression As String)
        Dim pattern As String = "\d"
        Dim match As Match
        Dim strValue As String
        Dim intValue, intIndex As Integer
        Dim strReplace As String

        For Each match In Regex.Matches(selExpression, pattern)
            strValue = match.Value

            If Integer.TryParse(strValue, intValue) Then
                intIndex = PowerCharsList(intValue)
                strReplace = ChrW(intIndex)

                selExpression = selExpression.Replace(strValue, strReplace)
            End If
        Next
        Return selExpression
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
