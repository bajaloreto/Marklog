Public Class DoubleValue
    Private intIdDoubleValue As Integer = -1
    Private intIdTarget As Integer
    Private objParentGuid As Guid
    Private dblValue As Double
    Private objGuid As Guid

    Public Sub New()

    End Sub

    Public Sub New(ByVal value As Double)
        Me.Value = value
    End Sub

    Public Property idDoubleValue As Integer
        Get
            Return intIdDoubleValue
        End Get
        Set(ByVal value As Integer)
            intIdDoubleValue = value
        End Set
    End Property

    Public Property idParent As Integer
        Get
            Return intIdTarget
        End Get
        Set(ByVal value As Integer)
            intIdTarget = value
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

    Public Property ParentGuid As Guid
        Get
            Return objParentGuid
        End Get
        Set(ByVal value As Guid)
            objParentGuid = value
        End Set
    End Property

    Public Property Value() As Double
        Get
            Return dblValue
        End Get
        Set(ByVal value As Double)
            dblValue = value
        End Set
    End Property
End Class

Public Class DoubleValues
    Inherits System.Collections.CollectionBase

    Public Event DoubleValueAdded(ByVal sender As Object, ByVal e As DoubleValueAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal doublevalue As DoubleValue)
        List.Add(doublevalue)
        RaiseEvent DoubleValueAdded(Me, New DoubleValueAddedEventArgs(doublevalue))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal doublevalue As DoubleValue)
        List.Insert(intIndex, doublevalue)
        RaiseEvent DoubleValueAdded(Me, New DoubleValueAddedEventArgs(doublevalue))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal doublevalue As DoubleValue)
        If Me.List.Contains(doublevalue) Then
            Me.List.Remove(doublevalue)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As DoubleValue
        Get
            If index >= 0 And index <= Me.Count - 1 Then
                Return CType(List.Item(index), DoubleValue)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal doublevalue As DoubleValue) As Integer
        Return List.IndexOf(doublevalue)
    End Function

    Public Function Contains(ByVal doublevalue As DoubleValue) As Boolean
        Return List.Contains(doublevalue)
    End Function

    Public Function GetDoubleValueByGuid(ByVal objGuid As Guid) As DoubleValue
        Dim selDoubleValue As DoubleValue = Nothing
        For Each objTarget As DoubleValue In Me.List
            If objTarget.Guid = objGuid Then
                selDoubleValue = objTarget
                Exit For
            End If
        Next
        Return selDoubleValue
    End Function
End Class

Public Class DoubleValuesMatrixRow
    Private objDoubleValues As New DoubleValues
    Private objGuid, objParentGuid As Guid

    Public Property DoubleValues As DoubleValues
        Get
            Return objDoubleValues
        End Get
        Set(ByVal value As DoubleValues)
            objDoubleValues = value
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

    Public Property ParentGuid As Guid
        Get
            Return objParentGuid
        End Get
        Set(ByVal value As Guid)
            objParentGuid = value
        End Set
    End Property

    Public ReadOnly Property Total As Double
        Get
            Dim dblTotal As Double

            For i = 0 To Me.DoubleValues.Count - 1
                dblTotal += DoubleValues(i).Value
            Next

            Return dblTotal
        End Get
    End Property

    Public ReadOnly Property Median As Integer
        Get
            Dim intMedian As Integer

            intMedian = Math.Ceiling(Me.Total / 2)

            Return intMedian
        End Get
    End Property

    Public ReadOnly Property FirstQuartile As Integer
        Get
            Dim intFirstQuartile As Integer

            intFirstQuartile = Math.Ceiling(Me.Total / 4)

            Return intFirstQuartile
        End Get
    End Property

    Public ReadOnly Property ThirdQuartile As Integer
        Get
            Dim intThirdQuartile As Integer

            intThirdQuartile = Math.Ceiling(Me.Total * 0.75)

            Return intThirdQuartile
        End Get
    End Property

    Public ReadOnly Property InterQuartileRange As Integer
        Get
            Dim intIndexOfFirstQuartile As Integer = GetIndexOfFirstQuartileClass()
            Dim intIndexOfThirdQuartile As Integer = GetIndexOfThirdQuartileClass()
            Dim intInterQuartileRange As Integer = intIndexOfThirdQuartile - intIndexOfFirstQuartile

            Return intInterQuartileRange
        End Get
    End Property

    Public Function GetIndexOfMedianClass() As Integer
        Dim intMedian As Integer = Me.Median
        Dim intClass As Integer
        Dim intTotal As Integer
        Dim boolFound As Boolean


        For intClass = 0 To DoubleValues.Count - 1
            intTotal += DoubleValues(intClass).Value
            If intTotal >= intMedian Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = False Then intClass = 0

        Return intClass
    End Function

    Public Function GetIndexOfFirstQuartileClass() As Integer
        Dim intFirstQuartile As Integer = Me.FirstQuartile
        Dim intClass As Integer
        Dim intTotal As Integer
        Dim boolFound As Boolean

        For intClass = 0 To DoubleValues.Count - 1
            intTotal += DoubleValues(intClass).Value
            If intTotal >= intFirstQuartile Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = False Then intClass = 0

        Return intClass
    End Function

    Public Function GetIndexOfThirdQuartileClass() As Integer
        Dim intThirdQuartile As Integer = Me.ThirdQuartile
        Dim intClass As Integer
        Dim intTotal As Integer
        Dim boolFound As Boolean

        For intClass = 0 To DoubleValues.Count - 1
            intTotal += DoubleValues(intClass).Value
            If intTotal >= intThirdQuartile Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = False Then intClass = 0

        Return intClass
    End Function
End Class

Public Class DoubleValuesMatrix
    Inherits System.Collections.CollectionBase

    Public Event DoubleValuesMatrixRowAdded(ByVal sender As Object, ByVal e As DoubleValuesMatrixRowAddedEventArgs)

    Public Sub New()

    End Sub

#Region "Calculations"
    Public Function GetRowCount() As Integer
        Return List.Count
    End Function

    Public Function GetColumnCount() As Integer
        Dim intColumnCount As Integer

        If List.Count > 0 Then
            Dim FirstRow As DoubleValuesMatrixRow = List(0)

            intColumnCount = FirstRow.DoubleValues.Count
        End If

        Return intColumnCount
    End Function

    Public Function CalculateColumnTotal(ByVal intIndex As Integer) As Double
        Dim dblTotal As Double

        For Each selRow As DoubleValuesMatrixRow In Me.List
            If intIndex <= selRow.DoubleValues.Count - 1 Then _
                dblTotal += selRow.DoubleValues(intIndex).Value
        Next

        Return dblTotal
    End Function

    Public Function CalculateTotal() As Double
        Dim intColumnCount As Integer = GetColumnCount()
        Dim dblTotal As Double

        For i = 0 To intColumnCount - 1
            dblTotal += CalculateColumnTotal(i)
        Next

        Return dblTotal
    End Function

    Public Function CalculateMedian() As Integer
        Dim dblTotal As Integer = CalculateTotal()
        Dim intMedian As Integer

        intMedian = Math.Ceiling(dblTotal / 2)

        Return intMedian
    End Function

    Public Function CalculateFirstQuartile() As Integer
        Dim dblTotal As Integer = CalculateTotal()
        Dim intFirstQuartile As Integer

        intFirstQuartile = Math.Ceiling(dblTotal / 4)

        Return intFirstQuartile

    End Function

    Public Function CalculateThirdQuartile() As Integer
        Dim dblTotal As Integer = CalculateTotal()
        Dim intThirdQuartile As Integer

        intThirdQuartile = Math.Ceiling(dblTotal * 0.75)

        Return intThirdQuartile

    End Function

    Public Function CalculateInterQuartileRange() As Integer

        Dim intIndexOfFirstQuartile As Integer = GetIndexOfFirstQuartileClass()
        Dim intIndexOfThirdQuartile As Integer = GetIndexOfThirdQuartileClass()
        Dim intInterQuartileRange As Integer = intIndexOfThirdQuartile - intIndexOfFirstQuartile

        Return intInterQuartileRange

    End Function

    Public Function GetIndexOfMedianClass() As Integer
        Dim intMedian As Integer = CalculateMedian()
        Dim intColumnCount As Integer = GetColumnCount()
        Dim intClass As Integer
        Dim intTotal As Integer
        Dim boolFound As Boolean


        For intClass = 0 To intColumnCount - 1
            intTotal += CalculateColumnTotal(intClass)
            If intTotal >= intMedian Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = False Then intClass = 0

        Return intClass
    End Function

    Public Function GetIndexOfFirstQuartileClass() As Integer
        Dim intFirstQuartile As Integer = CalculateFirstQuartile()
        Dim intColumnCount As Integer = GetColumnCount()
        Dim intClass As Integer
        Dim intTotal As Integer
        Dim boolFound As Boolean


        For intClass = 0 To intColumnCount - 1
            intTotal += CalculateColumnTotal(intClass)
            If intTotal >= intFirstQuartile Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = False Then intClass = 0

        Return intClass
    End Function

    Public Function GetIndexOfThirdQuartileClass() As Integer
        Dim intThirdQuartile As Integer = CalculateThirdQuartile()
        Dim intColumnCount As Integer = GetColumnCount()
        Dim intClass As Integer
        Dim intTotal As Integer
        Dim boolFound As Boolean

        For intClass = 0 To intColumnCount - 1
            intTotal += CalculateColumnTotal(intClass)
            If intTotal >= intThirdQuartile Then
                boolFound = True
                Exit For
            End If
        Next

        If boolFound = False Then intClass = 0

        Return intClass
    End Function
#End Region

#Region "Manage rows"
    Public Sub Add(ByVal doublevaluesmatrixrow As DoubleValuesMatrixRow)
        List.Add(doublevaluesmatrixrow)
        RaiseEvent DoubleValuesMatrixRowAdded(Me, New DoubleValuesMatrixRowAddedEventArgs(doublevaluesmatrixrow))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal doublevaluesmatrixrow As DoubleValuesMatrixRow)
        List.Insert(intIndex, doublevaluesmatrixrow)
        RaiseEvent DoubleValuesMatrixRowAdded(Me, New DoubleValuesMatrixRowAddedEventArgs(doublevaluesmatrixrow))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal doublevaluesmatrixrow As DoubleValuesMatrixRow)
        If Me.List.Contains(doublevaluesmatrixrow) Then
            Me.List.Remove(doublevaluesmatrixrow)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As DoubleValuesMatrixRow
        Get
            If index >= 0 And index <= Me.Count - 1 Then
                Return CType(List.Item(index), DoubleValuesMatrixRow)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal doublevaluesmatrixrow As DoubleValuesMatrixRow) As Integer
        Return List.IndexOf(doublevaluesmatrixrow)
    End Function

    Public Function Contains(ByVal doublevaluesmatrixrow As DoubleValuesMatrixRow) As Boolean
        Return List.Contains(doublevaluesmatrixrow)
    End Function

    Public Function GetDoubleValuesMatrixRowByGuid(ByVal objGuid As Guid) As DoubleValuesMatrixRow
        Dim selDoubleValuesMatrixRow As DoubleValuesMatrixRow = Nothing
        For Each objMatrixRow As DoubleValuesMatrixRow In Me.List
            If objMatrixRow.Guid = objGuid Then
                selDoubleValuesMatrixRow = objMatrixRow
                Exit For
            End If
        Next
        Return selDoubleValuesMatrixRow
    End Function

    Public Function GetParentMatrixRowOfValue(ByVal objDoubleValue As DoubleValue) As DoubleValuesMatrixRow
        Dim selDoubleValuesMatrixRow As DoubleValuesMatrixRow = Nothing
        For Each objMatrixRow As DoubleValuesMatrixRow In Me.List
            If objMatrixRow.DoubleValues.Contains(objDoubleValue) Then
                selDoubleValuesMatrixRow = objMatrixRow
                Exit For
            End If
        Next
        Return selDoubleValuesMatrixRow
    End Function
#End Region
End Class

Public Class DoubleValueAddedEventArgs
    Inherits EventArgs

    Public Property DoubleValue As DoubleValue

    Public Sub New(ByVal objDoubleValue As DoubleValue)
        MyBase.New()

        Me.DoubleValue = objDoubleValue
    End Sub
End Class

Public Class DoubleValuesMatrixRowAddedEventArgs
    Inherits EventArgs

    Public Property DoubleValuesMatrixRow As DoubleValuesMatrixRow

    Public Sub New(ByVal objDoubleValuesMatrixRow As DoubleValuesMatrixRow)
        MyBase.New()

        Me.DoubleValuesMatrixRow = objDoubleValuesMatrixRow
    End Sub
End Class

