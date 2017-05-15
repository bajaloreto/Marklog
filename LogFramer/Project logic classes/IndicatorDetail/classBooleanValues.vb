Public Class BooleanValue
    Private intIdBooleanValue As Integer = -1
    Private intIdTarget As Integer
    Private objParentGuid As Guid
    Private boolValue As Boolean
    Private objGuid As Guid

    Public Sub New()

    End Sub

    Public Sub New(ByVal value As Boolean)
        Me.Value = value
    End Sub

    Public Property idBooleanValue As Integer
        Get
            Return intIdBooleanValue
        End Get
        Set(ByVal value As Integer)
            intIdBooleanValue = value
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

    Public Property Value() As Boolean
        Get
            Return boolValue
        End Get
        Set(ByVal value As Boolean)
            boolValue = value
        End Set
    End Property
End Class

Public Class BooleanValues
    Inherits System.Collections.CollectionBase

    Public Event BooleanValueAdded(ByVal sender As Object, ByVal e As BooleanValueAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal booleanvalue As BooleanValue)
        If booleanvalue IsNot Nothing Then
            List.Add(booleanvalue)
            RaiseEvent BooleanValueAdded(Me, New BooleanValueAddedEventArgs(booleanvalue))
        End If
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal booleanvalue As BooleanValue)
        List.Insert(intIndex, booleanvalue)
        RaiseEvent BooleanValueAdded(Me, New BooleanValueAddedEventArgs(booleanvalue))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal booleanvalue As BooleanValue)
        If Me.List.Contains(booleanvalue) Then
            Me.List.Remove(booleanvalue)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As BooleanValue
        Get
            If index >= 0 And index <= Me.Count - 1 Then
                Return CType(List.Item(index), BooleanValue)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal booleanvalue As BooleanValue) As Integer
        Return List.IndexOf(booleanvalue)
    End Function

    Public Function Contains(ByVal booleanvalue As BooleanValue) As Boolean
        Return List.Contains(booleanvalue)
    End Function

    Public Sub SetValue(ByVal intIndex As Integer, ByVal boolValue As Boolean, ByVal boolMaxScoreIsSum As Boolean)
        If intIndex <= Me.Count - 1 Then
            Dim selValue As BooleanValue = Me.List(intIndex)
            If boolMaxScoreIsSum = False And boolValue = True Then
                For Each objValue As BooleanValue In Me.List
                    objValue.Value = False
                Next
            End If
            selValue.Value = boolValue
        End If

    End Sub

    Public Function GetBooleanValueByGuid(ByVal objGuid As Guid) As BooleanValue
        Dim selBooleanValue As BooleanValue = Nothing
        For Each objTarget As BooleanValue In Me.List
            If objTarget.Guid = objGuid Then
                selBooleanValue = objTarget
                Exit For
            End If
        Next
        Return selBooleanValue
    End Function
End Class

Public Class BooleanValuesMatrixRow
    Private objBooleanValues As New BooleanValues
    Private objGuid, objParentGuid As Guid

    Public Property BooleanValues As BooleanValues
        Get
            Return objBooleanValues
        End Get
        Set(ByVal value As BooleanValues)
            objBooleanValues = value
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
End Class

Public Class BooleanValuesMatrix
    Inherits System.Collections.CollectionBase

    Public Event BooleanValuesMatrixRowAdded(ByVal sender As Object, ByVal e As BooleanValuesMatrixRowAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal booleanvaluesmatrixrow As BooleanValuesMatrixRow)
        List.Add(booleanvaluesmatrixrow)
        RaiseEvent BooleanValuesMatrixRowAdded(Me, New BooleanValuesMatrixRowAddedEventArgs(booleanvaluesmatrixrow))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal booleanvaluesmatrixrow As BooleanValuesMatrixRow)
        List.Insert(intIndex, booleanvaluesmatrixrow)
        RaiseEvent BooleanValuesMatrixRowAdded(Me, New BooleanValuesMatrixRowAddedEventArgs(booleanvaluesmatrixrow))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal booleanvaluesmatrixrow As BooleanValuesMatrixRow)
        If Me.List.Contains(booleanvaluesmatrixrow) Then
            Me.List.Remove(booleanvaluesmatrixrow)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As BooleanValuesMatrixRow
        Get
            If index >= 0 And index <= Me.Count - 1 Then
                Return CType(List.Item(index), BooleanValuesMatrixRow)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal booleanvaluesmatrixrow As BooleanValuesMatrixRow) As Integer
        Return List.IndexOf(booleanvaluesmatrixrow)
    End Function

    Public Function Contains(ByVal booleanvaluesmatrixrow As BooleanValuesMatrixRow) As Boolean
        Return List.Contains(booleanvaluesmatrixrow)
    End Function

    Public Function GetBooleanValuesMatrixRowByGuid(ByVal objGuid As Guid) As BooleanValuesMatrixRow
        Dim selBooleanValuesMatrixRow As BooleanValuesMatrixRow = Nothing
        For Each objTarget As BooleanValuesMatrixRow In Me.List
            If objTarget.Guid = objGuid Then
                selBooleanValuesMatrixRow = objTarget
                Exit For
            End If
        Next
        Return selBooleanValuesMatrixRow
    End Function
End Class

Public Class BooleanValueAddedEventArgs
    Inherits EventArgs

    Public Property BooleanValue As BooleanValue

    Public Sub New(ByVal objBooleanValue As BooleanValue)
        MyBase.New()

        Me.BooleanValue = objBooleanValue
    End Sub
End Class

Public Class BooleanValuesMatrixRowAddedEventArgs
    Inherits EventArgs

    Public Property BooleanValuesMatrixRow As BooleanValuesMatrixRow

    Public Sub New(ByVal objBooleanValuesMatrixRow As BooleanValuesMatrixRow)
        MyBase.New()

        Me.BooleanValuesMatrixRow = objBooleanValuesMatrixRow
    End Sub
End Class

