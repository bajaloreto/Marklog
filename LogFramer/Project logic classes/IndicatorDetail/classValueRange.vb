Imports System.Xml.Serialization
Imports System.Web.Script.Serialization
Imports System.Globalization

Public Class ValueRange
    Private dblMinValue As Double, dblMaxValue As Double
    Private strOpMin As String = String.Empty, strOpMax As String = String.Empty
    Private objGuid As Guid

    Public Event ErrorParsingDecimal()

    Public Sub New()

    End Sub

#Region "Properties"
    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Sub New(ByVal opmin As String, ByVal minvalue As Double, ByVal opmax As String, ByVal maxvalue As Double)
        Me.OpMin = opmin
        Me.MinValue = minvalue
        Me.OpMax = opmax
        Me.MaxValue = maxvalue
    End Sub

    Public Property OpMin() As String
        Get
            Return strOpMin
        End Get
        Set(ByVal value As String)
            strOpMin = value
        End Set
    End Property

    Public Property MinValue() As Double
        Get
            Return dblMinValue
        End Get
        Set(ByVal value As Double)
            dblMinValue = value
        End Set
    End Property

    Public Property OpMax() As String
        Get
            Return strOpMax
        End Get
        Set(ByVal value As String)
            strOpMax = value
        End Set
    End Property

    Public Property MaxValue() As Double
        Get
            Return dblMaxValue
        End Get
        Set(ByVal value As Double)
            dblMaxValue = value
        End Set
    End Property

    Public ReadOnly Property ValueRangeSet As Boolean
        Get
            If String.IsNullOrEmpty(OpMin) = False Or String.IsNullOrEmpty(OpMax) = False Then
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
            Return LANG_ValueRange
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_ValueRanges
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function CheckWithinRange(ByVal dblNumber As Double) As Boolean
        Dim boolWithinRange As Boolean

        If String.IsNullOrEmpty(OpMin) = False Then
            Select Case OpMin
                Case CONST_Equals
                    If dblNumber = MinValue Then Return True
                Case CONST_LargerThan
                    If dblNumber > MinValue Then boolWithinRange = True
                Case CONST_LargerThanOrEqual
                    If dblNumber >= MinValue Then boolWithinRange = True
            End Select
        Else
            boolWithinRange = True
        End If

        If String.IsNullOrEmpty(OpMax) = False Then
            Select Case OpMax
                Case CONST_Equals
                    If boolWithinRange = True And dblNumber = MaxValue Then boolWithinRange = True Else boolWithinRange = False
                Case CONST_SmallerThan
                    If boolWithinRange = True And dblNumber < MaxValue Then boolWithinRange = True Else boolWithinRange = False
                Case CONST_SmallerThanOrEqual
                    If boolWithinRange = True And dblNumber <= MaxValue Then boolWithinRange = True Else boolWithinRange = False
            End Select
        End If

        Return boolWithinRange
    End Function

    Public Function MakeWithinRange(ByVal dblNumber As Double) As Double

        If String.IsNullOrEmpty(OpMin) = False Then
            Select Case OpMin
                Case CONST_Equals, CONST_LargerThanOrEqual
                    If dblNumber < MinValue Then dblNumber = MinValue
                Case CONST_LargerThan
                    If dblNumber < MinValue Then dblNumber = MinValue + 1
            End Select
        End If

        If String.IsNullOrEmpty(OpMax) = False Then
            Select Case OpMax
                Case CONST_Equals, CONST_SmallerThanOrEqual
                    If dblNumber > MaxValue Then dblNumber = MaxValue
                Case CONST_SmallerThan
                    If dblNumber > MaxValue Then dblNumber = MaxValue - 1
            End Select
        End If

        Return dblNumber
    End Function
#End Region
End Class

Public Class ValueRanges
    Inherits System.Collections.CollectionBase

    Public Event ValueRangeAdded(ByVal sender As Object, ByVal e As ValueRangeAddedEventArgs)

    Public Sub Add(ByVal valuerange As ValueRange)
        List.Add(valuerange)
        RaiseEvent ValueRangeAdded(Me, New ValueRangeAddedEventArgs(valuerange))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal valuerange As ValueRange)
        List.Insert(intIndex, valuerange)
        RaiseEvent ValueRangeAdded(Me, New ValueRangeAddedEventArgs(valuerange))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal valuerange As ValueRange)
        If Me.List.Contains(valuerange) Then
            Me.List.Remove(valuerange)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ValueRange
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), ValueRange)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal valuerange As ValueRange) As Integer
        Return List.IndexOf(valuerange)
    End Function

    Public Function Contains(ByVal valuerange As ValueRange) As Boolean
        Return List.Contains(valuerange)
    End Function

    Public Function GetValueRangeByGuid(ByVal objGuid As Guid) As ValueRange
        Dim selValueRange As ValueRange = Nothing

        For Each objValueRange As ValueRange In Me.List
            If objValueRange.Guid = objGuid Then
                selValueRange = objValueRange
                Exit For
            End If
        Next
        Return selValueRange
    End Function
End Class

Public Class ValueRangeAddedEventArgs
    Inherits EventArgs

    Public Property ValueRange As ValueRange

    Public Sub New(ByVal objValueRange As ValueRange)
        MyBase.New()

        Me.ValueRange = objValueRange
    End Sub
End Class


