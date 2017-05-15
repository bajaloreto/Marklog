Public Class PrintLogframeRow
    Inherits LogframeRow
    Private intRowHeight As Integer
    Private boolClippedRow As Boolean
    Private intRowCountStruct As Integer
    Private intRowCountIndicator As Integer

    Public Sub New()

    End Sub

    Public Sub New(ByVal section As Integer)
        Me.Section = section
    End Sub

    Public Sub New(ByVal section As Integer, ByVal structsort As String, ByVal struct As Struct, ByVal indicatorsort As String, ByVal indicator As Indicator, _
                   ByVal verificationsourcesort As String, ByVal verificationsource As VerificationSource, ByVal resourcesort As String, ByVal resource As Resource, _
                   ByVal assumptionsort As String, ByVal assumption As Assumption)
        Me.Section = section
        Me.StructSort = structsort
        Me.Struct = struct
        Me.IndicatorSort = indicatorsort
        Me.Indicator = indicator
        Me.VerificationSourceSort = verificationsourcesort
        Me.VerificationSource = verificationsource
        Me.ResourceSort = resourcesort
        Me.Resource = resource
        Me.AssumptionSort = assumptionsort
        Me.Assumption = assumption
    End Sub

#Region "Properties"
    Public Property ClippedRow As Boolean
        Get
            Return boolClippedRow
        End Get
        Set(value As Boolean)
            boolClippedRow = value
        End Set
    End Property

    Public Property RowHeight As Integer
        Get
            Return intRowHeight
        End Get
        Set(ByVal value As Integer)
            intRowHeight = value
        End Set
    End Property

    Public Property RowCountStruct As Integer
        Get
            Return intRowCountStruct
        End Get
        Set(value As Integer)
            intRowCountStruct = value
        End Set
    End Property

    Public Property RowCountIndicator As Integer
        Get
            Return intRowCountIndicator
        End Get
        Set(value As Integer)
            intRowCountIndicator = value
        End Set
    End Property
#End Region
End Class

Public Class PrintLogframeRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal PrintLogframeRow As PrintLogframeRow)
        List.Add(PrintLogframeRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal PrintLogframeRow As PrintLogframeRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(PrintLogframeRow)
        Else
            List.Insert(index, PrintLogframeRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal PrintLogframeRow As PrintLogframeRow)
        If List.Contains(PrintLogframeRow) = False Then
            System.Windows.Forms.MessageBox.Show("PrintLogframeRow not in list!")
        Else
            List.Remove(PrintLogframeRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintLogframeRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintLogframeRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal LogframeRow As PrintLogframeRow) As Integer
        Return List.IndexOf(LogframeRow)
    End Function

    Public Function Contains(ByVal LogframeRow As PrintLogframeRow) As Boolean
        Return List.Contains(LogframeRow)
    End Function
End Class