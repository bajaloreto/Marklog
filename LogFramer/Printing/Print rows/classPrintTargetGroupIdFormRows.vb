Public Class PrintTargetGroupIdFormRow
    Private intRowHeight As Integer
    Private strPropertyName, strPropertyValue As String
    Private intPropertyType As Integer = -1
    Private boolBold As Boolean

    Public Sub New()

    End Sub

    Public Sub New(ByVal propertyname As String, ByVal fontbold As Boolean)
        Me.PropertyName = propertyname
        Me.IsTitle = fontbold
    End Sub

    Public Sub New(ByVal propertyname As String, ByVal propertytype As Integer, Optional ByVal propertyvalue As String = "")
        Me.PropertyName = propertyname
        Me.PropertyType = propertytype
        Me.PropertyValue = propertyvalue
    End Sub


    Public Property RowHeight As Integer
        Get
            Return intRowHeight
        End Get
        Set(ByVal value As Integer)
            intRowHeight = value
        End Set
    End Property

    Public Property PropertyName() As String
        Get
            Return strPropertyName
        End Get
        Set(ByVal value As String)
            strPropertyName = value
        End Set
    End Property

    Public Property PropertyType() As Integer
        Get
            Return intPropertyType
        End Get
        Set(ByVal value As Integer)
            intPropertyType = value
        End Set
    End Property

    Public Property PropertyValue() As String
        Get
            Return strPropertyValue
        End Get
        Set(ByVal value As String)
            strPropertyValue = value
        End Set
    End Property

    Public Property IsTitle() As Boolean
        Get
            Return boolBold
        End Get
        Set(ByVal value As Boolean)
            boolBold = value
        End Set
    End Property
End Class

Public Class PrintTargetGroupIdFormRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintTargetGroupIdFormRow)
        List.Add(row)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintTargetGroupIdFormRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(row)
        Else
            List.Insert(index, row)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal row As PrintTargetGroupIdFormRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("PrintTargetGroupIdFormRow not in list!")
        Else
            List.Remove(row)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintTargetGroupIdFormRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintTargetGroupIdFormRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintTargetGroupIdFormRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintTargetGroupIdFormRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
