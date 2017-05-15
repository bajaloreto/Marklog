Public Class PrintPartnerRow
    Private intRowHeight As Integer
    Private strPropertyName As String
    Private strPropertyValue As String
    Private intObjectType As Integer
    Private boolBold As Boolean

    Public Sub New()

    End Sub

    Public Sub New(ByVal propertyvalue As String, ByVal propertyname As String, _
                   ByVal propertytype As Integer, Optional ByVal fontbold As Boolean = False)
        Me.PropertyName = propertyname
        Me.PropertyValue = propertyvalue
        Me.PropertyType = propertytype
        Me.Bold = fontbold
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

    Public Property PropertyValue() As String
        Get
            Return strPropertyValue
        End Get
        Set(ByVal value As String)
            strPropertyValue = value
        End Set
    End Property

    Public Property PropertyType() As Integer
        Get
            Return intObjectType
        End Get
        Set(ByVal value As Integer)
            intObjectType = value
        End Set
    End Property

    Public Property Bold() As Boolean
        Get
            Return boolBold
        End Get
        Set(ByVal value As Boolean)
            boolBold = value
        End Set
    End Property
End Class

Public Class PrintPartnerRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintPartnerRow)
        List.Add(row)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintPartnerRow)
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

    Public Sub Remove(ByVal row As PrintPartnerRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("PrintPartnerRow not in list!")
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

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintPartnerRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintPartnerRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintPartnerRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintPartnerRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
