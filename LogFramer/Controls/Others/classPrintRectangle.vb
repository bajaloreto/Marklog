Public Class PrintRectangle
    Dim objRectangle As New Rectangle
    Private intHorizontalPageIndex As Integer = -1
    Private boolTextRectangle As Boolean

    Public Sub New()

    End Sub

    Public Sub New(ByVal location As Point, ByVal size As Size)
        With objRectangle
            .Location = location
            .Size = size
        End With
    End Sub

    Public Sub New(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)
        With objRectangle
            .X = x
            .Y = y
            .Width = width
            .Height = height
        End With
    End Sub

    Public Sub New(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal horizontalpageindex As Integer, Optional ByVal istextrectangle As Boolean = False)
        With objRectangle
            .X = x
            .Y = y
            .Width = width
            .Height = height
        End With
        Me.HorizontalPageIndex = horizontalpageindex
        Me.IsTextRectangle = istextrectangle
    End Sub

    Public ReadOnly Property Bottom As Integer
        Get
            Return objRectangle.Bottom
        End Get
    End Property

    Public Property Height As Integer
        Get
            Return objRectangle.Height
        End Get
        Set(ByVal value As Integer)
            objRectangle.Height = value
        End Set
    End Property

    Public Property HorizontalPageIndex As Integer
        Get
            Return intHorizontalPageIndex
        End Get
        Set(ByVal value As Integer)
            intHorizontalPageIndex = value
        End Set
    End Property

    Public ReadOnly Property IsEmpty As Boolean
        Get
            Return objRectangle.IsEmpty
        End Get
    End Property

    Public Property IsTextRectangle As Boolean
        Get
            Return boolTextRectangle
        End Get
        Set(ByVal value As Boolean)
            boolTextRectangle = value
        End Set
    End Property

    Public ReadOnly Property Left As Integer
        Get
            Return objRectangle.Left
        End Get
    End Property

    Public Property Location As Point
        Get
            Return objRectangle.Location
        End Get
        Set(ByVal value As Point)
            objRectangle.Location = value
        End Set
    End Property

    Public ReadOnly Property Right As Integer
        Get
            Return objRectangle.Right
        End Get
    End Property

    Public Property Size As Size
        Get
            Return objRectangle.Size
        End Get
        Set(ByVal value As Size)
            objRectangle.Size = value
        End Set
    End Property

    Public ReadOnly Property Top As Integer
        Get
            Return objRectangle.Top
        End Get
    End Property

    Public Property Width As Integer
        Get
            Return objRectangle.Width
        End Get
        Set(ByVal value As Integer)
            objRectangle.Width = value
        End Set
    End Property

    Public Property X As Integer
        Get
            Return objRectangle.X
        End Get
        Set(ByVal value As Integer)
            objRectangle.X = value
        End Set
    End Property

    Public Property Y As Integer
        Get
            Return objRectangle.Y
        End Get
        Set(ByVal value As Integer)
            objRectangle.Y = value
        End Set
    End Property
End Class

Public Class PrintRectangles
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal printrectangle As PrintRectangle)
        If printrectangle IsNot Nothing Then _
            List.Add(printrectangle)
    End Sub

    Public Sub AddRange(ByVal printrectangles As PrintRectangles)
        InnerList.AddRange(printrectangles)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal printrectangle As PrintRectangle)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(printrectangle)
        Else
            List.Insert(index, printrectangle)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal printrectangle As PrintRectangle)
        If List.Contains(printrectangle) = False Then
            System.Windows.Forms.MessageBox.Show("Print rectangle not in list!")
        Else
            List.Remove(printrectangle)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Public Function GetRectanglesOfPage(ByVal intHorizontalPageIndex As Integer)
        Dim SelectedRectangles As New PrintRectangles

        For Each selRectangle As PrintRectangle In Me.List
            If selRectangle.HorizontalPageIndex = intHorizontalPageIndex Then
                SelectedRectangles.Add(selRectangle)
            End If
        Next

        Return SelectedRectangles
    End Function

    Public Function GetTotalWidth()
        Dim intWidth As Integer
        For Each selRectangle As PrintRectangle In Me.List
            intWidth += selRectangle.Width
        Next

        Return intWidth
    End Function

    Public ReadOnly Property CountTextRectangles
        Get
            Dim intCount As Integer
            For Each selRectangle As PrintRectangle In Me.List
                If selRectangle.IsTextRectangle = True Then
                    intCount += 1
                End If
            Next

            Return intCount
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintRectangle
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintRectangle)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal printrectangle As PrintRectangle) As Integer
        Return List.IndexOf(printrectangle)
    End Function

    Public Function Contains(ByVal printrectangle As PrintRectangle) As Boolean
        Return List.Contains(printrectangle)
    End Function
End Class
