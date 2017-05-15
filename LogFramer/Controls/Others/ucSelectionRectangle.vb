Public Class SelectionRectangle
    Private intFirstRowIndex, intLastRowIndex As Integer
    Private intFirstColumnIndex, intLastColumnIndex As Integer
    Private rRectangle As New Rectangle

    Public Property FirstRowIndex As Integer
        Get
            Return intFirstRowIndex
        End Get
        Set(value As Integer)
            intFirstRowIndex = value
        End Set
    End Property

    Public Property LastRowIndex As Integer
        Get
            Return intLastRowIndex
        End Get
        Set(value As Integer)
            intLastRowIndex = value
        End Set
    End Property

    Public Property FirstColumnIndex As Integer
        Get
            Return intFirstColumnIndex
        End Get
        Set(value As Integer)
            intFirstColumnIndex = value
        End Set
    End Property

    Public Property LastColumnIndex As Integer
        Get
            Return intLastColumnIndex
        End Get
        Set(value As Integer)
            intLastColumnIndex = value
        End Set
    End Property

    Public Property Rectangle As Rectangle
        Get
            Return rRectangle
        End Get
        Set(value As Rectangle)
            rRectangle = value
        End Set
    End Property

    Public Property X As Integer
        Get
            Return rRectangle.X
        End Get
        Set(value As Integer)
            rRectangle.X = value
        End Set
    End Property

    Public Property Y As Integer
        Get
            Return rRectangle.Y
        End Get
        Set(value As Integer)
            rRectangle.Y = value
        End Set
    End Property

    Public Property Width As Integer
        Get
            Return rRectangle.Width
        End Get
        Set(value As Integer)
            rRectangle.Width = value
        End Set
    End Property

    Public Property Height As Integer
        Get
            Return rRectangle.Height
        End Get
        Set(value As Integer)
            rRectangle.Height = value
        End Set
    End Property
End Class
