Public MustInherit Class DataGridViewBaseClassRichText
    Inherits DataGridViewBaseClass

    Private intTextSelectionIndex As Integer
    Private objDragBoxFromMouseDown As Rectangle
    Private boolDragMultipleCells As Boolean
    Private bytDragOperator As Byte
    Private boolDragReleased As Boolean = True
    Private strDragColName As String

    Public Enum DragOperatorValues As Integer
        Move = 0
        Copy = 1
    End Enum

    Public Property TextSelectionIndex() As Integer
        Get
            Return intTextSelectionIndex
        End Get
        Set(ByVal value As Integer)
            intTextSelectionIndex = value
        End Set
    End Property

    Public Property DragBoxFromMouseDown As Rectangle
        Get
            Return objDragBoxFromMouseDown
        End Get
        Set(ByVal value As Rectangle)
            objDragBoxFromMouseDown = value
        End Set
    End Property

    Public Property DragMultipleCells As Boolean
        Get
            Return boolDragMultipleCells
        End Get
        Set(ByVal value As Boolean)
            boolDragMultipleCells = value
        End Set
    End Property

    Public Property DragOperator As Byte
        Get
            Return bytDragOperator
        End Get
        Set(ByVal value As Byte)
            bytDragOperator = value
        End Set
    End Property

    Public Property DragReleased As Boolean
        Get
            Return boolDragReleased
        End Get
        Set(ByVal value As Boolean)
            boolDragReleased = value
        End Set
    End Property

    Public Function GetTextRectangle(ByVal rCell As Rectangle) As Rectangle
        Dim rText As Rectangle = Nothing

        If rCell <> Nothing Then _
            rText = New Rectangle(rCell.X + CONST_Padding, rCell.Y + CONST_Padding, rCell.Width - CONST_HorizontalPadding, rCell.Height - CONST_VerticalPadding)

        Return rText
    End Function
End Class
