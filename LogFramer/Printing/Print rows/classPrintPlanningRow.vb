Public Class PrintPlanningRow
    Inherits PlanningGridRow

    Private intSection As Integer
    Private intRowHeight As Integer
    Private bmKeyMomentCellImage As System.Drawing.Bitmap

    Public Sub New()

    End Sub

    Public Sub New(ByVal section As Integer)
        Me.Section = section
    End Sub

    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
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

    Public Property KeyMomentCellImage() As Bitmap
        Get
            Return bmKeyMomentCellImage
        End Get
        Set(ByVal value As Bitmap)
            bmKeyMomentCellImage = value
        End Set
    End Property
End Class

Public Class PrintPlanningRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal printplanningrow As PrintPlanningRow)
        List.Add(printplanningrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal printplanningrow As PrintPlanningRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(printplanningrow)
        Else
            List.Insert(index, printplanningrow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal printplanningrow As PrintPlanningRow)
        If List.Contains(printplanningrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(printplanningrow)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintPlanningRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintPlanningRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal printplanningrow As PrintPlanningRow) As Integer
        Return List.IndexOf(printplanningrow)
    End Function

    Public Function Contains(ByVal printplanningrow As PrintPlanningRow) As Boolean
        Return List.Contains(printplanningrow)
    End Function

    Public Function GetByGuid(ByVal FindGuid As Guid) As PrintPlanningRow
        Dim selGridRow As PrintPlanningRow
        For Each selGridRow In Me.List
            'If selGridRow.Guid = FindGuid Then
            '    Return selGridRow
            'End If
        Next
        Return Nothing
    End Function

    Public Function GetReferenceIndex(ByVal selRow As PrintPlanningRow) As Integer
        Dim intIndex As Integer = -1

        For Each objRow As PrintPlanningRow In Me.List
            If objRow.GuidReferenceMoment = selRow.GuidReferenceMoment Then intIndex += 1
            If objRow Is selRow Then Exit For
        Next

        Return intIndex
    End Function
End Class
