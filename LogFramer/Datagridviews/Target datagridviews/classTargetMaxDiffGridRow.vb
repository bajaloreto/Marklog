Public Class TargetMaxDiffGridRow
    Private objStatement As Statement
    Private bmResponseImage As Bitmap
    Private boolResponseImageDirty As Boolean = True
    Private intResponseHeight As Integer
    Private objLeastImportantValue As BooleanValue
    Private objMostImportantValue As BooleanValue

    Public Sub New()

    End Sub

    Public Sub New(ByVal statement As Statement, ByVal leastimportantvalue As BooleanValue, ByVal mostimportantvalue As BooleanValue)
        Me.Response = statement
        Me.LeastImportantValue = leastimportantvalue
        Me.MostImportantValue = mostimportantvalue
    End Sub

    Public Property Response As Statement
        Get
            Return objStatement
        End Get
        Set(ByVal value As Statement)
            objStatement = value
        End Set
    End Property

    Public Property ResponseImage As Bitmap
        Get
            Return bmResponseImage
        End Get
        Set(ByVal value As Bitmap)
            bmResponseImage = value
        End Set
    End Property

    Public Property ResponseImageDirty As Boolean
        Get
            Return boolResponseImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolResponseImageDirty = value
        End Set
    End Property

    Public Property ResponseHeight As Integer
        Get
            Return intResponseHeight
        End Get
        Set(ByVal value As Integer)
            intResponseHeight = value
        End Set
    End Property

    Public Property FirstLabel As String
        Get
            If Response IsNot Nothing Then
                Return Response.FirstLabel
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As String)
            If Response IsNot Nothing Then
                Response.FirstLabel = value
            End If
        End Set
    End Property

    Public Property LeastImportantValue As BooleanValue
        Get
            Return objLeastImportantValue
        End Get
        Set(ByVal value As BooleanValue)
            objLeastImportantValue = value
        End Set
    End Property

    Public Property MostImportantValue As BooleanValue
        Get
            Return objMostImportantValue
        End Get
        Set(ByVal value As BooleanValue)
            objMostImportantValue = value
        End Set
    End Property
End Class

Public Class TargetMaxDiffGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As TargetMaxDiffGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As TargetMaxDiffGridRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(gridrow)
        Else
            List.Insert(index, gridrow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal gridrow As TargetMaxDiffGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As TargetMaxDiffGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), TargetMaxDiffGridRow)
            End If
        End Get
        Set(ByVal value As TargetMaxDiffGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As TargetMaxDiffGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As TargetMaxDiffGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
End Class
