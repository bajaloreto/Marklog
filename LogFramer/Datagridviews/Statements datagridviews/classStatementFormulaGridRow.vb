Public Class StatementFormulaGridRow
    Private objStatement As Statement
    Private bmResponseImage As Bitmap
    Private boolStatementImageDirty As Boolean = True
    Private intStatementHeight As Integer

    Public Sub New()

    End Sub

    Public Sub New(ByVal statement As Statement)
        Me.Statement = Statement
    End Sub

    Public Property Statement As Statement
        Get
            Return objStatement
        End Get
        Set(ByVal value As Statement)
            objStatement = value
        End Set
    End Property

    Public Property StatementImage As Bitmap
        Get
            Return bmResponseImage
        End Get
        Set(ByVal value As Bitmap)
            bmResponseImage = value
        End Set
    End Property

    Public Property StatementImageDirty As Boolean
        Get
            Return boolStatementImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolStatementImageDirty = value
        End Set
    End Property

    Public Property StatementHeight As Integer
        Get
            Return intStatementHeight
        End Get
        Set(ByVal value As Integer)
            intStatementHeight = value
        End Set
    End Property

    Public Property FirstLabel As String
        Get
            If Statement IsNot Nothing Then
                Return Statement.FirstLabel
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As String)
            If Statement IsNot Nothing Then
                Statement.FirstLabel = value
            End If
        End Set
    End Property

    Public Property NrDecimals As Double
        Get
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing Then
                Return Statement.ValuesDetail.NrDecimals
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Double)
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing Then
                Statement.ValuesDetail.NrDecimals = value
            End If
        End Set
    End Property

    Public Property Unit As String
        Get
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing Then
                Return Statement.ValuesDetail.Unit
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As String)
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing Then
                Statement.ValuesDetail.Unit = value
            End If
        End Set
    End Property

    Public Property OpMin As String
        Get
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Return Statement.ValuesDetail.ValueRange.OpMin
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As String)
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Statement.ValuesDetail.ValueRange.OpMin = value
            End If
        End Set
    End Property

    Public Property MinValue As Double
        Get
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Return Statement.ValuesDetail.ValueRange.MinValue
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Double)
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Statement.ValuesDetail.ValueRange.MinValue = value
            End If
        End Set
    End Property

    Public Property OpMax As String
        Get
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Return Statement.ValuesDetail.ValueRange.OpMax
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As String)
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Statement.ValuesDetail.ValueRange.OpMax = value
            End If
        End Set
    End Property

    Public Property MaxValue As Double
        Get
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Return Statement.ValuesDetail.ValueRange.MaxValue
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Double)
            If Statement IsNot Nothing AndAlso Statement.ValuesDetail IsNot Nothing AndAlso Statement.ValuesDetail.ValueRange IsNot Nothing Then
                Statement.ValuesDetail.ValueRange.MaxValue = value
            End If
        End Set
    End Property
End Class

Public Class StatementFormulaGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As StatementFormulaGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As StatementFormulaGridRow)
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

    Public Sub Remove(ByVal gridrow As StatementFormulaGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As StatementFormulaGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), StatementFormulaGridRow)
            End If
        End Get
        Set(ByVal value As StatementFormulaGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As StatementFormulaGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As StatementFormulaGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
End Class
