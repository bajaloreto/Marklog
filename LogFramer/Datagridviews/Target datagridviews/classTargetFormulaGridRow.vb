Public Class TargetFormulaGridRow
    Private objStatement As New Statement
    Private bmStatementImage As Bitmap
    Private boolStatementImageDirty As Boolean = True
    Private intStatementHeight As Integer
    Private objBaselineDoubleValue As DoubleValue
    Private objTargetDoubleValues As DoubleValues

    Public Sub New()

    End Sub

    Public Sub New(ByVal statement As Statement, ByVal baselinebooleanvalue As DoubleValue, ByVal targetbooleanvalues As DoubleValues)
        Me.Statement = statement
        Me.BaselineDoubleValue = baselinebooleanvalue
        Me.TargetDoubleValues = targetbooleanvalues
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
            Return bmStatementImage
        End Get
        Set(ByVal value As Bitmap)
            bmStatementImage = value
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

    Public Property BaselineDoubleValue As DoubleValue
        Get
            Return objBaselineDoubleValue
        End Get
        Set(ByVal value As DoubleValue)
            objBaselineDoubleValue = value
        End Set
    End Property

    Public Property TargetDoubleValues As DoubleValues
        Get
            Return objTargetDoubleValues
        End Get
        Set(ByVal value As DoubleValues)
            objTargetDoubleValues = value
        End Set
    End Property
End Class

Public Class TargetFormulaGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As TargetFormulaGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As TargetFormulaGridRow)
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

    Public Sub Remove(ByVal gridrow As TargetFormulaGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As TargetFormulaGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), TargetFormulaGridRow)
            End If
        End Get
        Set(ByVal value As TargetFormulaGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As TargetFormulaGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As TargetFormulaGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
End Class
