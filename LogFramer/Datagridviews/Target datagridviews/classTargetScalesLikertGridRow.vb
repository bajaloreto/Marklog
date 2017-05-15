Public Class TargetScalesLikertGridRow
    Private objStatement As New Statement
    Private bmFirstLabelImage As Bitmap
    Private boolFirstLabelImageDirty As Boolean = True
    Private intFirstLabelHeight As Integer
    Private bmSecondLabelImage As Bitmap
    Private boolSecondLabelImageDirty As Boolean = True
    Private intSecondLabelHeight As Integer
    Private objBooleanValuesMatrixRow As new BooleanValuesMatrixRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal statement As Statement, ByVal booleanvaluesmatrixrow As BooleanValuesMatrixRow)
        Me.Statement = statement
        Me.BooleanValuesMatrixRow = booleanvaluesmatrixrow
    End Sub

    Public Property Statement As Statement
        Get
            Return objStatement
        End Get
        Set(ByVal value As Statement)
            objStatement = value
        End Set
    End Property

    Public Property FirstLabelImage As Bitmap
        Get
            Return bmFirstLabelImage
        End Get
        Set(ByVal value As Bitmap)
            bmFirstLabelImage = value
        End Set
    End Property

    Public Property FirstLabelImageDirty As Boolean
        Get
            Return boolFirstLabelImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolFirstLabelImageDirty = value
        End Set
    End Property

    Public Property FirstLabelHeight As Integer
        Get
            Return intFirstLabelHeight
        End Get
        Set(ByVal value As Integer)
            intFirstLabelHeight = value
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

    Public Property SecondLabelImage As Bitmap
        Get
            Return bmSecondLabelImage
        End Get
        Set(ByVal value As Bitmap)
            bmSecondLabelImage = value
        End Set
    End Property

    Public Property SecondLabelImageDirty As Boolean
        Get
            Return boolSecondLabelImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolSecondLabelImageDirty = value
        End Set
    End Property

    Public Property SecondLabelHeight As Integer
        Get
            Return intSecondLabelHeight
        End Get
        Set(ByVal value As Integer)
            intSecondLabelHeight = value
        End Set
    End Property

    Public Property SecondLabel As String
        Get
            If Statement IsNot Nothing Then
                Return Statement.SecondLabel
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As String)
            If Statement IsNot Nothing Then
                Statement.SecondLabel = value
            End If
        End Set
    End Property

    Public ReadOnly Property OneLabelContainsText As Boolean
        Get
            Dim strFirst, strSecond As String

            strFirst = RichTextToText(Me.FirstLabel)
            strSecond = RichTextToText(Me.SecondLabel)

            If String.IsNullOrEmpty(strFirst) = False Or String.IsNullOrEmpty(strSecond) = False Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public Property BooleanValuesMatrixRow As BooleanValuesMatrixRow
        Get
            Return objBooleanValuesMatrixRow
        End Get
        Set(ByVal value As BooleanValuesMatrixRow)
            objBooleanValuesMatrixRow = value
        End Set
    End Property
End Class

Public Class TargetScalesLikertGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As TargetScalesLikertGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As TargetScalesLikertGridRow)
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

    Public Sub Remove(ByVal gridrow As TargetScalesLikertGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As TargetScalesLikertGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), TargetScalesLikertGridRow)
            End If
        End Get
        Set(ByVal value As TargetScalesLikertGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As TargetScalesLikertGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As TargetScalesLikertGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
End Class
