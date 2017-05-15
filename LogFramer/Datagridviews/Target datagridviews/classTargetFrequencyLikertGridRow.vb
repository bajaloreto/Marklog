Public Class TargetFrequencyLikertGridRow
    Private objStatement As New Statement
    Private bmFirstLabelImage As Bitmap
    Private boolFirstLabelImageDirty As Boolean = True
    Private intFirstLabelHeight As Integer
    Private objDoubleValuesMatrixRow As New DoubleValuesMatrixRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal statement As Statement, ByVal doublevaluesmatrixrow As DoubleValuesMatrixRow)
        Me.Statement = statement
        Me.DoubleValuesMatrixRow = doublevaluesmatrixrow
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

    Public ReadOnly Property LabelContainsText As Boolean
        Get
            Dim strFirst As String

            strFirst = RichTextToText(Me.FirstLabel)

            If String.IsNullOrEmpty(strFirst) = False Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public Property DoubleValuesMatrixRow As DoubleValuesMatrixRow
        Get
            Return objDoubleValuesMatrixRow
        End Get
        Set(ByVal value As DoubleValuesMatrixRow)
            objDoubleValuesMatrixRow = value
        End Set
    End Property

    Public ReadOnly Property Total As Double
        Get
            Dim dblTotal As Double

            If DoubleValuesMatrixRow IsNot Nothing Then
                dblTotal = DoubleValuesMatrixRow.Total
            End If

            Return dblTotal
        End Get
    End Property

    Public ReadOnly Property IndexOfMedianClass As Integer
        Get
            Dim intClass As Integer

            If DoubleValuesMatrixRow IsNot Nothing Then
                intClass = DoubleValuesMatrixRow.GetIndexOfMedianClass
            End If

            Return intClass
        End Get
    End Property

    Public ReadOnly Property IndexOfFirstQuartileClass As Integer
        Get
            Dim intClass As Integer

            If DoubleValuesMatrixRow IsNot Nothing Then
                intClass = DoubleValuesMatrixRow.GetIndexOfFirstQuartileClass
            End If

            Return intClass
        End Get
    End Property

    Public ReadOnly Property IndexOfThirdQuartileClass As Integer
        Get
            Dim intClass As Integer

            If DoubleValuesMatrixRow IsNot Nothing Then
                intClass = DoubleValuesMatrixRow.GetIndexOfThirdQuartileClass
            End If

            Return intClass
        End Get
    End Property

    Public ReadOnly Property InterQuartileRange As Integer
        Get
            Dim intRange As Integer

            If DoubleValuesMatrixRow IsNot Nothing Then
                intRange = DoubleValuesMatrixRow.InterQuartileRange
            End If

            Return intRange
        End Get
    End Property
End Class

Public Class TargetFrequencyLikertGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As TargetFrequencyLikertGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As TargetFrequencyLikertGridRow)
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

    Public Sub Remove(ByVal gridrow As TargetFrequencyLikertGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As TargetFrequencyLikertGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), TargetFrequencyLikertGridRow)
            End If
        End Get
        Set(ByVal value As TargetFrequencyLikertGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As TargetFrequencyLikertGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As TargetFrequencyLikertGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
End Class
