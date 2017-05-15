Public Class QuestionnaireGridRow
    Private intSection As Integer
    Private objStruct As Struct
    Private strSortNumber As String
    Private intIndent As Integer
    Private bmStructImage As Bitmap
    Private boolStructImageDirty As Boolean = True
    Private intStructHeight As Integer
    Private objIndicator As Indicator
    Private bmIndicatorImage As Bitmap
    Private boolIndicatorImageDirty As Boolean = True
    Private intIndicatorHeight As Integer
    Private intQuestionType As Integer
    Private boolIsIndicator As Boolean
    Private intRowType As Integer

    Public Enum RowTypes
        Section = 0
        Goal = 1
        Purpose = 2
        Output = 3
        Activity = 4
        Indicator = 5
    End Enum

    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property SortNumber() As String
        Get
            Return strSortNumber
        End Get
        Set(ByVal value As String)
            strSortNumber = value
        End Set
    End Property

    Public Property Indent As Integer
        Get
            Return intIndent
        End Get
        Set(value As Integer)
            intIndent = value
        End Set
    End Property

    Public Property Struct As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
            If objStruct IsNot Nothing Then objIndicator = Nothing
        End Set
    End Property

    Public Property StructImage As Bitmap
        Get
            Return bmStructImage
        End Get
        Set(ByVal value As Bitmap)
            bmStructImage = value
        End Set
    End Property

    Public Property StructImageDirty As Boolean
        Get
            Return boolStructImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolStructImageDirty = value
        End Set
    End Property

    Public Property StructHeight As Integer
        Get
            Return intStructHeight
        End Get
        Set(ByVal value As Integer)
            intStructHeight = value
        End Set
    End Property

    Public Property Indicator As Indicator
        Get
            Return objIndicator
        End Get
        Set(value As Indicator)
            objIndicator = value
            If objIndicator IsNot Nothing Then objStruct = Nothing
        End Set
    End Property

    Public Property IndicatorImage As Bitmap
        Get
            Return bmIndicatorImage
        End Get
        Set(ByVal value As Bitmap)
            bmIndicatorImage = value
        End Set
    End Property

    Public Property IndicatorImageDirty As Boolean
        Get
            Return boolIndicatorImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolIndicatorImageDirty = value
        End Set
    End Property

    Public Property IndicatorHeight As Integer
        Get
            Return intIndicatorHeight
        End Get
        Set(ByVal value As Integer)
            intIndicatorHeight = value
        End Set
    End Property

    Public ReadOnly Property IsIndicator() As Boolean
        Get
            If Indicator Is Nothing Then Return False Else Return True
        End Get
    End Property

    Public Property RowType As Integer
        Get
            Return intRowType
        End Get
        Set(ByVal value As Integer)
            intRowType = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal intRowType As Integer)
        Me.RowType = intRowType
    End Sub
End Class

Public Class QuestionnaireGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As QuestionnaireGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As QuestionnaireGridRow)
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

    Public Sub Remove(ByVal gridrow As QuestionnaireGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As QuestionnaireGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), QuestionnaireGridRow)
            End If
        End Get
        Set(value As QuestionnaireGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As QuestionnaireGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As QuestionnaireGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function

    Public Function GetByGuid(ByVal FindGuid As Guid) As QuestionnaireGridRow
        Dim selGridRow As QuestionnaireGridRow
        For Each selGridRow In Me.List
            Select Case selGridRow.RowType
                Case QuestionnaireGridRow.RowTypes.Indicator
                    If selGridRow.Indicator.Guid = FindGuid Then Return selGridRow
                    'Case QuestionnaireGridRow.RowTypes.Activity
                    '    Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
                    '    If selActivity.Guid = FindGuid Then Return selGridRow
            End Select
        Next
        Return Nothing
    End Function

    Public Function GetPreviousStruct(ByVal intGridRowIndex As Integer) As Struct
        intGridRowIndex -= 1
        Dim selStruct As Struct = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As QuestionnaireGridRow = Me(i)
            If selGridRow.Struct IsNot Nothing Then
                selStruct = selGridRow.Struct
                Exit For
            End If
        Next

        Return selStruct
    End Function

    Public Function GetNextStruct(ByVal intGridRowIndex As Integer) As Struct
        intGridRowIndex += 1
        Dim selStruct As Struct = Nothing

        For i = intGridRowIndex To Me.Count - 1
            Dim selGridRow As QuestionnaireGridRow = Me(i)
            If selGridRow.Struct IsNot Nothing Then
                selStruct = selGridRow.Struct
                Exit For
            End If
        Next

        Return selStruct
    End Function

    Public Function GetPreviousIndicator(ByVal intGridRowIndex As Integer) As Indicator
        intGridRowIndex -= 1
        Dim selIndicator As Indicator = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As QuestionnaireGridRow = Me(i)
            If selGridRow.Indicator IsNot Nothing Then
                selIndicator = selGridRow.Indicator
                Exit For
            End If
        Next

        Return selIndicator
    End Function

    Public Function GetNextIndicator(ByVal intGridRowIndex As Integer) As Indicator
        intGridRowIndex += 1
        Dim selIndicator As Indicator = Nothing

        For i = intGridRowIndex To Me.Count - 1
            Dim selGridRow As QuestionnaireGridRow = Me(i)
            If selGridRow.Indicator IsNot Nothing Then
                selIndicator = selGridRow.Indicator
                Exit For
            End If
        Next

        Return selIndicator
    End Function
End Class
