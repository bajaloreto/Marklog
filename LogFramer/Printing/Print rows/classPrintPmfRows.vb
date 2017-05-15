Public Class PrintPmfRow
    Private intSection As Integer
    Private intRowHeight As Integer
    Private boolClippedRow As Boolean
    Private strStructSortNumber As String
    Private objStruct As Struct
    Private intStructHeight As Integer
    Private intStructIndent As Integer
    Private strIndSortNumber As String
    Private objIndicator As Indicator
    Private intIndicatorHeight As Integer
    Private intIndicatorIndent As Integer
    Private intTargetsHeight As Integer
    Private strVerSortNumber As String
    Private objVerificationSource As VerificationSource
    Private intVerificationSourceHeight As Integer
    Private strCollectionMethod As String
    Private intCollectionMethodHeight As Integer
    Private bmCollectionMethodImage As System.Drawing.Bitmap
    Private strFrequency As String
    Private strResponsibility As String
    Private intRowCountStruct As Integer
    Private intRowCountIndicator As Integer

    Public Sub New()

    End Sub

    Public Sub New(ByVal section As Integer)
        Me.Section = section
    End Sub

    Public Sub New(ByVal section As Integer, ByVal structsortnumber As String, ByVal struct As Struct)
        Me.Section = section
        Me.StructSort = structsortnumber
        Me.Struct = struct
    End Sub

    Public Sub New(ByVal section As Integer, ByVal structsort As String, ByVal struct As Struct, ByVal indicatorsort As String, ByVal indicator As Indicator, _
                   ByVal verificationsourcesort As String, ByVal verificationsource As VerificationSource)
        Me.Section = section
        Me.StructSort = structsort
        Me.Struct = struct
        Me.IndicatorSort = indicatorsort
        Me.Indicator = indicator
        Me.VerificationSourceSort = verificationsourcesort
        Me.VerificationSource = verificationsource
    End Sub

    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property ClippedRow As Boolean
        Get
            Return boolClippedRow
        End Get
        Set(value As Boolean)
            boolClippedRow = value
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

    Public Property StructSort() As String
        Get
            Return strStructSortNumber
        End Get
        Set(ByVal value As String)
            strStructSortNumber = value
        End Set
    End Property

    Public Property Struct() As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
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

    Public Property StructIndent As Integer
        Get
            Return intStructIndent
        End Get
        Set(value As Integer)
            intStructIndent = value
        End Set
    End Property

    Public Property IndicatorSort() As String
        Get
            Return strIndSortNumber
        End Get
        Set(ByVal value As String)
            strIndSortNumber = value
        End Set
    End Property

    Public Property Indicator() As Indicator
        Get
            Return objIndicator
        End Get
        Set(ByVal value As Indicator)
            objIndicator = value
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

    Public Property IndicatorIndent As Integer
        Get
            Return intIndicatorIndent
        End Get
        Set(value As Integer)
            intIndicatorIndent = value
        End Set
    End Property

    Public ReadOnly Property Baseline As Baseline
        Get
            If objIndicator IsNot Nothing Then
                Return objIndicator.Baseline
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Targets() As Targets
        Get
            If objIndicator IsNot Nothing Then
                Return objIndicator.Targets
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property TargetsHeight As Integer
        Get
            Return intTargetsHeight
        End Get
        Set(ByVal value As Integer)
            intTargetsHeight = value
        End Set
    End Property

    Public Property VerificationSourceSort() As String
        Get
            Return strVerSortNumber
        End Get
        Set(ByVal value As String)
            strVerSortNumber = value
        End Set
    End Property

    Public Property VerificationSource() As VerificationSource
        Get
            Return objVerificationSource
        End Get
        Set(ByVal value As VerificationSource)
            objVerificationSource = value
        End Set
    End Property

    Public Property VerificationSourceHeight As Integer
        Get
            Return intVerificationSourceHeight
        End Get
        Set(ByVal value As Integer)
            intVerificationSourceHeight = value
        End Set
    End Property

    Public Property CollectionMethod() As String
        Get
            Return strCollectionMethod
        End Get
        Set(ByVal value As String)
            strCollectionMethod = value
        End Set
    End Property

    Public Property CollectionMethodImage() As Bitmap
        Get
            Return bmCollectionMethodImage
        End Get
        Set(ByVal value As Bitmap)
            bmCollectionMethodImage = value
        End Set
    End Property

    Public Property CollectionMethodHeight As Integer
        Get
            Return intCollectionMethodHeight
        End Get
        Set(ByVal value As Integer)
            intCollectionMethodHeight = value
        End Set
    End Property

    Public Property Frequency() As String
        Get
            Return strFrequency
        End Get
        Set(ByVal value As String)
            strFrequency = value
        End Set
    End Property

    Public Property Responsibility() As String
        Get
            Return strResponsibility
        End Get
        Set(ByVal value As String)
            strResponsibility = value
        End Set
    End Property

    Public Property RowCountStruct As Integer
        Get
            Return intRowCountStruct
        End Get
        Set(value As Integer)
            intRowCountStruct = value
        End Set
    End Property

    Public Property RowCountIndicator As Integer
        Get
            Return intRowCountIndicator
        End Get
        Set(value As Integer)
            intRowCountIndicator = value
        End Set
    End Property
End Class

Public Class PrintPmfRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintPmfRow)
        List.Add(row)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintPmfRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(row)
        Else
            List.Insert(index, row)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal row As PrintPmfRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("PmfRow not in list!")
        Else
            List.Remove(row)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintPmfRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintPmfRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintPmfRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintPmfRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
