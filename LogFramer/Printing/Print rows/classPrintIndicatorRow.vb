Public Class PrintIndicatorRow
    Private intRowHeight As Integer
    Private strSortNumber As String = String.Empty
    Private strRtf As String = String.Empty
    Private bmFirstImage, bmSecondImage As System.Drawing.Bitmap
    Private objIndicator As Indicator
    Private objTargetDeadline As TargetDeadline
    Private strTargetDeadlineDate As String
    Private boolClippedRow As Boolean
    Private strTableHeader As String = String.Empty
    Private strColumnHeaders() As String
    Private strRowHeader As String
    Private strCells() As String
    Private intRowHeaderWidth As String
    Private intColumnWidth As String

    Private objStatement As Statement
    Private objResponseClasses As ResponseClasses
    Private intLeftLabelWidth As Integer, intRightLabelWidth As Integer
    Private intResponseColumnWidth As Integer
    Private intResponseClassesNumber As Integer
    Private boolNoCheckBox As Boolean
    Private intObjectType As Integer

    Public Enum ObjectTypes
        NotSet = -1
        TableHeader = 1
        ColumnHeader = 2
        RowHeader = 3
        WhiteSpace = 4

        Struct = 10
        Goal = 11
        Purpose = 12
        Output = 13
        Activity = 14

        Indicator = 20
        'Statement = 21
        'BooleanResponse = 22
        'IntegerResponse = 23
        'ValueRange = 24
        'ResponseClass = 25
        'Target = 26
        'Baseline = 27

        AnswerAreaOpenEnded = 30
        AnswerAreaMaxDiff = 31
        AnswerAreaValue = 40
        AnswerAreaFormula = 41
        AnswerAreaMultipleOptions = 42
        AnswerAreaLikertType = 43
        AnswerAreaSemanticDiff = 44
        AnswerAreaScale = 45
        AnswerAreaLikertScale = 46
        AnswerAreaCumulativeScale = 47
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal sortnumber As String, ByVal rtf As String)
        Me.ObjectType = objecttype
        Me.SortNumber = sortnumber
        Me.Rtf = rtf
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal indicator As Indicator)
        Me.ObjectType = objecttype
        Me.Indicator = indicator
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal tableheader As String)
        Me.ObjectType = objecttype
        Me.TableHeader = tableheader
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal columnheaders() As String)
        Me.ObjectType = objecttype
        Me.ColumnHeaders = columnheaders
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal columnheaders() As String, ByVal rowheaderwidth As Integer, ByVal columnwidth As Integer)
        Me.ObjectType = objecttype
        Me.ColumnHeaders = columnheaders
        Me.RowHeaderWidth = rowheaderwidth
        Me.ColumnWidth = columnwidth
    End Sub

    Public Sub New(ByVal objecttype As Integer, ByVal rowheader As String, ByVal cells() As String, ByVal rowheaderwidth As Integer, ByVal columnwidth As Integer, Optional ByVal nocheckbox As Boolean = False)
        Me.ObjectType = objecttype
        Me.RowHeader = rowheader
        Me.Cells = cells
        Me.RowHeaderWidth = rowheaderwidth
        Me.ColumnWidth = columnwidth
        Me.NoCheckBox = nocheckbox
    End Sub

    Public Sub New(ByVal whitespace As Integer)
        Me.ObjectType = ObjectTypes.WhiteSpace
        Me.RowHeight = whitespace
    End Sub

#Region "Properties"
    Public Property ObjectType As Integer
        Get
            Return intObjectType
        End Get
        Set(ByVal value As Integer)
            intObjectType = value
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

    Public Property SortNumber As String
        Get
            Return strSortNumber
        End Get
        Set(ByVal value As String)
            If value = Nothing Then value = String.Empty
            strSortNumber = value
        End Set
    End Property

    Public Property Rtf As String
        Get
            Return strRtf
        End Get
        Set(ByVal value As String)
            If value = Nothing Then value = String.Empty
            strRtf = value
        End Set
    End Property

    Public Property FirstImage As Bitmap
        Get
            Return bmFirstImage
        End Get
        Set(ByVal value As Bitmap)
            bmFirstImage = value
        End Set
    End Property

    Public Property SecondImage As Bitmap
        Get
            Return bmSecondImage
        End Get
        Set(ByVal value As Bitmap)
            bmSecondImage = value
        End Set
    End Property

    Public Property Indicator As Indicator
        Get
            Return objIndicator
        End Get
        Set(ByVal value As Indicator)
            objIndicator = value
        End Set
    End Property

    Public Property TargetDeadline As TargetDeadline
        Get
            Return objTargetDeadline
        End Get
        Set(ByVal value As TargetDeadline)
            objTargetDeadline = value
        End Set
    End Property

    Public Property TargetDeadlineDate As String
        Get
            Return strTargetDeadlineDate
        End Get
        Set(ByVal value As String)
            strTargetDeadlineDate = value
        End Set
    End Property

    Public Property NoCheckBox As Boolean
        Get
            Return boolNoCheckBox
        End Get
        Set(ByVal value As Boolean)
            boolNoCheckBox = value
        End Set
    End Property

    Public Property TableHeader As String
        Get
            Return strTableHeader
        End Get
        Set(ByVal value As String)
            strTableHeader = value
        End Set
    End Property

    Public Property ColumnHeaders As String()
        Get
            Return strColumnHeaders
        End Get
        Set(ByVal value As String())
            strColumnHeaders = value
        End Set
    End Property

    Public Property ColumnWidth As String
        Get
            Return intColumnWidth
        End Get
        Set(ByVal value As String)
            intColumnWidth = value
        End Set
    End Property

    Public Property RowHeader As String
        Get
            Return strRowHeader
        End Get
        Set(ByVal value As String)
            strRowHeader = value
        End Set
    End Property

    Public Property RowHeaderWidth As String
        Get
            Return intRowHeaderWidth
        End Get
        Set(ByVal value As String)
            intRowHeaderWidth = value
        End Set
    End Property

    Public Property Cells As String()
        Get
            Return strCells
        End Get
        Set(ByVal value As String())
            strCells = value
        End Set
    End Property

    Public ReadOnly Property QuestionType() As Integer
        Get
            If objIndicator IsNot Nothing Then
                Return objIndicator.QuestionType
            Else
                Return ObjectTypes.NotSet
            End If
        End Get
    End Property

    Public Property ClippedRow As Boolean
        Get
            Return boolClippedRow
        End Get
        Set(value As Boolean)
            boolClippedRow = value
        End Set
    End Property

    Public Property Statement() As Statement
        Get
            Return objStatement
        End Get
        Set(ByVal value As Statement)
            objStatement = value
        End Set
    End Property

    Public Property ResponseClasses() As ResponseClasses
        Get
            Return objResponseClasses
        End Get
        Set(ByVal value As ResponseClasses)
            objResponseClasses = value
        End Set
    End Property

    Public Property ResponseClassesNumber() As Integer
        Get
            Return intResponseClassesNumber
        End Get
        Set(ByVal value As Integer)
            intResponseClassesNumber = value
        End Set
    End Property

    Public Property LeftLabelWidth() As Integer
        Get
            Return intLeftLabelWidth
        End Get
        Set(ByVal value As Integer)
            intLeftLabelWidth = value
        End Set
    End Property

    Public Property RightLabelWidth() As Integer
        Get
            Return intRightLabelWidth
        End Get
        Set(ByVal value As Integer)
            intRightLabelWidth = value
        End Set
    End Property

    Public Property ResponseClassWidth() As Integer
        Get
            Return intResponseColumnWidth
        End Get
        Set(ByVal value As Integer)
            intResponseColumnWidth = value
        End Set
    End Property
#End Region

End Class

Public Class PrintIndicatorRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintIndicatorRow)
        List.Add(row)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintIndicatorRow)
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

    Public Sub Remove(ByVal row As PrintIndicatorRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("PrintIndicatorRow not in list!")
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

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintIndicatorRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintIndicatorRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintIndicatorRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintIndicatorRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
