Public Class StructuredComboBoxItem
    Private strText As String
    Private intType As Integer
    Private strUnit As String
    Private boolHeader As Boolean
    Private boolSpace As Boolean

    Public Sub New()

    End Sub

    Public Sub New(ByVal Description As String, ByVal IsHeader As Boolean, ByVal WhiteSpace As Boolean, _
                   Optional ByVal Type As Integer = Nothing, Optional ByVal Unit As String = Nothing)
        Me.Description = Description
        Me.Type = Type
        Me.Unit = Unit
        Me.IsHeader = IsHeader
        Me.WhiteSpace = WhiteSpace
    End Sub

    Public Property Description() As String
        Get
            Return strText
        End Get
        Set(ByVal value As String)
            strText = value
        End Set
    End Property

    Public Property Type() As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    Public Property Unit() As String
        Get
            Return strUnit
        End Get
        Set(ByVal value As String)
            strUnit = value
        End Set
    End Property

    Public Property IsHeader() As Boolean
        Get
            Return boolHeader
        End Get
        Set(ByVal value As Boolean)
            boolHeader = value
        End Set
    End Property

    Public Property WhiteSpace() As Boolean
        Get
            Return boolSpace
        End Get
        Set(ByVal value As Boolean)
            boolSpace = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return Me.Description
    End Function
End Class
