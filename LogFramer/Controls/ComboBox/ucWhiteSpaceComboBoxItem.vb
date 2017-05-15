Public Class WhiteSpaceComboBoxItem
    Private strWhiteSpace As String
    Private strDescription As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal Description As String, ByVal WhiteSpace As String)
        Me.Description = Description
        Me.WhiteSpace = WhiteSpace
    End Sub

    Public Property WhiteSpace() As String
        Get
            Return strWhiteSpace
        End Get
        Set(ByVal value As String)
            strWhiteSpace = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return strDescription
        End Get
        Set(ByVal value As String)
            strDescription = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return Me.WhiteSpace

    End Function
End Class
