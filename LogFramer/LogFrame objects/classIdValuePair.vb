Public Class IdValuePair
    Private objId As Object
    Private strValue As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal id As Object, ByVal value As String)
        Me.Id = id
        Me.Value = value
    End Sub

    Public Property Id() As Object
        Get
            Return objId
        End Get
        Set(ByVal value As Object)
            objId = value
        End Set
    End Property

    Public Property Value() As String
        Get
            Return strValue
        End Get
        Set(ByVal value As String)
            strValue = value
        End Set
    End Property
End Class
