Public Class UndoListItem
    Private strDescription As String
    Private objOldParentItem, objNewParentItem As Object
    Private intOldIndex, intNewIndex As Integer
    Private objItem, objOldItem As Object
    Private strPropertyName As String
    Private objOldValue, objNewValue As Object
    Private strSelectedText As String
    Private intActionIndex As Integer
    Private objParameter As Object

    Public Sub New()

    End Sub

    Public Sub New(ByVal parentitem As Object, ByVal index As Integer, ByVal item As Object)
        Me.OldParentItem = parentitem
        Me.OldIndex = index
        Me.Item = item
    End Sub

#Region "Properties"
    Public Property Description As String
        Get
            Return strDescription
        End Get
        Set(ByVal value As String)
            strDescription = value
        End Set
    End Property

    Public Property OldParentItem As Object
        Get
            Return objOldParentItem
        End Get
        Set(ByVal value As Object)
            objOldParentItem = value
        End Set
    End Property

    Public Property NewParentItem As Object
        Get
            Return objNewParentItem
        End Get
        Set(ByVal value As Object)
            objNewParentItem = value
        End Set
    End Property

    Public Property OldIndex As Integer
        Get
            Return intOldIndex
        End Get
        Set(ByVal value As Integer)
            intOldIndex = value
        End Set
    End Property

    Public Property NewIndex As Integer
        Get
            Return intNewIndex
        End Get
        Set(ByVal value As Integer)
            intNewIndex = value
        End Set
    End Property

    Public Property Item As Object
        Get
            Return objItem
        End Get
        Set(ByVal value As Object)
            objItem = value
        End Set
    End Property

    Public Property OldItem As Object
        Get
            Return objOldItem
        End Get
        Set(ByVal value As Object)
            objOldItem = value
        End Set
    End Property

    Public Property PropertyName As String
        Get
            Return strPropertyName
        End Get
        Set(ByVal value As String)
            strPropertyName = value
        End Set
    End Property

    Public Property OldValue As Object
        Get
            Return objOldValue
        End Get
        Set(ByVal value As Object)
            objOldValue = value
        End Set
    End Property

    Public Property NewValue As Object
        Get
            Return objNewValue
        End Get
        Set(ByVal value As Object)
            objNewValue = value
        End Set
    End Property

    Public Property SelectedText As String
        Get
            Return strSelectedText
        End Get
        Set(ByVal value As String)
            strSelectedText = value
        End Set
    End Property

    Public Property ActionIndex() As Integer
        Get
            Return intActionIndex
        End Get
        Set(ByVal value As Integer)
            intActionIndex = value
        End Set
    End Property

    Public Property Parameter As Object
        Get
            Return objParameter
        End Get
        Set(value As Object)
            objParameter = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function GetActionText() As String
        If Me.ActionIndex >= 0 And Me.ActionIndex <= LIST_UndoActions.Count - 1 Then
            Return LIST_UndoActions(intActionIndex)
        Else
            Return "-*-"
        End If
    End Function
#End Region
End Class
