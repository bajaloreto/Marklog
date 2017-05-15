Public MustInherit Class BasicObject
    Public Event GuidChanged(ByVal sender As Object)

    Private objGuid As Guid

    Public MustOverride ReadOnly Property ItemName As String
    Public MustOverride ReadOnly Property ItemNamePlural As String

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
                RaiseEvent GuidChanged(Me)
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
            RaiseEvent GuidChanged(Me)
        End Set
    End Property
End Class
