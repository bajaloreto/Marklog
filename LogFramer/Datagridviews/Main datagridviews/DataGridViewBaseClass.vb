Public MustInherit Class DataGridViewBaseClass
    Inherits DataGridView

    Public MustOverride ReadOnly Property CurrentItem(ByVal OnlyIfTextShows As Boolean) As Object

    Public MustOverride Sub CutItems(ByVal ShowWarning As Boolean)
    Public MustOverride Sub CopyItems()
    Public MustOverride Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal PasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
    Public MustOverride Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)
        CurrentControl = Me
    End Sub
End Class
