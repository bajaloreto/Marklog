Public Class DetailClipboard
    Public Event CloseButtonClicked()
    Private intContentType As Integer

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        RaiseEvent CloseButtonClicked()
    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        Select Case dgClipboard.ContentType
            Case DataGridViewClipboard.ContentTypes.Text
                TextClipboard.Clear()
            Case DataGridViewClipboard.ContentTypes.Items
                ItemClipboard.Clear()
        End Select

        dgClipboard.LoadColumns(dgClipboard.ContentType)
    End Sub

    Private Sub dgClipboard_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgClipboard.CellClick
        Select Case dgClipboard.ContentType
            Case DataGridViewClipboard.ContentTypes.Text
                frmParent.PasteText(dgClipboard.CurrentRow.Index)
            Case DataGridViewClipboard.ContentTypes.Items
                Dim datCopyGroup As Date
                If Date.TryParse(dgClipboard.CurrentRow.Cells(2).Value, datCopyGroup) Then
                    frmParent.PasteItem(datCopyGroup)
                End If
        End Select
    End Sub

    Private Sub DetailClipboard_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dgClipboard
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False

            .BackgroundColor = Color.White
            .GridColor = Color.White
            .ReadOnly = True
            .RowHeadersVisible = False
            .ColumnHeadersVisible = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            .LoadColumns()
        End With
    End Sub
End Class
