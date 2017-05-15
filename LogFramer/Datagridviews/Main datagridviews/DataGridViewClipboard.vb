Public Class DataGridViewClipboard
    Inherits DataGridView

    Private intContentType As Integer

    Public Enum ContentTypes
        Text
        Items
    End Enum

    Public Property ContentType As Integer
        Get
            Return intContentType
        End Get
        Set(ByVal value As Integer)
            intContentType = value
            LoadColumns()
        End Set
    End Property

    Public Sub New()
        
    End Sub

    Public Sub LoadColumns(Optional ByVal contenttype As Integer = 0)
        Me.Rows.Clear()
        Me.Columns.Clear()
        intContentType = contenttype

        With DefaultCellStyle
            .Alignment = DataGridViewContentAlignment.TopLeft
            .WrapMode = DataGridViewTriState.True
        End With

        Dim colSortNumber As New DataGridViewTextBoxColumn
        With colSortNumber
            .MinimumWidth = 60
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .SortMode = DataGridViewColumnSortMode.NotSortable
        End With
        Me.Columns.Add(colSortNumber)

        Dim colText As New DataGridViewTextBoxColumn
        With colText
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .SortMode = DataGridViewColumnSortMode.NotSortable
            
        End With
        Me.Columns.Add(colText)

        Dim colCopyGroup As New DataGridViewTextBoxColumn
        With colCopyGroup
            .Width = 20
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            .SortMode = DataGridViewColumnSortMode.NotSortable

        End With
        Me.Columns.Add(colCopyGroup)

        If contenttype = ContentTypes.Text Then
            colSortNumber.Visible = False
        Else
            colSortNumber.Visible = True
        End If
        colCopyGroup.Visible = False

        ReloadClipboard()
    End Sub

    Private Sub ReloadClipboard()
        Select Case ContentType
            Case ContentTypes.Text
                For Each selItem As ClipboardTextItem In TextClipboard
                    Rows.Add({Nothing, selItem.Text, Nothing})
                Next
            Case ContentTypes.Items
                If ItemClipboard.Count > 0 Then
                    Dim datCopyGroup As Date = ItemClipboard(0).CopyGroup
                    Dim strSort As String = String.Empty
                    Dim strText As String = String.Empty

                    For Each selItem As ClipboardItem In ItemClipboard
                        If selItem.CopyGroup <> datCopyGroup Then
                            Rows.Add({strSort, strText, datCopyGroup})
                            strSort = String.Empty
                            strText = String.Empty
                        End If

                        strSort = selItem.SortNumber & IIf(String.IsNullOrEmpty(strSort), String.Empty, vbLf & strSort)
                        strText = selItem.Text & IIf(String.IsNullOrEmpty(strText), String.Empty, vbLf & strText)

                        datCopyGroup = selItem.CopyGroup
                    Next
                    If String.IsNullOrEmpty(strText) = False Then Rows.Add({strSort, strText})
                End If
        End Select
    End Sub
End Class
