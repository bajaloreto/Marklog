Public Class DataGridViewExcelPreview
    Inherits DataGridView

    Private dtWorkSheet As System.Data.DataTable

    Public Property WorkSheet() As System.Data.DataTable
        Get
            Return dtWorkSheet
        End Get
        Set(ByVal value As System.Data.DataTable)
            dtWorkSheet = value
            Me.DataSource = dtWorkSheet
        End Set
    End Property

    Public Sub New()

        Me.Columns.Clear()
        Me.AutoGenerateColumns = True
        Me.AllowUserToResizeRows = True
        Me.AllowUserToResizeColumns = True
        Me.AllowUserToAddRows = False
        Me.AllowUserToDeleteRows = False
        Me.ShowCellToolTips = False
        Me.MultiSelect = True
        Me.ReadOnly = True
        Me.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.BackgroundColor = Color.White

        With DefaultCellStyle
            .Alignment = DataGridViewContentAlignment.TopLeft
        End With
        With ColumnHeadersDefaultCellStyle
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .WrapMode = DataGridViewTriState.True
        End With
    End Sub

    Public Sub LoadColumns(ByVal dtExcelSheet As System.Data.DataTable)
        Dim strColName As String = String.Empty

        For i = 0 To dtExcelSheet.Columns.Count - 1
            If i < 26 Then
                strColName = Chr(Asc("A") + i)
            ElseIf i >= 26 Then
                Dim intIndexFirstChar As Integer = Int(i / 26) - 1
                strColName = Chr(Asc("A") + intIndexFirstChar)
                Dim intIndexSecondChar As Integer = i - (Int(i / 26) * 26)
                strColName &= Chr(Asc("A") + intIndexSecondChar)
            End If
            dtExcelSheet.Columns(i).ColumnName = strColName
        Next
        Me.WorkSheet = dtExcelSheet
    End Sub

    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim rCell As System.Drawing.Rectangle = e.CellBounds
        Dim penLight As New Pen(SystemColors.ControlLight, 1)
        Dim penLightLight As Pen = New Pen(SystemColors.ControlLightLight, 1)
        Dim penDark As New Pen(SystemColors.ControlDark, 1)
        Dim penDarkDark As Pen = New Pen(SystemColors.ControlDarkDark, 1)
        Dim x As Integer = rCell.Left
        Dim y As Integer = rCell.Top
        Dim p As Integer = rCell.Right - 1
        Dim q As Integer = rCell.Bottom - 1

        If e.ColumnIndex = -1 And e.RowIndex > -1 Then
            e.Graphics.FillRectangle(SystemBrushes.Control, e.CellBounds)

            Dim sfRowHeader As New StringFormat
            sfRowHeader.Alignment = StringAlignment.Center
            sfRowHeader.LineAlignment = StringAlignment.Center
            Dim strRowName As String = (e.RowIndex + 1).ToString
            e.Graphics.DrawString(strRowName, Me.Font, Brushes.Black, e.CellBounds, sfRowHeader)

            e.Graphics.DrawLine(penLightLight, x, y, p, y)
            e.Graphics.DrawLine(penDark, x, q, p, q)

            e.Graphics.DrawLine(penDark, x, y, x, q)
            e.Graphics.DrawLine(penLightLight, x + 1, y, x + 1, q)

            e.Graphics.DrawLine(penDark, p, y, p, q)

            e.Handled = True
        End If
        MyBase.OnCellPainting(e)

    End Sub

    Public ReadOnly Property FirstDataRow() As Integer
        Get
            Dim intRowIndex As Integer
            Dim selRow As System.Data.DataRow
            For i = 0 To Me.Rows.Count - 1
                selRow = Me.WorkSheet.Rows(i)
                If AreAllColumnsEmpty(selRow) = False Then
                    intRowIndex = i
                    Exit For
                End If
            Next

            Return intRowIndex
        End Get
    End Property

    Public ReadOnly Property LastDataRow() As Integer
        Get
            Dim intRowIndex As Integer
            Dim selRow As System.Data.DataRow
            For i = Me.Rows.Count - 1 To 0 Step -1
                selRow = Me.WorkSheet.Rows(i)
                If AreAllColumnsEmpty(selRow) = False Then
                    intRowIndex = i
                    Exit For
                End If
            Next

            Return intRowIndex
        End Get
    End Property

    Private Function AreAllColumnsEmpty(ByVal dr As DataRow) As Boolean
        If dr Is Nothing Then
            Return True
        Else
            For Each value In dr.ItemArray
                If IsDBNull(value) = False Then
                    Return False
                End If
            Next
            Return True
        End If
    End Function
End Class
