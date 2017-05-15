Public Class StructuredComboBox
    Inherits ComboBoxSelectValue

    Public Sub New()
        Me.DrawMode = Windows.Forms.DrawMode.OwnerDrawVariable
        Me.DropDownStyle = ComboBoxStyle.DropDown
    End Sub

    Protected Overrides Sub OnMeasureItem(ByVal e As System.Windows.Forms.MeasureItemEventArgs)
        If e.Index >= 0 Then
            Dim strText As String = Items(e.Index).Description
            Dim boolIsHeader As Boolean = Items(e.Index).IsHeader
            Dim boolWhiteSpace As Boolean = Items(e.Index).WhiteSpace
            Dim fntNormal As New Font(Me.Font, FontStyle.Regular)
            Dim fntBold As New Font(Me.Font, FontStyle.Bold)
            Dim intHeight As Integer

            If boolIsHeader = False Then
                intHeight = e.Graphics.MeasureString(strText, fntNormal).Height * 1.1
            Else
                intHeight = e.Graphics.MeasureString(strText, fntBold).Height * 1.1
            End If
            If boolWhiteSpace = True Then intHeight += 4
            e.ItemHeight = intHeight
        End If
        MyBase.OnMeasureItem(e)
    End Sub

    Protected Overrides Sub OnDrawItem(ByVal e As System.Windows.Forms.DrawItemEventArgs)
        Dim graph As Graphics = e.Graphics
        Dim strText As String = String.Empty
        Dim boolIsHeader As Boolean
        Dim fntNormal As New Font(e.Font, FontStyle.Regular)
        Dim fntBold As New Font(e.Font, FontStyle.Bold)
        Dim brNormalBack As SolidBrush
        Dim brNormalFront As SolidBrush

        If e.Index >= 0 Then
            strText = Items(e.Index).Description
            boolIsHeader = Items(e.Index).IsHeader
        End If

        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            brNormalBack = SystemBrushes.Highlight
            brNormalFront = SystemBrushes.HighlightText
        Else
            brNormalBack = Brushes.White
            brNormalFront = Brushes.Black
        End If
        If boolIsHeader = False Then
            graph.FillRectangle(brNormalBack, e.Bounds)
            If String.IsNullOrEmpty(strText) = False Then graph.DrawString(strText, fntNormal, brNormalFront, e.Bounds.X + 10, e.Bounds.Y)
        Else
            graph.FillRectangle(Brushes.Gainsboro, e.Bounds)
            If String.IsNullOrEmpty(strText) = False Then graph.DrawString(strText, fntBold, Brushes.Black, e.Bounds.X, e.Bounds.Y)
        End If

        MyBase.OnDrawItem(e)
    End Sub

    Public Sub SelectItemByType(ByVal intType As Integer)
        For i = 0 To Items.Count - 1
            If Items(i).Type = intType Then
                SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Public Sub SelectItemByUnit(ByVal strUnit As String)
        For i = 0 To Items.Count - 1
            If Items(i).Unit = strUnit Then
                SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Public Sub SelectItemByText(ByVal strText As String)
        If strText = Nothing Then Exit Sub

        For i = 0 To Items.Count - 1
            If Items(i).Description = strText Then
                SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
End Class
