
Imports System.Drawing.Drawing2D

Public Class ucTabControlProject
    Inherits TabControl

    Private objTabSize As Size

    Public Sub New()
        DrawMode = TabDrawMode.OwnerDrawFixed
        SizeMode = TabSizeMode.Normal

        'If TabSize.IsEmpty() Then
        '    TabSize = New Size(200, 36)
        'End If
    End Sub

    Public Property TabSize As Size
        Get
            Return objTabSize
        End Get
        Set(ByVal value As Size)
            objTabSize = value
            Me.ItemSize = TabSize
        End Set
    End Property

    Protected Overrides Sub OnDrawItem(ByVal e As System.Windows.Forms.DrawItemEventArgs)
        MyBase.OnDrawItem(e)

        Dim fntText As Font = Me.Font
        Dim intIndex As Integer = e.Index
        Dim rTab As Rectangle = GetTabRect(intIndex)
        Dim graph As Graphics = e.Graphics
        Dim objImage As Image = Nothing
        Dim rText As Rectangle = New Rectangle(rTab.Left + CONST_Padding, rTab.Bottom - fntText.Height - CONST_VerticalPadding, rTab.Width - CONST_HorizontalPadding, fntText.Height)
        Dim strLabel As String = String.Empty
        Dim intHeight As Integer = (e.Bounds.Height / 2) + 1
        Dim colDark As Color = Color.FromArgb(255, 201, 218, 237)
        Dim colLight As Color = Color.FromArgb(255, 233, 243, 255)
        Dim rTop As New Rectangle(e.Bounds.Left, e.Bounds.Top, e.Bounds.Width, intHeight)
        Dim rBottom As New Rectangle(e.Bounds.Left, rTop.Bottom - 1, e.Bounds.Width, intHeight)

        Dim brBackgroundTop As New LinearGradientBrush(rTop, colLight, colDark, LinearGradientMode.Vertical)
        Dim brBackgroundBottom As New LinearGradientBrush(rBottom, colDark, colLight, LinearGradientMode.Vertical)


        graph.FillRectangle(brBackgroundBottom, rBottom)
        graph.FillRectangle(brBackgroundTop, rTop)

        Select Case intIndex
            Case 0 'Project info
                objImage = My.Resources.ProjectInfo_large
                strLabel = LANG_Project
            Case 1 'Logframe
                objImage = My.Resources.FileProject_large
                strLabel = LANG_Logframe
            Case 2 'Planning
                objImage = My.Resources.Planning_large
                strLabel = LANG_Planning
            Case 3 'Budget
                objImage = My.Resources.Budget_large
                strLabel = LANG_Budget
        End Select

        If objImage IsNot Nothing Then
            Dim rImage As New Rectangle(rTab.Left + CONST_Padding, rTab.Top + CONST_Padding, objImage.Width + CONST_HorizontalPadding, objImage.Height + CONST_VerticalPadding)
            graph.DrawImage(objImage, rImage)

            rText.X = rImage.Right + CONST_Padding
        End If

        graph.DrawString(strLabel, fntText, Brushes.Black, rText)
    End Sub
End Class
