Imports System.Drawing.Printing

Public Class ucPrintDetail
    Friend WithEvents PrintSettingsBar As New PrintSettings

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        'PrintSettingsBar.Dock = DockStyle.Fill
        SplitContainerPrint.Panel1.Controls.Add(PrintSettingsBar)
    End Sub

    Private Sub ToolStripButtonFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonFirst.Click
        CustomPrintPreview.StartPage = 0
    End Sub

    Private Sub ToolStripButtonPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonPrevious.Click
        CustomPrintPreview.StartPage -= 1
    End Sub

    Private Sub ToolStripButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonNext.Click
        CustomPrintPreview.StartPage += 1
    End Sub

    Private Sub ToolStripButtonLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonLast.Click
        CustomPrintPreview.StartPage = CustomPrintPreview.PageCount - 1
    End Sub

    Private Sub PageCountChanged(ByVal sender As Object, ByVal e As EventArgs) Handles CustomPrintPreview.PageCountChanged
        MyBase.Update()
        Application.DoEvents()

        ToolStripLabelPageCount.Text = String.Format("{0} {1}", LANG_Of, CustomPrintPreview.PageCount)
    End Sub

    Private Sub StartPageChanged(ByVal sender As Object, ByVal e As EventArgs) Handles CustomPrintPreview.StartPageChanged
        ToolStripTextBoxPage.Text = (Me.CustomPrintPreview.StartPage + 1).ToString
    End Sub

    Private Sub ToolStripTextBoxPage_Enter(ByVal sender As Object, ByVal e As EventArgs) Handles ToolStripTextBoxPage.Enter
        ToolStripTextBoxPage.SelectAll()
    End Sub

    Private Sub ToolStripTextBoxPage_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles ToolStripTextBoxPage.KeyPress
        Dim c As Char = e.KeyChar
        If (c = ChrW(13)) Then
            Me.CommitPageNumber()
            e.Handled = True
        ElseIf Not ((c <= " "c) OrElse Char.IsDigit(c)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ToolStripTextBoxPage_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ToolStripTextBoxPage.Validating
        Me.CommitPageNumber()
    End Sub

    Private Sub CommitPageNumber()
        Dim page As Integer
        If Integer.TryParse(ToolStripTextBoxPage.Text, page) Then
            Me.CustomPrintPreview.StartPage = (page - 1)
        End If
    End Sub

    Private Sub CmbZoom_ButtonClick(ByVal sender As Object, ByVal e As EventArgs) Handles cmbZoom.ButtonClick
        CustomPrintPreview.ZoomMode = IIf((CustomPrintPreview.ZoomMode = ZoomMode.ActualSize), ZoomMode.FullPage, ZoomMode.ActualSize)
    End Sub

    Private Sub CmbZoom_DropDownItemClicked(ByVal sender As Object, ByVal e As ToolStripItemClickedEventArgs) Handles cmbZoom.DropDownItemClicked
        If (e.ClickedItem Is ItemActualSize) Then
            CustomPrintPreview.ZoomMode = ZoomMode.ActualSize
        ElseIf (e.ClickedItem Is ItemFullPage) Then
            CustomPrintPreview.ZoomMode = ZoomMode.FullPage
        ElseIf (e.ClickedItem Is ItemPageWidth) Then
            CustomPrintPreview.ZoomMode = ZoomMode.PageWidth
        ElseIf (e.ClickedItem Is ItemTwoPages) Then
            CustomPrintPreview.ZoomMode = ZoomMode.TwoPages
        End If
        If (e.ClickedItem Is Zoom10) Then
            CustomPrintPreview.Zoom = 0.1
        ElseIf (e.ClickedItem Is Zoom100) Then
            CustomPrintPreview.Zoom = 1
        ElseIf (e.ClickedItem Is Zoom150) Then
            CustomPrintPreview.Zoom = 1.5
        ElseIf (e.ClickedItem Is Zoom200) Then
            CustomPrintPreview.Zoom = 2
        ElseIf (e.ClickedItem Is Zoom25) Then
            CustomPrintPreview.Zoom = 0.25
        ElseIf (e.ClickedItem Is Zoom50) Then
            CustomPrintPreview.Zoom = 0.5
        ElseIf (e.ClickedItem Is Zoom500) Then
            CustomPrintPreview.Zoom = 5
        ElseIf (e.ClickedItem Is Zoom75) Then
            CustomPrintPreview.Zoom = 0.75
        End If
    End Sub

    Private Sub PrintSettingsBar_PrintButtonClicked() Handles PrintSettingsBar.PrintButtonClicked
        CustomPrintPreview.Print()
    End Sub

    Private Sub PrintSettingsBar_SelectPrinterButtonClicked() Handles PrintSettingsBar.SelectPrinterButtonClicked

        Dim PrintDialog As New PrintDialog
        With PrintDialog
            .UseEXDialog = True
            .AllowCurrentPage = True
            .AllowSomePages = True
            .PrinterSettings = MyPrintersettings
            .Document = CustomPrintPreview.Document

            Dim result As DialogResult = .ShowDialog(Me)
            If result = DialogResult.OK Then
                CustomPrintPreview.Print()
            End If
            MyPrintersettings = .PrinterSettings
        End With
    End Sub

    Private Sub PrintSettingsBar_ReportSetupChanged() Handles PrintSettingsBar.ReportSetupChanged
        CustomPrintPreview.RefreshPreview()
    End Sub
End Class
