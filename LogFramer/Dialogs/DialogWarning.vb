Imports System.Windows

Public Class DialogWarning

    Private typeValue As Integer

    Public Enum WarningDialogTypes As Byte
        wdDeleteChildren = 0
    End Enum

    Public Property Type() As Integer
        Get
            Return typeValue
        End Get
        Set(ByVal value As Integer)
            typeValue = value
        End Set
    End Property

    Public Sub New(ByVal strMsg As String, ByVal strTitle As String)
        Me.InitializeComponent()
        CenterToScreen()
        lblTitle.Text = strTitle
        lblWarning.Text = strMsg
    End Sub

    'Public Sub WarnDeleteChildrenDialog(ByVal strMsg As String, ByVal strTitle As String)
    '    CenterToScreen()
    '    lblTitle.Text = strTitle
    '    lblWarning.Text = strMsg
    'End Sub

    Private Sub chkNotShow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNotShow.CheckedChanged
        Select Case typeValue
            Case WarningDialogTypes.wdDeleteChildren
                If chkNotShow.CheckState = CheckState.Checked Then My.Settings.setWarnLinkedObjectDelete = False Else _
                    My.Settings.setWarnLinkedObjectDelete = True
        End Select
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYes.Click
        DialogResult = Windows.Forms.DialogResult.Yes
        Close()
    End Sub

    Private Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNo.Click
        DialogResult = Windows.Forms.DialogResult.No
        Close()
    End Sub

    Private Sub lblWarning_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblWarning.SizeChanged

        Me.Height = lblWarning.Height + 150
    End Sub
End Class