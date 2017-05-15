Imports System.Windows.Forms

Public Class DialogNewUnit

    Public Property UnitText() As String
        Get
            Return tbName.Text
        End Get
        Set(ByVal value As String)
            tbName.Text = value
        End Set
    End Property

    Public Property Unit() As String
        Get
            Dim strUnit As String
            If tbAbbreviation.Text = "" Then
                strUnit = Microsoft.VisualBasic.Left(UnitText, 4).ToLower
            Else
                strUnit = tbAbbreviation.Text
            End If

            Return strUnit
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
