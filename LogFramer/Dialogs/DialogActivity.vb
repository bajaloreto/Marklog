Imports System.Windows.Forms

Public Class DialogActivity
    Private objCurrentActivity As Activity
    Private ActivityBindingSource As New BindingSource

    Friend WithEvents ActivityDetail As DetailActivity

    Public Property CurrentActivity() As Activity
        Get
            Return objCurrentActivity
        End Get
        Set(ByVal value As Activity)
            objCurrentActivity = value
        End Set
    End Property

    Public Sub New(ByVal activity As Activity)
        InitializeComponent()
        CurrentActivity = activity

        ActivityDetail = New DetailActivity(CurrentActivity)
        With ActivityDetail
            .TabControlActivity.SelectTab(0)
            .lblActivity.Visible = False
            .tbActivity.Visible = False
        End With

        PanelActivity.Controls.Add(ActivityDetail)
        tbRtf.Font = CurrentText.Font

        LoadItems()
    End Sub

    Public Sub LoadItems()
        If CurrentActivity IsNot Nothing Then
            ActivityBindingSource.DataSource = CurrentActivity

            tbRtf.DataBindings.Add("Rtf", ActivityBindingSource, "RTF")
        End If
    End Sub

    Private Sub btnReady_Click(sender As Object, e As System.EventArgs) Handles btnReady.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class
