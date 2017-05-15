Imports System.Windows.Forms

Public Class DialogProgress
    Private intIndex As Integer

    Public Property Index As Integer
        Get
            Return intIndex
        End Get
        Set(ByVal value As Integer)
            intIndex = value
            ProgressBar.Value = intIndex
            'If ProgressBar.Value = ProgressBar.Maximum Then Me.Hide()
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub SetMaximum(ByVal RowCount As Integer)
        With ProgressBar
            .Maximum = RowCount
            .Value = 0
        End With
    End Sub

    Public Sub SetStatus(ByVal strStatus As String)
        lblStatus.Text = strStatus
    End Sub
End Class
