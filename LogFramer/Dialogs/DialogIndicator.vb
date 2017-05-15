Imports System.Windows.Forms

Public Class DialogIndicator
    Private objCurrentIndicator As Indicator
    Private IndicatorBindingSource As New BindingSource

    Friend WithEvents IndicatorDetail As DetailIndicator

    Public Property CurrentIndicator() As Indicator
        Get
            Return objCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            objCurrentIndicator = value
        End Set
    End Property

    Public Sub New(ByVal indicator As Indicator)
        InitializeComponent()
        Me.CurrentIndicator = indicator

        IndicatorDetail = New DetailIndicator(CurrentIndicator)
        With IndicatorDetail
            .TabControlIndicator.SelectTab(0)
            .tbIndicator.Visible = False
        End With

        PanelIndicator.Controls.Add(IndicatorDetail)
        tbRtf.Font = CurrentText.Font

        LoadItems()
    End Sub

    Public Sub LoadItems()
        If CurrentIndicator IsNot Nothing Then
            IndicatorBindingSource.DataSource = CurrentIndicator

            tbRtf.DataBindings.Add("Rtf", IndicatorBindingSource, "RTF")
        End If
    End Sub

    Private Sub btnReady_Click(sender As Object, e As System.EventArgs) Handles btnReady.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class
