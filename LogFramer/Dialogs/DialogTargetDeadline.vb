Imports System.Windows.Forms

Public Class DialogTargetDeadline
    Private datDeadline As Date

#Region "Properties"
    Public Property Deadline As Date
        Get
            Return datDeadline
        End Get
        Set(value As Date)
            datDeadline = value
        End Set
    End Property
#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal deadline As Date)
        InitializeComponent()

        Me.Deadline = deadline
        
        Text = LANG_Target
        dtbDeadline.DateValue = Me.Deadline
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dtbDeadline_Validated(sender As Object, e As System.EventArgs) Handles dtbDeadline.Validated
        Date.TryParse(dtbDeadline.DateValue, Deadline)
    End Sub
End Class
