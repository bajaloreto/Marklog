Imports System.Windows.Forms

Public Class DialogResponseClassValues
    Private objResponseClasses As ResponseClasses
    Private bindClasses As New BindingSource

    Public Property ResponseClasses() As ResponseClasses
        Get
            Return objResponseClasses
        End Get
        Set(ByVal value As ResponseClasses)
            objResponseClasses = value
        End Set
    End Property

    Public Sub New(ByVal responseclasses As ResponseClasses)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.ResponseClasses = responseclasses
        If Me.ResponseClasses IsNot Nothing Then
            bindClasses.DataSource = Me.ResponseClasses
            Dim strClass, strValue As String

            Select Case UserLanguage
                Case "fr"
                    strClass = "Classe"
                    strValue = "Valeur"
                Case Else
                    strClass = "Class"
                    strValue = "Value"
            End Select

            With dgvResponseClassValues
                .BackgroundColor = Color.White
                .AutoGenerateColumns = False
                .DataSource = bindClasses

                .Columns.Add("ClassName", strClass)
                .Columns.Add("Value", strValue)
                .Columns(0).DataPropertyName = "ClassName"
                .Columns(1).DataPropertyName = "Value"
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
            End With
        End If

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
