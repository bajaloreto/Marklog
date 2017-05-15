Public Class DetailOutput
    Private OutputBindingSource As New BindingSource
    Private indCurrentOutput As Output

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal Output As Output)
        InitializeComponent()

        Me.CurrentOutput = Output
        Me.lvDetailKeyMoments.KeyMoments = Me.CurrentOutput.KeyMoments
        Me.lvDetailKeyMoments.OutputText = Me.CurrentOutput.Text
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Property CurrentOutput() As Output
        Get
            Return indCurrentOutput
        End Get
        Set(ByVal value As Output)
            indCurrentOutput = value
        End Set
    End Property

    Public Sub LoadItems()
        If CurrentOutput IsNot Nothing Then
            OutputBindingSource.DataSource = CurrentOutput

            tbOutput.DataBindings.Add("Text", OutputBindingSource, "Text")
        End If
    End Sub
End Class
