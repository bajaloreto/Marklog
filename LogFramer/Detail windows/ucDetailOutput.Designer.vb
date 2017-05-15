<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailOutput
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailOutput))
        Me.lblOutput = New System.Windows.Forms.Label()
        Me.tbOutput = New System.Windows.Forms.TextBox()
        Me.lvDetailKeyMoments = New FaciliDev.LogFramer.ListViewKeyMoments()
        Me.SuspendLayout()
        '
        'lblOutput
        '
        resources.ApplyResources(Me.lblOutput, "lblOutput")
        Me.lblOutput.Name = "lblOutput"
        '
        'tbOutput
        '
        resources.ApplyResources(Me.tbOutput, "tbOutput")
        Me.tbOutput.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbOutput.ForeColor = System.Drawing.Color.Blue
        Me.tbOutput.Name = "tbOutput"
        Me.tbOutput.ReadOnly = True
        '
        'lvDetailKeyMoments
        '
        resources.ApplyResources(Me.lvDetailKeyMoments, "lvDetailKeyMoments")
        Me.lvDetailKeyMoments.FullRowSelect = True
        Me.lvDetailKeyMoments.KeyMoments = Nothing
        Me.lvDetailKeyMoments.Name = "lvDetailKeyMoments"
        Me.lvDetailKeyMoments.OutputText = Nothing
        Me.lvDetailKeyMoments.SortColumnIndex = -1
        Me.lvDetailKeyMoments.UseCompatibleStateImageBehavior = False
        Me.lvDetailKeyMoments.View = System.Windows.Forms.View.Details
        '
        'DetailOutput
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.lvDetailKeyMoments)
        Me.Controls.Add(Me.tbOutput)
        Me.Controls.Add(Me.lblOutput)
        Me.Name = "DetailOutput"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblOutput As System.Windows.Forms.Label
    Friend WithEvents tbOutput As System.Windows.Forms.TextBox
    Friend WithEvents lvDetailKeyMoments As FaciliDev.LogFramer.ListViewKeyMoments

End Class
