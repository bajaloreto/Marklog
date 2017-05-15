<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogActivity
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogActivity))
        Me.btnReady = New System.Windows.Forms.Button()
        Me.tbRtf = New System.Windows.Forms.RichTextBox()
        Me.lblActivity = New System.Windows.Forms.Label()
        Me.PanelActivity = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        '
        'btnReady
        '
        resources.ApplyResources(Me.btnReady, "btnReady")
        Me.btnReady.Name = "btnReady"
        '
        'tbRtf
        '
        resources.ApplyResources(Me.tbRtf, "tbRtf")
        Me.tbRtf.Name = "tbRtf"
        '
        'lblActivity
        '
        resources.ApplyResources(Me.lblActivity, "lblActivity")
        Me.lblActivity.Name = "lblActivity"
        '
        'PanelActivity
        '
        resources.ApplyResources(Me.PanelActivity, "PanelActivity")
        Me.PanelActivity.Name = "PanelActivity"
        '
        'DialogActivity
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.PanelActivity)
        Me.Controls.Add(Me.lblActivity)
        Me.Controls.Add(Me.tbRtf)
        Me.Controls.Add(Me.btnReady)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogActivity"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnReady As System.Windows.Forms.Button
    Friend WithEvents tbRtf As System.Windows.Forms.RichTextBox
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents PanelActivity As System.Windows.Forms.Panel

End Class
