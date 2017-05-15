<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogIndicator
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogIndicator))
        Me.btnReady = New System.Windows.Forms.Button()
        Me.tbRtf = New System.Windows.Forms.RichTextBox()
        Me.lblIndicator = New System.Windows.Forms.Label()
        Me.PanelIndicator = New System.Windows.Forms.Panel()
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
        'lblIndicator
        '
        resources.ApplyResources(Me.lblIndicator, "lblIndicator")
        Me.lblIndicator.Name = "lblIndicator"
        '
        'PanelIndicator
        '
        resources.ApplyResources(Me.PanelIndicator, "PanelIndicator")
        Me.PanelIndicator.Name = "PanelIndicator"
        '
        'DialogIndicator
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.PanelIndicator)
        Me.Controls.Add(Me.lblIndicator)
        Me.Controls.Add(Me.tbRtf)
        Me.Controls.Add(Me.btnReady)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogIndicator"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnReady As System.Windows.Forms.Button
    Friend WithEvents tbRtf As System.Windows.Forms.RichTextBox
    Friend WithEvents lblIndicator As System.Windows.Forms.Label
    Friend WithEvents PanelIndicator As System.Windows.Forms.Panel

End Class
