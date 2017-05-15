<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailPurpose
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailPurpose))
        Me.lblPurpose = New System.Windows.Forms.Label()
        Me.tbPurpose = New System.Windows.Forms.TextBox()
        Me.lvDetailTargetGroups = New FaciliDev.LogFramer.ListViewTargetGroups()
        Me.SuspendLayout()
        '
        'lblPurpose
        '
        resources.ApplyResources(Me.lblPurpose, "lblPurpose")
        Me.lblPurpose.Name = "lblPurpose"
        '
        'tbPurpose
        '
        resources.ApplyResources(Me.tbPurpose, "tbPurpose")
        Me.tbPurpose.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbPurpose.ForeColor = System.Drawing.Color.Blue
        Me.tbPurpose.Name = "tbPurpose"
        Me.tbPurpose.ReadOnly = True
        '
        'lvDetailTargetGroups
        '
        resources.ApplyResources(Me.lvDetailTargetGroups, "lvDetailTargetGroups")
        Me.lvDetailTargetGroups.AllTargetGroups = False
        Me.lvDetailTargetGroups.FullRowSelect = True
        Me.lvDetailTargetGroups.Name = "lvDetailTargetGroups"
        Me.lvDetailTargetGroups.SortColumnIndex = -1
        Me.lvDetailTargetGroups.TargetGroups = Nothing
        Me.lvDetailTargetGroups.UseCompatibleStateImageBehavior = False
        Me.lvDetailTargetGroups.View = System.Windows.Forms.View.Details
        '
        'DetailPurpose
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.lvDetailTargetGroups)
        Me.Controls.Add(Me.tbPurpose)
        Me.Controls.Add(Me.lblPurpose)
        Me.Name = "DetailPurpose"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblPurpose As System.Windows.Forms.Label
    Friend WithEvents tbPurpose As System.Windows.Forms.TextBox
    Friend WithEvents lvDetailTargetGroups As FaciliDev.LogFramer.ListViewTargetGroups

End Class
