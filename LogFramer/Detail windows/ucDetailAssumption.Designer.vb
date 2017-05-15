<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailAssumption
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailAssumption))
        Me.tbAssumption = New FaciliDev.LogFramer.TextBoxLF()
        Me.cmbRaidType = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.lblRaidType = New System.Windows.Forms.Label()
        Me.PanelRaid = New System.Windows.Forms.Panel()
        Me.SuspendLayout()
        '
        'tbAssumption
        '
        resources.ApplyResources(Me.tbAssumption, "tbAssumption")
        Me.tbAssumption.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbAssumption.ForeColor = System.Drawing.Color.Blue
        Me.tbAssumption.Name = "tbAssumption"
        Me.tbAssumption.ReadOnly = True
        '
        'cmbRaidType
        '
        Me.cmbRaidType.FormattingEnabled = True
        resources.ApplyResources(Me.cmbRaidType, "cmbRaidType")
        Me.cmbRaidType.Name = "cmbRaidType"
        '
        'lblRaidType
        '
        resources.ApplyResources(Me.lblRaidType, "lblRaidType")
        Me.lblRaidType.Name = "lblRaidType"
        '
        'PanelRaid
        '
        resources.ApplyResources(Me.PanelRaid, "PanelRaid")
        Me.PanelRaid.Name = "PanelRaid"
        '
        'DetailAssumption
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.PanelRaid)
        Me.Controls.Add(Me.lblRaidType)
        Me.Controls.Add(Me.cmbRaidType)
        Me.Controls.Add(Me.tbAssumption)
        Me.Name = "DetailAssumption"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbAssumption As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents cmbRaidType As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents lblRaidType As System.Windows.Forms.Label
    Friend WithEvents PanelRaid As System.Windows.Forms.Panel

End Class
