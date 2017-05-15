<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucPrintSettingsTargetGroupIdForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucPrintSettingsTargetGroupIdForm))
        Me.GroupBoxTargetGroups = New System.Windows.Forms.GroupBox()
        Me.cmbSelectTargetGroup = New System.Windows.Forms.ComboBox()
        Me.gbAppearance = New System.Windows.Forms.GroupBox()
        Me.chkFillCells = New System.Windows.Forms.CheckBox()
        Me.chkPrintBorders = New System.Windows.Forms.CheckBox()
        Me.GroupBoxPurposes = New System.Windows.Forms.GroupBox()
        Me.cmbSelectPurpose = New System.Windows.Forms.ComboBox()
        Me.GroupBoxTargetGroups.SuspendLayout()
        Me.gbAppearance.SuspendLayout()
        Me.GroupBoxPurposes.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBoxTargetGroups
        '
        resources.ApplyResources(Me.GroupBoxTargetGroups, "GroupBoxTargetGroups")
        Me.GroupBoxTargetGroups.Controls.Add(Me.cmbSelectTargetGroup)
        Me.GroupBoxTargetGroups.Name = "GroupBoxTargetGroups"
        Me.GroupBoxTargetGroups.TabStop = False
        '
        'cmbSelectTargetGroup
        '
        resources.ApplyResources(Me.cmbSelectTargetGroup, "cmbSelectTargetGroup")
        Me.cmbSelectTargetGroup.FormattingEnabled = True
        Me.cmbSelectTargetGroup.Name = "cmbSelectTargetGroup"
        '
        'gbAppearance
        '
        resources.ApplyResources(Me.gbAppearance, "gbAppearance")
        Me.gbAppearance.Controls.Add(Me.chkFillCells)
        Me.gbAppearance.Controls.Add(Me.chkPrintBorders)
        Me.gbAppearance.Name = "gbAppearance"
        Me.gbAppearance.TabStop = False
        '
        'chkFillCells
        '
        resources.ApplyResources(Me.chkFillCells, "chkFillCells")
        Me.chkFillCells.Checked = True
        Me.chkFillCells.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFillCells.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_table_fill
        Me.chkFillCells.Name = "chkFillCells"
        Me.chkFillCells.UseVisualStyleBackColor = True
        '
        'chkPrintBorders
        '
        resources.ApplyResources(Me.chkPrintBorders, "chkPrintBorders")
        Me.chkPrintBorders.Checked = True
        Me.chkPrintBorders.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPrintBorders.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_table_grid
        Me.chkPrintBorders.Name = "chkPrintBorders"
        Me.chkPrintBorders.UseVisualStyleBackColor = True
        '
        'GroupBoxPurposes
        '
        resources.ApplyResources(Me.GroupBoxPurposes, "GroupBoxPurposes")
        Me.GroupBoxPurposes.Controls.Add(Me.cmbSelectPurpose)
        Me.GroupBoxPurposes.Name = "GroupBoxPurposes"
        Me.GroupBoxPurposes.TabStop = False
        '
        'cmbSelectPurpose
        '
        resources.ApplyResources(Me.cmbSelectPurpose, "cmbSelectPurpose")
        Me.cmbSelectPurpose.FormattingEnabled = True
        Me.cmbSelectPurpose.Name = "cmbSelectPurpose"
        '
        'ucPrintSettingsTargetGroupIdForm
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBoxPurposes)
        Me.Controls.Add(Me.gbAppearance)
        Me.Controls.Add(Me.GroupBoxTargetGroups)
        Me.Name = "ucPrintSettingsTargetGroupIdForm"
        Me.GroupBoxTargetGroups.ResumeLayout(False)
        Me.gbAppearance.ResumeLayout(False)
        Me.gbAppearance.PerformLayout()
        Me.GroupBoxPurposes.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxTargetGroups As System.Windows.Forms.GroupBox
    Friend WithEvents cmbSelectTargetGroup As System.Windows.Forms.ComboBox
    Friend WithEvents gbAppearance As System.Windows.Forms.GroupBox
    Friend WithEvents chkFillCells As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintBorders As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBoxPurposes As System.Windows.Forms.GroupBox
    Friend WithEvents cmbSelectPurpose As System.Windows.Forms.ComboBox

End Class
