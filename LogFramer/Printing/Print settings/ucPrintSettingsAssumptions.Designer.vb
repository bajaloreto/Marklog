<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucPrintSettingsAssumptions
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucPrintSettingsAssumptions))
        Me.GroupBoxPrintSection = New System.Windows.Forms.GroupBox()
        Me.cmbPrintSection = New System.Windows.Forms.ComboBox()
        Me.GroupBoxWidth = New System.Windows.Forms.GroupBox()
        Me.lblPages = New System.Windows.Forms.Label()
        Me.nudPagesWide = New System.Windows.Forms.NumericUpDown()
        Me.GroupBoxPrintSection.SuspendLayout()
        Me.GroupBoxWidth.SuspendLayout()
        CType(Me.nudPagesWide, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBoxPrintSection
        '
        resources.ApplyResources(Me.GroupBoxPrintSection, "GroupBoxPrintSection")
        Me.GroupBoxPrintSection.Controls.Add(Me.cmbPrintSection)
        Me.GroupBoxPrintSection.Name = "GroupBoxPrintSection"
        Me.GroupBoxPrintSection.TabStop = False
        '
        'cmbPrintSection
        '
        resources.ApplyResources(Me.cmbPrintSection, "cmbPrintSection")
        Me.cmbPrintSection.FormattingEnabled = True
        Me.cmbPrintSection.Name = "cmbPrintSection"
        '
        'GroupBoxWidth
        '
        resources.ApplyResources(Me.GroupBoxWidth, "GroupBoxWidth")
        Me.GroupBoxWidth.Controls.Add(Me.lblPages)
        Me.GroupBoxWidth.Controls.Add(Me.nudPagesWide)
        Me.GroupBoxWidth.Name = "GroupBoxWidth"
        Me.GroupBoxWidth.TabStop = False
        '
        'lblPages
        '
        resources.ApplyResources(Me.lblPages, "lblPages")
        Me.lblPages.Name = "lblPages"
        '
        'nudPagesWide
        '
        resources.ApplyResources(Me.nudPagesWide, "nudPagesWide")
        Me.nudPagesWide.Name = "nudPagesWide"
        '
        'ucPrintSettingsAssumptions
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBoxWidth)
        Me.Controls.Add(Me.GroupBoxPrintSection)
        Me.Name = "ucPrintSettingsAssumptions"
        Me.GroupBoxPrintSection.ResumeLayout(False)
        Me.GroupBoxWidth.ResumeLayout(False)
        Me.GroupBoxWidth.PerformLayout()
        CType(Me.nudPagesWide, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxPrintSection As System.Windows.Forms.GroupBox
    Friend WithEvents cmbPrintSection As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBoxWidth As System.Windows.Forms.GroupBox
    Friend WithEvents lblPages As System.Windows.Forms.Label
    Friend WithEvents nudPagesWide As System.Windows.Forms.NumericUpDown

End Class
