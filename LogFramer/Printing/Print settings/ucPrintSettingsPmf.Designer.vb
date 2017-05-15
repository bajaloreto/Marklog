<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSettingsPmf
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintSettingsPmf))
        Me.GroupBoxStruct = New System.Windows.Forms.GroupBox()
        Me.cmbPrintSection = New System.Windows.Forms.ComboBox()
        Me.GroupBoxWidth = New System.Windows.Forms.GroupBox()
        Me.lblPages = New System.Windows.Forms.Label()
        Me.nudPagesWide = New System.Windows.Forms.NumericUpDown()
        Me.GroupBoxShow = New System.Windows.Forms.GroupBox()
        Me.chkShowTargetRowTitles = New System.Windows.Forms.CheckBox()
        Me.GroupBoxStruct.SuspendLayout()
        Me.GroupBoxWidth.SuspendLayout()
        CType(Me.nudPagesWide, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxShow.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBoxStruct
        '
        resources.ApplyResources(Me.GroupBoxStruct, "GroupBoxStruct")
        Me.GroupBoxStruct.Controls.Add(Me.cmbPrintSection)
        Me.GroupBoxStruct.Name = "GroupBoxStruct"
        Me.GroupBoxStruct.TabStop = False
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
        'GroupBoxShow
        '
        resources.ApplyResources(Me.GroupBoxShow, "GroupBoxShow")
        Me.GroupBoxShow.Controls.Add(Me.chkShowTargetRowTitles)
        Me.GroupBoxShow.Name = "GroupBoxShow"
        Me.GroupBoxShow.TabStop = False
        '
        'chkShowTargetRowTitles
        '
        resources.ApplyResources(Me.chkShowTargetRowTitles, "chkShowTargetRowTitles")
        Me.chkShowTargetRowTitles.Name = "chkShowTargetRowTitles"
        Me.chkShowTargetRowTitles.UseVisualStyleBackColor = True
        '
        'PrintSettingsPmf
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBoxShow)
        Me.Controls.Add(Me.GroupBoxWidth)
        Me.Controls.Add(Me.GroupBoxStruct)
        Me.Name = "PrintSettingsPmf"
        Me.GroupBoxStruct.ResumeLayout(False)
        Me.GroupBoxWidth.ResumeLayout(False)
        Me.GroupBoxWidth.PerformLayout()
        CType(Me.nudPagesWide, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxShow.ResumeLayout(False)
        Me.GroupBoxShow.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxStruct As System.Windows.Forms.GroupBox
    Friend WithEvents cmbPrintSection As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBoxWidth As System.Windows.Forms.GroupBox
    Friend WithEvents lblPages As System.Windows.Forms.Label
    Friend WithEvents nudPagesWide As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBoxShow As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowTargetRowTitles As System.Windows.Forms.CheckBox

End Class
