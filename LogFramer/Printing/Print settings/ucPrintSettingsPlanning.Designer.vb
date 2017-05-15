<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSettingsPlanning
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintSettingsPlanning))
        Me.GroupBoxOutput = New System.Windows.Forms.GroupBox()
        Me.cmbSelectOutput = New System.Windows.Forms.ComboBox()
        Me.GroupBoxView = New System.Windows.Forms.GroupBox()
        Me.chkShowKeyMomentLinks = New System.Windows.Forms.CheckBox()
        Me.chkShowDatesColumns = New System.Windows.Forms.CheckBox()
        Me.chkRepeatRowHeaders = New System.Windows.Forms.CheckBox()
        Me.chkShowActivityLinks = New System.Windows.Forms.CheckBox()
        Me.cmbSelectPeriodView = New System.Windows.Forms.ComboBox()
        Me.cmbShowElements = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.dtpPeriodUntil = New System.Windows.Forms.DateTimePicker()
        Me.dtpPeriodFrom = New System.Windows.Forms.DateTimePicker()
        Me.lblUntil = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.GroupBoxOutput.SuspendLayout()
        Me.GroupBoxView.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBoxOutput
        '
        resources.ApplyResources(Me.GroupBoxOutput, "GroupBoxOutput")
        Me.GroupBoxOutput.Controls.Add(Me.cmbSelectOutput)
        Me.GroupBoxOutput.Name = "GroupBoxOutput"
        Me.GroupBoxOutput.TabStop = False
        '
        'cmbSelectOutput
        '
        resources.ApplyResources(Me.cmbSelectOutput, "cmbSelectOutput")
        Me.cmbSelectOutput.FormattingEnabled = True
        Me.cmbSelectOutput.Name = "cmbSelectOutput"
        '
        'GroupBoxView
        '
        resources.ApplyResources(Me.GroupBoxView, "GroupBoxView")
        Me.GroupBoxView.Controls.Add(Me.chkShowKeyMomentLinks)
        Me.GroupBoxView.Controls.Add(Me.chkShowDatesColumns)
        Me.GroupBoxView.Controls.Add(Me.chkRepeatRowHeaders)
        Me.GroupBoxView.Controls.Add(Me.chkShowActivityLinks)
        Me.GroupBoxView.Controls.Add(Me.cmbSelectPeriodView)
        Me.GroupBoxView.Controls.Add(Me.cmbShowElements)
        Me.GroupBoxView.Name = "GroupBoxView"
        Me.GroupBoxView.TabStop = False
        '
        'chkShowKeyMomentLinks
        '
        resources.ApplyResources(Me.chkShowKeyMomentLinks, "chkShowKeyMomentLinks")
        Me.chkShowKeyMomentLinks.Name = "chkShowKeyMomentLinks"
        Me.chkShowKeyMomentLinks.UseVisualStyleBackColor = True
        '
        'chkShowDatesColumns
        '
        resources.ApplyResources(Me.chkShowDatesColumns, "chkShowDatesColumns")
        Me.chkShowDatesColumns.Name = "chkShowDatesColumns"
        Me.chkShowDatesColumns.UseVisualStyleBackColor = True
        '
        'chkRepeatRowHeaders
        '
        resources.ApplyResources(Me.chkRepeatRowHeaders, "chkRepeatRowHeaders")
        Me.chkRepeatRowHeaders.Name = "chkRepeatRowHeaders"
        Me.chkRepeatRowHeaders.UseVisualStyleBackColor = True
        '
        'chkShowActivityLinks
        '
        resources.ApplyResources(Me.chkShowActivityLinks, "chkShowActivityLinks")
        Me.chkShowActivityLinks.Name = "chkShowActivityLinks"
        Me.chkShowActivityLinks.UseVisualStyleBackColor = True
        '
        'cmbSelectPeriodView
        '
        resources.ApplyResources(Me.cmbSelectPeriodView, "cmbSelectPeriodView")
        Me.cmbSelectPeriodView.FormattingEnabled = True
        Me.cmbSelectPeriodView.Name = "cmbSelectPeriodView"
        '
        'cmbShowElements
        '
        resources.ApplyResources(Me.cmbShowElements, "cmbShowElements")
        Me.cmbShowElements.FormattingEnabled = True
        Me.cmbShowElements.Name = "cmbShowElements"
        '
        'GroupBox1
        '
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Controls.Add(Me.dtpPeriodUntil)
        Me.GroupBox1.Controls.Add(Me.dtpPeriodFrom)
        Me.GroupBox1.Controls.Add(Me.lblUntil)
        Me.GroupBox1.Controls.Add(Me.lblFrom)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'dtpPeriodUntil
        '
        resources.ApplyResources(Me.dtpPeriodUntil, "dtpPeriodUntil")
        Me.dtpPeriodUntil.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpPeriodUntil.Name = "dtpPeriodUntil"
        '
        'dtpPeriodFrom
        '
        resources.ApplyResources(Me.dtpPeriodFrom, "dtpPeriodFrom")
        Me.dtpPeriodFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpPeriodFrom.Name = "dtpPeriodFrom"
        '
        'lblUntil
        '
        resources.ApplyResources(Me.lblUntil, "lblUntil")
        Me.lblUntil.Name = "lblUntil"
        '
        'lblFrom
        '
        resources.ApplyResources(Me.lblFrom, "lblFrom")
        Me.lblFrom.Name = "lblFrom"
        '
        'PrintSettingsPlanning
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBoxView)
        Me.Controls.Add(Me.GroupBoxOutput)
        Me.Name = "PrintSettingsPlanning"
        Me.GroupBoxOutput.ResumeLayout(False)
        Me.GroupBoxView.ResumeLayout(False)
        Me.GroupBoxView.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxOutput As System.Windows.Forms.GroupBox
    Friend WithEvents cmbSelectOutput As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBoxView As System.Windows.Forms.GroupBox
    Friend WithEvents cmbShowElements As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSelectPeriodView As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblFrom As System.Windows.Forms.Label
    Friend WithEvents lblUntil As System.Windows.Forms.Label
    Friend WithEvents dtpPeriodUntil As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpPeriodFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkShowActivityLinks As System.Windows.Forms.CheckBox
    Friend WithEvents chkRepeatRowHeaders As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowDatesColumns As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowKeyMomentLinks As System.Windows.Forms.CheckBox

End Class
