<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucRiskMonitoringDeadlines
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucRiskMonitoringDeadlines))
        Me.lblAfterProjectStarts = New System.Windows.Forms.Label()
        Me.lblStarting = New System.Windows.Forms.Label()
        Me.dtpStartDate = New System.Windows.Forms.DateTimePicker()
        Me.rbtnReferenceDateStart = New System.Windows.Forms.RadioButton()
        Me.rbtnExactDateStart = New System.Windows.Forms.RadioButton()
        Me.lblTargetSystem = New System.Windows.Forms.Label()
        Me.cmbRepetition = New System.Windows.Forms.ComboBox()
        Me.lblAfterProjectEnds = New System.Windows.Forms.Label()
        Me.lblEnding = New System.Windows.Forms.Label()
        Me.dtpEndDate = New System.Windows.Forms.DateTimePicker()
        Me.rbtnReferenceDateEnd = New System.Windows.Forms.RadioButton()
        Me.rbtnExactDateEnd = New System.Windows.Forms.RadioButton()
        Me.gbStarting = New System.Windows.Forms.GroupBox()
        Me.cmbPeriodUnitStart = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.ntbPeriodStart = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.gbEnding = New System.Windows.Forms.GroupBox()
        Me.cmbPeriodUnitEnd = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.ntbPeriodEnd = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.lvRiskMonitoringDeadlines = New FaciliDev.LogFramer.ListViewRiskMonitoringDeadlines()
        Me.gbStarting.SuspendLayout()
        Me.gbEnding.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblAfterProjectStarts
        '
        resources.ApplyResources(Me.lblAfterProjectStarts, "lblAfterProjectStarts")
        Me.lblAfterProjectStarts.Name = "lblAfterProjectStarts"
        '
        'lblStarting
        '
        resources.ApplyResources(Me.lblStarting, "lblStarting")
        Me.lblStarting.Name = "lblStarting"
        '
        'dtpStartDate
        '
        resources.ApplyResources(Me.dtpStartDate, "dtpStartDate")
        Me.dtpStartDate.Name = "dtpStartDate"
        '
        'rbtnReferenceDateStart
        '
        resources.ApplyResources(Me.rbtnReferenceDateStart, "rbtnReferenceDateStart")
        Me.rbtnReferenceDateStart.Checked = True
        Me.rbtnReferenceDateStart.Name = "rbtnReferenceDateStart"
        Me.rbtnReferenceDateStart.TabStop = True
        Me.rbtnReferenceDateStart.UseVisualStyleBackColor = True
        '
        'rbtnExactDateStart
        '
        resources.ApplyResources(Me.rbtnExactDateStart, "rbtnExactDateStart")
        Me.rbtnExactDateStart.Checked = True
        Me.rbtnExactDateStart.Name = "rbtnExactDateStart"
        Me.rbtnExactDateStart.TabStop = True
        Me.rbtnExactDateStart.UseVisualStyleBackColor = True
        '
        'lblTargetSystem
        '
        resources.ApplyResources(Me.lblTargetSystem, "lblTargetSystem")
        Me.lblTargetSystem.Name = "lblTargetSystem"
        '
        'cmbRepetition
        '
        resources.ApplyResources(Me.cmbRepetition, "cmbRepetition")
        Me.cmbRepetition.Name = "cmbRepetition"
        '
        'lblAfterProjectEnds
        '
        resources.ApplyResources(Me.lblAfterProjectEnds, "lblAfterProjectEnds")
        Me.lblAfterProjectEnds.Name = "lblAfterProjectEnds"
        '
        'lblEnding
        '
        resources.ApplyResources(Me.lblEnding, "lblEnding")
        Me.lblEnding.Name = "lblEnding"
        '
        'dtpEndDate
        '
        resources.ApplyResources(Me.dtpEndDate, "dtpEndDate")
        Me.dtpEndDate.Name = "dtpEndDate"
        '
        'rbtnReferenceDateEnd
        '
        resources.ApplyResources(Me.rbtnReferenceDateEnd, "rbtnReferenceDateEnd")
        Me.rbtnReferenceDateEnd.Checked = True
        Me.rbtnReferenceDateEnd.Name = "rbtnReferenceDateEnd"
        Me.rbtnReferenceDateEnd.TabStop = True
        Me.rbtnReferenceDateEnd.UseVisualStyleBackColor = True
        '
        'rbtnExactDateEnd
        '
        resources.ApplyResources(Me.rbtnExactDateEnd, "rbtnExactDateEnd")
        Me.rbtnExactDateEnd.Checked = True
        Me.rbtnExactDateEnd.Name = "rbtnExactDateEnd"
        Me.rbtnExactDateEnd.TabStop = True
        Me.rbtnExactDateEnd.UseVisualStyleBackColor = True
        '
        'gbStarting
        '
        resources.ApplyResources(Me.gbStarting, "gbStarting")
        Me.gbStarting.Controls.Add(Me.lblStarting)
        Me.gbStarting.Controls.Add(Me.cmbPeriodUnitStart)
        Me.gbStarting.Controls.Add(Me.rbtnExactDateStart)
        Me.gbStarting.Controls.Add(Me.rbtnReferenceDateStart)
        Me.gbStarting.Controls.Add(Me.ntbPeriodStart)
        Me.gbStarting.Controls.Add(Me.dtpStartDate)
        Me.gbStarting.Controls.Add(Me.lblAfterProjectStarts)
        Me.gbStarting.Name = "gbStarting"
        Me.gbStarting.TabStop = False
        '
        'cmbPeriodUnitStart
        '
        resources.ApplyResources(Me.cmbPeriodUnitStart, "cmbPeriodUnitStart")
        Me.cmbPeriodUnitStart.FormattingEnabled = True
        Me.cmbPeriodUnitStart.Name = "cmbPeriodUnitStart"
        '
        'ntbPeriodStart
        '
        resources.ApplyResources(Me.ntbPeriodStart, "ntbPeriodStart")
        Me.ntbPeriodStart.AllowSpace = True
        Me.ntbPeriodStart.IsCurrency = False
        Me.ntbPeriodStart.IsPercentage = False
        Me.ntbPeriodStart.Name = "ntbPeriodStart"
        '
        'gbEnding
        '
        resources.ApplyResources(Me.gbEnding, "gbEnding")
        Me.gbEnding.Controls.Add(Me.cmbPeriodUnitEnd)
        Me.gbEnding.Controls.Add(Me.rbtnExactDateEnd)
        Me.gbEnding.Controls.Add(Me.lblAfterProjectEnds)
        Me.gbEnding.Controls.Add(Me.rbtnReferenceDateEnd)
        Me.gbEnding.Controls.Add(Me.lblEnding)
        Me.gbEnding.Controls.Add(Me.ntbPeriodEnd)
        Me.gbEnding.Controls.Add(Me.dtpEndDate)
        Me.gbEnding.Name = "gbEnding"
        Me.gbEnding.TabStop = False
        '
        'cmbPeriodUnitEnd
        '
        resources.ApplyResources(Me.cmbPeriodUnitEnd, "cmbPeriodUnitEnd")
        Me.cmbPeriodUnitEnd.FormattingEnabled = True
        Me.cmbPeriodUnitEnd.Name = "cmbPeriodUnitEnd"
        '
        'ntbPeriodEnd
        '
        resources.ApplyResources(Me.ntbPeriodEnd, "ntbPeriodEnd")
        Me.ntbPeriodEnd.AllowSpace = True
        Me.ntbPeriodEnd.IsCurrency = False
        Me.ntbPeriodEnd.IsPercentage = False
        Me.ntbPeriodEnd.Name = "ntbPeriodEnd"
        '
        'lvRiskMonitoringDeadlines
        '
        resources.ApplyResources(Me.lvRiskMonitoringDeadlines, "lvRiskMonitoringDeadlines")
        Me.lvRiskMonitoringDeadlines.FullRowSelect = True
        Me.lvRiskMonitoringDeadlines.Name = "lvRiskMonitoringDeadlines"
        Me.lvRiskMonitoringDeadlines.RiskMonitoring = Nothing
        Me.lvRiskMonitoringDeadlines.SortColumnIndex = -1
        Me.lvRiskMonitoringDeadlines.UseCompatibleStateImageBehavior = False
        Me.lvRiskMonitoringDeadlines.View = System.Windows.Forms.View.Details
        '
        'ucRiskMonitoringDeadlines
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Controls.Add(Me.lvRiskMonitoringDeadlines)
        Me.Controls.Add(Me.gbEnding)
        Me.Controls.Add(Me.gbStarting)
        Me.Controls.Add(Me.lblTargetSystem)
        Me.Controls.Add(Me.cmbRepetition)
        Me.Name = "ucRiskMonitoringDeadlines"
        Me.gbStarting.ResumeLayout(False)
        Me.gbStarting.PerformLayout()
        Me.gbEnding.ResumeLayout(False)
        Me.gbEnding.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblAfterProjectStarts As System.Windows.Forms.Label
    Friend WithEvents lblStarting As System.Windows.Forms.Label
    Friend WithEvents dtpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbPeriodUnitStart As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents ntbPeriodStart As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents rbtnReferenceDateStart As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnExactDateStart As System.Windows.Forms.RadioButton
    Friend WithEvents lblTargetSystem As System.Windows.Forms.Label
    Friend WithEvents cmbRepetition As System.Windows.Forms.ComboBox
    Friend WithEvents lblAfterProjectEnds As System.Windows.Forms.Label
    Friend WithEvents lblEnding As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbPeriodUnitEnd As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents ntbPeriodEnd As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents rbtnReferenceDateEnd As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnExactDateEnd As System.Windows.Forms.RadioButton
    Friend WithEvents gbStarting As System.Windows.Forms.GroupBox
    Friend WithEvents gbEnding As System.Windows.Forms.GroupBox
    Friend WithEvents lvRiskMonitoringDeadlines As FaciliDev.LogFramer.ListViewRiskMonitoringDeadlines

End Class
