<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailProjectInfo
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailProjectInfo))
        Me.SplitContainerTargetGroups = New System.Windows.Forms.SplitContainer()
        Me.lvTargetGroups = New FaciliDev.LogFramer.ListViewTargetGroups()
        Me.lvTargetGroupStatistics = New FaciliDev.LogFramer.ListViewTargetGroupStatistics()
        Me.SplitContainerPartners = New System.Windows.Forms.SplitContainer()
        Me.lvPartners = New FaciliDev.LogFramer.ListViewPartners()
        Me.lvContacts = New FaciliDev.LogFramer.ListViewContacts()
        Me.lblCode = New System.Windows.Forms.Label()
        Me.lblShortTitle = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.lblDuration = New System.Windows.Forms.Label()
        Me.lblProjectTitle = New System.Windows.Forms.Label()
        Me.TabControlProjectInfo = New System.Windows.Forms.TabControl()
        Me.TabPageTargetGroups = New System.Windows.Forms.TabPage()
        Me.TabPagePartners = New System.Windows.Forms.TabPage()
        Me.TabPageDonors = New System.Windows.Forms.TabPage()
        Me.TabPageTargetDeadlines = New System.Windows.Forms.TabPage()
        Me.PanelTargetDeadlines = New System.Windows.Forms.Panel()
        Me.PanelRiskMonitoring = New System.Windows.Forms.Panel()
        Me.gbRiskMonitoring = New System.Windows.Forms.GroupBox()
        Me.ucRiskMonitoring = New FaciliDev.LogFramer.ucRiskMonitoringDeadlines()
        Me.gbTargetDeadlinesActivities = New System.Windows.Forms.GroupBox()
        Me.ucTargetDeadlinesActivities = New FaciliDev.LogFramer.ucTargetDeadlines()
        Me.gbTargetDeadlinesOutputs = New System.Windows.Forms.GroupBox()
        Me.ucTargetDeadlinesOutputs = New FaciliDev.LogFramer.ucTargetDeadlines()
        Me.gbTargetDeadlinesPurposes = New System.Windows.Forms.GroupBox()
        Me.ucTargetDeadlinesPurposes = New FaciliDev.LogFramer.ucTargetDeadlines()
        Me.gbTargetDeadlinesGoals = New System.Windows.Forms.GroupBox()
        Me.ucTargetDeadlinesGoals = New FaciliDev.LogFramer.ucTargetDeadlines()
        Me.tbCode = New FaciliDev.LogFramer.TextBoxLF()
        Me.tbShortTitle = New FaciliDev.LogFramer.TextBoxLF()
        Me.dtpEndDate = New FaciliDev.LogFramer.DateTimePickerLF()
        Me.dtpStartDate = New FaciliDev.LogFramer.DateTimePickerLF()
        Me.cmbDurationUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.tbProjectTitle = New FaciliDev.LogFramer.TextBoxLF()
        Me.ntbDuration = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        CType(Me.SplitContainerTargetGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerTargetGroups.Panel1.SuspendLayout()
        Me.SplitContainerTargetGroups.Panel2.SuspendLayout()
        Me.SplitContainerTargetGroups.SuspendLayout()
        CType(Me.SplitContainerPartners, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerPartners.Panel1.SuspendLayout()
        Me.SplitContainerPartners.Panel2.SuspendLayout()
        Me.SplitContainerPartners.SuspendLayout()
        Me.TabControlProjectInfo.SuspendLayout()
        Me.TabPageTargetGroups.SuspendLayout()
        Me.TabPagePartners.SuspendLayout()
        Me.TabPageTargetDeadlines.SuspendLayout()
        Me.PanelTargetDeadlines.SuspendLayout()
        Me.PanelRiskMonitoring.SuspendLayout()
        Me.gbRiskMonitoring.SuspendLayout()
        Me.gbTargetDeadlinesActivities.SuspendLayout()
        Me.gbTargetDeadlinesOutputs.SuspendLayout()
        Me.gbTargetDeadlinesPurposes.SuspendLayout()
        Me.gbTargetDeadlinesGoals.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainerTargetGroups
        '
        resources.ApplyResources(Me.SplitContainerTargetGroups, "SplitContainerTargetGroups")
        Me.SplitContainerTargetGroups.Name = "SplitContainerTargetGroups"
        '
        'SplitContainerTargetGroups.Panel1
        '
        resources.ApplyResources(Me.SplitContainerTargetGroups.Panel1, "SplitContainerTargetGroups.Panel1")
        Me.SplitContainerTargetGroups.Panel1.Controls.Add(Me.lvTargetGroups)
        '
        'SplitContainerTargetGroups.Panel2
        '
        resources.ApplyResources(Me.SplitContainerTargetGroups.Panel2, "SplitContainerTargetGroups.Panel2")
        Me.SplitContainerTargetGroups.Panel2.Controls.Add(Me.lvTargetGroupStatistics)
        '
        'lvTargetGroups
        '
        resources.ApplyResources(Me.lvTargetGroups, "lvTargetGroups")
        Me.lvTargetGroups.AllTargetGroups = False
        Me.lvTargetGroups.FullRowSelect = True
        Me.lvTargetGroups.Name = "lvTargetGroups"
        Me.lvTargetGroups.SortColumnIndex = -1
        Me.lvTargetGroups.TargetGroups = Nothing
        Me.lvTargetGroups.UseCompatibleStateImageBehavior = False
        Me.lvTargetGroups.View = System.Windows.Forms.View.Details
        '
        'lvTargetGroupStatistics
        '
        resources.ApplyResources(Me.lvTargetGroupStatistics, "lvTargetGroupStatistics")
        Me.lvTargetGroupStatistics.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.lvTargetGroupStatistics.ForeColor = System.Drawing.Color.White
        Me.lvTargetGroupStatistics.FullRowSelect = True
        Me.lvTargetGroupStatistics.Name = "lvTargetGroupStatistics"
        Me.lvTargetGroupStatistics.TargetGroups = Nothing
        Me.lvTargetGroupStatistics.UseCompatibleStateImageBehavior = False
        Me.lvTargetGroupStatistics.View = System.Windows.Forms.View.Details
        '
        'SplitContainerPartners
        '
        resources.ApplyResources(Me.SplitContainerPartners, "SplitContainerPartners")
        Me.SplitContainerPartners.Name = "SplitContainerPartners"
        '
        'SplitContainerPartners.Panel1
        '
        resources.ApplyResources(Me.SplitContainerPartners.Panel1, "SplitContainerPartners.Panel1")
        Me.SplitContainerPartners.Panel1.Controls.Add(Me.lvPartners)
        '
        'SplitContainerPartners.Panel2
        '
        resources.ApplyResources(Me.SplitContainerPartners.Panel2, "SplitContainerPartners.Panel2")
        Me.SplitContainerPartners.Panel2.Controls.Add(Me.lvContacts)
        '
        'lvPartners
        '
        resources.ApplyResources(Me.lvPartners, "lvPartners")
        Me.lvPartners.FullRowSelect = True
        Me.lvPartners.Name = "lvPartners"
        Me.lvPartners.ProjectPartners = Nothing
        Me.lvPartners.SortColumnIndex = -1
        Me.lvPartners.UseCompatibleStateImageBehavior = False
        Me.lvPartners.View = System.Windows.Forms.View.Details
        '
        'lvContacts
        '
        resources.ApplyResources(Me.lvContacts, "lvContacts")
        Me.lvContacts.Contacts = Nothing
        Me.lvContacts.FullRowSelect = True
        Me.lvContacts.Name = "lvContacts"
        Me.lvContacts.SortColumnIndex = -1
        Me.lvContacts.UseCompatibleStateImageBehavior = False
        Me.lvContacts.View = System.Windows.Forms.View.Details
        '
        'lblCode
        '
        resources.ApplyResources(Me.lblCode, "lblCode")
        Me.lblCode.Name = "lblCode"
        '
        'lblShortTitle
        '
        resources.ApplyResources(Me.lblShortTitle, "lblShortTitle")
        Me.lblShortTitle.Name = "lblShortTitle"
        '
        'lblEndDate
        '
        resources.ApplyResources(Me.lblEndDate, "lblEndDate")
        Me.lblEndDate.Name = "lblEndDate"
        '
        'lblStartDate
        '
        resources.ApplyResources(Me.lblStartDate, "lblStartDate")
        Me.lblStartDate.Name = "lblStartDate"
        '
        'lblDuration
        '
        resources.ApplyResources(Me.lblDuration, "lblDuration")
        Me.lblDuration.Name = "lblDuration"
        '
        'lblProjectTitle
        '
        resources.ApplyResources(Me.lblProjectTitle, "lblProjectTitle")
        Me.lblProjectTitle.Name = "lblProjectTitle"
        '
        'TabControlProjectInfo
        '
        resources.ApplyResources(Me.TabControlProjectInfo, "TabControlProjectInfo")
        Me.TabControlProjectInfo.Controls.Add(Me.TabPageTargetGroups)
        Me.TabControlProjectInfo.Controls.Add(Me.TabPagePartners)
        Me.TabControlProjectInfo.Controls.Add(Me.TabPageDonors)
        Me.TabControlProjectInfo.Controls.Add(Me.TabPageTargetDeadlines)
        Me.TabControlProjectInfo.Multiline = True
        Me.TabControlProjectInfo.Name = "TabControlProjectInfo"
        Me.TabControlProjectInfo.SelectedIndex = 0
        '
        'TabPageTargetGroups
        '
        resources.ApplyResources(Me.TabPageTargetGroups, "TabPageTargetGroups")
        Me.TabPageTargetGroups.Controls.Add(Me.SplitContainerTargetGroups)
        Me.TabPageTargetGroups.Name = "TabPageTargetGroups"
        Me.TabPageTargetGroups.UseVisualStyleBackColor = True
        '
        'TabPagePartners
        '
        resources.ApplyResources(Me.TabPagePartners, "TabPagePartners")
        Me.TabPagePartners.Controls.Add(Me.SplitContainerPartners)
        Me.TabPagePartners.Name = "TabPagePartners"
        Me.TabPagePartners.UseVisualStyleBackColor = True
        '
        'TabPageDonors
        '
        resources.ApplyResources(Me.TabPageDonors, "TabPageDonors")
        Me.TabPageDonors.Name = "TabPageDonors"
        Me.TabPageDonors.UseVisualStyleBackColor = True
        '
        'TabPageTargetDeadlines
        '
        resources.ApplyResources(Me.TabPageTargetDeadlines, "TabPageTargetDeadlines")
        Me.TabPageTargetDeadlines.Controls.Add(Me.PanelTargetDeadlines)
        Me.TabPageTargetDeadlines.Name = "TabPageTargetDeadlines"
        Me.TabPageTargetDeadlines.UseVisualStyleBackColor = True
        '
        'PanelTargetDeadlines
        '
        resources.ApplyResources(Me.PanelTargetDeadlines, "PanelTargetDeadlines")
        Me.PanelTargetDeadlines.Controls.Add(Me.PanelRiskMonitoring)
        Me.PanelTargetDeadlines.Controls.Add(Me.gbTargetDeadlinesActivities)
        Me.PanelTargetDeadlines.Controls.Add(Me.gbTargetDeadlinesOutputs)
        Me.PanelTargetDeadlines.Controls.Add(Me.gbTargetDeadlinesPurposes)
        Me.PanelTargetDeadlines.Controls.Add(Me.gbTargetDeadlinesGoals)
        Me.PanelTargetDeadlines.Name = "PanelTargetDeadlines"
        '
        'PanelRiskMonitoring
        '
        resources.ApplyResources(Me.PanelRiskMonitoring, "PanelRiskMonitoring")
        Me.PanelRiskMonitoring.BackColor = System.Drawing.SystemColors.ControlDark
        Me.PanelRiskMonitoring.Controls.Add(Me.gbRiskMonitoring)
        Me.PanelRiskMonitoring.Name = "PanelRiskMonitoring"
        '
        'gbRiskMonitoring
        '
        resources.ApplyResources(Me.gbRiskMonitoring, "gbRiskMonitoring")
        Me.gbRiskMonitoring.BackColor = System.Drawing.Color.Transparent
        Me.gbRiskMonitoring.Controls.Add(Me.ucRiskMonitoring)
        Me.gbRiskMonitoring.Name = "gbRiskMonitoring"
        Me.gbRiskMonitoring.TabStop = False
        '
        'ucRiskMonitoring
        '
        resources.ApplyResources(Me.ucRiskMonitoring, "ucRiskMonitoring")
        Me.ucRiskMonitoring.BackColor = System.Drawing.SystemColors.ControlDark
        Me.ucRiskMonitoring.Name = "ucRiskMonitoring"
        Me.ucRiskMonitoring.RiskMonitoring = Nothing
        '
        'gbTargetDeadlinesActivities
        '
        resources.ApplyResources(Me.gbTargetDeadlinesActivities, "gbTargetDeadlinesActivities")
        Me.gbTargetDeadlinesActivities.Controls.Add(Me.ucTargetDeadlinesActivities)
        Me.gbTargetDeadlinesActivities.Name = "gbTargetDeadlinesActivities"
        Me.gbTargetDeadlinesActivities.TabStop = False
        '
        'ucTargetDeadlinesActivities
        '
        resources.ApplyResources(Me.ucTargetDeadlinesActivities, "ucTargetDeadlinesActivities")
        Me.ucTargetDeadlinesActivities.Name = "ucTargetDeadlinesActivities"
        Me.ucTargetDeadlinesActivities.TargetDeadlinesSection = Nothing
        '
        'gbTargetDeadlinesOutputs
        '
        resources.ApplyResources(Me.gbTargetDeadlinesOutputs, "gbTargetDeadlinesOutputs")
        Me.gbTargetDeadlinesOutputs.Controls.Add(Me.ucTargetDeadlinesOutputs)
        Me.gbTargetDeadlinesOutputs.Name = "gbTargetDeadlinesOutputs"
        Me.gbTargetDeadlinesOutputs.TabStop = False
        '
        'ucTargetDeadlinesOutputs
        '
        resources.ApplyResources(Me.ucTargetDeadlinesOutputs, "ucTargetDeadlinesOutputs")
        Me.ucTargetDeadlinesOutputs.Name = "ucTargetDeadlinesOutputs"
        Me.ucTargetDeadlinesOutputs.TargetDeadlinesSection = Nothing
        '
        'gbTargetDeadlinesPurposes
        '
        resources.ApplyResources(Me.gbTargetDeadlinesPurposes, "gbTargetDeadlinesPurposes")
        Me.gbTargetDeadlinesPurposes.Controls.Add(Me.ucTargetDeadlinesPurposes)
        Me.gbTargetDeadlinesPurposes.Name = "gbTargetDeadlinesPurposes"
        Me.gbTargetDeadlinesPurposes.TabStop = False
        '
        'ucTargetDeadlinesPurposes
        '
        resources.ApplyResources(Me.ucTargetDeadlinesPurposes, "ucTargetDeadlinesPurposes")
        Me.ucTargetDeadlinesPurposes.Name = "ucTargetDeadlinesPurposes"
        Me.ucTargetDeadlinesPurposes.TargetDeadlinesSection = Nothing
        '
        'gbTargetDeadlinesGoals
        '
        resources.ApplyResources(Me.gbTargetDeadlinesGoals, "gbTargetDeadlinesGoals")
        Me.gbTargetDeadlinesGoals.Controls.Add(Me.ucTargetDeadlinesGoals)
        Me.gbTargetDeadlinesGoals.Name = "gbTargetDeadlinesGoals"
        Me.gbTargetDeadlinesGoals.TabStop = False
        '
        'ucTargetDeadlinesGoals
        '
        resources.ApplyResources(Me.ucTargetDeadlinesGoals, "ucTargetDeadlinesGoals")
        Me.ucTargetDeadlinesGoals.Name = "ucTargetDeadlinesGoals"
        Me.ucTargetDeadlinesGoals.TargetDeadlinesSection = Nothing
        '
        'tbCode
        '
        resources.ApplyResources(Me.tbCode, "tbCode")
        Me.tbCode.Name = "tbCode"
        '
        'tbShortTitle
        '
        resources.ApplyResources(Me.tbShortTitle, "tbShortTitle")
        Me.tbShortTitle.Name = "tbShortTitle"
        '
        'dtpEndDate
        '
        resources.ApplyResources(Me.dtpEndDate, "dtpEndDate")
        Me.dtpEndDate.Name = "dtpEndDate"
        '
        'dtpStartDate
        '
        resources.ApplyResources(Me.dtpStartDate, "dtpStartDate")
        Me.dtpStartDate.Name = "dtpStartDate"
        '
        'cmbDurationUnit
        '
        resources.ApplyResources(Me.cmbDurationUnit, "cmbDurationUnit")
        Me.cmbDurationUnit.FormattingEnabled = True
        Me.cmbDurationUnit.Name = "cmbDurationUnit"
        '
        'tbProjectTitle
        '
        resources.ApplyResources(Me.tbProjectTitle, "tbProjectTitle")
        Me.tbProjectTitle.Name = "tbProjectTitle"
        '
        'ntbDuration
        '
        resources.ApplyResources(Me.ntbDuration, "ntbDuration")
        Me.ntbDuration.Name = "ntbDuration"
        '
        'DetailProjectInfo
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ntbDuration)
        Me.Controls.Add(Me.TabControlProjectInfo)
        Me.Controls.Add(Me.tbCode)
        Me.Controls.Add(Me.lblCode)
        Me.Controls.Add(Me.tbShortTitle)
        Me.Controls.Add(Me.lblShortTitle)
        Me.Controls.Add(Me.dtpEndDate)
        Me.Controls.Add(Me.dtpStartDate)
        Me.Controls.Add(Me.lblEndDate)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.cmbDurationUnit)
        Me.Controls.Add(Me.lblDuration)
        Me.Controls.Add(Me.tbProjectTitle)
        Me.Controls.Add(Me.lblProjectTitle)
        Me.Name = "DetailProjectInfo"
        Me.SplitContainerTargetGroups.Panel1.ResumeLayout(False)
        Me.SplitContainerTargetGroups.Panel2.ResumeLayout(False)
        CType(Me.SplitContainerTargetGroups, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerTargetGroups.ResumeLayout(False)
        Me.SplitContainerPartners.Panel1.ResumeLayout(False)
        Me.SplitContainerPartners.Panel2.ResumeLayout(False)
        CType(Me.SplitContainerPartners, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerPartners.ResumeLayout(False)
        Me.TabControlProjectInfo.ResumeLayout(False)
        Me.TabPageTargetGroups.ResumeLayout(False)
        Me.TabPagePartners.ResumeLayout(False)
        Me.TabPageTargetDeadlines.ResumeLayout(False)
        Me.PanelTargetDeadlines.ResumeLayout(False)
        Me.PanelRiskMonitoring.ResumeLayout(False)
        Me.gbRiskMonitoring.ResumeLayout(False)
        Me.gbTargetDeadlinesActivities.ResumeLayout(False)
        Me.gbTargetDeadlinesOutputs.ResumeLayout(False)
        Me.gbTargetDeadlinesPurposes.ResumeLayout(False)
        Me.gbTargetDeadlinesGoals.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbCode As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents tbShortTitle As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblShortTitle As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As FaciliDev.LogFramer.DateTimePickerLF
    Friend WithEvents dtpStartDate As FaciliDev.LogFramer.DateTimePickerLF
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents cmbDurationUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents tbProjectTitle As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblProjectTitle As System.Windows.Forms.Label
    Friend WithEvents TabControlProjectInfo As System.Windows.Forms.TabControl
    Friend WithEvents TabPageTargetGroups As System.Windows.Forms.TabPage
    Friend WithEvents TabPagePartners As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainerTargetGroups As System.Windows.Forms.SplitContainer
    Friend WithEvents lvTargetGroups As FaciliDev.LogFramer.ListViewTargetGroups
    Friend WithEvents lvTargetGroupStatistics As FaciliDev.LogFramer.ListViewTargetGroupStatistics
    Friend WithEvents TabPageDonors As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainerPartners As System.Windows.Forms.SplitContainer
    Friend WithEvents lvPartners As FaciliDev.LogFramer.ListViewPartners
    Friend WithEvents lvContacts As FaciliDev.LogFramer.ListViewContacts
    Friend WithEvents TabPageTargetDeadlines As System.Windows.Forms.TabPage
    Friend WithEvents PanelTargetDeadlines As System.Windows.Forms.Panel
    Friend WithEvents gbTargetDeadlinesActivities As System.Windows.Forms.GroupBox
    Friend WithEvents ucTargetDeadlinesActivities As FaciliDev.LogFramer.ucTargetDeadlines
    Friend WithEvents gbTargetDeadlinesOutputs As System.Windows.Forms.GroupBox
    Friend WithEvents ucTargetDeadlinesOutputs As FaciliDev.LogFramer.ucTargetDeadlines
    Friend WithEvents gbTargetDeadlinesPurposes As System.Windows.Forms.GroupBox
    Friend WithEvents ucTargetDeadlinesPurposes As FaciliDev.LogFramer.ucTargetDeadlines
    Friend WithEvents gbTargetDeadlinesGoals As System.Windows.Forms.GroupBox
    Friend WithEvents ucTargetDeadlinesGoals As FaciliDev.LogFramer.ucTargetDeadlines
    Friend WithEvents PanelRiskMonitoring As System.Windows.Forms.Panel
    Friend WithEvents gbRiskMonitoring As System.Windows.Forms.GroupBox
    Friend WithEvents ucRiskMonitoring As FaciliDev.LogFramer.ucRiskMonitoringDeadlines
    Friend WithEvents ntbDuration As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
