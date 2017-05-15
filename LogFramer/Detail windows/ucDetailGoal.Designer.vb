<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailGoal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailGoal))
        Me.TabControlProjectInfo = New System.Windows.Forms.TabControl()
        Me.TabPageProjectInfo = New System.Windows.Forms.TabPage()
        Me.tbCode = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblCode = New System.Windows.Forms.Label()
        Me.tbShortTitle = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblShortTitle = New System.Windows.Forms.Label()
        Me.ntbDuration = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.dtpEndDate = New FaciliDev.LogFramer.DateTimePickerLF()
        Me.dtpStartDate = New FaciliDev.LogFramer.DateTimePickerLF()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.cmbDurationUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblDuration = New System.Windows.Forms.Label()
        Me.tbProjectTitle = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblProjectTitle = New System.Windows.Forms.Label()
        Me.TabPagePartners = New System.Windows.Forms.TabPage()
        Me.lvPartners = New FaciliDev.LogFramer.ListViewPartners()
        Me.TabControlProjectInfo.SuspendLayout()
        Me.TabPageProjectInfo.SuspendLayout()
        Me.TabPagePartners.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlProjectInfo
        '
        resources.ApplyResources(Me.TabControlProjectInfo, "TabControlProjectInfo")
        Me.TabControlProjectInfo.Controls.Add(Me.TabPageProjectInfo)
        Me.TabControlProjectInfo.Controls.Add(Me.TabPagePartners)
        Me.TabControlProjectInfo.Name = "TabControlProjectInfo"
        Me.TabControlProjectInfo.SelectedIndex = 0
        '
        'TabPageProjectInfo
        '
        resources.ApplyResources(Me.TabPageProjectInfo, "TabPageProjectInfo")
        Me.TabPageProjectInfo.Controls.Add(Me.tbCode)
        Me.TabPageProjectInfo.Controls.Add(Me.lblCode)
        Me.TabPageProjectInfo.Controls.Add(Me.tbShortTitle)
        Me.TabPageProjectInfo.Controls.Add(Me.lblShortTitle)
        Me.TabPageProjectInfo.Controls.Add(Me.ntbDuration)
        Me.TabPageProjectInfo.Controls.Add(Me.dtpEndDate)
        Me.TabPageProjectInfo.Controls.Add(Me.dtpStartDate)
        Me.TabPageProjectInfo.Controls.Add(Me.lblEndDate)
        Me.TabPageProjectInfo.Controls.Add(Me.lblStartDate)
        Me.TabPageProjectInfo.Controls.Add(Me.cmbDurationUnit)
        Me.TabPageProjectInfo.Controls.Add(Me.lblDuration)
        Me.TabPageProjectInfo.Controls.Add(Me.tbProjectTitle)
        Me.TabPageProjectInfo.Controls.Add(Me.lblProjectTitle)
        Me.TabPageProjectInfo.Name = "TabPageProjectInfo"
        Me.TabPageProjectInfo.UseVisualStyleBackColor = True
        '
        'tbCode
        '
        resources.ApplyResources(Me.tbCode, "tbCode")
        Me.tbCode.Name = "tbCode"
        '
        'lblCode
        '
        resources.ApplyResources(Me.lblCode, "lblCode")
        Me.lblCode.Name = "lblCode"
        '
        'tbShortTitle
        '
        resources.ApplyResources(Me.tbShortTitle, "tbShortTitle")
        Me.tbShortTitle.Name = "tbShortTitle"
        '
        'lblShortTitle
        '
        resources.ApplyResources(Me.lblShortTitle, "lblShortTitle")
        Me.lblShortTitle.Name = "lblShortTitle"
        '
        'ntbDuration
        '
        resources.ApplyResources(Me.ntbDuration, "ntbDuration")
        Me.ntbDuration.AllowSpace = True
        Me.ntbDuration.IsCurrency = False
        Me.ntbDuration.IsPercentage = False
        Me.ntbDuration.Name = "ntbDuration"
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
        'cmbDurationUnit
        '
        resources.ApplyResources(Me.cmbDurationUnit, "cmbDurationUnit")
        Me.cmbDurationUnit.FormattingEnabled = True
        Me.cmbDurationUnit.Name = "cmbDurationUnit"
        '
        'lblDuration
        '
        resources.ApplyResources(Me.lblDuration, "lblDuration")
        Me.lblDuration.Name = "lblDuration"
        '
        'tbProjectTitle
        '
        resources.ApplyResources(Me.tbProjectTitle, "tbProjectTitle")
        Me.tbProjectTitle.Name = "tbProjectTitle"
        '
        'lblProjectTitle
        '
        resources.ApplyResources(Me.lblProjectTitle, "lblProjectTitle")
        Me.lblProjectTitle.Name = "lblProjectTitle"
        '
        'TabPagePartners
        '
        resources.ApplyResources(Me.TabPagePartners, "TabPagePartners")
        Me.TabPagePartners.Controls.Add(Me.lvPartners)
        Me.TabPagePartners.Name = "TabPagePartners"
        Me.TabPagePartners.UseVisualStyleBackColor = True
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
        'DetailGoal
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.TabControlProjectInfo)
        Me.Name = "DetailGoal"
        Me.TabControlProjectInfo.ResumeLayout(False)
        Me.TabPageProjectInfo.ResumeLayout(False)
        Me.TabPageProjectInfo.PerformLayout()
        Me.TabPagePartners.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControlProjectInfo As System.Windows.Forms.TabControl
    Friend WithEvents TabPageProjectInfo As System.Windows.Forms.TabPage
    Friend WithEvents TabPagePartners As System.Windows.Forms.TabPage
    Friend WithEvents tbProjectTitle As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblProjectTitle As System.Windows.Forms.Label
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents cmbDurationUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents dtpEndDate As FaciliDev.LogFramer.DateTimePickerLF
    Friend WithEvents dtpStartDate As FaciliDev.LogFramer.DateTimePickerLF
    Friend WithEvents lvPartners As FaciliDev.LogFramer.ListViewPartners
    Friend WithEvents ntbDuration As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents tbCode As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblCode As System.Windows.Forms.Label
    Friend WithEvents tbShortTitle As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblShortTitle As System.Windows.Forms.Label

End Class
