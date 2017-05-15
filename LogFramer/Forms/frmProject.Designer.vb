<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProject
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProject))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.SplitContainerUtilities = New System.Windows.Forms.SplitContainer()
        Me.Clipboard = New FaciliDev.LogFramer.DetailClipboard()
        Me.SplitContainerMain = New System.Windows.Forms.SplitContainer()
        Me.TabControlProject = New FaciliDev.LogFramer.ucTabControlProject()
        Me.TabPageProject = New System.Windows.Forms.TabPage()
        Me.ProjectInfo = New FaciliDev.LogFramer.DetailProjectInfo()
        Me.TabPageLogframe = New System.Windows.Forms.TabPage()
        Me.SplitContainerLogFrame = New System.Windows.Forms.SplitContainer()
        Me.dgvLogframe = New FaciliDev.LogFramer.DataGridViewLogframe()
        Me.TabPagePlanning = New System.Windows.Forms.TabPage()
        Me.SplitContainerPlanning = New System.Windows.Forms.SplitContainer()
        Me.dgvPlanning = New FaciliDev.LogFramer.DataGridViewPlanning()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RichTextColumnLogframe1 = New FaciliDev.LogFramer.RichTextColumnLogframe()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabPageBudget = New System.Windows.Forms.TabPage()
        Me.SplitContainerExchangeRates = New System.Windows.Forms.SplitContainer()
        Me.TabControlBudget = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPageMonitoring = New System.Windows.Forms.TabPage()
        Me.TabControlQuestionnaires = New System.Windows.Forms.TabControl()
        Me.TabPageProgrammeLevel = New System.Windows.Forms.TabPage()
        Me.TabPageTeamLevel = New System.Windows.Forms.TabPage()
        Me.TabPageActions = New System.Windows.Forms.TabPage()
        Me.ImageListTabs = New System.Windows.Forms.ImageList(Me.components)
        CType(Me.SplitContainerUtilities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerUtilities.Panel1.SuspendLayout()
        Me.SplitContainerUtilities.Panel2.SuspendLayout()
        Me.SplitContainerUtilities.SuspendLayout()
        CType(Me.SplitContainerMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerMain.Panel1.SuspendLayout()
        Me.SplitContainerMain.SuspendLayout()
        Me.TabControlProject.SuspendLayout()
        Me.TabPageProject.SuspendLayout()
        Me.TabPageLogframe.SuspendLayout()
        CType(Me.SplitContainerLogFrame, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerLogFrame.Panel1.SuspendLayout()
        Me.SplitContainerLogFrame.SuspendLayout()
        CType(Me.dgvLogframe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPagePlanning.SuspendLayout()
        CType(Me.SplitContainerPlanning, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerPlanning.Panel1.SuspendLayout()
        Me.SplitContainerPlanning.SuspendLayout()
        CType(Me.dgvPlanning, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageBudget.SuspendLayout()
        CType(Me.SplitContainerExchangeRates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerExchangeRates.Panel1.SuspendLayout()
        Me.SplitContainerExchangeRates.SuspendLayout()
        Me.TabControlBudget.SuspendLayout()
        Me.TabPageMonitoring.SuspendLayout()
        Me.TabControlQuestionnaires.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainerUtilities
        '
        resources.ApplyResources(Me.SplitContainerUtilities, "SplitContainerUtilities")
        Me.SplitContainerUtilities.Name = "SplitContainerUtilities"
        '
        'SplitContainerUtilities.Panel1
        '
        resources.ApplyResources(Me.SplitContainerUtilities.Panel1, "SplitContainerUtilities.Panel1")
        Me.SplitContainerUtilities.Panel1.Controls.Add(Me.Clipboard)
        '
        'SplitContainerUtilities.Panel2
        '
        resources.ApplyResources(Me.SplitContainerUtilities.Panel2, "SplitContainerUtilities.Panel2")
        Me.SplitContainerUtilities.Panel2.Controls.Add(Me.SplitContainerMain)
        '
        'Clipboard
        '
        resources.ApplyResources(Me.Clipboard, "Clipboard")
        Me.Clipboard.Name = "Clipboard"
        '
        'SplitContainerMain
        '
        resources.ApplyResources(Me.SplitContainerMain, "SplitContainerMain")
        Me.SplitContainerMain.Name = "SplitContainerMain"
        '
        'SplitContainerMain.Panel1
        '
        resources.ApplyResources(Me.SplitContainerMain.Panel1, "SplitContainerMain.Panel1")
        Me.SplitContainerMain.Panel1.Controls.Add(Me.TabControlProject)
        '
        'SplitContainerMain.Panel2
        '
        resources.ApplyResources(Me.SplitContainerMain.Panel2, "SplitContainerMain.Panel2")
        '
        'TabControlProject
        '
        resources.ApplyResources(Me.TabControlProject, "TabControlProject")
        Me.TabControlProject.Controls.Add(Me.TabPageProject)
        Me.TabControlProject.Controls.Add(Me.TabPageLogframe)
        Me.TabControlProject.Controls.Add(Me.TabPagePlanning)
        Me.TabControlProject.Controls.Add(Me.TabPageBudget)
        Me.TabControlProject.Controls.Add(Me.TabPageMonitoring)
        Me.TabControlProject.Controls.Add(Me.TabPageActions)
        Me.TabControlProject.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
        Me.TabControlProject.Name = "TabControlProject"
        Me.TabControlProject.SelectedIndex = 0
        Me.TabControlProject.TabSize = New System.Drawing.Size(200, 40)
        '
        'TabPageProject
        '
        resources.ApplyResources(Me.TabPageProject, "TabPageProject")
        Me.TabPageProject.Controls.Add(Me.ProjectInfo)
        Me.TabPageProject.Name = "TabPageProject"
        Me.TabPageProject.UseVisualStyleBackColor = True
        '
        'ProjectInfo
        '
        resources.ApplyResources(Me.ProjectInfo, "ProjectInfo")
        Me.ProjectInfo.Name = "ProjectInfo"
        '
        'TabPageLogframe
        '
        resources.ApplyResources(Me.TabPageLogframe, "TabPageLogframe")
        Me.TabPageLogframe.Controls.Add(Me.SplitContainerLogFrame)
        Me.TabPageLogframe.Name = "TabPageLogframe"
        Me.TabPageLogframe.UseVisualStyleBackColor = True
        '
        'SplitContainerLogFrame
        '
        resources.ApplyResources(Me.SplitContainerLogFrame, "SplitContainerLogFrame")
        Me.SplitContainerLogFrame.Name = "SplitContainerLogFrame"
        '
        'SplitContainerLogFrame.Panel1
        '
        resources.ApplyResources(Me.SplitContainerLogFrame.Panel1, "SplitContainerLogFrame.Panel1")
        Me.SplitContainerLogFrame.Panel1.Controls.Add(Me.dgvLogframe)
        '
        'SplitContainerLogFrame.Panel2
        '
        resources.ApplyResources(Me.SplitContainerLogFrame.Panel2, "SplitContainerLogFrame.Panel2")
        '
        'dgvLogframe
        '
        resources.ApplyResources(Me.dgvLogframe, "dgvLogframe")
        Me.dgvLogframe.AllowUserToAddRows = False
        Me.dgvLogframe.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvLogframe.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvLogframe.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Verdana", 9.75!)
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(2)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvLogframe.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvLogframe.DragBoxFromMouseDown = New System.Drawing.Rectangle(0, 0, 0, 0)
        Me.dgvLogframe.DragMultipleCells = False
        Me.dgvLogframe.DragOperator = CType(0, Byte)
        Me.dgvLogframe.DragReleased = True
        Me.dgvLogframe.HideEmptyCells = False
        Me.dgvLogframe.Logframe = Nothing
        Me.dgvLogframe.Name = "dgvLogframe"
        Me.dgvLogframe.RowHeadersVisible = False
        Me.dgvLogframe.RowModifyIndex = -1
        Me.dgvLogframe.ShowActivities = True
        Me.dgvLogframe.ShowAssumptionColumn = True
        Me.dgvLogframe.ShowCellToolTips = False
        Me.dgvLogframe.ShowGoals = True
        Me.dgvLogframe.ShowIndicatorColumn = True
        Me.dgvLogframe.ShowOutputs = True
        Me.dgvLogframe.ShowPurposes = True
        Me.dgvLogframe.ShowResourcesBudget = True
        Me.dgvLogframe.ShowVerificationSourceColumn = True
        Me.dgvLogframe.VirtualMode = True
        '
        'TabPagePlanning
        '
        resources.ApplyResources(Me.TabPagePlanning, "TabPagePlanning")
        Me.TabPagePlanning.Controls.Add(Me.SplitContainerPlanning)
        Me.TabPagePlanning.Name = "TabPagePlanning"
        Me.TabPagePlanning.UseVisualStyleBackColor = True
        '
        'SplitContainerPlanning
        '
        resources.ApplyResources(Me.SplitContainerPlanning, "SplitContainerPlanning")
        Me.SplitContainerPlanning.Name = "SplitContainerPlanning"
        '
        'SplitContainerPlanning.Panel1
        '
        resources.ApplyResources(Me.SplitContainerPlanning.Panel1, "SplitContainerPlanning.Panel1")
        Me.SplitContainerPlanning.Panel1.Controls.Add(Me.dgvPlanning)
        '
        'SplitContainerPlanning.Panel2
        '
        resources.ApplyResources(Me.SplitContainerPlanning.Panel2, "SplitContainerPlanning.Panel2")
        '
        'dgvPlanning
        '
        resources.ApplyResources(Me.dgvPlanning, "dgvPlanning")
        Me.dgvPlanning.AllowUserToAddRows = False
        Me.dgvPlanning.AllowUserToDeleteRows = False
        Me.dgvPlanning.AllowUserToResizeRows = False
        Me.dgvPlanning.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvPlanning.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvPlanning.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPlanning.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.RichTextColumnLogframe1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5})
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvPlanning.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvPlanning.DragBoxFromMouseDown = New System.Drawing.Rectangle(0, 0, 0, 0)
        Me.dgvPlanning.DragLink = False
        Me.dgvPlanning.DragMultipleCells = False
        Me.dgvPlanning.DragOperator = CType(0, Byte)
        Me.dgvPlanning.DragReleased = True
        Me.dgvPlanning.ElementsView = 0
        Me.dgvPlanning.EndDate = New Date(CType(0, Long))
        Me.dgvPlanning.HideEmptyCells = False
        Me.dgvPlanning.Name = "dgvPlanning"
        Me.dgvPlanning.PeriodEnd = New Date(CType(0, Long))
        Me.dgvPlanning.PeriodStart = New Date(CType(0, Long))
        Me.dgvPlanning.PeriodView = 2
        Me.dgvPlanning.RowHeadersVisible = False
        Me.dgvPlanning.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvPlanning.ShowActivityLinks = False
        Me.dgvPlanning.ShowCellToolTips = False
        Me.dgvPlanning.ShowDatesColumns = False
        Me.dgvPlanning.ShowKeyMomentLinks = False
        Me.dgvPlanning.StartDate = New Date(CType(0, Long))
        Me.dgvPlanning.Unlink = False
        Me.dgvPlanning.VirtualMode = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn1.Frozen = True
        resources.ApplyResources(Me.DataGridViewTextBoxColumn1, "DataGridViewTextBoxColumn1")
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'RichTextColumnLogframe1
        '
        Me.RichTextColumnLogframe1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.RichTextColumnLogframe1.Frozen = True
        resources.ApplyResources(Me.RichTextColumnLogframe1, "RichTextColumnLogframe1")
        Me.RichTextColumnLogframe1.Name = "RichTextColumnLogframe1"
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn2.Frozen = True
        resources.ApplyResources(Me.DataGridViewTextBoxColumn2, "DataGridViewTextBoxColumn2")
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.DataGridViewTextBoxColumn3.Frozen = True
        resources.ApplyResources(Me.DataGridViewTextBoxColumn3, "DataGridViewTextBoxColumn3")
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        resources.ApplyResources(Me.DataGridViewTextBoxColumn4, "DataGridViewTextBoxColumn4")
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        resources.ApplyResources(Me.DataGridViewTextBoxColumn5, "DataGridViewTextBoxColumn5")
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        '
        'TabPageBudget
        '
        resources.ApplyResources(Me.TabPageBudget, "TabPageBudget")
        Me.TabPageBudget.Controls.Add(Me.SplitContainerExchangeRates)
        Me.TabPageBudget.Name = "TabPageBudget"
        Me.TabPageBudget.UseVisualStyleBackColor = True
        '
        'SplitContainerExchangeRates
        '
        resources.ApplyResources(Me.SplitContainerExchangeRates, "SplitContainerExchangeRates")
        Me.SplitContainerExchangeRates.Name = "SplitContainerExchangeRates"
        '
        'SplitContainerExchangeRates.Panel1
        '
        resources.ApplyResources(Me.SplitContainerExchangeRates.Panel1, "SplitContainerExchangeRates.Panel1")
        Me.SplitContainerExchangeRates.Panel1.Controls.Add(Me.TabControlBudget)
        '
        'SplitContainerExchangeRates.Panel2
        '
        resources.ApplyResources(Me.SplitContainerExchangeRates.Panel2, "SplitContainerExchangeRates.Panel2")
        '
        'TabControlBudget
        '
        resources.ApplyResources(Me.TabControlBudget, "TabControlBudget")
        Me.TabControlBudget.Controls.Add(Me.TabPage1)
        Me.TabControlBudget.Name = "TabControlBudget"
        Me.TabControlBudget.SelectedIndex = 0
        '
        'TabPage1
        '
        resources.ApplyResources(Me.TabPage1, "TabPage1")
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPageMonitoring
        '
        resources.ApplyResources(Me.TabPageMonitoring, "TabPageMonitoring")
        Me.TabPageMonitoring.Controls.Add(Me.TabControlQuestionnaires)
        Me.TabPageMonitoring.Name = "TabPageMonitoring"
        Me.TabPageMonitoring.UseVisualStyleBackColor = True
        '
        'TabControlQuestionnaires
        '
        resources.ApplyResources(Me.TabControlQuestionnaires, "TabControlQuestionnaires")
        Me.TabControlQuestionnaires.Controls.Add(Me.TabPageProgrammeLevel)
        Me.TabControlQuestionnaires.Controls.Add(Me.TabPageTeamLevel)
        Me.TabControlQuestionnaires.Name = "TabControlQuestionnaires"
        Me.TabControlQuestionnaires.SelectedIndex = 0
        '
        'TabPageProgrammeLevel
        '
        resources.ApplyResources(Me.TabPageProgrammeLevel, "TabPageProgrammeLevel")
        Me.TabPageProgrammeLevel.Name = "TabPageProgrammeLevel"
        Me.TabPageProgrammeLevel.UseVisualStyleBackColor = True
        '
        'TabPageTeamLevel
        '
        resources.ApplyResources(Me.TabPageTeamLevel, "TabPageTeamLevel")
        Me.TabPageTeamLevel.Name = "TabPageTeamLevel"
        Me.TabPageTeamLevel.UseVisualStyleBackColor = True
        '
        'TabPageActions
        '
        resources.ApplyResources(Me.TabPageActions, "TabPageActions")
        Me.TabPageActions.Name = "TabPageActions"
        Me.TabPageActions.UseVisualStyleBackColor = True
        '
        'ImageListTabs
        '
        Me.ImageListTabs.ImageStream = CType(resources.GetObject("ImageListTabs.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListTabs.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListTabs.Images.SetKeyName(0, "LF 2.0 - File buttons - project groot.png")
        Me.ImageListTabs.Images.SetKeyName(1, "LF 2.0 - File buttons - print planning groot.png")
        Me.ImageListTabs.Images.SetKeyName(2, "LF 2.0 - File buttons - print budget groot.png")
        '
        'frmProject
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SplitContainerUtilities)
        Me.Name = "frmProject"
        Me.SplitContainerUtilities.Panel1.ResumeLayout(False)
        Me.SplitContainerUtilities.Panel2.ResumeLayout(False)
        CType(Me.SplitContainerUtilities, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerUtilities.ResumeLayout(False)
        Me.SplitContainerMain.Panel1.ResumeLayout(False)
        CType(Me.SplitContainerMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerMain.ResumeLayout(False)
        Me.TabControlProject.ResumeLayout(False)
        Me.TabPageProject.ResumeLayout(False)
        Me.TabPageLogframe.ResumeLayout(False)
        Me.SplitContainerLogFrame.Panel1.ResumeLayout(False)
        CType(Me.SplitContainerLogFrame, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerLogFrame.ResumeLayout(False)
        CType(Me.dgvLogframe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPagePlanning.ResumeLayout(False)
        Me.SplitContainerPlanning.Panel1.ResumeLayout(False)
        CType(Me.SplitContainerPlanning, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerPlanning.ResumeLayout(False)
        CType(Me.dgvPlanning, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageBudget.ResumeLayout(False)
        Me.SplitContainerExchangeRates.Panel1.ResumeLayout(False)
        CType(Me.SplitContainerExchangeRates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerExchangeRates.ResumeLayout(False)
        Me.TabControlBudget.ResumeLayout(False)
        Me.TabPageMonitoring.ResumeLayout(False)
        Me.TabControlQuestionnaires.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainerUtilities As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainerMain As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControlProject As FaciliDev.LogFramer.ucTabControlProject
    Friend WithEvents TabPageProject As System.Windows.Forms.TabPage
    Friend WithEvents ProjectInfo As FaciliDev.LogFramer.DetailProjectInfo
    Friend WithEvents TabPageLogframe As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainerLogFrame As System.Windows.Forms.SplitContainer
    Friend WithEvents TabPagePlanning As System.Windows.Forms.TabPage
    Friend WithEvents TabPageBudget As System.Windows.Forms.TabPage
    Friend WithEvents TabPageActions As System.Windows.Forms.TabPage
    Friend WithEvents Clipboard As FaciliDev.LogFramer.DetailClipboard
    Friend WithEvents dgvLogframe As FaciliDev.LogFramer.DataGridViewLogframe
    Friend WithEvents SplitContainerPlanning As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvPlanning As FaciliDev.LogFramer.DataGridViewPlanning
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RichTextColumnLogframe1 As FaciliDev.LogFramer.RichTextColumnLogframe
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPageMonitoring As System.Windows.Forms.TabPage
    Friend WithEvents TabControlQuestionnaires As System.Windows.Forms.TabControl
    Friend WithEvents TabPageProgrammeLevel As System.Windows.Forms.TabPage
    Friend WithEvents TabPageTeamLevel As System.Windows.Forms.TabPage
    Friend WithEvents SplitContainerExchangeRates As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControlBudget As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents ImageListTabs As System.Windows.Forms.ImageList
End Class
