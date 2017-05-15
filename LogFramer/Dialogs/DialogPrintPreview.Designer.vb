<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogPrintPreview
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogPrintPreview))
        Me.TabControlReports = New System.Windows.Forms.TabControl()
        Me.TabPageLogFrame = New System.Windows.Forms.TabPage()
        Me.TabPagePlanning = New System.Windows.Forms.TabPage()
        Me.TabPageBudget = New System.Windows.Forms.TabPage()
        Me.TabPagePmf = New System.Windows.Forms.TabPage()
        Me.TabPageIndicators = New System.Windows.Forms.TabPage()
        Me.TabPageResources = New System.Windows.Forms.TabPage()
        Me.TabPageRiskRegister = New System.Windows.Forms.TabPage()
        Me.TabPageAssumptions = New System.Windows.Forms.TabPage()
        Me.TabPageDependencies = New System.Windows.Forms.TabPage()
        Me.TabPagePartnerList = New System.Windows.Forms.TabPage()
        Me.TabPageTargetGroupIdForm = New System.Windows.Forms.TabPage()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.TabControlReports.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlReports
        '
        resources.ApplyResources(Me.TabControlReports, "TabControlReports")
        Me.TabControlReports.Controls.Add(Me.TabPageLogFrame)
        Me.TabControlReports.Controls.Add(Me.TabPagePlanning)
        Me.TabControlReports.Controls.Add(Me.TabPageBudget)
        Me.TabControlReports.Controls.Add(Me.TabPagePmf)
        Me.TabControlReports.Controls.Add(Me.TabPageIndicators)
        Me.TabControlReports.Controls.Add(Me.TabPageResources)
        Me.TabControlReports.Controls.Add(Me.TabPageRiskRegister)
        Me.TabControlReports.Controls.Add(Me.TabPageAssumptions)
        Me.TabControlReports.Controls.Add(Me.TabPageDependencies)
        Me.TabControlReports.Controls.Add(Me.TabPagePartnerList)
        Me.TabControlReports.Controls.Add(Me.TabPageTargetGroupIdForm)
        Me.TabControlReports.Multiline = True
        Me.TabControlReports.Name = "TabControlReports"
        Me.TabControlReports.SelectedIndex = 0
        '
        'TabPageLogFrame
        '
        resources.ApplyResources(Me.TabPageLogFrame, "TabPageLogFrame")
        Me.TabPageLogFrame.Name = "TabPageLogFrame"
        Me.TabPageLogFrame.UseVisualStyleBackColor = True
        '
        'TabPagePlanning
        '
        resources.ApplyResources(Me.TabPagePlanning, "TabPagePlanning")
        Me.TabPagePlanning.Name = "TabPagePlanning"
        Me.TabPagePlanning.UseVisualStyleBackColor = True
        '
        'TabPageBudget
        '
        resources.ApplyResources(Me.TabPageBudget, "TabPageBudget")
        Me.TabPageBudget.Name = "TabPageBudget"
        Me.TabPageBudget.UseVisualStyleBackColor = True
        '
        'TabPagePmf
        '
        resources.ApplyResources(Me.TabPagePmf, "TabPagePmf")
        Me.TabPagePmf.Name = "TabPagePmf"
        Me.TabPagePmf.UseVisualStyleBackColor = True
        '
        'TabPageIndicators
        '
        resources.ApplyResources(Me.TabPageIndicators, "TabPageIndicators")
        Me.TabPageIndicators.Name = "TabPageIndicators"
        Me.TabPageIndicators.UseVisualStyleBackColor = True
        '
        'TabPageResources
        '
        resources.ApplyResources(Me.TabPageResources, "TabPageResources")
        Me.TabPageResources.Name = "TabPageResources"
        Me.TabPageResources.UseVisualStyleBackColor = True
        '
        'TabPageRiskRegister
        '
        resources.ApplyResources(Me.TabPageRiskRegister, "TabPageRiskRegister")
        Me.TabPageRiskRegister.Name = "TabPageRiskRegister"
        Me.TabPageRiskRegister.UseVisualStyleBackColor = True
        '
        'TabPageAssumptions
        '
        resources.ApplyResources(Me.TabPageAssumptions, "TabPageAssumptions")
        Me.TabPageAssumptions.Name = "TabPageAssumptions"
        Me.TabPageAssumptions.UseVisualStyleBackColor = True
        '
        'TabPageDependencies
        '
        resources.ApplyResources(Me.TabPageDependencies, "TabPageDependencies")
        Me.TabPageDependencies.Name = "TabPageDependencies"
        Me.TabPageDependencies.UseVisualStyleBackColor = True
        '
        'TabPagePartnerList
        '
        resources.ApplyResources(Me.TabPagePartnerList, "TabPagePartnerList")
        Me.TabPagePartnerList.Name = "TabPagePartnerList"
        Me.TabPagePartnerList.UseVisualStyleBackColor = True
        '
        'TabPageTargetGroupIdForm
        '
        resources.ApplyResources(Me.TabPageTargetGroupIdForm, "TabPageTargetGroupIdForm")
        Me.TabPageTargetGroupIdForm.Name = "TabPageTargetGroupIdForm"
        Me.TabPageTargetGroupIdForm.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        resources.ApplyResources(Me.btnClose, "btnClose")
        Me.btnClose.Name = "btnClose"
        '
        'DialogPrintPreview
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.TabControlReports)
        Me.DoubleBuffered = True
        Me.MinimizeBox = False
        Me.Name = "DialogPrintPreview"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.TabControlReports.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControlReports As System.Windows.Forms.TabControl
    Friend WithEvents TabPageLogFrame As System.Windows.Forms.TabPage
    Friend WithEvents TabPageIndicators As System.Windows.Forms.TabPage
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents TabPagePmf As System.Windows.Forms.TabPage
    Friend WithEvents TabPageRiskRegister As System.Windows.Forms.TabPage
    Friend WithEvents TabPagePartnerList As System.Windows.Forms.TabPage
    Friend WithEvents TabPageTargetGroupIdForm As System.Windows.Forms.TabPage
    Friend WithEvents TabPagePlanning As System.Windows.Forms.TabPage
    Friend WithEvents TabPageResources As System.Windows.Forms.TabPage
    Friend WithEvents TabPageBudget As System.Windows.Forms.TabPage
    Friend WithEvents TabPageAssumptions As System.Windows.Forms.TabPage
    Friend WithEvents TabPageDependencies As System.Windows.Forms.TabPage
End Class
