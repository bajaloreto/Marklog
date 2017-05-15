<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AssumptionRisks
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssumptionRisks))
        Me.TableLayoutPanelAssumptions = New System.Windows.Forms.TableLayoutPanel()
        Me.gbResponse = New System.Windows.Forms.GroupBox()
        Me.cmbOwner = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblOwner = New System.Windows.Forms.Label()
        Me.tbResponseStrategy = New FaciliDev.LogFramer.TextBoxLF()
        Me.cmbRiskResponse = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblResponseStrategy = New System.Windows.Forms.Label()
        Me.lblRiskResponse = New System.Windows.Forms.Label()
        Me.gbAssessment = New System.Windows.Forms.GroupBox()
        Me.tbImpact = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblImpact = New System.Windows.Forms.Label()
        Me.tbRiskLevel = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblRiskImpact = New System.Windows.Forms.Label()
        Me.cmbRiskImpact = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblLikelihood = New System.Windows.Forms.Label()
        Me.lblRiskCategory = New System.Windows.Forms.Label()
        Me.cmbRiskCategory = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblRiskLevel = New System.Windows.Forms.Label()
        Me.cmbLikelihood = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.TableLayoutPanelAssumptions.SuspendLayout()
        Me.gbResponse.SuspendLayout()
        Me.gbAssessment.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanelAssumptions
        '
        resources.ApplyResources(Me.TableLayoutPanelAssumptions, "TableLayoutPanelAssumptions")
        Me.TableLayoutPanelAssumptions.Controls.Add(Me.gbResponse, 0, 0)
        Me.TableLayoutPanelAssumptions.Controls.Add(Me.gbAssessment, 0, 0)
        Me.TableLayoutPanelAssumptions.Name = "TableLayoutPanelAssumptions"
        '
        'gbResponse
        '
        resources.ApplyResources(Me.gbResponse, "gbResponse")
        Me.gbResponse.Controls.Add(Me.cmbOwner)
        Me.gbResponse.Controls.Add(Me.lblOwner)
        Me.gbResponse.Controls.Add(Me.tbResponseStrategy)
        Me.gbResponse.Controls.Add(Me.cmbRiskResponse)
        Me.gbResponse.Controls.Add(Me.lblResponseStrategy)
        Me.gbResponse.Controls.Add(Me.lblRiskResponse)
        Me.gbResponse.Name = "gbResponse"
        Me.gbResponse.TabStop = False
        '
        'cmbOwner
        '
        resources.ApplyResources(Me.cmbOwner, "cmbOwner")
        Me.cmbOwner.FormattingEnabled = True
        Me.cmbOwner.Name = "cmbOwner"
        '
        'lblOwner
        '
        resources.ApplyResources(Me.lblOwner, "lblOwner")
        Me.lblOwner.Name = "lblOwner"
        '
        'tbResponseStrategy
        '
        Me.tbResponseStrategy.AcceptsReturn = True
        resources.ApplyResources(Me.tbResponseStrategy, "tbResponseStrategy")
        Me.tbResponseStrategy.Name = "tbResponseStrategy"
        '
        'cmbRiskResponse
        '
        resources.ApplyResources(Me.cmbRiskResponse, "cmbRiskResponse")
        Me.cmbRiskResponse.FormattingEnabled = True
        Me.cmbRiskResponse.Name = "cmbRiskResponse"
        '
        'lblResponseStrategy
        '
        resources.ApplyResources(Me.lblResponseStrategy, "lblResponseStrategy")
        Me.lblResponseStrategy.Name = "lblResponseStrategy"
        '
        'lblRiskResponse
        '
        resources.ApplyResources(Me.lblRiskResponse, "lblRiskResponse")
        Me.lblRiskResponse.Name = "lblRiskResponse"
        '
        'gbAssessment
        '
        resources.ApplyResources(Me.gbAssessment, "gbAssessment")
        Me.gbAssessment.Controls.Add(Me.tbImpact)
        Me.gbAssessment.Controls.Add(Me.lblImpact)
        Me.gbAssessment.Controls.Add(Me.tbRiskLevel)
        Me.gbAssessment.Controls.Add(Me.lblRiskImpact)
        Me.gbAssessment.Controls.Add(Me.cmbRiskImpact)
        Me.gbAssessment.Controls.Add(Me.lblLikelihood)
        Me.gbAssessment.Controls.Add(Me.lblRiskCategory)
        Me.gbAssessment.Controls.Add(Me.cmbRiskCategory)
        Me.gbAssessment.Controls.Add(Me.lblRiskLevel)
        Me.gbAssessment.Controls.Add(Me.cmbLikelihood)
        Me.gbAssessment.Name = "gbAssessment"
        Me.gbAssessment.TabStop = False
        '
        'tbImpact
        '
        Me.tbImpact.AcceptsReturn = True
        resources.ApplyResources(Me.tbImpact, "tbImpact")
        Me.tbImpact.Name = "tbImpact"
        '
        'lblImpact
        '
        resources.ApplyResources(Me.lblImpact, "lblImpact")
        Me.lblImpact.Name = "lblImpact"
        '
        'tbRiskLevel
        '
        resources.ApplyResources(Me.tbRiskLevel, "tbRiskLevel")
        Me.tbRiskLevel.Name = "tbRiskLevel"
        Me.tbRiskLevel.ReadOnly = True
        '
        'lblRiskImpact
        '
        resources.ApplyResources(Me.lblRiskImpact, "lblRiskImpact")
        Me.lblRiskImpact.Name = "lblRiskImpact"
        '
        'cmbRiskImpact
        '
        resources.ApplyResources(Me.cmbRiskImpact, "cmbRiskImpact")
        Me.cmbRiskImpact.FormattingEnabled = True
        Me.cmbRiskImpact.Name = "cmbRiskImpact"
        '
        'lblLikelihood
        '
        resources.ApplyResources(Me.lblLikelihood, "lblLikelihood")
        Me.lblLikelihood.Name = "lblLikelihood"
        '
        'lblRiskCategory
        '
        resources.ApplyResources(Me.lblRiskCategory, "lblRiskCategory")
        Me.lblRiskCategory.Name = "lblRiskCategory"
        '
        'cmbRiskCategory
        '
        resources.ApplyResources(Me.cmbRiskCategory, "cmbRiskCategory")
        Me.cmbRiskCategory.FormattingEnabled = True
        Me.cmbRiskCategory.Name = "cmbRiskCategory"
        '
        'lblRiskLevel
        '
        resources.ApplyResources(Me.lblRiskLevel, "lblRiskLevel")
        Me.lblRiskLevel.Name = "lblRiskLevel"
        '
        'cmbLikelihood
        '
        resources.ApplyResources(Me.cmbLikelihood, "cmbLikelihood")
        Me.cmbLikelihood.FormattingEnabled = True
        Me.cmbLikelihood.Name = "cmbLikelihood"
        '
        'AssumptionRisks
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.TableLayoutPanelAssumptions)
        Me.Name = "AssumptionRisks"
        Me.TableLayoutPanelAssumptions.ResumeLayout(False)
        Me.gbResponse.ResumeLayout(False)
        Me.gbResponse.PerformLayout()
        Me.gbAssessment.ResumeLayout(False)
        Me.gbAssessment.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanelAssumptions As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents gbResponse As System.Windows.Forms.GroupBox
    Friend WithEvents cmbOwner As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents tbResponseStrategy As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents cmbRiskResponse As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblResponseStrategy As System.Windows.Forms.Label
    Friend WithEvents lblRiskResponse As System.Windows.Forms.Label
    Friend WithEvents gbAssessment As System.Windows.Forms.GroupBox
    Friend WithEvents tbRiskLevel As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblRiskImpact As System.Windows.Forms.Label
    Friend WithEvents cmbRiskImpact As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblLikelihood As System.Windows.Forms.Label
    Friend WithEvents lblRiskCategory As System.Windows.Forms.Label
    Friend WithEvents cmbRiskCategory As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblRiskLevel As System.Windows.Forms.Label
    Friend WithEvents cmbLikelihood As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents tbImpact As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblImpact As System.Windows.Forms.Label

End Class
