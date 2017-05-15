<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AssumptionAssumptions
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssumptionAssumptions))
        Me.TableLayoutPanelAssumptions = New System.Windows.Forms.TableLayoutPanel()
        Me.gbResponse = New System.Windows.Forms.GroupBox()
        Me.cmbOwner = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblRiskOwner = New System.Windows.Forms.Label()
        Me.tbResponseStrategy = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblResponseStrategy = New System.Windows.Forms.Label()
        Me.gbAssessment = New System.Windows.Forms.GroupBox()
        Me.tbImpact = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblImpact = New System.Windows.Forms.Label()
        Me.tbHowToValidate = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblHowToValidate = New System.Windows.Forms.Label()
        Me.chkValidated = New FaciliDev.LogFramer.CheckBoxLF()
        Me.tbReason = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblReason = New System.Windows.Forms.Label()
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
        Me.gbResponse.Controls.Add(Me.lblRiskOwner)
        Me.gbResponse.Controls.Add(Me.tbResponseStrategy)
        Me.gbResponse.Controls.Add(Me.lblResponseStrategy)
        Me.gbResponse.Name = "gbResponse"
        Me.gbResponse.TabStop = False
        '
        'cmbOwner
        '
        resources.ApplyResources(Me.cmbOwner, "cmbOwner")
        Me.cmbOwner.FormattingEnabled = True
        Me.cmbOwner.Name = "cmbOwner"
        '
        'lblRiskOwner
        '
        resources.ApplyResources(Me.lblRiskOwner, "lblRiskOwner")
        Me.lblRiskOwner.Name = "lblRiskOwner"
        '
        'tbResponseStrategy
        '
        Me.tbResponseStrategy.AcceptsReturn = True
        resources.ApplyResources(Me.tbResponseStrategy, "tbResponseStrategy")
        Me.tbResponseStrategy.Name = "tbResponseStrategy"
        '
        'lblResponseStrategy
        '
        resources.ApplyResources(Me.lblResponseStrategy, "lblResponseStrategy")
        Me.lblResponseStrategy.Name = "lblResponseStrategy"
        '
        'gbAssessment
        '
        resources.ApplyResources(Me.gbAssessment, "gbAssessment")
        Me.gbAssessment.Controls.Add(Me.tbImpact)
        Me.gbAssessment.Controls.Add(Me.lblImpact)
        Me.gbAssessment.Controls.Add(Me.tbHowToValidate)
        Me.gbAssessment.Controls.Add(Me.lblHowToValidate)
        Me.gbAssessment.Controls.Add(Me.chkValidated)
        Me.gbAssessment.Controls.Add(Me.tbReason)
        Me.gbAssessment.Controls.Add(Me.lblReason)
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
        'tbHowToValidate
        '
        Me.tbHowToValidate.AcceptsReturn = True
        resources.ApplyResources(Me.tbHowToValidate, "tbHowToValidate")
        Me.tbHowToValidate.Name = "tbHowToValidate"
        '
        'lblHowToValidate
        '
        resources.ApplyResources(Me.lblHowToValidate, "lblHowToValidate")
        Me.lblHowToValidate.Name = "lblHowToValidate"
        '
        'chkValidated
        '
        resources.ApplyResources(Me.chkValidated, "chkValidated")
        Me.chkValidated.Name = "chkValidated"
        Me.chkValidated.Type = 0
        Me.chkValidated.UseVisualStyleBackColor = True
        '
        'tbReason
        '
        Me.tbReason.AcceptsReturn = True
        resources.ApplyResources(Me.tbReason, "tbReason")
        Me.tbReason.Name = "tbReason"
        '
        'lblReason
        '
        resources.ApplyResources(Me.lblReason, "lblReason")
        Me.lblReason.Name = "lblReason"
        '
        'AssumptionAssumptions
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.TableLayoutPanelAssumptions)
        Me.Name = "AssumptionAssumptions"
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
    Friend WithEvents lblRiskOwner As System.Windows.Forms.Label
    Friend WithEvents tbResponseStrategy As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblResponseStrategy As System.Windows.Forms.Label
    Friend WithEvents gbAssessment As System.Windows.Forms.GroupBox
    Friend WithEvents tbReason As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblReason As System.Windows.Forms.Label
    Friend WithEvents chkValidated As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents tbHowToValidate As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblHowToValidate As System.Windows.Forms.Label
    Friend WithEvents tbImpact As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblImpact As System.Windows.Forms.Label

End Class
