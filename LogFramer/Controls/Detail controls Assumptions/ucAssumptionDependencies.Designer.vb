<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AssumptionDependencies
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AssumptionDependencies))
        Me.TableLayoutPanelAssumptions = New System.Windows.Forms.TableLayoutPanel()
        Me.gbDeliverables = New System.Windows.Forms.GroupBox()
        Me.lblDateDelivered = New System.Windows.Forms.Label()
        Me.lblDateExpected = New System.Windows.Forms.Label()
        Me.dtbDateDelivered = New FaciliDev.LogFramer.DateTextBox()
        Me.dtbDateExpected = New FaciliDev.LogFramer.DateTextBox()
        Me.lblDeliverableType = New System.Windows.Forms.Label()
        Me.cmbDeliverableType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbSupplier = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblSupplier = New System.Windows.Forms.Label()
        Me.tbDeliverables = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblDeliverables = New System.Windows.Forms.Label()
        Me.gbResponse = New System.Windows.Forms.GroupBox()
        Me.cmbOwner = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblRiskOwner = New System.Windows.Forms.Label()
        Me.tbResponseStrategy = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblResponseStrategy = New System.Windows.Forms.Label()
        Me.gbDependency = New System.Windows.Forms.GroupBox()
        Me.lblInputType = New System.Windows.Forms.Label()
        Me.cmbInputType = New FaciliDev.LogFramer.StructuredComboBox()
        Me.lblDependencyType = New System.Windows.Forms.Label()
        Me.cmbDependencyType = New FaciliDev.LogFramer.StructuredComboBox()
        Me.tbImpact = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblImpact = New System.Windows.Forms.Label()
        Me.lblImportanceLevel = New System.Windows.Forms.Label()
        Me.cmbImportanceLevel = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.TableLayoutPanelAssumptions.SuspendLayout()
        Me.gbDeliverables.SuspendLayout()
        Me.gbResponse.SuspendLayout()
        Me.gbDependency.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanelAssumptions
        '
        resources.ApplyResources(Me.TableLayoutPanelAssumptions, "TableLayoutPanelAssumptions")
        Me.TableLayoutPanelAssumptions.Controls.Add(Me.gbDeliverables, 0, 0)
        Me.TableLayoutPanelAssumptions.Controls.Add(Me.gbResponse, 2, 0)
        Me.TableLayoutPanelAssumptions.Controls.Add(Me.gbDependency, 0, 0)
        Me.TableLayoutPanelAssumptions.Name = "TableLayoutPanelAssumptions"
        '
        'gbDeliverables
        '
        resources.ApplyResources(Me.gbDeliverables, "gbDeliverables")
        Me.gbDeliverables.Controls.Add(Me.lblDateDelivered)
        Me.gbDeliverables.Controls.Add(Me.lblDateExpected)
        Me.gbDeliverables.Controls.Add(Me.dtbDateDelivered)
        Me.gbDeliverables.Controls.Add(Me.dtbDateExpected)
        Me.gbDeliverables.Controls.Add(Me.lblDeliverableType)
        Me.gbDeliverables.Controls.Add(Me.cmbDeliverableType)
        Me.gbDeliverables.Controls.Add(Me.cmbSupplier)
        Me.gbDeliverables.Controls.Add(Me.lblSupplier)
        Me.gbDeliverables.Controls.Add(Me.tbDeliverables)
        Me.gbDeliverables.Controls.Add(Me.lblDeliverables)
        Me.gbDeliverables.Name = "gbDeliverables"
        Me.gbDeliverables.TabStop = False
        '
        'lblDateDelivered
        '
        resources.ApplyResources(Me.lblDateDelivered, "lblDateDelivered")
        Me.lblDateDelivered.Name = "lblDateDelivered"
        '
        'lblDateExpected
        '
        resources.ApplyResources(Me.lblDateExpected, "lblDateExpected")
        Me.lblDateExpected.Name = "lblDateExpected"
        '
        'dtbDateDelivered
        '
        resources.ApplyResources(Me.dtbDateDelivered, "dtbDateDelivered")
        Me.dtbDateDelivered.DateValue = New Date(CType(0, Long))
        Me.dtbDateDelivered.ForeColor = System.Drawing.SystemColors.Window
        Me.dtbDateDelivered.Name = "dtbDateDelivered"
        '
        'dtbDateExpected
        '
        resources.ApplyResources(Me.dtbDateExpected, "dtbDateExpected")
        Me.dtbDateExpected.DateValue = New Date(CType(0, Long))
        Me.dtbDateExpected.ForeColor = System.Drawing.SystemColors.Window
        Me.dtbDateExpected.Name = "dtbDateExpected"
        '
        'lblDeliverableType
        '
        resources.ApplyResources(Me.lblDeliverableType, "lblDeliverableType")
        Me.lblDeliverableType.Name = "lblDeliverableType"
        '
        'cmbDeliverableType
        '
        resources.ApplyResources(Me.cmbDeliverableType, "cmbDeliverableType")
        Me.cmbDeliverableType.FormattingEnabled = True
        Me.cmbDeliverableType.Name = "cmbDeliverableType"
        '
        'cmbSupplier
        '
        resources.ApplyResources(Me.cmbSupplier, "cmbSupplier")
        Me.cmbSupplier.FormattingEnabled = True
        Me.cmbSupplier.Name = "cmbSupplier"
        '
        'lblSupplier
        '
        resources.ApplyResources(Me.lblSupplier, "lblSupplier")
        Me.lblSupplier.Name = "lblSupplier"
        '
        'tbDeliverables
        '
        Me.tbDeliverables.AcceptsReturn = True
        resources.ApplyResources(Me.tbDeliverables, "tbDeliverables")
        Me.tbDeliverables.Name = "tbDeliverables"
        '
        'lblDeliverables
        '
        resources.ApplyResources(Me.lblDeliverables, "lblDeliverables")
        Me.lblDeliverables.Name = "lblDeliverables"
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
        'gbDependency
        '
        resources.ApplyResources(Me.gbDependency, "gbDependency")
        Me.gbDependency.Controls.Add(Me.lblInputType)
        Me.gbDependency.Controls.Add(Me.cmbInputType)
        Me.gbDependency.Controls.Add(Me.lblDependencyType)
        Me.gbDependency.Controls.Add(Me.cmbDependencyType)
        Me.gbDependency.Controls.Add(Me.tbImpact)
        Me.gbDependency.Controls.Add(Me.lblImpact)
        Me.gbDependency.Controls.Add(Me.lblImportanceLevel)
        Me.gbDependency.Controls.Add(Me.cmbImportanceLevel)
        Me.gbDependency.Name = "gbDependency"
        Me.gbDependency.TabStop = False
        '
        'lblInputType
        '
        resources.ApplyResources(Me.lblInputType, "lblInputType")
        Me.lblInputType.Name = "lblInputType"
        '
        'cmbInputType
        '
        resources.ApplyResources(Me.cmbInputType, "cmbInputType")
        Me.cmbInputType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbInputType.FormattingEnabled = True
        Me.cmbInputType.Name = "cmbInputType"
        '
        'lblDependencyType
        '
        resources.ApplyResources(Me.lblDependencyType, "lblDependencyType")
        Me.lblDependencyType.Name = "lblDependencyType"
        '
        'cmbDependencyType
        '
        resources.ApplyResources(Me.cmbDependencyType, "cmbDependencyType")
        Me.cmbDependencyType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbDependencyType.FormattingEnabled = True
        Me.cmbDependencyType.Name = "cmbDependencyType"
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
        'lblImportanceLevel
        '
        resources.ApplyResources(Me.lblImportanceLevel, "lblImportanceLevel")
        Me.lblImportanceLevel.Name = "lblImportanceLevel"
        '
        'cmbImportanceLevel
        '
        resources.ApplyResources(Me.cmbImportanceLevel, "cmbImportanceLevel")
        Me.cmbImportanceLevel.FormattingEnabled = True
        Me.cmbImportanceLevel.Name = "cmbImportanceLevel"
        '
        'AssumptionDependencies
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.TableLayoutPanelAssumptions)
        Me.Name = "AssumptionDependencies"
        Me.TableLayoutPanelAssumptions.ResumeLayout(False)
        Me.gbDeliverables.ResumeLayout(False)
        Me.gbDeliverables.PerformLayout()
        Me.gbResponse.ResumeLayout(False)
        Me.gbResponse.PerformLayout()
        Me.gbDependency.ResumeLayout(False)
        Me.gbDependency.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanelAssumptions As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents gbDependency As System.Windows.Forms.GroupBox
    Friend WithEvents lblImportanceLevel As System.Windows.Forms.Label
    Friend WithEvents cmbImportanceLevel As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents tbImpact As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblImpact As System.Windows.Forms.Label
    Friend WithEvents lblDependencyType As System.Windows.Forms.Label
    Friend WithEvents cmbDependencyType As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents lblInputType As System.Windows.Forms.Label
    Friend WithEvents cmbInputType As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents gbResponse As System.Windows.Forms.GroupBox
    Friend WithEvents cmbOwner As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblRiskOwner As System.Windows.Forms.Label
    Friend WithEvents tbResponseStrategy As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblResponseStrategy As System.Windows.Forms.Label
    Friend WithEvents gbDeliverables As System.Windows.Forms.GroupBox
    Friend WithEvents cmbSupplier As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblSupplier As System.Windows.Forms.Label
    Friend WithEvents tbDeliverables As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblDeliverables As System.Windows.Forms.Label
    Friend WithEvents lblDeliverableType As System.Windows.Forms.Label
    Friend WithEvents cmbDeliverableType As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents dtbDateDelivered As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents dtbDateExpected As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents lblDateExpected As System.Windows.Forms.Label
    Friend WithEvents lblDateDelivered As System.Windows.Forms.Label

End Class
