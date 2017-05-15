<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailIndicator
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailIndicator))
        Me.lblTargetGroupName = New System.Windows.Forms.Label()
        Me.tbIndicator = New System.Windows.Forms.TextBox()
        Me.TabControlIndicator = New System.Windows.Forms.TabControl()
        Me.TabPageScores = New System.Windows.Forms.TabPage()
        Me.chkAdvanced = New System.Windows.Forms.CheckBox()
        Me.cmbQuestionType = New FaciliDev.LogFramer.StructuredComboBox()
        Me.PanelQuestionType = New System.Windows.Forms.Panel()
        Me.lblResponseType = New System.Windows.Forms.Label()
        Me.TabPageRegistration = New System.Windows.Forms.TabPage()
        Me.cmbAggregateHorizontal = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblAggregateHorizontal = New System.Windows.Forms.Label()
        Me.lblRegistration = New System.Windows.Forms.Label()
        Me.cmbRegistration = New FaciliDev.LogFramer.ComboBoxText()
        Me.cmbTargetGroupGuid = New FaciliDev.LogFramer.ComboBoxText()
        Me.TabPageTargets = New System.Windows.Forms.TabPage()
        Me.PanelTargets = New System.Windows.Forms.Panel()
        Me.TabPageSubIndicators = New System.Windows.Forms.TabPage()
        Me.gbWeightingChildren = New System.Windows.Forms.GroupBox()
        Me.lblWeightingFactorChildren = New System.Windows.Forms.Label()
        Me.ntbWeightingFactorChildren = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.gbSubIndicators = New System.Windows.Forms.GroupBox()
        Me.gbTotals = New System.Windows.Forms.GroupBox()
        Me.cmbAggregateVertical = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblAggregateVertical = New System.Windows.Forms.Label()
        Me.TabControlIndicator.SuspendLayout()
        Me.TabPageScores.SuspendLayout()
        Me.TabPageRegistration.SuspendLayout()
        Me.TabPageTargets.SuspendLayout()
        Me.TabPageSubIndicators.SuspendLayout()
        Me.gbWeightingChildren.SuspendLayout()
        Me.gbTotals.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblTargetGroupName
        '
        resources.ApplyResources(Me.lblTargetGroupName, "lblTargetGroupName")
        Me.lblTargetGroupName.Name = "lblTargetGroupName"
        '
        'tbIndicator
        '
        resources.ApplyResources(Me.tbIndicator, "tbIndicator")
        Me.tbIndicator.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbIndicator.ForeColor = System.Drawing.Color.Blue
        Me.tbIndicator.Name = "tbIndicator"
        Me.tbIndicator.ReadOnly = True
        '
        'TabControlIndicator
        '
        resources.ApplyResources(Me.TabControlIndicator, "TabControlIndicator")
        Me.TabControlIndicator.Controls.Add(Me.TabPageScores)
        Me.TabControlIndicator.Controls.Add(Me.TabPageRegistration)
        Me.TabControlIndicator.Controls.Add(Me.TabPageTargets)
        Me.TabControlIndicator.Controls.Add(Me.TabPageSubIndicators)
        Me.TabControlIndicator.Name = "TabControlIndicator"
        Me.TabControlIndicator.SelectedIndex = 0
        '
        'TabPageScores
        '
        resources.ApplyResources(Me.TabPageScores, "TabPageScores")
        Me.TabPageScores.Controls.Add(Me.chkAdvanced)
        Me.TabPageScores.Controls.Add(Me.cmbQuestionType)
        Me.TabPageScores.Controls.Add(Me.PanelQuestionType)
        Me.TabPageScores.Controls.Add(Me.lblResponseType)
        Me.TabPageScores.Name = "TabPageScores"
        Me.TabPageScores.UseVisualStyleBackColor = True
        '
        'chkAdvanced
        '
        resources.ApplyResources(Me.chkAdvanced, "chkAdvanced")
        Me.chkAdvanced.Name = "chkAdvanced"
        Me.chkAdvanced.UseVisualStyleBackColor = True
        '
        'cmbQuestionType
        '
        resources.ApplyResources(Me.cmbQuestionType, "cmbQuestionType")
        Me.cmbQuestionType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbQuestionType.FormattingEnabled = True
        Me.cmbQuestionType.Name = "cmbQuestionType"
        '
        'PanelQuestionType
        '
        resources.ApplyResources(Me.PanelQuestionType, "PanelQuestionType")
        Me.PanelQuestionType.Name = "PanelQuestionType"
        '
        'lblResponseType
        '
        resources.ApplyResources(Me.lblResponseType, "lblResponseType")
        Me.lblResponseType.Name = "lblResponseType"
        '
        'TabPageRegistration
        '
        resources.ApplyResources(Me.TabPageRegistration, "TabPageRegistration")
        Me.TabPageRegistration.Controls.Add(Me.cmbAggregateHorizontal)
        Me.TabPageRegistration.Controls.Add(Me.lblAggregateHorizontal)
        Me.TabPageRegistration.Controls.Add(Me.lblRegistration)
        Me.TabPageRegistration.Controls.Add(Me.lblTargetGroupName)
        Me.TabPageRegistration.Controls.Add(Me.cmbRegistration)
        Me.TabPageRegistration.Controls.Add(Me.cmbTargetGroupGuid)
        Me.TabPageRegistration.Name = "TabPageRegistration"
        Me.TabPageRegistration.UseVisualStyleBackColor = True
        '
        'cmbAggregateHorizontal
        '
        resources.ApplyResources(Me.cmbAggregateHorizontal, "cmbAggregateHorizontal")
        Me.cmbAggregateHorizontal.FormattingEnabled = True
        Me.cmbAggregateHorizontal.Name = "cmbAggregateHorizontal"
        '
        'lblAggregateHorizontal
        '
        resources.ApplyResources(Me.lblAggregateHorizontal, "lblAggregateHorizontal")
        Me.lblAggregateHorizontal.Name = "lblAggregateHorizontal"
        '
        'lblRegistration
        '
        resources.ApplyResources(Me.lblRegistration, "lblRegistration")
        Me.lblRegistration.Name = "lblRegistration"
        '
        'cmbRegistration
        '
        resources.ApplyResources(Me.cmbRegistration, "cmbRegistration")
        Me.cmbRegistration.FormattingEnabled = True
        Me.cmbRegistration.Name = "cmbRegistration"
        '
        'cmbTargetGroupGuid
        '
        resources.ApplyResources(Me.cmbTargetGroupGuid, "cmbTargetGroupGuid")
        Me.cmbTargetGroupGuid.FormattingEnabled = True
        Me.cmbTargetGroupGuid.Name = "cmbTargetGroupGuid"
        '
        'TabPageTargets
        '
        resources.ApplyResources(Me.TabPageTargets, "TabPageTargets")
        Me.TabPageTargets.Controls.Add(Me.PanelTargets)
        Me.TabPageTargets.Name = "TabPageTargets"
        Me.TabPageTargets.UseVisualStyleBackColor = True
        '
        'PanelTargets
        '
        resources.ApplyResources(Me.PanelTargets, "PanelTargets")
        Me.PanelTargets.Name = "PanelTargets"
        '
        'TabPageSubIndicators
        '
        resources.ApplyResources(Me.TabPageSubIndicators, "TabPageSubIndicators")
        Me.TabPageSubIndicators.Controls.Add(Me.gbWeightingChildren)
        Me.TabPageSubIndicators.Controls.Add(Me.gbSubIndicators)
        Me.TabPageSubIndicators.Controls.Add(Me.gbTotals)
        Me.TabPageSubIndicators.Name = "TabPageSubIndicators"
        Me.TabPageSubIndicators.UseVisualStyleBackColor = True
        '
        'gbWeightingChildren
        '
        resources.ApplyResources(Me.gbWeightingChildren, "gbWeightingChildren")
        Me.gbWeightingChildren.Controls.Add(Me.lblWeightingFactorChildren)
        Me.gbWeightingChildren.Controls.Add(Me.ntbWeightingFactorChildren)
        Me.gbWeightingChildren.Name = "gbWeightingChildren"
        Me.gbWeightingChildren.TabStop = False
        '
        'lblWeightingFactorChildren
        '
        resources.ApplyResources(Me.lblWeightingFactorChildren, "lblWeightingFactorChildren")
        Me.lblWeightingFactorChildren.Name = "lblWeightingFactorChildren"
        '
        'ntbWeightingFactorChildren
        '
        resources.ApplyResources(Me.ntbWeightingFactorChildren, "ntbWeightingFactorChildren")
        Me.ntbWeightingFactorChildren.AllowSpace = True
        Me.ntbWeightingFactorChildren.IsCurrency = False
        Me.ntbWeightingFactorChildren.IsPercentage = False
        Me.ntbWeightingFactorChildren.Name = "ntbWeightingFactorChildren"
        '
        'gbSubIndicators
        '
        resources.ApplyResources(Me.gbSubIndicators, "gbSubIndicators")
        Me.gbSubIndicators.Name = "gbSubIndicators"
        Me.gbSubIndicators.TabStop = False
        '
        'gbTotals
        '
        resources.ApplyResources(Me.gbTotals, "gbTotals")
        Me.gbTotals.Controls.Add(Me.cmbAggregateVertical)
        Me.gbTotals.Controls.Add(Me.lblAggregateVertical)
        Me.gbTotals.Name = "gbTotals"
        Me.gbTotals.TabStop = False
        '
        'cmbAggregateVertical
        '
        resources.ApplyResources(Me.cmbAggregateVertical, "cmbAggregateVertical")
        Me.cmbAggregateVertical.FormattingEnabled = True
        Me.cmbAggregateVertical.Name = "cmbAggregateVertical"
        '
        'lblAggregateVertical
        '
        resources.ApplyResources(Me.lblAggregateVertical, "lblAggregateVertical")
        Me.lblAggregateVertical.Name = "lblAggregateVertical"
        '
        'DetailIndicator
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.TabControlIndicator)
        Me.Controls.Add(Me.tbIndicator)
        Me.Name = "DetailIndicator"
        Me.TabControlIndicator.ResumeLayout(False)
        Me.TabPageScores.ResumeLayout(False)
        Me.TabPageScores.PerformLayout()
        Me.TabPageRegistration.ResumeLayout(False)
        Me.TabPageRegistration.PerformLayout()
        Me.TabPageTargets.ResumeLayout(False)
        Me.TabPageSubIndicators.ResumeLayout(False)
        Me.gbWeightingChildren.ResumeLayout(False)
        Me.gbWeightingChildren.PerformLayout()
        Me.gbTotals.ResumeLayout(False)
        Me.gbTotals.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblTargetGroupName As System.Windows.Forms.Label
    Friend WithEvents cmbTargetGroupGuid As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents tbIndicator As System.Windows.Forms.TextBox
    Friend WithEvents TabControlIndicator As System.Windows.Forms.TabControl
    Friend WithEvents TabPageScores As System.Windows.Forms.TabPage
    Friend WithEvents TabPageSubIndicators As System.Windows.Forms.TabPage
    Friend WithEvents TabPageRegistration As System.Windows.Forms.TabPage
    Friend WithEvents cmbRegistration As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblRegistration As System.Windows.Forms.Label
    Friend WithEvents lblResponseType As System.Windows.Forms.Label
    Friend WithEvents cmbAggregateHorizontal As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblAggregateHorizontal As System.Windows.Forms.Label
    Friend WithEvents PanelQuestionType As System.Windows.Forms.Panel
    Friend WithEvents TabPageTargets As System.Windows.Forms.TabPage
    Friend WithEvents cmbQuestionType As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents PanelTargets As System.Windows.Forms.Panel
    Friend WithEvents chkAdvanced As System.Windows.Forms.CheckBox
    Friend WithEvents gbSubIndicators As System.Windows.Forms.GroupBox
    Friend WithEvents gbTotals As System.Windows.Forms.GroupBox
    Friend WithEvents cmbAggregateVertical As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblAggregateVertical As System.Windows.Forms.Label
    Friend WithEvents gbWeightingChildren As System.Windows.Forms.GroupBox
    Friend WithEvents lblWeightingFactorChildren As System.Windows.Forms.Label
    Friend WithEvents ntbWeightingFactorChildren As FaciliDev.LogFramer.NumericTextBoxLF

End Class
