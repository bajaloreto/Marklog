<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSettingsIndicators
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintSettingsIndicators))
        Me.GroupBoxStruct = New System.Windows.Forms.GroupBox()
        Me.cmbPrintSection = New System.Windows.Forms.ComboBox()
        Me.GroupBoxTargetGroup = New System.Windows.Forms.GroupBox()
        Me.cmbTargetGroupGuid = New System.Windows.Forms.ComboBox()
        Me.GroupBoxInclude = New System.Windows.Forms.GroupBox()
        Me.chkPrintActivities = New System.Windows.Forms.CheckBox()
        Me.chkPrintTargets = New System.Windows.Forms.CheckBox()
        Me.chkPrintRanges = New System.Windows.Forms.CheckBox()
        Me.chkPrintOptionValues = New System.Windows.Forms.CheckBox()
        Me.chkPrintOutputs = New System.Windows.Forms.CheckBox()
        Me.chkPrintPurposes = New System.Windows.Forms.CheckBox()
        Me.GroupBoxMeasurement = New System.Windows.Forms.GroupBox()
        Me.rbtnAllMeasurements = New System.Windows.Forms.RadioButton()
        Me.rbtnSingleMeasurement = New System.Windows.Forms.RadioButton()
        Me.GroupBoxStruct.SuspendLayout()
        Me.GroupBoxTargetGroup.SuspendLayout()
        Me.GroupBoxInclude.SuspendLayout()
        Me.GroupBoxMeasurement.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBoxStruct
        '
        resources.ApplyResources(Me.GroupBoxStruct, "GroupBoxStruct")
        Me.GroupBoxStruct.Controls.Add(Me.cmbPrintSection)
        Me.GroupBoxStruct.Name = "GroupBoxStruct"
        Me.GroupBoxStruct.TabStop = False
        '
        'cmbPrintSection
        '
        resources.ApplyResources(Me.cmbPrintSection, "cmbPrintSection")
        Me.cmbPrintSection.FormattingEnabled = True
        Me.cmbPrintSection.Name = "cmbPrintSection"
        '
        'GroupBoxTargetGroup
        '
        resources.ApplyResources(Me.GroupBoxTargetGroup, "GroupBoxTargetGroup")
        Me.GroupBoxTargetGroup.Controls.Add(Me.cmbTargetGroupGuid)
        Me.GroupBoxTargetGroup.Name = "GroupBoxTargetGroup"
        Me.GroupBoxTargetGroup.TabStop = False
        '
        'cmbTargetGroupGuid
        '
        resources.ApplyResources(Me.cmbTargetGroupGuid, "cmbTargetGroupGuid")
        Me.cmbTargetGroupGuid.FormattingEnabled = True
        Me.cmbTargetGroupGuid.Name = "cmbTargetGroupGuid"
        '
        'GroupBoxInclude
        '
        resources.ApplyResources(Me.GroupBoxInclude, "GroupBoxInclude")
        Me.GroupBoxInclude.Controls.Add(Me.chkPrintActivities)
        Me.GroupBoxInclude.Controls.Add(Me.chkPrintTargets)
        Me.GroupBoxInclude.Controls.Add(Me.chkPrintRanges)
        Me.GroupBoxInclude.Controls.Add(Me.chkPrintOptionValues)
        Me.GroupBoxInclude.Controls.Add(Me.chkPrintOutputs)
        Me.GroupBoxInclude.Controls.Add(Me.chkPrintPurposes)
        Me.GroupBoxInclude.Name = "GroupBoxInclude"
        Me.GroupBoxInclude.TabStop = False
        '
        'chkPrintActivities
        '
        resources.ApplyResources(Me.chkPrintActivities, "chkPrintActivities")
        Me.chkPrintActivities.Name = "chkPrintActivities"
        Me.chkPrintActivities.UseVisualStyleBackColor = True
        '
        'chkPrintTargets
        '
        resources.ApplyResources(Me.chkPrintTargets, "chkPrintTargets")
        Me.chkPrintTargets.Name = "chkPrintTargets"
        Me.chkPrintTargets.UseVisualStyleBackColor = True
        '
        'chkPrintRanges
        '
        resources.ApplyResources(Me.chkPrintRanges, "chkPrintRanges")
        Me.chkPrintRanges.Name = "chkPrintRanges"
        Me.chkPrintRanges.UseVisualStyleBackColor = True
        '
        'chkPrintOptionValues
        '
        resources.ApplyResources(Me.chkPrintOptionValues, "chkPrintOptionValues")
        Me.chkPrintOptionValues.Name = "chkPrintOptionValues"
        Me.chkPrintOptionValues.UseVisualStyleBackColor = True
        '
        'chkPrintOutputs
        '
        resources.ApplyResources(Me.chkPrintOutputs, "chkPrintOutputs")
        Me.chkPrintOutputs.Name = "chkPrintOutputs"
        Me.chkPrintOutputs.UseVisualStyleBackColor = True
        '
        'chkPrintPurposes
        '
        resources.ApplyResources(Me.chkPrintPurposes, "chkPrintPurposes")
        Me.chkPrintPurposes.Name = "chkPrintPurposes"
        Me.chkPrintPurposes.UseVisualStyleBackColor = True
        '
        'GroupBoxMeasurement
        '
        resources.ApplyResources(Me.GroupBoxMeasurement, "GroupBoxMeasurement")
        Me.GroupBoxMeasurement.Controls.Add(Me.rbtnAllMeasurements)
        Me.GroupBoxMeasurement.Controls.Add(Me.rbtnSingleMeasurement)
        Me.GroupBoxMeasurement.Name = "GroupBoxMeasurement"
        Me.GroupBoxMeasurement.TabStop = False
        '
        'rbtnAllMeasurements
        '
        resources.ApplyResources(Me.rbtnAllMeasurements, "rbtnAllMeasurements")
        Me.rbtnAllMeasurements.Name = "rbtnAllMeasurements"
        Me.rbtnAllMeasurements.TabStop = True
        Me.rbtnAllMeasurements.UseVisualStyleBackColor = True
        '
        'rbtnSingleMeasurement
        '
        resources.ApplyResources(Me.rbtnSingleMeasurement, "rbtnSingleMeasurement")
        Me.rbtnSingleMeasurement.Name = "rbtnSingleMeasurement"
        Me.rbtnSingleMeasurement.TabStop = True
        Me.rbtnSingleMeasurement.UseVisualStyleBackColor = True
        '
        'PrintSettingsIndicators
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBoxMeasurement)
        Me.Controls.Add(Me.GroupBoxInclude)
        Me.Controls.Add(Me.GroupBoxTargetGroup)
        Me.Controls.Add(Me.GroupBoxStruct)
        Me.Name = "PrintSettingsIndicators"
        Me.GroupBoxStruct.ResumeLayout(False)
        Me.GroupBoxTargetGroup.ResumeLayout(False)
        Me.GroupBoxInclude.ResumeLayout(False)
        Me.GroupBoxInclude.PerformLayout()
        Me.GroupBoxMeasurement.ResumeLayout(False)
        Me.GroupBoxMeasurement.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxStruct As System.Windows.Forms.GroupBox
    Friend WithEvents cmbPrintSection As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBoxTargetGroup As System.Windows.Forms.GroupBox
    Friend WithEvents cmbTargetGroupGuid As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBoxInclude As System.Windows.Forms.GroupBox
    Friend WithEvents chkPrintOptionValues As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintOutputs As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintPurposes As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintTargets As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintRanges As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBoxMeasurement As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnAllMeasurements As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnSingleMeasurement As System.Windows.Forms.RadioButton
    Friend WithEvents chkPrintActivities As System.Windows.Forms.CheckBox

End Class
