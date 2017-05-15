<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSettingsLogframe
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintSettingsLogframe))
        Me.GroupBoxColumns = New System.Windows.Forms.GroupBox()
        Me.chkShowAsm = New System.Windows.Forms.CheckBox()
        Me.chkShowVer = New System.Windows.Forms.CheckBox()
        Me.chkShowInd = New System.Windows.Forms.CheckBox()
        Me.chkShowStruct = New System.Windows.Forms.CheckBox()
        Me.gbStructure = New System.Windows.Forms.GroupBox()
        Me.rbtnPrintIndicators = New System.Windows.Forms.RadioButton()
        Me.rbtnPrintResourcesBudget = New System.Windows.Forms.RadioButton()
        Me.gbSections = New System.Windows.Forms.GroupBox()
        Me.chkShowPurposes = New System.Windows.Forms.CheckBox()
        Me.chkShowActivities = New System.Windows.Forms.CheckBox()
        Me.chkShowOutputs = New System.Windows.Forms.CheckBox()
        Me.chkShowGoals = New System.Windows.Forms.CheckBox()
        Me.GroupBoxColumns.SuspendLayout()
        Me.gbStructure.SuspendLayout()
        Me.gbSections.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBoxColumns
        '
        resources.ApplyResources(Me.GroupBoxColumns, "GroupBoxColumns")
        Me.GroupBoxColumns.Controls.Add(Me.chkShowAsm)
        Me.GroupBoxColumns.Controls.Add(Me.chkShowVer)
        Me.GroupBoxColumns.Controls.Add(Me.chkShowInd)
        Me.GroupBoxColumns.Controls.Add(Me.chkShowStruct)
        Me.GroupBoxColumns.Name = "GroupBoxColumns"
        Me.GroupBoxColumns.TabStop = False
        '
        'chkShowAsm
        '
        resources.ApplyResources(Me.chkShowAsm, "chkShowAsm")
        Me.chkShowAsm.Checked = True
        Me.chkShowAsm.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowAsm.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutAssumptions
        Me.chkShowAsm.Name = "chkShowAsm"
        Me.chkShowAsm.UseVisualStyleBackColor = True
        '
        'chkShowVer
        '
        resources.ApplyResources(Me.chkShowVer, "chkShowVer")
        Me.chkShowVer.Checked = True
        Me.chkShowVer.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowVer.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutVerificationSources
        Me.chkShowVer.Name = "chkShowVer"
        Me.chkShowVer.UseVisualStyleBackColor = True
        '
        'chkShowInd
        '
        resources.ApplyResources(Me.chkShowInd, "chkShowInd")
        Me.chkShowInd.Checked = True
        Me.chkShowInd.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowInd.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutIndicators
        Me.chkShowInd.Name = "chkShowInd"
        Me.chkShowInd.UseVisualStyleBackColor = True
        '
        'chkShowStruct
        '
        resources.ApplyResources(Me.chkShowStruct, "chkShowStruct")
        Me.chkShowStruct.Checked = True
        Me.chkShowStruct.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowStruct.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutProjectLogic
        Me.chkShowStruct.Name = "chkShowStruct"
        Me.chkShowStruct.UseVisualStyleBackColor = True
        '
        'gbStructure
        '
        resources.ApplyResources(Me.gbStructure, "gbStructure")
        Me.gbStructure.Controls.Add(Me.rbtnPrintIndicators)
        Me.gbStructure.Controls.Add(Me.rbtnPrintResourcesBudget)
        Me.gbStructure.Name = "gbStructure"
        Me.gbStructure.TabStop = False
        '
        'rbtnPrintIndicators
        '
        resources.ApplyResources(Me.rbtnPrintIndicators, "rbtnPrintIndicators")
        Me.rbtnPrintIndicators.Name = "rbtnPrintIndicators"
        Me.rbtnPrintIndicators.TabStop = True
        Me.rbtnPrintIndicators.UseVisualStyleBackColor = True
        '
        'rbtnPrintResourcesBudget
        '
        resources.ApplyResources(Me.rbtnPrintResourcesBudget, "rbtnPrintResourcesBudget")
        Me.rbtnPrintResourcesBudget.Name = "rbtnPrintResourcesBudget"
        Me.rbtnPrintResourcesBudget.TabStop = True
        Me.rbtnPrintResourcesBudget.UseVisualStyleBackColor = True
        '
        'gbSections
        '
        resources.ApplyResources(Me.gbSections, "gbSections")
        Me.gbSections.Controls.Add(Me.chkShowPurposes)
        Me.gbSections.Controls.Add(Me.chkShowActivities)
        Me.gbSections.Controls.Add(Me.chkShowOutputs)
        Me.gbSections.Controls.Add(Me.chkShowGoals)
        Me.gbSections.Name = "gbSections"
        Me.gbSections.TabStop = False
        '
        'chkShowPurposes
        '
        resources.ApplyResources(Me.chkShowPurposes, "chkShowPurposes")
        Me.chkShowPurposes.Checked = True
        Me.chkShowPurposes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowPurposes.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutPurposesSection
        Me.chkShowPurposes.Name = "chkShowPurposes"
        Me.chkShowPurposes.UseVisualStyleBackColor = True
        '
        'chkShowActivities
        '
        resources.ApplyResources(Me.chkShowActivities, "chkShowActivities")
        Me.chkShowActivities.Checked = True
        Me.chkShowActivities.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowActivities.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutActivitiesSection
        Me.chkShowActivities.Name = "chkShowActivities"
        Me.chkShowActivities.UseVisualStyleBackColor = True
        '
        'chkShowOutputs
        '
        resources.ApplyResources(Me.chkShowOutputs, "chkShowOutputs")
        Me.chkShowOutputs.Checked = True
        Me.chkShowOutputs.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowOutputs.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutOutputsSection
        Me.chkShowOutputs.Name = "chkShowOutputs"
        Me.chkShowOutputs.UseVisualStyleBackColor = True
        '
        'chkShowGoals
        '
        resources.ApplyResources(Me.chkShowGoals, "chkShowGoals")
        Me.chkShowGoals.Checked = True
        Me.chkShowGoals.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowGoals.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.LayOutGoalsSection
        Me.chkShowGoals.Name = "chkShowGoals"
        Me.chkShowGoals.UseVisualStyleBackColor = True
        '
        'PrintSettingsLogframe
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbSections)
        Me.Controls.Add(Me.gbStructure)
        Me.Controls.Add(Me.GroupBoxColumns)
        Me.Name = "PrintSettingsLogframe"
        Me.GroupBoxColumns.ResumeLayout(False)
        Me.gbStructure.ResumeLayout(False)
        Me.gbStructure.PerformLayout()
        Me.gbSections.ResumeLayout(False)
        Me.gbSections.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxColumns As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowAsm As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowVer As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowInd As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowStruct As System.Windows.Forms.CheckBox
    Friend WithEvents gbStructure As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnPrintIndicators As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPrintResourcesBudget As System.Windows.Forms.RadioButton
    Friend WithEvents gbSections As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowPurposes As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowActivities As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowOutputs As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowGoals As System.Windows.Forms.CheckBox

End Class
