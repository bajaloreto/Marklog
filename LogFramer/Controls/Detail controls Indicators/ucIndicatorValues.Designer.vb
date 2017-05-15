<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorValues
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorValues))
        Me.gbValueRange = New System.Windows.Forms.GroupBox()
        Me.tbMaxValue = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.tbMinValue = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.cmbOpMax = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbOpMin = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.lblMaxValue = New System.Windows.Forms.Label()
        Me.lblMinValue = New System.Windows.Forms.Label()
        Me.gbNumber = New System.Windows.Forms.GroupBox()
        Me.tbNrDecimals = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblUnit = New System.Windows.Forms.Label()
        Me.cmbUnit = New FaciliDev.LogFramer.StructuredComboBox()
        Me.lblNrDecimals = New System.Windows.Forms.Label()
        Me.gbTargetSystem = New System.Windows.Forms.GroupBox()
        Me.rbtnTargetSystemFormula = New System.Windows.Forms.RadioButton()
        Me.rbtnTargetSystemValueRange = New System.Windows.Forms.RadioButton()
        Me.rbtnTargetSystemSimple = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringValue = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringPercentage = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringScore = New System.Windows.Forms.RadioButton()
        Me.gbScoring = New System.Windows.Forms.GroupBox()
        Me.gbValueRange.SuspendLayout()
        Me.gbNumber.SuspendLayout()
        Me.gbTargetSystem.SuspendLayout()
        Me.gbScoring.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbValueRange
        '
        resources.ApplyResources(Me.gbValueRange, "gbValueRange")
        Me.gbValueRange.Controls.Add(Me.tbMaxValue)
        Me.gbValueRange.Controls.Add(Me.tbMinValue)
        Me.gbValueRange.Controls.Add(Me.cmbOpMax)
        Me.gbValueRange.Controls.Add(Me.cmbOpMin)
        Me.gbValueRange.Controls.Add(Me.lblMaxValue)
        Me.gbValueRange.Controls.Add(Me.lblMinValue)
        Me.gbValueRange.Name = "gbValueRange"
        Me.gbValueRange.TabStop = False
        '
        'tbMaxValue
        '
        resources.ApplyResources(Me.tbMaxValue, "tbMaxValue")
        Me.tbMaxValue.Name = "tbMaxValue"
        '
        'tbMinValue
        '
        resources.ApplyResources(Me.tbMinValue, "tbMinValue")
        Me.tbMinValue.Name = "tbMinValue"
        '
        'cmbOpMax
        '
        resources.ApplyResources(Me.cmbOpMax, "cmbOpMax")
        Me.cmbOpMax.FormattingEnabled = True
        Me.cmbOpMax.Name = "cmbOpMax"
        '
        'cmbOpMin
        '
        resources.ApplyResources(Me.cmbOpMin, "cmbOpMin")
        Me.cmbOpMin.FormattingEnabled = True
        Me.cmbOpMin.Name = "cmbOpMin"
        '
        'lblMaxValue
        '
        resources.ApplyResources(Me.lblMaxValue, "lblMaxValue")
        Me.lblMaxValue.Name = "lblMaxValue"
        '
        'lblMinValue
        '
        resources.ApplyResources(Me.lblMinValue, "lblMinValue")
        Me.lblMinValue.Name = "lblMinValue"
        '
        'gbNumber
        '
        resources.ApplyResources(Me.gbNumber, "gbNumber")
        Me.gbNumber.Controls.Add(Me.tbNrDecimals)
        Me.gbNumber.Controls.Add(Me.lblUnit)
        Me.gbNumber.Controls.Add(Me.cmbUnit)
        Me.gbNumber.Controls.Add(Me.lblNrDecimals)
        Me.gbNumber.Name = "gbNumber"
        Me.gbNumber.TabStop = False
        '
        'tbNrDecimals
        '
        resources.ApplyResources(Me.tbNrDecimals, "tbNrDecimals")
        Me.tbNrDecimals.Name = "tbNrDecimals"
        '
        'lblUnit
        '
        resources.ApplyResources(Me.lblUnit, "lblUnit")
        Me.lblUnit.Name = "lblUnit"
        '
        'cmbUnit
        '
        resources.ApplyResources(Me.cmbUnit, "cmbUnit")
        Me.cmbUnit.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbUnit.FormattingEnabled = True
        Me.cmbUnit.Name = "cmbUnit"
        '
        'lblNrDecimals
        '
        resources.ApplyResources(Me.lblNrDecimals, "lblNrDecimals")
        Me.lblNrDecimals.Name = "lblNrDecimals"
        '
        'gbTargetSystem
        '
        resources.ApplyResources(Me.gbTargetSystem, "gbTargetSystem")
        Me.gbTargetSystem.Controls.Add(Me.rbtnTargetSystemFormula)
        Me.gbTargetSystem.Controls.Add(Me.rbtnTargetSystemValueRange)
        Me.gbTargetSystem.Controls.Add(Me.rbtnTargetSystemSimple)
        Me.gbTargetSystem.Name = "gbTargetSystem"
        Me.gbTargetSystem.TabStop = False
        '
        'rbtnTargetSystemFormula
        '
        resources.ApplyResources(Me.rbtnTargetSystemFormula, "rbtnTargetSystemFormula")
        Me.rbtnTargetSystemFormula.Name = "rbtnTargetSystemFormula"
        Me.rbtnTargetSystemFormula.TabStop = True
        Me.rbtnTargetSystemFormula.UseVisualStyleBackColor = True
        '
        'rbtnTargetSystemValueRange
        '
        resources.ApplyResources(Me.rbtnTargetSystemValueRange, "rbtnTargetSystemValueRange")
        Me.rbtnTargetSystemValueRange.Name = "rbtnTargetSystemValueRange"
        Me.rbtnTargetSystemValueRange.TabStop = True
        Me.rbtnTargetSystemValueRange.UseVisualStyleBackColor = True
        '
        'rbtnTargetSystemSimple
        '
        resources.ApplyResources(Me.rbtnTargetSystemSimple, "rbtnTargetSystemSimple")
        Me.rbtnTargetSystemSimple.Name = "rbtnTargetSystemSimple"
        Me.rbtnTargetSystemSimple.TabStop = True
        Me.rbtnTargetSystemSimple.UseVisualStyleBackColor = True
        '
        'rbtnScoringValue
        '
        resources.ApplyResources(Me.rbtnScoringValue, "rbtnScoringValue")
        Me.rbtnScoringValue.Name = "rbtnScoringValue"
        Me.rbtnScoringValue.TabStop = True
        Me.rbtnScoringValue.UseVisualStyleBackColor = True
        '
        'rbtnScoringPercentage
        '
        resources.ApplyResources(Me.rbtnScoringPercentage, "rbtnScoringPercentage")
        Me.rbtnScoringPercentage.Name = "rbtnScoringPercentage"
        Me.rbtnScoringPercentage.TabStop = True
        Me.rbtnScoringPercentage.UseVisualStyleBackColor = True
        '
        'rbtnScoringScore
        '
        resources.ApplyResources(Me.rbtnScoringScore, "rbtnScoringScore")
        Me.rbtnScoringScore.Name = "rbtnScoringScore"
        Me.rbtnScoringScore.TabStop = True
        Me.rbtnScoringScore.UseVisualStyleBackColor = True
        '
        'gbScoring
        '
        resources.ApplyResources(Me.gbScoring, "gbScoring")
        Me.gbScoring.Controls.Add(Me.rbtnScoringScore)
        Me.gbScoring.Controls.Add(Me.rbtnScoringPercentage)
        Me.gbScoring.Controls.Add(Me.rbtnScoringValue)
        Me.gbScoring.Name = "gbScoring"
        Me.gbScoring.TabStop = False
        '
        'IndicatorValues
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbTargetSystem)
        Me.Controls.Add(Me.gbScoring)
        Me.Controls.Add(Me.gbNumber)
        Me.Controls.Add(Me.gbValueRange)
        Me.Name = "IndicatorValues"
        Me.gbValueRange.ResumeLayout(False)
        Me.gbValueRange.PerformLayout()
        Me.gbNumber.ResumeLayout(False)
        Me.gbNumber.PerformLayout()
        Me.gbTargetSystem.ResumeLayout(False)
        Me.gbTargetSystem.PerformLayout()
        Me.gbScoring.ResumeLayout(False)
        Me.gbScoring.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbValueRange As System.Windows.Forms.GroupBox
    Friend WithEvents lblMinValue As System.Windows.Forms.Label
    Friend WithEvents lblMaxValue As System.Windows.Forms.Label
    Friend WithEvents gbNumber As System.Windows.Forms.GroupBox
    Friend WithEvents lblUnit As System.Windows.Forms.Label
    Friend WithEvents cmbUnit As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents lblNrDecimals As System.Windows.Forms.Label
    Friend WithEvents gbTargetSystem As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnTargetSystemFormula As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnTargetSystemValueRange As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnTargetSystemSimple As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringValue As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringScore As System.Windows.Forms.RadioButton
    Friend WithEvents gbScoring As System.Windows.Forms.GroupBox
    Friend WithEvents cmbOpMax As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbOpMin As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents tbNrDecimals As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents tbMaxValue As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents tbMinValue As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
