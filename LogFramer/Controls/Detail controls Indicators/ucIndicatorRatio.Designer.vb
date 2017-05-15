<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorRatio
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorRatio))
        Me.gbValueRangeFirstStatement = New System.Windows.Forms.GroupBox()
        Me.ntbMaxValue1 = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.ntbMinValue1 = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.cmbOpMax1 = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbOpMin1 = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.lblMaxValue1 = New System.Windows.Forms.Label()
        Me.lblMinValue1 = New System.Windows.Forms.Label()
        Me.gbNumberFirstStatement = New System.Windows.Forms.GroupBox()
        Me.ntbNrDecimals1 = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblUnit1 = New System.Windows.Forms.Label()
        Me.cmbUnit1 = New FaciliDev.LogFramer.StructuredComboBox()
        Me.lblNrDecimals1 = New System.Windows.Forms.Label()
        Me.gbTargetSystem = New System.Windows.Forms.GroupBox()
        Me.rbtnTargetSystemFormula = New System.Windows.Forms.RadioButton()
        Me.rbtnTargetSystemValueRange = New System.Windows.Forms.RadioButton()
        Me.rbtnTargetSystemSimple = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringValue = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringPercentage = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringScore = New System.Windows.Forms.RadioButton()
        Me.gbScoring = New System.Windows.Forms.GroupBox()
        Me.gbFirstStatement = New System.Windows.Forms.GroupBox()
        Me.rtbFirstStatement = New System.Windows.Forms.RichTextBox()
        Me.gbSecondStatement = New System.Windows.Forms.GroupBox()
        Me.rtbSecondStatement = New System.Windows.Forms.RichTextBox()
        Me.gbNumberSecondStatement = New System.Windows.Forms.GroupBox()
        Me.ntbNrDecimals2 = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblUnit2 = New System.Windows.Forms.Label()
        Me.cmbUnit2 = New FaciliDev.LogFramer.StructuredComboBox()
        Me.lblNrDecimals2 = New System.Windows.Forms.Label()
        Me.gbValueRangeSecondStatement = New System.Windows.Forms.GroupBox()
        Me.ntbMaxValue2 = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.ntbMinValue2 = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.cmbOpMax2 = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbOpMin2 = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.lblMaxValue2 = New System.Windows.Forms.Label()
        Me.lblMinValue2 = New System.Windows.Forms.Label()
        Me.gbValueRangeFirstStatement.SuspendLayout()
        Me.gbNumberFirstStatement.SuspendLayout()
        Me.gbTargetSystem.SuspendLayout()
        Me.gbScoring.SuspendLayout()
        Me.gbFirstStatement.SuspendLayout()
        Me.gbSecondStatement.SuspendLayout()
        Me.gbNumberSecondStatement.SuspendLayout()
        Me.gbValueRangeSecondStatement.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbValueRangeFirstStatement
        '
        resources.ApplyResources(Me.gbValueRangeFirstStatement, "gbValueRangeFirstStatement")
        Me.gbValueRangeFirstStatement.Controls.Add(Me.ntbMaxValue1)
        Me.gbValueRangeFirstStatement.Controls.Add(Me.ntbMinValue1)
        Me.gbValueRangeFirstStatement.Controls.Add(Me.cmbOpMax1)
        Me.gbValueRangeFirstStatement.Controls.Add(Me.cmbOpMin1)
        Me.gbValueRangeFirstStatement.Controls.Add(Me.lblMaxValue1)
        Me.gbValueRangeFirstStatement.Controls.Add(Me.lblMinValue1)
        Me.gbValueRangeFirstStatement.Name = "gbValueRangeFirstStatement"
        Me.gbValueRangeFirstStatement.TabStop = False
        '
        'ntbMaxValue1
        '
        resources.ApplyResources(Me.ntbMaxValue1, "ntbMaxValue1")
        Me.ntbMaxValue1.Name = "ntbMaxValue1"
        '
        'ntbMinValue1
        '
        resources.ApplyResources(Me.ntbMinValue1, "ntbMinValue1")
        Me.ntbMinValue1.Name = "ntbMinValue1"
        '
        'cmbOpMax1
        '
        resources.ApplyResources(Me.cmbOpMax1, "cmbOpMax1")
        Me.cmbOpMax1.FormattingEnabled = True
        Me.cmbOpMax1.Name = "cmbOpMax1"
        '
        'cmbOpMin1
        '
        resources.ApplyResources(Me.cmbOpMin1, "cmbOpMin1")
        Me.cmbOpMin1.FormattingEnabled = True
        Me.cmbOpMin1.Name = "cmbOpMin1"
        '
        'lblMaxValue1
        '
        resources.ApplyResources(Me.lblMaxValue1, "lblMaxValue1")
        Me.lblMaxValue1.Name = "lblMaxValue1"
        '
        'lblMinValue1
        '
        resources.ApplyResources(Me.lblMinValue1, "lblMinValue1")
        Me.lblMinValue1.Name = "lblMinValue1"
        '
        'gbNumberFirstStatement
        '
        resources.ApplyResources(Me.gbNumberFirstStatement, "gbNumberFirstStatement")
        Me.gbNumberFirstStatement.Controls.Add(Me.ntbNrDecimals1)
        Me.gbNumberFirstStatement.Controls.Add(Me.lblUnit1)
        Me.gbNumberFirstStatement.Controls.Add(Me.cmbUnit1)
        Me.gbNumberFirstStatement.Controls.Add(Me.lblNrDecimals1)
        Me.gbNumberFirstStatement.Name = "gbNumberFirstStatement"
        Me.gbNumberFirstStatement.TabStop = False
        '
        'ntbNrDecimals1
        '
        resources.ApplyResources(Me.ntbNrDecimals1, "ntbNrDecimals1")
        Me.ntbNrDecimals1.Name = "ntbNrDecimals1"
        '
        'lblUnit1
        '
        resources.ApplyResources(Me.lblUnit1, "lblUnit1")
        Me.lblUnit1.Name = "lblUnit1"
        '
        'cmbUnit1
        '
        resources.ApplyResources(Me.cmbUnit1, "cmbUnit1")
        Me.cmbUnit1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbUnit1.FormattingEnabled = True
        Me.cmbUnit1.Name = "cmbUnit1"
        '
        'lblNrDecimals1
        '
        resources.ApplyResources(Me.lblNrDecimals1, "lblNrDecimals1")
        Me.lblNrDecimals1.Name = "lblNrDecimals1"
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
        'gbFirstStatement
        '
        resources.ApplyResources(Me.gbFirstStatement, "gbFirstStatement")
        Me.gbFirstStatement.Controls.Add(Me.rtbFirstStatement)
        Me.gbFirstStatement.Name = "gbFirstStatement"
        Me.gbFirstStatement.TabStop = False
        '
        'rtbFirstStatement
        '
        resources.ApplyResources(Me.rtbFirstStatement, "rtbFirstStatement")
        Me.rtbFirstStatement.Name = "rtbFirstStatement"
        '
        'gbSecondStatement
        '
        resources.ApplyResources(Me.gbSecondStatement, "gbSecondStatement")
        Me.gbSecondStatement.Controls.Add(Me.rtbSecondStatement)
        Me.gbSecondStatement.Name = "gbSecondStatement"
        Me.gbSecondStatement.TabStop = False
        '
        'rtbSecondStatement
        '
        resources.ApplyResources(Me.rtbSecondStatement, "rtbSecondStatement")
        Me.rtbSecondStatement.Name = "rtbSecondStatement"
        '
        'gbNumberSecondStatement
        '
        resources.ApplyResources(Me.gbNumberSecondStatement, "gbNumberSecondStatement")
        Me.gbNumberSecondStatement.Controls.Add(Me.ntbNrDecimals2)
        Me.gbNumberSecondStatement.Controls.Add(Me.lblUnit2)
        Me.gbNumberSecondStatement.Controls.Add(Me.cmbUnit2)
        Me.gbNumberSecondStatement.Controls.Add(Me.lblNrDecimals2)
        Me.gbNumberSecondStatement.Name = "gbNumberSecondStatement"
        Me.gbNumberSecondStatement.TabStop = False
        '
        'ntbNrDecimals2
        '
        resources.ApplyResources(Me.ntbNrDecimals2, "ntbNrDecimals2")
        Me.ntbNrDecimals2.Name = "ntbNrDecimals2"
        '
        'lblUnit2
        '
        resources.ApplyResources(Me.lblUnit2, "lblUnit2")
        Me.lblUnit2.Name = "lblUnit2"
        '
        'cmbUnit2
        '
        resources.ApplyResources(Me.cmbUnit2, "cmbUnit2")
        Me.cmbUnit2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbUnit2.FormattingEnabled = True
        Me.cmbUnit2.Name = "cmbUnit2"
        '
        'lblNrDecimals2
        '
        resources.ApplyResources(Me.lblNrDecimals2, "lblNrDecimals2")
        Me.lblNrDecimals2.Name = "lblNrDecimals2"
        '
        'gbValueRangeSecondStatement
        '
        resources.ApplyResources(Me.gbValueRangeSecondStatement, "gbValueRangeSecondStatement")
        Me.gbValueRangeSecondStatement.Controls.Add(Me.ntbMaxValue2)
        Me.gbValueRangeSecondStatement.Controls.Add(Me.ntbMinValue2)
        Me.gbValueRangeSecondStatement.Controls.Add(Me.cmbOpMax2)
        Me.gbValueRangeSecondStatement.Controls.Add(Me.cmbOpMin2)
        Me.gbValueRangeSecondStatement.Controls.Add(Me.lblMaxValue2)
        Me.gbValueRangeSecondStatement.Controls.Add(Me.lblMinValue2)
        Me.gbValueRangeSecondStatement.Name = "gbValueRangeSecondStatement"
        Me.gbValueRangeSecondStatement.TabStop = False
        '
        'ntbMaxValue2
        '
        resources.ApplyResources(Me.ntbMaxValue2, "ntbMaxValue2")
        Me.ntbMaxValue2.Name = "ntbMaxValue2"
        '
        'ntbMinValue2
        '
        resources.ApplyResources(Me.ntbMinValue2, "ntbMinValue2")
        Me.ntbMinValue2.Name = "ntbMinValue2"
        '
        'cmbOpMax2
        '
        resources.ApplyResources(Me.cmbOpMax2, "cmbOpMax2")
        Me.cmbOpMax2.FormattingEnabled = True
        Me.cmbOpMax2.Name = "cmbOpMax2"
        '
        'cmbOpMin2
        '
        resources.ApplyResources(Me.cmbOpMin2, "cmbOpMin2")
        Me.cmbOpMin2.FormattingEnabled = True
        Me.cmbOpMin2.Name = "cmbOpMin2"
        '
        'lblMaxValue2
        '
        resources.ApplyResources(Me.lblMaxValue2, "lblMaxValue2")
        Me.lblMaxValue2.Name = "lblMaxValue2"
        '
        'lblMinValue2
        '
        resources.ApplyResources(Me.lblMinValue2, "lblMinValue2")
        Me.lblMinValue2.Name = "lblMinValue2"
        '
        'IndicatorRatio
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbSecondStatement)
        Me.Controls.Add(Me.gbNumberSecondStatement)
        Me.Controls.Add(Me.gbValueRangeSecondStatement)
        Me.Controls.Add(Me.gbFirstStatement)
        Me.Controls.Add(Me.gbTargetSystem)
        Me.Controls.Add(Me.gbScoring)
        Me.Controls.Add(Me.gbNumberFirstStatement)
        Me.Controls.Add(Me.gbValueRangeFirstStatement)
        Me.Name = "IndicatorRatio"
        Me.gbValueRangeFirstStatement.ResumeLayout(False)
        Me.gbValueRangeFirstStatement.PerformLayout()
        Me.gbNumberFirstStatement.ResumeLayout(False)
        Me.gbNumberFirstStatement.PerformLayout()
        Me.gbTargetSystem.ResumeLayout(False)
        Me.gbTargetSystem.PerformLayout()
        Me.gbScoring.ResumeLayout(False)
        Me.gbScoring.PerformLayout()
        Me.gbFirstStatement.ResumeLayout(False)
        Me.gbSecondStatement.ResumeLayout(False)
        Me.gbNumberSecondStatement.ResumeLayout(False)
        Me.gbNumberSecondStatement.PerformLayout()
        Me.gbValueRangeSecondStatement.ResumeLayout(False)
        Me.gbValueRangeSecondStatement.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbValueRangeFirstStatement As System.Windows.Forms.GroupBox
    Friend WithEvents lblMinValue1 As System.Windows.Forms.Label
    Friend WithEvents lblMaxValue1 As System.Windows.Forms.Label
    Friend WithEvents gbNumberFirstStatement As System.Windows.Forms.GroupBox
    Friend WithEvents lblUnit1 As System.Windows.Forms.Label
    Friend WithEvents cmbUnit1 As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents lblNrDecimals1 As System.Windows.Forms.Label
    Friend WithEvents gbTargetSystem As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnTargetSystemFormula As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnTargetSystemValueRange As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnTargetSystemSimple As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringValue As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringScore As System.Windows.Forms.RadioButton
    Friend WithEvents gbScoring As System.Windows.Forms.GroupBox
    Friend WithEvents gbFirstStatement As System.Windows.Forms.GroupBox
    Friend WithEvents rtbFirstStatement As System.Windows.Forms.RichTextBox
    Friend WithEvents gbSecondStatement As System.Windows.Forms.GroupBox
    Friend WithEvents rtbSecondStatement As System.Windows.Forms.RichTextBox
    Friend WithEvents gbNumberSecondStatement As System.Windows.Forms.GroupBox
    Friend WithEvents lblUnit2 As System.Windows.Forms.Label
    Friend WithEvents cmbUnit2 As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents lblNrDecimals2 As System.Windows.Forms.Label
    Friend WithEvents gbValueRangeSecondStatement As System.Windows.Forms.GroupBox
    Friend WithEvents lblMaxValue2 As System.Windows.Forms.Label
    Friend WithEvents lblMinValue2 As System.Windows.Forms.Label
    Friend WithEvents cmbOpMin1 As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbOpMax1 As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbOpMax2 As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbOpMin2 As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents ntbNrDecimals1 As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents ntbNrDecimals2 As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents ntbMaxValue1 As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents ntbMinValue1 As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents ntbMaxValue2 As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents ntbMinValue2 As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
