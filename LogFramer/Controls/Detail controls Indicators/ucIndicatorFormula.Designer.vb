<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorFormula
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorFormula))
        Me.gbTargetSystem = New System.Windows.Forms.GroupBox()
        Me.rbtnTargetSystemSimple = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringValue = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringPercentage = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringScore = New System.Windows.Forms.RadioButton()
        Me.gbScoring = New System.Windows.Forms.GroupBox()
        Me.gbQuestions = New System.Windows.Forms.GroupBox()
        Me.gbFormula = New System.Windows.Forms.GroupBox()
        Me.lblFormula = New System.Windows.Forms.Label()
        Me.tbFormula = New FaciliDev.LogFramer.TextBoxLF()
        Me.gbNumber = New System.Windows.Forms.GroupBox()
        Me.ntbNrDecimals = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblUnit = New System.Windows.Forms.Label()
        Me.cmbUnit = New FaciliDev.LogFramer.StructuredComboBox()
        Me.lblNrDecimals = New System.Windows.Forms.Label()
        Me.gbTargetSystem.SuspendLayout()
        Me.gbScoring.SuspendLayout()
        Me.gbFormula.SuspendLayout()
        Me.gbNumber.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbTargetSystem
        '
        resources.ApplyResources(Me.gbTargetSystem, "gbTargetSystem")
        Me.gbTargetSystem.Controls.Add(Me.rbtnTargetSystemSimple)
        Me.gbTargetSystem.Name = "gbTargetSystem"
        Me.gbTargetSystem.TabStop = False
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
        'gbQuestions
        '
        resources.ApplyResources(Me.gbQuestions, "gbQuestions")
        Me.gbQuestions.Name = "gbQuestions"
        Me.gbQuestions.TabStop = False
        '
        'gbFormula
        '
        resources.ApplyResources(Me.gbFormula, "gbFormula")
        Me.gbFormula.Controls.Add(Me.lblFormula)
        Me.gbFormula.Controls.Add(Me.tbFormula)
        Me.gbFormula.Name = "gbFormula"
        Me.gbFormula.TabStop = False
        '
        'lblFormula
        '
        resources.ApplyResources(Me.lblFormula, "lblFormula")
        Me.lblFormula.Name = "lblFormula"
        '
        'tbFormula
        '
        resources.ApplyResources(Me.tbFormula, "tbFormula")
        Me.tbFormula.Name = "tbFormula"
        '
        'gbNumber
        '
        resources.ApplyResources(Me.gbNumber, "gbNumber")
        Me.gbNumber.Controls.Add(Me.ntbNrDecimals)
        Me.gbNumber.Controls.Add(Me.lblUnit)
        Me.gbNumber.Controls.Add(Me.cmbUnit)
        Me.gbNumber.Controls.Add(Me.lblNrDecimals)
        Me.gbNumber.Name = "gbNumber"
        Me.gbNumber.TabStop = False
        '
        'ntbNrDecimals
        '
        resources.ApplyResources(Me.ntbNrDecimals, "ntbNrDecimals")
        Me.ntbNrDecimals.Name = "ntbNrDecimals"
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
        'IndicatorFormula
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbNumber)
        Me.Controls.Add(Me.gbFormula)
        Me.Controls.Add(Me.gbQuestions)
        Me.Controls.Add(Me.gbTargetSystem)
        Me.Controls.Add(Me.gbScoring)
        Me.Name = "IndicatorFormula"
        Me.gbTargetSystem.ResumeLayout(False)
        Me.gbTargetSystem.PerformLayout()
        Me.gbScoring.ResumeLayout(False)
        Me.gbScoring.PerformLayout()
        Me.gbFormula.ResumeLayout(False)
        Me.gbFormula.PerformLayout()
        Me.gbNumber.ResumeLayout(False)
        Me.gbNumber.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbTargetSystem As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnTargetSystemSimple As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringValue As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringScore As System.Windows.Forms.RadioButton
    Friend WithEvents gbScoring As System.Windows.Forms.GroupBox
    Friend WithEvents gbQuestions As System.Windows.Forms.GroupBox
    Friend WithEvents gbFormula As System.Windows.Forms.GroupBox
    Friend WithEvents lblFormula As System.Windows.Forms.Label
    Friend WithEvents tbFormula As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents gbNumber As System.Windows.Forms.GroupBox
    Friend WithEvents lblUnit As System.Windows.Forms.Label
    Friend WithEvents cmbUnit As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents lblNrDecimals As System.Windows.Forms.Label
    Friend WithEvents ntbNrDecimals As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
