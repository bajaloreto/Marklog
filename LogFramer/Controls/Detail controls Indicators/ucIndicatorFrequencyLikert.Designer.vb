<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorFrequencyLikert
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorFrequencyLikert))
        Me.gbScoreValues = New System.Windows.Forms.GroupBox()
        Me.gbScoring = New System.Windows.Forms.GroupBox()
        Me.rbtnScoringScore = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringPercentage = New System.Windows.Forms.RadioButton()
        Me.gbNumber = New System.Windows.Forms.GroupBox()
        Me.ntbNrDecimals = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblUnit = New System.Windows.Forms.Label()
        Me.cmbUnit = New FaciliDev.LogFramer.StructuredComboBox()
        Me.lblNrDecimals = New System.Windows.Forms.Label()
        Me.gbScoring.SuspendLayout()
        Me.gbNumber.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbScoreValues
        '
        resources.ApplyResources(Me.gbScoreValues, "gbScoreValues")
        Me.gbScoreValues.Name = "gbScoreValues"
        Me.gbScoreValues.TabStop = False
        '
        'gbScoring
        '
        resources.ApplyResources(Me.gbScoring, "gbScoring")
        Me.gbScoring.Controls.Add(Me.rbtnScoringScore)
        Me.gbScoring.Controls.Add(Me.rbtnScoringPercentage)
        Me.gbScoring.Name = "gbScoring"
        Me.gbScoring.TabStop = False
        '
        'rbtnScoringScore
        '
        resources.ApplyResources(Me.rbtnScoringScore, "rbtnScoringScore")
        Me.rbtnScoringScore.Name = "rbtnScoringScore"
        Me.rbtnScoringScore.TabStop = True
        Me.rbtnScoringScore.UseVisualStyleBackColor = True
        '
        'rbtnScoringPercentage
        '
        resources.ApplyResources(Me.rbtnScoringPercentage, "rbtnScoringPercentage")
        Me.rbtnScoringPercentage.Name = "rbtnScoringPercentage"
        Me.rbtnScoringPercentage.TabStop = True
        Me.rbtnScoringPercentage.UseVisualStyleBackColor = True
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
        'IndicatorFrequencyLikert
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbNumber)
        Me.Controls.Add(Me.gbScoring)
        Me.Controls.Add(Me.gbScoreValues)
        Me.Name = "IndicatorFrequencyLikert"
        Me.gbScoring.ResumeLayout(False)
        Me.gbScoring.PerformLayout()
        Me.gbNumber.ResumeLayout(False)
        Me.gbNumber.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbScoreValues As System.Windows.Forms.GroupBox
    Friend WithEvents gbScoring As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnScoringScore As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents gbNumber As System.Windows.Forms.GroupBox
    Friend WithEvents ntbNrDecimals As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents lblUnit As System.Windows.Forms.Label
    Friend WithEvents cmbUnit As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents lblNrDecimals As System.Windows.Forms.Label

End Class
