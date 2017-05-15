<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorScales
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorScales))
        Me.gbScoreValues = New System.Windows.Forms.GroupBox()
        Me.gbScoring = New System.Windows.Forms.GroupBox()
        Me.rbtnScoringScore = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringPercentage = New System.Windows.Forms.RadioButton()
        Me.gbAgreement = New System.Windows.Forms.GroupBox()
        Me.tbDisagreeText = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblDisagreeText = New System.Windows.Forms.Label()
        Me.tbAgreeText = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblAgreeText = New System.Windows.Forms.Label()
        Me.gbNumber = New System.Windows.Forms.GroupBox()
        Me.ntbNrDecimals = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblNrDecimals = New System.Windows.Forms.Label()
        Me.gbScoring.SuspendLayout()
        Me.gbAgreement.SuspendLayout()
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
        'gbAgreement
        '
        resources.ApplyResources(Me.gbAgreement, "gbAgreement")
        Me.gbAgreement.Controls.Add(Me.tbDisagreeText)
        Me.gbAgreement.Controls.Add(Me.lblDisagreeText)
        Me.gbAgreement.Controls.Add(Me.tbAgreeText)
        Me.gbAgreement.Controls.Add(Me.lblAgreeText)
        Me.gbAgreement.Name = "gbAgreement"
        Me.gbAgreement.TabStop = False
        '
        'tbDisagreeText
        '
        resources.ApplyResources(Me.tbDisagreeText, "tbDisagreeText")
        Me.tbDisagreeText.Name = "tbDisagreeText"
        '
        'lblDisagreeText
        '
        resources.ApplyResources(Me.lblDisagreeText, "lblDisagreeText")
        Me.lblDisagreeText.Name = "lblDisagreeText"
        '
        'tbAgreeText
        '
        resources.ApplyResources(Me.tbAgreeText, "tbAgreeText")
        Me.tbAgreeText.Name = "tbAgreeText"
        '
        'lblAgreeText
        '
        resources.ApplyResources(Me.lblAgreeText, "lblAgreeText")
        Me.lblAgreeText.Name = "lblAgreeText"
        '
        'gbNumber
        '
        resources.ApplyResources(Me.gbNumber, "gbNumber")
        Me.gbNumber.Controls.Add(Me.ntbNrDecimals)
        Me.gbNumber.Controls.Add(Me.lblNrDecimals)
        Me.gbNumber.Name = "gbNumber"
        Me.gbNumber.TabStop = False
        '
        'ntbNrDecimals
        '
        resources.ApplyResources(Me.ntbNrDecimals, "ntbNrDecimals")
        Me.ntbNrDecimals.Name = "ntbNrDecimals"
        '
        'lblNrDecimals
        '
        resources.ApplyResources(Me.lblNrDecimals, "lblNrDecimals")
        Me.lblNrDecimals.Name = "lblNrDecimals"
        '
        'IndicatorScales
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbNumber)
        Me.Controls.Add(Me.gbAgreement)
        Me.Controls.Add(Me.gbScoring)
        Me.Controls.Add(Me.gbScoreValues)
        Me.Name = "IndicatorScales"
        Me.gbScoring.ResumeLayout(False)
        Me.gbScoring.PerformLayout()
        Me.gbAgreement.ResumeLayout(False)
        Me.gbAgreement.PerformLayout()
        Me.gbNumber.ResumeLayout(False)
        Me.gbNumber.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbScoreValues As System.Windows.Forms.GroupBox
    Friend WithEvents gbScoring As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnScoringScore As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents gbAgreement As System.Windows.Forms.GroupBox
    Friend WithEvents tbDisagreeText As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblDisagreeText As System.Windows.Forms.Label
    Friend WithEvents tbAgreeText As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblAgreeText As System.Windows.Forms.Label
    Friend WithEvents gbNumber As System.Windows.Forms.GroupBox
    Friend WithEvents lblNrDecimals As System.Windows.Forms.Label
    Friend WithEvents ntbNrDecimals As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
