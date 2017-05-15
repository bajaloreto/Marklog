<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorYesNo
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorYesNo))
        Me.lblNo = New System.Windows.Forms.Label()
        Me.lblYes = New System.Windows.Forms.Label()
        Me.gbScoreValues = New System.Windows.Forms.GroupBox()
        Me.ntbNo = New FaciliDev.LogFramer.NumericTextBox()
        Me.ntbYes = New FaciliDev.LogFramer.NumericTextBox()
        Me.gbScoring = New System.Windows.Forms.GroupBox()
        Me.rbtnScoringScore = New System.Windows.Forms.RadioButton()
        Me.rbtnScoringPercentage = New System.Windows.Forms.RadioButton()
        Me.gbNumber = New System.Windows.Forms.GroupBox()
        Me.ntbNrDecimals = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblNrDecimals = New System.Windows.Forms.Label()
        Me.gbScoreValues.SuspendLayout()
        Me.gbScoring.SuspendLayout()
        Me.gbNumber.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblNo
        '
        resources.ApplyResources(Me.lblNo, "lblNo")
        Me.lblNo.Name = "lblNo"
        '
        'lblYes
        '
        resources.ApplyResources(Me.lblYes, "lblYes")
        Me.lblYes.Name = "lblYes"
        '
        'gbScoreValues
        '
        resources.ApplyResources(Me.gbScoreValues, "gbScoreValues")
        Me.gbScoreValues.Controls.Add(Me.ntbNo)
        Me.gbScoreValues.Controls.Add(Me.ntbYes)
        Me.gbScoreValues.Controls.Add(Me.lblNo)
        Me.gbScoreValues.Controls.Add(Me.lblYes)
        Me.gbScoreValues.Name = "gbScoreValues"
        Me.gbScoreValues.TabStop = False
        '
        'ntbNo
        '
        resources.ApplyResources(Me.ntbNo, "ntbNo")
        Me.ntbNo.AllowSpace = True
        Me.ntbNo.DoubleValue = 0.0R
        Me.ntbNo.IntegerValue = 0
        Me.ntbNo.IsCurrency = False
        Me.ntbNo.IsPercentage = False
        Me.ntbNo.Name = "ntbNo"
        Me.ntbNo.NrDecimals = 0
        Me.ntbNo.SetDecimals = False
        Me.ntbNo.SingleValue = 0.0!
        Me.ntbNo.Unit = Nothing
        Me.ntbNo.ValueType = 0
        '
        'ntbYes
        '
        resources.ApplyResources(Me.ntbYes, "ntbYes")
        Me.ntbYes.AllowSpace = True
        Me.ntbYes.DoubleValue = 0.0R
        Me.ntbYes.IntegerValue = 0
        Me.ntbYes.IsCurrency = False
        Me.ntbYes.IsPercentage = False
        Me.ntbYes.Name = "ntbYes"
        Me.ntbYes.NrDecimals = 0
        Me.ntbYes.SetDecimals = False
        Me.ntbYes.SingleValue = 0.0!
        Me.ntbYes.Unit = Nothing
        Me.ntbYes.ValueType = 0
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
        'IndicatorYesNo
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbNumber)
        Me.Controls.Add(Me.gbScoring)
        Me.Controls.Add(Me.gbScoreValues)
        Me.Name = "IndicatorYesNo"
        Me.gbScoreValues.ResumeLayout(False)
        Me.gbScoreValues.PerformLayout()
        Me.gbScoring.ResumeLayout(False)
        Me.gbScoring.PerformLayout()
        Me.gbNumber.ResumeLayout(False)
        Me.gbNumber.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblYes As System.Windows.Forms.Label
    Friend WithEvents ntbNo As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents lblNo As System.Windows.Forms.Label
    Friend WithEvents ntbYes As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents gbScoreValues As System.Windows.Forms.GroupBox
    Friend WithEvents gbScoring As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnScoringScore As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnScoringPercentage As System.Windows.Forms.RadioButton
    Friend WithEvents gbNumber As System.Windows.Forms.GroupBox
    Friend WithEvents lblNrDecimals As System.Windows.Forms.Label
    Friend WithEvents ntbNrDecimals As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
