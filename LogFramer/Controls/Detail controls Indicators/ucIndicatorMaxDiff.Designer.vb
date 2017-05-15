<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorMaxDiff
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorMaxDiff))
        Me.gbStatements = New System.Windows.Forms.GroupBox()
        Me.gbAgreement = New System.Windows.Forms.GroupBox()
        Me.tbDisagreeText = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblWorstOption = New System.Windows.Forms.Label()
        Me.tbAgreeText = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblBestOption = New System.Windows.Forms.Label()
        Me.gbScoreValues = New System.Windows.Forms.GroupBox()
        Me.ntbNotSelectedScore = New FaciliDev.LogFramer.NumericTextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ntbWorstOptionScore = New FaciliDev.LogFramer.NumericTextBox()
        Me.ntbBestOptionScore = New FaciliDev.LogFramer.NumericTextBox()
        Me.lblNo = New System.Windows.Forms.Label()
        Me.lblBestValueScore = New System.Windows.Forms.Label()
        Me.gbAgreement.SuspendLayout()
        Me.gbScoreValues.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbStatements
        '
        resources.ApplyResources(Me.gbStatements, "gbStatements")
        Me.gbStatements.Name = "gbStatements"
        Me.gbStatements.TabStop = False
        '
        'gbAgreement
        '
        resources.ApplyResources(Me.gbAgreement, "gbAgreement")
        Me.gbAgreement.Controls.Add(Me.tbDisagreeText)
        Me.gbAgreement.Controls.Add(Me.lblWorstOption)
        Me.gbAgreement.Controls.Add(Me.tbAgreeText)
        Me.gbAgreement.Controls.Add(Me.lblBestOption)
        Me.gbAgreement.Name = "gbAgreement"
        Me.gbAgreement.TabStop = False
        '
        'tbDisagreeText
        '
        resources.ApplyResources(Me.tbDisagreeText, "tbDisagreeText")
        Me.tbDisagreeText.Name = "tbDisagreeText"
        '
        'lblWorstOption
        '
        resources.ApplyResources(Me.lblWorstOption, "lblWorstOption")
        Me.lblWorstOption.Name = "lblWorstOption"
        '
        'tbAgreeText
        '
        resources.ApplyResources(Me.tbAgreeText, "tbAgreeText")
        Me.tbAgreeText.Name = "tbAgreeText"
        '
        'lblBestOption
        '
        resources.ApplyResources(Me.lblBestOption, "lblBestOption")
        Me.lblBestOption.Name = "lblBestOption"
        '
        'gbScoreValues
        '
        resources.ApplyResources(Me.gbScoreValues, "gbScoreValues")
        Me.gbScoreValues.Controls.Add(Me.ntbNotSelectedScore)
        Me.gbScoreValues.Controls.Add(Me.Label1)
        Me.gbScoreValues.Controls.Add(Me.ntbWorstOptionScore)
        Me.gbScoreValues.Controls.Add(Me.ntbBestOptionScore)
        Me.gbScoreValues.Controls.Add(Me.lblNo)
        Me.gbScoreValues.Controls.Add(Me.lblBestValueScore)
        Me.gbScoreValues.Name = "gbScoreValues"
        Me.gbScoreValues.TabStop = False
        '
        'ntbNotSelectedScore
        '
        resources.ApplyResources(Me.ntbNotSelectedScore, "ntbNotSelectedScore")
        Me.ntbNotSelectedScore.AllowSpace = True
        Me.ntbNotSelectedScore.DoubleValue = 0.0R
        Me.ntbNotSelectedScore.IntegerValue = 0
        Me.ntbNotSelectedScore.IsCurrency = False
        Me.ntbNotSelectedScore.IsPercentage = False
        Me.ntbNotSelectedScore.Name = "ntbNotSelectedScore"
        Me.ntbNotSelectedScore.NrDecimals = 0
        Me.ntbNotSelectedScore.ReadOnly = True
        Me.ntbNotSelectedScore.SetDecimals = True
        Me.ntbNotSelectedScore.SingleValue = 0.0!
        Me.ntbNotSelectedScore.Unit = Nothing
        Me.ntbNotSelectedScore.ValueType = 0
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'ntbWorstOptionScore
        '
        resources.ApplyResources(Me.ntbWorstOptionScore, "ntbWorstOptionScore")
        Me.ntbWorstOptionScore.AllowSpace = True
        Me.ntbWorstOptionScore.DoubleValue = 0.0R
        Me.ntbWorstOptionScore.IntegerValue = 0
        Me.ntbWorstOptionScore.IsCurrency = False
        Me.ntbWorstOptionScore.IsPercentage = False
        Me.ntbWorstOptionScore.Name = "ntbWorstOptionScore"
        Me.ntbWorstOptionScore.NrDecimals = 0
        Me.ntbWorstOptionScore.ReadOnly = True
        Me.ntbWorstOptionScore.SetDecimals = True
        Me.ntbWorstOptionScore.SingleValue = 0.0!
        Me.ntbWorstOptionScore.Unit = Nothing
        Me.ntbWorstOptionScore.ValueType = 0
        '
        'ntbBestOptionScore
        '
        resources.ApplyResources(Me.ntbBestOptionScore, "ntbBestOptionScore")
        Me.ntbBestOptionScore.AllowSpace = True
        Me.ntbBestOptionScore.DoubleValue = 0.0R
        Me.ntbBestOptionScore.IntegerValue = 0
        Me.ntbBestOptionScore.IsCurrency = False
        Me.ntbBestOptionScore.IsPercentage = False
        Me.ntbBestOptionScore.Name = "ntbBestOptionScore"
        Me.ntbBestOptionScore.NrDecimals = 0
        Me.ntbBestOptionScore.SetDecimals = True
        Me.ntbBestOptionScore.SingleValue = 0.0!
        Me.ntbBestOptionScore.Unit = Nothing
        Me.ntbBestOptionScore.ValueType = 0
        '
        'lblNo
        '
        resources.ApplyResources(Me.lblNo, "lblNo")
        Me.lblNo.Name = "lblNo"
        '
        'lblBestValueScore
        '
        resources.ApplyResources(Me.lblBestValueScore, "lblBestValueScore")
        Me.lblBestValueScore.Name = "lblBestValueScore"
        '
        'IndicatorMaxDiff
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbScoreValues)
        Me.Controls.Add(Me.gbAgreement)
        Me.Controls.Add(Me.gbStatements)
        Me.Name = "IndicatorMaxDiff"
        Me.gbAgreement.ResumeLayout(False)
        Me.gbAgreement.PerformLayout()
        Me.gbScoreValues.ResumeLayout(False)
        Me.gbScoreValues.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbStatements As System.Windows.Forms.GroupBox
    Friend WithEvents gbAgreement As System.Windows.Forms.GroupBox
    Friend WithEvents tbDisagreeText As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblWorstOption As System.Windows.Forms.Label
    Friend WithEvents tbAgreeText As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblBestOption As System.Windows.Forms.Label
    Friend WithEvents gbScoreValues As System.Windows.Forms.GroupBox
    Friend WithEvents ntbNotSelectedScore As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ntbWorstOptionScore As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbBestOptionScore As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents lblNo As System.Windows.Forms.Label
    Friend WithEvents lblBestValueScore As System.Windows.Forms.Label

End Class
