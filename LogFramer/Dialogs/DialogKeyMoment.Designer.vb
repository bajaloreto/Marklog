<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogKeyMoment
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogKeyMoment))
        Me.lblOutput = New System.Windows.Forms.Label()
        Me.tbDescription = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.gbRelative = New System.Windows.Forms.GroupBox()
        Me.dtpKeyMoment = New FaciliDev.LogFramer.DateTimePickerLF()
        Me.cmbReferenceMoment = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbPeriodDirection = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbPeriodUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.ntbPeriod = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.rbtnReferenceDate = New System.Windows.Forms.RadioButton()
        Me.rbtnExactDate = New System.Windows.Forms.RadioButton()
        Me.cmbOutput = New System.Windows.Forms.ComboBox()
        Me.btnReady = New System.Windows.Forms.Button()
        Me.gbRelative.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblOutput
        '
        resources.ApplyResources(Me.lblOutput, "lblOutput")
        Me.lblOutput.Name = "lblOutput"
        '
        'tbDescription
        '
        resources.ApplyResources(Me.tbDescription, "tbDescription")
        Me.tbDescription.Name = "tbDescription"
        '
        'lblDescription
        '
        resources.ApplyResources(Me.lblDescription, "lblDescription")
        Me.lblDescription.Name = "lblDescription"
        '
        'gbRelative
        '
        resources.ApplyResources(Me.gbRelative, "gbRelative")
        Me.gbRelative.Controls.Add(Me.dtpKeyMoment)
        Me.gbRelative.Controls.Add(Me.cmbReferenceMoment)
        Me.gbRelative.Controls.Add(Me.cmbPeriodDirection)
        Me.gbRelative.Controls.Add(Me.cmbPeriodUnit)
        Me.gbRelative.Controls.Add(Me.ntbPeriod)
        Me.gbRelative.Controls.Add(Me.rbtnReferenceDate)
        Me.gbRelative.Controls.Add(Me.rbtnExactDate)
        Me.gbRelative.Name = "gbRelative"
        Me.gbRelative.TabStop = False
        '
        'dtpKeyMoment
        '
        resources.ApplyResources(Me.dtpKeyMoment, "dtpKeyMoment")
        Me.dtpKeyMoment.Name = "dtpKeyMoment"
        '
        'cmbReferenceMoment
        '
        resources.ApplyResources(Me.cmbReferenceMoment, "cmbReferenceMoment")
        Me.cmbReferenceMoment.FormattingEnabled = True
        Me.cmbReferenceMoment.Name = "cmbReferenceMoment"
        '
        'cmbPeriodDirection
        '
        resources.ApplyResources(Me.cmbPeriodDirection, "cmbPeriodDirection")
        Me.cmbPeriodDirection.FormattingEnabled = True
        Me.cmbPeriodDirection.Name = "cmbPeriodDirection"
        '
        'cmbPeriodUnit
        '
        resources.ApplyResources(Me.cmbPeriodUnit, "cmbPeriodUnit")
        Me.cmbPeriodUnit.FormattingEnabled = True
        Me.cmbPeriodUnit.Name = "cmbPeriodUnit"
        '
        'ntbPeriod
        '
        resources.ApplyResources(Me.ntbPeriod, "ntbPeriod")
        Me.ntbPeriod.AllowSpace = True
        Me.ntbPeriod.IsCurrency = False
        Me.ntbPeriod.IsPercentage = False
        Me.ntbPeriod.Name = "ntbPeriod"
        '
        'rbtnReferenceDate
        '
        resources.ApplyResources(Me.rbtnReferenceDate, "rbtnReferenceDate")
        Me.rbtnReferenceDate.Checked = True
        Me.rbtnReferenceDate.Name = "rbtnReferenceDate"
        Me.rbtnReferenceDate.TabStop = True
        Me.rbtnReferenceDate.UseVisualStyleBackColor = True
        '
        'rbtnExactDate
        '
        resources.ApplyResources(Me.rbtnExactDate, "rbtnExactDate")
        Me.rbtnExactDate.Checked = True
        Me.rbtnExactDate.Name = "rbtnExactDate"
        Me.rbtnExactDate.TabStop = True
        Me.rbtnExactDate.UseVisualStyleBackColor = True
        '
        'cmbOutput
        '
        resources.ApplyResources(Me.cmbOutput, "cmbOutput")
        Me.cmbOutput.FormattingEnabled = True
        Me.cmbOutput.Name = "cmbOutput"
        '
        'btnReady
        '
        resources.ApplyResources(Me.btnReady, "btnReady")
        Me.btnReady.Name = "btnReady"
        '
        'DialogKeyMoment
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnReady)
        Me.Controls.Add(Me.cmbOutput)
        Me.Controls.Add(Me.gbRelative)
        Me.Controls.Add(Me.tbDescription)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblOutput)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogKeyMoment"
        Me.ShowInTaskbar = False
        Me.gbRelative.ResumeLayout(False)
        Me.gbRelative.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblOutput As System.Windows.Forms.Label
    Friend WithEvents tbDescription As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents gbRelative As System.Windows.Forms.GroupBox
    Friend WithEvents dtpKeyMoment As FaciliDev.LogFramer.DateTimePickerLF
    Friend WithEvents cmbReferenceMoment As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbPeriodDirection As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents cmbPeriodUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents ntbPeriod As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents rbtnReferenceDate As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnExactDate As System.Windows.Forms.RadioButton
    Friend WithEvents cmbOutput As System.Windows.Forms.ComboBox
    Friend WithEvents btnReady As System.Windows.Forms.Button

End Class
