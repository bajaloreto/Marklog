<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailKeyMoment
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailKeyMoment))
        Me.dtpKeyMoment = New System.Windows.Forms.DateTimePicker()
        Me.cmbReferenceMoment = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbPeriodDirection = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbPeriodUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.ntbPeriod = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.rbtnReferenceDate = New System.Windows.Forms.RadioButton()
        Me.rbtnExactDate = New System.Windows.Forms.RadioButton()
        Me.lblDate = New System.Windows.Forms.Label()
        Me.tbDescription = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.dtbExactDateKeyMoment = New FaciliDev.LogFramer.DateTextBox()
        Me.SuspendLayout()
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
        'lblDate
        '
        resources.ApplyResources(Me.lblDate, "lblDate")
        Me.lblDate.Name = "lblDate"
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
        'dtbExactDateKeyMoment
        '
        resources.ApplyResources(Me.dtbExactDateKeyMoment, "dtbExactDateKeyMoment")
        Me.dtbExactDateKeyMoment.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbExactDateKeyMoment.ForeColor = System.Drawing.Color.Black
        Me.dtbExactDateKeyMoment.Name = "dtbExactDateKeyMoment"
        Me.dtbExactDateKeyMoment.ReadOnly = True
        Me.dtbExactDateKeyMoment.TabStop = False
        '
        'DetailKeyMoment
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.dtbExactDateKeyMoment)
        Me.Controls.Add(Me.tbDescription)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.dtpKeyMoment)
        Me.Controls.Add(Me.cmbReferenceMoment)
        Me.Controls.Add(Me.cmbPeriodDirection)
        Me.Controls.Add(Me.cmbPeriodUnit)
        Me.Controls.Add(Me.ntbPeriod)
        Me.Controls.Add(Me.rbtnReferenceDate)
        Me.Controls.Add(Me.rbtnExactDate)
        Me.Name = "DetailKeyMoment"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpKeyMoment As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbReferenceMoment As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbPeriodDirection As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents cmbPeriodUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents ntbPeriod As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents rbtnReferenceDate As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnExactDate As System.Windows.Forms.RadioButton
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents tbDescription As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents dtbExactDateKeyMoment As FaciliDev.LogFramer.DateTextBox

End Class
