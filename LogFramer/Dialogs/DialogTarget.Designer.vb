<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogTarget
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogTarget))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.lblMinValue = New System.Windows.Forms.Label()
        Me.lblMinValueDescription = New System.Windows.Forms.Label()
        Me.cmbOpMin = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbOpMax = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.lblMaxValueDescription = New System.Windows.Forms.Label()
        Me.lblMaxValue = New System.Windows.Forms.Label()
        Me.ntbMinValue = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.ntbMaxValue = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblDeadline = New System.Windows.Forms.Label()
        Me.dtbDeadline = New FaciliDev.LogFramer.DateTextBox()
        Me.ntbScoringValue = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblScoringValueDescription = New System.Windows.Forms.Label()
        Me.lblScoringValue = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'OK_Button
        '
        resources.ApplyResources(Me.OK_Button, "OK_Button")
        Me.OK_Button.Name = "OK_Button"
        '
        'Cancel_Button
        '
        resources.ApplyResources(Me.Cancel_Button, "Cancel_Button")
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Name = "Cancel_Button"
        '
        'lblMinValue
        '
        resources.ApplyResources(Me.lblMinValue, "lblMinValue")
        Me.lblMinValue.Name = "lblMinValue"
        '
        'lblMinValueDescription
        '
        resources.ApplyResources(Me.lblMinValueDescription, "lblMinValueDescription")
        Me.lblMinValueDescription.Name = "lblMinValueDescription"
        '
        'cmbOpMin
        '
        resources.ApplyResources(Me.cmbOpMin, "cmbOpMin")
        Me.cmbOpMin.FormattingEnabled = True
        Me.cmbOpMin.Name = "cmbOpMin"
        '
        'cmbOpMax
        '
        resources.ApplyResources(Me.cmbOpMax, "cmbOpMax")
        Me.cmbOpMax.FormattingEnabled = True
        Me.cmbOpMax.Name = "cmbOpMax"
        '
        'lblMaxValueDescription
        '
        resources.ApplyResources(Me.lblMaxValueDescription, "lblMaxValueDescription")
        Me.lblMaxValueDescription.Name = "lblMaxValueDescription"
        '
        'lblMaxValue
        '
        resources.ApplyResources(Me.lblMaxValue, "lblMaxValue")
        Me.lblMaxValue.Name = "lblMaxValue"
        '
        'ntbMinValue
        '
        resources.ApplyResources(Me.ntbMinValue, "ntbMinValue")
        Me.ntbMinValue.Name = "ntbMinValue"
        '
        'ntbMaxValue
        '
        resources.ApplyResources(Me.ntbMaxValue, "ntbMaxValue")
        Me.ntbMaxValue.Name = "ntbMaxValue"
        '
        'lblDeadline
        '
        resources.ApplyResources(Me.lblDeadline, "lblDeadline")
        Me.lblDeadline.Name = "lblDeadline"
        '
        'dtbDeadline
        '
        resources.ApplyResources(Me.dtbDeadline, "dtbDeadline")
        Me.dtbDeadline.DateValue = New Date(CType(0, Long))
        Me.dtbDeadline.ForeColor = System.Drawing.SystemColors.Window
        Me.dtbDeadline.Name = "dtbDeadline"
        Me.dtbDeadline.ReadOnly = True
        '
        'ntbScoringValue
        '
        resources.ApplyResources(Me.ntbScoringValue, "ntbScoringValue")
        Me.ntbScoringValue.Name = "ntbScoringValue"
        '
        'lblScoringValueDescription
        '
        resources.ApplyResources(Me.lblScoringValueDescription, "lblScoringValueDescription")
        Me.lblScoringValueDescription.Name = "lblScoringValueDescription"
        '
        'lblScoringValue
        '
        resources.ApplyResources(Me.lblScoringValue, "lblScoringValue")
        Me.lblScoringValue.Name = "lblScoringValue"
        '
        'DialogTarget
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.ntbScoringValue)
        Me.Controls.Add(Me.lblScoringValueDescription)
        Me.Controls.Add(Me.lblScoringValue)
        Me.Controls.Add(Me.dtbDeadline)
        Me.Controls.Add(Me.lblDeadline)
        Me.Controls.Add(Me.ntbMaxValue)
        Me.Controls.Add(Me.ntbMinValue)
        Me.Controls.Add(Me.cmbOpMax)
        Me.Controls.Add(Me.lblMaxValueDescription)
        Me.Controls.Add(Me.lblMaxValue)
        Me.Controls.Add(Me.cmbOpMin)
        Me.Controls.Add(Me.lblMinValueDescription)
        Me.Controls.Add(Me.lblMinValue)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogTarget"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents lblMinValue As System.Windows.Forms.Label
    Friend WithEvents lblMinValueDescription As System.Windows.Forms.Label
    Friend WithEvents cmbOpMin As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbOpMax As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents lblMaxValueDescription As System.Windows.Forms.Label
    Friend WithEvents lblMaxValue As System.Windows.Forms.Label
    Friend WithEvents ntbMinValue As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents ntbMaxValue As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents lblDeadline As System.Windows.Forms.Label
    Friend WithEvents dtbDeadline As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents ntbScoringValue As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents lblScoringValueDescription As System.Windows.Forms.Label
    Friend WithEvents lblScoringValue As System.Windows.Forms.Label

End Class
