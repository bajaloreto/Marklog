<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogTargetFormula
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogTargetFormula))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.lblFormula = New System.Windows.Forms.Label()
        Me.lblFormulaDescription = New System.Windows.Forms.Label()
        Me.cmbOpMin = New System.Windows.Forms.ComboBox()
        Me.lblDeadline = New System.Windows.Forms.Label()
        Me.lblScoringValueDescription = New System.Windows.Forms.Label()
        Me.lblScoringValue = New System.Windows.Forms.Label()
        Me.ntbScoringValue = New FaciliDev.LogFramer.NumericTextBox()
        Me.dtbDeadline = New FaciliDev.LogFramer.DateTextBox()
        Me.tbFormula = New System.Windows.Forms.TextBox()
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
        'lblFormula
        '
        resources.ApplyResources(Me.lblFormula, "lblFormula")
        Me.lblFormula.Name = "lblFormula"
        '
        'lblFormulaDescription
        '
        resources.ApplyResources(Me.lblFormulaDescription, "lblFormulaDescription")
        Me.lblFormulaDescription.Name = "lblFormulaDescription"
        '
        'cmbOpMin
        '
        resources.ApplyResources(Me.cmbOpMin, "cmbOpMin")
        Me.cmbOpMin.FormattingEnabled = True
        Me.cmbOpMin.Name = "cmbOpMin"
        '
        'lblDeadline
        '
        resources.ApplyResources(Me.lblDeadline, "lblDeadline")
        Me.lblDeadline.Name = "lblDeadline"
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
        'ntbScoringValue
        '
        resources.ApplyResources(Me.ntbScoringValue, "ntbScoringValue")
        Me.ntbScoringValue.AllowSpace = False
        Me.ntbScoringValue.DoubleValue = 0.0R
        Me.ntbScoringValue.IntegerValue = 0
        Me.ntbScoringValue.IsCurrency = False
        Me.ntbScoringValue.IsPercentage = False
        Me.ntbScoringValue.Name = "ntbScoringValue"
        Me.ntbScoringValue.NrDecimals = 0
        Me.ntbScoringValue.SetDecimals = False
        Me.ntbScoringValue.SingleValue = 0.0!
        Me.ntbScoringValue.Unit = Nothing
        Me.ntbScoringValue.ValueType = 0
        '
        'dtbDeadline
        '
        resources.ApplyResources(Me.dtbDeadline, "dtbDeadline")
        Me.dtbDeadline.DateValue = New Date(CType(0, Long))
        Me.dtbDeadline.ForeColor = System.Drawing.SystemColors.Window
        Me.dtbDeadline.Name = "dtbDeadline"
        Me.dtbDeadline.ReadOnly = True
        '
        'tbFormula
        '
        resources.ApplyResources(Me.tbFormula, "tbFormula")
        Me.tbFormula.Name = "tbFormula"
        '
        'DialogTargetFormula
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.tbFormula)
        Me.Controls.Add(Me.ntbScoringValue)
        Me.Controls.Add(Me.lblScoringValueDescription)
        Me.Controls.Add(Me.lblScoringValue)
        Me.Controls.Add(Me.dtbDeadline)
        Me.Controls.Add(Me.lblDeadline)
        Me.Controls.Add(Me.cmbOpMin)
        Me.Controls.Add(Me.lblFormulaDescription)
        Me.Controls.Add(Me.lblFormula)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogTargetFormula"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents lblFormula As System.Windows.Forms.Label
    Friend WithEvents lblFormulaDescription As System.Windows.Forms.Label
    Friend WithEvents cmbOpMin As System.Windows.Forms.ComboBox
    Friend WithEvents lblDeadline As System.Windows.Forms.Label
    Friend WithEvents dtbDeadline As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents ntbScoringValue As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents lblScoringValueDescription As System.Windows.Forms.Label
    Friend WithEvents lblScoringValue As System.Windows.Forms.Label
    Friend WithEvents tbFormula As System.Windows.Forms.TextBox

End Class
