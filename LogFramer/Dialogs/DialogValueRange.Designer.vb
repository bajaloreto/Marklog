<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogValueRange
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogValueRange))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.lblMinValue = New System.Windows.Forms.Label()
        Me.lblMinValueDescription = New System.Windows.Forms.Label()
        Me.cmbOpMin = New System.Windows.Forms.ComboBox()
        Me.cmbOpMax = New System.Windows.Forms.ComboBox()
        Me.lblMaxValueDescription = New System.Windows.Forms.Label()
        Me.lblMaxValue = New System.Windows.Forms.Label()
        Me.tbMinValue = New FaciliDev.LogFramer.NumericTextBox()
        Me.tbMaxValue = New FaciliDev.LogFramer.NumericTextBox()
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
        'tbMinValue
        '
        resources.ApplyResources(Me.tbMinValue, "tbMinValue")
        Me.tbMinValue.AllowSpace = False
        Me.tbMinValue.DoubleValue = 0.0R
        Me.tbMinValue.IntegerValue = 0
        Me.tbMinValue.IsCurrency = False
        Me.tbMinValue.IsPercentage = False
        Me.tbMinValue.Name = "tbMinValue"
        Me.tbMinValue.NrDecimals = 0
        Me.tbMinValue.SetDecimals = False
        Me.tbMinValue.SingleValue = 0.0!
        Me.tbMinValue.Unit = Nothing
        Me.tbMinValue.ValueType = 0
        '
        'tbMaxValue
        '
        resources.ApplyResources(Me.tbMaxValue, "tbMaxValue")
        Me.tbMaxValue.AllowSpace = False
        Me.tbMaxValue.DoubleValue = 0.0R
        Me.tbMaxValue.IntegerValue = 0
        Me.tbMaxValue.IsCurrency = False
        Me.tbMaxValue.IsPercentage = False
        Me.tbMaxValue.Name = "tbMaxValue"
        Me.tbMaxValue.NrDecimals = 0
        Me.tbMaxValue.SetDecimals = False
        Me.tbMaxValue.SingleValue = 0.0!
        Me.tbMaxValue.Unit = Nothing
        Me.tbMaxValue.ValueType = 0
        '
        'DialogValueRange
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.tbMaxValue)
        Me.Controls.Add(Me.tbMinValue)
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
        Me.Name = "DialogValueRange"
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
    Friend WithEvents cmbOpMin As System.Windows.Forms.ComboBox
    Friend WithEvents cmbOpMax As System.Windows.Forms.ComboBox
    Friend WithEvents lblMaxValueDescription As System.Windows.Forms.Label
    Friend WithEvents lblMaxValue As System.Windows.Forms.Label
    Friend WithEvents tbMinValue As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents tbMaxValue As FaciliDev.LogFramer.NumericTextBox

End Class
