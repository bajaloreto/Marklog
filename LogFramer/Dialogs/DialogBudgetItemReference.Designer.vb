<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogBudgetItemReference
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogBudgetItemReference))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.cmbActivity = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbResource = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblActivity = New System.Windows.Forms.Label()
        Me.lblResource = New System.Windows.Forms.Label()
        Me.ntbPercentage = New FaciliDev.LogFramer.NumericTextBox()
        Me.lblPercentage = New System.Windows.Forms.Label()
        Me.lblTotalCost = New System.Windows.Forms.Label()
        Me.ntbTotalCost = New FaciliDev.LogFramer.NumericTextBox()
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
        'cmbActivity
        '
        resources.ApplyResources(Me.cmbActivity, "cmbActivity")
        Me.cmbActivity.FormattingEnabled = True
        Me.cmbActivity.Name = "cmbActivity"
        '
        'cmbResource
        '
        resources.ApplyResources(Me.cmbResource, "cmbResource")
        Me.cmbResource.FormattingEnabled = True
        Me.cmbResource.Name = "cmbResource"
        '
        'lblActivity
        '
        resources.ApplyResources(Me.lblActivity, "lblActivity")
        Me.lblActivity.Name = "lblActivity"
        '
        'lblResource
        '
        resources.ApplyResources(Me.lblResource, "lblResource")
        Me.lblResource.Name = "lblResource"
        '
        'ntbPercentage
        '
        resources.ApplyResources(Me.ntbPercentage, "ntbPercentage")
        Me.ntbPercentage.AllowSpace = False
        Me.ntbPercentage.DoubleValue = 0.0R
        Me.ntbPercentage.IntegerValue = 0
        Me.ntbPercentage.IsCurrency = False
        Me.ntbPercentage.IsPercentage = False
        Me.ntbPercentage.Name = "ntbPercentage"
        Me.ntbPercentage.NrDecimals = 0
        Me.ntbPercentage.SetDecimals = False
        Me.ntbPercentage.SingleValue = 0.0!
        Me.ntbPercentage.Unit = Nothing
        Me.ntbPercentage.ValueType = 0
        '
        'lblPercentage
        '
        resources.ApplyResources(Me.lblPercentage, "lblPercentage")
        Me.lblPercentage.Name = "lblPercentage"
        '
        'lblTotalCost
        '
        resources.ApplyResources(Me.lblTotalCost, "lblTotalCost")
        Me.lblTotalCost.Name = "lblTotalCost"
        '
        'ntbTotalCost
        '
        resources.ApplyResources(Me.ntbTotalCost, "ntbTotalCost")
        Me.ntbTotalCost.AllowSpace = False
        Me.ntbTotalCost.DoubleValue = 0.0R
        Me.ntbTotalCost.IntegerValue = 0
        Me.ntbTotalCost.IsCurrency = False
        Me.ntbTotalCost.IsPercentage = False
        Me.ntbTotalCost.Name = "ntbTotalCost"
        Me.ntbTotalCost.NrDecimals = 0
        Me.ntbTotalCost.SetDecimals = False
        Me.ntbTotalCost.SingleValue = 0.0!
        Me.ntbTotalCost.Unit = Nothing
        Me.ntbTotalCost.ValueType = 0
        '
        'DialogBudgetItemReference
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.lblTotalCost)
        Me.Controls.Add(Me.ntbTotalCost)
        Me.Controls.Add(Me.lblPercentage)
        Me.Controls.Add(Me.ntbPercentage)
        Me.Controls.Add(Me.lblResource)
        Me.Controls.Add(Me.lblActivity)
        Me.Controls.Add(Me.cmbResource)
        Me.Controls.Add(Me.cmbActivity)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogBudgetItemReference"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents cmbActivity As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents cmbResource As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents lblResource As System.Windows.Forms.Label
    Friend WithEvents ntbPercentage As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents lblPercentage As System.Windows.Forms.Label
    Friend WithEvents lblTotalCost As System.Windows.Forms.Label
    Friend WithEvents ntbTotalCost As FaciliDev.LogFramer.NumericTextBox

End Class
