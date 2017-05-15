<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BudgetItemRatio
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BudgetItemRatio))
        Me.lblRatio = New System.Windows.Forms.Label()
        Me.lblReferenceRow = New System.Windows.Forms.Label()
        Me.cmbReferenceRow = New System.Windows.Forms.ComboBox()
        Me.lblAmount = New System.Windows.Forms.Label()
        Me.ntbTotalCost = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.tblRatio = New System.Windows.Forms.TableLayoutPanel()
        Me.ntbReferenceAmount = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.lblReferenceAmount = New System.Windows.Forms.Label()
        Me.ntbRatio = New FaciliDev.LogFramer.NumericBoundTextBoxLF()
        Me.tblRatio.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblRatio
        '
        resources.ApplyResources(Me.lblRatio, "lblRatio")
        Me.lblRatio.Name = "lblRatio"
        '
        'lblReferenceRow
        '
        resources.ApplyResources(Me.lblReferenceRow, "lblReferenceRow")
        Me.lblReferenceRow.Name = "lblReferenceRow"
        '
        'cmbReferenceRow
        '
        resources.ApplyResources(Me.cmbReferenceRow, "cmbReferenceRow")
        Me.cmbReferenceRow.FormattingEnabled = True
        Me.cmbReferenceRow.Name = "cmbReferenceRow"
        '
        'lblAmount
        '
        resources.ApplyResources(Me.lblAmount, "lblAmount")
        Me.lblAmount.Name = "lblAmount"
        '
        'ntbTotalCost
        '
        resources.ApplyResources(Me.ntbTotalCost, "ntbTotalCost")
        Me.ntbTotalCost.Name = "ntbTotalCost"
        Me.ntbTotalCost.ReadOnly = True
        '
        'tblRatio
        '
        resources.ApplyResources(Me.tblRatio, "tblRatio")
        Me.tblRatio.Controls.Add(Me.ntbReferenceAmount, 1, 1)
        Me.tblRatio.Controls.Add(Me.lblReferenceAmount, 1, 0)
        Me.tblRatio.Controls.Add(Me.lblReferenceRow, 0, 0)
        Me.tblRatio.Controls.Add(Me.cmbReferenceRow, 0, 1)
        Me.tblRatio.Controls.Add(Me.ntbTotalCost, 3, 1)
        Me.tblRatio.Controls.Add(Me.lblAmount, 3, 0)
        Me.tblRatio.Controls.Add(Me.lblRatio, 2, 0)
        Me.tblRatio.Controls.Add(Me.ntbRatio, 2, 1)
        Me.tblRatio.Name = "tblRatio"
        '
        'ntbReferenceAmount
        '
        resources.ApplyResources(Me.ntbReferenceAmount, "ntbReferenceAmount")
        Me.ntbReferenceAmount.Name = "ntbReferenceAmount"
        Me.ntbReferenceAmount.ReadOnly = True
        '
        'lblReferenceAmount
        '
        resources.ApplyResources(Me.lblReferenceAmount, "lblReferenceAmount")
        Me.lblReferenceAmount.Name = "lblReferenceAmount"
        '
        'ntbRatio
        '
        resources.ApplyResources(Me.ntbRatio, "ntbRatio")
        Me.ntbRatio.Name = "ntbRatio"
        '
        'BudgetItemRatio
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.tblRatio)
        Me.Name = "BudgetItemRatio"
        Me.tblRatio.ResumeLayout(False)
        Me.tblRatio.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblRatio As System.Windows.Forms.Label
    Friend WithEvents lblReferenceRow As System.Windows.Forms.Label
    Friend WithEvents cmbReferenceRow As System.Windows.Forms.ComboBox
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents ntbTotalCost As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents tblRatio As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents ntbReferenceAmount As FaciliDev.LogFramer.NumericBoundTextBoxLF
    Friend WithEvents lblReferenceAmount As System.Windows.Forms.Label
    Friend WithEvents ntbRatio As FaciliDev.LogFramer.NumericBoundTextBoxLF

End Class
