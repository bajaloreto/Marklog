<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailResource
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailResource))
        Me.tbResource = New System.Windows.Forms.TextBox()
        Me.PanelBudgetItems = New System.Windows.Forms.Panel()
        Me.lblResource = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'tbResource
        '
        resources.ApplyResources(Me.tbResource, "tbResource")
        Me.tbResource.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbResource.ForeColor = System.Drawing.Color.Blue
        Me.tbResource.Name = "tbResource"
        Me.tbResource.ReadOnly = True
        '
        'PanelBudgetItems
        '
        resources.ApplyResources(Me.PanelBudgetItems, "PanelBudgetItems")
        Me.PanelBudgetItems.Name = "PanelBudgetItems"
        '
        'lblResource
        '
        resources.ApplyResources(Me.lblResource, "lblResource")
        Me.lblResource.Name = "lblResource"
        '
        'DetailResource
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.lblResource)
        Me.Controls.Add(Me.PanelBudgetItems)
        Me.Controls.Add(Me.tbResource)
        Me.Name = "DetailResource"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbResource As System.Windows.Forms.TextBox
    Friend WithEvents PanelBudgetItems As System.Windows.Forms.Panel
    Friend WithEvents lblResource As System.Windows.Forms.Label

End Class
