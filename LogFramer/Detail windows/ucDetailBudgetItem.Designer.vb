<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailBudgetItem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailBudgetItem))
        Me.tbBudgetItem = New FaciliDev.LogFramer.TextBoxLF()
        Me.TabControlBudgetItem = New System.Windows.Forms.TabControl()
        Me.TabPageReferences = New System.Windows.Forms.TabPage()
        Me.TabPageType = New System.Windows.Forms.TabPage()
        Me.lblType = New System.Windows.Forms.Label()
        Me.cmbType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.PanelBudgetItemType = New System.Windows.Forms.Panel()
        Me.TabControlBudgetItem.SuspendLayout()
        Me.TabPageType.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbBudgetItem
        '
        resources.ApplyResources(Me.tbBudgetItem, "tbBudgetItem")
        Me.tbBudgetItem.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbBudgetItem.ForeColor = System.Drawing.Color.Blue
        Me.tbBudgetItem.Name = "tbBudgetItem"
        Me.tbBudgetItem.ReadOnly = True
        '
        'TabControlBudgetItem
        '
        resources.ApplyResources(Me.TabControlBudgetItem, "TabControlBudgetItem")
        Me.TabControlBudgetItem.Controls.Add(Me.TabPageReferences)
        Me.TabControlBudgetItem.Controls.Add(Me.TabPageType)
        Me.TabControlBudgetItem.Name = "TabControlBudgetItem"
        Me.TabControlBudgetItem.SelectedIndex = 0
        '
        'TabPageReferences
        '
        resources.ApplyResources(Me.TabPageReferences, "TabPageReferences")
        Me.TabPageReferences.Name = "TabPageReferences"
        Me.TabPageReferences.UseVisualStyleBackColor = True
        '
        'TabPageType
        '
        resources.ApplyResources(Me.TabPageType, "TabPageType")
        Me.TabPageType.Controls.Add(Me.lblType)
        Me.TabPageType.Controls.Add(Me.cmbType)
        Me.TabPageType.Controls.Add(Me.PanelBudgetItemType)
        Me.TabPageType.Name = "TabPageType"
        Me.TabPageType.UseVisualStyleBackColor = True
        '
        'lblType
        '
        resources.ApplyResources(Me.lblType, "lblType")
        Me.lblType.Name = "lblType"
        '
        'cmbType
        '
        resources.ApplyResources(Me.cmbType, "cmbType")
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Name = "cmbType"
        '
        'PanelBudgetItemType
        '
        resources.ApplyResources(Me.PanelBudgetItemType, "PanelBudgetItemType")
        Me.PanelBudgetItemType.Name = "PanelBudgetItemType"
        '
        'DetailBudgetItem
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.TabControlBudgetItem)
        Me.Controls.Add(Me.tbBudgetItem)
        Me.Name = "DetailBudgetItem"
        Me.TabControlBudgetItem.ResumeLayout(False)
        Me.TabPageType.ResumeLayout(False)
        Me.TabPageType.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbBudgetItem As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents TabControlBudgetItem As System.Windows.Forms.TabControl
    Friend WithEvents TabPageType As System.Windows.Forms.TabPage
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cmbType As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents PanelBudgetItemType As System.Windows.Forms.Panel
    Friend WithEvents TabPageReferences As System.Windows.Forms.TabPage

End Class
