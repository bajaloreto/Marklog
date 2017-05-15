<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSettingsBudget
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintSettingsBudget))
        Me.GroupBoxShow = New System.Windows.Forms.GroupBox()
        Me.chkShowExchangeRates = New System.Windows.Forms.CheckBox()
        Me.chkShowDurationColumns = New System.Windows.Forms.CheckBox()
        Me.chkShowLocalCurrencyColumns = New System.Windows.Forms.CheckBox()
        Me.cmbBudgetYear = New System.Windows.Forms.ComboBox()
        Me.GroupBoxBudgetYear = New System.Windows.Forms.GroupBox()
        Me.GroupBoxWidth = New System.Windows.Forms.GroupBox()
        Me.lblPages = New System.Windows.Forms.Label()
        Me.nudPagesWide = New System.Windows.Forms.NumericUpDown()
        Me.GroupBoxShow.SuspendLayout()
        Me.GroupBoxBudgetYear.SuspendLayout()
        Me.GroupBoxWidth.SuspendLayout()
        CType(Me.nudPagesWide, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBoxShow
        '
        resources.ApplyResources(Me.GroupBoxShow, "GroupBoxShow")
        Me.GroupBoxShow.Controls.Add(Me.chkShowExchangeRates)
        Me.GroupBoxShow.Controls.Add(Me.chkShowDurationColumns)
        Me.GroupBoxShow.Controls.Add(Me.chkShowLocalCurrencyColumns)
        Me.GroupBoxShow.Name = "GroupBoxShow"
        Me.GroupBoxShow.TabStop = False
        '
        'chkShowExchangeRates
        '
        resources.ApplyResources(Me.chkShowExchangeRates, "chkShowExchangeRates")
        Me.chkShowExchangeRates.Name = "chkShowExchangeRates"
        Me.chkShowExchangeRates.UseVisualStyleBackColor = True
        '
        'chkShowDurationColumns
        '
        resources.ApplyResources(Me.chkShowDurationColumns, "chkShowDurationColumns")
        Me.chkShowDurationColumns.Name = "chkShowDurationColumns"
        Me.chkShowDurationColumns.UseVisualStyleBackColor = True
        '
        'chkShowLocalCurrencyColumns
        '
        resources.ApplyResources(Me.chkShowLocalCurrencyColumns, "chkShowLocalCurrencyColumns")
        Me.chkShowLocalCurrencyColumns.Name = "chkShowLocalCurrencyColumns"
        Me.chkShowLocalCurrencyColumns.UseVisualStyleBackColor = True
        '
        'cmbBudgetYear
        '
        resources.ApplyResources(Me.cmbBudgetYear, "cmbBudgetYear")
        Me.cmbBudgetYear.FormattingEnabled = True
        Me.cmbBudgetYear.Name = "cmbBudgetYear"
        '
        'GroupBoxBudgetYear
        '
        resources.ApplyResources(Me.GroupBoxBudgetYear, "GroupBoxBudgetYear")
        Me.GroupBoxBudgetYear.Controls.Add(Me.cmbBudgetYear)
        Me.GroupBoxBudgetYear.Name = "GroupBoxBudgetYear"
        Me.GroupBoxBudgetYear.TabStop = False
        '
        'GroupBoxWidth
        '
        resources.ApplyResources(Me.GroupBoxWidth, "GroupBoxWidth")
        Me.GroupBoxWidth.Controls.Add(Me.lblPages)
        Me.GroupBoxWidth.Controls.Add(Me.nudPagesWide)
        Me.GroupBoxWidth.Name = "GroupBoxWidth"
        Me.GroupBoxWidth.TabStop = False
        '
        'lblPages
        '
        resources.ApplyResources(Me.lblPages, "lblPages")
        Me.lblPages.Name = "lblPages"
        '
        'nudPagesWide
        '
        resources.ApplyResources(Me.nudPagesWide, "nudPagesWide")
        Me.nudPagesWide.Name = "nudPagesWide"
        '
        'PrintSettingsBudget
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBoxWidth)
        Me.Controls.Add(Me.GroupBoxShow)
        Me.Controls.Add(Me.GroupBoxBudgetYear)
        Me.Name = "PrintSettingsBudget"
        Me.GroupBoxShow.ResumeLayout(False)
        Me.GroupBoxShow.PerformLayout()
        Me.GroupBoxBudgetYear.ResumeLayout(False)
        Me.GroupBoxWidth.ResumeLayout(False)
        Me.GroupBoxWidth.PerformLayout()
        CType(Me.nudPagesWide, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBoxShow As System.Windows.Forms.GroupBox
    Friend WithEvents chkShowLocalCurrencyColumns As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowDurationColumns As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowExchangeRates As System.Windows.Forms.CheckBox
    Friend WithEvents cmbBudgetYear As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBoxBudgetYear As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBoxWidth As System.Windows.Forms.GroupBox
    Friend WithEvents lblPages As System.Windows.Forms.Label
    Friend WithEvents nudPagesWide As System.Windows.Forms.NumericUpDown

End Class
