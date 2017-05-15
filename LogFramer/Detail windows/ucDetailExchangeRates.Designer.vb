<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailExchangeRates
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailExchangeRates))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.lblDefaultCurrency = New System.Windows.Forms.Label()
        Me.cmbDefaultCurrencyCode = New System.Windows.Forms.ComboBox()
        Me.lblExchangeRates = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.dgvExchangeRates = New FaciliDev.LogFramer.DataGridViewExchangeRates()
        CType(Me.dgvExchangeRates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblDefaultCurrency
        '
        resources.ApplyResources(Me.lblDefaultCurrency, "lblDefaultCurrency")
        Me.lblDefaultCurrency.Name = "lblDefaultCurrency"
        '
        'cmbDefaultCurrencyCode
        '
        resources.ApplyResources(Me.cmbDefaultCurrencyCode, "cmbDefaultCurrencyCode")
        Me.cmbDefaultCurrencyCode.FormattingEnabled = True
        Me.cmbDefaultCurrencyCode.Name = "cmbDefaultCurrencyCode"
        '
        'lblExchangeRates
        '
        resources.ApplyResources(Me.lblExchangeRates, "lblExchangeRates")
        Me.lblExchangeRates.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.lblExchangeRates.Name = "lblExchangeRates"
        '
        'btnClose
        '
        resources.ApplyResources(Me.btnClose, "btnClose")
        Me.btnClose.Name = "btnClose"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'dgvExchangeRates
        '
        resources.ApplyResources(Me.dgvExchangeRates, "dgvExchangeRates")
        Me.dgvExchangeRates.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvExchangeRates.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvExchangeRates.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvExchangeRates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(2)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvExchangeRates.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvExchangeRates.DefaultCurrencyCode = Nothing
        Me.dgvExchangeRates.ExchangeRates = Nothing
        Me.dgvExchangeRates.Name = "dgvExchangeRates"
        Me.dgvExchangeRates.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvExchangeRates.ShowCellToolTips = False
        Me.dgvExchangeRates.VirtualMode = True
        '
        'DetailExchangeRates
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.dgvExchangeRates)
        Me.Controls.Add(Me.lblExchangeRates)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.cmbDefaultCurrencyCode)
        Me.Controls.Add(Me.lblDefaultCurrency)
        Me.Name = "DetailExchangeRates"
        CType(Me.dgvExchangeRates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblDefaultCurrency As System.Windows.Forms.Label
    Friend WithEvents cmbDefaultCurrencyCode As System.Windows.Forms.ComboBox
    Friend WithEvents lblExchangeRates As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents dgvExchangeRates As FaciliDev.LogFramer.DataGridViewExchangeRates

End Class
