<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogImportBudget
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogImportBudget))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.TabControlSteps = New System.Windows.Forms.TabControl()
        Me.TabPageSelectDocument = New System.Windows.Forms.TabPage()
        Me.btnToStep2 = New System.Windows.Forms.Button()
        Me.lblStep1 = New System.Windows.Forms.Label()
        Me.btnLoadFile = New System.Windows.Forms.Button()
        Me.lblFileName = New System.Windows.Forms.Label()
        Me.tbFileName = New System.Windows.Forms.TextBox()
        Me.TabPageSelectSpreadsheet = New System.Windows.Forms.TabPage()
        Me.btnToStep3 = New System.Windows.Forms.Button()
        Me.cmbSelectWorksheet = New System.Windows.Forms.ComboBox()
        Me.lblSelectWorksheet = New System.Windows.Forms.Label()
        Me.lblStep2 = New System.Windows.Forms.Label()
        Me.TabPageSelectColumns = New System.Windows.Forms.TabPage()
        Me.btnToStep4 = New System.Windows.Forms.Button()
        Me.TableLayoutPanelColumns = New System.Windows.Forms.TableLayoutPanel()
        Me.cmbColumnDuration = New System.Windows.Forms.ComboBox()
        Me.cmbColumnDescription = New System.Windows.Forms.ComboBox()
        Me.cmbColumnNumber = New System.Windows.Forms.ComboBox()
        Me.cmbColumnUnitCost = New System.Windows.Forms.ComboBox()
        Me.cmbColumnTotalCost = New System.Windows.Forms.ComboBox()
        Me.lblTotalCost = New System.Windows.Forms.Label()
        Me.lblUnitCost = New System.Windows.Forms.Label()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.lblDuration = New System.Windows.Forms.Label()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblStep3 = New System.Windows.Forms.Label()
        Me.TabPageSelectRows = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblFirstRow = New System.Windows.Forms.Label()
        Me.ntbFirstRow = New FaciliDev.LogFramer.NumericTextBox()
        Me.lblLastRow = New System.Windows.Forms.Label()
        Me.ntbLastRow = New FaciliDev.LogFramer.NumericTextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblStep4 = New System.Windows.Forms.Label()
        Me.dgWorkSheet = New FaciliDev.LogFramer.DataGridViewExcelPreview()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControlSteps.SuspendLayout()
        Me.TabPageSelectDocument.SuspendLayout()
        Me.TabPageSelectSpreadsheet.SuspendLayout()
        Me.TabPageSelectColumns.SuspendLayout()
        Me.TableLayoutPanelColumns.SuspendLayout()
        Me.TabPageSelectRows.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.dgWorkSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.btnImport, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'btnImport
        '
        resources.ApplyResources(Me.btnImport, "btnImport")
        Me.btnImport.Name = "btnImport"
        '
        'Cancel_Button
        '
        resources.ApplyResources(Me.Cancel_Button, "Cancel_Button")
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Name = "Cancel_Button"
        '
        'TabControlSteps
        '
        resources.ApplyResources(Me.TabControlSteps, "TabControlSteps")
        Me.TabControlSteps.Controls.Add(Me.TabPageSelectDocument)
        Me.TabControlSteps.Controls.Add(Me.TabPageSelectSpreadsheet)
        Me.TabControlSteps.Controls.Add(Me.TabPageSelectColumns)
        Me.TabControlSteps.Controls.Add(Me.TabPageSelectRows)
        Me.TabControlSteps.Name = "TabControlSteps"
        Me.TabControlSteps.SelectedIndex = 0
        '
        'TabPageSelectDocument
        '
        resources.ApplyResources(Me.TabPageSelectDocument, "TabPageSelectDocument")
        Me.TabPageSelectDocument.Controls.Add(Me.btnToStep2)
        Me.TabPageSelectDocument.Controls.Add(Me.lblStep1)
        Me.TabPageSelectDocument.Controls.Add(Me.btnLoadFile)
        Me.TabPageSelectDocument.Controls.Add(Me.lblFileName)
        Me.TabPageSelectDocument.Controls.Add(Me.tbFileName)
        Me.TabPageSelectDocument.Name = "TabPageSelectDocument"
        Me.TabPageSelectDocument.UseVisualStyleBackColor = True
        '
        'btnToStep2
        '
        resources.ApplyResources(Me.btnToStep2, "btnToStep2")
        Me.btnToStep2.Name = "btnToStep2"
        Me.btnToStep2.UseVisualStyleBackColor = True
        '
        'lblStep1
        '
        resources.ApplyResources(Me.lblStep1, "lblStep1")
        Me.lblStep1.Name = "lblStep1"
        '
        'btnLoadFile
        '
        resources.ApplyResources(Me.btnLoadFile, "btnLoadFile")
        Me.btnLoadFile.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.FileOpen
        Me.btnLoadFile.Name = "btnLoadFile"
        Me.btnLoadFile.UseVisualStyleBackColor = True
        '
        'lblFileName
        '
        resources.ApplyResources(Me.lblFileName, "lblFileName")
        Me.lblFileName.Name = "lblFileName"
        '
        'tbFileName
        '
        resources.ApplyResources(Me.tbFileName, "tbFileName")
        Me.tbFileName.BackColor = System.Drawing.SystemColors.Window
        Me.tbFileName.Name = "tbFileName"
        Me.tbFileName.ReadOnly = True
        '
        'TabPageSelectSpreadsheet
        '
        resources.ApplyResources(Me.TabPageSelectSpreadsheet, "TabPageSelectSpreadsheet")
        Me.TabPageSelectSpreadsheet.Controls.Add(Me.btnToStep3)
        Me.TabPageSelectSpreadsheet.Controls.Add(Me.cmbSelectWorksheet)
        Me.TabPageSelectSpreadsheet.Controls.Add(Me.lblSelectWorksheet)
        Me.TabPageSelectSpreadsheet.Controls.Add(Me.lblStep2)
        Me.TabPageSelectSpreadsheet.Name = "TabPageSelectSpreadsheet"
        Me.TabPageSelectSpreadsheet.UseVisualStyleBackColor = True
        '
        'btnToStep3
        '
        resources.ApplyResources(Me.btnToStep3, "btnToStep3")
        Me.btnToStep3.Name = "btnToStep3"
        Me.btnToStep3.UseVisualStyleBackColor = True
        '
        'cmbSelectWorksheet
        '
        resources.ApplyResources(Me.cmbSelectWorksheet, "cmbSelectWorksheet")
        Me.cmbSelectWorksheet.FormattingEnabled = True
        Me.cmbSelectWorksheet.Name = "cmbSelectWorksheet"
        '
        'lblSelectWorksheet
        '
        resources.ApplyResources(Me.lblSelectWorksheet, "lblSelectWorksheet")
        Me.lblSelectWorksheet.Name = "lblSelectWorksheet"
        '
        'lblStep2
        '
        resources.ApplyResources(Me.lblStep2, "lblStep2")
        Me.lblStep2.Name = "lblStep2"
        '
        'TabPageSelectColumns
        '
        resources.ApplyResources(Me.TabPageSelectColumns, "TabPageSelectColumns")
        Me.TabPageSelectColumns.Controls.Add(Me.btnToStep4)
        Me.TabPageSelectColumns.Controls.Add(Me.TableLayoutPanelColumns)
        Me.TabPageSelectColumns.Controls.Add(Me.lblStep3)
        Me.TabPageSelectColumns.Name = "TabPageSelectColumns"
        Me.TabPageSelectColumns.UseVisualStyleBackColor = True
        '
        'btnToStep4
        '
        resources.ApplyResources(Me.btnToStep4, "btnToStep4")
        Me.btnToStep4.Name = "btnToStep4"
        Me.btnToStep4.UseVisualStyleBackColor = True
        '
        'TableLayoutPanelColumns
        '
        resources.ApplyResources(Me.TableLayoutPanelColumns, "TableLayoutPanelColumns")
        Me.TableLayoutPanelColumns.Controls.Add(Me.cmbColumnDuration, 1, 1)
        Me.TableLayoutPanelColumns.Controls.Add(Me.cmbColumnDescription, 0, 1)
        Me.TableLayoutPanelColumns.Controls.Add(Me.cmbColumnNumber, 2, 1)
        Me.TableLayoutPanelColumns.Controls.Add(Me.cmbColumnUnitCost, 3, 1)
        Me.TableLayoutPanelColumns.Controls.Add(Me.cmbColumnTotalCost, 4, 1)
        Me.TableLayoutPanelColumns.Controls.Add(Me.lblTotalCost, 4, 0)
        Me.TableLayoutPanelColumns.Controls.Add(Me.lblUnitCost, 3, 0)
        Me.TableLayoutPanelColumns.Controls.Add(Me.lblNumber, 2, 0)
        Me.TableLayoutPanelColumns.Controls.Add(Me.lblDuration, 1, 0)
        Me.TableLayoutPanelColumns.Controls.Add(Me.lblDescription, 0, 0)
        Me.TableLayoutPanelColumns.Name = "TableLayoutPanelColumns"
        '
        'cmbColumnDuration
        '
        resources.ApplyResources(Me.cmbColumnDuration, "cmbColumnDuration")
        Me.cmbColumnDuration.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnDuration.FormattingEnabled = True
        Me.cmbColumnDuration.Name = "cmbColumnDuration"
        '
        'cmbColumnDescription
        '
        resources.ApplyResources(Me.cmbColumnDescription, "cmbColumnDescription")
        Me.cmbColumnDescription.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnDescription.FormattingEnabled = True
        Me.cmbColumnDescription.Name = "cmbColumnDescription"
        '
        'cmbColumnNumber
        '
        resources.ApplyResources(Me.cmbColumnNumber, "cmbColumnNumber")
        Me.cmbColumnNumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnNumber.FormattingEnabled = True
        Me.cmbColumnNumber.Name = "cmbColumnNumber"
        '
        'cmbColumnUnitCost
        '
        resources.ApplyResources(Me.cmbColumnUnitCost, "cmbColumnUnitCost")
        Me.cmbColumnUnitCost.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnUnitCost.FormattingEnabled = True
        Me.cmbColumnUnitCost.Name = "cmbColumnUnitCost"
        '
        'cmbColumnTotalCost
        '
        resources.ApplyResources(Me.cmbColumnTotalCost, "cmbColumnTotalCost")
        Me.cmbColumnTotalCost.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnTotalCost.FormattingEnabled = True
        Me.cmbColumnTotalCost.Name = "cmbColumnTotalCost"
        '
        'lblTotalCost
        '
        resources.ApplyResources(Me.lblTotalCost, "lblTotalCost")
        Me.lblTotalCost.Name = "lblTotalCost"
        '
        'lblUnitCost
        '
        resources.ApplyResources(Me.lblUnitCost, "lblUnitCost")
        Me.lblUnitCost.Name = "lblUnitCost"
        '
        'lblNumber
        '
        resources.ApplyResources(Me.lblNumber, "lblNumber")
        Me.lblNumber.Name = "lblNumber"
        '
        'lblDuration
        '
        resources.ApplyResources(Me.lblDuration, "lblDuration")
        Me.lblDuration.Name = "lblDuration"
        '
        'lblDescription
        '
        resources.ApplyResources(Me.lblDescription, "lblDescription")
        Me.lblDescription.Name = "lblDescription"
        '
        'lblStep3
        '
        resources.ApplyResources(Me.lblStep3, "lblStep3")
        Me.lblStep3.Name = "lblStep3"
        '
        'TabPageSelectRows
        '
        resources.ApplyResources(Me.TabPageSelectRows, "TabPageSelectRows")
        Me.TabPageSelectRows.Controls.Add(Me.TableLayoutPanel2)
        Me.TabPageSelectRows.Controls.Add(Me.lblStep4)
        Me.TabPageSelectRows.Name = "TabPageSelectRows"
        Me.TabPageSelectRows.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        resources.ApplyResources(Me.TableLayoutPanel2, "TableLayoutPanel2")
        Me.TableLayoutPanel2.Controls.Add(Me.lblFirstRow, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbFirstRow, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblLastRow, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbLastRow, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        '
        'lblFirstRow
        '
        resources.ApplyResources(Me.lblFirstRow, "lblFirstRow")
        Me.lblFirstRow.Name = "lblFirstRow"
        '
        'ntbFirstRow
        '
        resources.ApplyResources(Me.ntbFirstRow, "ntbFirstRow")
        Me.ntbFirstRow.AllowSpace = True
        Me.ntbFirstRow.DoubleValue = 0.0R
        Me.ntbFirstRow.IntegerValue = 0
        Me.ntbFirstRow.IsCurrency = False
        Me.ntbFirstRow.IsPercentage = False
        Me.ntbFirstRow.Name = "ntbFirstRow"
        Me.ntbFirstRow.NrDecimals = 0
        Me.ntbFirstRow.SetDecimals = True
        Me.ntbFirstRow.SingleValue = 0.0!
        Me.ntbFirstRow.Unit = Nothing
        Me.ntbFirstRow.ValueType = 0
        '
        'lblLastRow
        '
        resources.ApplyResources(Me.lblLastRow, "lblLastRow")
        Me.lblLastRow.Name = "lblLastRow"
        '
        'ntbLastRow
        '
        resources.ApplyResources(Me.ntbLastRow, "ntbLastRow")
        Me.ntbLastRow.AllowSpace = True
        Me.ntbLastRow.DoubleValue = 0.0R
        Me.ntbLastRow.IntegerValue = 0
        Me.ntbLastRow.IsCurrency = False
        Me.ntbLastRow.IsPercentage = False
        Me.ntbLastRow.Name = "ntbLastRow"
        Me.ntbLastRow.NrDecimals = 0
        Me.ntbLastRow.SetDecimals = False
        Me.ntbLastRow.SingleValue = 0.0!
        Me.ntbLastRow.Unit = Nothing
        Me.ntbLastRow.ValueType = 0
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'lblStep4
        '
        resources.ApplyResources(Me.lblStep4, "lblStep4")
        Me.lblStep4.Name = "lblStep4"
        '
        'dgWorkSheet
        '
        resources.ApplyResources(Me.dgWorkSheet, "dgWorkSheet")
        Me.dgWorkSheet.AllowUserToAddRows = False
        Me.dgWorkSheet.AllowUserToDeleteRows = False
        Me.dgWorkSheet.BackgroundColor = System.Drawing.Color.White
        Me.dgWorkSheet.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgWorkSheet.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgWorkSheet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgWorkSheet.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgWorkSheet.MultiSelect = False
        Me.dgWorkSheet.Name = "dgWorkSheet"
        Me.dgWorkSheet.ReadOnly = True
        Me.dgWorkSheet.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgWorkSheet.ShowCellToolTips = False
        Me.dgWorkSheet.WorkSheet = Nothing
        '
        'DialogImportBudget
        '
        Me.AcceptButton = Me.btnImport
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.dgWorkSheet)
        Me.Controls.Add(Me.TabControlSteps)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogImportBudget"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControlSteps.ResumeLayout(False)
        Me.TabPageSelectDocument.ResumeLayout(False)
        Me.TabPageSelectDocument.PerformLayout()
        Me.TabPageSelectSpreadsheet.ResumeLayout(False)
        Me.TabPageSelectSpreadsheet.PerformLayout()
        Me.TabPageSelectColumns.ResumeLayout(False)
        Me.TabPageSelectColumns.PerformLayout()
        Me.TableLayoutPanelColumns.ResumeLayout(False)
        Me.TableLayoutPanelColumns.PerformLayout()
        Me.TabPageSelectRows.ResumeLayout(False)
        Me.TabPageSelectRows.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        CType(Me.dgWorkSheet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents TabControlSteps As System.Windows.Forms.TabControl
    Friend WithEvents TabPageSelectDocument As System.Windows.Forms.TabPage
    Friend WithEvents TabPageSelectSpreadsheet As System.Windows.Forms.TabPage
    Friend WithEvents TabPageSelectColumns As System.Windows.Forms.TabPage
    Friend WithEvents TabPageSelectRows As System.Windows.Forms.TabPage
    Friend WithEvents btnLoadFile As System.Windows.Forms.Button
    Friend WithEvents lblFileName As System.Windows.Forms.Label
    Friend WithEvents tbFileName As System.Windows.Forms.TextBox
    Friend WithEvents btnToStep2 As System.Windows.Forms.Button
    Friend WithEvents lblStep1 As System.Windows.Forms.Label
    Friend WithEvents lblStep2 As System.Windows.Forms.Label
    Friend WithEvents cmbSelectWorksheet As System.Windows.Forms.ComboBox
    Friend WithEvents lblSelectWorksheet As System.Windows.Forms.Label
    Friend WithEvents lblStep3 As System.Windows.Forms.Label
    Friend WithEvents btnToStep3 As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanelColumns As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lblNumber As System.Windows.Forms.Label
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents dgWorkSheet As FaciliDev.LogFramer.DataGridViewExcelPreview
    Friend WithEvents btnToStep4 As System.Windows.Forms.Button
    Friend WithEvents lblStep4 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lblFirstRow As System.Windows.Forms.Label
    Friend WithEvents ntbFirstRow As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbLastRow As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblLastRow As System.Windows.Forms.Label
    Friend WithEvents cmbColumnDuration As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnDescription As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnNumber As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnUnitCost As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnTotalCost As System.Windows.Forms.ComboBox
    Friend WithEvents lblTotalCost As System.Windows.Forms.Label
    Friend WithEvents lblUnitCost As System.Windows.Forms.Label

End Class
