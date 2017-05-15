<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogImportLogframe
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogImportLogframe))
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
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.cmbColumnIndicators = New System.Windows.Forms.ComboBox()
        Me.cmbColumnVerificationSources = New System.Windows.Forms.ComboBox()
        Me.cmbColumnAssumptions = New System.Windows.Forms.ComboBox()
        Me.cmbColumnObjectives = New System.Windows.Forms.ComboBox()
        Me.lblObjectives = New System.Windows.Forms.Label()
        Me.lblStep3 = New System.Windows.Forms.Label()
        Me.TabPageSelectRows = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.cmbMeansBudget = New System.Windows.Forms.ComboBox()
        Me.lblFirstRow = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.ntbFirstRowGoals = New FaciliDev.LogFramer.NumericTextBox()
        Me.lblLastRow = New System.Windows.Forms.Label()
        Me.ntbLastRowGoals = New FaciliDev.LogFramer.NumericTextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ntbFirstRowPurposes = New FaciliDev.LogFramer.NumericTextBox()
        Me.ntbLastRowPurposes = New FaciliDev.LogFramer.NumericTextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ntbFirstRowOutputs = New FaciliDev.LogFramer.NumericTextBox()
        Me.ntbLastRowOutputs = New FaciliDev.LogFramer.NumericTextBox()
        Me.ntbFirstRowActivities = New FaciliDev.LogFramer.NumericTextBox()
        Me.ntbLastRowActivities = New FaciliDev.LogFramer.NumericTextBox()
        Me.lblStep4 = New System.Windows.Forms.Label()
        Me.dgWorkSheet = New FaciliDev.LogFramer.DataGridViewExcelPreview()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControlSteps.SuspendLayout()
        Me.TabPageSelectDocument.SuspendLayout()
        Me.TabPageSelectSpreadsheet.SuspendLayout()
        Me.TabPageSelectColumns.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
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
        Me.TabPageSelectColumns.Controls.Add(Me.TableLayoutPanel3)
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
        'TableLayoutPanel3
        '
        resources.ApplyResources(Me.TableLayoutPanel3, "TableLayoutPanel3")
        Me.TableLayoutPanel3.Controls.Add(Me.Label20, 3, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Label19, 2, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Label18, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.cmbColumnIndicators, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.cmbColumnVerificationSources, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.cmbColumnAssumptions, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.cmbColumnObjectives, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.lblObjectives, 0, 0)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        '
        'Label20
        '
        resources.ApplyResources(Me.Label20, "Label20")
        Me.Label20.Name = "Label20"
        '
        'Label19
        '
        resources.ApplyResources(Me.Label19, "Label19")
        Me.Label19.Name = "Label19"
        '
        'Label18
        '
        resources.ApplyResources(Me.Label18, "Label18")
        Me.Label18.Name = "Label18"
        '
        'cmbColumnIndicators
        '
        resources.ApplyResources(Me.cmbColumnIndicators, "cmbColumnIndicators")
        Me.cmbColumnIndicators.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnIndicators.FormattingEnabled = True
        Me.cmbColumnIndicators.Name = "cmbColumnIndicators"
        '
        'cmbColumnVerificationSources
        '
        resources.ApplyResources(Me.cmbColumnVerificationSources, "cmbColumnVerificationSources")
        Me.cmbColumnVerificationSources.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnVerificationSources.FormattingEnabled = True
        Me.cmbColumnVerificationSources.Name = "cmbColumnVerificationSources"
        '
        'cmbColumnAssumptions
        '
        resources.ApplyResources(Me.cmbColumnAssumptions, "cmbColumnAssumptions")
        Me.cmbColumnAssumptions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnAssumptions.FormattingEnabled = True
        Me.cmbColumnAssumptions.Name = "cmbColumnAssumptions"
        '
        'cmbColumnObjectives
        '
        resources.ApplyResources(Me.cmbColumnObjectives, "cmbColumnObjectives")
        Me.cmbColumnObjectives.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbColumnObjectives.FormattingEnabled = True
        Me.cmbColumnObjectives.Name = "cmbColumnObjectives"
        '
        'lblObjectives
        '
        resources.ApplyResources(Me.lblObjectives, "lblObjectives")
        Me.lblObjectives.Name = "lblObjectives"
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
        Me.TableLayoutPanel2.Controls.Add(Me.cmbMeansBudget, 4, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.lblFirstRow, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label15, 3, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbFirstRowGoals, 1, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.lblLastRow, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbLastRowGoals, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Label10, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.Label11, 0, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbFirstRowPurposes, 1, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbLastRowPurposes, 2, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbFirstRowOutputs, 1, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbLastRowOutputs, 2, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbFirstRowActivities, 1, 4)
        Me.TableLayoutPanel2.Controls.Add(Me.ntbLastRowActivities, 2, 4)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        '
        'cmbMeansBudget
        '
        resources.ApplyResources(Me.cmbMeansBudget, "cmbMeansBudget")
        Me.cmbMeansBudget.FormattingEnabled = True
        Me.cmbMeansBudget.Name = "cmbMeansBudget"
        '
        'lblFirstRow
        '
        resources.ApplyResources(Me.lblFirstRow, "lblFirstRow")
        Me.lblFirstRow.Name = "lblFirstRow"
        '
        'Label15
        '
        resources.ApplyResources(Me.Label15, "Label15")
        Me.Label15.Name = "Label15"
        '
        'ntbFirstRowGoals
        '
        resources.ApplyResources(Me.ntbFirstRowGoals, "ntbFirstRowGoals")
        Me.ntbFirstRowGoals.AllowSpace = True
        Me.ntbFirstRowGoals.DoubleValue = 0.0R
        Me.ntbFirstRowGoals.IntegerValue = 0
        Me.ntbFirstRowGoals.IsCurrency = False
        Me.ntbFirstRowGoals.IsPercentage = False
        Me.ntbFirstRowGoals.Name = "ntbFirstRowGoals"
        Me.ntbFirstRowGoals.NrDecimals = 0
        Me.ntbFirstRowGoals.SetDecimals = True
        Me.ntbFirstRowGoals.SingleValue = 0.0!
        Me.ntbFirstRowGoals.Unit = Nothing
        Me.ntbFirstRowGoals.ValueType = 0
        '
        'lblLastRow
        '
        resources.ApplyResources(Me.lblLastRow, "lblLastRow")
        Me.lblLastRow.Name = "lblLastRow"
        '
        'ntbLastRowGoals
        '
        resources.ApplyResources(Me.ntbLastRowGoals, "ntbLastRowGoals")
        Me.ntbLastRowGoals.AllowSpace = True
        Me.ntbLastRowGoals.DoubleValue = 0.0R
        Me.ntbLastRowGoals.IntegerValue = 0
        Me.ntbLastRowGoals.IsCurrency = False
        Me.ntbLastRowGoals.IsPercentage = False
        Me.ntbLastRowGoals.Name = "ntbLastRowGoals"
        Me.ntbLastRowGoals.NrDecimals = 0
        Me.ntbLastRowGoals.SetDecimals = False
        Me.ntbLastRowGoals.SingleValue = 0.0!
        Me.ntbLastRowGoals.Unit = Nothing
        Me.ntbLastRowGoals.ValueType = 0
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.Name = "Label10"
        '
        'Label11
        '
        resources.ApplyResources(Me.Label11, "Label11")
        Me.Label11.Name = "Label11"
        '
        'ntbFirstRowPurposes
        '
        resources.ApplyResources(Me.ntbFirstRowPurposes, "ntbFirstRowPurposes")
        Me.ntbFirstRowPurposes.AllowSpace = True
        Me.ntbFirstRowPurposes.DoubleValue = 0.0R
        Me.ntbFirstRowPurposes.IntegerValue = 0
        Me.ntbFirstRowPurposes.IsCurrency = False
        Me.ntbFirstRowPurposes.IsPercentage = False
        Me.ntbFirstRowPurposes.Name = "ntbFirstRowPurposes"
        Me.ntbFirstRowPurposes.NrDecimals = 0
        Me.ntbFirstRowPurposes.SetDecimals = False
        Me.ntbFirstRowPurposes.SingleValue = 0.0!
        Me.ntbFirstRowPurposes.Unit = Nothing
        Me.ntbFirstRowPurposes.ValueType = 0
        '
        'ntbLastRowPurposes
        '
        resources.ApplyResources(Me.ntbLastRowPurposes, "ntbLastRowPurposes")
        Me.ntbLastRowPurposes.AllowSpace = True
        Me.ntbLastRowPurposes.DoubleValue = 0.0R
        Me.ntbLastRowPurposes.IntegerValue = 0
        Me.ntbLastRowPurposes.IsCurrency = False
        Me.ntbLastRowPurposes.IsPercentage = False
        Me.ntbLastRowPurposes.Name = "ntbLastRowPurposes"
        Me.ntbLastRowPurposes.NrDecimals = 0
        Me.ntbLastRowPurposes.SetDecimals = False
        Me.ntbLastRowPurposes.SingleValue = 0.0!
        Me.ntbLastRowPurposes.Unit = Nothing
        Me.ntbLastRowPurposes.ValueType = 0
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'ntbFirstRowOutputs
        '
        resources.ApplyResources(Me.ntbFirstRowOutputs, "ntbFirstRowOutputs")
        Me.ntbFirstRowOutputs.AllowSpace = True
        Me.ntbFirstRowOutputs.DoubleValue = 0.0R
        Me.ntbFirstRowOutputs.IntegerValue = 0
        Me.ntbFirstRowOutputs.IsCurrency = False
        Me.ntbFirstRowOutputs.IsPercentage = False
        Me.ntbFirstRowOutputs.Name = "ntbFirstRowOutputs"
        Me.ntbFirstRowOutputs.NrDecimals = 0
        Me.ntbFirstRowOutputs.SetDecimals = False
        Me.ntbFirstRowOutputs.SingleValue = 0.0!
        Me.ntbFirstRowOutputs.Unit = Nothing
        Me.ntbFirstRowOutputs.ValueType = 0
        '
        'ntbLastRowOutputs
        '
        resources.ApplyResources(Me.ntbLastRowOutputs, "ntbLastRowOutputs")
        Me.ntbLastRowOutputs.AllowSpace = True
        Me.ntbLastRowOutputs.DoubleValue = 0.0R
        Me.ntbLastRowOutputs.IntegerValue = 0
        Me.ntbLastRowOutputs.IsCurrency = False
        Me.ntbLastRowOutputs.IsPercentage = False
        Me.ntbLastRowOutputs.Name = "ntbLastRowOutputs"
        Me.ntbLastRowOutputs.NrDecimals = 0
        Me.ntbLastRowOutputs.SetDecimals = False
        Me.ntbLastRowOutputs.SingleValue = 0.0!
        Me.ntbLastRowOutputs.Unit = Nothing
        Me.ntbLastRowOutputs.ValueType = 0
        '
        'ntbFirstRowActivities
        '
        resources.ApplyResources(Me.ntbFirstRowActivities, "ntbFirstRowActivities")
        Me.ntbFirstRowActivities.AllowSpace = True
        Me.ntbFirstRowActivities.DoubleValue = 0.0R
        Me.ntbFirstRowActivities.IntegerValue = 0
        Me.ntbFirstRowActivities.IsCurrency = False
        Me.ntbFirstRowActivities.IsPercentage = False
        Me.ntbFirstRowActivities.Name = "ntbFirstRowActivities"
        Me.ntbFirstRowActivities.NrDecimals = 0
        Me.ntbFirstRowActivities.SetDecimals = False
        Me.ntbFirstRowActivities.SingleValue = 0.0!
        Me.ntbFirstRowActivities.Unit = Nothing
        Me.ntbFirstRowActivities.ValueType = 0
        '
        'ntbLastRowActivities
        '
        resources.ApplyResources(Me.ntbLastRowActivities, "ntbLastRowActivities")
        Me.ntbLastRowActivities.AllowSpace = True
        Me.ntbLastRowActivities.DoubleValue = 0.0R
        Me.ntbLastRowActivities.IntegerValue = 0
        Me.ntbLastRowActivities.IsCurrency = False
        Me.ntbLastRowActivities.IsPercentage = False
        Me.ntbLastRowActivities.Name = "ntbLastRowActivities"
        Me.ntbLastRowActivities.NrDecimals = 0
        Me.ntbLastRowActivities.SetDecimals = False
        Me.ntbLastRowActivities.SingleValue = 0.0!
        Me.ntbLastRowActivities.Unit = Nothing
        Me.ntbLastRowActivities.ValueType = 0
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
        'DialogImportLogframe
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
        Me.Name = "DialogImportLogframe"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControlSteps.ResumeLayout(False)
        Me.TabPageSelectDocument.ResumeLayout(False)
        Me.TabPageSelectDocument.PerformLayout()
        Me.TabPageSelectSpreadsheet.ResumeLayout(False)
        Me.TabPageSelectSpreadsheet.PerformLayout()
        Me.TabPageSelectColumns.ResumeLayout(False)
        Me.TabPageSelectColumns.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
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
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cmbColumnIndicators As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnVerificationSources As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnAssumptions As System.Windows.Forms.ComboBox
    Friend WithEvents cmbColumnObjectives As System.Windows.Forms.ComboBox
    Friend WithEvents lblObjectives As System.Windows.Forms.Label
    Friend WithEvents dgWorkSheet As FaciliDev.LogFramer.DataGridViewExcelPreview
    Friend WithEvents btnToStep4 As System.Windows.Forms.Button
    Friend WithEvents lblStep4 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lblFirstRow As System.Windows.Forms.Label
    Friend WithEvents ntbFirstRowGoals As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbLastRowGoals As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ntbFirstRowPurposes As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbFirstRowOutputs As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbLastRowPurposes As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbLastRowOutputs As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbMeansBudget As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents lblLastRow As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ntbFirstRowActivities As FaciliDev.LogFramer.NumericTextBox
    Friend WithEvents ntbLastRowActivities As FaciliDev.LogFramer.NumericTextBox

End Class
