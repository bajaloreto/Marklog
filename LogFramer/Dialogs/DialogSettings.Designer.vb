<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogSettings))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.TabControlSettings = New System.Windows.Forms.TabControl()
        Me.TabPageDesign = New System.Windows.Forms.TabPage()
        Me.lblAssumptionName = New System.Windows.Forms.Label()
        Me.tbAssumptionName = New System.Windows.Forms.TextBox()
        Me.lblBudgetName = New System.Windows.Forms.Label()
        Me.tbBudgetName = New System.Windows.Forms.TextBox()
        Me.lblResourceName = New System.Windows.Forms.Label()
        Me.tbResourceName = New System.Windows.Forms.TextBox()
        Me.lblVerificationSourceName = New System.Windows.Forms.Label()
        Me.tbVerificationSourceName = New System.Windows.Forms.TextBox()
        Me.lblIndicatorName = New System.Windows.Forms.Label()
        Me.tbIndicatorName = New System.Windows.Forms.TextBox()
        Me.chkDefaultScheme = New System.Windows.Forms.CheckBox()
        Me.lblUse = New System.Windows.Forms.Label()
        Me.dgvProjectLogic = New FaciliDev.LogFramer.DataGridViewProjectLogic()
        Me.TabPageCurrentLogFrame = New System.Windows.Forms.TabPage()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.cmbCurrency = New System.Windows.Forms.ComboBox()
        Me.lblCurrency = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.tbDetailsFont = New System.Windows.Forms.TextBox()
        Me.lblDetailsFont = New System.Windows.Forms.Label()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.tbSectionTitleFont = New System.Windows.Forms.TextBox()
        Me.lblSectionTitleFont = New System.Windows.Forms.Label()
        Me.btnEqualColorLf = New System.Windows.Forms.Button()
        Me.lblSectionColor = New System.Windows.Forms.Label()
        Me.tbSectionTitleFontColor = New System.Windows.Forms.TextBox()
        Me.tbSectionColorTop = New System.Windows.Forms.TextBox()
        Me.lblSectionTitleFontColor = New System.Windows.Forms.Label()
        Me.tbSectionColorBottom = New System.Windows.Forms.TextBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.lblSortNumberExample = New System.Windows.Forms.Label()
        Me.chkSortNumberRepeatParent = New System.Windows.Forms.CheckBox()
        Me.chkSortNumberTerminateWithDivider = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbSortNumberDivider = New System.Windows.Forms.TextBox()
        Me.tbSortNumberFont = New System.Windows.Forms.TextBox()
        Me.lblSortNumberFont = New System.Windows.Forms.Label()
        Me.TabPageGeneral = New System.Windows.Forms.TabPage()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.btnEqualColor = New System.Windows.Forms.Button()
        Me.tbDefaultFontSectionColor = New System.Windows.Forms.TextBox()
        Me.lblDefaultFontSectionColor = New System.Windows.Forms.Label()
        Me.chkRepeatPurposes = New System.Windows.Forms.CheckBox()
        Me.tbDefaultColorSectionBottom = New System.Windows.Forms.TextBox()
        Me.chkRepeatOutputs = New System.Windows.Forms.CheckBox()
        Me.tbDefaultColorSectionTop = New System.Windows.Forms.TextBox()
        Me.lblDefaultFontSections = New System.Windows.Forms.Label()
        Me.lblDefaultColorSection = New System.Windows.Forms.Label()
        Me.tbDefaultFontSections = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.lblAutoSaveMinutes = New System.Windows.Forms.Label()
        Me.ntbTimerAutoSave = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.lblTimerAutoSave = New System.Windows.Forms.Label()
        Me.ntbRecentFilesMax = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.lblRecentFilesMax = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmbDefaultIndType = New FaciliDev.LogFramer.StructuredComboBox()
        Me.tbDefaultFontUI = New System.Windows.Forms.TextBox()
        Me.lblDefaultCurrency = New System.Windows.Forms.Label()
        Me.lblDefaultFontUI = New System.Windows.Forms.Label()
        Me.cmbDefaultCurrency = New System.Windows.Forms.ComboBox()
        Me.lblDefaultIndType = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.tbDefaultFont = New System.Windows.Forms.TextBox()
        Me.lblDefaultFont = New System.Windows.Forms.Label()
        Me.tbDefaultFontSortNumbers = New System.Windows.Forms.TextBox()
        Me.chkWarnLinkedObjDelete = New System.Windows.Forms.CheckBox()
        Me.lblDefaultFontSortNumbers = New System.Windows.Forms.Label()
        Me.FontDialogSettings = New System.Windows.Forms.FontDialog()
        Me.ColorDialogSettings = New System.Windows.Forms.ColorDialog()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControlSettings.SuspendLayout()
        Me.TabPageDesign.SuspendLayout()
        CType(Me.dgvProjectLogic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageCurrentLogFrame.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.TabPageGeneral.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
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
        'TabControlSettings
        '
        resources.ApplyResources(Me.TabControlSettings, "TabControlSettings")
        Me.TabControlSettings.Controls.Add(Me.TabPageDesign)
        Me.TabControlSettings.Controls.Add(Me.TabPageCurrentLogFrame)
        Me.TabControlSettings.Controls.Add(Me.TabPageGeneral)
        Me.TabControlSettings.Name = "TabControlSettings"
        Me.TabControlSettings.SelectedIndex = 0
        '
        'TabPageDesign
        '
        resources.ApplyResources(Me.TabPageDesign, "TabPageDesign")
        Me.TabPageDesign.Controls.Add(Me.lblAssumptionName)
        Me.TabPageDesign.Controls.Add(Me.tbAssumptionName)
        Me.TabPageDesign.Controls.Add(Me.lblBudgetName)
        Me.TabPageDesign.Controls.Add(Me.tbBudgetName)
        Me.TabPageDesign.Controls.Add(Me.lblResourceName)
        Me.TabPageDesign.Controls.Add(Me.tbResourceName)
        Me.TabPageDesign.Controls.Add(Me.lblVerificationSourceName)
        Me.TabPageDesign.Controls.Add(Me.tbVerificationSourceName)
        Me.TabPageDesign.Controls.Add(Me.lblIndicatorName)
        Me.TabPageDesign.Controls.Add(Me.tbIndicatorName)
        Me.TabPageDesign.Controls.Add(Me.chkDefaultScheme)
        Me.TabPageDesign.Controls.Add(Me.lblUse)
        Me.TabPageDesign.Controls.Add(Me.dgvProjectLogic)
        Me.TabPageDesign.Name = "TabPageDesign"
        Me.TabPageDesign.UseVisualStyleBackColor = True
        '
        'lblAssumptionName
        '
        resources.ApplyResources(Me.lblAssumptionName, "lblAssumptionName")
        Me.lblAssumptionName.Name = "lblAssumptionName"
        '
        'tbAssumptionName
        '
        resources.ApplyResources(Me.tbAssumptionName, "tbAssumptionName")
        Me.tbAssumptionName.Name = "tbAssumptionName"
        '
        'lblBudgetName
        '
        resources.ApplyResources(Me.lblBudgetName, "lblBudgetName")
        Me.lblBudgetName.Name = "lblBudgetName"
        '
        'tbBudgetName
        '
        resources.ApplyResources(Me.tbBudgetName, "tbBudgetName")
        Me.tbBudgetName.Name = "tbBudgetName"
        '
        'lblResourceName
        '
        resources.ApplyResources(Me.lblResourceName, "lblResourceName")
        Me.lblResourceName.Name = "lblResourceName"
        '
        'tbResourceName
        '
        resources.ApplyResources(Me.tbResourceName, "tbResourceName")
        Me.tbResourceName.Name = "tbResourceName"
        '
        'lblVerificationSourceName
        '
        resources.ApplyResources(Me.lblVerificationSourceName, "lblVerificationSourceName")
        Me.lblVerificationSourceName.Name = "lblVerificationSourceName"
        '
        'tbVerificationSourceName
        '
        resources.ApplyResources(Me.tbVerificationSourceName, "tbVerificationSourceName")
        Me.tbVerificationSourceName.Name = "tbVerificationSourceName"
        '
        'lblIndicatorName
        '
        resources.ApplyResources(Me.lblIndicatorName, "lblIndicatorName")
        Me.lblIndicatorName.Name = "lblIndicatorName"
        '
        'tbIndicatorName
        '
        resources.ApplyResources(Me.tbIndicatorName, "tbIndicatorName")
        Me.tbIndicatorName.Name = "tbIndicatorName"
        '
        'chkDefaultScheme
        '
        resources.ApplyResources(Me.chkDefaultScheme, "chkDefaultScheme")
        Me.chkDefaultScheme.Name = "chkDefaultScheme"
        Me.chkDefaultScheme.UseVisualStyleBackColor = True
        '
        'lblUse
        '
        resources.ApplyResources(Me.lblUse, "lblUse")
        Me.lblUse.Name = "lblUse"
        '
        'dgvProjectLogic
        '
        resources.ApplyResources(Me.dgvProjectLogic, "dgvProjectLogic")
        Me.dgvProjectLogic.AllowUserToAddRows = False
        Me.dgvProjectLogic.AllowUserToDeleteRows = False
        Me.dgvProjectLogic.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvProjectLogic.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvProjectLogic.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvProjectLogic.MultiSelect = False
        Me.dgvProjectLogic.Name = "dgvProjectLogic"
        Me.dgvProjectLogic.ReadOnly = True
        Me.dgvProjectLogic.RowHeadersVisible = False
        Me.dgvProjectLogic.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullColumnSelect
        '
        'TabPageCurrentLogFrame
        '
        resources.ApplyResources(Me.TabPageCurrentLogFrame, "TabPageCurrentLogFrame")
        Me.TabPageCurrentLogFrame.Controls.Add(Me.Panel7)
        Me.TabPageCurrentLogFrame.Controls.Add(Me.Panel6)
        Me.TabPageCurrentLogFrame.Controls.Add(Me.Panel5)
        Me.TabPageCurrentLogFrame.Controls.Add(Me.Panel4)
        Me.TabPageCurrentLogFrame.Name = "TabPageCurrentLogFrame"
        Me.TabPageCurrentLogFrame.UseVisualStyleBackColor = True
        '
        'Panel7
        '
        resources.ApplyResources(Me.Panel7, "Panel7")
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel7.Controls.Add(Me.cmbCurrency)
        Me.Panel7.Controls.Add(Me.lblCurrency)
        Me.Panel7.Name = "Panel7"
        '
        'cmbCurrency
        '
        resources.ApplyResources(Me.cmbCurrency, "cmbCurrency")
        Me.cmbCurrency.FormattingEnabled = True
        Me.cmbCurrency.Name = "cmbCurrency"
        '
        'lblCurrency
        '
        resources.ApplyResources(Me.lblCurrency, "lblCurrency")
        Me.lblCurrency.Name = "lblCurrency"
        '
        'Panel6
        '
        resources.ApplyResources(Me.Panel6, "Panel6")
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Controls.Add(Me.tbDetailsFont)
        Me.Panel6.Controls.Add(Me.lblDetailsFont)
        Me.Panel6.Name = "Panel6"
        '
        'tbDetailsFont
        '
        resources.ApplyResources(Me.tbDetailsFont, "tbDetailsFont")
        Me.tbDetailsFont.Name = "tbDetailsFont"
        Me.tbDetailsFont.ReadOnly = True
        '
        'lblDetailsFont
        '
        resources.ApplyResources(Me.lblDetailsFont, "lblDetailsFont")
        Me.lblDetailsFont.Name = "lblDetailsFont"
        '
        'Panel5
        '
        resources.ApplyResources(Me.Panel5, "Panel5")
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Controls.Add(Me.tbSectionTitleFont)
        Me.Panel5.Controls.Add(Me.lblSectionTitleFont)
        Me.Panel5.Controls.Add(Me.btnEqualColorLf)
        Me.Panel5.Controls.Add(Me.lblSectionColor)
        Me.Panel5.Controls.Add(Me.tbSectionTitleFontColor)
        Me.Panel5.Controls.Add(Me.tbSectionColorTop)
        Me.Panel5.Controls.Add(Me.lblSectionTitleFontColor)
        Me.Panel5.Controls.Add(Me.tbSectionColorBottom)
        Me.Panel5.Name = "Panel5"
        '
        'tbSectionTitleFont
        '
        resources.ApplyResources(Me.tbSectionTitleFont, "tbSectionTitleFont")
        Me.tbSectionTitleFont.Name = "tbSectionTitleFont"
        Me.tbSectionTitleFont.ReadOnly = True
        '
        'lblSectionTitleFont
        '
        resources.ApplyResources(Me.lblSectionTitleFont, "lblSectionTitleFont")
        Me.lblSectionTitleFont.Name = "lblSectionTitleFont"
        '
        'btnEqualColorLf
        '
        resources.ApplyResources(Me.btnEqualColorLf, "btnEqualColorLf")
        Me.btnEqualColorLf.Name = "btnEqualColorLf"
        Me.btnEqualColorLf.UseVisualStyleBackColor = True
        '
        'lblSectionColor
        '
        resources.ApplyResources(Me.lblSectionColor, "lblSectionColor")
        Me.lblSectionColor.Name = "lblSectionColor"
        '
        'tbSectionTitleFontColor
        '
        resources.ApplyResources(Me.tbSectionTitleFontColor, "tbSectionTitleFontColor")
        Me.tbSectionTitleFontColor.Name = "tbSectionTitleFontColor"
        Me.tbSectionTitleFontColor.ReadOnly = True
        '
        'tbSectionColorTop
        '
        resources.ApplyResources(Me.tbSectionColorTop, "tbSectionColorTop")
        Me.tbSectionColorTop.Name = "tbSectionColorTop"
        Me.tbSectionColorTop.ReadOnly = True
        '
        'lblSectionTitleFontColor
        '
        resources.ApplyResources(Me.lblSectionTitleFontColor, "lblSectionTitleFontColor")
        Me.lblSectionTitleFontColor.Name = "lblSectionTitleFontColor"
        '
        'tbSectionColorBottom
        '
        resources.ApplyResources(Me.tbSectionColorBottom, "tbSectionColorBottom")
        Me.tbSectionColorBottom.Name = "tbSectionColorBottom"
        Me.tbSectionColorBottom.ReadOnly = True
        '
        'Panel4
        '
        resources.ApplyResources(Me.Panel4, "Panel4")
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.lblSortNumberExample)
        Me.Panel4.Controls.Add(Me.chkSortNumberRepeatParent)
        Me.Panel4.Controls.Add(Me.chkSortNumberTerminateWithDivider)
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Controls.Add(Me.tbSortNumberDivider)
        Me.Panel4.Controls.Add(Me.tbSortNumberFont)
        Me.Panel4.Controls.Add(Me.lblSortNumberFont)
        Me.Panel4.Name = "Panel4"
        '
        'lblSortNumberExample
        '
        resources.ApplyResources(Me.lblSortNumberExample, "lblSortNumberExample")
        Me.lblSortNumberExample.Name = "lblSortNumberExample"
        '
        'chkSortNumberRepeatParent
        '
        resources.ApplyResources(Me.chkSortNumberRepeatParent, "chkSortNumberRepeatParent")
        Me.chkSortNumberRepeatParent.Name = "chkSortNumberRepeatParent"
        Me.chkSortNumberRepeatParent.UseVisualStyleBackColor = True
        '
        'chkSortNumberTerminateWithDivider
        '
        resources.ApplyResources(Me.chkSortNumberTerminateWithDivider, "chkSortNumberTerminateWithDivider")
        Me.chkSortNumberTerminateWithDivider.Name = "chkSortNumberTerminateWithDivider"
        Me.chkSortNumberTerminateWithDivider.UseVisualStyleBackColor = True
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'tbSortNumberDivider
        '
        resources.ApplyResources(Me.tbSortNumberDivider, "tbSortNumberDivider")
        Me.tbSortNumberDivider.Name = "tbSortNumberDivider"
        '
        'tbSortNumberFont
        '
        resources.ApplyResources(Me.tbSortNumberFont, "tbSortNumberFont")
        Me.tbSortNumberFont.Name = "tbSortNumberFont"
        Me.tbSortNumberFont.ReadOnly = True
        '
        'lblSortNumberFont
        '
        resources.ApplyResources(Me.lblSortNumberFont, "lblSortNumberFont")
        Me.lblSortNumberFont.Name = "lblSortNumberFont"
        '
        'TabPageGeneral
        '
        resources.ApplyResources(Me.TabPageGeneral, "TabPageGeneral")
        Me.TabPageGeneral.Controls.Add(Me.Panel8)
        Me.TabPageGeneral.Controls.Add(Me.Panel3)
        Me.TabPageGeneral.Controls.Add(Me.Panel2)
        Me.TabPageGeneral.Controls.Add(Me.Panel1)
        Me.TabPageGeneral.Name = "TabPageGeneral"
        Me.TabPageGeneral.UseVisualStyleBackColor = True
        '
        'Panel8
        '
        resources.ApplyResources(Me.Panel8, "Panel8")
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel8.Controls.Add(Me.btnEqualColor)
        Me.Panel8.Controls.Add(Me.tbDefaultFontSectionColor)
        Me.Panel8.Controls.Add(Me.lblDefaultFontSectionColor)
        Me.Panel8.Controls.Add(Me.chkRepeatPurposes)
        Me.Panel8.Controls.Add(Me.tbDefaultColorSectionBottom)
        Me.Panel8.Controls.Add(Me.chkRepeatOutputs)
        Me.Panel8.Controls.Add(Me.tbDefaultColorSectionTop)
        Me.Panel8.Controls.Add(Me.lblDefaultFontSections)
        Me.Panel8.Controls.Add(Me.lblDefaultColorSection)
        Me.Panel8.Controls.Add(Me.tbDefaultFontSections)
        Me.Panel8.Name = "Panel8"
        '
        'btnEqualColor
        '
        resources.ApplyResources(Me.btnEqualColor, "btnEqualColor")
        Me.btnEqualColor.Name = "btnEqualColor"
        Me.btnEqualColor.UseVisualStyleBackColor = True
        '
        'tbDefaultFontSectionColor
        '
        resources.ApplyResources(Me.tbDefaultFontSectionColor, "tbDefaultFontSectionColor")
        Me.tbDefaultFontSectionColor.Name = "tbDefaultFontSectionColor"
        Me.tbDefaultFontSectionColor.ReadOnly = True
        '
        'lblDefaultFontSectionColor
        '
        resources.ApplyResources(Me.lblDefaultFontSectionColor, "lblDefaultFontSectionColor")
        Me.lblDefaultFontSectionColor.Name = "lblDefaultFontSectionColor"
        '
        'chkRepeatPurposes
        '
        resources.ApplyResources(Me.chkRepeatPurposes, "chkRepeatPurposes")
        Me.chkRepeatPurposes.Name = "chkRepeatPurposes"
        Me.chkRepeatPurposes.UseVisualStyleBackColor = True
        '
        'tbDefaultColorSectionBottom
        '
        resources.ApplyResources(Me.tbDefaultColorSectionBottom, "tbDefaultColorSectionBottom")
        Me.tbDefaultColorSectionBottom.Name = "tbDefaultColorSectionBottom"
        Me.tbDefaultColorSectionBottom.ReadOnly = True
        '
        'chkRepeatOutputs
        '
        resources.ApplyResources(Me.chkRepeatOutputs, "chkRepeatOutputs")
        Me.chkRepeatOutputs.Name = "chkRepeatOutputs"
        Me.chkRepeatOutputs.UseVisualStyleBackColor = True
        '
        'tbDefaultColorSectionTop
        '
        resources.ApplyResources(Me.tbDefaultColorSectionTop, "tbDefaultColorSectionTop")
        Me.tbDefaultColorSectionTop.Name = "tbDefaultColorSectionTop"
        Me.tbDefaultColorSectionTop.ReadOnly = True
        '
        'lblDefaultFontSections
        '
        resources.ApplyResources(Me.lblDefaultFontSections, "lblDefaultFontSections")
        Me.lblDefaultFontSections.Name = "lblDefaultFontSections"
        '
        'lblDefaultColorSection
        '
        resources.ApplyResources(Me.lblDefaultColorSection, "lblDefaultColorSection")
        Me.lblDefaultColorSection.Name = "lblDefaultColorSection"
        '
        'tbDefaultFontSections
        '
        resources.ApplyResources(Me.tbDefaultFontSections, "tbDefaultFontSections")
        Me.tbDefaultFontSections.Name = "tbDefaultFontSections"
        Me.tbDefaultFontSections.ReadOnly = True
        '
        'Panel3
        '
        resources.ApplyResources(Me.Panel3, "Panel3")
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.lblAutoSaveMinutes)
        Me.Panel3.Controls.Add(Me.ntbTimerAutoSave)
        Me.Panel3.Controls.Add(Me.lblTimerAutoSave)
        Me.Panel3.Controls.Add(Me.ntbRecentFilesMax)
        Me.Panel3.Controls.Add(Me.lblRecentFilesMax)
        Me.Panel3.Name = "Panel3"
        '
        'lblAutoSaveMinutes
        '
        resources.ApplyResources(Me.lblAutoSaveMinutes, "lblAutoSaveMinutes")
        Me.lblAutoSaveMinutes.Name = "lblAutoSaveMinutes"
        '
        'ntbTimerAutoSave
        '
        resources.ApplyResources(Me.ntbTimerAutoSave, "ntbTimerAutoSave")
        Me.ntbTimerAutoSave.AllowSpace = False
        Me.ntbTimerAutoSave.IsCurrency = False
        Me.ntbTimerAutoSave.IsPercentage = False
        Me.ntbTimerAutoSave.Name = "ntbTimerAutoSave"
        '
        'lblTimerAutoSave
        '
        resources.ApplyResources(Me.lblTimerAutoSave, "lblTimerAutoSave")
        Me.lblTimerAutoSave.Name = "lblTimerAutoSave"
        '
        'ntbRecentFilesMax
        '
        resources.ApplyResources(Me.ntbRecentFilesMax, "ntbRecentFilesMax")
        Me.ntbRecentFilesMax.AllowSpace = True
        Me.ntbRecentFilesMax.IsCurrency = False
        Me.ntbRecentFilesMax.IsPercentage = False
        Me.ntbRecentFilesMax.Name = "ntbRecentFilesMax"
        '
        'lblRecentFilesMax
        '
        resources.ApplyResources(Me.lblRecentFilesMax, "lblRecentFilesMax")
        Me.lblRecentFilesMax.Name = "lblRecentFilesMax"
        '
        'Panel2
        '
        resources.ApplyResources(Me.Panel2, "Panel2")
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cmbDefaultIndType)
        Me.Panel2.Controls.Add(Me.tbDefaultFontUI)
        Me.Panel2.Controls.Add(Me.lblDefaultCurrency)
        Me.Panel2.Controls.Add(Me.lblDefaultFontUI)
        Me.Panel2.Controls.Add(Me.cmbDefaultCurrency)
        Me.Panel2.Controls.Add(Me.lblDefaultIndType)
        Me.Panel2.Name = "Panel2"
        '
        'cmbDefaultIndType
        '
        resources.ApplyResources(Me.cmbDefaultIndType, "cmbDefaultIndType")
        Me.cmbDefaultIndType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.cmbDefaultIndType.FormattingEnabled = True
        Me.cmbDefaultIndType.Name = "cmbDefaultIndType"
        '
        'tbDefaultFontUI
        '
        resources.ApplyResources(Me.tbDefaultFontUI, "tbDefaultFontUI")
        Me.tbDefaultFontUI.Name = "tbDefaultFontUI"
        Me.tbDefaultFontUI.ReadOnly = True
        '
        'lblDefaultCurrency
        '
        resources.ApplyResources(Me.lblDefaultCurrency, "lblDefaultCurrency")
        Me.lblDefaultCurrency.Name = "lblDefaultCurrency"
        '
        'lblDefaultFontUI
        '
        resources.ApplyResources(Me.lblDefaultFontUI, "lblDefaultFontUI")
        Me.lblDefaultFontUI.Name = "lblDefaultFontUI"
        '
        'cmbDefaultCurrency
        '
        resources.ApplyResources(Me.cmbDefaultCurrency, "cmbDefaultCurrency")
        Me.cmbDefaultCurrency.FormattingEnabled = True
        Me.cmbDefaultCurrency.Name = "cmbDefaultCurrency"
        '
        'lblDefaultIndType
        '
        resources.ApplyResources(Me.lblDefaultIndType, "lblDefaultIndType")
        Me.lblDefaultIndType.Name = "lblDefaultIndType"
        '
        'Panel1
        '
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.tbDefaultFont)
        Me.Panel1.Controls.Add(Me.lblDefaultFont)
        Me.Panel1.Controls.Add(Me.tbDefaultFontSortNumbers)
        Me.Panel1.Controls.Add(Me.chkWarnLinkedObjDelete)
        Me.Panel1.Controls.Add(Me.lblDefaultFontSortNumbers)
        Me.Panel1.Name = "Panel1"
        '
        'tbDefaultFont
        '
        resources.ApplyResources(Me.tbDefaultFont, "tbDefaultFont")
        Me.tbDefaultFont.Name = "tbDefaultFont"
        Me.tbDefaultFont.ReadOnly = True
        '
        'lblDefaultFont
        '
        resources.ApplyResources(Me.lblDefaultFont, "lblDefaultFont")
        Me.lblDefaultFont.Name = "lblDefaultFont"
        '
        'tbDefaultFontSortNumbers
        '
        resources.ApplyResources(Me.tbDefaultFontSortNumbers, "tbDefaultFontSortNumbers")
        Me.tbDefaultFontSortNumbers.Name = "tbDefaultFontSortNumbers"
        Me.tbDefaultFontSortNumbers.ReadOnly = True
        '
        'chkWarnLinkedObjDelete
        '
        resources.ApplyResources(Me.chkWarnLinkedObjDelete, "chkWarnLinkedObjDelete")
        Me.chkWarnLinkedObjDelete.Name = "chkWarnLinkedObjDelete"
        Me.chkWarnLinkedObjDelete.UseVisualStyleBackColor = True
        '
        'lblDefaultFontSortNumbers
        '
        resources.ApplyResources(Me.lblDefaultFontSortNumbers, "lblDefaultFontSortNumbers")
        Me.lblDefaultFontSortNumbers.Name = "lblDefaultFontSortNumbers"
        '
        'DialogSettings
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.TabControlSettings)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogSettings"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControlSettings.ResumeLayout(False)
        Me.TabPageDesign.ResumeLayout(False)
        Me.TabPageDesign.PerformLayout()
        CType(Me.dgvProjectLogic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageCurrentLogFrame.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.Panel7.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.TabPageGeneral.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents TabControlSettings As System.Windows.Forms.TabControl
    Friend WithEvents TabPageGeneral As System.Windows.Forms.TabPage
    Friend WithEvents TabPageDesign As System.Windows.Forms.TabPage
    Friend WithEvents dgvProjectLogic As FaciliDev.LogFramer.DataGridViewProjectLogic
    Friend WithEvents lblUse As System.Windows.Forms.Label
    Friend WithEvents FontDialogSettings As System.Windows.Forms.FontDialog
    Friend WithEvents ColorDialogSettings As System.Windows.Forms.ColorDialog
    Friend WithEvents chkDefaultScheme As System.Windows.Forms.CheckBox
    Friend WithEvents TabPageCurrentLogFrame As System.Windows.Forms.TabPage
    Friend WithEvents btnEqualColorLf As System.Windows.Forms.Button
    Friend WithEvents tbSectionTitleFontColor As System.Windows.Forms.TextBox
    Friend WithEvents lblSectionTitleFontColor As System.Windows.Forms.Label
    Friend WithEvents tbSectionColorBottom As System.Windows.Forms.TextBox
    Friend WithEvents tbSectionColorTop As System.Windows.Forms.TextBox
    Friend WithEvents lblSectionColor As System.Windows.Forms.Label
    Friend WithEvents tbSectionTitleFont As System.Windows.Forms.TextBox
    Friend WithEvents lblSectionTitleFont As System.Windows.Forms.Label
    Friend WithEvents tbSortNumberFont As System.Windows.Forms.TextBox
    Friend WithEvents lblCurrency As System.Windows.Forms.Label
    Friend WithEvents cmbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents lblSortNumberFont As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ntbRecentFilesMax As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents lblRecentFilesMax As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblDefaultCurrency As System.Windows.Forms.Label
    Friend WithEvents cmbDefaultCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents lblDefaultIndType As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnEqualColor As System.Windows.Forms.Button
    Friend WithEvents tbDefaultFontSectionColor As System.Windows.Forms.TextBox
    Friend WithEvents lblDefaultFontSectionColor As System.Windows.Forms.Label
    Friend WithEvents tbDefaultColorSectionBottom As System.Windows.Forms.TextBox
    Friend WithEvents tbDefaultColorSectionTop As System.Windows.Forms.TextBox
    Friend WithEvents lblDefaultColorSection As System.Windows.Forms.Label
    Friend WithEvents tbDefaultFontSections As System.Windows.Forms.TextBox
    Friend WithEvents lblDefaultFontSections As System.Windows.Forms.Label
    Friend WithEvents tbDefaultFontUI As System.Windows.Forms.TextBox
    Friend WithEvents tbDefaultFontSortNumbers As System.Windows.Forms.TextBox
    Friend WithEvents chkWarnLinkedObjDelete As System.Windows.Forms.CheckBox
    Friend WithEvents chkRepeatOutputs As System.Windows.Forms.CheckBox
    Friend WithEvents chkRepeatPurposes As System.Windows.Forms.CheckBox
    Friend WithEvents lblDefaultFontUI As System.Windows.Forms.Label
    Friend WithEvents lblDefaultFontSortNumbers As System.Windows.Forms.Label
    Friend WithEvents ntbTimerAutoSave As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents lblTimerAutoSave As System.Windows.Forms.Label
    Friend WithEvents lblAutoSaveMinutes As System.Windows.Forms.Label
    Friend WithEvents cmbDefaultIndType As FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbSortNumberDivider As System.Windows.Forms.TextBox
    Friend WithEvents chkSortNumberTerminateWithDivider As System.Windows.Forms.CheckBox
    Friend WithEvents chkSortNumberRepeatParent As System.Windows.Forms.CheckBox
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents lblSortNumberExample As System.Windows.Forms.Label
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents tbDetailsFont As System.Windows.Forms.TextBox
    Friend WithEvents lblDetailsFont As System.Windows.Forms.Label
    Friend WithEvents lblIndicatorName As System.Windows.Forms.Label
    Friend WithEvents tbIndicatorName As System.Windows.Forms.TextBox
    Friend WithEvents lblBudgetName As System.Windows.Forms.Label
    Friend WithEvents tbBudgetName As System.Windows.Forms.TextBox
    Friend WithEvents lblResourceName As System.Windows.Forms.Label
    Friend WithEvents tbResourceName As System.Windows.Forms.TextBox
    Friend WithEvents lblVerificationSourceName As System.Windows.Forms.Label
    Friend WithEvents tbVerificationSourceName As System.Windows.Forms.TextBox
    Friend WithEvents lblAssumptionName As System.Windows.Forms.Label
    Friend WithEvents tbAssumptionName As System.Windows.Forms.TextBox
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents tbDefaultFont As System.Windows.Forms.TextBox
    Friend WithEvents lblDefaultFont As System.Windows.Forms.Label

End Class
