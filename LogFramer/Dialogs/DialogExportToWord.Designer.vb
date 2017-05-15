<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogExportToWord
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogExportToWord))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.TabControlOptions = New System.Windows.Forms.TabControl()
        Me.TabPageWordOptions = New System.Windows.Forms.TabPage()
        Me.GroupBoxMainOptions = New System.Windows.Forms.GroupBox()
        Me.PanelExportOptions = New System.Windows.Forms.Panel()
        Me.gbInsertOptions = New System.Windows.Forms.GroupBox()
        Me.cmbBookmark = New System.Windows.Forms.ComboBox()
        Me.rbtnBookmark = New System.Windows.Forms.RadioButton()
        Me.rbtnEndDoc = New System.Windows.Forms.RadioButton()
        Me.rbtnStartDoc = New System.Windows.Forms.RadioButton()
        Me.btnLoadFile = New System.Windows.Forms.Button()
        Me.lblFileName = New System.Windows.Forms.Label()
        Me.tbFileName = New System.Windows.Forms.TextBox()
        Me.rbtnWordDoc = New System.Windows.Forms.RadioButton()
        Me.rbtnNewWordDoc = New System.Windows.Forms.RadioButton()
        Me.TabPageOrientation = New System.Windows.Forms.TabPage()
        Me.gbOrientation = New System.Windows.Forms.GroupBox()
        Me.rbtnLandscape = New System.Windows.Forms.RadioButton()
        Me.rbtnPortrait = New System.Windows.Forms.RadioButton()
        Me.TabPageIndicatorOptions = New System.Windows.Forms.TabPage()
        Me.lblOptions = New System.Windows.Forms.Label()
        Me.chkPrintTarget = New System.Windows.Forms.CheckBox()
        Me.chkPrintRange = New System.Windows.Forms.CheckBox()
        Me.chkPrintResponse = New System.Windows.Forms.CheckBox()
        Me.chkPrintOutput = New System.Windows.Forms.CheckBox()
        Me.chkPrintPurpose = New System.Windows.Forms.CheckBox()
        Me.lblPrintLevel = New System.Windows.Forms.Label()
        Me.cmbPrintLevel = New System.Windows.Forms.ComboBox()
        Me.TabPagePlanningOptions = New System.Windows.Forms.TabPage()
        Me.lblShowElements = New System.Windows.Forms.Label()
        Me.cmbShowElements = New System.Windows.Forms.ComboBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControlOptions.SuspendLayout()
        Me.TabPageWordOptions.SuspendLayout()
        Me.GroupBoxMainOptions.SuspendLayout()
        Me.PanelExportOptions.SuspendLayout()
        Me.gbInsertOptions.SuspendLayout()
        Me.TabPageOrientation.SuspendLayout()
        Me.gbOrientation.SuspendLayout()
        Me.TabPageIndicatorOptions.SuspendLayout()
        Me.TabPagePlanningOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.btnExport, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'btnExport
        '
        resources.ApplyResources(Me.btnExport, "btnExport")
        Me.btnExport.Name = "btnExport"
        '
        'Cancel_Button
        '
        resources.ApplyResources(Me.Cancel_Button, "Cancel_Button")
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Name = "Cancel_Button"
        '
        'TabControlOptions
        '
        resources.ApplyResources(Me.TabControlOptions, "TabControlOptions")
        Me.TabControlOptions.Controls.Add(Me.TabPageWordOptions)
        Me.TabControlOptions.Controls.Add(Me.TabPageOrientation)
        Me.TabControlOptions.Controls.Add(Me.TabPageIndicatorOptions)
        Me.TabControlOptions.Controls.Add(Me.TabPagePlanningOptions)
        Me.TabControlOptions.Name = "TabControlOptions"
        Me.TabControlOptions.SelectedIndex = 0
        '
        'TabPageWordOptions
        '
        resources.ApplyResources(Me.TabPageWordOptions, "TabPageWordOptions")
        Me.TabPageWordOptions.Controls.Add(Me.GroupBoxMainOptions)
        Me.TabPageWordOptions.Name = "TabPageWordOptions"
        Me.TabPageWordOptions.UseVisualStyleBackColor = True
        '
        'GroupBoxMainOptions
        '
        resources.ApplyResources(Me.GroupBoxMainOptions, "GroupBoxMainOptions")
        Me.GroupBoxMainOptions.Controls.Add(Me.PanelExportOptions)
        Me.GroupBoxMainOptions.Controls.Add(Me.rbtnWordDoc)
        Me.GroupBoxMainOptions.Controls.Add(Me.rbtnNewWordDoc)
        Me.GroupBoxMainOptions.Name = "GroupBoxMainOptions"
        Me.GroupBoxMainOptions.TabStop = False
        '
        'PanelExportOptions
        '
        resources.ApplyResources(Me.PanelExportOptions, "PanelExportOptions")
        Me.PanelExportOptions.Controls.Add(Me.gbInsertOptions)
        Me.PanelExportOptions.Controls.Add(Me.btnLoadFile)
        Me.PanelExportOptions.Controls.Add(Me.lblFileName)
        Me.PanelExportOptions.Controls.Add(Me.tbFileName)
        Me.PanelExportOptions.Name = "PanelExportOptions"
        '
        'gbInsertOptions
        '
        resources.ApplyResources(Me.gbInsertOptions, "gbInsertOptions")
        Me.gbInsertOptions.Controls.Add(Me.cmbBookmark)
        Me.gbInsertOptions.Controls.Add(Me.rbtnBookmark)
        Me.gbInsertOptions.Controls.Add(Me.rbtnEndDoc)
        Me.gbInsertOptions.Controls.Add(Me.rbtnStartDoc)
        Me.gbInsertOptions.Name = "gbInsertOptions"
        Me.gbInsertOptions.TabStop = False
        '
        'cmbBookmark
        '
        resources.ApplyResources(Me.cmbBookmark, "cmbBookmark")
        Me.cmbBookmark.FormattingEnabled = True
        Me.cmbBookmark.Name = "cmbBookmark"
        '
        'rbtnBookmark
        '
        resources.ApplyResources(Me.rbtnBookmark, "rbtnBookmark")
        Me.rbtnBookmark.Name = "rbtnBookmark"
        Me.rbtnBookmark.TabStop = True
        Me.rbtnBookmark.UseVisualStyleBackColor = True
        '
        'rbtnEndDoc
        '
        resources.ApplyResources(Me.rbtnEndDoc, "rbtnEndDoc")
        Me.rbtnEndDoc.Name = "rbtnEndDoc"
        Me.rbtnEndDoc.TabStop = True
        Me.rbtnEndDoc.UseVisualStyleBackColor = True
        '
        'rbtnStartDoc
        '
        resources.ApplyResources(Me.rbtnStartDoc, "rbtnStartDoc")
        Me.rbtnStartDoc.Name = "rbtnStartDoc"
        Me.rbtnStartDoc.TabStop = True
        Me.rbtnStartDoc.UseVisualStyleBackColor = True
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
        'rbtnWordDoc
        '
        resources.ApplyResources(Me.rbtnWordDoc, "rbtnWordDoc")
        Me.rbtnWordDoc.Name = "rbtnWordDoc"
        Me.rbtnWordDoc.TabStop = True
        Me.rbtnWordDoc.UseVisualStyleBackColor = True
        '
        'rbtnNewWordDoc
        '
        resources.ApplyResources(Me.rbtnNewWordDoc, "rbtnNewWordDoc")
        Me.rbtnNewWordDoc.Name = "rbtnNewWordDoc"
        Me.rbtnNewWordDoc.TabStop = True
        Me.rbtnNewWordDoc.UseVisualStyleBackColor = True
        '
        'TabPageOrientation
        '
        resources.ApplyResources(Me.TabPageOrientation, "TabPageOrientation")
        Me.TabPageOrientation.Controls.Add(Me.gbOrientation)
        Me.TabPageOrientation.Name = "TabPageOrientation"
        Me.TabPageOrientation.UseVisualStyleBackColor = True
        '
        'gbOrientation
        '
        resources.ApplyResources(Me.gbOrientation, "gbOrientation")
        Me.gbOrientation.Controls.Add(Me.rbtnLandscape)
        Me.gbOrientation.Controls.Add(Me.rbtnPortrait)
        Me.gbOrientation.Name = "gbOrientation"
        Me.gbOrientation.TabStop = False
        '
        'rbtnLandscape
        '
        resources.ApplyResources(Me.rbtnLandscape, "rbtnLandscape")
        Me.rbtnLandscape.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_landscape1
        Me.rbtnLandscape.Name = "rbtnLandscape"
        Me.rbtnLandscape.TabStop = True
        Me.rbtnLandscape.UseVisualStyleBackColor = True
        '
        'rbtnPortrait
        '
        resources.ApplyResources(Me.rbtnPortrait, "rbtnPortrait")
        Me.rbtnPortrait.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_portrait2
        Me.rbtnPortrait.Name = "rbtnPortrait"
        Me.rbtnPortrait.TabStop = True
        Me.rbtnPortrait.UseVisualStyleBackColor = True
        '
        'TabPageIndicatorOptions
        '
        resources.ApplyResources(Me.TabPageIndicatorOptions, "TabPageIndicatorOptions")
        Me.TabPageIndicatorOptions.Controls.Add(Me.lblOptions)
        Me.TabPageIndicatorOptions.Controls.Add(Me.chkPrintTarget)
        Me.TabPageIndicatorOptions.Controls.Add(Me.chkPrintRange)
        Me.TabPageIndicatorOptions.Controls.Add(Me.chkPrintResponse)
        Me.TabPageIndicatorOptions.Controls.Add(Me.chkPrintOutput)
        Me.TabPageIndicatorOptions.Controls.Add(Me.chkPrintPurpose)
        Me.TabPageIndicatorOptions.Controls.Add(Me.lblPrintLevel)
        Me.TabPageIndicatorOptions.Controls.Add(Me.cmbPrintLevel)
        Me.TabPageIndicatorOptions.Name = "TabPageIndicatorOptions"
        Me.TabPageIndicatorOptions.UseVisualStyleBackColor = True
        '
        'lblOptions
        '
        resources.ApplyResources(Me.lblOptions, "lblOptions")
        Me.lblOptions.Name = "lblOptions"
        '
        'chkPrintTarget
        '
        resources.ApplyResources(Me.chkPrintTarget, "chkPrintTarget")
        Me.chkPrintTarget.Name = "chkPrintTarget"
        Me.chkPrintTarget.UseVisualStyleBackColor = True
        '
        'chkPrintRange
        '
        resources.ApplyResources(Me.chkPrintRange, "chkPrintRange")
        Me.chkPrintRange.Name = "chkPrintRange"
        Me.chkPrintRange.UseVisualStyleBackColor = True
        '
        'chkPrintResponse
        '
        resources.ApplyResources(Me.chkPrintResponse, "chkPrintResponse")
        Me.chkPrintResponse.Name = "chkPrintResponse"
        Me.chkPrintResponse.UseVisualStyleBackColor = True
        '
        'chkPrintOutput
        '
        resources.ApplyResources(Me.chkPrintOutput, "chkPrintOutput")
        Me.chkPrintOutput.Name = "chkPrintOutput"
        Me.chkPrintOutput.UseVisualStyleBackColor = True
        '
        'chkPrintPurpose
        '
        resources.ApplyResources(Me.chkPrintPurpose, "chkPrintPurpose")
        Me.chkPrintPurpose.Name = "chkPrintPurpose"
        Me.chkPrintPurpose.UseVisualStyleBackColor = True
        '
        'lblPrintLevel
        '
        resources.ApplyResources(Me.lblPrintLevel, "lblPrintLevel")
        Me.lblPrintLevel.Name = "lblPrintLevel"
        '
        'cmbPrintLevel
        '
        resources.ApplyResources(Me.cmbPrintLevel, "cmbPrintLevel")
        Me.cmbPrintLevel.FormattingEnabled = True
        Me.cmbPrintLevel.Name = "cmbPrintLevel"
        '
        'TabPagePlanningOptions
        '
        resources.ApplyResources(Me.TabPagePlanningOptions, "TabPagePlanningOptions")
        Me.TabPagePlanningOptions.Controls.Add(Me.lblShowElements)
        Me.TabPagePlanningOptions.Controls.Add(Me.cmbShowElements)
        Me.TabPagePlanningOptions.Name = "TabPagePlanningOptions"
        Me.TabPagePlanningOptions.UseVisualStyleBackColor = True
        '
        'lblShowElements
        '
        resources.ApplyResources(Me.lblShowElements, "lblShowElements")
        Me.lblShowElements.Name = "lblShowElements"
        '
        'cmbShowElements
        '
        resources.ApplyResources(Me.cmbShowElements, "cmbShowElements")
        Me.cmbShowElements.FormattingEnabled = True
        Me.cmbShowElements.Name = "cmbShowElements"
        '
        'DialogExportToWord
        '
        Me.AcceptButton = Me.btnExport
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.TabControlOptions)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogExportToWord"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControlOptions.ResumeLayout(False)
        Me.TabPageWordOptions.ResumeLayout(False)
        Me.GroupBoxMainOptions.ResumeLayout(False)
        Me.GroupBoxMainOptions.PerformLayout()
        Me.PanelExportOptions.ResumeLayout(False)
        Me.PanelExportOptions.PerformLayout()
        Me.gbInsertOptions.ResumeLayout(False)
        Me.gbInsertOptions.PerformLayout()
        Me.TabPageOrientation.ResumeLayout(False)
        Me.gbOrientation.ResumeLayout(False)
        Me.gbOrientation.PerformLayout()
        Me.TabPageIndicatorOptions.ResumeLayout(False)
        Me.TabPageIndicatorOptions.PerformLayout()
        Me.TabPagePlanningOptions.ResumeLayout(False)
        Me.TabPagePlanningOptions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents TabControlOptions As System.Windows.Forms.TabControl
    Friend WithEvents TabPageWordOptions As System.Windows.Forms.TabPage
    Friend WithEvents GroupBoxMainOptions As System.Windows.Forms.GroupBox
    Friend WithEvents PanelExportOptions As System.Windows.Forms.Panel
    Friend WithEvents gbInsertOptions As System.Windows.Forms.GroupBox
    Friend WithEvents cmbBookmark As System.Windows.Forms.ComboBox
    Friend WithEvents rbtnBookmark As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnEndDoc As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnStartDoc As System.Windows.Forms.RadioButton
    Friend WithEvents btnLoadFile As System.Windows.Forms.Button
    Friend WithEvents lblFileName As System.Windows.Forms.Label
    Friend WithEvents tbFileName As System.Windows.Forms.TextBox
    Friend WithEvents rbtnWordDoc As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNewWordDoc As System.Windows.Forms.RadioButton
    Friend WithEvents TabPageIndicatorOptions As System.Windows.Forms.TabPage
    Friend WithEvents lblOptions As System.Windows.Forms.Label
    Friend WithEvents chkPrintTarget As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintRange As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintResponse As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintOutput As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrintPurpose As System.Windows.Forms.CheckBox
    Friend WithEvents lblPrintLevel As System.Windows.Forms.Label
    Friend WithEvents cmbPrintLevel As System.Windows.Forms.ComboBox
    Friend WithEvents TabPageOrientation As System.Windows.Forms.TabPage
    Friend WithEvents gbOrientation As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnLandscape As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPortrait As System.Windows.Forms.RadioButton
    Friend WithEvents TabPagePlanningOptions As System.Windows.Forms.TabPage
    Friend WithEvents lblShowElements As System.Windows.Forms.Label
    Friend WithEvents cmbShowElements As System.Windows.Forms.ComboBox

End Class
