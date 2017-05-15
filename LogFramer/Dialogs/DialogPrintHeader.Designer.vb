<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogPrintHeader
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogPrintHeader))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.rtbLeft = New System.Windows.Forms.RichTextBox()
        Me.tspFormatting = New System.Windows.Forms.ToolStrip()
        Me.cmbFonts = New System.Windows.Forms.ToolStripComboBox()
        Me.cmbFontSize = New System.Windows.Forms.ToolStripComboBox()
        Me.buttonTextBold = New System.Windows.Forms.ToolStripButton()
        Me.buttonTextItalics = New System.Windows.Forms.ToolStripButton()
        Me.buttonTextUnderlined = New System.Windows.Forms.ToolStripButton()
        Me.buttonTextStrikeThrough = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.buttonTextSuperscript = New System.Windows.Forms.ToolStripButton()
        Me.buttonTextSubscript = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.buttonTextColor = New System.Windows.Forms.ToolStripSplitButton()
        Me.buttonTextBackGround = New System.Windows.Forms.ToolStripSplitButton()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.buttonParAlignLeft = New System.Windows.Forms.ToolStripButton()
        Me.buttonParAlignCentre = New System.Windows.Forms.ToolStripButton()
        Me.buttonParAlignRight = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.buttonAddPageNumber = New System.Windows.Forms.ToolStripButton()
        Me.buttonAddTotalPages = New System.Windows.Forms.ToolStripButton()
        Me.rtbMiddle = New System.Windows.Forms.RichTextBox()
        Me.rtbRight = New System.Windows.Forms.RichTextBox()
        Me.chkCopyToAllReports = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.tspFormatting.SuspendLayout()
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
        'rtbLeft
        '
        Me.rtbLeft.AcceptsTab = True
        resources.ApplyResources(Me.rtbLeft, "rtbLeft")
        Me.rtbLeft.Name = "rtbLeft"
        '
        'tspFormatting
        '
        resources.ApplyResources(Me.tspFormatting, "tspFormatting")
        Me.tspFormatting.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmbFonts, Me.cmbFontSize, Me.buttonTextBold, Me.buttonTextItalics, Me.buttonTextUnderlined, Me.buttonTextStrikeThrough, Me.ToolStripSeparator6, Me.buttonTextSuperscript, Me.buttonTextSubscript, Me.ToolStripSeparator7, Me.buttonTextColor, Me.buttonTextBackGround, Me.ToolStripSeparator8, Me.buttonParAlignLeft, Me.buttonParAlignCentre, Me.buttonParAlignRight, Me.ToolStripSeparator1, Me.buttonAddPageNumber, Me.buttonAddTotalPages})
        Me.tspFormatting.Name = "tspFormatting"
        '
        'cmbFonts
        '
        resources.ApplyResources(Me.cmbFonts, "cmbFonts")
        Me.cmbFonts.Name = "cmbFonts"
        '
        'cmbFontSize
        '
        resources.ApplyResources(Me.cmbFontSize, "cmbFontSize")
        Me.cmbFontSize.Name = "cmbFontSize"
        '
        'buttonTextBold
        '
        resources.ApplyResources(Me.buttonTextBold, "buttonTextBold")
        Me.buttonTextBold.CheckOnClick = True
        Me.buttonTextBold.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextBold.Name = "buttonTextBold"
        '
        'buttonTextItalics
        '
        resources.ApplyResources(Me.buttonTextItalics, "buttonTextItalics")
        Me.buttonTextItalics.CheckOnClick = True
        Me.buttonTextItalics.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextItalics.Name = "buttonTextItalics"
        '
        'buttonTextUnderlined
        '
        resources.ApplyResources(Me.buttonTextUnderlined, "buttonTextUnderlined")
        Me.buttonTextUnderlined.CheckOnClick = True
        Me.buttonTextUnderlined.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextUnderlined.Name = "buttonTextUnderlined"
        '
        'buttonTextStrikeThrough
        '
        resources.ApplyResources(Me.buttonTextStrikeThrough, "buttonTextStrikeThrough")
        Me.buttonTextStrikeThrough.CheckOnClick = True
        Me.buttonTextStrikeThrough.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextStrikeThrough.Name = "buttonTextStrikeThrough"
        '
        'ToolStripSeparator6
        '
        resources.ApplyResources(Me.ToolStripSeparator6, "ToolStripSeparator6")
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        '
        'buttonTextSuperscript
        '
        resources.ApplyResources(Me.buttonTextSuperscript, "buttonTextSuperscript")
        Me.buttonTextSuperscript.CheckOnClick = True
        Me.buttonTextSuperscript.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextSuperscript.Name = "buttonTextSuperscript"
        '
        'buttonTextSubscript
        '
        resources.ApplyResources(Me.buttonTextSubscript, "buttonTextSubscript")
        Me.buttonTextSubscript.CheckOnClick = True
        Me.buttonTextSubscript.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextSubscript.Name = "buttonTextSubscript"
        '
        'ToolStripSeparator7
        '
        resources.ApplyResources(Me.ToolStripSeparator7, "ToolStripSeparator7")
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        '
        'buttonTextColor
        '
        resources.ApplyResources(Me.buttonTextColor, "buttonTextColor")
        Me.buttonTextColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextColor.Name = "buttonTextColor"
        '
        'buttonTextBackGround
        '
        resources.ApplyResources(Me.buttonTextBackGround, "buttonTextBackGround")
        Me.buttonTextBackGround.BackColor = System.Drawing.Color.Yellow
        Me.buttonTextBackGround.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonTextBackGround.Name = "buttonTextBackGround"
        '
        'ToolStripSeparator8
        '
        resources.ApplyResources(Me.ToolStripSeparator8, "ToolStripSeparator8")
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        '
        'buttonParAlignLeft
        '
        resources.ApplyResources(Me.buttonParAlignLeft, "buttonParAlignLeft")
        Me.buttonParAlignLeft.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonParAlignLeft.Name = "buttonParAlignLeft"
        '
        'buttonParAlignCentre
        '
        resources.ApplyResources(Me.buttonParAlignCentre, "buttonParAlignCentre")
        Me.buttonParAlignCentre.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonParAlignCentre.Name = "buttonParAlignCentre"
        '
        'buttonParAlignRight
        '
        resources.ApplyResources(Me.buttonParAlignRight, "buttonParAlignRight")
        Me.buttonParAlignRight.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.buttonParAlignRight.Name = "buttonParAlignRight"
        '
        'ToolStripSeparator1
        '
        resources.ApplyResources(Me.ToolStripSeparator1, "ToolStripSeparator1")
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        '
        'buttonAddPageNumber
        '
        resources.ApplyResources(Me.buttonAddPageNumber, "buttonAddPageNumber")
        Me.buttonAddPageNumber.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.buttonAddPageNumber.Name = "buttonAddPageNumber"
        '
        'buttonAddTotalPages
        '
        resources.ApplyResources(Me.buttonAddTotalPages, "buttonAddTotalPages")
        Me.buttonAddTotalPages.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.buttonAddTotalPages.Name = "buttonAddTotalPages"
        '
        'rtbMiddle
        '
        Me.rtbMiddle.AcceptsTab = True
        resources.ApplyResources(Me.rtbMiddle, "rtbMiddle")
        Me.rtbMiddle.Name = "rtbMiddle"
        '
        'rtbRight
        '
        Me.rtbRight.AcceptsTab = True
        resources.ApplyResources(Me.rtbRight, "rtbRight")
        Me.rtbRight.Name = "rtbRight"
        '
        'chkCopyToAllReports
        '
        resources.ApplyResources(Me.chkCopyToAllReports, "chkCopyToAllReports")
        Me.chkCopyToAllReports.Name = "chkCopyToAllReports"
        Me.chkCopyToAllReports.UseVisualStyleBackColor = True
        '
        'DialogPrintHeader
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.chkCopyToAllReports)
        Me.Controls.Add(Me.rtbRight)
        Me.Controls.Add(Me.rtbMiddle)
        Me.Controls.Add(Me.tspFormatting)
        Me.Controls.Add(Me.rtbLeft)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogPrintHeader"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.tspFormatting.ResumeLayout(False)
        Me.tspFormatting.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents rtbLeft As System.Windows.Forms.RichTextBox
    Friend WithEvents tspFormatting As System.Windows.Forms.ToolStrip
    Friend WithEvents cmbFonts As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents cmbFontSize As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents buttonTextBold As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonTextItalics As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonTextUnderlined As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonTextStrikeThrough As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents buttonTextSuperscript As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonTextSubscript As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents buttonTextColor As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents buttonTextBackGround As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents buttonParAlignLeft As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonParAlignCentre As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonParAlignRight As System.Windows.Forms.ToolStripButton
    Friend WithEvents rtbMiddle As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbRight As System.Windows.Forms.RichTextBox
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents buttonAddPageNumber As System.Windows.Forms.ToolStripButton
    Friend WithEvents buttonAddTotalPages As System.Windows.Forms.ToolStripButton
    Friend WithEvents chkCopyToAllReports As System.Windows.Forms.CheckBox

End Class
