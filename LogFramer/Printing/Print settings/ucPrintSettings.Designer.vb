<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PrintSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PrintSettings))
        Me.btnPageMargins = New System.Windows.Forms.Button()
        Me.btnHeaderText = New System.Windows.Forms.Button()
        Me.btnFooterText = New System.Windows.Forms.Button()
        Me.gbOrientation = New System.Windows.Forms.GroupBox()
        Me.rbtnLandscape = New System.Windows.Forms.RadioButton()
        Me.rbtnPortrait = New System.Windows.Forms.RadioButton()
        Me.gbMargins = New System.Windows.Forms.GroupBox()
        Me.rbtnWide = New System.Windows.Forms.RadioButton()
        Me.rbtnNormal = New System.Windows.Forms.RadioButton()
        Me.rbtnThin = New System.Windows.Forms.RadioButton()
        Me.ProgressBarDocument = New System.Windows.Forms.ProgressBar()
        Me.btnSelectPrinter = New System.Windows.Forms.Button()
        Me.PrintDialog = New System.Windows.Forms.PrintDialog()
        Me.btnPrint = New System.Windows.Forms.Button()
        Me.gbOrientation.SuspendLayout()
        Me.gbMargins.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPageMargins
        '
        resources.ApplyResources(Me.btnPageMargins, "btnPageMargins")
        Me.btnPageMargins.Name = "btnPageMargins"
        Me.btnPageMargins.UseVisualStyleBackColor = True
        '
        'btnHeaderText
        '
        resources.ApplyResources(Me.btnHeaderText, "btnHeaderText")
        Me.btnHeaderText.Name = "btnHeaderText"
        Me.btnHeaderText.UseVisualStyleBackColor = True
        '
        'btnFooterText
        '
        resources.ApplyResources(Me.btnFooterText, "btnFooterText")
        Me.btnFooterText.Name = "btnFooterText"
        Me.btnFooterText.UseVisualStyleBackColor = True
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
        'gbMargins
        '
        resources.ApplyResources(Me.gbMargins, "gbMargins")
        Me.gbMargins.Controls.Add(Me.rbtnWide)
        Me.gbMargins.Controls.Add(Me.rbtnNormal)
        Me.gbMargins.Controls.Add(Me.rbtnThin)
        Me.gbMargins.Controls.Add(Me.btnPageMargins)
        Me.gbMargins.Name = "gbMargins"
        Me.gbMargins.TabStop = False
        '
        'rbtnWide
        '
        resources.ApplyResources(Me.rbtnWide, "rbtnWide")
        Me.rbtnWide.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_margins_wide
        Me.rbtnWide.Name = "rbtnWide"
        Me.rbtnWide.TabStop = True
        Me.rbtnWide.UseVisualStyleBackColor = True
        '
        'rbtnNormal
        '
        resources.ApplyResources(Me.rbtnNormal, "rbtnNormal")
        Me.rbtnNormal.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_margins_medium
        Me.rbtnNormal.Name = "rbtnNormal"
        Me.rbtnNormal.TabStop = True
        Me.rbtnNormal.UseVisualStyleBackColor = True
        '
        'rbtnThin
        '
        resources.ApplyResources(Me.rbtnThin, "rbtnThin")
        Me.rbtnThin.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Print_margins_small
        Me.rbtnThin.Name = "rbtnThin"
        Me.rbtnThin.TabStop = True
        Me.rbtnThin.UseVisualStyleBackColor = True
        '
        'ProgressBarDocument
        '
        resources.ApplyResources(Me.ProgressBarDocument, "ProgressBarDocument")
        Me.ProgressBarDocument.Name = "ProgressBarDocument"
        '
        'btnSelectPrinter
        '
        resources.ApplyResources(Me.btnSelectPrinter, "btnSelectPrinter")
        Me.btnSelectPrinter.Name = "btnSelectPrinter"
        Me.btnSelectPrinter.UseVisualStyleBackColor = True
        '
        'PrintDialog
        '
        Me.PrintDialog.UseEXDialog = True
        '
        'btnPrint
        '
        resources.ApplyResources(Me.btnPrint, "btnPrint")
        Me.btnPrint.Image = Global.FaciliDev.LogFramer.My.Resources.Resources.Printer_middelgroot1
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'PrintSettings
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.btnSelectPrinter)
        Me.Controls.Add(Me.ProgressBarDocument)
        Me.Controls.Add(Me.gbMargins)
        Me.Controls.Add(Me.gbOrientation)
        Me.Controls.Add(Me.btnFooterText)
        Me.Controls.Add(Me.btnHeaderText)
        Me.Controls.Add(Me.btnPrint)
        Me.Name = "PrintSettings"
        Me.gbOrientation.ResumeLayout(False)
        Me.gbOrientation.PerformLayout()
        Me.gbMargins.ResumeLayout(False)
        Me.gbMargins.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnPageMargins As System.Windows.Forms.Button
    Friend WithEvents btnHeaderText As System.Windows.Forms.Button
    Friend WithEvents btnFooterText As System.Windows.Forms.Button
    Friend WithEvents gbOrientation As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnLandscape As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnPortrait As System.Windows.Forms.RadioButton
    Friend WithEvents gbMargins As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnWide As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNormal As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnThin As System.Windows.Forms.RadioButton
    Friend WithEvents ProgressBarDocument As System.Windows.Forms.ProgressBar
    Friend WithEvents btnSelectPrinter As System.Windows.Forms.Button
    Friend WithEvents PrintDialog As System.Windows.Forms.PrintDialog

End Class
