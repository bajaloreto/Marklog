<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogImageSelect
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogImageSelect))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.gbSelect = New System.Windows.Forms.GroupBox()
        Me.rbtnAllImages = New System.Windows.Forms.RadioButton()
        Me.rbtnSingleFile = New System.Windows.Forms.RadioButton()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.gbSelect.SuspendLayout()
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
        'gbSelect
        '
        resources.ApplyResources(Me.gbSelect, "gbSelect")
        Me.gbSelect.Controls.Add(Me.rbtnAllImages)
        Me.gbSelect.Controls.Add(Me.rbtnSingleFile)
        Me.gbSelect.Name = "gbSelect"
        Me.gbSelect.TabStop = False
        '
        'rbtnAllImages
        '
        resources.ApplyResources(Me.rbtnAllImages, "rbtnAllImages")
        Me.rbtnAllImages.Name = "rbtnAllImages"
        Me.rbtnAllImages.UseVisualStyleBackColor = True
        '
        'rbtnSingleFile
        '
        resources.ApplyResources(Me.rbtnSingleFile, "rbtnSingleFile")
        Me.rbtnSingleFile.Checked = True
        Me.rbtnSingleFile.Name = "rbtnSingleFile"
        Me.rbtnSingleFile.TabStop = True
        Me.rbtnSingleFile.UseVisualStyleBackColor = True
        '
        'DialogImageSelect
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.gbSelect)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogImageSelect"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.gbSelect.ResumeLayout(False)
        Me.gbSelect.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents gbSelect As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnAllImages As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnSingleFile As System.Windows.Forms.RadioButton

End Class
