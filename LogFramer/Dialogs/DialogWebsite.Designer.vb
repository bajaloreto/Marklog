<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogWebsite
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogWebsite))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.tbWebsitePath = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblWebsitePath = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.cmbType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.tbWebsiteName = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblWebsiteName = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
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
        'tbWebsitePath
        '
        resources.ApplyResources(Me.tbWebsitePath, "tbWebsitePath")
        Me.tbWebsitePath.ForeColor = System.Drawing.SystemColors.Highlight
        Me.tbWebsitePath.Name = "tbWebsitePath"
        '
        'lblWebsitePath
        '
        resources.ApplyResources(Me.lblWebsitePath, "lblWebsitePath")
        Me.lblWebsitePath.Name = "lblWebsitePath"
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
        'tbWebsiteName
        '
        resources.ApplyResources(Me.tbWebsiteName, "tbWebsiteName")
        Me.tbWebsiteName.Name = "tbWebsiteName"
        '
        'lblWebsiteName
        '
        resources.ApplyResources(Me.lblWebsiteName, "lblWebsiteName")
        Me.lblWebsiteName.Name = "lblWebsiteName"
        '
        'DialogWebsite
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.tbWebsiteName)
        Me.Controls.Add(Me.lblWebsiteName)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.tbWebsitePath)
        Me.Controls.Add(Me.lblWebsitePath)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogWebsite"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents tbWebsitePath As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblWebsitePath As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cmbType As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents tbWebsiteName As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblWebsiteName As System.Windows.Forms.Label

End Class
