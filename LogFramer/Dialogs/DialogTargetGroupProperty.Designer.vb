<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogTargetGroupInformation
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogTargetGroupInformation))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.lblPropertyName = New System.Windows.Forms.Label()
        Me.tbName = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblPropertyType = New System.Windows.Forms.Label()
        Me.cmbPropertyType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblOption2 = New System.Windows.Forms.Label()
        Me.lblOption1 = New System.Windows.Forms.Label()
        Me.PanelOptions = New System.Windows.Forms.Panel()
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
        'lblPropertyName
        '
        resources.ApplyResources(Me.lblPropertyName, "lblPropertyName")
        Me.lblPropertyName.Name = "lblPropertyName"
        '
        'tbName
        '
        resources.ApplyResources(Me.tbName, "tbName")
        Me.tbName.Name = "tbName"
        '
        'lblPropertyType
        '
        resources.ApplyResources(Me.lblPropertyType, "lblPropertyType")
        Me.lblPropertyType.Name = "lblPropertyType"
        '
        'cmbPropertyType
        '
        resources.ApplyResources(Me.cmbPropertyType, "cmbPropertyType")
        Me.cmbPropertyType.FormattingEnabled = True
        Me.cmbPropertyType.Name = "cmbPropertyType"
        '
        'lblOption2
        '
        resources.ApplyResources(Me.lblOption2, "lblOption2")
        Me.lblOption2.Name = "lblOption2"
        '
        'lblOption1
        '
        resources.ApplyResources(Me.lblOption1, "lblOption1")
        Me.lblOption1.Name = "lblOption1"
        '
        'PanelOptions
        '
        resources.ApplyResources(Me.PanelOptions, "PanelOptions")
        Me.PanelOptions.Name = "PanelOptions"
        '
        'DialogTargetGroupInformation
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.PanelOptions)
        Me.Controls.Add(Me.lblOption1)
        Me.Controls.Add(Me.lblOption2)
        Me.Controls.Add(Me.cmbPropertyType)
        Me.Controls.Add(Me.lblPropertyType)
        Me.Controls.Add(Me.tbName)
        Me.Controls.Add(Me.lblPropertyName)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogTargetGroupInformation"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents lblPropertyName As System.Windows.Forms.Label
    Friend WithEvents tbName As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblPropertyType As System.Windows.Forms.Label
    Friend WithEvents cmbPropertyType As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblOption2 As System.Windows.Forms.Label
    Friend WithEvents lblOption1 As System.Windows.Forms.Label
    Friend WithEvents PanelOptions As System.Windows.Forms.Panel

End Class
