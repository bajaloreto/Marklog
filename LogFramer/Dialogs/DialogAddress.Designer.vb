<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogAddress
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogAddress))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.lblStreet = New System.Windows.Forms.Label()
        Me.tbStreet = New FaciliDev.LogFramer.TextBoxLF()
        Me.tbNumber = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.tbPostNumber = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblPostNumber = New System.Windows.Forms.Label()
        Me.tbDistrict = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblDistrict = New System.Windows.Forms.Label()
        Me.tbTown = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblTown = New System.Windows.Forms.Label()
        Me.tbCountry = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblCountry = New System.Windows.Forms.Label()
        Me.lblType = New System.Windows.Forms.Label()
        Me.cmbType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
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
        'lblStreet
        '
        resources.ApplyResources(Me.lblStreet, "lblStreet")
        Me.lblStreet.Name = "lblStreet"
        '
        'tbStreet
        '
        resources.ApplyResources(Me.tbStreet, "tbStreet")
        Me.tbStreet.Name = "tbStreet"
        '
        'tbNumber
        '
        resources.ApplyResources(Me.tbNumber, "tbNumber")
        Me.tbNumber.Name = "tbNumber"
        '
        'lblNumber
        '
        resources.ApplyResources(Me.lblNumber, "lblNumber")
        Me.lblNumber.Name = "lblNumber"
        '
        'tbPostNumber
        '
        resources.ApplyResources(Me.tbPostNumber, "tbPostNumber")
        Me.tbPostNumber.Name = "tbPostNumber"
        '
        'lblPostNumber
        '
        resources.ApplyResources(Me.lblPostNumber, "lblPostNumber")
        Me.lblPostNumber.Name = "lblPostNumber"
        '
        'tbDistrict
        '
        resources.ApplyResources(Me.tbDistrict, "tbDistrict")
        Me.tbDistrict.Name = "tbDistrict"
        '
        'lblDistrict
        '
        resources.ApplyResources(Me.lblDistrict, "lblDistrict")
        Me.lblDistrict.Name = "lblDistrict"
        '
        'tbTown
        '
        resources.ApplyResources(Me.tbTown, "tbTown")
        Me.tbTown.Name = "tbTown"
        '
        'lblTown
        '
        resources.ApplyResources(Me.lblTown, "lblTown")
        Me.lblTown.Name = "lblTown"
        '
        'tbCountry
        '
        resources.ApplyResources(Me.tbCountry, "tbCountry")
        Me.tbCountry.Name = "tbCountry"
        '
        'lblCountry
        '
        resources.ApplyResources(Me.lblCountry, "lblCountry")
        Me.lblCountry.Name = "lblCountry"
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
        'DialogAddress
        '
        Me.AcceptButton = Me.OK_Button
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.tbCountry)
        Me.Controls.Add(Me.lblCountry)
        Me.Controls.Add(Me.tbTown)
        Me.Controls.Add(Me.lblTown)
        Me.Controls.Add(Me.tbDistrict)
        Me.Controls.Add(Me.lblDistrict)
        Me.Controls.Add(Me.tbPostNumber)
        Me.Controls.Add(Me.lblPostNumber)
        Me.Controls.Add(Me.tbNumber)
        Me.Controls.Add(Me.lblNumber)
        Me.Controls.Add(Me.tbStreet)
        Me.Controls.Add(Me.lblStreet)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogAddress"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents lblStreet As System.Windows.Forms.Label
    Friend WithEvents tbStreet As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents tbNumber As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblNumber As System.Windows.Forms.Label
    Friend WithEvents tbPostNumber As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblPostNumber As System.Windows.Forms.Label
    Friend WithEvents tbDistrict As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblDistrict As System.Windows.Forms.Label
    Friend WithEvents tbTown As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblTown As System.Windows.Forms.Label
    Friend WithEvents tbCountry As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblCountry As System.Windows.Forms.Label
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cmbType As FaciliDev.LogFramer.ComboBoxSelectIndex

End Class
