<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailVerificationSource
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailVerificationSource))
        Me.lblVerificationSource = New System.Windows.Forms.Label()
        Me.tbVerificationSource = New FaciliDev.LogFramer.TextBoxLF()
        Me.GroupBoxVerificationSource = New System.Windows.Forms.GroupBox()
        Me.tbCollectionMethod = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblCollectionMethod = New System.Windows.Forms.Label()
        Me.lblResponsibility = New System.Windows.Forms.Label()
        Me.cmbResponsibility = New FaciliDev.LogFramer.ComboBoxText()
        Me.tbWebsite = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblWebsite = New System.Windows.Forms.Label()
        Me.lblFrequency = New System.Windows.Forms.Label()
        Me.cmbFrequency = New FaciliDev.LogFramer.ComboBoxText()
        Me.lblSource = New System.Windows.Forms.Label()
        Me.cmbSource = New FaciliDev.LogFramer.ComboBoxText()
        Me.GroupBoxVerificationSource.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblVerificationSource
        '
        resources.ApplyResources(Me.lblVerificationSource, "lblVerificationSource")
        Me.lblVerificationSource.Name = "lblVerificationSource"
        '
        'tbVerificationSource
        '
        resources.ApplyResources(Me.tbVerificationSource, "tbVerificationSource")
        Me.tbVerificationSource.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbVerificationSource.ForeColor = System.Drawing.Color.Blue
        Me.tbVerificationSource.Name = "tbVerificationSource"
        Me.tbVerificationSource.ReadOnly = True
        '
        'GroupBoxVerificationSource
        '
        resources.ApplyResources(Me.GroupBoxVerificationSource, "GroupBoxVerificationSource")
        Me.GroupBoxVerificationSource.Controls.Add(Me.tbCollectionMethod)
        Me.GroupBoxVerificationSource.Controls.Add(Me.lblCollectionMethod)
        Me.GroupBoxVerificationSource.Controls.Add(Me.lblResponsibility)
        Me.GroupBoxVerificationSource.Controls.Add(Me.cmbResponsibility)
        Me.GroupBoxVerificationSource.Controls.Add(Me.tbWebsite)
        Me.GroupBoxVerificationSource.Controls.Add(Me.lblWebsite)
        Me.GroupBoxVerificationSource.Controls.Add(Me.lblFrequency)
        Me.GroupBoxVerificationSource.Controls.Add(Me.cmbFrequency)
        Me.GroupBoxVerificationSource.Controls.Add(Me.lblSource)
        Me.GroupBoxVerificationSource.Controls.Add(Me.cmbSource)
        Me.GroupBoxVerificationSource.Name = "GroupBoxVerificationSource"
        Me.GroupBoxVerificationSource.TabStop = False
        '
        'tbCollectionMethod
        '
        Me.tbCollectionMethod.AcceptsReturn = True
        resources.ApplyResources(Me.tbCollectionMethod, "tbCollectionMethod")
        Me.tbCollectionMethod.Name = "tbCollectionMethod"
        '
        'lblCollectionMethod
        '
        resources.ApplyResources(Me.lblCollectionMethod, "lblCollectionMethod")
        Me.lblCollectionMethod.Name = "lblCollectionMethod"
        '
        'lblResponsibility
        '
        resources.ApplyResources(Me.lblResponsibility, "lblResponsibility")
        Me.lblResponsibility.Name = "lblResponsibility"
        '
        'cmbResponsibility
        '
        resources.ApplyResources(Me.cmbResponsibility, "cmbResponsibility")
        Me.cmbResponsibility.FormattingEnabled = True
        Me.cmbResponsibility.Name = "cmbResponsibility"
        '
        'tbWebsite
        '
        resources.ApplyResources(Me.tbWebsite, "tbWebsite")
        Me.tbWebsite.Name = "tbWebsite"
        '
        'lblWebsite
        '
        resources.ApplyResources(Me.lblWebsite, "lblWebsite")
        Me.lblWebsite.Name = "lblWebsite"
        '
        'lblFrequency
        '
        resources.ApplyResources(Me.lblFrequency, "lblFrequency")
        Me.lblFrequency.Name = "lblFrequency"
        '
        'cmbFrequency
        '
        resources.ApplyResources(Me.cmbFrequency, "cmbFrequency")
        Me.cmbFrequency.FormattingEnabled = True
        Me.cmbFrequency.Name = "cmbFrequency"
        '
        'lblSource
        '
        resources.ApplyResources(Me.lblSource, "lblSource")
        Me.lblSource.Name = "lblSource"
        '
        'cmbSource
        '
        resources.ApplyResources(Me.cmbSource, "cmbSource")
        Me.cmbSource.FormattingEnabled = True
        Me.cmbSource.Name = "cmbSource"
        '
        'DetailVerificationSource
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.GroupBoxVerificationSource)
        Me.Controls.Add(Me.tbVerificationSource)
        Me.Controls.Add(Me.lblVerificationSource)
        Me.Name = "DetailVerificationSource"
        Me.GroupBoxVerificationSource.ResumeLayout(False)
        Me.GroupBoxVerificationSource.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblVerificationSource As System.Windows.Forms.Label
    Friend WithEvents tbVerificationSource As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents GroupBoxVerificationSource As System.Windows.Forms.GroupBox
    Friend WithEvents tbCollectionMethod As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblCollectionMethod As System.Windows.Forms.Label
    Friend WithEvents lblResponsibility As System.Windows.Forms.Label
    Friend WithEvents cmbResponsibility As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents tbWebsite As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblWebsite As System.Windows.Forms.Label
    Friend WithEvents lblFrequency As System.Windows.Forms.Label
    Friend WithEvents cmbFrequency As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblSource As System.Windows.Forms.Label
    Friend WithEvents cmbSource As FaciliDev.LogFramer.ComboBoxText

End Class
