<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogWarning
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogWarning))
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.chkNotShow = New System.Windows.Forms.CheckBox()
        Me.btnYes = New System.Windows.Forms.Button()
        Me.btnNo = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblWarning
        '
        resources.ApplyResources(Me.lblWarning, "lblWarning")
        Me.lblWarning.Name = "lblWarning"
        '
        'chkNotShow
        '
        resources.ApplyResources(Me.chkNotShow, "chkNotShow")
        Me.chkNotShow.Name = "chkNotShow"
        Me.chkNotShow.UseVisualStyleBackColor = True
        '
        'btnYes
        '
        resources.ApplyResources(Me.btnYes, "btnYes")
        Me.btnYes.Name = "btnYes"
        Me.btnYes.UseVisualStyleBackColor = True
        '
        'btnNo
        '
        resources.ApplyResources(Me.btnNo, "btnNo")
        Me.btnNo.Name = "btnNo"
        Me.btnNo.UseVisualStyleBackColor = True
        '
        'lblTitle
        '
        resources.ApplyResources(Me.lblTitle, "lblTitle")
        Me.lblTitle.Name = "lblTitle"
        '
        'DialogWarning
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ControlBox = False
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnNo)
        Me.Controls.Add(Me.btnYes)
        Me.Controls.Add(Me.chkNotShow)
        Me.Controls.Add(Me.lblWarning)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogWarning"
        Me.ShowIcon = False
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblWarning As System.Windows.Forms.Label
    Friend WithEvents chkNotShow As System.Windows.Forms.CheckBox
    Friend WithEvents btnYes As System.Windows.Forms.Button
    Friend WithEvents btnNo As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
End Class
