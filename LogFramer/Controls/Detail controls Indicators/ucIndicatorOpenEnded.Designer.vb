<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorOpenEnded
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorOpenEnded))
        Me.lblWhiteSpace = New System.Windows.Forms.Label()
        Me.gbWhiteSpace = New System.Windows.Forms.GroupBox()
        Me.cmbWhiteSpace = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.gbWhiteSpace.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblWhiteSpace
        '
        resources.ApplyResources(Me.lblWhiteSpace, "lblWhiteSpace")
        Me.lblWhiteSpace.Name = "lblWhiteSpace"
        '
        'gbWhiteSpace
        '
        resources.ApplyResources(Me.gbWhiteSpace, "gbWhiteSpace")
        Me.gbWhiteSpace.Controls.Add(Me.cmbWhiteSpace)
        Me.gbWhiteSpace.Controls.Add(Me.lblWhiteSpace)
        Me.gbWhiteSpace.Name = "gbWhiteSpace"
        Me.gbWhiteSpace.TabStop = False
        '
        'cmbWhiteSpace
        '
        resources.ApplyResources(Me.cmbWhiteSpace, "cmbWhiteSpace")
        Me.cmbWhiteSpace.FormattingEnabled = True
        Me.cmbWhiteSpace.Name = "cmbWhiteSpace"
        '
        'IndicatorOpenEnded
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.gbWhiteSpace)
        Me.Name = "IndicatorOpenEnded"
        Me.gbWhiteSpace.ResumeLayout(False)
        Me.gbWhiteSpace.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblWhiteSpace As System.Windows.Forms.Label
    Friend WithEvents gbWhiteSpace As System.Windows.Forms.GroupBox
    Friend WithEvents cmbWhiteSpace As FaciliDev.LogFramer.ComboBoxSelectIndex

End Class
