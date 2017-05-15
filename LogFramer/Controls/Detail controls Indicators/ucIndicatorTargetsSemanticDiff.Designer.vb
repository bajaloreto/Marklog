<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IndicatorTargetsSemanticDiff
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(IndicatorTargetsSemanticDiff))
        Me.TabControlTargets = New System.Windows.Forms.TabControl()
        Me.TabPageBaseline = New System.Windows.Forms.TabPage()
        Me.TabControlTargets.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlTargets
        '
        resources.ApplyResources(Me.TabControlTargets, "TabControlTargets")
        Me.TabControlTargets.Controls.Add(Me.TabPageBaseline)
        Me.TabControlTargets.Name = "TabControlTargets"
        Me.TabControlTargets.SelectedIndex = 0
        '
        'TabPageBaseline
        '
        resources.ApplyResources(Me.TabPageBaseline, "TabPageBaseline")
        Me.TabPageBaseline.Name = "TabPageBaseline"
        Me.TabPageBaseline.UseVisualStyleBackColor = True
        '
        'IndicatorTargetsSemanticDiff
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TabControlTargets)
        Me.Name = "IndicatorTargetsSemanticDiff"
        Me.TabControlTargets.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControlTargets As System.Windows.Forms.TabControl
    Friend WithEvents TabPageBaseline As System.Windows.Forms.TabPage

End Class
