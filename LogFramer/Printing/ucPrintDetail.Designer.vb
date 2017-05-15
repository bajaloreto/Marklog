<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucPrintDetail
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucPrintDetail))
        Me.SplitContainerPrint = New System.Windows.Forms.SplitContainer()
        Me.CustomPrintPreview = New FaciliDev.LogFramer.CustomPrintPreview()
        Me.ToolStripPreview = New System.Windows.Forms.ToolStrip()
        Me.ToolStripButtonFirst = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonPrevious = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripTextBoxPage = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripLabelPageCount = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripButtonNext = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonLast = New System.Windows.Forms.ToolStripButton()
        Me.cmbZoom = New System.Windows.Forms.ToolStripSplitButton()
        Me.ItemActualSize = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemFullPage = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemPageWidth = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemTwoPages = New System.Windows.Forms.ToolStripMenuItem()
        Me.toolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.Zoom500 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom200 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom150 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom100 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom75 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom50 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom25 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Zoom10 = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.SplitContainerPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainerPrint.Panel2.SuspendLayout()
        Me.SplitContainerPrint.SuspendLayout()
        Me.ToolStripPreview.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainerPrint
        '
        resources.ApplyResources(Me.SplitContainerPrint, "SplitContainerPrint")
        Me.SplitContainerPrint.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainerPrint.Name = "SplitContainerPrint"
        '
        'SplitContainerPrint.Panel1
        '
        resources.ApplyResources(Me.SplitContainerPrint.Panel1, "SplitContainerPrint.Panel1")
        '
        'SplitContainerPrint.Panel2
        '
        resources.ApplyResources(Me.SplitContainerPrint.Panel2, "SplitContainerPrint.Panel2")
        Me.SplitContainerPrint.Panel2.Controls.Add(Me.CustomPrintPreview)
        Me.SplitContainerPrint.Panel2.Controls.Add(Me.ToolStripPreview)
        '
        'CustomPrintPreview
        '
        resources.ApplyResources(Me.CustomPrintPreview, "CustomPrintPreview")
        Me.CustomPrintPreview.Document = Nothing
        Me.CustomPrintPreview.Name = "CustomPrintPreview"
        '
        'ToolStripPreview
        '
        resources.ApplyResources(Me.ToolStripPreview, "ToolStripPreview")
        Me.ToolStripPreview.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonFirst, Me.ToolStripButtonPrevious, Me.ToolStripTextBoxPage, Me.ToolStripLabelPageCount, Me.ToolStripButtonNext, Me.ToolStripButtonLast, Me.cmbZoom})
        Me.ToolStripPreview.Name = "ToolStripPreview"
        '
        'ToolStripButtonFirst
        '
        resources.ApplyResources(Me.ToolStripButtonFirst, "ToolStripButtonFirst")
        Me.ToolStripButtonFirst.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonFirst.Name = "ToolStripButtonFirst"
        '
        'ToolStripButtonPrevious
        '
        resources.ApplyResources(Me.ToolStripButtonPrevious, "ToolStripButtonPrevious")
        Me.ToolStripButtonPrevious.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonPrevious.Name = "ToolStripButtonPrevious"
        '
        'ToolStripTextBoxPage
        '
        resources.ApplyResources(Me.ToolStripTextBoxPage, "ToolStripTextBoxPage")
        Me.ToolStripTextBoxPage.Name = "ToolStripTextBoxPage"
        '
        'ToolStripLabelPageCount
        '
        resources.ApplyResources(Me.ToolStripLabelPageCount, "ToolStripLabelPageCount")
        Me.ToolStripLabelPageCount.Name = "ToolStripLabelPageCount"
        '
        'ToolStripButtonNext
        '
        resources.ApplyResources(Me.ToolStripButtonNext, "ToolStripButtonNext")
        Me.ToolStripButtonNext.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonNext.Name = "ToolStripButtonNext"
        '
        'ToolStripButtonLast
        '
        resources.ApplyResources(Me.ToolStripButtonLast, "ToolStripButtonLast")
        Me.ToolStripButtonLast.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripButtonLast.Name = "ToolStripButtonLast"
        '
        'cmbZoom
        '
        resources.ApplyResources(Me.cmbZoom, "cmbZoom")
        Me.cmbZoom.AutoToolTip = False
        Me.cmbZoom.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ItemActualSize, Me.ItemFullPage, Me.ItemPageWidth, Me.ItemTwoPages, Me.toolStripMenuItem1, Me.Zoom500, Me.Zoom200, Me.Zoom150, Me.Zoom100, Me.Zoom75, Me.Zoom50, Me.Zoom25, Me.Zoom10})
        Me.cmbZoom.Name = "cmbZoom"
        '
        'ItemActualSize
        '
        resources.ApplyResources(Me.ItemActualSize, "ItemActualSize")
        Me.ItemActualSize.Name = "ItemActualSize"
        '
        'ItemFullPage
        '
        resources.ApplyResources(Me.ItemFullPage, "ItemFullPage")
        Me.ItemFullPage.Name = "ItemFullPage"
        '
        'ItemPageWidth
        '
        resources.ApplyResources(Me.ItemPageWidth, "ItemPageWidth")
        Me.ItemPageWidth.Name = "ItemPageWidth"
        '
        'ItemTwoPages
        '
        resources.ApplyResources(Me.ItemTwoPages, "ItemTwoPages")
        Me.ItemTwoPages.Name = "ItemTwoPages"
        '
        'toolStripMenuItem1
        '
        resources.ApplyResources(Me.toolStripMenuItem1, "toolStripMenuItem1")
        Me.toolStripMenuItem1.Name = "toolStripMenuItem1"
        '
        'Zoom500
        '
        resources.ApplyResources(Me.Zoom500, "Zoom500")
        Me.Zoom500.Name = "Zoom500"
        '
        'Zoom200
        '
        resources.ApplyResources(Me.Zoom200, "Zoom200")
        Me.Zoom200.Name = "Zoom200"
        '
        'Zoom150
        '
        resources.ApplyResources(Me.Zoom150, "Zoom150")
        Me.Zoom150.Name = "Zoom150"
        '
        'Zoom100
        '
        resources.ApplyResources(Me.Zoom100, "Zoom100")
        Me.Zoom100.Name = "Zoom100"
        '
        'Zoom75
        '
        resources.ApplyResources(Me.Zoom75, "Zoom75")
        Me.Zoom75.Name = "Zoom75"
        '
        'Zoom50
        '
        resources.ApplyResources(Me.Zoom50, "Zoom50")
        Me.Zoom50.Name = "Zoom50"
        '
        'Zoom25
        '
        resources.ApplyResources(Me.Zoom25, "Zoom25")
        Me.Zoom25.Name = "Zoom25"
        '
        'Zoom10
        '
        resources.ApplyResources(Me.Zoom10, "Zoom10")
        Me.Zoom10.Name = "Zoom10"
        '
        'ucPrintDetail
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SplitContainerPrint)
        Me.Name = "ucPrintDetail"
        Me.SplitContainerPrint.Panel2.ResumeLayout(False)
        CType(Me.SplitContainerPrint, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainerPrint.ResumeLayout(False)
        Me.ToolStripPreview.ResumeLayout(False)
        Me.ToolStripPreview.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainerPrint As System.Windows.Forms.SplitContainer
    Friend WithEvents CustomPrintPreview As FaciliDev.LogFramer.CustomPrintPreview
    Friend WithEvents ToolStripPreview As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripButtonFirst As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonPrevious As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripTextBoxPage As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents ToolStripLabelPageCount As System.Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripButtonNext As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonLast As System.Windows.Forms.ToolStripButton
    Private WithEvents cmbZoom As System.Windows.Forms.ToolStripSplitButton
    Private WithEvents ItemActualSize As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ItemFullPage As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ItemPageWidth As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents ItemTwoPages As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents toolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Private WithEvents Zoom500 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom200 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom150 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom100 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom75 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom50 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom25 As System.Windows.Forms.ToolStripMenuItem
    Private WithEvents Zoom10 As System.Windows.Forms.ToolStripMenuItem

End Class
