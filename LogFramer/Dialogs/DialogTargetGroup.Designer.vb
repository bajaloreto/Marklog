<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogTargetGroup
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
        Dim tspEditProperties As System.Windows.Forms.ToolStrip
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogTargetGroup))
        Me.ToolStripButtonNewProperty = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonEditProperty = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButtonDeleteProperty = New System.Windows.Forms.ToolStripButton()
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButtonPrintIdForm = New System.Windows.Forms.ToolStripButton()
        Me.lblName = New System.Windows.Forms.Label()
        Me.tbName = New FaciliDev.LogFramer.TextBoxLF()
        Me.cmbType = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.lblNumberOfFemales = New System.Windows.Forms.Label()
        Me.lblNumberOfMales = New System.Windows.Forms.Label()
        Me.lblNumberOfPeople = New System.Windows.Forms.Label()
        Me.tbLocation = New FaciliDev.LogFramer.TextBoxLF()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblOtherProperties = New System.Windows.Forms.Label()
        Me.ntbNumber = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.ntbNumberOfPeople = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.ntbNumberOfMales = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.ntbNumberOfFemales = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.PanelProperties = New System.Windows.Forms.Panel()
        Me.lblPurpose = New System.Windows.Forms.Label()
        Me.cmbPurpose = New System.Windows.Forms.ComboBox()
        Me.btnReady = New System.Windows.Forms.Button()
        tspEditProperties = New System.Windows.Forms.ToolStrip()
        tspEditProperties.SuspendLayout()
        Me.PanelProperties.SuspendLayout()
        Me.SuspendLayout()
        '
        'tspEditProperties
        '
        resources.ApplyResources(tspEditProperties, "tspEditProperties")
        tspEditProperties.BackgroundImage = Global.FaciliDev.LogFramer.My.Resources.Resources.Background
        tspEditProperties.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButtonNewProperty, Me.ToolStripButtonEditProperty, Me.ToolStripButtonDeleteProperty, Me.toolStripSeparator, Me.ToolStripButtonPrintIdForm})
        tspEditProperties.Name = "tspEditProperties"
        '
        'ToolStripButtonNewProperty
        '
        resources.ApplyResources(Me.ToolStripButtonNewProperty, "ToolStripButtonNewProperty")
        Me.ToolStripButtonNewProperty.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonNewProperty.Name = "ToolStripButtonNewProperty"
        '
        'ToolStripButtonEditProperty
        '
        resources.ApplyResources(Me.ToolStripButtonEditProperty, "ToolStripButtonEditProperty")
        Me.ToolStripButtonEditProperty.CheckOnClick = True
        Me.ToolStripButtonEditProperty.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonEditProperty.Name = "ToolStripButtonEditProperty"
        '
        'ToolStripButtonDeleteProperty
        '
        resources.ApplyResources(Me.ToolStripButtonDeleteProperty, "ToolStripButtonDeleteProperty")
        Me.ToolStripButtonDeleteProperty.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonDeleteProperty.Name = "ToolStripButtonDeleteProperty"
        '
        'toolStripSeparator
        '
        resources.ApplyResources(Me.toolStripSeparator, "toolStripSeparator")
        Me.toolStripSeparator.Name = "toolStripSeparator"
        '
        'ToolStripButtonPrintIdForm
        '
        resources.ApplyResources(Me.ToolStripButtonPrintIdForm, "ToolStripButtonPrintIdForm")
        Me.ToolStripButtonPrintIdForm.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButtonPrintIdForm.Name = "ToolStripButtonPrintIdForm"
        '
        'lblName
        '
        resources.ApplyResources(Me.lblName, "lblName")
        Me.lblName.Name = "lblName"
        '
        'tbName
        '
        resources.ApplyResources(Me.tbName, "tbName")
        Me.tbName.Name = "tbName"
        '
        'cmbType
        '
        resources.ApplyResources(Me.cmbType, "cmbType")
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Name = "cmbType"
        '
        'lblType
        '
        resources.ApplyResources(Me.lblType, "lblType")
        Me.lblType.Name = "lblType"
        '
        'lblNumber
        '
        resources.ApplyResources(Me.lblNumber, "lblNumber")
        Me.lblNumber.Name = "lblNumber"
        '
        'lblNumberOfFemales
        '
        resources.ApplyResources(Me.lblNumberOfFemales, "lblNumberOfFemales")
        Me.lblNumberOfFemales.Name = "lblNumberOfFemales"
        '
        'lblNumberOfMales
        '
        resources.ApplyResources(Me.lblNumberOfMales, "lblNumberOfMales")
        Me.lblNumberOfMales.Name = "lblNumberOfMales"
        '
        'lblNumberOfPeople
        '
        resources.ApplyResources(Me.lblNumberOfPeople, "lblNumberOfPeople")
        Me.lblNumberOfPeople.Name = "lblNumberOfPeople"
        '
        'tbLocation
        '
        resources.ApplyResources(Me.tbLocation, "tbLocation")
        Me.tbLocation.Name = "tbLocation"
        '
        'lblLocation
        '
        resources.ApplyResources(Me.lblLocation, "lblLocation")
        Me.lblLocation.Name = "lblLocation"
        '
        'lblOtherProperties
        '
        resources.ApplyResources(Me.lblOtherProperties, "lblOtherProperties")
        Me.lblOtherProperties.Name = "lblOtherProperties"
        '
        'ntbNumber
        '
        resources.ApplyResources(Me.ntbNumber, "ntbNumber")
        Me.ntbNumber.AllowSpace = True
        Me.ntbNumber.IsCurrency = False
        Me.ntbNumber.IsPercentage = False
        Me.ntbNumber.Name = "ntbNumber"
        '
        'ntbNumberOfPeople
        '
        resources.ApplyResources(Me.ntbNumberOfPeople, "ntbNumberOfPeople")
        Me.ntbNumberOfPeople.AllowSpace = True
        Me.ntbNumberOfPeople.IsCurrency = False
        Me.ntbNumberOfPeople.IsPercentage = False
        Me.ntbNumberOfPeople.Name = "ntbNumberOfPeople"
        '
        'ntbNumberOfMales
        '
        resources.ApplyResources(Me.ntbNumberOfMales, "ntbNumberOfMales")
        Me.ntbNumberOfMales.AllowSpace = True
        Me.ntbNumberOfMales.IsCurrency = False
        Me.ntbNumberOfMales.IsPercentage = False
        Me.ntbNumberOfMales.Name = "ntbNumberOfMales"
        '
        'ntbNumberOfFemales
        '
        resources.ApplyResources(Me.ntbNumberOfFemales, "ntbNumberOfFemales")
        Me.ntbNumberOfFemales.AllowSpace = True
        Me.ntbNumberOfFemales.IsCurrency = False
        Me.ntbNumberOfFemales.IsPercentage = False
        Me.ntbNumberOfFemales.Name = "ntbNumberOfFemales"
        '
        'PanelProperties
        '
        resources.ApplyResources(Me.PanelProperties, "PanelProperties")
        Me.PanelProperties.Controls.Add(tspEditProperties)
        Me.PanelProperties.Name = "PanelProperties"
        '
        'lblPurpose
        '
        resources.ApplyResources(Me.lblPurpose, "lblPurpose")
        Me.lblPurpose.Name = "lblPurpose"
        '
        'cmbPurpose
        '
        resources.ApplyResources(Me.cmbPurpose, "cmbPurpose")
        Me.cmbPurpose.FormattingEnabled = True
        Me.cmbPurpose.Name = "cmbPurpose"
        '
        'btnReady
        '
        resources.ApplyResources(Me.btnReady, "btnReady")
        Me.btnReady.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnReady.Name = "btnReady"
        '
        'DialogTargetGroup
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnReady)
        Me.Controls.Add(Me.lblPurpose)
        Me.Controls.Add(Me.cmbPurpose)
        Me.Controls.Add(Me.PanelProperties)
        Me.Controls.Add(Me.lblOtherProperties)
        Me.Controls.Add(Me.tbLocation)
        Me.Controls.Add(Me.lblLocation)
        Me.Controls.Add(Me.lblNumber)
        Me.Controls.Add(Me.ntbNumber)
        Me.Controls.Add(Me.lblNumberOfPeople)
        Me.Controls.Add(Me.ntbNumberOfPeople)
        Me.Controls.Add(Me.lblNumberOfMales)
        Me.Controls.Add(Me.ntbNumberOfMales)
        Me.Controls.Add(Me.lblNumberOfFemales)
        Me.Controls.Add(Me.ntbNumberOfFemales)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.tbName)
        Me.Controls.Add(Me.lblName)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogTargetGroup"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        tspEditProperties.ResumeLayout(False)
        tspEditProperties.PerformLayout()
        Me.PanelProperties.ResumeLayout(False)
        Me.PanelProperties.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents tbName As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents cmbType As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents ntbNumber As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents lblNumber As System.Windows.Forms.Label
    Friend WithEvents lblNumberOfFemales As System.Windows.Forms.Label
    Friend WithEvents ntbNumberOfFemales As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents lblNumberOfMales As System.Windows.Forms.Label
    Friend WithEvents ntbNumberOfMales As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents lblNumberOfPeople As System.Windows.Forms.Label
    Friend WithEvents ntbNumberOfPeople As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents tbLocation As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblOtherProperties As System.Windows.Forms.Label
    Friend WithEvents PanelProperties As System.Windows.Forms.Panel
    Friend WithEvents ToolStripButtonNewProperty As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonEditProperty As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButtonDeleteProperty As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripButtonPrintIdForm As System.Windows.Forms.ToolStripButton
    Friend WithEvents lblPurpose As System.Windows.Forms.Label
    Friend WithEvents cmbPurpose As System.Windows.Forms.ComboBox
    Friend WithEvents btnReady As System.Windows.Forms.Button

End Class
