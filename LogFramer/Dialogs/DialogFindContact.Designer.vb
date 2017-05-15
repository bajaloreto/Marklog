<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogFindContact
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogFindContact))
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.chkMatchWholeWord = New System.Windows.Forms.CheckBox()
        Me.chkMatchCase = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnFindNext = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnFindFirst = New System.Windows.Forms.Button()
        Me.tbFind = New System.Windows.Forms.TextBox()
        Me.lblFind = New System.Windows.Forms.Label()
        Me.GroupBoxSearchWhat = New System.Windows.Forms.GroupBox()
        Me.RadioButtonContacts = New System.Windows.Forms.RadioButton()
        Me.RadioButtonOrganisations = New System.Windows.Forms.RadioButton()
        Me.RadioButtonAll = New System.Windows.Forms.RadioButton()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.GroupBoxSearchWhat.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblMessage
        '
        resources.ApplyResources(Me.lblMessage, "lblMessage")
        Me.lblMessage.Name = "lblMessage"
        '
        'chkMatchWholeWord
        '
        resources.ApplyResources(Me.chkMatchWholeWord, "chkMatchWholeWord")
        Me.chkMatchWholeWord.Name = "chkMatchWholeWord"
        Me.chkMatchWholeWord.UseVisualStyleBackColor = True
        '
        'chkMatchCase
        '
        resources.ApplyResources(Me.chkMatchCase, "chkMatchCase")
        Me.chkMatchCase.Name = "chkMatchCase"
        Me.chkMatchCase.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        resources.ApplyResources(Me.TableLayoutPanel2, "TableLayoutPanel2")
        Me.TableLayoutPanel2.Controls.Add(Me.btnFindNext, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnClose, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnFindFirst, 1, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        '
        'btnFindNext
        '
        resources.ApplyResources(Me.btnFindNext, "btnFindNext")
        Me.btnFindNext.Name = "btnFindNext"
        '
        'btnClose
        '
        resources.ApplyResources(Me.btnClose, "btnClose")
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Name = "btnClose"
        '
        'btnFindFirst
        '
        resources.ApplyResources(Me.btnFindFirst, "btnFindFirst")
        Me.btnFindFirst.Name = "btnFindFirst"
        '
        'tbFind
        '
        resources.ApplyResources(Me.tbFind, "tbFind")
        Me.tbFind.Name = "tbFind"
        '
        'lblFind
        '
        resources.ApplyResources(Me.lblFind, "lblFind")
        Me.lblFind.Name = "lblFind"
        '
        'GroupBoxSearchWhat
        '
        resources.ApplyResources(Me.GroupBoxSearchWhat, "GroupBoxSearchWhat")
        Me.GroupBoxSearchWhat.Controls.Add(Me.RadioButtonContacts)
        Me.GroupBoxSearchWhat.Controls.Add(Me.RadioButtonOrganisations)
        Me.GroupBoxSearchWhat.Controls.Add(Me.RadioButtonAll)
        Me.GroupBoxSearchWhat.Name = "GroupBoxSearchWhat"
        Me.GroupBoxSearchWhat.TabStop = False
        '
        'RadioButtonContacts
        '
        resources.ApplyResources(Me.RadioButtonContacts, "RadioButtonContacts")
        Me.RadioButtonContacts.Name = "RadioButtonContacts"
        Me.RadioButtonContacts.TabStop = True
        Me.RadioButtonContacts.UseVisualStyleBackColor = True
        '
        'RadioButtonOrganisations
        '
        resources.ApplyResources(Me.RadioButtonOrganisations, "RadioButtonOrganisations")
        Me.RadioButtonOrganisations.Name = "RadioButtonOrganisations"
        Me.RadioButtonOrganisations.TabStop = True
        Me.RadioButtonOrganisations.UseVisualStyleBackColor = True
        '
        'RadioButtonAll
        '
        resources.ApplyResources(Me.RadioButtonAll, "RadioButtonAll")
        Me.RadioButtonAll.Name = "RadioButtonAll"
        Me.RadioButtonAll.TabStop = True
        Me.RadioButtonAll.UseVisualStyleBackColor = True
        '
        'DialogFindContact
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBoxSearchWhat)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.chkMatchWholeWord)
        Me.Controls.Add(Me.chkMatchCase)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Controls.Add(Me.tbFind)
        Me.Controls.Add(Me.lblFind)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogFindContact"
        Me.ShowInTaskbar = False
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.GroupBoxSearchWhat.ResumeLayout(False)
        Me.GroupBoxSearchWhat.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents chkMatchWholeWord As System.Windows.Forms.CheckBox
    Friend WithEvents chkMatchCase As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnFindNext As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFindFirst As System.Windows.Forms.Button
    Friend WithEvents tbFind As System.Windows.Forms.TextBox
    Friend WithEvents lblFind As System.Windows.Forms.Label
    Friend WithEvents GroupBoxSearchWhat As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonContacts As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonOrganisations As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonAll As System.Windows.Forms.RadioButton

End Class
