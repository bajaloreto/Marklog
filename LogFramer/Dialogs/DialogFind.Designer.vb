<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogFind
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogFind))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnFindFirst = New System.Windows.Forms.Button()
        Me.btnFindNext = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.TabControlFindReplace = New System.Windows.Forms.TabControl()
        Me.TabPageFind = New System.Windows.Forms.TabPage()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.chkMatchWholeWord = New System.Windows.Forms.CheckBox()
        Me.chkMatchCase = New System.Windows.Forms.CheckBox()
        Me.chkSearchAllPanes = New System.Windows.Forms.CheckBox()
        Me.tbFind = New System.Windows.Forms.TextBox()
        Me.lblFind = New System.Windows.Forms.Label()
        Me.TabPageReplace = New System.Windows.Forms.TabPage()
        Me.tbFindReplace = New System.Windows.Forms.TextBox()
        Me.lblFindReplace = New System.Windows.Forms.Label()
        Me.lblMessageReplace = New System.Windows.Forms.Label()
        Me.chkReplaceWholeWord = New System.Windows.Forms.CheckBox()
        Me.chkReplaceMatchCase = New System.Windows.Forms.CheckBox()
        Me.chkReplaceDetails = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnReplaceAll = New System.Windows.Forms.Button()
        Me.btnReplaceNext = New System.Windows.Forms.Button()
        Me.btnReplaceClose = New System.Windows.Forms.Button()
        Me.tbReplace = New System.Windows.Forms.TextBox()
        Me.lblReplace = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TabControlFindReplace.SuspendLayout()
        Me.TabPageFind.SuspendLayout()
        Me.TabPageReplace.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.btnFindFirst, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnFindNext, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.btnClose, 2, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'btnFindFirst
        '
        resources.ApplyResources(Me.btnFindFirst, "btnFindFirst")
        Me.btnFindFirst.Name = "btnFindFirst"
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
        'TabControlFindReplace
        '
        resources.ApplyResources(Me.TabControlFindReplace, "TabControlFindReplace")
        Me.TabControlFindReplace.Controls.Add(Me.TabPageFind)
        Me.TabControlFindReplace.Controls.Add(Me.TabPageReplace)
        Me.TabControlFindReplace.Name = "TabControlFindReplace"
        Me.TabControlFindReplace.SelectedIndex = 0
        '
        'TabPageFind
        '
        resources.ApplyResources(Me.TabPageFind, "TabPageFind")
        Me.TabPageFind.Controls.Add(Me.lblMessage)
        Me.TabPageFind.Controls.Add(Me.chkMatchWholeWord)
        Me.TabPageFind.Controls.Add(Me.chkMatchCase)
        Me.TabPageFind.Controls.Add(Me.chkSearchAllPanes)
        Me.TabPageFind.Controls.Add(Me.TableLayoutPanel1)
        Me.TabPageFind.Controls.Add(Me.tbFind)
        Me.TabPageFind.Controls.Add(Me.lblFind)
        Me.TabPageFind.Name = "TabPageFind"
        Me.TabPageFind.UseVisualStyleBackColor = True
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
        'chkSearchAllPanes
        '
        resources.ApplyResources(Me.chkSearchAllPanes, "chkSearchAllPanes")
        Me.chkSearchAllPanes.Name = "chkSearchAllPanes"
        Me.chkSearchAllPanes.UseVisualStyleBackColor = True
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
        'TabPageReplace
        '
        resources.ApplyResources(Me.TabPageReplace, "TabPageReplace")
        Me.TabPageReplace.Controls.Add(Me.tbFindReplace)
        Me.TabPageReplace.Controls.Add(Me.lblFindReplace)
        Me.TabPageReplace.Controls.Add(Me.lblMessageReplace)
        Me.TabPageReplace.Controls.Add(Me.chkReplaceWholeWord)
        Me.TabPageReplace.Controls.Add(Me.chkReplaceMatchCase)
        Me.TabPageReplace.Controls.Add(Me.chkReplaceDetails)
        Me.TabPageReplace.Controls.Add(Me.TableLayoutPanel2)
        Me.TabPageReplace.Controls.Add(Me.tbReplace)
        Me.TabPageReplace.Controls.Add(Me.lblReplace)
        Me.TabPageReplace.Name = "TabPageReplace"
        Me.TabPageReplace.UseVisualStyleBackColor = True
        '
        'tbFindReplace
        '
        resources.ApplyResources(Me.tbFindReplace, "tbFindReplace")
        Me.tbFindReplace.Name = "tbFindReplace"
        '
        'lblFindReplace
        '
        resources.ApplyResources(Me.lblFindReplace, "lblFindReplace")
        Me.lblFindReplace.Name = "lblFindReplace"
        '
        'lblMessageReplace
        '
        resources.ApplyResources(Me.lblMessageReplace, "lblMessageReplace")
        Me.lblMessageReplace.Name = "lblMessageReplace"
        '
        'chkReplaceWholeWord
        '
        resources.ApplyResources(Me.chkReplaceWholeWord, "chkReplaceWholeWord")
        Me.chkReplaceWholeWord.Name = "chkReplaceWholeWord"
        Me.chkReplaceWholeWord.UseVisualStyleBackColor = True
        '
        'chkReplaceMatchCase
        '
        resources.ApplyResources(Me.chkReplaceMatchCase, "chkReplaceMatchCase")
        Me.chkReplaceMatchCase.Name = "chkReplaceMatchCase"
        Me.chkReplaceMatchCase.UseVisualStyleBackColor = True
        '
        'chkReplaceDetails
        '
        resources.ApplyResources(Me.chkReplaceDetails, "chkReplaceDetails")
        Me.chkReplaceDetails.Name = "chkReplaceDetails"
        Me.chkReplaceDetails.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        resources.ApplyResources(Me.TableLayoutPanel2, "TableLayoutPanel2")
        Me.TableLayoutPanel2.Controls.Add(Me.btnReplaceAll, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnReplaceNext, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnReplaceClose, 2, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        '
        'btnReplaceAll
        '
        resources.ApplyResources(Me.btnReplaceAll, "btnReplaceAll")
        Me.btnReplaceAll.Name = "btnReplaceAll"
        '
        'btnReplaceNext
        '
        resources.ApplyResources(Me.btnReplaceNext, "btnReplaceNext")
        Me.btnReplaceNext.Name = "btnReplaceNext"
        '
        'btnReplaceClose
        '
        resources.ApplyResources(Me.btnReplaceClose, "btnReplaceClose")
        Me.btnReplaceClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnReplaceClose.Name = "btnReplaceClose"
        '
        'tbReplace
        '
        resources.ApplyResources(Me.tbReplace, "tbReplace")
        Me.tbReplace.Name = "tbReplace"
        '
        'lblReplace
        '
        resources.ApplyResources(Me.lblReplace, "lblReplace")
        Me.lblReplace.Name = "lblReplace"
        '
        'DialogFind
        '
        Me.AcceptButton = Me.btnFindNext
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TabControlFindReplace)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogFind"
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TabControlFindReplace.ResumeLayout(False)
        Me.TabPageFind.ResumeLayout(False)
        Me.TabPageFind.PerformLayout()
        Me.TabPageReplace.ResumeLayout(False)
        Me.TabPageReplace.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnFindNext As System.Windows.Forms.Button
    Friend WithEvents TabControlFindReplace As System.Windows.Forms.TabControl
    Friend WithEvents TabPageFind As System.Windows.Forms.TabPage
    Friend WithEvents TabPageReplace As System.Windows.Forms.TabPage
    Friend WithEvents chkSearchAllPanes As System.Windows.Forms.CheckBox
    Friend WithEvents tbFind As System.Windows.Forms.TextBox
    Friend WithEvents lblFind As System.Windows.Forms.Label
    Friend WithEvents chkMatchWholeWord As System.Windows.Forms.CheckBox
    Friend WithEvents chkMatchCase As System.Windows.Forms.CheckBox
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents btnFindFirst As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents tbFindReplace As System.Windows.Forms.TextBox
    Friend WithEvents lblFindReplace As System.Windows.Forms.Label
    Friend WithEvents lblMessageReplace As System.Windows.Forms.Label
    Friend WithEvents chkReplaceWholeWord As System.Windows.Forms.CheckBox
    Friend WithEvents chkReplaceMatchCase As System.Windows.Forms.CheckBox
    Friend WithEvents chkReplaceDetails As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnReplaceAll As System.Windows.Forms.Button
    Friend WithEvents btnReplaceNext As System.Windows.Forms.Button
    Friend WithEvents btnReplaceClose As System.Windows.Forms.Button
    Friend WithEvents tbReplace As System.Windows.Forms.TextBox
    Friend WithEvents lblReplace As System.Windows.Forms.Label

End Class
