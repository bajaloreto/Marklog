Imports System.Windows.Forms

Public Class DialogTargetGroupInformation
    Private objProperty As TargetGroupInformation
    Private bindTargetGroupInformation As New BindingSource
    Private dtValues As New DataTable
    Friend WithEvents cmbUnit As New FaciliDev.LogFramer.StructuredComboBox
    Friend WithEvents dgvValues As New System.Windows.Forms.DataGridView
    Friend WithEvents cmbNrDecimals As New System.Windows.Forms.ComboBox

    Public Property TargetGroupInformation() As TargetGroupInformation
        Get
            Return objProperty
        End Get
        Set(ByVal value As TargetGroupInformation)
            objProperty = value
            LoadDataGrid()
        End Set
    End Property

    Public Sub New(ByVal targetgroupinfo As TargetGroupInformation)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        dtValues.Columns.Add(New DataColumn("Values", GetType(System.String)))
        Me.TargetGroupInformation = targetgroupinfo

        If targetgroupinfo IsNot Nothing Then
            bindTargetGroupInformation.DataSource = Me.TargetGroupInformation

            Me.tbName.DataBindings.Add("Text", bindTargetGroupInformation, "Name")
            With Me.cmbPropertyType
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDownList

                .Items.AddRange(LIST_TargetGroupInformationTypes)
                .DataBindings.Add("SelectedIndex", bindTargetGroupInformation, "Type")
            End With
            With Me.dgvValues
                .Dock = DockStyle.Fill
                .BackgroundColor = Color.White
                .ColumnHeadersVisible = False
                .RowHeadersVisible = True
                .DataSource = dtValues
                .AutoGenerateColumns = True
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
            End With

            With Me.cmbNrDecimals
                .Location = New Point(0, 0)
                For i = 0 To 6
                    .Items.Add(i)
                Next
                .DataBindings.Add("SelectedIndex", bindTargetGroupInformation, "NrDecimals")
            End With

            Dim MeasureUnits As List(Of StructuredComboBoxItem) = LoadMeasureUnits()
            With Me.cmbUnit
                .Location = New Point(0, 27)
                .Width = PanelOptions.Width
                .Anchor = AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Top
                .Items.AddRange(MeasureUnits.ToArray)
                .DisplayMember = "Description"
                .ValueMember = "Unit"
                .DataBindings.Add("SelectedText", bindTargetGroupInformation, "Unit")
            End With
        End If
    End Sub

    Private Sub LoadDataGrid()
        dtValues.Rows.Clear()
        If cmbPropertyType.SelectedIndex = TargetGroupInformation.PropertyTypes.List And _
            (Me.TargetGroupInformation.CheckListOptions IsNot Nothing AndAlso _
             Me.TargetGroupInformation.CheckListOptions.Count > 0) Then

            For Each selOption As ChecklistOption In Me.TargetGroupInformation.CheckListOptions
                dtValues.Rows.Add(selOption.OptionName)
            Next
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.TargetGroupInformation.Type = TargetGroupInformation.PropertyTypes.List Then
            Me.TargetGroupInformation.CheckListOptions.Clear()
            For Each selRow As DataRow In dtValues.Rows
                Me.TargetGroupInformation.CheckListOptions.Add(selRow(0))
            Next
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub cmbPropertyType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPropertyType.SelectedIndexChanged
        PanelOptions.Controls.Clear()
        
        Select Case cmbPropertyType.SelectedIndex
            Case TargetGroupInformation.PropertyTypes.Number
                lblOption1.Text = ToLabel(LANG_Decimals)
                lblOption2.Text = ToLabel(LANG_Unit)
                PanelOptions.Controls.Add(cmbNrDecimals)
                PanelOptions.Controls.Add(cmbUnit)
            Case TargetGroupInformation.PropertyTypes.List
                lblOption1.Text = ToLabel(LANG_Options)
                lblOption2.Text = String.Empty
                PanelOptions.Controls.Add(dgvValues)
            Case Else
                lblOption1.Text = String.Empty
                lblOption2.Text = String.Empty
        End Select

        LoadDataGrid()
    End Sub
End Class
