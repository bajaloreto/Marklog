Public Class DetailResource
    Private ResourceBindingSource As New BindingSource
    Private objCurrentResource As Resource
    Private boolTextSelected As Boolean
    Friend WithEvents dgvBudgetItemReferences As DataGridViewBudgetItemReferences

    Public Sub New(ByRef resource As Resource)
        InitializeComponent()

        Me.CurrentResource = resource
        dgvBudgetItemReferences = New DataGridViewBudgetItemReferences(CurrentResource)
        With dgvBudgetItemReferences
            .Name = "dgvBudgetItemReferences"
            .Dock = DockStyle.Fill
        End With


        Me.PanelBudgetItems.Controls.Add(dgvBudgetItemReferences)
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Property CurrentResource() As Resource
        Get
            Return objCurrentResource
        End Get
        Set(ByVal value As Resource)
            objCurrentResource = value
        End Set
    End Property

    Public Property TextSelected() As Boolean
        Get
            Return boolTextSelected
        End Get
        Set(ByVal value As Boolean)
            boolTextSelected = value
            Me.Refresh()
        End Set
    End Property

    Public Sub LoadItems()
        If CurrentResource IsNot Nothing Then
            ResourceBindingSource.DataSource = CurrentResource

            tbResource.DataBindings.Add("Text", ResourceBindingSource, "Text")

            dgvBudgetItemReferences.Reload()
        End If
    End Sub

    Private Sub dgvBudgetItems_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvBudgetItemReferences.RowValidated
        UpdateResourceBudgetInLogFrame()
    End Sub

    Private Sub dgvBudgetItems_TotalCostChanged() Handles dgvBudgetItemReferences.TotalCostChanged
        UpdateResourceBudgetInLogFrame()
    End Sub

    Private Sub UpdateResourceBudgetInLogFrame()
        If dgvBudgetItemReferences.CurrentResource.BudgetItemReferences.Count > 0 Then
            CurrentResource.CellImageBudget = Nothing
            CurrentResource.SetTotalCostAmount()
            'CurrentResource.TotalCostAmount = dgvBudgetItemReferences.CurrentResource.BudgetItemReferences.GetTotalCost.Amount
            With CurrentProjectForm.dgvLogframe
                .InvalidateRow(.CurrentRow.Index)
            End With

            dgvBudgetItemReferences.InvalidateRow(dgvBudgetItemReferences.IndexTotal)
        End If
    End Sub
End Class
