Public Class DetailBudgetItem
    Friend WithEvents lvBudgetItemReferences As ListViewBudgetItemReferences
    Public Event BudgetUpdateParentTotalsNeeded(ByVal sender As Object, ByVal e As UpdateParentTotalsEventArgs)

    Private BudgetItemBindingSource As New BindingSource
    Private objCurrentBudgetItem As BudgetItem
    Private objBudgetItemReferences As New Dictionary(Of Guid, String)
    Private intCurrentType As Integer
    Private boolTypeChanged As Boolean
    Private boolTextSelected As Boolean

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal budgetitem As BudgetItem, ByVal budgetitemreferences As Dictionary(Of Guid, String))
        InitializeComponent()

        Me.CurrentBudgetItem = budgetitem
        Me.BudgetItemReferences = budgetitemreferences
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

#Region "Properties"
    Public Property CurrentBudgetItem() As BudgetItem
        Get
            Return objCurrentBudgetItem
        End Get
        Set(ByVal value As BudgetItem)
            objCurrentBudgetItem = value
        End Set
    End Property

    Private Property CurrentType As Integer
        Get
            Return intCurrentType
        End Get
        Set(ByVal value As Integer)
            intCurrentType = value
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

    Public Property BudgetItemReferences As Dictionary(Of Guid, String)
        Get
            Return objBudgetItemReferences
        End Get
        Set(ByVal value As Dictionary(Of Guid, String))
            objBudgetItemReferences = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub LoadItems()
        If CurrentBudgetItem IsNot Nothing Then
            BudgetItemBindingSource.DataSource = CurrentBudgetItem

            tbBudgetItem.DataBindings.Add("Text", BudgetItemBindingSource, "Text")
            lvBudgetItemReferences = New ListViewBudgetItemReferences(CurrentBudgetItem)
            With lvBudgetItemReferences
                .Name = "lvBudgetItemReferences"
                .Dock = DockStyle.Fill
            End With

            TabPageReferences.Controls.Add(lvBudgetItemReferences)

            Select Case CurrentBudgetItem.Type
                Case BudgetItem.BudgetItemTypes.Normal, BudgetItem.BudgetItemTypes.Ratio, BudgetItem.BudgetItemTypes.Formula
                    With cmbType
                        .Items.AddRange(LIST_BudgetItemTypes)
                        .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                        .DropDownStyle = ComboBoxStyle.DropDownList
                        .DataBindings.Add("SelectedIndex", BudgetItemBindingSource, "Type")
                    End With

                    cmbType.Focus()

                    'set sub forms
                    ChangeUserControls(CurrentBudgetItem.Type)
                Case Else
                    TabControlBudgetItem.TabPages.Remove(TabPageType)
            End Select

        End If
    End Sub
#End Region

#Region "Events"
    Private Sub cmbType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        Dim intRaidType As Integer = cmbType.SelectedIndex

        If intRaidType <> CurrentType Then
            ChangeUserControls(intRaidType)
            CurrentType = intRaidType
        End If
    End Sub
#End Region

#Region "Add custom controls"
    Public Sub ChangeUserControls(ByVal selResponseType As Integer)
        'Change detail windows of tab pages
        PanelBudgetItemType.Controls.Clear()

        Select Case selResponseType
            Case BudgetItem.BudgetItemTypes.Normal

            Case BudgetItem.BudgetItemTypes.Ratio
                Dim ctlBudgetItem As New BudgetItemRatio(CurrentBudgetItem, Me.BudgetItemReferences)

                AddHandler ctlBudgetItem.RatioChanged, AddressOf RatioChanged
                ctlBudgetItem.Dock = DockStyle.Fill
                PanelBudgetItemType.Controls.Add(ctlBudgetItem)

            Case BudgetItem.BudgetItemTypes.Formula

        End Select

    End Sub

    Private Sub RatioChanged()
        RaiseEvent BudgetUpdateParentTotalsNeeded(Me, New UpdateParentTotalsEventArgs(CurrentBudgetItem))
    End Sub
#End Region
End Class
