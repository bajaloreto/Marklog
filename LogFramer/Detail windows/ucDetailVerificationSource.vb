Public Class DetailVerificationSource
    Public VerificationSourceBindingSource As New BindingSource
    Private ResponsesBindingSource As New BindingSource
    Private objCurrentVerificationSource As VerificationSource
    Private boolTypeChanged As Boolean
    Private boolTextSelected As Boolean

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal verificationsource As VerificationSource)
        InitializeComponent()

        Me.CurrentVerificationSource = verificationsource
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Property CurrentVerificationSource() As VerificationSource
        Get
            Return objCurrentVerificationSource
        End Get
        Set(ByVal value As VerificationSource)
            objCurrentVerificationSource = value
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
        If CurrentVerificationSource IsNot Nothing Then
            VerificationSourceBindingSource.DataSource = CurrentVerificationSource

            tbVerificationSource.DataBindings.Add("Text", VerificationSourceBindingSource, "Text")
            With cmbResponsibility
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .DataSource = CurrentLogframe.PartnerNamesList
                .DataBindings.Add("Text", VerificationSourceBindingSource, "Responsibility")
            End With
            With cmbFrequency
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .Items.AddRange(LIST_MonitoringFrequency)
                .DataBindings.Add("Text", VerificationSourceBindingSource, "Frequency")
            End With
            With cmbSource
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .DropDownStyle = ComboBoxStyle.DropDown
                .Items.AddRange(LIST_MonitoringSource)
                .DataBindings.Add("Text", VerificationSourceBindingSource, "Source")
            End With
            tbWebsite.DataBindings.Add("Text", VerificationSourceBindingSource, "Website")
            tbCollectionMethod.DataBindings.Add("Text", VerificationSourceBindingSource, "CollectionMethod")
        End If
    End Sub

    Private Sub cmbSource_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSource.SelectedIndexChanged
        If cmbSource.Text = LANG_Website Then
            tbWebsite.Enabled = True
        Else
            tbWebsite.Enabled = False
        End If
    End Sub
End Class
