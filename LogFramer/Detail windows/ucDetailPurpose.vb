Public Class DetailPurpose
    Private PurposeBindingSource As New BindingSource
    Private indCurrentPurpose As Purpose

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        Me.Dock = DockStyle.Fill
    End Sub

    Public Sub New(ByVal purpose As Purpose)
        InitializeComponent()

        Me.CurrentPurpose = purpose
        Me.lvDetailTargetGroups.TargetGroups = Me.CurrentPurpose.TargetGroups
        Me.Dock = DockStyle.Fill
        LoadItems()
    End Sub

    Public Property CurrentPurpose() As Purpose
        Get
            Return indCurrentPurpose
        End Get
        Set(ByVal value As Purpose)
            indCurrentPurpose = value
        End Set
    End Property

    Public Sub LoadItems()
        If CurrentPurpose IsNot Nothing Then
            PurposeBindingSource.DataSource = CurrentPurpose

            tbPurpose.DataBindings.Add("Text", PurposeBindingSource, "Text")
            lvDetailTargetGroups.LoadTargetGroups(False)
        End If
    End Sub

    Private Sub lvDetailTargetGroups_Updated() Handles lvDetailTargetGroups.Updated
        CurrentProjectForm.ProjectInfo.LoadItems_TargetGroups()
    End Sub
End Class
