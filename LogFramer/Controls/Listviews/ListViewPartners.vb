Public Class ListViewPartners
    Inherits ListViewSortable

    Public Event ListUpdated()

    Private objPartners As ProjectPartners

    Public Property ProjectPartners() As ProjectPartners
        Get
            Return objPartners
        End Get
        Set(ByVal value As ProjectPartners)
            objPartners = value
            Me.LoadItems()
        End Set
    End Property

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Name, 250, HorizontalAlignment.Left)
            .Columns.Add(LANG_Role, 150, HorizontalAlignment.Left)
            .Columns.Add(LANG_Address, 250, HorizontalAlignment.Left)
        End With
    End Sub

    Public ReadOnly Property SelectedProjectPartner() As ProjectPartner
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selProjectPartner As ProjectPartner = Me.ProjectPartners.GetProjectPartnerByGuid(Me.SelectedGuid)
                Return selProjectPartner
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedProjectPartners() As ProjectPartner()
        Get
            Dim objProjectPartners(SelectedItems.Count - 1) As ProjectPartner

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objProjectPartners(i) = Me.ProjectPartners.GetProjectPartnerByGuid(Me.SelectedGuids(i))
                Next

                Return objProjectPartners
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub LoadItems()
        Me.Items.Clear()
        If objPartners IsNot Nothing Then
            For Each selProjectPartner As ProjectPartner In Me.ProjectPartners
                Dim newItem As New ListViewItem(selProjectPartner.Organisation.FullName)
                newItem.Name = selProjectPartner.Guid.ToString
                newItem.SubItems.Add(selProjectPartner.RoleName)
                newItem.SubItems.Add(selProjectPartner.Organisation.MainAddress)
                Me.Items.Add(newItem)
            Next
        End If
        If Items.Count > 0 Then
            AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
    End Sub

    Public Overrides Sub NewItem()
        'AddNewProjectPartner() --> Logframer 3.0
        PopUpProjectPartner(Nothing)

        RaiseEvent ListUpdated()
    End Sub

    Public Overrides Sub EditItem()
        If Me.ProjectPartners.Count > 0 AndAlso Me.SelectedProjectPartner IsNot Nothing Then
            PopUpProjectPartner(Me.SelectedProjectPartner)
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.ProjectPartners.Count > 0 AndAlso Me.SelectedProjectPartner IsNot Nothing Then
            If MsgBox(LANG_RemoveProjectPartner, MsgBoxStyle.OkCancel, LANG_Remove) = _
                MsgBoxResult.Ok Then

                UndoRedo.ItemRemoved(SelectedProjectPartner, Me.ProjectPartners)
                Me.ProjectPartners.Remove(Me.SelectedProjectPartner)
                LoadItems()

                RaiseEvent ListUpdated()
            End If
        End If
    End Sub

    Private Sub AddNewProjectPartner()
        'Logframer 3.0

        'Dim dialogAddPartner As New DialogAddProjectPartner

        'With dialogAddPartner
        '    .ShowDialog(Me)
        'End With
    End Sub

    Private Sub PopUpProjectPartner(ByVal selProjectPartner As ProjectPartner)
        Dim boolNew As Boolean
        Dim CopyProjectPartner As New ProjectPartner

        If selProjectPartner Is Nothing Then
            boolNew = True
            selProjectPartner = New ProjectPartner
        End If

        Dim dialogProjectPartner As New DialogProjectPartner(selProjectPartner)

        With dialogProjectPartner
            .ShowDialog(Me)
            If .DialogResult = vbOK Then
                If boolNew = True Then
                    Me.ProjectPartners.Add(selProjectPartner)

                    UndoRedo.ItemInserted(selProjectPartner, Me.ProjectPartners)
                End If

                Me.LoadItems()
            End If
        End With
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selProjectPartner As ProjectPartner In SelectedProjectPartners
            UndoRedo.ItemCut(selProjectPartner, ProjectPartners)

            ProjectPartners.Remove(selProjectPartner)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selProjectPartner As ProjectPartner In SelectedProjectPartners
            Dim NewItem As New ClipboardItem(CopyGroup, selProjectPartner, ProjectPartner.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selProjectPartner As ProjectPartner
        'Dim intCountName As Integer

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(ProjectPartner)
                    selProjectPartner = CType(selItem.Item, ProjectPartner)
                    Dim NewProjectPartner As New ProjectPartner

                    Using copier As New ObjectCopy
                        NewProjectPartner = copier.CopyObject(selProjectPartner)
                    End Using

                    'intCountName = ProjectPartners.VerifyIfNameExists(NewProjectPartner.Name)
                    'If intCountName > 0 Then
                    '    NewProjectPartner.Name &= String.Format(" ({0})", {intCountName})
                    'End If
                    Me.ProjectPartners.Add(NewProjectPartner)

                    UndoRedo.ItemPasted(NewProjectPartner, Me.ProjectPartners)
            End Select
        Next

        Me.LoadItems()
    End Sub
End Class

