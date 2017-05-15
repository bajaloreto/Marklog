Public Class ListViewEmails
    Inherits ListViewSortable

    Private objEmails As Emails

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Type, 100, HorizontalAlignment.Left)
            .Columns.Add(LANG_Email, 300, HorizontalAlignment.Left)
        End With
    End Sub

    Public Property Emails() As Emails
        Get
            Return objEmails
        End Get
        Set(ByVal value As Emails)
            objEmails = value
            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedEmail() As Email
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selEmail As Email = Me.Emails.GetEmailByGuid(Me.SelectedGuid)
                Return selEmail
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedEmails() As Email()
        Get
            Dim objEmails(SelectedItems.Count - 1) As Email

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objEmails(i) = Me.Emails.GetEmailByGuid(Me.SelectedGuids(i))
                Next

                Return objEmails
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub LoadItems()
        Items.Clear()
        For Each selEmail As Email In Me.Emails
            Dim newItem As New ListViewItem(selEmail.TypeName)
            newItem.Name = selEmail.Guid.ToString
            newItem.SubItems.Add(selEmail.Email)
            Items.Add(newItem)
        Next
        If Items.Count > 0 Then
            Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
    End Sub

    Public Overrides Sub NewItem()
        PopUpEmailDetails(Nothing)
    End Sub

    Public Overrides Sub EditItem()
        If Me.Emails.Count > 0 AndAlso Me.SelectedEmail IsNot Nothing Then
            PopUpEmailDetails(Me.SelectedEmail)
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.Emails.Count > 0 AndAlso Me.SelectedEmail IsNot Nothing Then
            If MsgBox(LANG_RemoveEmail, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selEmail As Email = Me.SelectedEmails(0)
                UndoRedo.ItemRemoved(selEmail, Me.Emails)

                Me.Emails.Remove(selEmail)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpEmailDetails(ByVal selEmail As Email)
        Dim boolNew As Boolean

        If selEmail Is Nothing Then
            boolNew = True
            selEmail = New Email
        End If

        Dim dialogEmail As New DialogEmail(selEmail)

        dialogEmail.ShowDialog()
        If dialogEmail.DialogResult = vbOK Then
            If boolNew = True Then
                Me.Emails.Add(selEmail)
                UndoRedo.ItemInserted(selEmail, Me.Emails)
            End If

            Me.LoadItems()
        End If
        dialogEmail.Dispose()
        dialogEmail = Nothing
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selEmail As Email In SelectedEmails
            UndoRedo.ItemCut(selEmail, Me.Emails)

            Emails.Remove(selEmail)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selEmail As Email In SelectedEmails
            Dim NewItem As New ClipboardItem(CopyGroup, selEmail, Email.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selEmail As Email

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Email)
                    selEmail = CType(selItem.Item, Email)
                    Dim NewEmail As New Email

                    Using copier As New ObjectCopy
                        NewEmail = copier.CopyObject(selEmail)
                    End Using
                    Me.Emails.Add(NewEmail)

                    UndoRedo.ItemPasted(NewEmail, Me.Emails)
            End Select
        Next

        Me.LoadItems()
    End Sub

    Public Sub SendEmail()
        Dim selEmail As Email = Nothing

        If SelectedEmail IsNot Nothing Then
            selEmail = SelectedEmail
        Else
            selEmail = Emails.GetMainEmail
        End If

        If selEmail IsNot Nothing Then
            Using objEmailIO As New EmailIO
                objEmailIO.StartMailClient(selEmail)
            End Using
        End If
    End Sub
End Class
