Public Class ListViewWebsites
    Inherits ListViewSortable

    Private objWebsites As Websites

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .OwnerDraw = True
            .Columns.Add(LANG_Type, 100, HorizontalAlignment.Left)
            .Columns.Add(LANG_Website, 200, HorizontalAlignment.Left)
            .Columns.Add(LANG_Link, 200, HorizontalAlignment.Left)
        End With
    End Sub

    Public Property Websites() As Websites
        Get
            Return objWebsites
        End Get
        Set(ByVal value As Websites)
            objWebsites = value
            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedWebsite() As Website
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selWebsite As Website = Me.Websites.GetWebsiteByGuid(Me.SelectedGuid)
                Return selWebsite
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedWebsites() As Website()
        Get
            Dim objWebsites(SelectedItems.Count - 1) As Website

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objWebsites(i) = Me.Websites.GetWebsiteByGuid(Me.SelectedGuids(i))
                Next

                Return objWebsites
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub LoadItems()
        Items.Clear()
        For Each selWebsite As Website In Me.Websites
            Dim newItem As New ListViewItem(selWebsite.TypeName)
            newItem.Name = selWebsite.Guid.ToString
            newItem.SubItems.Add(selWebsite.WebsiteName)
            Dim objUrl As New ListViewItem.ListViewSubItem
            objUrl.Text = selWebsite.WebsiteUrl
            objUrl.Tag = "link"
            newItem.SubItems.Add(objUrl)

            Items.Add(newItem)
        Next
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
    End Sub

    Public Overrides Sub NewItem()
        PopUpWebsiteDetails(Nothing)
    End Sub

    Public Overrides Sub EditItem()
        If Me.Websites.Count > 0 AndAlso Me.SelectedWebsite IsNot Nothing Then
            PopUpWebsiteDetails(Me.SelectedWebsite)
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.Websites.Count > 0 AndAlso Me.SelectedWebsite IsNot Nothing Then
            If MsgBox(LANG_RemoveWebsite, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selWebsite As Website = Me.SelectedWebsites(0)
                UndoRedo.ItemRemoved(selWebsite, Me.Websites)

                Me.Websites.Remove(selWebsite)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpWebsiteDetails(ByVal selWebsite As Website)
        Dim boolNew As Boolean

        If selWebsite Is Nothing Then
            boolNew = True
            selWebsite = New Website
        End If

        Dim dialogWebsite As New DialogWebsite(selWebsite)

        dialogWebsite.ShowDialog()
        If dialogWebsite.DialogResult = vbOK Then
            If boolNew = True Then
                Me.Websites.Add(selWebsite)
                UndoRedo.ItemInserted(selWebsite, Me.Websites)
            End If

            Me.LoadItems()
        End If
        dialogWebsite.Dispose()
        dialogWebsite = Nothing
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selWebsite As Website In SelectedWebsites
            UndoRedo.ItemCut(selWebsite, Me.Websites)

            Websites.Remove(selWebsite)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selWebsite As Website In SelectedWebsites
            Dim NewItem As New ClipboardItem(CopyGroup, selWebsite, Website.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selWebsite As Website

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Website)
                    selWebsite = CType(selItem.Item, Website)
                    Dim NewWebsite As New Website

                    Using copier As New ObjectCopy
                        NewWebsite = copier.CopyObject(selWebsite)
                    End Using
                    Me.Websites.Add(NewWebsite)

                    UndoRedo.ItemPasted(NewWebsite, Me.Websites)
            End Select
        Next

        Me.LoadItems()
    End Sub

    Private Sub ucContactsListViewWebsites_DrawColumnHeader(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs) Handles Me.DrawColumnHeader
        e.DrawDefault = True
    End Sub

    Private Sub ucContactsListViewWebsites_DrawSubItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawListViewSubItemEventArgs) Handles Me.DrawSubItem
        If e.ColumnIndex = 2 Then
            Dim fntLink As New Font(Me.Font, FontStyle.Underline)
            If (e.ItemState And ListViewItemStates.Selected) Then
                e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds)
                e.Graphics.DrawString(e.SubItem.Text, fntLink, SystemBrushes.HighlightText, e.Bounds)
            Else
                e.DrawBackground()
                e.Graphics.DrawString(e.SubItem.Text, fntLink, SystemBrushes.HotTrack, e.Bounds)
            End If


        Else
            e.DrawDefault = True
        End If
    End Sub

    Private Sub ucContactsListViewWebsites_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        Dim hitTest As ListViewHitTestInfo = Me.HitTest(e.X, e.Y)
        If hitTest.SubItem IsNot Nothing And (Control.ModifierKeys And Keys.Control) = Keys.Control Then
            If hitTest.SubItem.Tag = "link" Then
                OpenWebsite()
            End If
        End If
    End Sub

    Public Sub OpenWebsite()
        If Me.Websites.Count > 0 Then
            Dim intIndex As Integer
            If Me.SelectedIndices.Count > 0 Then intIndex = Me.SelectedIndices(0)
            Dim selWebsite As Website = Me.Websites(intIndex)
            Try
                Process.Start(selWebsite.WebsiteUrl)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub ucContactsListViewWebsites_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        IndicateHyperlink(New Point(e.X, e.Y))
    End Sub

    Private Sub IndicateHyperlink(ByVal hitPoint As Point)
        Dim hitTest As ListViewHitTestInfo = Me.HitTest(hitPoint)
        If hitTest.SubItem IsNot Nothing And (Control.ModifierKeys And Keys.Control) = Keys.Control Then
            If hitTest.SubItem.Tag = "link" Then
                Cursor = Cursors.Hand
            Else
                Cursor = Cursors.Default
            End If
        Else
            Cursor = Cursors.Default
        End If
    End Sub
End Class
