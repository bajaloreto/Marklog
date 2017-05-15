Public Class ListViewCloudProjects
    Inherits ListViewSortable

    Private objLogframes As Logframes

    Public Property Logframes() As Logframes
        Get
            Return objLogframes
        End Get
        Set(ByVal value As Logframes)
            objLogframes = value
            Me.LoadItems()
        End Set
    End Property

    Public Sub New()
        Dim strProjectTitle, strShortTitle, strCode, strStartDate, strEndDate As String

        Select Case UserLanguage
            Case "fr"
                strCode = "Code"
                strProjectTitle = "Titre"
                strShortTitle = "Titre abrégé"
                strStartDate = "Date de début"
                strEndDate = "Date de fin"
            Case Else
                strCode = "Code"
                strProjectTitle = "Title"
                strShortTitle = "Short title"
                strStartDate = "Start date"
                strEndDate = "End date"
        End Select

        With Me
            .View = View.Details
            .FullRowSelect = True
            '.Columns.Add(strCode, 100, HorizontalAlignment.Left)
            '.Columns.Add(strShortTitle, 150, HorizontalAlignment.Left)
            .Columns.Add(strProjectTitle, 250, HorizontalAlignment.Left)
            .Columns.Add(strStartDate, 100, HorizontalAlignment.Left)
            .Columns.Add(strEndDate, 100, HorizontalAlignment.Left)
        End With
    End Sub

    Public ReadOnly Property SelectedLogframe() As LogFrame
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selLogframe As Logframe = Me.Logframes.GetLogframeByGuid(Me.SelectedGuid)
                Return selLogframe
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub LoadItems()
        Me.Items.Clear()
        If objLogframes IsNot Nothing Then
            For Each selLogframe As Logframe In Me.Logframes
                Dim newItem As New ListViewItem(selLogframe.ProjectTitle)
                newItem.Name = selLogframe.Guid.ToString
                newItem.SubItems.Add(selLogframe.StartDate)
                newItem.SubItems.Add(selLogframe.EndDate)
                Me.Items.Add(newItem)
            Next
        End If
        If Items.Count > 0 Then
            AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)

    End Sub

    Public Overrides Sub NewItem()

    End Sub

    Public Overrides Sub EditItem()

    End Sub

    Public Overrides Sub RemoveItem()

    End Sub

    Public Overrides Sub CutItems()

    End Sub

    Public Overrides Sub CopyItems()
        
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        
    End Sub

    'Public Sub RemoveLogframe()
    '    Dim strMsg, strTitle As String

    '    Select Case UserLanguage
    '        Case "fr"
    '            strMsg = "Supprimer cette Logframe?"
    '            strTitle = "Supprimer"
    '        Case Else
    '            strMsg = "Remove the selected Logframe?"
    '            strTitle = "Remove"
    '    End Select

    '    If Me.SelectedLogframe IsNot Nothing Then
    '        If MsgBox(strMsg, MsgBoxStyle.OkCancel, strTitle) = _
    '            MsgBoxResult.Ok Then
    '            UndoListLF.SetUndoBuffer(Me.SelectedLogframe.Guid, LogFrame.ObjectTypes.LogFrame, Me.SelectedLogframe.ProjectTitle, UndoListItem.Actions.Delete, Nothing, _
    '                                      Me.SelectedLogframe, Me.Logframes, Me.Logframes.IndexOf(Me.SelectedLogframe), "")
    '            UndoListLF.AddToUndoList()
    '            Me.Logframes.Remove(Me.SelectedLogframe)
    '            LoadItems()
    '        End If
    '    End If
    'End Sub


End Class


