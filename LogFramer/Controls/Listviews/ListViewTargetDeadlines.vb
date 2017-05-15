Public Class ListViewTargetDeadlines
    Inherits ListViewSortable

    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private strUnit, strValueName As String
    Private intNrDecimals As Integer
    Private intScoreSystem As Integer
    Private boolFromChildIndicator As Boolean

    Public Property TargetDeadlinesSection() As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesSection
        End Get
        Set(ByVal value As TargetDeadlinesSection)
            objTargetDeadlinesSection = value

            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedTargetDeadlines() As TargetDeadline()
        Get
            Dim objTargetDeadlines(SelectedItems.Count - 1) As TargetDeadline

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objTargetDeadlines(i) = TargetDeadlinesSection.TargetDeadlines.GetTargetDeadlineByGuid(Me.SelectedGuids(i))
                Next

                Return objTargetDeadlines
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Dates, 150, HorizontalAlignment.Center)
        End With
    End Sub

    Public Sub New(ByVal targetdeadlinessection As TargetDeadlinesSection)
        Me.TargetDeadlinesSection = targetdeadlinessection

        View = View.Details
        FullRowSelect = True
        Columns.Add(LANG_Dates, 150, HorizontalAlignment.Center)

        LoadItems()
    End Sub

    Public Sub LoadItems()

        Me.Items.Clear()
        If Me.TargetDeadlinesSection IsNot Nothing Then
            For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines.Sort
                Dim newItem As New ListViewItem(selTargetDeadline.ExactDeadline.ToShortDateString)
                newItem.Name = selTargetDeadline.Guid.ToString

                Me.Items.Add(newItem)
            Next
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        If Me.TargetDeadlinesSection.Repetition = TargetDeadlinesSection.RepetitionOptions.UserSelect Then
            PopUpTargetDeadlineDetails(Nothing)
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub EditItem()
        If Me.TargetDeadlinesSection.Repetition = TargetDeadlinesSection.RepetitionOptions.UserSelect Then
            If Me.TargetDeadlinesSection.TargetDeadlines.Count > 0 AndAlso Me.SelectedTargetDeadlines.Length > 0 Then
                PopUpTargetDeadlineDetails(Me.SelectedTargetDeadlines(0))
                CurrentControl = Me
            End If
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.TargetDeadlinesSection.Repetition = TargetDeadlinesSection.RepetitionOptions.UserSelect Then
            If Me.TargetDeadlinesSection.TargetDeadlines.Count > 0 AndAlso Me.SelectedTargetDeadlines.Length > 0 Then
                If MsgBox(LANG_RemoveTargetDeadline, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                    Dim selTargetDeadline As TargetDeadline = Me.SelectedTargetDeadlines(0)
                    'CurrentUndoList.DeleteOperation(selTargetDeadline, Me.TargetDeadlinesSection.TargetDeadlines, , , True)

                    TargetDeadlinesSection.TargetDeadlines.Remove(selTargetDeadline)
                    Me.LoadItems()
                End If
            End If
        End If
    End Sub

    Private Sub PopUpTargetDeadlineDetails(ByVal selTargetDeadline As TargetDeadline)
        Dim boolNew As Boolean

        If selTargetDeadline Is Nothing Then
            boolNew = True
            selTargetDeadline = New TargetDeadline()
        End If

        Dim dialogTargetDeadline As New DialogTargetDeadline(selTargetDeadline.Deadline)

        With dialogTargetDeadline
            If .ShowDialog = DialogResult.OK Then
                selTargetDeadline.Deadline = .Deadline


                If boolNew = True Then
                    TargetDeadlinesSection.TargetDeadlines.Add(selTargetDeadline)

                    'CurrentUndoList.InsertNewOperation(selTargetDeadline, TargetDeadlinesSection.TargetDeadlines, , , True)
                End If

                Me.LoadItems()
            End If
        End With
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        If Me.TargetDeadlinesSection.Repetition = TargetDeadlinesSection.RepetitionOptions.UserSelect Then
            For Each selTargetDeadline As TargetDeadline In SelectedTargetDeadlines
                'CurrentUndoList.CutOperation(selTargetDeadline, TargetDeadlinesSection.TargetDeadlines, TargetDeadlinesSection.TargetDeadlines.IndexOf(selTargetDeadline), , True)

                TargetDeadlinesSection.TargetDeadlines.Remove(selTargetDeadline)
            Next

            LoadItems()
        End If
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selTargetDeadline As TargetDeadline In SelectedTargetDeadlines
            Dim NewItem As New ClipboardItem(CopyGroup, selTargetDeadline, TargetDeadline.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selTargetDeadline As TargetDeadline

        If Me.TargetDeadlinesSection.Repetition = TargetDeadlinesSection.RepetitionOptions.UserSelect Then
            For i = 0 To PasteItems.Count - 1
                selItem = PasteItems(i)
                Select Case selItem.Item.GetType
                    Case GetType(TargetDeadline)
                        selTargetDeadline = CType(selItem.Item, TargetDeadline)
                        Dim NewTargetDeadline As New TargetDeadline

                        Using copier As New ObjectCopy
                            NewTargetDeadline = copier.CopyObject(selTargetDeadline)
                        End Using

                        TargetDeadlinesSection.TargetDeadlines.Add(NewTargetDeadline)

                        'CurrentUndoList.PasteOperation(NewTargetDeadline, TargetDeadlinesSection.TargetDeadlines, 0, , True)
                End Select
            Next

            Me.LoadItems()
        End If
    End Sub
End Class

