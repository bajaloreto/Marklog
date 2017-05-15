Public Class ListViewProcesses
    Inherits ListViewSortable

    Private objActivities As Activities
    Private strOutputString

    Public Event Updated()
    Public Event ActivityModified()

    Public Property OutputText() As String
        Get
            Return strOutputString
        End Get
        Set(ByVal value As String)
            strOutputString = value
        End Set
    End Property

    Public Property Activities() As Activities
        Get
            Return objActivities
        End Get
        Set(ByVal value As Activities)
            objActivities = value
            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedActivities() As Activity()
        Get
            Dim objActivities(SelectedItems.Count - 1) As Activity

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objActivities(i) = Me.Activities.GetActivityByGuid(Me.SelectedGuids(i))
                Next

                Return objActivities
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Activity, 300, HorizontalAlignment.Left)
            .Columns.Add(LANG_Type, 150, HorizontalAlignment.Left)
            .Columns.Add(LANG_StartDate, 120, HorizontalAlignment.Right)
            .Columns.Add(LANG_EndDate, 120, HorizontalAlignment.Right)
        End With
    End Sub

    Public Sub LoadItems()

        Me.Items.Clear()
        If Me.Activities IsNot Nothing Then
            For Each selActivity As Activity In Me.Activities
                Dim newItem As New ListViewItem(selActivity.Text)
                newItem.Name = selActivity.Guid.ToString
                newItem.SubItems.Add(selActivity.ActivityType)

                If selActivity.IsProcess Then
                    newItem.SubItems.Add(selActivity.GetProcessStartDate)
                    newItem.SubItems.Add(selActivity.GetProcessEndDate)
                Else
                    newItem.SubItems.Add(selActivity.ExactStartDate)
                    newItem.SubItems.Add(selActivity.ExactEndDate)
                End If
                Me.Items.Add(newItem)
            Next
            RaiseEvent Updated()
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        PopUpActivityDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.Activities.Count > 0 AndAlso Me.SelectedActivities.Length > 0 Then
            PopUpActivityDetails(Me.SelectedActivities(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.Activities.Count > 0 AndAlso Me.SelectedActivities.Length > 0 Then
            If MsgBox(LANG_RemoveActivity, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selActivity As Activity = Me.SelectedActivities(0)

                UndoRedo.ItemRemoved(selActivity, Me.Activities)

                Me.Activities.Remove(selActivity)
                Me.LoadItems()

                RaiseEvent ActivityModified()
            End If
        End If
    End Sub

    Private Sub PopUpActivityDetails(ByVal selActivity As Activity)
        Dim boolNew As Boolean

        If selActivity Is Nothing Then
            boolNew = True
            selActivity = New Activity()
        End If

        Dim dialogActivity As New DialogActivity(selActivity)

        dialogActivity.ShowDialog()
        If dialogActivity.DialogResult = vbOK Then
            If boolNew = True Then
                Me.Activities.AddToProcess(selActivity)

                UndoRedo.ItemInserted(selActivity, Me.Activities)
            End If

            Me.LoadItems()
            RaiseEvent ActivityModified()
        End If
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selActivity As Activity In SelectedActivities
            UndoRedo.ItemCut(selActivity, Me.Activities)

            Activities.Remove(selActivity)
        Next

        LoadItems()
        RaiseEvent ActivityModified()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selActivity As Activity In SelectedActivities
            Dim NewItem As New ClipboardItem(CopyGroup, selActivity, Activity.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selActivity As Activity

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Activity)
                    selActivity = CType(selItem.Item, Activity)
                    Dim NewActivity As New Activity

                    Using copier As New ObjectCopy
                        NewActivity = copier.CopyObject(selActivity)
                    End Using

                    Me.Activities.AddToProcess(NewActivity)

                    UndoRedo.ItemPasted(NewActivity, Me.Activities)
            End Select
        Next

        Me.LoadItems()
        RaiseEvent ActivityModified()
    End Sub
End Class

