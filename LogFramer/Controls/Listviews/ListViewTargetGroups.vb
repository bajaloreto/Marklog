Public Class ListViewTargetGroups
    Inherits ListViewSortable

    Private _TargetGroups As TargetGroups
    Private _AllTargetGroups As Boolean

    Public Event Updated()

#Region "Properties"
    Public Property TargetGroups() As TargetGroups
        Get
            Return _TargetGroups
        End Get
        Set(ByVal value As TargetGroups)
            _TargetGroups = value
        End Set
    End Property

    Public Property AllTargetGroups As Boolean
        Get
            Return _AllTargetGroups
        End Get
        Set(ByVal value As Boolean)
            _AllTargetGroups = value
        End Set
    End Property

    Public ReadOnly Property SelectedTargetGroups() As TargetGroup()
        Get
            Dim objTargetGroups(SelectedItems.Count - 1) As TargetGroup

            If Me.SelectedGuids IsNot Nothing AndAlso Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objTargetGroups(i) = Me.TargetGroups.GetTargetGroupByGuid(Me.SelectedGuids(i))
                Next

                Return objTargetGroups
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .MultiSelect = True

            .Columns.Add(LANG_TargetGroup, 250, HorizontalAlignment.Center)
            .Columns.Add(LANG_Type, 150, HorizontalAlignment.Left)
            .Columns.Add(LANG_Number, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_NumberFemales, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_NumberMales, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_NumberPeople, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_Location, 200, HorizontalAlignment.Left)
        End With
    End Sub

    Public Sub LoadTargetGroups(ByVal boolAllTargetGroups As Boolean)
        Me.AllTargetGroups = boolAllTargetGroups

        If Me.AllTargetGroups = True Then
            Me.TargetGroups = New TargetGroups
            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                    If selTargetGroup.ParentPurposeGuid = Nothing Or selTargetGroup.ParentPurposeGuid = Guid.Empty Then
                        selTargetGroup.ParentPurposeGuid = selPurpose.Guid
                    End If
                    Me.TargetGroups.Add(selTargetGroup)
                Next
            Next
        End If

        LoadItems()
    End Sub

    Public Sub LoadItems()
        Dim strText As String
        Dim boolHideTotalNumber As Boolean = True
        Dim boolHideNumberOfFemales As Boolean = True, boolHideNumberOfMales As Boolean = True
        Dim boolHideNumberOfPeople As Boolean = True

        If Me.TargetGroups IsNot Nothing Then
            Me.Items.Clear()
            For Each selTargetGroup As TargetGroup In Me.TargetGroups
                Dim newItem As New ListViewItem(selTargetGroup.Name)
                newItem.Name = selTargetGroup.Guid.ToString
                newItem.Tag = selTargetGroup.ParentPurposeGuid

                newItem.SubItems.Add(selTargetGroup.TypeName)
                If selTargetGroup.Number > 0 Then
                    strText = selTargetGroup.Number.ToString("N0")
                    boolHideTotalNumber = False
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)

                If selTargetGroup.NumberOfFemales > 0 Then
                    strText = selTargetGroup.NumberOfFemales.ToString("N0") & " (" & _
                        selTargetGroup.PercentageFemales.ToString("P2") & ")"
                    boolHideNumberOfFemales = False
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)

                If selTargetGroup.NumberOfMales > 0 Then
                    strText = selTargetGroup.NumberOfMales.ToString("N0") & " (" & _
                        selTargetGroup.PercentageMales.ToString("P2") & ")"
                    boolHideNumberOfMales = False
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)

                If selTargetGroup.NumberOfPeople > 0 Then
                    strText = selTargetGroup.NumberOfPeople.ToString("N0")
                    boolHideNumberOfPeople = False
                Else
                    strText = String.Empty
                End If
                newItem.SubItems.Add(strText)
                newItem.SubItems.Add(selTargetGroup.Location)
                Me.Items.Add(newItem)
            Next

            If Me.Columns.Count > 0 Then
                If boolHideTotalNumber = True Then Me.Columns(2).Width = 0 Else _
                    Me.Columns(2).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
                If boolHideNumberOfFemales = True Then Me.Columns(3).Width = 0 Else _
                    Me.Columns(3).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
                If boolHideNumberOfMales = True Then Me.Columns(4).Width = 0 Else _
                    Me.Columns(4).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
                If boolHideNumberOfPeople = True Then Me.Columns(5).Width = 0 Else _
                    Me.Columns(5).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
            End If

            RaiseEvent Updated()
        End If
    End Sub
#End Region

#Region "Events and methods"
    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Sub SetFocusOnItem(ByVal selTargetGroup As TargetGroup)
        Dim intIndex As Integer = Me.TargetGroups.IndexOf(selTargetGroup)

        If intIndex >= 0 And intIndex < Me.Items.Count Then
            Me.Items(intIndex).Focused = True
            Me.Items(intIndex).Selected = True
        End If
    End Sub

    Public Overrides Sub NewItem()
        PopUpTargetGroupDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.TargetGroups.Count > 0 AndAlso Me.SelectedTargetGroups.Length > 0 Then
            PopUpTargetGroupDetails(Me.SelectedTargetGroups(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.TargetGroups.Count > 0 AndAlso Me.SelectedTargetGroups.Length > 0 Then

            If MsgBox(LANG_RemoveTargetGroup, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selTargetGroup As TargetGroup = Me.SelectedTargetGroups(0)
                Dim objParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selTargetGroup.ParentPurposeGuid)

                UndoRedo.ItemRemoved(selTargetGroup, objParentPurpose.TargetGroups)

                objParentPurpose.TargetGroups.Remove(selTargetGroup)

                Me.LoadTargetGroups(Me.AllTargetGroups)
            End If
        End If
    End Sub

    Private Sub PopUpTargetGroupDetails(ByVal selTargetGroup As TargetGroup)
        Dim boolNew As Boolean

        If selTargetGroup Is Nothing Then
            boolNew = True
            selTargetGroup = New TargetGroup()
            selTargetGroup.Type = TargetGroup.TargetGroupTypes.Individual
            selTargetGroup.TargetGroupInformations.SetDefaultInformations(selTargetGroup.Type)

            If CurrentLogFrame.Purposes.Count = 0 Then
                CurrentLogFrame.Purposes.Add(New Purpose)
            End If

            selTargetGroup.ParentPurposeGuid = CurrentLogFrame.Purposes(0).Guid
        End If

        Dim objDialogTargetGroup As New DialogTargetGroup(selTargetGroup)
        Dim OldParentPurposeGuid As Guid = selTargetGroup.ParentPurposeGuid

        objDialogTargetGroup.ShowDialog(Me)
        If objDialogTargetGroup.DialogResult = vbOK Then
            Dim objParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selTargetGroup.ParentPurposeGuid)

            If boolNew = True Then
                objParentPurpose.TargetGroups.Add(selTargetGroup)
                UndoRedo.ItemInserted(selTargetGroup, objParentPurpose.TargetGroups)
            Else
                'user changed parent purpose (dropdown list)
                If selTargetGroup.ParentPurposeGuid <> OldParentPurposeGuid And AllTargetGroups = False Then
                    Me.TargetGroups.Remove(selTargetGroup)
                    'CurrentUndoList.DeleteOperation(selTargetGroup, Me.TargetGroups, , , True)

                    objParentPurpose.TargetGroups.Add(selTargetGroup)
                    'CurrentUndoList.InsertNewOperation(selTargetGroup, objParentPurpose.TargetGroups, , , True)
                End If
            End If

            LoadTargetGroups(Me.AllTargetGroups)
        End If

        objDialogTargetGroup = Nothing
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selTargetGroup As TargetGroup In SelectedTargetGroups
            Dim objParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(selTargetGroup.ParentPurposeGuid)

            UndoRedo.ItemCut(selTargetGroup, objParentPurpose.TargetGroups)

            objParentPurpose.TargetGroups.Remove(selTargetGroup)
        Next

        LoadTargetGroups(Me.AllTargetGroups)
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selTargetGroup As TargetGroup In SelectedTargetGroups
            Dim NewItem As New ClipboardItem(CopyGroup, selTargetGroup, TargetGroup.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selTargetGroup As TargetGroup
        Dim intCountName As Integer

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(TargetGroup)
                    selTargetGroup = CType(selItem.Item, TargetGroup)
                    Dim NewTargetGroup As New TargetGroup

                    Using copier As New ObjectCopy
                        NewTargetGroup = copier.CopyObject(selTargetGroup)
                    End Using

                    intCountName = TargetGroups.VerifyIfNameExists(NewTargetGroup.Name)
                    If intCountName > 0 Then
                        NewTargetGroup.Name &= String.Format(" ({0})", {intCountName})
                    End If

                    If AllTargetGroups = True Then
                        Dim objParentPurpose As Purpose = CurrentLogFrame.GetPurposeByGuid(NewTargetGroup.ParentPurposeGuid)
                        objParentPurpose.TargetGroups.Add(NewTargetGroup)

                        UndoRedo.ItemPasted(NewTargetGroup, objParentPurpose.TargetGroups)
                    Else
                        Me.TargetGroups.Add(NewTargetGroup)

                        UndoRedo.ItemPasted(NewTargetGroup, Me.TargetGroups)
                    End If
            End Select
        Next

        LoadTargetGroups(Me.AllTargetGroups)
    End Sub
#End Region
End Class



