Public Class ListViewTargetGroupInformations
    Inherits ListViewSortable

    Private objTargetGroupInformations As TargetGroupInformations

    Public Property TargetGroupInformations() As TargetGroupInformations
        Get
            Return objTargetGroupInformations
        End Get
        Set(ByVal value As TargetGroupInformations)
            objTargetGroupInformations = value
            Me.LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedTargetGroupInformation() As TargetGroupInformation
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selTargetGroupInformation As TargetGroupInformation = Me.TargetGroupInformations.GetTargetGroupInformationByGuid(Me.SelectedGuid)
                Return selTargetGroupInformation
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedTargetGroupInformations() As TargetGroupInformation()
        Get
            Dim objTargetGroupInformations(SelectedItems.Count - 1) As TargetGroupInformation

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objTargetGroupInformations(i) = Me.TargetGroupInformations.GetTargetGroupInformationByGuid(Me.SelectedGuids(i))
                Next

                Return objTargetGroupInformations
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Name, 250, HorizontalAlignment.Left)
            .Columns.Add(LANG_ValueType, 150, HorizontalAlignment.Left)
            .Columns.Add(LANG_Values, 250, HorizontalAlignment.Left)
        End With
    End Sub

    Public Sub LoadItems()
        Dim strText As String = String.Empty


        Me.Items.Clear()
        For Each selTargetGroupInformation As TargetGroupInformation In Me.TargetGroupInformations
            Dim newItem As New ListViewItem(selTargetGroupInformation.Name)
            newItem.Name = selTargetGroupInformation.Guid.ToString
            newItem.SubItems.Add(selTargetGroupInformation.TypeName)
            If selTargetGroupInformation.CheckListOptions IsNot Nothing AndAlso selTargetGroupInformation.CheckListOptions.Count > 0 Then
                strText = String.Empty
                For Each selOption As ChecklistOption In selTargetGroupInformation.CheckListOptions
                    If String.IsNullOrEmpty(strText) = False Then strText &= ", "
                    strText &= selOption.OptionName
                Next
                newItem.SubItems.Add(strText)
            End If
            Me.Items.Add(newItem)
        Next

        If Me.Items.Count > 0 Then Me.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
    End Sub

    Public Overrides Sub NewItem()
        PopUpTargetGroupInformationDetails(Nothing)
    End Sub

    Public Overrides Sub EditItem()
        If Me.TargetGroupInformations.Count > 0 AndAlso Me.SelectedTargetGroupInformation IsNot Nothing Then
            PopUpTargetGroupInformationDetails(Me.SelectedTargetGroupInformation)
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.TargetGroupInformations.Count > 0 AndAlso Me.SelectedTargetGroupInformation IsNot Nothing Then
            If MsgBox(LANG_RemoveProperty, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selTargetGroupInformation As TargetGroupInformation = Me.SelectedTargetGroupInformations(0)
                Dim objParentTargetGroup As TargetGroup = CurrentLogFrame.GetTargetGroupByGuid(selTargetGroupInformation.ParentTargetGroupGuid)

                UndoRedo.ItemRemoved(selTargetGroupInformation, objParentTargetGroup.TargetGroupInformations)

                Me.TargetGroupInformations.Remove(Me.SelectedTargetGroupInformation)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpTargetGroupInformationDetails(ByVal selTargetGroupInformation As TargetGroupInformation)
        Dim boolNew As Boolean

        If selTargetGroupInformation Is Nothing Then
            boolNew = True
            selTargetGroupInformation = New TargetGroupInformation()
            selTargetGroupInformation.Type = TargetGroupInformation.PropertyTypes.Text
        End If

        Dim dialogTargetGroupInformation As New DialogTargetGroupInformation(selTargetGroupInformation)

        dialogTargetGroupInformation.ShowDialog(Me)
        If dialogTargetGroupInformation.DialogResult = vbOK Then
            If boolNew = True Then
                Me.TargetGroupInformations.Add(selTargetGroupInformation)
                UndoRedo.ItemInserted(selTargetGroupInformation, Me.TargetGroupInformations)
            End If

            Me.LoadItems()
        End If
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selTargetgroupInformation As TargetGroupInformation In SelectedTargetGroupInformations
            UndoRedo.ItemCut(selTargetgroupInformation, TargetGroupInformations)

            TargetGroupInformations.Remove(selTargetgroupInformation)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selTargetgroupInformation As TargetGroupInformation In SelectedTargetGroupInformations
            Dim NewItem As New ClipboardItem(CopyGroup, selTargetgroupInformation, TargetGroup.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selTargetGroupInformation As TargetGroupInformation
        Dim intCountName As Integer

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(TargetGroupInformation)
                    selTargetGroupInformation = CType(selItem.Item, TargetGroupInformation)
                    Dim NewTargetGroupInformation As New TargetGroupInformation

                    Using copier As New ObjectCopy
                        NewTargetGroupInformation = copier.CopyObject(selTargetGroupInformation)
                    End Using

                    intCountName = TargetGroupInformations.VerifyIfNameExists(NewTargetGroupInformation.Name)
                    If intCountName > 0 Then
                        NewTargetGroupInformation.Name &= String.Format(" ({0})", {intCountName})
                    End If
                    Me.TargetGroupInformations.Add(NewTargetGroupInformation)

                    UndoRedo.ItemPasted(NewTargetGroupInformation, TargetGroupInformations)
            End Select
        Next

        Me.LoadItems()
    End Sub
End Class
