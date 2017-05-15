Public Class ListViewKeyMoments
    Inherits ListViewSortable

    Private objKeyMoments As KeyMoments
    Private strOutputString

#Region "Properties"
    Public Property OutputText() As String
        Get
            Return strOutputString
        End Get
        Set(ByVal value As String)
            strOutputString = value
        End Set
    End Property

    Public Property KeyMoments() As KeyMoments
        Get
            Return objKeyMoments
        End Get
        Set(ByVal value As KeyMoments)
            objKeyMoments = value
            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedKeyMoments() As KeyMoment()
        Get
            Dim objKeyMoments(SelectedItems.Count - 1) As KeyMoment

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objKeyMoments(i) = Me.KeyMoments.GetKeyMomentByGuid(Me.SelectedGuids(i))
                Next

                Return objKeyMoments
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
            .Columns.Add(LANG_Date, 150, HorizontalAlignment.Center)
            .Columns.Add(LANG_KeyMoment, 500, HorizontalAlignment.Left)
        End With
    End Sub

    Public Sub LoadItems()

        Me.Items.Clear()
        If Me.KeyMoments IsNot Nothing Then
            For Each selKeyMoment As KeyMoment In Me.KeyMoments
                Dim newItem As New ListViewItem(selKeyMoment.ExactDateKeyMoment.ToShortDateString)
                newItem.Name = selKeyMoment.Guid.ToString
                newItem.SubItems.Add(selKeyMoment.Description)
                Me.Items.Add(newItem)
            Next
        End If
    End Sub
#End Region

#Region "Events and methods"
    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Sub SetFocusOnItem(ByVal selKeyMoment As KeyMoment)
        Dim intIndex As Integer = Me.KeyMoments.IndexOf(selKeyMoment)

        If intIndex >= 0 And intIndex < Me.Items.Count Then
            Me.Items(intIndex).Focused = True
            Me.Items(intIndex).Selected = True
        End If
    End Sub

    Public Overrides Sub NewItem()
        PopUpKeyMomentDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.KeyMoments.Count > 0 AndAlso Me.SelectedKeyMoments.Length > 0 Then
            PopUpKeyMomentDetails(Me.SelectedKeyMoments(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.KeyMoments.Count > 0 AndAlso Me.SelectedKeyMoments.Length > 0 Then
            If MsgBox(LANG_RemoveKeyMoment, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selKeyMoment As KeyMoment = Me.SelectedKeyMoments(0)
                Dim objParentOutput As Output = CurrentLogFrame.GetOutputByGuid(selKeyMoment.ParentOutputGuid)

                'change relative start date settings of items that refer to this key moment
                RemoveItem_Referers(selKeyMoment.Guid)

                UndoRedo.ItemRemoved(selKeyMoment, objParentOutput.KeyMoments)

                Me.KeyMoments.Remove(selKeyMoment)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub RemoveItem_Referers(ByVal objReferenceGuid As Guid)
        Dim ReferersList As ArrayList = CurrentLogFrame.GetReferingMomentsByReferenceGuid(objReferenceGuid)

        If ReferersList.Count > 0 Then
            For i = 0 To ReferersList.Count - 1
                Select Case ReferersList(i).GetType
                    Case GetType(KeyMoment)
                        Dim selKeyMoment As KeyMoment = ReferersList(i)

                        With selKeyMoment
                            .KeyMoment = .ExactDateKeyMoment
                            .Relative = False
                            .Period = 0
                            .PeriodDirection = KeyMoment.DirectionValues.NotDefined
                            .PeriodUnit = DurationUnits.NotDefined
                            .GuidReferenceMoment = Guid.Empty
                        End With
                    Case GetType(Activity)
                        Dim selActivity As Activity = ReferersList(i)

                        With selActivity.ActivityDetail
                            .StartDate = .StartDateMainActivity
                            .Relative = False
                            .Period = 0
                            .PeriodDirection = ActivityDetail.DirectionValues.NotDefined
                            .PeriodUnit = ActivityDetail.DurationUnits.NotDefined
                            .GuidReferenceMoment = Guid.Empty
                        End With
                End Select
            Next
        End If
    End Sub

    Private Sub PopUpKeyMomentDetails(ByVal selKeyMoment As KeyMoment)
        Dim boolNew As Boolean

        If selKeyMoment Is Nothing Then
            boolNew = True
            selKeyMoment = New KeyMoment()
        End If

        Dim dialogKeyMoment As New DialogKeyMoment(selKeyMoment)
        Dim OldParentOutputGuid As Guid = selKeyMoment.ParentOutputGuid

        dialogKeyMoment.ShowDialog()
        If dialogKeyMoment.DialogResult = vbOK Then
            Dim objParentOutput As Output = CurrentLogFrame.GetOutputByGuid(selKeyMoment.ParentOutputGuid)

            If boolNew = True Then
                objParentOutput.KeyMoments.Add(selKeyMoment)

                UndoRedo.ItemInserted(selKeyMoment, objParentOutput.KeyMoments)
            Else
                'user changed parent purpose (dropdown list)
                If selKeyMoment.ParentOutputGuid <> OldParentOutputGuid Then
                    Me.KeyMoments.Remove(selKeyMoment)

                    objParentOutput.KeyMoments.Add(selKeyMoment)
                End If
            End If

            Me.LoadItems()
        End If
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selKeyMoment As KeyMoment In SelectedKeyMoments
            UndoRedo.ItemCut(selKeyMoment, KeyMoments)

            KeyMoments.Remove(selKeyMoment)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selKeyMoment As KeyMoment In SelectedKeyMoments
            Dim NewItem As New ClipboardItem(CopyGroup, selKeyMoment, KeyMoment.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selKeyMoment As KeyMoment
        Dim intCountDescription As Integer

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(KeyMoment)
                    selKeyMoment = CType(selItem.Item, KeyMoment)
                    Dim NewKeyMoment As New KeyMoment

                    Using copier As New ObjectCopy
                        NewKeyMoment = copier.CopyObject(selKeyMoment)
                    End Using

                    intCountDescription = KeyMoments.VerifyIfDescriptionExists(NewKeyMoment.Description)
                    If intCountDescription > 0 Then
                        NewKeyMoment.Description &= String.Format(" ({0})", {intCountDescription})
                    End If

                    Me.KeyMoments.Add(NewKeyMoment)

                    UndoRedo.ItemPasted(NewKeyMoment, Me.KeyMoments)
            End Select
        Next

        Me.LoadItems()
    End Sub
#End Region
End Class

