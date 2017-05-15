Public Class ListViewBudgetItemReferences
    Inherits ListViewSortable

    Private objCurrentBudgetItem As BudgetItem
    Private objBudgetItemReferences As List(Of BudgetItemReference)

#Region "Properties"
    Public Property CurrentBudgetItem As BudgetItem
        Get
            Return objCurrentBudgetItem
        End Get
        Set(value As BudgetItem)
            objCurrentBudgetItem = value
        End Set
    End Property

    Public Property BudgetItemReferences() As List(Of BudgetItemReference)
        Get
            Return objBudgetItemReferences
        End Get
        Set(ByVal value As List(Of BudgetItemReference))
            objBudgetItemReferences = value
        End Set
    End Property

    Public ReadOnly Property SelectedBudgetItemReferences() As BudgetItemReference()
        Get
            Dim objBudgetItemReferences(SelectedItems.Count - 1) As BudgetItemReference

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objBudgetItemReferences(i) = GetBudgetItemReferenceByGuid(Me.SelectedGuids(i))
                Next

                Return objBudgetItemReferences
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New(ByVal budgetitem As BudgetItem)
        View = View.Details
        FullRowSelect = True
        OwnerDraw = True
        Columns.Add(CurrentLogFrame.StructNamePlural(3), 150, HorizontalAlignment.Left) 'Activity
        Columns.Add(LANG_Resource, 150, HorizontalAlignment.Left) 'Resource
        Columns.Add(LANG_Percentage, 150, HorizontalAlignment.Right)
        Columns.Add(LANG_Amount, 150, HorizontalAlignment.Right)

        Me.CurrentBudgetItem = budgetitem

        LoadItems()
    End Sub

    Public Sub LoadItems()
        Dim sngTotalCost As Single
        Dim sngTotalPercentage As Single
        Dim strCurrency As String = String.Empty

        Me.Items.Clear()
        If Me.CurrentBudgetItem IsNot Nothing Then
            Me.BudgetItemReferences = CurrentLogFrame.GetReferingResourcesByReferenceGuid(CurrentBudgetItem.Guid)

            For Each selBudgetItemReference As BudgetItemReference In Me.BudgetItemReferences
                If BudgetItemReferences.IndexOf(selBudgetItemReference) = 0 AndAlso selBudgetItemReference.TotalCost IsNot Nothing Then _
                    strCurrency = selBudgetItemReference.TotalCost.CurrencyCode

                Dim ParentActivity As Activity
                Dim ParentResource As Resource = CurrentLogFrame.GetResourceByGuid(selBudgetItemReference.ParentResourceGuid)
                Dim strActivityText As String = String.Empty
                Dim strResourceText As String = String.Empty

                If ParentResource IsNot Nothing Then
                    strResourceText = ParentResource.Text
                    ParentActivity = CurrentLogFrame.GetActivityByGuid(ParentResource.ParentStructGuid)

                    If ParentActivity IsNot Nothing Then strActivityText = ParentActivity.Text
                End If

                Dim newItem As New ListViewItem(strActivityText)
                With newItem
                    .Name = selBudgetItemReference.Guid.ToString
                    .SubItems.Add(strResourceText)
                    .SubItems.Add(selBudgetItemReference.Percentage.ToString("P2"))
                    .SubItems.Add(selBudgetItemReference.TotalCost.ToString)
                End With

                Me.Items.Add(newItem)
                sngTotalPercentage += selBudgetItemReference.Percentage
                sngTotalCost += selBudgetItemReference.TotalCost.Amount
            Next

            If BudgetItemReferences.Count > 0 Then
                Dim TotalItem As New ListViewItem(LANG_Total)
                Dim curTotal As New Currency(sngTotalCost, strCurrency)
                With TotalItem
                    'General information
                    .Name = "Total"
                    .SubItems.Add(String.Empty)
                    .SubItems.Add(sngTotalPercentage.ToString("P2"))
                    .SubItems.Add(curTotal.ToString)
                End With

                Me.Items.Add(TotalItem)
                AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent)
            End If
        End If
    End Sub
#End Region

#Region "Methods & Events"
    Private Function GetBudgetItemReferenceByGuid(ByVal objGuid As Guid)
        For Each selBudgetItemReference As BudgetItemReference In Me.BudgetItemReferences
            If selBudgetItemReference.Guid = objGuid Then Return selBudgetItemReference
        Next

        Return Nothing
    End Function

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        PopUpBudgetItemReferenceDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.BudgetItemReferences.Count > 0 AndAlso Me.SelectedBudgetItemReferences.Length > 0 Then
            PopUpBudgetItemReferenceDetails(Me.SelectedBudgetItemReferences(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.BudgetItemReferences.Count > 0 AndAlso Me.SelectedBudgetItemReferences.Length > 0 Then
            If MsgBox(LANG_RemoveBudgetItemReference, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selBudgetItemReference As BudgetItemReference = Me.SelectedBudgetItemReferences(0)

                Dim objParentResource As Resource = CurrentLogFrame.GetResourceByGuid(selBudgetItemReference.ParentResourceGuid)

                If objParentResource IsNot Nothing Then
                    UndoRedo.ItemRemoved(selBudgetItemReference, objParentResource.BudgetItemReferences)

                    objParentResource.BudgetItemReferences.Remove(selBudgetItemReference)
                    objParentResource.SetTotalCostAmount()
                End If

                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpBudgetItemReferenceDetails(ByVal selBudgetItemReference As BudgetItemReference)
        Dim boolNew As Boolean

        If selBudgetItemReference Is Nothing Then
            boolNew = True
            With CurrentBudgetItem
                selBudgetItemReference = New BudgetItemReference(.Guid)
                selBudgetItemReference.TotalCost = New Currency(.TotalCost.Amount, .TotalCost.CurrencyCode)
            End With
        End If

        Dim dialogBudgetItemReference As New DialogBudgetItemReference(selBudgetItemReference)

        dialogBudgetItemReference.ShowDialog()
        If dialogBudgetItemReference.DialogResult = vbOK Then
            If selBudgetItemReference.ParentResourceGuid = Guid.Empty Then Exit Sub
            Dim objParentResource As Resource = CurrentLogFrame.GetResourceByGuid(selBudgetItemReference.ParentResourceGuid)

            If boolNew = True Then
                objParentResource.BudgetItemReferences.Add(selBudgetItemReference)

                UndoRedo.ItemInserted(selBudgetItemReference, objParentResource.BudgetItemReferences)
            End If

            objParentResource.SetTotalCostAmount()
            Me.LoadItems()
        End If
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selBudgetItemReference As BudgetItemReference In SelectedBudgetItemReferences
            UndoRedo.ItemCut(selBudgetItemReference, Me.BudgetItemReferences)

            BudgetItemReferences.Remove(selBudgetItemReference)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selBudgetItemReference As BudgetItemReference In SelectedBudgetItemReferences
            Dim NewItem As New ClipboardItem(CopyGroup, selBudgetItemReference, BudgetItemReference.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selBudgetItemReference As BudgetItemReference

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(BudgetItemReference)
                    selBudgetItemReference = CType(selItem.Item, BudgetItemReference)
                    Dim NewBudgetItemReference As New BudgetItemReference

                    Using copier As New ObjectCopy
                        NewBudgetItemReference = copier.CopyObject(selBudgetItemReference)
                    End Using

                    Me.BudgetItemReferences.Add(NewBudgetItemReference)

                    UndoRedo.ItemPasted(NewBudgetItemReference, Me.BudgetItemReferences)
            End Select
        Next

        Me.LoadItems()
    End Sub
#End Region

#Region "Draw"
    Protected Overrides Sub OnDrawColumnHeader(e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs)
        MyBase.OnDrawColumnHeader(e)

        e.DrawDefault = True
    End Sub

    Protected Overrides Sub OnDrawSubItem(e As System.Windows.Forms.DrawListViewSubItemEventArgs)
        Dim graphics As Graphics = e.Graphics
        Dim sfSubItem As New StringFormat
        Dim brText As New SolidBrush(Color.Black)
        Dim rCell As Rectangle = e.Bounds
        rCell.X += 4
        rCell.Y += 2
        rCell.Width -= 8
        rCell.Height -= 4

        sfSubItem.LineAlignment = StringAlignment.Center
        If e.ColumnIndex = 0 Then
            sfSubItem.Alignment = StringAlignment.Near
        Else
            sfSubItem.Alignment = StringAlignment.Far
        End If

        If e.Item.Name = "Total" Then
            If e.ColumnIndex = 2 Then
                Dim strValue As String = e.SubItem.Text
                strValue = strValue.Split("%")(0)
                Dim sngValue As Single = ParseSingle(strValue)
                If sngValue > 100 Then brText.Color = Color.Red
            End If

            graphics.FillRectangle(SystemBrushes.ControlDark, e.Bounds)
            graphics.DrawString(e.SubItem.Text, Me.Font, brText, rCell, sfSubItem)
        Else
            e.DrawDefault = True
        End If
    End Sub
#End Region
End Class

