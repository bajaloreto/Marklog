Public Class ListViewSubIndicatorsBase
    Inherits ListViewSortable

    Protected objParentIndicator As Indicator

    Public Event Updated()
    Public Event IndicatorModified()


#Region "Properties"
    Public Property ParentIndicator() As Indicator
        Get
            Return objParentIndicator
        End Get
        Set(ByVal value As Indicator)
            objParentIndicator = value

        End Set
    End Property

    Public ReadOnly Property SelectedIndicators() As Indicator()
        Get
            Dim objIndicators(SelectedItems.Count - 1) As Indicator
            Dim intIndex As Integer

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    If SelectedGuids(i) <> Guid.Empty Then
                        objIndicators(intIndex) = Me.ParentIndicator.Indicators.GetIndicatorByGuid(Me.SelectedGuids(intIndex))
                        intIndex += 1
                    End If

                Next
                ReDim Preserve objIndicators(intIndex)
                Return objIndicators
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Initialise"
    Public Sub New()
        View = View.Details
        FullRowSelect = True
    End Sub

    Public Sub New(ByVal indicator As Indicator)
        View = View.Details
        FullRowSelect = True

        Me.ParentIndicator = indicator

        LoadColumns()
    End Sub

    Public Sub LoadColumns()

        Columns.Clear()

        If ParentIndicator IsNot Nothing Then
            Columns.Add(LANG_Indicator, 300, HorizontalAlignment.Left)
            Columns.Add(LANG_ResponseType, 150, HorizontalAlignment.Left)
            Columns.Add(LANG_TargetGroup, 120, HorizontalAlignment.Left)
            Columns.Add(LANG_RegistrationOption, 120, HorizontalAlignment.Left)

        End If
    End Sub

    Public Overridable Sub LoadItems()
        Dim objTargetScoresTotal As Targets = Nothing

        Me.Items.Clear()
        If ParentIndicator IsNot Nothing Then
            LoadItems_SubIndicators()

            If ParentIndicator.Indicators.Count > 0 Then
                For i = 0 To 2
                    Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                Next
            End If

        End If
    End Sub

    Protected Overridable Sub LoadItems_SubIndicators()

        For Each selIndicator As Indicator In Me.ParentIndicator.Indicators
            Dim newItem As New ListViewItem(selIndicator.Text)
            Dim selTargetGroup As TargetGroup = CurrentLogFrame.Purposes.GetTargetGroupByGuid(selIndicator.TargetGroupGuid)
            Dim strTargetGroupName As String = String.Empty

            'values: set number of decimals and unit
            If selTargetGroup IsNot Nothing Then strTargetGroupName = selTargetGroup.Name

            With newItem
                'General information
                .Name = selIndicator.Guid.ToString
                .SubItems.Add(selIndicator.QuestionTypeName)
                .SubItems.Add(strTargetGroupName)
                .SubItems.Add(selIndicator.RegistrationOption)

            End With
            Me.Items.Add(newItem)
        Next
    End Sub

    
#End Region

#Region "Methods & Events"
    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        PopUpIndicatorDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.SelectedIndicators.Length > 0 AndAlso Me.SelectedIndicators(0) IsNot Nothing Then
            PopUpIndicatorDetails(Me.SelectedIndicators(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.SelectedIndicators.Length > 0 Then
            If MsgBox(LANG_RemoveIndicator, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selIndicator As Indicator = Me.SelectedIndicators(0)

                UndoRedo.ItemRemoved(selIndicator, Me.ParentIndicator.Indicators)

                Me.ParentIndicator.Indicators.Remove(selIndicator)
                Me.LoadItems()

                RaiseEvent IndicatorModified()
            End If
        End If
    End Sub

    Private Sub PopUpIndicatorDetails(ByVal selIndicator As Indicator)
        Dim boolNew As Boolean

        If selIndicator Is Nothing Then
            boolNew = True
            selIndicator = New Indicator()

            Me.ParentIndicator.Indicators.Add(selIndicator)

            UndoRedo.ItemInserted(selIndicator, Me.ParentIndicator.Indicators)
        End If

        Dim dialogIndicator As New DialogIndicator(selIndicator)
        Dim OldParentIndicatorGuid As Guid = selIndicator.ParentIndicatorGuid

        If dialogIndicator.ShowDialog() = DialogResult.OK Then
            Me.LoadItems()
            RaiseEvent IndicatorModified()
        End If
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selIndicator As Indicator In SelectedIndicators
            UndoRedo.ItemCut(selIndicator, Me.ParentIndicator.Indicators)

            Me.ParentIndicator.Indicators.Remove(selIndicator)
        Next

        LoadItems()
        RaiseEvent IndicatorModified()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selIndicator As Indicator In SelectedIndicators
            Dim NewItem As New ClipboardItem(CopyGroup, selIndicator, Indicator.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selIndicator As Indicator

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Indicator)
                    selIndicator = CType(selItem.Item, Indicator)
                    Dim NewIndicator As New Indicator

                    Using copier As New ObjectCopy
                        NewIndicator = copier.CopyObject(selIndicator)
                    End Using

                    Me.ParentIndicator.Indicators.Add(NewIndicator)

                    UndoRedo.ItemPasted(NewIndicator, Me.ParentIndicator.Indicators)
            End Select
        Next

        Me.LoadItems()
        RaiseEvent IndicatorModified()
    End Sub
#End Region


End Class

