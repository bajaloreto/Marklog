Public Class DataGridViewResponseClasses
    Inherits DataGridViewBaseClass

    Private objResponseClasses As ResponseClasses
    Private colClassName As New DataGridViewTextBoxColumn
    Private colValue As New DataGridViewTextBoxColumn
    Private boolShowValuesColumn As Boolean

    Public Event DataSourceUpdated()

#Region "Properties"
    Public Property ResponseClasses As ResponseClasses
        Get
            Return objResponseClasses
        End Get
        Set(ByVal value As ResponseClasses)
            objResponseClasses = value
        End Set
    End Property

    Public Overrides ReadOnly Property CurrentItem(OnlyIfTextShows As Boolean) As Object
        Get
            If CurrentRow IsNot Nothing And CurrentRow.Index <> NewRowIndex Then
                Return ResponseClasses(CurrentRow.Index)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Property ShowValuesColumn As Boolean
        Get
            Return boolShowValuesColumn
        End Get
        Set(value As Boolean)
            boolShowValuesColumn = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub New()

        ChooseSettings()
    End Sub

    Public Sub New(ByVal responseclasses As ResponseClasses, ByVal showvaluescolumn As Boolean)
        'load values
        Me.ResponseClasses = responseclasses
        Me.ShowValuesColumn = showvaluescolumn

        'datagridview settings
        ChooseSettings()

        LoadColumns()
    End Sub

    Protected Sub ChooseSettings()
        VirtualMode = False
        EditMode = DataGridViewEditMode.EditOnEnter
        AutoGenerateColumns = False
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = False

        ShowCellToolTips = False
        BackgroundColor = Color.White
        DefaultCellStyle.Padding = New Padding(2)
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        RowHeadersVisible = True
        RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        With RowHeadersDefaultCellStyle
            .WrapMode = DataGridViewTriState.True
        End With

        With ColumnHeadersDefaultCellStyle
            .Font = New Font(DefaultFont, FontStyle.Bold)
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .WrapMode = DataGridViewTriState.True
        End With
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize

    End Sub

    Public Overridable Sub LoadColumns()
        Columns.Clear()

        With colClassName
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .HeaderText = LANG_Option
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .DataPropertyName = "ClassName"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End With
        Me.Columns.Add(colClassName)

        If ShowValuesColumn = True Then
            With colValue
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                .HeaderText = LANG_Value
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .DataPropertyName = "Value"
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            End With

            Me.Columns.Add(colValue)
        End If
    End Sub
#End Region

#Region "General methods"
    Public Sub SetFocusOnItem(ByVal selResponseClass As ResponseClass, ByVal boolSelectValue As Boolean)
        Dim intColIndex, intRowIndex As Integer

        If boolSelectValue = True Then intColIndex = 1
        intRowIndex = ResponseClasses.IndexOf(selResponseClass)

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            CurrentCell = Me(intColIndex, intRowIndex)
        End If
    End Sub

    Public Sub AddNewItem()
        Dim intColIndex As Integer
        If CurrentCell IsNot Nothing Then
            intColIndex = CurrentCell.ColumnIndex
        End If

        CurrentCell = Me(intColIndex, RowCount - 1)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selObject As Object = Me.CurrentItem(False)
        Dim intIndex As Integer

        If CurrentCell IsNot Nothing Then
            intIndex = CurrentCell.RowIndex

            Dim NewResponseClass As New ResponseClass
            ResponseClasses.Insert(intIndex, NewResponseClass)

            UndoRedo.ItemInserted(NewResponseClass, ResponseClasses)

            RaiseEvent DataSourceUpdated()
        End If

        Me.BeginEdit(False)
    End Sub

    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selItem As Object = Me.CurrentItem(False)
        If selItem Is Nothing Then Exit Sub

        Dim intIndex As Integer = CurrentCell.RowIndex
        Dim intOldIndex As Integer = intIndex
        Dim selResponseClass As ResponseClass = ResponseClasses(intIndex)

        If intDirection < 0 Then
            intIndex -= 1
            If intIndex < 0 Then intIndex = 0
        Else
            intIndex += 1
            If intIndex > ResponseClasses.Count - 1 Then intIndex = ResponseClasses.Count - 1
        End If

        ResponseClasses.Remove(selResponseClass)
        ResponseClasses.Insert(intIndex, selResponseClass)

        If intDirection < 0 Then
            UndoRedo.ItemMovedUp(selResponseClass, selResponseClass, ResponseClasses, intOldIndex, ResponseClasses)
        Else
            UndoRedo.ItemMovedDown(selResponseClass, selResponseClass, ResponseClasses, intOldIndex, ResponseClasses)
        End If

        RaiseEvent DataSourceUpdated()

        CurrentCell = Me(0, intIndex)
    End Sub

    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        'If Me.IsCurrentCellInEditMode = False Then
        If CurrentRow IsNot Nothing AndAlso CurrentRow.Index <> NewRowIndex Then
            Dim selResponseClass As ResponseClass = ResponseClasses(CurrentRow.Index)

            UndoRedo.ItemRemoved(selResponseClass, Me.ResponseClasses)

            Rows.RemoveAt(CurrentRow.Index)
        End If
        'Else
        'With CurrentEditingControl
        '    Dim intSelectionStart As Integer = .SelectionStart
        '    Dim intLength As Integer = .SelectionLength

        '    If intLength = 0 Then intLength = 1
        '    .Text = .Text.Remove(intSelectionStart, intLength)

        '    .Select(intSelectionStart, 0)
        'End With
        'End If
    End Sub
#End Region

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)
        CopyItems()

        RemoveItems(False)
    End Sub

    Public Overrides Sub CopyItems()
        If CurrentRow IsNot Nothing AndAlso CurrentRow.Index <> NewRowIndex Then
            Dim selResponseClass As ResponseClass = ResponseClasses(CurrentRow.Index)
            Dim CopyGroup As Date = Now()
            Dim strSort As String

            If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selResponseClass Then
                strSort = ResponseClass.ItemName
                Dim NewItem As New ClipboardItem(CopyGroup, selResponseClass, strSort)
                ItemClipboard.Insert(0, NewItem)
            End If
        End If
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal PasteOption As Integer, Optional ByVal PasteCell As System.Windows.Forms.DataGridViewCell = Nothing)
        If PasteCell Is Nothing Then PasteCell = CurrentCell
        If PasteCell Is Nothing Then Exit Sub

        Dim intColumnIndex As Integer = PasteCell.ColumnIndex
        Dim intRowIndex As Integer = PasteCell.RowIndex
        Dim selItem As ClipboardItem

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(ResponseClass)
                    PasteItems_ResponseClass(selItem, PasteOption)
                Case Else
                    PasteItems_Other(selItem, PasteOption)
            End Select
        Next
    End Sub

    Private Sub PasteItems_ResponseClass(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selResponseClass As ResponseClass = DirectCast(selItem.Item, ResponseClass)
        Dim NewResponseClass As New ResponseClass
        Dim intIndex = CurrentRow.Index

        Using copier As New ObjectCopy
            NewResponseClass = copier.CopyObject(selResponseClass)
        End Using

        PasteResponseClass(NewResponseClass, intIndex)
    End Sub

    Private Sub PasteItems_Other(ByVal selItem As ClipboardItem, ByVal intPasteOption As Integer)
        Dim selObject As LogframeObject = TryCast(selItem.Item, LogframeObject)
        Dim strText As String = String.Empty
        Dim intIndex = CurrentRow.Index

        If selObject IsNot Nothing Then
            strText = selObject.Text
        Else
            strText = selItem.ToString
        End If

        If String.IsNullOrEmpty(strText) = False Then
            Dim NewResponseClass As New ResponseClass(strText)

            PasteResponseClass(NewResponseClass, intIndex)
        End If
    End Sub

    Private Sub PasteResponseClass(ByVal NewResponseClass As ResponseClass, ByVal intIndex As Integer)
        If intIndex = -1 Then intIndex = ResponseClasses.Count

        ResponseClasses.Insert(intIndex, NewResponseClass)

        UndoRedo.ItemPasted(NewResponseClass, Me.ResponseClasses)

        RaiseEvent DataSourceUpdated()
    End Sub
#End Region

#Region "Events"
    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)
        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnCellValidating(ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)
        MyBase.OnCellValidating(e)

        Dim selResponseClass As ResponseClass = ResponseClasses(e.RowIndex)

        If selResponseClass IsNot Nothing Then
            Select Case e.ColumnIndex
                Case 0
                    UndoRedo.UndoBuffer_Initialise(selResponseClass, "ClassName", selResponseClass.ClassName)
                Case 1
                    UndoRedo.UndoBuffer_Initialise(selResponseClass, "Value", selResponseClass.Value)
            End Select
        End If
    End Sub

    Protected Overrides Sub OnCellValidated(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellValidated(e)

        If CurrentItem(False) IsNot Nothing Then
            Dim selResponseClass As ResponseClass = CurrentItem(False)

            Select Case e.ColumnIndex
                Case 0
                    UndoRedo.TextChanged(Me(e.ColumnIndex, e.RowIndex).Value)
                Case 1
                    UndoRedo.ValueChanged(Me(e.ColumnIndex, e.RowIndex).Value)
            End Select

        End If
    End Sub
#End Region
End Class
