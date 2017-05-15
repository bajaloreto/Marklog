Public Class StructuredComboBoxCellNumberUnit
    Inherits DataGridViewComboBoxCell

    Public Sub New()
        Me.Value = ""
    End Sub

    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, _
        ByVal initialFormattedValue As Object, _
        ByVal dataGridViewCellStyle As DataGridViewCellStyle)

        ' Set the value of the editing control to the current cell value.
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, _
            dataGridViewCellStyle)

        Dim ctl As StructuredComboBoxEditingControlNumberUnit = CType(DataGridView.EditingControl, StructuredComboBoxEditingControlNumberUnit)

        Try
            ctl.DropDownHeight = 200
            ctl.DropDownStyle = ComboBoxStyle.DropDown
            ctl.DisplayMember = "Description"
            ctl.ValueMember = "Unit"
        Catch ex As Exception

        End Try

    End Sub

    Public Overrides ReadOnly Property EditType() As Type
        Get
            ' Return the type of the editing contol that 
            ' the DropDownListCell uses.
            Return GetType(StructuredComboBoxEditingControlNumberUnit)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType() As Type
        Get
            Return GetType(String)
        End Get
    End Property

    Protected Overrides Function GetFormattedValue(ByVal value As Object, ByVal rowIndex As Integer, ByRef cellStyle As DataGridViewCellStyle, ByVal valueTypeConverter As System.ComponentModel.TypeConverter, ByVal formattedValueTypeConverter As System.ComponentModel.TypeConverter, ByVal context As DataGridViewDataErrorContexts) As Object

        Return MyBase.GetFormattedValue(value, rowIndex, cellStyle, valueTypeConverter, formattedValueTypeConverter, context)
    End Function

    Public Overrides Function ParseFormattedValue(ByVal formattedValue As Object, ByVal cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal formattedValueTypeConverter As System.ComponentModel.TypeConverter, ByVal valueTypeConverter As System.ComponentModel.TypeConverter) As Object
        Dim strText As String = formattedValue.ToString
        Dim objFormattedValue As Object = Nothing
        Dim boolFound As Boolean

        For Each selItem As StructuredComboBoxItem In Me.Items
            If selItem.IsHeader = False Then
                If selItem.Description = strText Then boolFound = True
            End If
        Next
        If boolFound = False Then
            Dim NewItem As New StructuredComboBoxItem(strText, False, False, , strText)
            Me.Items.Add(NewItem)
            
            Dim ctl As StructuredComboBoxEditingControlNumberUnit = CType(DataGridView.EditingControl, StructuredComboBoxEditingControlNumberUnit)

            ctl.Items.Add(NewItem)
            With Me.DataGridView
                .SuspendLayout()
                Dim selCol As StructuredComboBoxColumnNumberUnit = CType(.Columns(Me.ColumnIndex), StructuredComboBoxColumnNumberUnit)
                selCol.Items.Add(NewItem)
                .ResumeLayout()
            End With
        End If
        'Try
        objFormattedValue = MyBase.ParseFormattedValue(formattedValue, cellStyle, formattedValueTypeConverter, valueTypeConverter)
        'Catch ex As Exception

        'End Try
        Return objFormattedValue
    End Function
End Class

Public Class StructuredComboBoxColumnNumberUnit
    Inherits DataGridViewComboBoxColumn

    Public Sub New()
        'Set the type used in the DataGridView
        MyBase.New()
        Me.CellTemplate = New StructuredComboBoxCellNumberUnit
    End Sub

    Public Overrides Property CellTemplate() As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As DataGridViewCell)

            ' Controleren of de cel die gebruikt wordt een RTF cel is.
            If (value IsNot Nothing) AndAlso _
                Not value.GetType().IsAssignableFrom(GetType(StructuredComboBoxCellNumberUnit)) _
                Then
                Throw New InvalidCastException("Must be a Number Unit ComboBox cell")
            End If
            MyBase.CellTemplate = value

        End Set
    End Property
End Class

Public Class StructuredComboBoxEditingControlNumberUnit
    Inherits ComboBox
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer

    Public Sub New()
        MyBase.New()
        'This will cause the DrawItem event to raise,
        'allowing changes to be made to the appearance
        'at run time
        Me.DrawMode = Windows.Forms.DrawMode.OwnerDrawVariable

        'Make this a DropDownList 
        Me.DropDownStyle = ComboBoxStyle.DropDown
    End Sub

#Region "Properties"
    Private ReadOnly Property DatagridFont() As Font
        Get
            Return Me.EditingControlDataGridView.DefaultCellStyle.Font
        End Get
    End Property

    Public Property EditingControlDataGridView() As System.Windows.Forms.DataGridView Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlDataGridView
        Get
            Return dataGridViewControl
        End Get

        Set(ByVal value As DataGridView)
            dataGridViewControl = value
        End Set
    End Property

    Public Property EditingControlFormattedValue() As Object Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlFormattedValue
        Get
            Return Me.Text
        End Get
        Set(ByVal value As Object)
            Me.Text = value.ToString
        End Set
    End Property

    Public Property EditingControlRowIndex() As Integer Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlRowIndex
        Get
            Return rowIndexNum
        End Get

        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set
    End Property

    Public Property EditingControlValueChanged() As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlValueChanged
        Get
            Return valueIsChanged
        End Get

        Set(ByVal value As Boolean)
            valueIsChanged = value
        End Set
    End Property

    Public ReadOnly Property EditingPanelCursor() As System.Windows.Forms.Cursor Implements System.Windows.Forms.IDataGridViewEditingControl.EditingPanelCursor
        Get
            Return MyBase.Cursor
        End Get
    End Property

    Public ReadOnly Property RepositionEditingControlOnValueChange() As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.RepositionEditingControlOnValueChange
        Get
            Return False
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As System.Windows.Forms.DataGridViewCellStyle) Implements System.Windows.Forms.IDataGridViewEditingControl.ApplyCellStyleToEditingControl

    End Sub

    Public Function EditingControlWantsInputKey(ByVal keyData As System.Windows.Forms.Keys, ByVal dataGridViewWantsInputKey As Boolean) As Boolean Implements System.Windows.Forms.IDataGridViewEditingControl.EditingControlWantsInputKey

    End Function

    Public Function GetEditingControlFormattedValue(ByVal context As System.Windows.Forms.DataGridViewDataErrorContexts) As Object Implements System.Windows.Forms.IDataGridViewEditingControl.GetEditingControlFormattedValue
        'CurrentUndoList.UndoBuffer.NewValue = Me.SelectedValue
        Return Me.Text
    End Function

    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) Implements System.Windows.Forms.IDataGridViewEditingControl.PrepareEditingControlForEdit

    End Sub

    Protected Overrides Sub OnTextChanged(ByVal eventargs As EventArgs)
        ' Notify the DataGridView that the contents of the cell have changed.
        valueIsChanged = True
        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
        MyBase.OnTextChanged(eventargs)

    End Sub

    Protected Overrides Sub OnSelectedValueChanged(ByVal e As System.EventArgs)
        MyBase.OnSelectedValueChanged(e)
        'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeText

    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)

    End Sub

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)
        'With CurrentUndoList.UndoBuffer
        '    If Me.SelectedItem IsNot Nothing Then
        '        .OldValue = CType(Me.SelectedItem, StructuredComboBoxItem).Unit
        '    Else
        '        .OldValue = ""
        '    End If
        '    .IsRtf = False
        'End With
    End Sub
#End Region

#Region "Draw customized dropdown list"
    Protected Overrides Sub OnMeasureItem(ByVal e As System.Windows.Forms.MeasureItemEventArgs)
        Dim SelItem As StructuredComboBoxItem = TryCast(Items(e.Index), StructuredComboBoxItem)
        If SelItem IsNot Nothing Then
            Dim strText As String = SelItem.Description
            Dim boolIsHeader As Boolean = SelItem.IsHeader
            Dim boolWhiteSpace As Boolean = SelItem.WhiteSpace
            Dim fntNormal As New Font(DatagridFont, FontStyle.Regular)
            Dim fntBold As New Font(DatagridFont, FontStyle.Bold)
            Dim intHeight As Integer

            If boolIsHeader = False Then
                intHeight = e.Graphics.MeasureString(strText, fntNormal).Height * 1.1
            Else
                intHeight = e.Graphics.MeasureString(strText, fntBold).Height * 1.1
            End If
            If boolWhiteSpace = True Then intHeight += 4
            e.ItemHeight = intHeight
        End If
        MyBase.OnMeasureItem(e)
    End Sub

    Protected Overrides Sub OnDrawItem(ByVal e As System.Windows.Forms.DrawItemEventArgs)
        If e.Index >= 0 Then
            Dim graph As Graphics = e.Graphics
            Dim SelItem As StructuredComboBoxItem = TryCast(Items(e.Index), StructuredComboBoxItem)
            If SelItem IsNot Nothing Then
                Dim strText As String = SelItem.Description
                Dim boolIsHeader As Boolean = SelItem.IsHeader
                Dim fntNormal As New Font(DatagridFont, FontStyle.Regular)
                Dim fntBold As New Font(DatagridFont, FontStyle.Bold)
                Dim brNormalBack As SolidBrush
                Dim brNormalFront As SolidBrush

                If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
                    brNormalBack = SystemBrushes.Highlight
                    brNormalFront = SystemBrushes.HighlightText
                Else
                    brNormalBack = Brushes.White
                    brNormalFront = Brushes.Black
                End If
                If boolIsHeader = False Then
                    graph.FillRectangle(brNormalBack, e.Bounds)
                    graph.DrawString(strText, fntNormal, brNormalFront, e.Bounds.X + 10, e.Bounds.Y)
                Else
                    graph.FillRectangle(Brushes.Gainsboro, e.Bounds)
                    graph.DrawString(strText, fntBold, Brushes.Black, e.Bounds.X, e.Bounds.Y)
                End If
            End If
        End If
        MyBase.OnDrawItem(e)

    End Sub

    Protected Overrides Sub OnDropDown(ByVal e As System.EventArgs)
        Dim intWidth As Integer = Me.DropDownWidth
        Dim g As Graphics = Me.CreateGraphics()
        Dim fntNormal As New Font(DatagridFont, FontStyle.Regular)
        Dim fntBold As New Font(DatagridFont, FontStyle.Bold)
        Dim font As Font = Me.Font
        Dim intVertScrollBarWidth As Integer = If((Me.Items.Count > Me.MaxDropDownItems), _
                                                  SystemInformation.VerticalScrollBarWidth, 0)
        Dim newWidth As Integer
        Dim strText As String

        For Each selItem As StructuredComboBoxItem In Me.Items
            strText = selItem.Description
            If selItem.IsHeader = False Then
                newWidth = CInt(g.MeasureString(strText, fntNormal).Width)
                newWidth += 10
            Else
                newWidth = CInt(g.MeasureString(strText, fntBold).Width)
            End If
            newWidth += intVertScrollBarWidth

            If intWidth < newWidth Then
                intWidth = newWidth
            End If
        Next
        Me.DropDownWidth = intWidth
        'MyBase.OnDropDown(e)
    End Sub
#End Region
End Class
