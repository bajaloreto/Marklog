Class RichTextEditingControlLogframe
    Inherits RichTextBox
    Implements IDataGridViewEditingControl

    Private dataGridViewControl As DataGridView
    Private valueIsChanged As Boolean = False
    Private rowIndexNum As Integer
    Public BeforeFirstStroke As Boolean

    Public Enum TextCases
        Normal = 0
        Uppercase = 1
        Lowercase = 2
        Sentence = 3
        FirstLetter = 4
    End Enum

    Public Sub New()

        Me.BorderStyle = Windows.Forms.BorderStyle.None
        Me.MaxLength = 400
        Me.Multiline = True
        Me.WordWrap = True
    End Sub

#Region "Properties"
    Public Property EditingControlFormattedValue() As Object _
        Implements IDataGridViewEditingControl.EditingControlFormattedValue

        Get
            Return Me.Rtf
        End Get

        Set(ByVal value As Object)
            If TypeOf value Is String Then
                Me.Rtf = value
            End If
        End Set
    End Property

    Public Property EditingControlRowIndex() As Integer Implements IDataGridViewEditingControl.EditingControlRowIndex
        Get
            Return rowIndexNum
        End Get

        Set(ByVal value As Integer)
            rowIndexNum = value
        End Set
    End Property

    Public ReadOnly Property RepositionEditingControlOnValueChange() As Boolean Implements _
        IDataGridViewEditingControl.RepositionEditingControlOnValueChange

        Get
            Return False
        End Get
    End Property

    Public Property EditingControlDataGridView() As DataGridView _
        Implements IDataGridViewEditingControl.EditingControlDataGridView

        Get
            Return dataGridViewControl
        End Get

        Set(ByVal value As DataGridView)
            dataGridViewControl = value
        End Set

    End Property

    Public Property EditingControlValueChanged() As Boolean Implements IDataGridViewEditingControl.EditingControlValueChanged
        Get
            Return valueIsChanged
        End Get

        Set(ByVal value As Boolean)
            valueIsChanged = value
        End Set
    End Property

    Public ReadOnly Property EditingControlCursor() As Cursor Implements IDataGridViewEditingControl.EditingPanelCursor
        Get
            Return MyBase.Cursor
        End Get
    End Property
#End Region 'Properties

#Region "Methods"
    Public Function GetEditingControlFormattedValue(ByVal context As DataGridViewDataErrorContexts) As Object _
        Implements IDataGridViewEditingControl.GetEditingControlFormattedValue

        'CurrentUndoList.UndoBuffer.NewValue = Me.Rtf
        Return Me.Rtf
    End Function

    Public Sub ApplyCellStyleToEditingControl(ByVal dataGridViewCellStyle As DataGridViewCellStyle) _
        Implements IDataGridViewEditingControl.ApplyCellStyleToEditingControl

    End Sub

    Public Sub PrepareEditingControlForEdit(ByVal selectAll As Boolean) _
        Implements IDataGridViewEditingControl.PrepareEditingControlForEdit

        ' No preparation needs to be done.
    End Sub

    Public Function EditingControlWantsInputKey(ByVal key As Keys, _
        ByVal dataGridViewWantsInputKey As Boolean) As Boolean _
        Implements IDataGridViewEditingControl.EditingControlWantsInputKey

        ' Let the RTF-cell handle the keys listed.
        Dim boolInputKey As Boolean
        Dim intTextLength As Integer = Me.TextLength

        Select Case key And Keys.KeyCode
            Case Keys.Left, Keys.Up, Keys.Down, Keys.Right, Keys.Home, Keys.End, Keys.PageDown, Keys.PageUp
                boolInputKey = True

            Case Keys.Delete, Keys.Back
                boolInputKey = False

                'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeText
            Case Keys.OemPeriod, Keys.D1, Keys.D3, Keys.D4, Keys.D5, Keys.D8, Keys.Oem1, Keys.Q, Keys.Oemtilde
                boolInputKey = True

                'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeText
            Case Keys.ShiftKey, Keys.ControlKey
                boolInputKey = False

            Case Else
                boolInputKey = False

                'CurrentUndoList.UndoBuffer.ActionIndex = UndoListItemOld.Actions.ChangeText

        End Select
        Return boolInputKey
    End Function

    Protected Overrides Sub OnTextChanged(ByVal eventargs As EventArgs)
        EditingControlValueChanged = True
        Me.EditingControlDataGridView.NotifyCurrentCellDirty(True)
        MyBase.OnTextChanged(eventargs)

    End Sub

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        If Me.Text = String.Empty Then
            SelectionFont = CurrentText.Font
        End If
        'With CurrentUndoList.UndoBuffer
        '    .OldValue = Me.Rtf
        '    .IsRtf = True
        'End With
    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)

        'cut/copy/paste is handled at level of frmParent
        If e.Control AndAlso (e.KeyCode = Keys.X Or e.KeyCode = Keys.C Or e.KeyCode = Keys.V) Then
            e.Handled = True
        End If
    End Sub
#End Region 'Methods

#Region "Selection"
    Protected Overrides Sub OnMouseUp(ByVal mevent As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseUp(mevent)
        Me.SelectionLength = RTrim(Me.SelectedText).Length
        SetCurrentText()
    End Sub
#End Region 'Selection

#Region "Font and paragraph"
    Public Sub ChangeFont()

        With CurrentText
            If .Font Is Nothing Then Exit Sub

            Dim intSelectionStart As Integer = SelectionStart
            Dim intSelectionLength As Integer = SelectionLength
            Dim objFontStyle As New FontStyle
            objFontStyle = FontStyle.Regular

            If .BoldComplex = False And .Font.Bold = True Then objFontStyle = FontStyle.Bold
            If .ItalicComplex = False And .Font.Italic = True Then objFontStyle += FontStyle.Italic
            If .UnderlineComplex = False And .Font.Underline = True Then objFontStyle += FontStyle.Underline
            If .StrikeoutComplex = False And .Font.Strikeout = True Then objFontStyle += FontStyle.Strikeout

            Me.Select(.SelectionStart, .SelectionLength)

            If .FontComplex = False Then
                If .FontSizeComplex = False Then
                    ChangeFont_Change(.Font.Name, .Font.Size, objFontStyle)
                Else
                    ChangeFont_ByCharacter(.Font.Name, .Font.Size, objFontStyle)
                End If
            Else
                ChangeFont_ByCharacter(.Font.Name, .Font.Size, objFontStyle)
            End If

            SetCurrentText()
            Me.Select(intSelectionStart, intSelectionLength)
        End With
    End Sub

    Public Sub ChangeFont_ByCharacter(ByVal strFontFamily As String, ByVal intFontSize As Integer, ByVal objFontStyle As FontStyle)
        With CurrentText
            Dim intSelStart As Integer, intSelStop As Integer
            Dim i As Integer

            intSelStart = .SelectionStart
            intSelStop = intSelStart + .SelectionLength

            For i = intSelStart To intSelStop - 1
                Me.Select(i, 1)
                If .FontComplex Then strFontFamily = Me.SelectionFont.Name

                If .FontSizeComplex Then
                    intFontSize = Me.SelectionFont.SizeInPoints
                Else
                    If .CharOffSet <> 0 Then intFontSize *= 0.6
                End If

                ChangeFont_Change(strFontFamily, intFontSize, objFontStyle)
            Next
        End With
    End Sub

    Public Sub ChangeFont_Change(ByVal strFontFamily As String, ByVal intFontSize As Integer, ByVal objFontStyle As FontStyle)
        Try
            Me.SelectionFont = New Font(strFontFamily, intFontSize, objFontStyle, GraphicsUnit.Point)
        Catch ex2 As ArgumentException
            Try
                Me.SelectionFont = New Font(strFontFamily, intFontSize, Me.SelectionFont.Style, GraphicsUnit.Point)
            Catch ex3 As ArgumentException
                MsgBox(ERR_FontStyleNotSupported, MsgBoxStyle.OkOnly)
                Me.Select(CurrentText.SelectionStart, CurrentText.SelectionLength)

                Exit Sub
            End Try
        End Try
    End Sub

    Public Sub ChangeFontCase(ByVal intCase As Integer)
        Dim intSelectionStart As Integer = SelectionStart
        Dim intSelectionLength As Integer = SelectionLength

        Select Case intCase
            Case TextCases.Uppercase
                Me.SelectedText = Me.SelectedText.ToUpper
            Case TextCases.Lowercase
                Me.SelectedText = Me.SelectedText.ToLower
            Case TextCases.FirstLetter
                Me.SelectedText = StrConv(Me.SelectedText, VbStrConv.ProperCase)
            Case TextCases.Sentence
                Dim sentenceStart As Boolean = True
                Dim strSentence() As Char = CStr(Me.SelectedText).ToLower
                For i = 0 To Len(strSentence) - 1
                    If Char.IsLetterOrDigit(strSentence(i)) Then
                        If sentenceStart Then
                            strSentence(i) = CStr(strSentence(i)).ToUpper
                            sentenceStart = False
                        End If
                    ElseIf InStr(".!?" & vbCrLf, strSentence(i)) Then
                        sentenceStart = True
                    End If
                Next
                Me.SelectedText = strSentence
        End Select

        Me.Select(intSelectionStart, intSelectionLength)
    End Sub

    Public Sub ChangeFont_CharOffset(ByVal intUpDown As Integer)
        Dim intSize As Integer
        Dim intStart As Integer, intLen As Integer
        Dim intSelectionStart As Integer = SelectionStart
        Dim intSelectionLength As Integer = SelectionLength

        With CurrentText
            If .Font Is Nothing Then Exit Sub
            If Me.SelectionCharOffset = 0 Then
                Me.SelectionCharOffset = .CharOffSet
                ChangeFont_Change(.Font.Name, .Font.Size, .Font.Style)
            Else
                'get size of previous character and set superscript/subscript character to that size
                intStart = Me.SelectionStart
                intLen = Me.SelectionLength
                With Me
                    If intStart > 0 Then
                        .SelectionStart = intStart - 1
                        .SelectionLength = 0
                    End If
                    intSize = Me.SelectionFont.SizeInPoints
                    .SelectionStart = intStart
                    .SelectionLength = intLen
                    .SelectionCharOffset = 0
                End With
                Me.SelectionFont = New Font(Me.SelectionFont.Name, intSize, Me.SelectionFont.Style, GraphicsUnit.Point)
            End If

            SetCurrentText()
            Me.Select(intSelectionStart, intSelectionLength)
        End With
    End Sub

    Public Sub SetCurrentText()
        With CurrentText
            .Reset()
            .SelectionStart = Me.SelectionStart
            .SelectionLength = Me.SelectionLength

            .HorizontalAlignment = Me.SelectionAlignment

            If Me.SelectionLength > 1 Then
                Dim intSelStart As Integer = Me.SelectionStart
                Dim intSelLength As Integer = Me.SelectionLength
                Dim intSelStop As Integer = intSelStart + intSelLength

                Me.Select(intSelStart, 1)
                .Font = SelectionFont
                .ChangeFontSize(SelectionFont.Size) 'make sure font size is an integer
                .CharOffSet = Me.SelectionCharOffset

                For i = intSelStart To intSelStop - 1
                    Me.Select(i, 1)
                    If SelectionFont.Name <> .Font.Name Then .FontComplex = True
                    If Int(SelectionFont.Size) <> .Font.Size Then .FontSizeComplex = True
                    If SelectionFont.Bold <> .Font.Bold Then .BoldComplex = True
                    If SelectionFont.Italic <> .Font.Italic Then .ItalicComplex = True
                    If SelectionFont.Underline <> .Font.Underline Then .UnderlineComplex = True
                    If SelectionFont.Strikeout <> .Font.Strikeout Then .StrikeoutComplex = True
                    If SelectionCharOffset <> .CharOffSet Then .CharOffSetComplex = True
                Next
                Me.Select(intSelStart, intSelLength)
            Else
                .Font = Me.SelectionFont
                .ChangeFontSize(SelectionFont.Size) 'make sure font size is an integer
                .CharOffSet = Me.SelectionCharOffset
            End If
        End With
    End Sub

    Public Sub ChangeTextColor(ByVal color As Color)
        Me.SelectionColor = color
    End Sub

    Public Sub ChangeMarkerColor(ByVal color As Color)
        Me.SelectionBackColor = color
    End Sub

    Public Sub ChangeTextAlignment(ByVal intAlignment As Integer)
        Me.SelectionAlignment = intAlignment

        SetCurrentText()
    End Sub

    Public Sub ChangeLeftIndent(ByVal intIncrement As Integer)
        Me.SelectionIndent += intIncrement
    End Sub
#End Region 'Font and paragraph



End Class
