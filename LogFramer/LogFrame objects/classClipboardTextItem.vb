Public Class ClipboardTextItem
    Private strText As String
    Private strRtf As String

    Public Property Text As String
        Get
            Return strText
        End Get
        Set(value As String)
            strText = value
        End Set
    End Property

    Public Property RTF As String
        Get
            Return strRtf
        End Get
        Set(value As String)
            strRtf = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal strText As String)
        Me.Text = strText
    End Sub

    Public Sub New(ByVal strText As String, strRtf As String)
        Me.Text = strText
        Me.RTF = strRtf
    End Sub
End Class

Public Class ClipboardTextItems
    Inherits System.Collections.CollectionBase

    Public Event ListUpdated()

    Public Enum PasteOptions As Integer
        KeepFormatting = 0
        MergeFormatting = 1
        TextOnly = 2
    End Enum

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As ClipboardTextItem
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ClipboardTextItem)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal clipboarditem As ClipboardTextItem)
        If clipboarditem IsNot Nothing Then
            List.Add(clipboarditem)
            RaiseEvent ListUpdated()
        End If
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal clipboarditem As ClipboardTextItem)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, LANG_Text.ToLower))
        ElseIf index = Count Then
            List.Add(clipboarditem)
            RaiseEvent ListUpdated()
        Else
            List.Insert(index, clipboarditem)
            RaiseEvent ListUpdated()
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, LANG_Text.ToLower))
        Else
            List.RemoveAt(index)
            RaiseEvent ListUpdated()
        End If
    End Sub

    Public Sub Remove(ByVal clipboarditem As ClipboardTextItem)
        If List.Contains(clipboarditem) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, LANG_Text.ToLower))
        Else
            List.Remove(clipboarditem)
            RaiseEvent ListUpdated()
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
        RaiseEvent ListUpdated()
    End Sub

    Public Function IndexOf(ByVal clipboarditem As ClipboardTextItem) As Integer
        Return List.IndexOf(clipboarditem)
    End Function

    Public Function Contains(ByVal clipboarditem As ClipboardTextItem) As Boolean
        Return List.Contains(clipboarditem)
    End Function
#End Region

#Region "Cut, copy and paste"
    Public Sub CutRichText(ByVal strText As String, ByVal strRtf As String)
        Me.Insert(0, New ClipboardTextItem(strText, strRtf))
    End Sub

    Public Sub CutRichText(ByVal strRtf As String)
        Dim strText As String

        Using objRichTextManager As New RichTextManager
            objRichTextManager.Rtf = strRtf
            strText = objRichTextManager.Text
        End Using

        Me.Insert(0, New ClipboardTextItem(strText, strRtf))
    End Sub

    Public Sub CutText(ByVal strText As String)
        Me.Insert(0, New ClipboardTextItem(strText))
    End Sub

    Public Sub CopyRichText(ByVal strText As String, ByVal strRtf As String)
        Me.Insert(0, New ClipboardTextItem(strText, strRtf))
    End Sub

    Public Sub CopyRichText(ByVal strRtf As String)
        Dim strText As String

        Using objRichTextManager As New RichTextManager
            objRichTextManager.Rtf = strRtf
            strText = objRichTextManager.Text
        End Using

        Me.Insert(0, New ClipboardTextItem(strText, strRtf))
    End Sub

    Public Sub CopyText(ByVal strText As String)
        Me.Insert(0, New ClipboardTextItem(strText))
    End Sub

    Public Sub CheckClipboardContentInList()
        Dim strText As String = String.Empty
        Dim strRtf As String = String.Empty

        If (Clipboard.GetDataObject().GetDataPresent(DataFormats.Text)) Then
            strText = Clipboard.GetDataObject().GetData(DataFormats.Text).ToString()

            If strText.StartsWith("{\rtf1") Then
                strRtf = strText
                Using objRichTextManager As New RichTextManager
                    With objRichTextManager
                        .Rtf = strRtf
                        .Text = .Text.Replace(vbTab, " ")
                        .Text = .Text.Replace(vbCrLf, " ")
                        strText = .Text
                    End With
                End Using
            Else
                strText = strText.Replace(vbTab, " ")
                strText = strText.Replace(vbCrLf, " ")
            End If
        Else
            Exit Sub
        End If

        Dim NewItem As ClipboardTextItem
        If Me.Count > 0 Then
            Dim FirstItem As ClipboardTextItem = Me(0)
            If strText <> FirstItem.Text Then
                NewItem = New ClipboardTextItem(strText, strRtf)
            Else
                NewItem = Nothing
            End If
        Else
            NewItem = New ClipboardTextItem(strText, strRtf)
        End If
        If NewItem IsNot Nothing Then
            Me.Insert(0, NewItem)
            If String.IsNullOrEmpty(strRtf) Then
                Clipboard.SetText(strText)
            Else
                Clipboard.SetText(strRtf)
            End If
        End If
    End Sub
#End Region
End Class
