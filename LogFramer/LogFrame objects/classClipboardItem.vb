Imports System.Reflection

Public Class ClipboardItem
    Private objItem As Object
    Private strSortNumber As String
    Private strText As String
    Private datCopyGroup As Date
    Private intChildren As Integer
    Private objReferenceGuid As Guid

    Public Property CopyGroup As Date
        Get
            Return datCopyGroup
        End Get
        Set(ByVal value As Date)
            datCopyGroup = value
        End Set
    End Property

    Public Property Item As Object
        Get
            Return objItem
        End Get
        Set(ByVal value As Object)
            Dim selItem As LogframeObject = TryCast(value, LogframeObject)

            If selItem IsNot Nothing Then
                objReferenceGuid = selItem.Guid
            End If

            Using copier As New ObjectCopy
                objItem = copier.CopyObject(value)
            End Using
        End Set
    End Property

    Public Property ReferenceGuid As Guid
        Get
            Return objReferenceGuid
        End Get
        Set(value As Guid)
            objReferenceGuid = value
        End Set
    End Property

    Public Property SortNumber As String
        Get
            Return strSortNumber
        End Get
        Set(ByVal value As String)
            strSortNumber = value
        End Set
    End Property

    Public Property Text As String
        Get
            Return strText
        End Get
        Set(ByVal value As String)
            strText = value
        End Set
    End Property

    Public Property Children As Integer
        Get
            Return intChildren
        End Get
        Set(ByVal value As Integer)
            intChildren = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal copygroup As Date, ByVal item As Object, ByVal sortnumber As String, Optional ByVal children As Integer = 0)
        Me.CopyGroup = copygroup
        Me.Item = item
        Me.SortNumber = sortnumber
        Me.Children = children

        Dim selType As Type = item.GetType
        Dim selPropInfo As PropertyInfo = selType.GetProperty("Text")
        Dim objValue As Object

        If selPropInfo IsNot Nothing Then
            objValue = selPropInfo.GetValue(item, Nothing)
            If objValue IsNot Nothing Then Me.Text = objValue.ToString
        Else
            Dim selMethodInfo As MethodInfo = selType.GetMethod("ToString")
            objValue = selMethodInfo.Invoke(item, Nothing)
            If objValue IsNot Nothing Then Me.Text = objValue.ToString
        End If
    End Sub
End Class

Public Class ClipboardItems
    Inherits System.Collections.CollectionBase

    Public Event ListUpdated()

    Public Enum PasteOptions As Integer
        PasteAll = 0
        PasteNoVert = 1
        PasteNoDetails = 2
        PasteNoDependencies = 3
    End Enum

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As ClipboardItem
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ClipboardItem)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal clipboarditem As ClipboardItem)
        If clipboarditem IsNot Nothing Then
            List.Add(clipboarditem)
            RaiseEvent ListUpdated()
        End If
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal clipboarditem As ClipboardItem)
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

    Public Sub Remove(ByVal clipboarditem As ClipboardItem)
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

    Public Function IndexOf(ByVal clipboarditem As ClipboardItem) As Integer
        Return List.IndexOf(clipboarditem)
    End Function

    Public Function Contains(ByVal clipboarditem As ClipboardItem) As Boolean
        Return List.Contains(clipboarditem)
    End Function

    Public Function Contains(ByVal selObject As Object) As Boolean
        For Each selItem As ClipboardItem In Me.List
            If selItem.Item = selObject Then
                Return True
            End If
        Next
        Return False
    End Function
#End Region

#Region "Cut, copy and paste"
    Public Function GetCopyGroup(Optional ByVal datCopyGroup As Date = Nothing) As ClipboardItems
        If Me.List.Count = 0 Then
            Return Nothing
        Else
            Dim PasteItems As New ClipboardItems

            If datCopyGroup = Nothing Then datCopyGroup = Me(0).CopyGroup

            For Each selItem As ClipboardItem In Me.List
                If selItem.CopyGroup = datCopyGroup Then
                    PasteItems.Add(selItem)
                End If
            Next

            Return PasteItems
        End If
    End Function
#End Region
End Class
