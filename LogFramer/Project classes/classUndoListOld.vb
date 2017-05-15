Imports System.Reflection

Public Class UndoListItemOld
    Private objGuid As Guid
    Private intObjectType As Integer
    Private strItem As String
    Private intActionIndex As Integer
    Private objOldValue As New Object
    Private objNewValue As New Object
    Private strPropertyName As String
    Private objParent As Object
    Private intIndex As Integer, intOldIndex As Integer
    Private boolIsRtf As Boolean
    Private intIncludeNext As Integer


    Public Enum Actions
        NotSet = -1
        ChangeText = 0
        ChangeValue = 1
        ChangeDate = 2
        FontName = 3
        FontSize = 4
        Bold = 5
        Italic = 6
        UnderLine = 7
        StrikeOut = 8
        SuperScript = 9
        SubScript = 10
        TextColor = 11
        BackColor = 12
        AlignLeft = 13
        AlignCenter = 14
        AlignRight = 15

        InsertNew = 16
        Delete = 17
        DeleteNotVertical = 18
        Paste = 19
        Cut = 20
        CutNotVertical = 21

        RevertValues = 22

        ChangeLeftIndent = 23

        CaseUpper = 24
        CaseLower = 25
        CaseSentence = 26
        CaseFirstLetter = 27
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal guid As Guid, ByVal objecttype As Integer, ByVal item As String, ByVal actionindex As Integer, ByVal oldvalue As Object, ByVal newvalue As Object, ByVal isrtf As Boolean, _
                   Optional ByVal propertyname As String = "", Optional ByVal parent As Object = Nothing, Optional ByVal index As Integer = 0, Optional ByVal includenext As Integer = 0)
        Me.Guid = guid
        Me.ObjectType = objecttype
        Me.Item = item
        Me.ActionIndex = actionindex
        Me.OldValue = oldvalue
        Me.NewValue = newvalue
        Me.IsRtf = isrtf
        If String.IsNullOrEmpty(propertyname) = False Then Me.PropertyName = propertyname
        If parent IsNot Nothing Then Me.Parent = parent
        Me.Index = index
        Me.IncludeNext = includenext
    End Sub

    Public Property Guid() As Guid
        Get
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ObjectType() As Integer
        Get
            Return intObjectType
        End Get
        Set(ByVal value As Integer)
            intObjectType = value
        End Set
    End Property

    Public Property Item() As String
        Get
            Return strItem
        End Get
        Set(ByVal value As String)
            strItem = value
        End Set
    End Property

    Public Property ActionIndex() As Integer
        Get
            Return intActionIndex
        End Get
        Set(ByVal value As Integer)
            intActionIndex = value
        End Set
    End Property

    Public ReadOnly Property Action() As String
        Get
            If Me.ActionIndex >= 0 Then
                Return LIST_UndoActions(intActionIndex)
            Else
                Return "-*-"
            End If
        End Get
    End Property

    Public ReadOnly Property Change() As String
        Get
            Dim strChange As String = String.Empty
            Select Case Me.ActionIndex
                Case Actions.ChangeText
                    Dim strOld As String = String.Empty, strNew As String = String.Empty
                    Dim intCount As Integer

                    If Me.IsRtf = True Then
                        Using objPainter As New RichTextManager
                            If OldValue IsNot Nothing Then
                                objPainter.Rtf = OldValue
                                strOld = objPainter.Text
                            End If
                            If NewValue IsNot Nothing Then
                                objPainter.Rtf = NewValue
                                strNew = objPainter.Text
                            End If
                        End Using
                    Else
                        If OldValue IsNot Nothing Then strOld = OldValue.ToString
                        If NewValue IsNot Nothing Then strNew = NewValue.ToString
                    End If

                    Dim strOldWords() As String = strOld.Split(" "c)
                    Dim strNewWords() As String = strNew.Split(" "c)

                    If strOldWords.Count < strNewWords.Count Then
                        intCount = strNewWords.Count - 1
                    Else
                        intCount = strOldWords.Count - 1
                    End If

                    For i = 0 To intCount
                        If i <= strOldWords.Count - 1 And i <= strNewWords.Count - 1 Then
                            If String.Equals(strOldWords(i).ToLower, strNewWords(i).ToLower) = False Then
                                strChange = Change_Text(strNewWords, i)
                                Exit For
                            End If
                        ElseIf i > strOldWords.Count - 1 Then
                            strChange = Change_Text(strNewWords, i)
                            Exit For
                        ElseIf i > strNewWords.Count - 1 Then
                            strChange = Change_Text(strOldWords, i)
                            Exit For
                        End If
                    Next
            End Select

            Return strChange
        End Get
    End Property

    Private Function Change_Text(ByVal strWords() As String, ByVal intIndex As Integer) As String
        Dim strChange As String = ""
        Dim intLastIndex As Integer

        intLastIndex = intIndex + 3
        If intLastIndex > strWords.Count - 1 Then intLastIndex = strWords.Count - 1
        For j = intIndex To intLastIndex
            If String.IsNullOrEmpty(strChange) = False Then strChange &= " "
            strChange &= strWords(j)
        Next
        strChange &= "..."

        Return strChange
    End Function

    Public Property OldValue() As Object
        Get
            Return objOldValue
        End Get
        Set(ByVal value As Object)
            objOldValue = value
        End Set
    End Property

    Public Property NewValue() As Object
        Get
            Return objNewValue
        End Get
        Set(ByVal value As Object)
            objNewValue = value
        End Set
    End Property

    Public Property IsRtf() As Boolean
        Get
            Return boolIsRtf
        End Get
        Set(ByVal value As Boolean)
            boolIsRtf = value
        End Set
    End Property

    Public Property PropertyName() As String
        Get
            Return strPropertyName
        End Get
        Set(ByVal value As String)
            strPropertyName = value
        End Set
    End Property

    Public Property Parent() As Object
        Get
            Return objParent
        End Get
        Set(ByVal value As Object)
            objParent = value
        End Set
    End Property

    Public Property OldIndex() As Integer
        Get
            Return intOldIndex
        End Get
        Set(ByVal value As Integer)
            intOldIndex = value
        End Set
    End Property

    Public Property Index() As Integer
        Get
            Return intIndex
        End Get
        Set(ByVal value As Integer)
            intIndex = value
        End Set
    End Property

    Public Property IncludeNext() As Integer
        Get
            Return intIncludeNext
        End Get
        Set(ByVal value As Integer)
            intIncludeNext = value
        End Set
    End Property

    Public Overrides Function Tostring() As String
        Dim strText As String = strItem & " -- " & Me.Action
        If String.IsNullOrEmpty(Me.Change) = False Then strText &= ": " & Me.Change
        Return strText
    End Function
End Class

Public Class RedoListOld
    Inherits System.Collections.CollectionBase

    Public Event UndoItemAdded()
    Public Event UndoItemRemoved()

    Public Sub Add(ByVal listitem As UndoListItemOld)
        List.Add(listitem)

        RaiseEvent UndoItemAdded()
    End Sub

    Public Function Contains(ByVal listitem As UndoListItemOld) As Boolean
        Return List.Contains(listitem)
    End Function

    Public Function IndexOf(ByVal listitem As UndoListItemOld) As Integer
        Return List.IndexOf(listitem)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal listitem As UndoListItemOld)
        If index > Count Or index < 0 Then
            Dim strMsg As String = String.Format(ERR_IndexNotValidCannotInsert, {LANG_UndoListItem})
            System.Windows.Forms.MessageBox.Show(strMsg)
        ElseIf index = Count Then
            List.Add(listitem)
            RaiseEvent UndoItemAdded()
        Else
            List.Insert(index, listitem)
            RaiseEvent UndoItemAdded()
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As UndoListItemOld
        Get
            Return CType(List.Item(index), UndoListItemOld)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
            RaiseEvent UndoItemRemoved()
        End If
    End Sub

    Public Overloads Sub RemoveAt(ByVal index As Integer)
        MyBase.RemoveAt(index)
        RaiseEvent UndoItemRemoved()
    End Sub

    Public Sub Remove(ByVal listitem As UndoListItemOld)
        If List.Contains(listitem) Then List.Remove(listitem)
        RaiseEvent UndoItemRemoved()
    End Sub
End Class

Public Class UndoListOld
    Inherits RedoListOld
    Private objUndoBuffer As New UndoListItemOld

    Public Property UndoBuffer() As UndoListItemOld
        Get
            Return objUndoBuffer
        End Get
        Set(ByVal value As UndoListItemOld)
            objUndoBuffer = value
        End Set
    End Property

    Public Sub SetUndoBuffer(ByVal objGuid As Guid, ByVal intObjectType As Integer, ByVal strItem As String, _
                             ByVal intActionIndex As Integer, ByVal objNewValue As Object, ByVal objOldValue As Object, _
                             ByVal objParent As Object, ByVal intIndex As Integer, ByVal strPropertyName As String, ByVal boolIsRtf As Boolean)
        With Me.UndoBuffer
            .Guid = objGuid
            .ObjectType = intObjectType
            .Item = strItem
            .ActionIndex = intActionIndex
            .NewValue = objNewValue
            .OldValue = objOldValue
            .Parent = objParent
            .Index = intIndex
            .PropertyName = strPropertyName
            .IsRtf = boolIsRtf
        End With
    End Sub

    Public Sub InitialiseUndoBuffer(ByVal selObject As Object, Optional ByVal strSortNumber As String = "", Optional ByVal boolAddToUndoList As Boolean = False)
        If selObject IsNot Nothing Then
            Dim propGuid As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("Guid")
            Dim propItemName As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("ItemName")
            Dim objGuid As Guid
            Dim strItem As String = String.Empty

            If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(selObject, Nothing)
            If propItemName IsNot Nothing Then strItem = propItemName.GetValue(selObject, Nothing)

            With Me.UndoBuffer
                .ObjectType = GetObjectType(selObject)
                .Guid = objGuid
                .Item = Trim(strItem & " " & strSortNumber)
                If .ActionIndex <> UndoListItemOld.Actions.FontName And .ActionIndex <> UndoListItemOld.Actions.FontSize Then _
                    .ActionIndex = UndoListItemOld.Actions.NotSet
                .NewValue = Nothing
                .OldValue = Nothing
                .Parent = Nothing
                .Index = 0
                .PropertyName = String.Empty
            End With

            If boolAddToUndoList = True Then
                Me.AddToUndoList()
            End If
        Else
            With Me.UndoBuffer
                .ObjectType = -1
                .Guid = Guid.Empty
                .Item = String.Empty
                .ActionIndex = UndoListItemOld.Actions.NotSet
                .NewValue = Nothing
                .OldValue = Nothing
                .Parent = Nothing
                .Index = 0
                .PropertyName = String.Empty
            End With
        End If
    End Sub

    Public Function CheckIfValueChanged() As Boolean
        If UndoBuffer.OldValue Is Nothing OrElse UndoBuffer.OldValue.GetType Is GetType(Object) Then Return False
        If UndoBuffer.NewValue Is Nothing OrElse UndoBuffer.NewValue.GetType Is GetType(Object) Then Return False
        If UndoBuffer.OldValue.GetType IsNot UndoBuffer.NewValue.GetType Then Return False

        If UndoBuffer.OldValue <> UndoBuffer.NewValue Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub AddToUndoList()

        If Me.UndoBuffer.Guid <> Guid.Empty Then 'And Me.UndoBuffer.NewValue IsNot Nothing Then
            If (UndoBuffer.ActionIndex = UndoListItemOld.Actions.Delete Or UndoBuffer.ActionIndex = UndoListItemOld.Actions.Cut Or UndoBuffer.ActionIndex = UndoListItemOld.Actions.RevertValues) _
                Xor (Me.UndoBuffer.NewValue IsNot Nothing) Then

                Dim NewUndo As New UndoListItemOld(UndoBuffer.Guid, UndoBuffer.ObjectType, UndoBuffer.Item, UndoBuffer.ActionIndex, _
                                                UndoBuffer.OldValue, UndoBuffer.NewValue, UndoBuffer.IsRtf, UndoBuffer.PropertyName, _
                                                UndoBuffer.Parent, UndoBuffer.Index, UndoBuffer.IncludeNext)
                If Me.Count > 0 Then
                    If NewUndo IsNot Me(0) Then Insert(0, NewUndo)
                Else
                    Insert(0, NewUndo)
                End If
                UndoBuffer.Guid = Guid.Empty
                UndoBuffer.ObjectType = LogFrame.ObjectTypes.NotSet
                UndoBuffer.Item = ""
                UndoBuffer.ActionIndex = UndoListItemOld.Actions.NotSet
                UndoBuffer.OldValue = Nothing
                UndoBuffer.NewValue = Nothing
                UndoBuffer.IsRtf = False
                UndoBuffer.PropertyName = String.Empty
                UndoBuffer.Parent = Nothing
                UndoBuffer.Index = 0
                UndoBuffer.IncludeNext = 0
            End If
        End If
    End Sub

    Private Function GetObjectType(ByVal selObject As Object) As Integer
        Dim intObjectType As Integer = -1
        Dim strTypeName As String = selObject.GetType.Name

        intObjectType = [Enum].Parse(GetType(LogFrame.ObjectTypes), strTypeName)
        Return intObjectType
    End Function

    Private Function GetItemName(ByVal strItem As String, ByVal strPropertyName As String)
        If String.IsNullOrEmpty(strItem) Then strItem = System.Text.RegularExpressions.Regex.Replace(strPropertyName, "([A-Z])", " $1", System.Text.RegularExpressions.RegexOptions.Compiled).Trim()
        strItem = strItem.Substring(0, 1) & strItem.Substring(1).ToLower

        Return strItem
    End Function

#Region "Operations"
    Public Sub InsertNewOperation(ByVal NewValue As Object, ByVal objParent As Object, _
                                  Optional ByVal intIncludeNext As Integer = 0, Optional ByVal strSortNumber As String = "", _
                                  Optional ByVal boolAddToUndoList As Boolean = False)
        Dim propGuid As System.Reflection.PropertyInfo = NewValue.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = NewValue.GetType.GetProperty("ItemName")
        Dim propText As System.Reflection.PropertyInfo = NewValue.GetType.GetProperty("Text")
        Dim objGuid As Guid
        Dim strItem As String = String.Empty
        Dim strText As String = String.Empty

        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(NewValue, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(NewValue, Nothing)
        If propText IsNot Nothing Then strText = propText.GetValue(NewValue, Nothing)
        If String.IsNullOrEmpty(strSortNumber) = False Then strText = strSortNumber

        With Me.UndoBuffer
            .ObjectType = GetObjectType(NewValue)
            .Guid = objGuid
            .Item = Trim(strItem & " " & strText)
            .ActionIndex = UndoListItemOld.Actions.InsertNew
            .NewValue = NewValue
            .OldValue = Nothing
            .Parent = objParent
            .Index = objParent.IndexOf(NewValue)
            .PropertyName = ""
            .IncludeNext = intIncludeNext
        End With

        If boolAddToUndoList = True Then
            Me.AddToUndoList()
        End If
    End Sub

    Public Sub DeleteOperation(ByVal OldValue As Object, ByVal objParent As Object, Optional ByVal intRemoveNext As Integer = 0, _
                               Optional ByVal strSortNumber As String = "", Optional ByVal boolAddToUndoList As Boolean = False)
        If OldValue IsNot Nothing Then
            Dim propGuid As System.Reflection.PropertyInfo = OldValue.GetType.GetProperty("Guid")
            Dim propItemName As System.Reflection.PropertyInfo = OldValue.GetType.GetProperty("ItemName")
            Dim propText As System.Reflection.PropertyInfo = OldValue.GetType.GetProperty("Text")
            Dim objGuid As Guid
            Dim strItem As String = String.Empty
            Dim strText As String = String.Empty

            If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(OldValue, Nothing)
            If propItemName IsNot Nothing Then strItem = propItemName.GetValue(OldValue, Nothing)

            With Me.UndoBuffer
                .ObjectType = GetObjectType(OldValue)
                .Guid = objGuid
                .Item = Trim(strItem & " " & strSortNumber)
                .ActionIndex = UndoListItemOld.Actions.Delete
                .NewValue = Nothing
                .OldValue = OldValue
                .Parent = objParent
                .Index = objParent.IndexOf(OldValue)
                .PropertyName = String.Empty
                .IncludeNext = intRemoveNext
            End With

            If boolAddToUndoList = True Then
                Me.AddToUndoList()
            End If
        End If
    End Sub

    Public Sub CutOperation(ByVal objOldValue As Object, ByVal objParent As Object, ByVal intIndex As Integer, _
                               Optional ByVal strSortNumber As String = "", Optional ByVal boolAddToUndoList As Boolean = False)
        Dim propGuid As System.Reflection.PropertyInfo = objOldValue.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = objOldValue.GetType.GetProperty("ItemName")
        Dim objGuid As Guid
        Dim strItem As String = String.Empty

        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(objOldValue, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(objOldValue, Nothing)

        With Me.UndoBuffer
            .ObjectType = GetObjectType(objOldValue)
            .Guid = objGuid
            .Item = Trim(strItem & " " & strSortNumber)
            .ActionIndex = UndoListItemOld.Actions.Cut
            .NewValue = Nothing
            .OldValue = objOldValue
            .Parent = objParent
            .Index = intIndex
            .PropertyName = String.Empty
        End With

        If boolAddToUndoList = True Then
            Me.AddToUndoList()
        End If
    End Sub

    Public Sub PasteOperation(ByVal objDuplicate As Object, ByVal objParent As Object, ByVal intIndex As Integer, _
                               Optional ByVal strSortNumber As String = "", Optional ByVal boolAddToUndoList As Boolean = False)
        Dim propGuid As System.Reflection.PropertyInfo = objDuplicate.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = objDuplicate.GetType.GetProperty("ItemName")
        Dim objGuid As Guid
        Dim strItem As String = String.Empty

        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(objDuplicate, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(objDuplicate, Nothing)

        With Me.UndoBuffer
            .ObjectType = GetObjectType(objDuplicate)
            .Guid = objGuid
            .Item = Trim(strItem & " " & strSortNumber)
            .ActionIndex = UndoListItemOld.Actions.Paste
            .NewValue = objDuplicate
            .OldValue = Nothing
            .Parent = objParent
            .Index = intIndex
            .PropertyName = String.Empty
        End With

        If boolAddToUndoList = True Then
            Me.AddToUndoList()
        End If
    End Sub

    Public Sub ChangeTextOperation(ByVal intActionIndex As Integer)
        If Me.UndoBuffer.Guid <> Guid.Empty Then
            Dim objOldValue As Object

            If UndoBuffer.ActionIndex <> -1 Then
                objOldValue = UndoBuffer.NewValue
                AddToUndoList()
                UndoBuffer.OldValue = objOldValue
            End If
            UndoBuffer.ActionIndex = intActionIndex
        End If
    End Sub

    Public Sub ChangeTextOperation(ByVal selObject As Object, ByVal strPropertyName As String, ByVal NewText As String, ByVal OldText As String, _
                                   ByVal objParent As Object, ByVal intIndex As Integer, _
                                   Optional ByVal strSortNumber As String = "", Optional ByVal boolAddToUndoList As Boolean = False)
        Dim propGuid As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("ItemName")
        Dim objGuid As Guid
        Dim strItem As String = String.Empty

        strItem = GetItemName(strItem, strPropertyName)
        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(selObject, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(selObject, Nothing)

        With Me.UndoBuffer
            .ObjectType = GetObjectType(selObject)
            .Guid = objGuid
            .Item = Trim(strItem & " " & strSortNumber)
            .ActionIndex = UndoListItemOld.Actions.ChangeText
            .NewValue = NewText
            .OldValue = OldText
            .Parent = objParent
            .Index = intIndex
            .PropertyName = strPropertyName
        End With

        If boolAddToUndoList = True Then
            Me.AddToUndoList()
        End If
    End Sub

    Public Sub ChangeValueOperation(ByVal selObject As Object, ByVal strPropertyName As String, ByVal strItem As String, ByVal objNewValue As Object, ByVal objOldValue As Object, _
                                    Optional ByVal boolAddToUndoList As Boolean = False)
        Dim propGuid As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("ItemName")
        Dim objGuid As Guid

        strItem = GetItemName(strItem, strPropertyName)
        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(selObject, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(selObject, Nothing) & " - " & strItem

        With Me.UndoBuffer
            .ObjectType = GetObjectType(selObject)
            .Guid = objGuid
            .Item = strItem
            .ActionIndex = UndoListItemOld.Actions.ChangeValue
            .NewValue = objNewValue
            .OldValue = objOldValue
            .Parent = Nothing
            .Index = 0
            .PropertyName = strPropertyName
        End With

        If boolAddToUndoList = True Then
            Me.AddToUndoList()
        End If
    End Sub

    Public Sub ChangeValueOperation_BeforeChange(ByVal selObject As Object, ByVal strPropertyName As String, ByVal strItem As String, ByVal objOldValue As Object)
        Dim propGuid As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("ItemName")
        Dim objGuid As Guid

        strItem = GetItemName(strItem, strPropertyName)
        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(selObject, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(selObject, Nothing) & strItem

        With Me.UndoBuffer
            .ObjectType = GetObjectType(selObject)
            .Guid = objGuid
            .Item = strItem
            .ActionIndex = UndoListItemOld.Actions.ChangeValue
            .NewValue = Nothing
            .OldValue = objOldValue
            .Parent = Nothing
            .Index = 0
            .PropertyName = strPropertyName
        End With
    End Sub

    Public Sub ChangeValueOperation_AfterChange(ByVal objNewValue As Object, Optional ByVal boolAddToUndoList As Boolean = False)

        With Me.UndoBuffer
            If .ActionIndex <> UndoListItemOld.Actions.ChangeValue Then Exit Sub
            .NewValue = objNewValue
        End With

        If boolAddToUndoList = True Then
            If UndoBuffer.OldValue IsNot UndoBuffer.NewValue Then AddToUndoList()
        End If
    End Sub

    Public Sub ChangeValueRangeOperation_BeforeChange(ByVal selObject As ValueRange, ByVal strPropertyName As String, ByVal strItem As String, _
                                                      ByVal objParent As Statement)
        Dim propGuid As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("ItemName")
        Dim objGuid As Guid

        strItem = GetItemName(strItem, strPropertyName)
        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(selObject, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(selObject, Nothing) & strItem

        With Me.UndoBuffer
            .ObjectType = GetObjectType(selObject)
            .Guid = objGuid
            .Item = strItem
            .ActionIndex = UndoListItemOld.Actions.ChangeValue
            .NewValue = Nothing
            .OldValue = selObject
            .Parent = objParent
            .Index = 0
            .PropertyName = strPropertyName
        End With
    End Sub

    Public Sub ChangeValueRangeOperation_AfterChange(ByVal objNewValue As Object, Optional ByVal boolAddToUndoList As Boolean = False)

        With Me.UndoBuffer
            If .ActionIndex <> UndoListItemOld.Actions.ChangeValue Then Exit Sub
            .NewValue = objNewValue
        End With

        If boolAddToUndoList = True Then
            If UndoBuffer.OldValue IsNot UndoBuffer.NewValue Then AddToUndoList()
        End If
    End Sub

    Public Sub RevertValuesOperation(ByVal selStatement As Statement, Optional ByVal boolAddToUndoList As Boolean = False)

        With Me.UndoBuffer
            .ObjectType = GetObjectType(selStatement)
            .Guid = selStatement.Guid
            .Item = Trim(Statement.ItemName & RichTextToText(selStatement.FirstLabel))
            .ActionIndex = UndoListItemOld.Actions.RevertValues
            .NewValue = Nothing
            .OldValue = Nothing
            .Parent = Nothing
            .Index = 0
            .PropertyName = String.Empty
            .IncludeNext = 0
        End With

        If boolAddToUndoList = True Then
            Me.AddToUndoList()
        End If
    End Sub

    Public Sub ChangeDateOperation_BeforeChange(ByVal selObject As Object, ByVal strPropertyName As String, ByVal strItem As String, ByVal objOldDate As Object)
        Dim propGuid As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("Guid")
        Dim propItemName As System.Reflection.PropertyInfo = selObject.GetType.GetProperty("ItemName")
        Dim objGuid As Guid

        strItem = GetItemName(strItem, strPropertyName)
        If propGuid IsNot Nothing Then objGuid = propGuid.GetValue(selObject, Nothing)
        If propItemName IsNot Nothing Then strItem = propItemName.GetValue(selObject, Nothing) & " - " & strItem

        With Me.UndoBuffer
            .ObjectType = GetObjectType(selObject)
            .Guid = objGuid
            .Item = strItem
            .ActionIndex = -1
            .NewValue = Nothing
            .OldValue = objOldDate
            .Parent = Nothing
            .Index = 0
            .PropertyName = strPropertyName
        End With
    End Sub

    Public Sub ChangeDateOperation_AfterChange(ByVal objNewDate As Object, Optional ByVal boolAddToUndoList As Boolean = False)

        With Me.UndoBuffer
            If .ActionIndex <> UndoListItemOld.Actions.ChangeDate Then Exit Sub
            .NewValue = objNewDate
        End With

        If boolAddToUndoList = True Then
            If UndoBuffer.OldValue <> UndoBuffer.NewValue Then AddToUndoList()
        End If
    End Sub

    Public Sub ChangeAlignmentOperation(ByVal CurrentObject As LogframeObject, ByVal intAlignment As Integer, ByVal strOldRtf As String)
        InitialiseUndoBuffer(CurrentObject)
        With UndoBuffer
            Select Case intAlignment
                Case HorizontalAlignment.Left
                    .ActionIndex = UndoListItemOld.Actions.AlignLeft
                Case HorizontalAlignment.Center
                    .ActionIndex = UndoListItemOld.Actions.AlignCenter
                Case HorizontalAlignment.Right
                    .ActionIndex = UndoListItemOld.Actions.AlignRight
            End Select

            .OldValue = strOldRtf
        End With
    End Sub

    Public Sub ChangeLeftIndentOperation(ByVal CurrentObject As LogframeObject, ByVal strOldRtf As String)
        InitialiseUndoBuffer(CurrentObject)

        With UndoBuffer
            .ActionIndex = UndoListItemOld.Actions.ChangeLeftIndent
            .OldValue = strOldRtf
        End With
    End Sub
#End Region
End Class
