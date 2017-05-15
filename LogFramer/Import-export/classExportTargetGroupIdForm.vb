Public Class ExportTargetGroupIdForm
    Implements System.IDisposable

    Private objPrintTargetGroupIdForm As PrintTargetGroupIdForm
    Private objTableGrid As New ExportTargetGroupIdRows

    Public Sub New()

    End Sub

    Public Sub New(ByVal objtargetgroups As TargetGroups)
        objPrintTargetGroupIdForm = New PrintTargetGroupIdForm(objtargetgroups)
        objPrintTargetGroupIdForm.CreateList()
    End Sub

    Public Property TableGrid As ExportTargetGroupIdRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportTargetGroupIdRows)
            objTableGrid = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintTargetGroupIdFormRow As PrintTargetGroupIdFormRow In objPrintTargetGroupIdForm.PrintList
            objTableGrid.Add(selPrintTargetGroupIdFormRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 2

        Return intColumns
    End Function

    Public Function GetRowCount() As Integer
        Return Me.TableGrid.Count
    End Function

    Public Function GetRowCountOfTargetGroup(ByVal intStartIndex As Integer) As Integer
        Dim intRowCount As Integer
        Dim selPrintTargetGroupIdFormRow As PrintTargetGroupIdFormRow

        For i = intStartIndex + 1 To objPrintTargetGroupIdForm.PrintList.Count - 1
            selPrintTargetGroupIdFormRow = objPrintTargetGroupIdForm.PrintList(i)

            If selPrintTargetGroupIdFormRow.IsTitle Then
                Exit For
            End If
            intRowCount += 1
        Next

        Return intRowCount
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

Public Class ExportTargetGroupIdRow
    Inherits PrintTargetGroupIdFormRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintTargetGroupIdFormRow As PrintTargetGroupIdFormRow)
        With objPrintTargetGroupIdFormRow
            Me.IsTitle = .IsTitle
            Me.PropertyName = .PropertyName
            Me.PropertyType = .PropertyType
            Me.PropertyValue = .PropertyValue
        End With
    End Sub
End Class

Public Class ExportTargetGroupIdRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportTargetGroupIdRow As ExportTargetGroupIdRow)
        List.Add(ExportTargetGroupIdRow)
    End Sub

    Public Sub Add(ByVal PrintTargetGroupIdFormRow As PrintTargetGroupIdFormRow)
        Dim NewExportRow As New ExportTargetGroupIdRow(PrintTargetGroupIdFormRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportTargetGroupIdRow As ExportTargetGroupIdRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportTargetGroupIdRow)
        Else
            List.Insert(index, ExportTargetGroupIdRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportTargetGroupIdRow As ExportTargetGroupIdRow)
        If List.Contains(ExportTargetGroupIdRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportTargetGroupIdRow not in list!")
        Else
            List.Remove(ExportTargetGroupIdRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportTargetGroupIdRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportTargetGroupIdRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal PartnerListRow As ExportTargetGroupIdRow) As Integer
        Return List.IndexOf(PartnerListRow)
    End Function

    Public Function Contains(ByVal PartnerListRow As ExportTargetGroupIdRow) As Boolean
        Return List.Contains(PartnerListRow)
    End Function
End Class
