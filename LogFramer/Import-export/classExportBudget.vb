Public Class ExportBudget
    Implements System.IDisposable

    Private objPrintBudget As PrintBudget
    Private objTableGrid As New ExportBudgetRows

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, ByVal budgetyearindex As Integer)
        objPrintBudget = New PrintBudget(CurrentLogFrame, budgetyearindex, True, True, False, 1)
        objPrintBudget.CreateList()
    End Sub

    Public Property TableGrid As ExportBudgetRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportBudgetRows)
            objTableGrid = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintBudgetRow As PrintBudgetRow In objPrintBudget.PrintList
            objTableGrid.Add(selPrintBudgetRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 12

        Return intColumns
    End Function

    Public Function GetRowCount() As Integer
        Return Me.TableGrid.Count
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

Public Class ExportBudgetRow
    Inherits PrintBudgetRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintBudgetRow As PrintBudgetRow)
        With objPrintBudgetRow
            Me.BudgetItem = .BudgetItem
            Me.SortNumber = .SortNumber
        End With
    End Sub
End Class

Public Class ExportBudgetRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportBudgetRow As ExportBudgetRow)
        List.Add(ExportBudgetRow)
    End Sub

    Public Sub Add(ByVal PrintBudgetRow As PrintBudgetRow)
        Dim NewExportRow As New ExportBudgetRow(PrintBudgetRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportBudgetRow As ExportBudgetRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportBudgetRow)
        Else
            List.Insert(index, ExportBudgetRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportBudgetRow As ExportBudgetRow)
        If List.Contains(ExportBudgetRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportBudgetRow not in list!")
        Else
            List.Remove(ExportBudgetRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportBudgetRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportBudgetRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal BudgetRow As ExportBudgetRow) As Integer
        Return List.IndexOf(BudgetRow)
    End Function

    Public Function Contains(ByVal BudgetRow As ExportBudgetRow) As Boolean
        Return List.Contains(BudgetRow)
    End Function
End Class
