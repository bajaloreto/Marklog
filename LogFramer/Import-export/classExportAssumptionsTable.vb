Public Class ExportAssumptionsTable
    Implements System.IDisposable

    Private objPrintAssumptions As PrintAssumptions
    Private objTableGrid As New ExportAssumptionTableRows

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, ByVal exportsection As Integer)
        objPrintAssumptions = New PrintAssumptions(CurrentLogFrame, exportsection, 1)
        objPrintAssumptions.CreateList()
    End Sub

    Public Property TableGrid As ExportAssumptionTableRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportAssumptionTableRows)
            objTableGrid = value
        End Set
    End Property

    Public Property ExportSection As Integer
        Get
            Return objPrintAssumptions.PrintSection
        End Get
        Set(ByVal value As Integer)
            objPrintAssumptions.PrintSection = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintAssumptionTableRow As PrintAssumptionRow In objPrintAssumptions.PrintList
            objTableGrid.Add(selPrintAssumptionTableRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 8

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

Public Class ExportAssumptionTableRow
    Inherits PrintAssumptionRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintAssumptionRow As PrintAssumptionRow)
        With objPrintAssumptionRow
            Me.Assumption = .Assumption
            Me.AssumptionSortNumber = .AssumptionSortNumber
            Me.Struct = .Struct
            Me.StructSortNumber = .StructSortNumber
        End With
    End Sub

    Public ReadOnly Property AssumptionText As String
        Get
            If Me.Assumption IsNot Nothing Then
                Return Me.Assumption.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property StructText As String
        Get
            If Me.Struct IsNot Nothing Then
                Return Me.Struct.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property
End Class

Public Class ExportAssumptionTableRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportAssumptionTableRow As ExportAssumptionTableRow)
        List.Add(ExportAssumptionTableRow)
    End Sub

    Public Sub Add(ByVal PrintAssumptionRow As PrintAssumptionRow)
        Dim NewExportRow As New ExportAssumptionTableRow(PrintAssumptionRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportAssumptionTableRow As ExportAssumptionTableRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportAssumptionTableRow)
        Else
            List.Insert(index, ExportAssumptionTableRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportAssumptionTableRow As ExportAssumptionTableRow)
        If List.Contains(ExportAssumptionTableRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportAssumptionTableRow not in list!")
        Else
            List.Remove(ExportAssumptionTableRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportAssumptionTableRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportAssumptionTableRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal AssumptionTableRow As ExportAssumptionTableRow) As Integer
        Return List.IndexOf(AssumptionTableRow)
    End Function

    Public Function Contains(ByVal AssumptionTableRow As ExportAssumptionTableRow) As Boolean
        Return List.Contains(AssumptionTableRow)
    End Function
End Class
