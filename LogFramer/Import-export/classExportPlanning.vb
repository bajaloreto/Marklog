Public Class ExportPlanning
    Implements System.IDisposable

    Private objPrintPlanning As PrintPlanning
    Private objTableGrid As New ExportPlanningRows

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, ByVal periodview As Integer, ByVal planningelements As Integer, ByVal periodfrom As Date, ByVal perioduntil As Date, ByVal repeatrowheaders As Boolean, ByVal showdatescolumns As Boolean)
        objPrintPlanning = New PrintPlanning(logframe, Guid.Empty, periodview, planningelements, periodfrom, perioduntil, False, False, repeatrowheaders, showdatescolumns)
    End Sub

    Public Property TableGrid As ExportPlanningRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportPlanningRows)
            objTableGrid = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()
        objPrintPlanning.CreateList()

        For Each selPrintPlanningRow As PrintPlanningRow In objPrintPlanning.PrintList
            objTableGrid.Add(selPrintPlanningRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intYears As Integer = Math.Abs((objPrintPlanning.PeriodUntil.Year - objPrintPlanning.PeriodFrom.Year))
        Dim intMonths As Integer = ((intYears * 12) + Math.Abs((objPrintPlanning.PeriodUntil.Month - objPrintPlanning.PeriodFrom.Month))) + 1

        Return intMonths
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

Public Class ExportPlanningRow
    Inherits PrintPlanningRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintPlanningRow As PrintPlanningRow)
        With objPrintPlanningRow
            Me.KeyMoment = .KeyMoment
            Me.RowType = .RowType
            Me.Section = .Section
            Me.SortNumber = .SortNumber
            Me.Struct = .Struct
        End With
    End Sub

    Public ReadOnly Property StructText As String
        Get
            If Me.Struct IsNot Nothing Then
                Return Me.Struct.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property KeyMomentText As String
        Get
            If Me.KeyMoment IsNot Nothing Then
                Return Me.KeyMoment.Description
            Else
                Return String.Empty
            End If
        End Get
    End Property
End Class

Public Class ExportPlanningRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportPlanningRow As ExportPlanningRow)
        List.Add(ExportPlanningRow)
    End Sub

    Public Sub Add(ByVal PrintPlanningRow As PrintPlanningRow)
        Dim NewExportRow As New ExportPlanningRow(PrintPlanningRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportPlanningRow As ExportPlanningRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportPlanningRow)
        Else
            List.Insert(index, ExportPlanningRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportPlanningRow As ExportPlanningRow)
        If List.Contains(ExportPlanningRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportPlanningRow not in list!")
        Else
            List.Remove(ExportPlanningRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportPlanningRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportPlanningRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal PlanningRow As ExportPlanningRow) As Integer
        Return List.IndexOf(PlanningRow)
    End Function

    Public Function Contains(ByVal PlanningRow As ExportPlanningRow) As Boolean
        Return List.Contains(PlanningRow)
    End Function
End Class
