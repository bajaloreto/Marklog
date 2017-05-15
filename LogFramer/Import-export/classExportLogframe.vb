Public Class ExportLogframe
    Implements System.IDisposable

    Private objPrintLogFrame As New PrintLogFrame
    Private objTableGrid As New ExportLogframeRows

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, _
                   ByVal boolShowIndicatorColumn As Boolean, ByVal boolShowVerificationSourceColumn As Boolean, ByVal boolShowAssumptionColumn As Boolean, _
                   ByVal boolShowGoals As Boolean, ByVal boolShowPurposes As Boolean, ByVal boolShowOutputs As Boolean, ByVal boolShowActivities As Boolean, _
                   ByVal boolShowResourcesBudget As Boolean)

        objPrintLogFrame.Logframe = logframe
        objPrintLogFrame.ShowIndicatorColumn = boolShowIndicatorColumn
        objPrintLogFrame.ShowVerificationSourceColumn = boolShowVerificationSourceColumn
        objPrintLogFrame.ShowAssumptionColumn = boolShowAssumptionColumn
        objPrintLogFrame.ShowGoals = boolShowGoals
        objPrintLogFrame.ShowPurposes = boolShowPurposes
        objPrintLogFrame.ShowOutputs = boolShowOutputs
        objPrintLogFrame.ShowActivities = boolShowActivities
        objPrintLogFrame.ShowResourcesBudget = boolShowResourcesBudget
    End Sub

    Public Property TableGrid As ExportLogframeRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportLogframeRows)
            objTableGrid = value
        End Set
    End Property

    Public ReadOnly Property ShowIndicatorColumn() As Boolean
        Get
            Return objPrintLogFrame.ShowIndicatorColumn
        End Get
    End Property

    Public ReadOnly Property ShowVerificationSourceColumn() As Boolean
        Get
            Return objPrintLogFrame.ShowVerificationSourceColumn
        End Get
    End Property

    Public ReadOnly Property ShowAssumptionColumn() As Boolean
        Get
            Return objPrintLogFrame.ShowAssumptionColumn
        End Get
    End Property

    Public ReadOnly Property ShowGoals() As Boolean
        Get
            Return objPrintLogFrame.ShowGoals
        End Get
    End Property

    Public ReadOnly Property ShowPurposes() As Boolean
        Get
            Return objPrintLogFrame.ShowPurposes
        End Get
    End Property

    Public ReadOnly Property ShowOutputs() As Boolean
        Get
            Return objPrintLogFrame.ShowOutputs
        End Get
    End Property

    Public ReadOnly Property ShowActivities() As Boolean
        Get
            Return objPrintLogFrame.ShowActivities
        End Get
    End Property

    Public ReadOnly Property ShowResourcesBudget() As Boolean
        Get
            Return objPrintLogFrame.ShowResourcesBudget
        End Get
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()
        objPrintLogFrame.LoadSections()

        For Each selPrintLogframeRow As PrintLogframeRow In objPrintLogFrame.PrintList
            objTableGrid.Add(selPrintLogframeRow)
        Next
    End Sub

    Public Function IsResourceBudgetRow(ByVal selRow As ExportLogframeRow) As Boolean
        If selRow.Section = LogFrame.SectionTypes.ActivitiesSection And Me.ShowResourcesBudget = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetColumnCount() As Integer
        Dim intColumnCount As Integer = 2

        With objPrintLogFrame
            If .ShowIndicatorColumn = True Then
                intColumnCount += 2
            End If
            If .ShowVerificationSourceColumn = True Then
                intColumnCount += 2
            End If
            If .ShowAssumptionColumn = True Then
                intColumnCount += 2
            End If
        End With

        Return intColumnCount
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

Public Class ExportLogframeRow
    Inherits PrintLogframeRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintLogframeRow As PrintLogframeRow)
        With objPrintLogframeRow
            Me.Assumption = .Assumption
            Me.AssumptionSort = .AssumptionSort
            Me.Indicator = .Indicator
            Me.IndicatorSort = .IndicatorSort
            Me.Resource = .Resource
            Me.ResourceSort = .ResourceSort
            Me.Struct = .Struct
            Me.StructSort = .StructSort
            Me.VerificationSource = .VerificationSource
            Me.VerificationSourceSort = .VerificationSourceSort
            Me.RowType = .RowType
            Me.Section = .Section
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

    Public ReadOnly Property IndicatorText As String
        Get
            If Me.Indicator IsNot Nothing Then
                Return Me.Indicator.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property VerificationSourceText As String
        Get
            If Me.VerificationSource IsNot Nothing Then
                Return Me.VerificationSource.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property AssumptionText As String
        Get
            If Me.Assumption IsNot Nothing Then
                Return Me.Assumption.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property ResourceText As String
        Get
            If Me.Resource IsNot Nothing Then
                Return Me.Resource.Text
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property BudgetText As String
        Get
            Dim strBudget As String = String.Empty

            If Me.Resource IsNot Nothing Then
                strBudget = String.Format("{0} {1}", Me.TotalCostAmount.ToString("N2"), CurrentLogFrame.CurrencyCode)
            End If

            Return strBudget
        End Get
    End Property
End Class

Public Class ExportLogframeRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportLogframeRow As ExportLogframeRow)
        List.Add(ExportLogframeRow)
    End Sub

    Public Sub Add(ByVal PrintLogframeRow As PrintLogframeRow)
        Dim NewExportRow As New ExportLogframeRow(PrintLogframeRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportLogframeRow As ExportLogframeRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportLogframeRow)
        Else
            List.Insert(index, ExportLogframeRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportLogframeRow As ExportLogframeRow)
        If List.Contains(ExportLogframeRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportLogframeRow not in list!")
        Else
            List.Remove(ExportLogframeRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportLogframeRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportLogframeRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal LogframeRow As ExportLogframeRow) As Integer
        Return List.IndexOf(LogframeRow)
    End Function

    Public Function Contains(ByVal LogframeRow As ExportLogframeRow) As Boolean
        Return List.Contains(LogframeRow)
    End Function
End Class
