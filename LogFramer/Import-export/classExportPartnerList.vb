Public Class ExportPartnerList
    Implements System.IDisposable

    Private objPrintPartnerList As PrintPartnerList
    Private objTableGrid As New ExportPartnerListRows

    Public Enum ObjectTypes
        Organisation = 0
        Address = 1
        TelephoneNumber = 2
        Email = 3
        Website = 4
        Contact = 5
        BankAccount = 6
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame)
        objPrintPartnerList = New PrintPartnerList(logframe.ProjectPartners)
        objPrintPartnerList.CreateList()
    End Sub

    Public Property TableGrid As ExportPartnerListRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportPartnerListRows)
            objTableGrid = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintPartnerListRow As PrintPartnerRow In objPrintPartnerList.PrintList
            objTableGrid.Add(selPrintPartnerListRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 2

        Return intColumns
    End Function

    Public Function GetRowCount() As Integer
        Return Me.TableGrid.Count
    End Function

    Public Function GetRowCountOfOrganisation(ByVal intStartIndex As Integer) As Integer
        Dim intRowCount As Integer
        Dim selPrintPartnerListRow As PrintPartnerRow

        For i = intStartIndex + 1 To objPrintPartnerList.PrintList.Count - 1
            selPrintPartnerListRow = objPrintPartnerList.PrintList(i)

            If selPrintPartnerListRow.PropertyName.StartsWith(LANG_Organisation) Then
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

Public Class ExportPartnerListRow
    Inherits PrintPartnerRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintPartnerRow As PrintPartnerRow)
        With objPrintPartnerRow
            Me.Bold = .Bold
            Me.PropertyName = .PropertyName
            Me.PropertyType = .PropertyType
            Me.PropertyValue = .PropertyValue
        End With
    End Sub
End Class

Public Class ExportPartnerListRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportPartnerListRow As ExportPartnerListRow)
        List.Add(ExportPartnerListRow)
    End Sub

    Public Sub Add(ByVal PrintPartnerRow As PrintPartnerRow)
        Dim NewExportRow As New ExportPartnerListRow(PrintPartnerRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportPartnerListRow As ExportPartnerListRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportPartnerListRow)
        Else
            List.Insert(index, ExportPartnerListRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportPartnerListRow As ExportPartnerListRow)
        If List.Contains(ExportPartnerListRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportPartnerListRow not in list!")
        Else
            List.Remove(ExportPartnerListRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportPartnerListRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportPartnerListRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal PartnerListRow As ExportPartnerListRow) As Integer
        Return List.IndexOf(PartnerListRow)
    End Function

    Public Function Contains(ByVal PartnerListRow As ExportPartnerListRow) As Boolean
        Return List.Contains(PartnerListRow)
    End Function
End Class
