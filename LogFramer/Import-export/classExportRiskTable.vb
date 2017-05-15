Public Class ExportRisksTable
    Implements System.IDisposable

    Private objPrintRisksTable As PrintRiskRegister
    Private objTableGrid As New ExportRiskTableRows

    Public Enum RiskCategories As Integer
        NotDefined = 0
        Operational = 1
        Financial = 2
        Objectives = 3
        Reputation = 4
        Other = 5
        All = 6
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, ByVal riskcategory As Integer)
        objPrintRisksTable = New PrintRiskRegister(CurrentLogFrame, riskcategory, 1)
        objPrintRisksTable.CreateList()
    End Sub

    Public Property TableGrid As ExportRiskTableRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportRiskTableRows)
            objTableGrid = value
        End Set
    End Property

    Public Property RiskCategory As Integer
        Get
            Return objPrintRisksTable.RiskCategory
        End Get
        Set(ByVal value As Integer)
            objPrintRisksTable.RiskCategory = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintRiskTableRow As PrintRiskRegisterRow In objPrintRisksTable.PrintList
            objTableGrid.Add(selPrintRiskTableRow)
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

Public Class ExportRiskTableRow
    Inherits PrintRiskRegisterRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintRiskTableRow As PrintRiskRegisterRow)
        With objPrintRiskTableRow
            Me.Assumption = .Assumption
            Me.AssumptionSortNumber = .AssumptionSortNumber
            Me.RiskLevel = .RiskLevel
            Me.RiskResponse = .RiskResponse
            Me.SectionTitle = .SectionTitle
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

    Public ReadOnly Property RiskResponseText As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.RiskDetail IsNot Nothing Then
                Dim strStatement As String
                strStatement = Me.Assumption.RiskDetail.RiskResponseText

                If String.IsNullOrEmpty(strStatement) = False Then strStatement &= ": "
                strStatement &= Me.Assumption.ResponseStrategy

                Return strStatement
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

    Public ReadOnly Property LikelihoodText As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.RiskDetail IsNot Nothing Then
                Dim strStatement As String = Me.Assumption.RiskDetail.LikelihoodText

                Return strStatement
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property ImpactText As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.RiskDetail IsNot Nothing Then
                Dim strStatement As String = String.Empty

                If String.IsNullOrEmpty(Me.Assumption.RiskDetail.RiskImpactText) = False Then
                    With Me.Assumption.RiskDetail
                        Dim strImpactText As String = .RiskImpactText
                        Dim strSplit() As String = strImpactText.Split("-"c)
                        strImpactText = Trim(strSplit(0))

                        strStatement = strImpactText
                    End With
                End If

                Return strStatement
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property RiskLevelText As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.RiskDetail IsNot Nothing Then
                Dim strStatement As String = String.Empty

                If String.IsNullOrEmpty(Me.Assumption.RiskDetail.RiskImpactText) = False Then
                    strStatement = Me.Assumption.RiskDetail.RiskLevel.ToString("P0")
                End If

                Return strStatement
            Else
                Return String.Empty
            End If
        End Get
    End Property
End Class

Public Class ExportRiskTableRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportRiskTableRow As ExportRiskTableRow)
        List.Add(ExportRiskTableRow)
    End Sub

    Public Sub Add(ByVal PrintRiskRegisterRow As PrintRiskRegisterRow)
        Dim NewExportRow As New ExportRiskTableRow(PrintRiskRegisterRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportRiskTableRow As ExportRiskTableRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportRiskTableRow)
        Else
            List.Insert(index, ExportRiskTableRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportRiskTableRow As ExportRiskTableRow)
        If List.Contains(ExportRiskTableRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportRiskTableRow not in list!")
        Else
            List.Remove(ExportRiskTableRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportRiskTableRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportRiskTableRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal RiskTableRow As ExportRiskTableRow) As Integer
        Return List.IndexOf(RiskTableRow)
    End Function

    Public Function Contains(ByVal RiskTableRow As ExportRiskTableRow) As Boolean
        Return List.Contains(RiskTableRow)
    End Function
End Class
