Public Class ExportRiskRegister
    Implements System.IDisposable

    Private objPrintRiskRegister As PrintRiskRegister
    Private objRiskMonitoring As RiskMonitoring
    Private objTableGrid As New ExportRiskRegisterRows

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
        objPrintRiskRegister = New PrintRiskRegister(logframe, riskcategory, 1)
        objRiskMonitoring = logframe.RiskMonitoring
        objPrintRiskRegister.CreateList()
    End Sub

    Public Property TableGrid As ExportRiskRegisterRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportRiskRegisterRows)
            objTableGrid = value
        End Set
    End Property

    Public Property RiskCategory As Integer
        Get
            Return objPrintRiskRegister.RiskCategory
        End Get
        Set(ByVal value As Integer)
            objPrintRiskRegister.RiskCategory = value
        End Set
    End Property

    Public Property RiskMonitoring As RiskMonitoring
        Get
            Return objRiskMonitoring
        End Get
        Set(ByVal value As RiskMonitoring)
            objRiskMonitoring = value
        End Set
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintRiskRegisterRow As PrintRiskRegisterRow In objPrintRiskRegister.PrintList
            objTableGrid.Add(selPrintRiskRegisterRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 9
        Dim intDeadlines As Integer = GetDeadlinesCount()

        intColumns += (intDeadlines * 3)

        Return intColumns
    End Function

    Public Function GetRowCount() As Integer
        Return Me.TableGrid.Count
    End Function

    Public Function GetDeadlinesCount() As Integer
        Dim intCount As Integer

        If Me.RiskMonitoring IsNot Nothing Then
            intCount = Me.RiskMonitoring.RiskMonitoringDeadlines.Count
        End If

        Return intCount
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

Public Class ExportRiskRegisterRow
    Inherits PrintRiskRegisterRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintRiskRegisterRow As PrintRiskRegisterRow)
        With objPrintRiskRegisterRow
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

Public Class ExportRiskRegisterRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportRiskRegisterRow As ExportRiskRegisterRow)
        List.Add(ExportRiskRegisterRow)
    End Sub

    Public Sub Add(ByVal PrintRiskRegisterRow As PrintRiskRegisterRow)
        Dim NewExportRow As New ExportRiskRegisterRow(PrintRiskRegisterRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportRiskRegisterRow As ExportRiskRegisterRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportRiskRegisterRow)
        Else
            List.Insert(index, ExportRiskRegisterRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportRiskRegisterRow As ExportRiskRegisterRow)
        If List.Contains(ExportRiskRegisterRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportRiskRegisterRow not in list!")
        Else
            List.Remove(ExportRiskRegisterRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportRiskRegisterRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportRiskRegisterRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal RiskRegisterRow As ExportRiskRegisterRow) As Integer
        Return List.IndexOf(RiskRegisterRow)
    End Function

    Public Function Contains(ByVal RiskRegisterRow As ExportRiskRegisterRow) As Boolean
        Return List.Contains(RiskRegisterRow)
    End Function
End Class
