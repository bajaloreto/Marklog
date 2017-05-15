Public Class ExportPmf
    Implements System.IDisposable

    Private objPrintPmf As PrintPMF
    Private objTableGrid As New ExportPmfRows

    Public Enum ExportSections As Integer
        NotSelected = -1
        ExportGoals = 0
        ExportPurposes = 1
        ExportOutputs = 2
        ExportActivities = 3
        ExportAll = 4
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal logframe As LogFrame, ByVal exportsection As Integer, ByVal boolShowTargetRowTitles As Boolean)
        objPrintPmf = New PrintPMF(logframe, exportsection, 1, boolShowTargetRowTitles)
        objPrintPmf.CreateList()
    End Sub

    Public Property TableGrid As ExportPmfRows
        Get
            Return objTableGrid
        End Get
        Set(ByVal value As ExportPmfRows)
            objTableGrid = value
        End Set
    End Property

    Public Property ExportSection As Integer
        Get
            Return objPrintPmf.PrintSection
        End Get
        Set(ByVal value As Integer)
            objPrintPmf.PrintSection = value
        End Set
    End Property

    Public ReadOnly Property TargetDeadlinesSection As TargetDeadlinesSection
        Get
            Return objPrintPmf.TargetDeadlinesSection
        End Get
    End Property

    Public Sub LoadTable()
        objTableGrid.Clear()

        For Each selPrintPmfRow As PrintPmfRow In objPrintPmf.PrintList
            objTableGrid.Add(selPrintPmfRow)
        Next
    End Sub

    Public Function GetColumnCount() As Integer
        Dim intColumns As Integer = 10

        If objPrintPmf.TargetDeadlinesSection IsNot Nothing Then
            intColumns += objPrintPmf.TargetDeadlinesSection.TargetDeadlines.Count
        End If

        Return intColumns
    End Function

    Public Function GetTargetColumnCount() As Integer
        Dim intColumns As Integer = objPrintPmf.TargetDeadlinesSection.TargetDeadlines.Count

        Return intColumns
    End Function

    Public Function GetRowCount(Optional ByVal ToWord As Boolean = False) As Integer
        Dim intCounter As Integer = Me.TableGrid.Count

        If ToWord = False Then
            intCounter += CountIndicatorsBeneficiaryLevel()
        End If

        Return intCounter
    End Function

    Private Function CountIndicatorsBeneficiaryLevel() As Integer
        Dim intCounter As Integer

        For Each selPrintPmfRow As PrintPmfRow In objPrintPmf.PrintList
            If selPrintPmfRow.Indicator IsNot Nothing AndAlso selPrintPmfRow.Indicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                Select Case selPrintPmfRow.Indicator.QuestionType
                    Case Indicator.QuestionTypes.OpenEnded, Indicator.QuestionTypes.MaxDiff, Indicator.QuestionTypes.Image, Indicator.QuestionTypes.ImageWithTargets

                    Case Else
                        intCounter += 1
                End Select
            End If
        Next

        Return intCounter
    End Function

    Public Function GetColumnTitleObjectives() As String
        Return Me.objPrintPmf.GetColumnTitleObjectives
    End Function

    Public Function GetColumnTitlesTargets() As String()
        Dim strHeaderText(objPrintPmf.TargetDeadlinesSection.TargetDeadlines.Count - 1) As String
        Dim strDate As String
        Dim intTargetIndex As Integer

        For Each selTargetDeadline As TargetDeadline In objPrintPmf.TargetDeadlinesSection.TargetDeadlines

            Select Case objPrintPmf.TargetDeadlinesSection.Repetition
                Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                    strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
                Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                    strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
                Case Else
                    strDate = selTargetDeadline.ExactDeadline.ToShortDateString
            End Select

            strHeaderText(intTargetIndex) = String.Format("{0} {1}", LANG_Target, strDate)

            intTargetIndex += 1
        Next

        Return strHeaderText
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

Public Class ExportPmfRow
    Inherits PrintPmfRow

    Public Sub New()

    End Sub

    Public Sub New(ByVal objPrintPmfRow As PrintPmfRow)
        With objPrintPmfRow
            Me.CollectionMethod = .CollectionMethod
            Me.Frequency = .Frequency
            Me.Indicator = .Indicator
            Me.IndicatorSort = .IndicatorSort
            Me.Responsibility = .Responsibility
            Me.Section = .Section
            Me.Struct = .Struct
            Me.StructSort = .StructSort
            Me.VerificationSource = .VerificationSource
            Me.VerificationSourceSort = .VerificationSourceSort
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
End Class

Public Class ExportPmfRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal ExportPmfRow As ExportPmfRow)
        List.Add(ExportPmfRow)
    End Sub

    Public Sub Add(ByVal PrintPmfRow As PrintPmfRow)
        Dim NewExportRow As New ExportPmfRow(PrintPmfRow)

        List.Add(NewExportRow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal ExportPmfRow As ExportPmfRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(ExportPmfRow)
        Else
            List.Insert(index, ExportPmfRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal ExportPmfRow As ExportPmfRow)
        If List.Contains(ExportPmfRow) = False Then
            System.Windows.Forms.MessageBox.Show("ExportPmfRow not in list!")
        Else
            List.Remove(ExportPmfRow)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExportPmfRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), ExportPmfRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal PmfRow As ExportPmfRow) As Integer
        Return List.IndexOf(PmfRow)
    End Function

    Public Function Contains(ByVal PmfRow As ExportPmfRow) As Boolean
        Return List.Contains(PmfRow)
    End Function
End Class
