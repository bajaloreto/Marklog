Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Public Class DialogImportBudget
    Private WorkSheets As List(Of String)
    Private objExcelIO As New ExcelIO
    Private strDocName, strDocPath As String
    Private objBudget As New Budget
    Private strColumnDescription, strColumnDuration As String
    Private strColumnNumber, strColumnUnitCost, strColumnTotalCost As String
    Private intFirstRowIndex As Integer, intLastRowIndex As Integer

#Region "Properties"
    Public Property DocName As String
        Get
            Return strDocName
        End Get
        Set(ByVal value As String)
            strDocName = value
        End Set
    End Property

    Public Property DocPath As String
        Get
            Return strDocPath
        End Get
        Set(ByVal value As String)
            strDocPath = value
        End Set
    End Property

    Public Property Budget() As Budget
        Get
            Return objBudget
        End Get
        Set(ByVal value As Budget)
            objBudget = value
        End Set
    End Property

    Private Property ColumnDescription() As String
        Get
            Return strColumnDescription
        End Get
        Set(ByVal value As String)
            strColumnDescription = value
        End Set
    End Property

    Private Property ColumnDuration() As String
        Get
            Return strColumnDuration
        End Get
        Set(ByVal value As String)
            strColumnDuration = value
        End Set
    End Property

    Private Property ColumnNumber() As String
        Get
            Return strColumnNumber
        End Get
        Set(ByVal value As String)
            strColumnNumber = value
        End Set
    End Property

    Private Property ColumnUnitCost() As String
        Get
            Return strColumnUnitCost
        End Get
        Set(ByVal value As String)
            strColumnUnitCost = value
        End Set
    End Property

    Private Property ColumnTotalCost() As String
        Get
            Return strColumnTotalCost
        End Get
        Set(ByVal value As String)
            strColumnTotalCost = value
        End Set
    End Property

    Private Property FirstRowIndex() As Integer
        Get
            Return intFirstRowIndex
        End Get
        Set(ByVal value As Integer)
            intFirstRowIndex = value
        End Set
    End Property

    Private Property LastRowIndex() As Integer
        Get
            Return intLastRowIndex
        End Get
        Set(ByVal value As Integer)
            intLastRowIndex = value
        End Set
    End Property
#End Region

#Region "Methods and events"
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.Text = LANG_ImportExcelBudgetTitle
        lblStep1.Text = LANG_ImportExcelBudgetStep1
        lblStep2.Text = LANG_ImportExcelBudgetStep2
        lblSelectWorksheet.Text = ToLabel(LANG_WorkSheet)

        dgWorkSheet.Visible = False
        btnImport.Enabled = False

        ntbFirstRow.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbLastRow.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbFirstRow.NrDecimals = 0
        ntbLastRow.NrDecimals = 0
        ntbFirstRow.SetDecimals = True
        ntbLastRow.SetDecimals = True
    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Function CheckAllInfoComplete() As Boolean
        Dim boolComplete As Boolean = True

        If String.IsNullOrEmpty(ColumnDescription) Then boolComplete = False
        If String.IsNullOrEmpty(ColumnTotalCost) Then boolComplete = False

        Return boolComplete
    End Function

    Private Sub btnLoadFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadFile.Click
        Dim DialogOpenWorkbook As New OpenFileDialog

        With DialogOpenWorkbook
            .Title = LANG_ImportExcelBudgetSelectFile
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Multiselect = False
            .Filter = (LANG_ImportExcelFilter)
            .FilterIndex = 0
        End With

        dgWorkSheet.Visible = False

        If DialogOpenWorkbook.ShowDialog() = DialogResult.OK And String.IsNullOrEmpty(DialogOpenWorkbook.FileName) = False Then
            Dim selFileInfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(DialogOpenWorkbook.FileName)
            DocName = selFileInfo.Name
            DocPath = selFileInfo.FullName

            tbFileName.Text = Me.DocName

            Dim objSheetNames As Object

            objExcelIO.FilePath = DialogOpenWorkbook.FileName
            objSheetNames = objExcelIO.GetExcelSheetNames

            If objSheetNames IsNot Nothing Then
                Dim strSheetNames() As String = objSheetNames
                If strSheetNames.Length > 0 Then
                    cmbSelectWorksheet.DataSource = strSheetNames
                    cmbSelectWorksheet.SelectedIndex = 0
                    dgWorkSheet.Visible = True
                End If
            End If
        End If

        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub btnToStep2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToStep2.Click
        TabControlSteps.SelectedIndex = 1
    End Sub

    Private Sub btnToStep3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToStep3.Click
        TabControlSteps.SelectedIndex = 2
    End Sub

    Private Sub btnToStep4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToStep4.Click
        TabControlSteps.SelectedIndex = 3
    End Sub

    Private Sub cmbSelectWorksheet_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectWorksheet.SelectedIndexChanged
        Dim strWorkSheet As String = cmbSelectWorksheet.SelectedItem
        Dim dtWorkSheet As System.Data.DataTable

        dtWorkSheet = objExcelIO.GetDataFromWorkSheet(strWorkSheet)

        If dtWorkSheet IsNot Nothing Then
            intFirstRowIndex = 0
            intLastRowIndex = dtWorkSheet.Rows.Count - 1
            dgWorkSheet.LoadColumns(dtWorkSheet)

            Dim lstColumns(dtWorkSheet.Columns.Count) As String
            Dim i As Integer
            lstColumns(0) = ""
            For i = 0 To dtWorkSheet.Columns.Count - 1
                lstColumns(i + 1) = dtWorkSheet.Columns(i).ColumnName
            Next

            cmbColumnDescription.Items.Clear()
            cmbColumnDescription.Items.AddRange(lstColumns.ToArray)
            cmbColumnDescription.SelectedItem = "A"
            cmbColumnDuration.Items.Clear()
            cmbColumnDuration.Items.AddRange(lstColumns.ToArray)
            'cmbColumnDuration.SelectedItem = "B"
            cmbColumnNumber.Items.Clear()
            cmbColumnNumber.Items.AddRange(lstColumns.ToArray)
            'cmbColumnVerificationMeans.SelectedItem = "C"
            cmbColumnUnitCost.Items.Clear()
            cmbColumnUnitCost.Items.AddRange(lstColumns.ToArray)
            'cmbColumnUnitCost.SelectedItem = "D"
            cmbColumnTotalCost.Items.Clear()
            cmbColumnTotalCost.Items.AddRange(lstColumns.ToArray)
            'cmbColumnTotalCost.SelectedItem = "E"

            ntbFirstRow.Text = "1"
            ntbLastRow.Text = (intLastRowIndex + 1).ToString
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnDescription_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnDescription.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnDescription_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnDescription.SelectedIndexChanged
        Dim strColumn As String = cmbColumnDescription.SelectedItem.ToString

        Me.ColumnDescription = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnDuration_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnDuration.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnDuration_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnDuration.SelectedIndexChanged
        Dim strColumn As String = cmbColumnDuration.SelectedItem.ToString
        Me.ColumnDuration = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnNumber.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnNumber_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnNumber.SelectedIndexChanged
        Dim strColumn As String = cmbColumnNumber.SelectedItem.ToString
        Me.ColumnNumber = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnUnitCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnUnitCost.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnUnitCost_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnUnitCost.SelectedIndexChanged
        Dim strColumn As String = cmbColumnUnitCost.SelectedItem.ToString
        Me.ColumnUnitCost = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnTotalCost_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnTotalCost.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnTotalCost_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnTotalCost.SelectedIndexChanged
        Dim strColumn As String = cmbColumnTotalCost.SelectedItem.ToString
        Me.ColumnTotalCost = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbFirstRow_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbFirstRow.Validated
        If String.IsNullOrEmpty(ntbFirstRow.Text) = False Then
            FirstRowIndex = ntbFirstRow.IntegerValue
            FirstRowIndex -= 1
            If FirstRowIndex < 0 Then ntbFirstRow.Text = 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbLastRow_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbLastRow.Validated
        If String.IsNullOrEmpty(ntbLastRow.Text) = False Then
            LastRowIndex = ntbLastRow.IntegerValue
            LastRowIndex -= 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub
#End Region

#Region "Import"
    Public Function ImportBudget() As Budget
        Cursor.Current = Cursors.WaitCursor

        With Me.Budget
            .MultiYearBudget = False
            .UpdateBudgetYears(CurrentLogFrame.StartDate, CurrentLogFrame.EndDate)
        End With

        ImportBudget_Import()

        Cursor.Current = Cursors.Default

        Return Me.Budget
    End Function

    Public Sub ImportBudget_Import()
        Dim selRow As DataRow
        Dim objBudgetYear As BudgetYear = Me.Budget.BudgetYears(0)
        Dim objLastBudgetItem As BudgetItem = Nothing
        Dim sngDuration As Single
        Dim sngNumber As Single
        Dim sngUnitCost As Single
        Dim sngTotalCost As Single
        Dim strCurrency As String = My.Settings.setDefaultCurrency
        Dim i As Integer

        If objBudgetYear Is Nothing Then Exit Sub

        If FirstRowIndex >= 0 And LastRowIndex >= FirstRowIndex Then
            For i = FirstRowIndex To LastRowIndex
                selRow = dgWorkSheet.WorkSheet.Rows(i)
                Dim objBudgetItem As New BudgetItem

                If String.IsNullOrEmpty(ColumnDescription) = False AndAlso IsDBNull(selRow(ColumnDescription)) = False Then
                    objBudgetItem.SetText(selRow(ColumnDescription).ToString)
                End If
                If String.IsNullOrEmpty(ColumnDuration) = False AndAlso IsDBNull(selRow(ColumnDuration)) = False Then
                    sngDuration = ParseSingle(selRow(ColumnDuration))

                    If sngDuration > 0 Then
                        objBudgetItem.Duration = sngDuration
                        objBudgetItem.DurationUnit = LANG_Months
                    End If
                End If
                If String.IsNullOrEmpty(ColumnNumber) = False AndAlso IsDBNull(selRow(ColumnNumber)) = False Then
                    sngNumber = ParseSingle(selRow(ColumnNumber))

                    If sngNumber > 0 Then
                        objBudgetItem.Number = sngNumber
                        objBudgetItem.NumberUnit = LANG_Units
                    End If
                End If
                If String.IsNullOrEmpty(ColumnUnitCost) = False AndAlso IsDBNull(selRow(ColumnUnitCost)) = False Then
                    sngUnitCost = ParseSingle(selRow(ColumnUnitCost))

                    If sngUnitCost <> 0 Then
                        objBudgetItem.UnitCost = New Currency(sngUnitCost, strCurrency)
                    End If
                End If

                If String.IsNullOrEmpty(ColumnTotalCost) = False AndAlso IsDBNull(selRow(ColumnTotalCost)) = False And sngUnitCost = 0 Then
                    sngTotalCost = ParseSingle(selRow(ColumnTotalCost))

                    If sngTotalCost <> 0 Then
                        objBudgetItem.UnitCost = New Currency(sngTotalCost, strCurrency)
                    End If
                End If
                If String.IsNullOrEmpty(objBudgetItem.Text) = False Then
                    objBudgetItem.SetTotalCost()
                    objBudgetYear.BudgetItems.Add(objBudgetItem)

                    objLastBudgetItem = objBudgetItem
                End If

                sngDuration = 0
                sngNumber = 0
                sngUnitCost = 0
                sngTotalCost = 0
            Next
        End If

        objBudgetYear.TotalCost = objBudgetYear.GetTotalCost()
    End Sub

    Private Function GetRowCount() As Integer
        Dim intRowCount As Integer

        If intFirstRowIndex >= 0 And intLastRowIndex >= intFirstRowIndex Then intRowCount = intLastRowIndex - intFirstRowIndex + 1

        Return intRowCount
    End Function
#End Region

End Class
