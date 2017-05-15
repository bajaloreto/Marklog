Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Public Class DialogImportLogframe
    Private FromWordDocument As Boolean
    Private WorkSheets As List(Of String)
    Private objExcelIO As New ExcelIO
    Private objWordIO As New WordIO
    Private strDocName, strDocPath As String
    Private objLogFrame As New LogFrame
    Private strColumnObjectives As String, strColumnIndicators As String
    Private strColumnVerificationSources As String, strColumnAssumptions As String
    Private intFirstRowGoals As Integer, intLastRowGoals As Integer
    Private intFirstRowPurposes As Integer, intLastRowPurposes As Integer
    Private intFirstRowOutputs As Integer, intLastRowOutputs As Integer
    Private intFirstRowActivities As Integer, intLastRowActivities As Integer
    Private intLastRowIndex As Integer

#Region "Properties"
    Public Property DocName As String
        Get
            Return strDocName
        End Get
        Set(value As String)
            strDocName = value
        End Set
    End Property

    Public Property DocPath As String
        Get
            Return strDocPath
        End Get
        Set(value As String)
            strDocPath = value
        End Set
    End Property

    Public Property LogFrame() As LogFrame
        Get
            Return objLogFrame
        End Get
        Set(ByVal value As LogFrame)
            objLogFrame = value
        End Set
    End Property

    Private Property ColumnObjectives() As String
        Get
            Return strColumnObjectives
        End Get
        Set(ByVal value As String)
            strColumnObjectives = value
        End Set
    End Property

    Private Property ColumnIndicators() As String
        Get
            Return strColumnIndicators
        End Get
        Set(ByVal value As String)
            strColumnIndicators = value
        End Set
    End Property

    Private Property ColumnVerificationSources() As String
        Get
            Return strColumnVerificationSources
        End Get
        Set(ByVal value As String)
            strColumnVerificationSources = value
        End Set
    End Property

    Private Property ColumnAssumptions() As String
        Get
            Return strColumnAssumptions
        End Get
        Set(ByVal value As String)
            strColumnAssumptions = value
        End Set
    End Property

    Private Property FirstRowGoals() As Integer
        Get
            Return intFirstRowGoals
        End Get
        Set(ByVal value As Integer)
            intFirstRowGoals = value
        End Set
    End Property

    Private Property LastRowGoals() As Integer
        Get
            Return intLastRowGoals
        End Get
        Set(ByVal value As Integer)
            intLastRowGoals = value
        End Set
    End Property

    Private Property FirstRowPurposes() As Integer
        Get
            Return intFirstRowPurposes
        End Get
        Set(ByVal value As Integer)
            intFirstRowPurposes = value
        End Set
    End Property

    Private Property LastRowPurposes() As Integer
        Get
            Return intLastRowPurposes
        End Get
        Set(ByVal value As Integer)
            intLastRowPurposes = value
        End Set
    End Property

    Private Property FirstRowOutputs() As Integer
        Get
            Return intFirstRowOutputs
        End Get
        Set(ByVal value As Integer)
            intFirstRowOutputs = value
        End Set
    End Property

    Private Property LastRowOutputs() As Integer
        Get
            Return intLastRowOutputs
        End Get
        Set(ByVal value As Integer)
            intLastRowOutputs = value
        End Set
    End Property

    Private Property FirstRowActivities() As Integer
        Get
            Return intFirstRowActivities
        End Get
        Set(ByVal value As Integer)
            intFirstRowActivities = value
        End Set
    End Property

    Private Property LastRowActivities() As Integer
        Get
            Return intLastRowActivities
        End Get
        Set(ByVal value As Integer)
            intLastRowActivities = value
        End Set
    End Property
#End Region

#Region "Methods and events"
    Public Sub New(ByVal boolFromWordDocument As Boolean)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.FromWordDocument = boolFromWordDocument

        Dim strTitle, strStep1, strStep2, strSelect As String

        If FromWordDocument = False Then
            strTitle = LANG_ImportExcelLogframeTitle
            strStep1 = LANG_ImportExcelLogframeStep1
            strStep2 = LANG_ImportExcelLogframeStep2
            strSelect = ToLabel(LANG_WorkSheet)
        Else
            strTitle = LANG_ImportWordLogframeTitle
            strStep1 = LANG_ImportWordLogframeStep1
            strStep2 = LANG_ImportWordLogframeStep2
            strSelect = ToLabel(LANG_Table)
        End If

        Me.Text = strTitle
        lblStep1.Text = strStep1
        lblStep2.Text = strStep2
        lblSelectWorksheet.Text = strSelect

        dgWorkSheet.Visible = False
        btnImport.Enabled = False

        With Me.cmbMeansBudget
            .Items.Add(LANG_ImportResourcesBudget)
            .Items.Add(LANG_ImportResourcesBudget)
            .SelectedIndex = 0
        End With

        ntbFirstRowActivities.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbLastRowActivities.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbFirstRowOutputs.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbLastRowOutputs.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbFirstRowPurposes.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbLastRowPurposes.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbFirstRowGoals.ValueType = NumericTextBox.ValueTypes.IntegerValue
        ntbLastRowGoals.ValueType = NumericTextBox.ValueTypes.IntegerValue

        ntbFirstRowActivities.NrDecimals = 0
        ntbLastRowActivities.NrDecimals = 0
        ntbFirstRowOutputs.NrDecimals = 0
        ntbLastRowOutputs.NrDecimals = 0
        ntbFirstRowPurposes.NrDecimals = 0
        ntbLastRowPurposes.NrDecimals = 0
        ntbFirstRowGoals.NrDecimals = 0
        ntbLastRowGoals.NrDecimals = 0

        ntbFirstRowActivities.SetDecimals = True
        ntbLastRowActivities.SetDecimals = True
        ntbFirstRowOutputs.SetDecimals = True
        ntbLastRowOutputs.SetDecimals = True
        ntbFirstRowPurposes.SetDecimals = True
        ntbLastRowPurposes.SetDecimals = True
        ntbFirstRowGoals.SetDecimals = True
        ntbLastRowGoals.SetDecimals = True
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

        If String.IsNullOrEmpty(ColumnObjectives) Then boolComplete = False
        If String.IsNullOrEmpty(ColumnVerificationSources) = False And String.IsNullOrEmpty(ColumnIndicators) Then boolComplete = False
        'If String.IsNullOrEmpty(ColumnIndicators) Then boolComplete = False
        'If String.IsNullOrEmpty(ColumnAssumptions) Then boolComplete = False
        If GoalRowsSet() = False And PurposeRowsSet() = False And OutputRowsSet() = False And ActivityRowsSet() = False Then boolComplete = False

        Return boolComplete
    End Function

    Private Sub btnLoadFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadFile.Click
        Dim DialogOpenWorkbook As New OpenFileDialog

        Dim strTitle, strFilter As String

        If FromWordDocument = False Then
            strTitle = LANG_ImportExcelLogframeSelectFile
            strFilter = LANG_ImportExcelFilter
        Else
            strTitle = LANG_ImportWordLogframeSelectFile
            strFilter = LANG_ImportWordLogframeFilter
        End If

        With DialogOpenWorkbook
            .Title = strTitle
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Multiselect = False
            .Filter = (strFilter)
            .FilterIndex = 0
        End With

        If DialogOpenWorkbook.ShowDialog() = DialogResult.OK And String.IsNullOrEmpty(DialogOpenWorkbook.FileName) = False Then
            Dim selFileInfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(DialogOpenWorkbook.FileName)
            DocName = selFileInfo.Name
            DocPath = selFileInfo.FullName

            Me.LogFrame.ProjectTitle = DocName.Replace(selFileInfo.Extension, String.Empty)
            tbFileName.Text = Me.DocName

            Dim objSheetNames As Object
            If FromWordDocument = False Then
                objExcelIO.FilePath = DialogOpenWorkbook.FileName
                objSheetNames = objExcelIO.GetExcelSheetNames
            Else
                objWordIO.FilePath = DialogOpenWorkbook.FileName
                objSheetNames = objWordIO.GetTableNames
            End If

            If objSheetNames IsNot Nothing Then
                Dim strSheetNames() As String = objSheetNames
                If strSheetNames.Length > 0 Then
                    cmbSelectWorksheet.DataSource = strSheetNames
                    cmbSelectWorksheet.SelectedIndex = 0
                    dgWorkSheet.Visible = True
                Else
                    dgWorkSheet.Visible = False
                End If
            Else
                dgWorkSheet.Visible = False
            End If
        Else
            dgWorkSheet.Visible = False
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

        If FromWordDocument = False Then
            dtWorkSheet = objExcelIO.GetDataFromWorkSheet(strWorkSheet)
        Else
            dtWorkSheet = objWordIO.GetTableContent(cmbSelectWorksheet.SelectedIndex + 1)
        End If

        If dtWorkSheet IsNot Nothing Then
            intLastRowIndex = dtWorkSheet.Rows.Count - 1
            dgWorkSheet.LoadColumns(dtWorkSheet)

            Dim lstColumns(dtWorkSheet.Columns.Count) As String
            Dim i As Integer
            lstColumns(0) = ""
            For i = 0 To dtWorkSheet.Columns.Count - 1
                lstColumns(i + 1) = dtWorkSheet.Columns(i).ColumnName
            Next

            cmbColumnObjectives.Items.Clear()
            cmbColumnObjectives.Items.AddRange(lstColumns.ToArray)
            cmbColumnObjectives.SelectedItem = "A"
            cmbColumnIndicators.Items.Clear()
            cmbColumnIndicators.Items.AddRange(lstColumns.ToArray)
            'cmbColumnIndicators.SelectedItem = "B"
            cmbColumnVerificationSources.Items.Clear()
            cmbColumnVerificationSources.Items.AddRange(lstColumns.ToArray)
            'cmbColumnVerificationMeans.SelectedItem = "C"
            cmbColumnAssumptions.Items.Clear()
            cmbColumnAssumptions.Items.AddRange(lstColumns.ToArray)
            'cmbColumnAssumptions.SelectedItem = "D"

        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnObjectives_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnObjectives.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnObjectives_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnObjectives.SelectedIndexChanged
        Dim strColumn As String = cmbColumnObjectives.SelectedItem.ToString

        Me.ColumnObjectives = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnIndicators_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnIndicators.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnIndicators_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnIndicators.SelectedIndexChanged
        Dim strColumn As String = cmbColumnIndicators.SelectedItem.ToString
        Me.ColumnIndicators = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnVerificationSources_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnVerificationSources.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnVerificationMeans_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnVerificationSources.SelectedIndexChanged
        Dim strColumn As String = cmbColumnVerificationSources.SelectedItem.ToString
        Me.ColumnVerificationSources = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub cmbColumnAssumptions_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbColumnAssumptions.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
        End If
    End Sub

    Private Sub cmbColumnAssumptions_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbColumnAssumptions.SelectedIndexChanged
        Dim strColumn As String = cmbColumnAssumptions.SelectedItem.ToString
        Me.ColumnAssumptions = strColumn
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbFirstRowGoals_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbFirstRowGoals.Validated
        If String.IsNullOrEmpty(ntbFirstRowGoals.Text) = False Then
            intFirstRowGoals = ntbFirstRowGoals.IntegerValue 'ntbFirstRowGoals.Text)
            intFirstRowGoals -= 1
            If intFirstRowGoals < 0 Then ntbFirstRowGoals.Text = 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbLastRowGoals_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbLastRowGoals.Validated
        If String.IsNullOrEmpty(ntbLastRowGoals.Text) = False Then
            intLastRowGoals = ntbLastRowGoals.IntegerValue
            intLastRowGoals -= 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbFirstRowPurposes_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbFirstRowPurposes.Validated
        If String.IsNullOrEmpty(ntbFirstRowPurposes.Text) = False Then
            intFirstRowPurposes = ntbFirstRowPurposes.IntegerValue
            intFirstRowPurposes -= 1
            If intFirstRowPurposes < 0 Then ntbFirstRowPurposes.Text = 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbLastRowPurposes_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbLastRowPurposes.Validated
        If String.IsNullOrEmpty(ntbLastRowPurposes.Text) = False Then
            intLastRowPurposes = ntbLastRowPurposes.IntegerValue
            intLastRowPurposes -= 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbFirstRowOutputs_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbFirstRowOutputs.Validated
        If String.IsNullOrEmpty(ntbFirstRowOutputs.Text) = False Then
            intFirstRowOutputs = ntbFirstRowOutputs.IntegerValue
            intFirstRowOutputs -= 1
            If intFirstRowOutputs < 0 Then ntbFirstRowOutputs.Text = 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbLastRowOutputs_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbLastRowOutputs.Validated
        If String.IsNullOrEmpty(ntbLastRowOutputs.Text) = False Then
            intLastRowOutputs = ntbLastRowOutputs.IntegerValue
            intLastRowOutputs -= 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbFirstRowActivities_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbFirstRowActivities.Validated
        If String.IsNullOrEmpty(ntbFirstRowActivities.Text) = False Then
            intFirstRowActivities = ntbFirstRowActivities.IntegerValue
            intFirstRowActivities -= 1
            If intFirstRowActivities < 0 Then ntbFirstRowActivities.Text = 1
            With ntbLastRowActivities
                '.SuspendLayout()
                .Text = (intLastRowIndex + 1).ToString
                '.ResumeLayout()
            End With
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Sub ntbLastRowActivities_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ntbLastRowActivities.Validated
        If String.IsNullOrEmpty(ntbLastRowActivities.Text) = False Then
            intLastRowActivities = ntbLastRowActivities.IntegerValue
            intLastRowActivities -= 1
        End If
        If CheckAllInfoComplete() = True Then btnImport.Enabled = True Else btnImport.Enabled = False
    End Sub

    Private Function GoalRowsSet() As Boolean
        Dim boolSet As Boolean
        If String.IsNullOrEmpty(ntbFirstRowGoals.Text) And String.IsNullOrEmpty(ntbLastRowGoals.Text) Then
            boolSet = False
        Else
            boolSet = True
        End If
        Return boolSet
    End Function

    Private Function PurposeRowsSet() As Boolean
        Dim boolSet As Boolean
        If String.IsNullOrEmpty(ntbFirstRowPurposes.Text) And String.IsNullOrEmpty(ntbLastRowPurposes.Text) Then
            boolSet = False
        Else
            boolSet = True
        End If
        Return boolSet
    End Function

    Private Function OutputRowsSet() As Boolean
        Dim boolSet As Boolean
        If String.IsNullOrEmpty(ntbFirstRowOutputs.Text) And String.IsNullOrEmpty(ntbLastRowOutputs.Text) Then
            boolSet = False
        Else
            boolSet = True
        End If
        Return boolSet
    End Function

    Private Function ActivityRowsSet() As Boolean
        Dim boolSet As Boolean
        If String.IsNullOrEmpty(ntbFirstRowActivities.Text) And String.IsNullOrEmpty(ntbLastRowActivities.Text) Then
            boolSet = False
        Else
            boolSet = True
        End If
        Return boolSet
    End Function
#End Region

#Region "Import"
    Public Function ImportLogFrame() As LogFrame
        Cursor.Current = Cursors.WaitCursor

        With Me.LogFrame
            .StartDate = Now.Date
            .EndDate = .StartDate.AddYears(1)
            .EndDate = .EndDate.AddDays(-1)
            .Duration = 1
            .DurationUnit = LogFrame.DurationUnits.Years
        End With

        'import goals
        If GoalRowsSet() = True Then ImportLogframe_Goals()

        'import purposes
        If PurposeRowsSet() = True Then ImportLogframe_Purposes()

        'import outputs
        If OutputRowsSet() = True Then ImportLogframe_Outputs()

        'import activities
        If ActivityRowsSet() = True Then ImportLogframe_Activities()

        Cursor.Current = Cursors.Default

        Return Me.LogFrame
    End Function

    Private Function GetRowCount() As Integer
        Dim intRowCount As Integer

        If GoalRowsSet() = True Then
            If intFirstRowGoals >= 0 And intLastRowGoals >= intFirstRowGoals Then intRowCount = intLastRowGoals - intFirstRowGoals + 1
        End If

        If PurposeRowsSet() = True Then
            If intFirstRowPurposes >= 0 And intLastRowPurposes >= intFirstRowPurposes Then intRowCount = intRowCount + intLastRowPurposes - intFirstRowPurposes + 1
        End If

        If OutputRowsSet() = True Then
            If intFirstRowOutputs >= 0 And intLastRowOutputs >= intFirstRowOutputs Then intRowCount = intRowCount + intLastRowOutputs - intFirstRowOutputs + 1
        End If

        If ActivityRowsSet() = True Then
            If intFirstRowActivities >= 0 And intLastRowActivities >= intFirstRowActivities Then intRowCount = intRowCount + intLastRowActivities - intFirstRowActivities + 1
        End If

        Return intRowCount
    End Function

    Public Sub ImportLogframe_Goals()
        Dim selRow As DataRow
        Dim LastGoal As Goal = Nothing
        Dim LastIndicator As Indicator = Nothing
        Dim i As Integer

        If intFirstRowGoals >= 0 And intLastRowGoals >= intFirstRowGoals Then
            For i = intFirstRowGoals To intLastRowGoals
                selRow = dgWorkSheet.WorkSheet.Rows(i)

                If String.IsNullOrEmpty(ColumnObjectives) = False AndAlso IsDBNull(selRow(ColumnObjectives)) = False Then
                    Dim objGoal As New Goal
                    objGoal.SetText(selRow(ColumnObjectives).ToString)
                    LastGoal = objGoal
                    Me.LogFrame.Goals.Add(objGoal)
                End If
                If String.IsNullOrEmpty(ColumnIndicators) = False AndAlso IsDBNull(selRow(ColumnIndicators)) = False Then
                    Dim objIndicator As New Indicator
                    objIndicator.SetText(selRow(ColumnIndicators).ToString)
                    LastIndicator = objIndicator
                    If LastGoal IsNot Nothing Then _
                        LastGoal.Indicators.Add(objIndicator)
                End If
                If String.IsNullOrEmpty(ColumnVerificationSources) = False AndAlso IsDBNull(selRow(ColumnVerificationSources)) = False Then
                    Dim objVerificationSource As New VerificationSource
                    objVerificationSource.SetText(selRow(ColumnVerificationSources).ToString)
                    If LastIndicator IsNot Nothing Then _
                        LastIndicator.VerificationSources.Add(objVerificationSource)
                End If
                If String.IsNullOrEmpty(ColumnAssumptions) = False AndAlso IsDBNull(selRow(ColumnAssumptions)) = False Then
                    Dim objAssumption As New Assumption
                    objAssumption.SetText(selRow(ColumnAssumptions).ToString)
                    If LastGoal IsNot Nothing Then _
                        LastGoal.Assumptions.Add(objAssumption)
                End If
            Next
        End If
    End Sub

    Public Sub ImportLogframe_Purposes()
        Dim selRow As DataRow
        Dim LastPurpose As Purpose = Nothing
        Dim LastIndicator As Indicator = Nothing
        Dim i As Integer

        If intFirstRowPurposes >= 0 And intLastRowPurposes >= intFirstRowPurposes Then

            For i = intFirstRowPurposes To intLastRowPurposes
                selRow = dgWorkSheet.WorkSheet.Rows(i)

                If IsDBNull(selRow(ColumnObjectives)) = False Then
                    Dim objPurpose As New Purpose
                    objPurpose.SetText(selRow(ColumnObjectives).ToString)
                    LastPurpose = objPurpose
                    Me.LogFrame.Purposes.Add(objPurpose)
                End If
                If IsDBNull(selRow(ColumnIndicators)) = False Then
                    Dim objIndicator As New Indicator
                    objIndicator.SetText(selRow(ColumnIndicators).ToString)
                    LastIndicator = objIndicator
                    If LastPurpose IsNot Nothing Then _
                        LastPurpose.Indicators.Add(objIndicator)
                End If
                If IsDBNull(selRow(ColumnVerificationSources)) = False Then
                    Dim objVerificationSource As New VerificationSource
                    objVerificationSource.SetText(selRow(ColumnVerificationSources).ToString)
                    If LastIndicator IsNot Nothing Then _
                        LastIndicator.VerificationSources.Add(objVerificationSource)
                End If
                If IsDBNull(selRow(ColumnAssumptions)) = False Then
                    Dim objAssumption As New Assumption
                    objAssumption.SetText(selRow(ColumnAssumptions).ToString)
                    If LastPurpose IsNot Nothing Then _
                        LastPurpose.Assumptions.Add(objAssumption)
                End If
            Next
        End If
    End Sub

    Private Sub ImportLogframe_Outputs()
        Dim selRow As DataRow
        Dim LastOutput As Output = Nothing
        Dim LastIndicator As Indicator = Nothing
        Dim i As Integer

        If intFirstRowOutputs >= 0 And intLastRowOutputs >= intFirstRowOutputs Then

            For i = intFirstRowOutputs To intLastRowOutputs
                selRow = dgWorkSheet.WorkSheet.Rows(i)

                If IsDBNull(selRow(ColumnObjectives)) = False Then
                    'intCount += 1
                    Dim objOutput As New Output
                    objOutput.SetText(selRow(ColumnObjectives).ToString)
                    LastOutput = objOutput
                    If Me.LogFrame.Purposes.Count = 0 Then
                        Me.LogFrame.Purposes.Add(New Purpose)
                    End If
                    Me.LogFrame.Purposes(0).Outputs.Add(objOutput)
                End If
                If IsDBNull(selRow(ColumnIndicators)) = False Then
                    Dim objIndicator As New Indicator
                    objIndicator.SetText(selRow(ColumnIndicators).ToString)
                    LastIndicator = objIndicator
                    If LastOutput IsNot Nothing Then _
                        LastOutput.Indicators.Add(objIndicator)
                End If
                If IsDBNull(selRow(ColumnVerificationSources)) = False Then
                    Dim objVerificationSource As New VerificationSource
                    objVerificationSource.SetText(selRow(ColumnVerificationSources).ToString)
                    If LastIndicator IsNot Nothing Then _
                        LastIndicator.VerificationSources.Add(objVerificationSource)
                End If
                If IsDBNull(selRow(ColumnAssumptions)) = False Then
                    Dim objAssumption As New Assumption
                    objAssumption.SetText(selRow(ColumnAssumptions).ToString)
                    If LastOutput IsNot Nothing Then _
                        LastOutput.Assumptions.Add(objAssumption)
                End If
            Next
        End If
    End Sub

    Private Sub ImportLogframe_Activities()
        Dim selRow As DataRow
        Dim LastActivity As Activity = Nothing
        Dim LastIndicator As Indicator = Nothing
        Dim LastResource As Resource = Nothing
        Dim i As Integer

        If intFirstRowActivities >= 0 And intLastRowActivities >= intFirstRowActivities Then

            For i = intFirstRowActivities To intLastRowActivities
                selRow = dgWorkSheet.WorkSheet.Rows(i)

                If IsDBNull(selRow(ColumnObjectives)) = False Then
                    Dim objActivity As New Activity
                    objActivity.SetText(selRow(ColumnObjectives).ToString)
                    LastActivity = objActivity
                    If Me.LogFrame.Purposes.Count = 0 Then
                        Me.LogFrame.Purposes.Add(New Purpose)
                    End If
                    If Me.LogFrame.Purposes(0).Outputs.Count = 0 Then
                        Me.LogFrame.Purposes(0).Outputs.Add(New Output)
                    End If
                    Me.LogFrame.Purposes(0).Outputs(0).Activities.Add(objActivity)
                End If
                If Me.cmbMeansBudget.SelectedIndex = 0 Then
                    'add resources and budgetitems
                    If IsDBNull(selRow(ColumnIndicators)) = False Then
                        Dim objResource As New Resource
                        objResource.SetText(selRow(ColumnIndicators).ToString)
                        LastResource = objResource
                        If LastActivity IsNot Nothing Then _
                            LastActivity.Resources.Add(objResource)
                    End If
                    If IsDBNull(selRow(ColumnVerificationSources)) = False Then
                        'Dim strBudget As String = selRow(ColumnVerificationSources).ToString
                        'strBudget = Regex.Replace(strBudget, "[^.0-9]", "")
                        'Dim sngBudget As Single

                        'If Single.TryParse(strBudget, sngBudget) Then
                        '    LastResource.BudgetValue = sngBudget
                        'Else
                        '    Dim objBudgetItem As New BudgetItem

                        '    objBudgetItem.SetText(selRow(ColumnVerificationSources).ToString)
                        '    If LastResource IsNot Nothing Then _
                        '        LastResource.BudgetItems.Add(objBudgetItem)
                        'End If
                    End If
                Else
                    If IsDBNull(selRow(ColumnIndicators)) = False Then
                        Dim objIndicator As New Indicator
                        objIndicator.SetText(selRow(ColumnIndicators).ToString)
                        LastIndicator = objIndicator
                        If LastActivity IsNot Nothing Then _
                            LastActivity.Indicators.Add(objIndicator)
                    End If
                    If IsDBNull(selRow(ColumnVerificationSources)) = False Then
                        Dim objVerificationSource As New VerificationSource
                        objVerificationSource.SetText(selRow(ColumnVerificationSources).ToString)
                        If LastIndicator IsNot Nothing Then _
                            LastIndicator.VerificationSources.Add(objVerificationSource)
                    End If
                End If

                If IsDBNull(selRow(ColumnAssumptions)) = False Then
                    Dim objAssumption As New Assumption
                    objAssumption.SetText(selRow(ColumnAssumptions).ToString)
                    If LastActivity IsNot Nothing Then _
                        LastActivity.Assumptions.Add(objAssumption)
                End If
            Next
        End If
    End Sub
#End Region



    
    
End Class
