Public Class DataGridViewExchangeRates
    Inherits DataGridView

    Public Event ExchangeRateUpdated()

    Private objExchangeRates As ExchangeRates
    Private strDefaultCurrencyCode As String
    Private objCellLocation As New Point
    Private boolExchangeRateUpdated As Boolean

    Private EditRow As ExchangeRate = Nothing
    Private EditRowFlag As Integer = -1
    Private rowScopeCommit As Boolean = False

#Region "Properties"
    Public Property ExchangeRates As ExchangeRates
        Get
            Return objExchangeRates
        End Get
        Set(ByVal value As ExchangeRates)
            objExchangeRates = value
        End Set
    End Property

    Public Property DefaultCurrencyCode() As String
        Get
            Return strDefaultCurrencyCode
        End Get
        Set(ByVal value As String)
            strDefaultCurrencyCode = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub New()
        Initialise()

    End Sub

    Public Sub New(ByVal exchangerates As ExchangeRates, ByVal defaultcurrencycode As String)
        Initialise()

        Me.ExchangeRates = exchangerates
        Me.DefaultCurrencyCode = defaultcurrencycode

        'Reload()
    End Sub

    Private Sub Initialise()
        'datagridview settings
        VirtualMode = True
        AutoGenerateColumns = False
        AllowUserToAddRows = True
        AllowUserToResizeColumns = True
        AllowUserToResizeRows = True

        ShowCellToolTips = False
        BackgroundColor = Color.White
        DefaultCellStyle.Padding = New Padding(2)
        AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        RowHeadersVisible = True
        RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        With RowHeadersDefaultCellStyle
            .WrapMode = DataGridViewTriState.True
        End With

        With ColumnHeadersDefaultCellStyle
            .Font = New Font(DefaultFont, FontStyle.Bold)
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .WrapMode = DataGridViewTriState.True
        End With
        ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
    End Sub

    Public Sub LoadColumns()
        Dim lstCurrencyCodes As List(Of IdValuePair) = LoadCurrencyCodesList()
        Me.Columns.Clear()

        With DefaultCellStyle
            .Alignment = DataGridViewContentAlignment.TopLeft
            .WrapMode = DataGridViewTriState.True
        End With

        Dim colCurrencyCode As New DataGridViewComboBoxColumn
        With colCurrencyCode
            .Name = "CurrencyCode"
            .HeaderText = LANG_Currency
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .AutoComplete = True
            .DataSource = lstCurrencyCodes
            .ValueMember = "Id"
            .DisplayMember = "Value"
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .DropDownWidth = 200
        End With
        Me.Columns.Add(colCurrencyCode)

        Dim colExchangeRate As New DataGridViewTextBoxColumn
        With colExchangeRate
            .Name = "ExchangeRate"
            .HeaderText = LANG_ExchangeRate
            '.ValueType = GetType(Double)
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        End With
        Me.Columns.Add(colExchangeRate)

        Dim colConversion1 As New DataGridViewTextBoxColumn
        With colConversion1
            .Name = "Conversion1"
            .HeaderText = LANG_Conversion
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        End With
        Me.Columns.Add(colConversion1)

        Dim colConversion2 As New DataGridViewTextBoxColumn
        With colConversion2
            .Name = "Conversion2"
            .HeaderText = LANG_Conversion
            .FillWeight = 1
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        End With
        Me.Columns.Add(colConversion2)
    End Sub

    Public Sub Reload()
        'remember current cell location
        objCellLocation = CurrentCellAddress
        If objCellLocation.X < 0 Then objCellLocation.X = 0
        If objCellLocation.Y < 0 Then objCellLocation.Y = 0

        '(re)load target columns and rows
        Me.SuspendLayout()

        'Rows.Clear()
        LoadColumns()
        RowCount = ExchangeRates.Count + 1

        Me.Invalidate()
        Me.ResumeLayout()

        'set current cell location to what it was before
        If objCellLocation.X <= Me.ColumnCount - 1 And objCellLocation.Y <= Me.RowCount - 1 Then
            If Me(objCellLocation.X, objCellLocation.Y).Displayed = True Then _
                CurrentCell = Me(objCellLocation.X, objCellLocation.Y)
        End If
    End Sub
#End Region

#Region "Virtual mode"
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As ExchangeRate = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name
        Dim strConversion As String
        Dim sngRate As Single

        If e.RowIndex >= RowCount - 1 Then Return

        ' Store a reference to the planning grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.ExchangeRates(e.RowIndex), ExchangeRate)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "CurrencyCode"
                e.Value = RowTmp.CurrencyCode
            Case "ExchangeRate"
                e.Value = RowTmp.ExchangeRate.ToString("N4")
            Case "Conversion1"
                If RowTmp.ExchangeRate <> 0 Then
                    sngRate = 1 / RowTmp.ExchangeRate
                    strConversion = String.Format("1 {0} = {1} {2}", RowTmp.CurrencyCode, sngRate.ToString("N4"), DefaultCurrencyCode)
                    e.Value = strConversion
                End If
            Case "Conversion2"
                If RowTmp.ExchangeRate <> 0 Then
                    sngRate = RowTmp.ExchangeRate
                    strConversion = String.Format("1 {0} = {1} {2}", DefaultCurrencyCode, sngRate.ToString("N4"), RowTmp.CurrencyCode)
                    e.Value = strConversion
                End If
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As ExchangeRate = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name
        Dim strCellValue As String = String.Empty
        Dim selRowIndex As Integer = e.RowIndex

        If selRowIndex < ExchangeRates.Count Then
            'If the user is editing a new row, create a new planning grid row object.
            If Me.EditRow Is Nothing Then
                Dim selExchangeRate As ExchangeRate = CType(ExchangeRates(selRowIndex), ExchangeRate)

                Me.EditRow = New ExchangeRate
                With EditRow
                    .CurrencyCode = selExchangeRate.CurrencyCode
                    .ExchangeRate = selExchangeRate.ExchangeRate
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        Select Case strColName
            Case "CurrencyCode"
                RowTmp.CurrencyCode = e.Value
            Case "ExchangeRate"
                RowTmp.ExchangeRate = ParseSingle(e.Value)
                boolExchangeRateUpdated = True
        End Select
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = ExchangeRates.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New ExchangeRate
        Else
            ' If the user has canceled the edit of an existing row, 
            ' release the corresponding logframe row.
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        End If
        'Me.Reload()
        MyBase.OnCancelRowEdit(e)

    End Sub

    Protected Overrides Sub OnNewRowNeeded(ByVal e As System.Windows.Forms.DataGridViewRowEventArgs)
        Me.EditRow = New ExchangeRate()
        Me.EditRowFlag = Me.Rows.Count - 1
    End Sub

    Protected Overrides Sub OnRowDirtyStateNeeded(ByVal e As System.Windows.Forms.QuestionEventArgs)
        If Not rowScopeCommit Then

            ' In cell-level commit scope, indicate whether the value
            ' of the current cell has been modified.
            e.Response = Me.IsCurrentCellDirty

        End If
    End Sub

    Protected Overrides Sub OnRowValidated(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ' Save row changes if any were made and release the edited 
        ' grid row if there is one.
        If e.RowIndex >= ExchangeRates.Count AndAlso e.RowIndex <> Me.Rows.Count - 1 Then

            ' Add the new planning grid row to grid.
            ExchangeRates.Add(Me.EditRow)

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            'Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < ExchangeRates.Count Then
            ' Save the modified planning grid row in grid.
            Dim selExchangeRate As ExchangeRate = ExchangeRates(e.RowIndex)
            selExchangeRate.CurrencyCode = Me.EditRow.CurrencyCode
            selExchangeRate.ExchangeRate = Me.EditRow.ExchangeRate
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            'Me.Reload()
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        Invalidate()

        If boolExchangeRateUpdated = True Then
            RaiseEvent ExchangeRateUpdated()
            boolExchangeRateUpdated = False
        End If
        'MyBase.OnRowValidated(e)
    End Sub

    Protected Overrides Sub OnUserDeletingRow(ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        If e.Row.Index < Me.ExchangeRates.Count Then

            ' If the user has deleted an existing row, remove the  
            ' corresponding exchange rate from the data store. 
            Me.ExchangeRates.RemoveAt(e.Row.Index)

        End If

        If e.Row.Index = EditRowFlag Then

            ' If the user has deleted a newly created row, release 
            ' the corresponding exchange rate.  
            Me.EditRowFlag = -1
            Me.EditRow = Nothing

        End If
    End Sub
#End Region 'Virtual mode
End Class
