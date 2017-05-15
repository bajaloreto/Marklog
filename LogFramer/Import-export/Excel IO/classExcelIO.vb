Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class ExcelIO
    Implements System.IDisposable

    Private objExcelApplication As Object
    Private objWorkBook As Object

    Private oldCI As System.Globalization.CultureInfo

    Private conStrs As New Dictionary(Of String, String)
    Private connExcel As New OleDbConnection
    Private strFilePath As String
    Private objExportPmfRows As New PrintPmfRows

    Private xlReferenceStyle_R1C1 As Integer = -4150
    Private xlReferenceStyle_A1 As Integer = 1

    Private xlDeleteShiftDirection_ShiftToLeft As Integer = -4159

    Private xlVerticalAlignment_Top As Integer = -4160
    Private xlVerticalAlignment_Center As Integer = -4108 'Excel.XlVAlign.xlVAlignCenter
    Private xlHorizontalAlignment_Left As Integer = -4131
    Private xlHorizontalAlignment_Center As Integer = -4108
    Private xlHorizontalAlignment_Right As Integer = -4152

    Private xlPattern_Solid As Integer = 1
    Private xlColorIndex_ColorIndexAutomatic As Integer = -4105

    Private xlBorders_EdgeLeft As Integer = 7
    Private xlBorders_EdgeTop As Integer = 8
    Private xlBorders_EdgeBottom As Integer = 9
    Private xlBorders_EdgeRight As Integer = 10
    Private xlBorders_InsideVertical As Integer = 11
    Private xlBorders_InsideHorizontal As Integer = 12
    Private xlBorderWeight_Thin As Integer = 2
    Private xlBorderWeight_Medium As Integer = -4138 'Excel.XlBorderWeight.xlMedium

    Private xlLineStyle_Continuous As Integer = 1
    Private xlLineStyle_None As Integer = -4142

    Private xlThemeColorAccent6 As Integer = 10 'Excel.XlThemeColor.xlThemeColorAccent6
    Private xlThemeColorLight1 As Integer = 2 'Excel.XlThemeColor.xlThemeColorLight1
    Private xlDVType_ValidateList As Integer = 3 'Excel.XlDVType.xlValidateList
    Private xlDVAlertStyle_Stop As Integer = 1
    Private xlFormatConditionOperator_Between As Integer = 1

    Public Sub New()
        Dim AceRegistryKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.ClassesRoot.OpenSubKey("Microsoft.ACE.OLEDB.12.0")
        If AceRegistryKey IsNot Nothing Then
            conStrs("xlsxace") = "Provider='Microsoft.ACE.OLEDB.12.0'; Data Source='{0}'; " & "Extended Properties=""Excel 12.0;HDR=NO;IMEX=1"";"
            conStrs("xlsace") = "Provider='Microsoft.ACE.OLEDB.12.0'; Data Source='{0}'; " & "Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"";"
        End If

        Dim JetRegistryKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.ClassesRoot.OpenSubKey("Microsoft.JET.OLEDB.4.0")
        If JetRegistryKey IsNot Nothing Then
            conStrs("xlsjet") = "Provider='Microsoft.JET.OLEDB.4.0'; Data Source='{0}'; " & "Extended Properties=""Excel 8.0;HDR=NO;IMEX=1"";"
        End If

    End Sub

    Private Function GetColumnName(ByVal intExcelColumnIndex As Integer) As String
        intExcelColumnIndex -= 1
        Dim strColumn As String = "A"
        Dim strFirst As String = String.Empty
        Dim strSecond As String = String.Empty
        Dim intIndexFirst As Integer
        Dim intIndexSecond As Integer = intExcelColumnIndex

        If intExcelColumnIndex >= 26 Then
            intIndexFirst = Math.Floor(intExcelColumnIndex / 26) - 1
            strFirst = Chr(intIndexFirst + 65)

            Math.DivRem(intExcelColumnIndex, 26, intIndexSecond)
        End If
        strSecond = Chr(intIndexSecond + 65)

        strColumn = strFirst & strSecond

        Return strColumn
    End Function

    Private Function GetColumnRange(ByVal intColumnIndex As Integer, Optional ByVal intLastColumnIndex As Integer = -1) As String
        Dim strSortName As String = GetColumnName(intColumnIndex)
        Dim strFinalSortName As String

        If intLastColumnIndex >= 1 Then
            strFinalSortName = GetColumnName(intLastColumnIndex)
        Else
            strFinalSortName = strSortName
        End If
        Dim strColumnRange As String = String.Format("{0}:{1}", strSortName, strFinalSortName)

        Return strColumnRange
    End Function

#Region "Properties"
    Public Property ExcelApplication As Object
        Get
            Return objExcelApplication
        End Get
        Set(ByVal value As Object)
            objExcelApplication = value
        End Set
    End Property

    Public Property WorkBook As Object
        Get
            Return objWorkBook
        End Get
        Set(ByVal value As Object)
            objWorkBook = value
        End Set
    End Property
#End Region

#Region "Connect to Excel file"
    Public Property FilePath() As String
        Get
            Return strFilePath
        End Get
        Set(ByVal value As String)
            strFilePath = value
        End Set
    End Property

    Private Function OpenExcelConnection() As Boolean
        If Me.FilePath = "" Then Exit Function
        Dim strConnect As String = GetConnectionString()
        Dim boolOpened As Boolean

        If strConnect <> String.Empty Then
            With connExcel
                If .State = ConnectionState.Open Then
                    If .ConnectionString <> strConnect Then .Close()
                End If
                If .State = ConnectionState.Closed Then
                    .ConnectionString = strConnect
                    .Open()
                End If
            End With
            boolOpened = True
        End If

        Return boolOpened
    End Function

    Private Function GetConnectionString() As String
        Dim conStr As String = String.Empty
        Dim filExt As String = Path.GetExtension(Me.FilePath)
        filExt = filExt.Substring(1)
        If conStrs.ContainsKey(filExt & "ace") Then
            conStr = conStrs(filExt & "ace")
        ElseIf conStrs.ContainsKey(filExt & "jet") Then
            conStr = conStrs(filExt & "jet")
        Else
            Dim strMsg, strTitle As String
            Select Case UserLanguage
                Case "fr"
                    strMsg = "Erreur de connexion au fichier MS Excel" & vbLf & vbLf & _
                        "Les pilotes nécessaires (Microsoft Jet et/ou Microsoft ACE) ne sont pas installées sur votre machine."
                    strTitle = "Erreur serveur base de données"
                Case Else
                    strMsg = "Error Connecting to the MS Excel file" & vbLf & vbLf & _
                        "The necessary drivers (Microsoft Jet and/or Microsoft ACE) are not installed on your computer."

                    strTitle = "Error Database Server"
            End Select

            MessageBox.Show(strMsg, strTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        If conStr <> String.Empty Then
            Return String.Format(conStr, FilePath)
        Else
            Return String.Empty
        End If
    End Function

    Private Sub CloseExcelConnection()
        Try
            connExcel.Close()
        Catch ex As Exception

        End Try
    End Sub

    Public Function GetDataFromWorkSheet(ByVal strSheetName As String) As System.Data.DataTable
        Dim strTableName As String = strSheetName
        strSheetName = Trim(strSheetName) & "$"
        If strSheetName.Contains(" ") Then
            strSheetName = "'" & strSheetName & "'"
        End If
        Dim sql As String = "SELECT * FROM [" & strSheetName & "]"
        If OpenExcelConnection() = True Then

            Dim daExcel As New OleDbDataAdapter(sql, connExcel)
            'Dim dsExcel As New DataSet()
            Dim dtExcelSheet As New System.Data.DataTable(strTableName)
            Try
                daExcel.Fill(dtExcelSheet)
            Catch ex As OleDbException
                Dim strMsg, strTitle As String
                Select Case UserLanguage
                    Case "fr"
                        strMsg = "Il y a trop de colonnes dans la feuille de calcul (max.254)"
                        strTitle = "Impossible de montrer la feuille de calcul"
                    Case Else
                        strMsg = "Excel spreadsheet has too many columns (max.254)"
                        strTitle = "Can't show the spreadsheet"
                End Select
                MessageBox.Show(strMsg, strTitle, MessageBoxButtons.OK, _
                            MessageBoxIcon.Exclamation)
                MsgBox(strMsg, MsgBoxStyle.Exclamation, strTitle)
            End Try


            Return dtExcelSheet

            CloseExcelConnection()
        Else
            Return Nothing
        End If
    End Function

    Public Function GetExcelSheetNames() As String()
        Dim dt As System.Data.DataTable = Nothing

        'Try
        If OpenExcelConnection() = True Then
            ' Get the data table containg the schema guid.

            dt = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)

            If dt Is Nothing Then
                Return Nothing
            End If

            Dim excelSheets(dt.Rows.Count - 1) As String
            Dim i As Integer = 0
            Dim strName As String
            Dim strSplit(1) As String
            For Each row As DataRow In dt.Rows
                strName = row("TABLE_NAME").ToString()
                strName = strName.Replace("'", "")
                strSplit = strName.Split("$")
                If excelSheets.Contains(strSplit(0)) = False Then
                    excelSheets(i) = strSplit(0)
                    i += 1
                End If
            Next

            Return excelSheets
        Else
            Return Nothing
            'Catch ex As Exception
            '    Return Nothing
            'Finally
            '    ' Clean up.

            '    CloseExcelConnection()
            '    If dt IsNot Nothing Then
            '        dt.Dispose()
            '    End If
            'End Try
        End If
    End Function
#End Region 'Connect to Excel file

#Region "Initialise new workbook"
    Private Sub InitialiseWorkBook()
        ExcelApplication = CreateObject("Excel.Application") 'New Excel.Application
        ExcelApplication.Visible = True
        ExcelApplication.UserControl = False

        WorkBook = ExcelApplication.Workbooks.Add()

        'necessary to avoid bug when english version of office is installed
        'on non-english version of Windows
        oldCI = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
    End Sub

    Private Sub FinaliseWorkBook()
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
    End Sub
#End Region

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




