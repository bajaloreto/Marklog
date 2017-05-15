Imports System.IO
Imports Word = Microsoft.Office.Interop.Word

Public Class WordIO
    Implements System.IDisposable

    Private strFilePath As String
    Private fntNormal As Object 'Word.Font
    Private LastTable As Object 'Object
    Private objExportPmfRows As New PrintPmfRows

    Private wdPageSetup_Orientation_Landscape As Integer = 1 'wdPageSetup_Orientation_Landscape
    Private wdPageSetup_Orientation_Portrait As Integer = 0 'wdPageSetup_Orientation_Portrait

    Private wdInsertBreak_SectionBreakNextPage As Integer = 2 'InsertBreak_SectionBreakNextPage
    Private wdUnits_Paragraph As Integer = 4
    Private wdUnits_Line As Integer = 5 'wdUnits_Line
    Private wdUnits_Story As Integer = 6 'wdUnits_Story

    Private wdLineSpaceSingle As Integer = 0 'WdLineSpacing.wdLineSpaceSingle
    Private wdStyleNormal As Integer = -1 'wdStyleNormal
    Private wdStyleParagraph As Integer = 1 'wdStyleParagraph
    Private wdStyleCharacter As Integer = 2 'wdStyleCharacter
    Private wdAlignParagraphLeft As Integer = 0 'WdParagraphAlignment.wdAlignParagraphLeft
    Private wdAlignParagraphCenter As Integer = 1 'WdParagraphAlignment.wdAlignParagraphCenter
    Private wdAlignParagraphRight As Integer = 2 'WdParagraphAlignment.wdAlignParagraphRight
    Private wdUnderline_None As Integer = 0 'wdUnderline_None
    Private wdUnderline_Dotted As Integer = 4 'wdUnderline_Dotted

    Private wdColorAutomatic As Integer = -16777216 'wdColor_Automatic 13421772
    Private wdColor_Gray20 As Integer = 13421772 'wdColor_Gray20
    Private wdColor_Gray25 As Integer = 12632256 'wdColor_Gray25 9211020
    Private wdColor_Gray45 As Integer = 9211020 'wdColor_Gray45
    Private wdColor_Gray70 As Integer = 5000268 'wdColor_Gray70 16764057
    Private wdColor_Red As Integer = 255 'Word.WdColor.wdColorRed
    Private wdColor_Green As Integer = 32768 'Word.WdColor.wdColorGreen
    Private wdColor_Blue As Integer = 16711680 'Word.WdColor.wdColorBlue
    Private wdColor_Yellow As Integer = 65535 'Word.WdColor.wdColorYellow
    Private wdColor_Orange As Integer = 26367 'Word.WdColor.wdColorOrange
    Private wdColor_DarkRed As Integer = 128 'Word.WdColor.wdColorDarkRed
    Private wdColor_DarkGreen As Integer = 13056 'Word.WdColor.wdColorDarkGreen
    Private wdColor_DarkBlue As Integer = 8388608 'Word.WdColor.wdColorDarkBlue
    Private wdColor_DarkYellow As Integer = 32896 'Word.WdColor.wdColorDarkYellow
    Private wdColor_DarkTeal As Integer = 6697728 'Word.WdColor.wdColorDarkTeal
    Private wdColor_LightTurquoise As Integer = 16777164 'Word.WdColor.wdColorLightTurquoise
    Private wdColor_LightGreen As Integer = 13434828 'Word.WdColor.wdColorLightGreen
    Private wdColor_LightBlue As Integer = 16737843 'Word.WdColor.wdColorLightBlue
    Private wdColor_LightYellow As Integer = 10092543 'Word.WdColor.wdColorLightYellow
    Private wdColor_LightOrange As Integer = 39423 'Word.WdColor.wdColorLightOrange
    Private wdColor_PaleBlue As Integer = 16764057 'Word.wdColor_PaleBlue

    Private wdLineStyleNone As Integer = 0 'wdLineStyleNone
    Private wdLineStyleSingle As Integer = 1 'wdLineStyle_Single
    Private wdLineStyleDot As Integer = 2 'wdLineStyle_Dot
    Private wdLineWidth025pt As Integer = 2 'WdLineWidth.wdLineWidth025pt
    Private wdLineWidth050pt As Integer = 4 'WdLineWidth.wdLineWidth050pt
    Private wdLineWidth100pt As Integer = 8 'WdLineWidth.wdLineWidt100pt
    Private wdLineWidth150pt As Integer = 12 'WdLineWidth.wdLineWidth150pt

    Private wdTableBehaviour9 As Integer = 1 'wdTableBehaviour9
    Private wdAlignRowLeft As Integer = 0 'WdRowAlignment.wdAlignRowLeft
    Private wdAlignRowCenter As Integer = 1 'WdRowAlignment.wdAlignRowCenter
    Private wdAlignRowRight As Integer = 2 'WdRowAlignment.wdAlignRowRight
    Private wdAutofitBehaviour_FitContent As Integer = 1 'wdAutofitBehaviour_FitContent
    Private wdAutofitBehaviour_Fixed As Integer = 0 'wdAutofitBehaviour_Fixed
    Private wdBorderVertical As Integer = -6 'WdBorderType.wdBorderVertical
    Private wdBorderHorizontal As Integer = -5 'WdBorderType.wdBorderHorizontal
    Private wdBorderRight As Integer = -4 'WdBorderType.wdBorderRight
    Private wdBorderBottom As Integer = -3 'WdBorderType.wdBorderBottom
    Private wdBorderLeft As Integer = -2 'WdBorderType.wdBorderLeft
    Private wdBorderTop As Integer = -1 'WdBorderType.wdBorderTop
    Private wdCollapseEnd As Integer = 0 'WdCollapseDirection.wdCollapseEnd
    Private wdEvenRowBanding As Integer = 3 'WdConditionCode.wdEvenRowBanding
    Private wdFirstRow As Integer = 0 'WdConditionCode.wdFirstRow
    Private wdTableFormatGrid1 As Integer = 16 'WdTableFormat.wdTableFormatGrid1
    Private wdOddRowBanding As Integer = 2 'WdConditionCode.wdOddRowBanding
    Private wdOutlineLevelBodyText As Integer = 10 'WdOutlineLevel.wdOutlineLevelBodyText
    Private wdPreferredWidthType_Points As Integer = 3 'wdPreferredWidthType_Points
    Private wdSeparateByDefaultListSeparator As Integer = 3 'Word.WdTableFieldSeparator.wdSeparateByDefaultListSeparator
    Private wdSeparateByTabs As Integer = 1 'wdSeparateByTabs
    Private wdStyleTypeTable As Integer = 3 'WdStyleType.wdStyleTypeTable
    Private wdTextureNone As Integer = 0 'WdTextureIndex.wdTextureNone
    Private wdPreferredWidthPercent As Integer = 2 'Word.WdPreferredWidthType.wdPreferredWidthPercent

    Public Enum InsertLocations As Integer
        AtBeginning = 0
        AtBookmark = 1
        AtEnd = 2
    End Enum

    Public Property FilePath() As String
        Get
            Return strFilePath
        End Get
        Set(ByVal value As String)
            strFilePath = value
        End Set
    End Property

#Region "Create or open document"
    Public Function GetDocumentName() As String
        Dim DialogOpenDocument As New OpenFileDialog
        Dim strTitle, strWordDoc, strAllDocs As String
        Dim strFilter As String

        Select Case UserLanguage
            Case "fr"
                strTitle = "Sélectionnez un document MS Word"
                strWordDoc = "Document MS Word"
                strAllDocs = "Tous les fichiers"
            Case Else
                strTitle = "Select a MS Word document"
                strWordDoc = "MS Word document"
                strAllDocs = "All files"
        End Select
        strFilter = strWordDoc & " (*.doc, *.docx)|*.doc; *.docx|" & strAllDocs & " (*.*)|*.*"
        With DialogOpenDocument
            .Title = strTitle
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Multiselect = False
            .Filter = (strFilter)
            .FilterIndex = 0
        End With

        If DialogOpenDocument.ShowDialog() = DialogResult.OK And DialogOpenDocument.FileName <> "" Then
            Return DialogOpenDocument.FileName
        Else
            Return String.Empty
        End If
    End Function

    Public Function GetBookMarksList(ByVal strFilePath As String) As ArrayList
        Dim BookMarksList As New ArrayList

        If File.Exists(strFilePath) Then
            Dim MsWord As Object = CreateObject("Word.Application") 'New Microsoft.Office.Interop.Word.Application
            MsWord.Visible = False
            Dim objDoc As Object 'Object 
            objDoc = MsWord.Documents.Open(strFilePath, False, True, False)
            If objDoc.Bookmarks.Count > 0 Then
                For i = 1 To objDoc.Bookmarks.Count
                    BookMarksList.Add(objDoc.Bookmarks(i).Name)
                Next
            End If
            objDoc.Close(False)
            MsWord.Quit()
            objDoc = Nothing
            MsWord = Nothing
        End If
        Return BookMarksList
    End Function

    Public Function GetTableNames() As String()
        Dim TablesList As New ArrayList

        If File.Exists(Me.FilePath) Then
            Dim MsWord As Object = CreateObject("Word.Application") 'New Microsoft.Office.Interop.Word.Application
            'Dim MsWord As Object = New Microsoft.Office.Interop.Word.Application
            MsWord.Visible = False
            Dim objDoc As Object 'Object 
            objDoc = MsWord.Documents.Open(Me.FilePath, False, True, False)
            If objDoc.Tables.Count > 0 Then
                For i = 1 To objDoc.Tables.Count
                    TablesList.Add(String.Format("Table {0}", i))
                Next
            End If
            objDoc.Close(False)
            MsWord.Quit()
            objDoc = Nothing
            MsWord = Nothing
        End If
        Dim strTableNames As String() = CType(TablesList.ToArray(GetType(String)), String())
        Return strTableNames
    End Function
#End Region

#Region "Import"
    Public Function GetTableContent(ByVal intTableIndex As Integer) As System.Data.DataTable
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim MsWord As Object = objDoc.Application
        Dim objWordRow As Object, objWordCell As Object
        Dim dtWordTable As New System.Data.DataTable()
        Dim intRowIndex, intCellIndex As Integer
        Dim strCellValue As String

        MsWord.visible = False

        If objDoc IsNot Nothing AndAlso objDoc.Tables.Count >= intTableIndex Then
            Dim objTable As Object = objDoc.Tables(intTableIndex)

            If objTable IsNot Nothing Then
                Dim intRowCount As Integer = objTable.Rows.Count
                Dim intColumnCount As Integer = GetColumnCount(objTable)

                'COM exception when vertical cells are merged
                If intColumnCount = -1 Then
                    objDoc.Close(False)
                    MsWord.Quit()
                    Return dtWordTable
                End If

                For i = 1 To intColumnCount
                    dtWordTable.Columns.Add(String.Format("Col{0}", i), GetType(String))
                Next
                For intRowIndex = 1 To intRowCount
                    objWordRow = objTable.Rows(intRowIndex)
                    Dim NewRow As DataRow = dtWordTable.NewRow

                    For intCellIndex = 1 To objWordRow.Cells.Count
                        objWordCell = objWordRow.Cells(intCellIndex)
                        strCellValue = objWordCell.Range.Text
                        If Asc(strCellValue.Substring(strCellValue.Length - 1)) = 7 Then
                            strCellValue = strCellValue.Remove(strCellValue.Length - 2)
                            strCellValue = Trim(strCellValue)
                        End If

                        'Debug.Print(Asc(strCellValue.Substring(strCellValue.Length - 1)).ToString & vbTab & strCellValue)
                        If strCellValue <> String.Empty Then NewRow.Item(intCellIndex - 1) = strCellValue
                    Next
                    dtWordTable.Rows.Add(NewRow)
                Next
            End If
        End If
        objDoc.Close(False)
        MsWord.Quit()

        Return dtWordTable
    End Function

    Private Function GetColumnCount(ByVal objTable As Object)
        Dim objRow As Object
        Dim intColumnCount As Integer
        Dim intIndex As Integer

        For intIndex = 1 To objTable.Rows.count
            Try
                objRow = objTable.Rows(intIndex)
                If objRow.Cells.Count > intColumnCount Then intColumnCount = objRow.Cells.Count
            Catch ex As System.Runtime.InteropServices.COMException
                Dim strMsg As String
                Select Case UserLanguage
                    Case "fr"
                        strMsg = "Impossible de lire la table sélectionnée." & vbLf & vbLf & _
                              "Cette table contient des cellules fusionnées verticalement qui s'étendent sur plusieurs lignes. Pour permettre à Logframer d'importer cette table, veuillez ouvrir le document en Microsoft Word " & _
                              "et éliminer les cellules fusionnées verticalement (en les supprimant ou en les partageant)."
                    Case Else
                        strMsg = "Unable to read the selected table." & vbLf & vbLf & _
                            "This table contains vertically merged cells that span multiple rows. To enable Logframer to import this table, please open the document in Microsoft Word and remove the vertically merged cells " & _
                            "(by deleting them or by splitting them up again)."
                End Select
                MsgBox(strMsg, MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                Return -1
            End Try


        Next
        'Debug.Print("column count: " & intColumnCount.ToString)
        Return intColumnCount
    End Function
#End Region

#Region "Export"
    Public Function GetNewDocument(ByVal intOrientation As Integer) As Object
        Dim MsWord As Object = CreateObject("Word.Application")
        MsWord.Visible = True
        Dim objDoc As Object = MsWord.Documents.Add()

        Select Case intOrientation
            Case DialogExportToWord.OrientationValues.Landscape
                objDoc.PageSetup.Orientation = wdPageSetup_Orientation_Landscape
            Case DialogExportToWord.OrientationValues.Portrait
                objDoc.PageSetup.Orientation = wdPageSetup_Orientation_Portrait
        End Select

        MsWord.Selection.TypeParagraph()
        AppActivate(MsWord.Windows(1).Caption)

        Return objDoc
    End Function

    Public Function OpenDocument(ByVal strFilePath As String) As Object
        Dim MsWord As Object = CreateObject("Word.Application")
        MsWord.Visible = True
        Dim objDoc As Object = Nothing

        Try
            objDoc = MsWord.Documents.Open(strFilePath)
        Catch ex As Exception
            Dim strMsg As String
            Select Case UserLanguage
                Case "fr"
                    strMsg = "Impossible d'ouvrir le document que vous avez sélectionné. Il se peut que votre document est en lecture seule, car il est déjà ouvert dans une autre fenêtre." & vbLf & vbLf & _
                        "S.v.p. fermez le document et réessayez."
                Case Else
                    strMsg = "Can't open the document you selected. It may be that your document is read-only, because it is already open in another window. " & vbLf & vbLf &
                        "Please close the document and try again."
            End Select
            MsgBox(strMsg, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation)
        End Try
        Return objDoc
    End Function

    Public Function GetRange(ByVal objDoc As Object, ByVal intLocation As Integer, ByVal intOrientation As Integer, _
                             Optional ByVal strBookmark As String = "") As Object
        Dim MsWord As Object = objDoc.Application
        Dim objRange As Object = Nothing

        Select Case intLocation
            Case InsertLocations.AtBeginning
                objRange = objDoc.Range(0, 0)
                objRange.Select()
                MsWord.Selection.TypeParagraph()
                MsWord.Selection.InsertBreak(wdInsertBreak_SectionBreakNextPage)
            Case InsertLocations.AtBookmark
                If objDoc.Bookmarks.Exists(strBookmark) Then
                    objDoc.Bookmarks(strBookmark).Select()
                    MsWord.Selection.InsertBreak(wdInsertBreak_SectionBreakNextPage)
                    MsWord.Selection.TypeParagraph()
                    objRange = MsWord.Selection.Range
                    MsWord.Selection.TypeParagraph()
                    MsWord.Selection.InsertBreak(wdInsertBreak_SectionBreakNextPage)
                    'MsWord.Selection.MoveUp(wdUnits_Line, 1)
                End If
            Case InsertLocations.AtEnd
                objRange = objDoc.Characters.Last
                objRange.Select()
                MsWord.Selection.InsertBreak(wdInsertBreak_SectionBreakNextPage)
                MsWord.Selection.TypeParagraph()
                objRange = objDoc.Characters.Last
        End Select

        Select Case intOrientation
            Case DialogExportToWord.OrientationValues.Landscape
                objRange.PageSetup.Orientation = wdPageSetup_Orientation_Landscape
            Case DialogExportToWord.OrientationValues.Portrait
                objRange.PageSetup.Orientation = wdPageSetup_Orientation_Portrait
        End Select
        objRange.Style = wdStyleNormal

        Return objRange
    End Function
#End Region

#Region "Create table"
    Public Function CreateTable(ByVal objDoc As Object, ByVal objRange As Object, ByVal objTableGrid As Object(,)) As Object
        Dim intRowCount As Integer = objTableGrid.GetUpperBound(0) + 1
        Dim intColumnCount As Integer = objTableGrid.GetUpperBound(1) + 1

        'create a string to hold what is the objTableGrid array
        Dim strTable As String = String.Empty
        Dim strValue As String

        For intRowIndex = 0 To intRowCount - 1
            For intColumnIndex = 0 To intColumnCount - 1
                If objTableGrid(intRowIndex, intColumnIndex) Is Nothing Then
                    strValue = String.Empty
                Else
                    strValue = objTableGrid(intRowIndex, intColumnIndex).ToString
                    
                    If strValue.Contains(Chr(10)) Or strValue.Contains(Chr(13)) Then
                        strValue = strValue.Replace(Chr(10), String.Empty)
                        strValue = strValue.Replace(Chr(13), String.Empty)
                    End If
                End If

                strTable &= strValue

                'If String.IsNullOrEmpty(strValue) = False Then strTable &= strValue
                'strTable &= vbTab
                If intRowIndex < intRowCount - 1 Then
                    strTable &= vbTab
                Else
                    If intColumnIndex < intColumnCount - 1 Then
                        strTable &= vbTab
                    End If
                End If

            Next
        Next
        'Debug.Print(strTable)
        'write the text to the document. It is not a table yet.
        objRange.Text = strTable

        'convert text to a table
        Dim objTable As Object = objRange.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=intColumnCount, NumRows:=intRowCount, _
                                                         Format:=wdTableFormatGrid1, ApplyBorders:=True, AutoFitBehavior:=wdAutofitBehaviour_FitContent)

        Return objTable
    End Function
#End Region

#Region "Create letter"
    Public Sub SendLetterToAddress(ByVal selContact As Contact, ByVal selOrganisation As Organisation, ByVal selAddress As Address)
        Dim objDoc As Object = GetNewDocument(wdPageSetup_Orientation_Portrait)
        If objDoc Is Nothing Then Exit Sub

        Dim MsWord As Object = objDoc.Application
        Dim objRange As Object = objDoc.Range(0, 0)
        Dim strContact As String = String.Empty
        Dim strOrganisationName As String = String.Empty

        If selContact IsNot Nothing Then
            strContact = Trim(String.Format("{0} {1}", selContact.Title, selContact.FullName))
        End If
        If selOrganisation IsNot Nothing Then
            strOrganisationName = selOrganisation.FullName
        End If

        objRange.Select()
        With objDoc.Application.Selection
            With .ParagraphFormat
                .LeftIndent = MsWord.CentimetersToPoints(9)
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
            End With

            If String.IsNullOrEmpty(strContact) = False Then
                .TypeText(strContact)
                .TypeParagraph()
            End If
            If String.IsNullOrEmpty(strOrganisationName) = False Then
                .TypeText(strOrganisationName)
                .TypeParagraph()
            End If
            .TypeText(selAddress.FullStreet)
            .TypeParagraph()
            .TypeText(selAddress.FullTown)
            .TypeParagraph()
            .TypeText(selAddress.Country.ToUpper)
            .TypeParagraph()

            .Style = wdStyleNormal


            objRange = .Range
        End With

        objRange.Collapse(wdCollapseEnd)
        objRange.InsertParagraphAfter()

        objRange.Select()
        With objDoc.Application.Selection
            .TypeParagraph()
            .Style = wdStyleNormal
            objRange = .Range
        End With
    End Sub
#End Region

#Region "Styles"
    Private Sub SetTableStyle_Basic(ByVal MsWord As Object, ByVal objDoc As Object)
        'Basic table
        Dim TableStyleBasic As Object = objDoc.Styles.Add("Logframer table basic", wdStyleTypeTable)

        With TableStyleBasic.Font
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = 0
            .UnderlineColor = wdColorAutomatic
            .StrikeThrough = False
            .Color = wdColorAutomatic
        End With
        With TableStyleBasic.ParagraphFormat
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RightIndent = MsWord.CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
        End With
        TableStyleBasic.NoSpaceBetweenParagraphsOfSameStyle = False
        TableStyleBasic.ParagraphFormat.TabStops.ClearAll()
        TableStyleBasic.Frame.Delete()

        With TableStyleBasic.Table
            .TableDirection = 1
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)
            .Alignment = 0
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowBreakAcrossPage = True
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RowStripe = 0
            .ColumnStripe = 0

            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderVertical)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            .Borders.Shadow = False

            With .Condition(wdFirstRow)
                .TopPadding = MsWord.CentimetersToPoints(0)
                .BottomPadding = MsWord.CentimetersToPoints(0)
                .LeftPadding = MsWord.CentimetersToPoints(0.19)
                .RightPadding = MsWord.CentimetersToPoints(0.19)
                .ParagraphFormat.TabStops.ClearAll()

                With .Font
                    .Name = ""
                    .Bold = True
                End With

                With .Shading
                    .Texture = wdTextureNone
                    .ForegroundPatternColor = wdColorAutomatic
                    .BackgroundPatternColor = -587137152
                End With
                .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                .Borders(wdBorderRight).LineStyle = wdLineStyleNone
                .Borders(wdBorderTop).LineStyle = wdLineStyleNone
                .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
                .Borders.Shadow = False

                With .ParagraphFormat
                    .LeftIndent = MsWord.CentimetersToPoints(0)
                    .RightIndent = MsWord.CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceSingle
                    .Alignment = wdAlignParagraphCenter
                    .WidowControl = True
                End With
            End With
        End With
    End Sub

    Private Sub SetTableStyle_StripedGreen(ByVal MsWord As Object, ByVal objDoc As Object)
        Dim TableStylePlanning As Object = objDoc.Styles.Add("Logframer table striped", wdStyleTypeTable)

        With TableStylePlanning.Font
            .Size = 11
            .Bold = False
            .Italic = False
            .Underline = 0
            .UnderlineColor = wdColorAutomatic
            .StrikeThrough = False
            .DoubleStrikeThrough = False
            .Outline = False
            .Emboss = False
            .Shadow = False
            .Hidden = False
            .SmallCaps = False
            .AllCaps = False
            .Color = wdColorAutomatic
        End With

        With TableStylePlanning.ParagraphFormat
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RightIndent = MsWord.CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .WidowControl = True
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = MsWord.CentimetersToPoints(0)
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .CollapsedByDefault = False
        End With
        TableStylePlanning.NoSpaceBetweenParagraphsOfSameStyle = False
        TableStylePlanning.ParagraphFormat.TabStops.ClearAll()
        TableStylePlanning.Frame.Delete()

        With TableStylePlanning.Table
            .TableDirection = 1
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)
            .Alignment = wdAlignRowLeft
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowBreakAcrossPage = True
            .LeftIndent = MsWord.CentimetersToPoints(0)
            .RowStripe = 1
            .ColumnStripe = 0

            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderHorizontal)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderVertical)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            .Borders.Shadow = False
        End With
        With TableStylePlanning.Table.Condition(wdFirstRow)
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)

            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = -654262273
            End With
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
            .Borders.Shadow = False

            With .ParagraphFormat
                .LeftIndent = MsWord.CentimetersToPoints(0)
                .RightIndent = MsWord.CentimetersToPoints(0)
                .SpaceBefore = 3
                .SpaceBeforeAuto = False
                .SpaceAfter = 3
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .Alignment = wdAlignParagraphLeft
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = MsWord.CentimetersToPoints(0)
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0.6
                .LineUnitAfter = 0.6
                .MirrorIndents = False
                .CollapsedByDefault = False
            End With

            With .Font
                .Name = ""
                .Bold = True
            End With
            .ParagraphFormat.TabStops.ClearAll()
        End With

        With TableStylePlanning.Table.Condition(wdOddRowBanding)
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)

            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = -654246093
            End With
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
            .Borders.Shadow = False

            With .ParagraphFormat
                .LeftIndent = MsWord.CentimetersToPoints(0)
                .RightIndent = MsWord.CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .Alignment = wdAlignParagraphLeft
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = MsWord.CentimetersToPoints(0)
                .OutlineLevel = wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = False
                .CollapsedByDefault = False
            End With

            .Font.Name = ""
            .ParagraphFormat.TabStops.ClearAll()
        End With

        With TableStylePlanning.Table.Condition(wdEvenRowBanding)
            .TopPadding = MsWord.CentimetersToPoints(0)
            .BottomPadding = MsWord.CentimetersToPoints(0)
            .LeftPadding = MsWord.CentimetersToPoints(0.19)
            .RightPadding = MsWord.CentimetersToPoints(0.19)

            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = -654246042
            End With
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone
            .Borders(wdBorderTop).LineStyle = wdLineStyleNone
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
            .Borders.Shadow = False

            With .ParagraphFormat
                .LeftIndent = MsWord.CentimetersToPoints(0)
                .RightIndent = MsWord.CentimetersToPoints(0)
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceSingle
                .Alignment = wdAlignParagraphLeft
                .WidowControl = True
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = False
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = MsWord.CentimetersToPoints(0)
                .OutlineLevel = wdOutlineLevelBodyText
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .MirrorIndents = False
                .CollapsedByDefault = False
            End With
            .Font.Name = ""
            .ParagraphFormat.TabStops.ClearAll()
        End With
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
