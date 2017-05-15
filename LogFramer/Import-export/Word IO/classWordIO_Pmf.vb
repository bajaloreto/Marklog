Partial Public Class WordIO
    Public Sub ExportPmfToNewDocument(ByVal intOrientation As Integer)
        Dim objDoc As Object = GetNewDocument(intOrientation)
        Dim objRange As Object = objDoc.Range(0, 0)

        If objDoc Is Nothing Then Exit Sub

        Export_Pmf(objDoc, objRange)
    End Sub

    Public Sub ExportPmfToDocument(ByVal strFilePath As String, ByVal intLocation As Integer, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, intLocation, intOrientation)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        Export_Pmf(objDoc, objRange)
    End Sub

    Public Sub ExportPmfToBookmark(ByVal strFilePath As String, ByVal strBookmark As String, ByVal intOrientation As Integer)
        Dim objDoc As Object = OpenDocument(strFilePath)
        Dim objRange As Object = GetRange(objDoc, InsertLocations.AtBookmark, intOrientation, strBookmark)

        If objDoc Is Nothing Or objRange Is Nothing Then Exit Sub

        Export_Pmf(objDoc, objRange)
    End Sub

    Private Sub Export_Pmf(ByVal objDoc As Object, ByVal objRange As Object)
        'Set styles
        SetTableStyle_Basic(objDoc.Application, objDoc)

        Using objExportPmf As New ExportPmf(CurrentLogFrame, My.Settings.setPrintPmfSection, My.Settings.setPrintPmfShowTargetRowTitles) 'PrintPMF.PrintSections.PrintAll
            objExportPmf.LoadTable()

            Dim objTableGrid As Object = ExportPmf_CreateTableGrid(objExportPmf)
            Dim objTable As Object = Nothing
            Dim intTargetCount As Integer = objExportPmf.GetTargetColumnCount()

            If objTableGrid IsNot Nothing Then objTable = CreateTable(objDoc, objRange, objTableGrid)

            If objTable IsNot Nothing Then
                With objTable
                    .Rows(1).HeadingFormat = True
                    .Cell(1, 6 + intTargetCount).Merge(.Cell(1, 7 + intTargetCount))
                    .Cell(1, 3).Merge(.Cell(1, 4))
                    .Cell(1, 1).Merge(.Cell(1, 2))
                    .Style = "Logframer table basic"

                    .AutoFitBehavior(wdAutofitBehaviour_FitContent)
                End With

                objRange.Collapse(wdCollapseEnd)
                objRange.InsertParagraphAfter()
            End If
        End Using


    End Sub

    Private Function ExportPmf_CreateTableGrid(ByVal objExportPmf As ExportPmf) As Object
        Dim intColumnCount As Integer = objExportPmf.GetColumnCount
        Dim intRowCount As Integer = objExportPmf.GetRowCount(True) + 1
        Dim objTableGrid(intRowCount - 1, intColumnCount - 1) As Object
        Dim intRow As Integer = 1

        If intRowCount = 0 Then Return Nothing

        Dim strObjectives As String = objExportPmf.GetColumnTitleObjectives()
        Dim intColumnIndex As Integer = 5
        Dim intTargetCount As Integer = objExportPmf.GetTargetColumnCount()
        Dim intTargetIndex As Integer
        Dim strHeaderText() As String = objExportPmf.GetColumnTitlesTargets
        Dim strValue As String

        'Title row
        objTableGrid(0, 1) = strObjectives
        objTableGrid(0, 3) = LANG_Indicators
        objTableGrid(0, 4) = LANG_Baseline

        For i = 0 To intTargetCount - 1
            objTableGrid(0, 5 + i) = strHeaderText(i)
            intColumnIndex += 1
        Next
        intColumnIndex += 1
        objTableGrid(0, intColumnIndex) = LANG_VerificationSources
        intColumnIndex += 1
        objTableGrid(0, intColumnIndex) = LANG_DataCollection
        intColumnIndex += 1
        objTableGrid(0, intColumnIndex) = LANG_Frequency
        intColumnIndex += 1
        objTableGrid(0, intColumnIndex) = LANG_Responsibility

        For Each selGridRow As ExportPmfRow In objExportPmf.TableGrid
            objTableGrid(intRow, 0) = selGridRow.StructSort
            objTableGrid(intRow, 1) = selGridRow.StructText
            objTableGrid(intRow, 2) = selGridRow.IndicatorSort
            objTableGrid(intRow, 3) = selGridRow.IndicatorText

            If selGridRow.Indicator IsNot Nothing Then
                If selGridRow.Indicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selGridRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                    strValue = selGridRow.Indicator.GetBaselineFormattedValue
                Else
                    strValue = selGridRow.Indicator.GetBaselineFormattedScore
                End If

                If selGridRow.Indicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                    strValue &= " | "
                    If selGridRow.Indicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selGridRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                        strValue &= selGridRow.Indicator.GetPopulationBaselineFormattedValue
                    Else
                        strValue &= selGridRow.Indicator.GetPopulationBaselineFormattedScore
                    End If
                End If
                objTableGrid(intRow, 4) = strValue

                intColumnIndex = 5
                For Each selTargetDeadline As TargetDeadline In objExportPmf.TargetDeadlinesSection.TargetDeadlines
                    intTargetIndex = 0
                    For Each selTarget As Target In selGridRow.Targets
                        If selTarget.TargetDeadlineGuid = selTargetDeadline.Guid Then
                            If selGridRow.Indicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selGridRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                                strValue = selGridRow.Indicator.GetTargetFormattedValue(intTargetIndex)
                            Else
                                strValue = selGridRow.Indicator.GetTargetFormattedScore(intTargetIndex)
                            End If

                            If selGridRow.Indicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then
                                strValue &= " | "
                                If selGridRow.Indicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selGridRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                                    strValue &= selGridRow.Indicator.GetPopulationTargetFormattedValue(intTargetIndex)
                                Else
                                    strValue &= DisplayAsUnit(selGridRow.Indicator.PopulationTargets(intTargetIndex).Percentage, selGridRow.Indicator.ValuesDetail.NrDecimals, "%")
                                End If

                                strValue &= " | "
                                If selGridRow.Indicator.QuestionType = Indicator.QuestionTypes.AbsoluteValue And selGridRow.Indicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                                    strValue &= selGridRow.Indicator.GetPopulationTargetFormattedValue(intTargetIndex)
                                Else
                                    strValue &= selGridRow.Indicator.GetPopulationTargetFormattedScore(intTargetIndex)
                                End If
                            End If

                            objTableGrid(intRow, intColumnIndex) = strValue
                            intColumnIndex += 1
                        End If

                        intTargetIndex += 1
                    Next
                Next
            End If

            intColumnIndex = 5 + intTargetCount

            If selGridRow.VerificationSource IsNot Nothing Then objTableGrid(intRow, intColumnIndex) = selGridRow.VerificationSourceSort
            intColumnIndex += 1
            If selGridRow.VerificationSource IsNot Nothing Then objTableGrid(intRow, intColumnIndex) = selGridRow.VerificationSourceText
            intColumnIndex += 1
            If selGridRow.VerificationSource IsNot Nothing Then objTableGrid(intRow, intColumnIndex) = selGridRow.VerificationSource.CollectionMethod
            intColumnIndex += 1
            If selGridRow.VerificationSource IsNot Nothing Then objTableGrid(intRow, intColumnIndex) = selGridRow.Frequency
            intColumnIndex += 1
            If selGridRow.VerificationSource IsNot Nothing Then objTableGrid(intRow, intColumnIndex) = selGridRow.Responsibility

            intRow += 1
        Next

        Return objTableGrid
    End Function

    
End Class
