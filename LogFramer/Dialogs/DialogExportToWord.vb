Imports System.Windows.Forms

Public Class DialogExportToWord
    Private Bookmarks As List(Of String)
    Private objWordIO As New WordIO
    Private strFilePath As String
    Private intSection As Integer
    Private intOrientation As Integer
    Private intReport As Integer
    Private intPlanningElements As Integer

    Public Enum Reports
        Logframe
        Planning
        Pmf
        RisksTable
        AssumptionsTable

        IndicatorList
        ResourcesTable
        PartnerList
        TargetGroupIdForm
    End Enum

    Public Enum SectionNames As Integer
        NotSelected = -1
        SelectGoals = 0
        SelectPurposes = 1
        SelectOutputs = 2
        SelectActivities = 3
        SelectAll = 4
    End Enum

    Public Enum OrientationValues As Integer
        Portrait = 0
        Landscape = 1
    End Enum

#Region "Properties"
    Public Property Report As Integer
        Get
            Return intReport
        End Get
        Set(ByVal value As Integer)
            intReport = value
        End Set
    End Property

    Public Property Orientation As Integer
        Get
            Return intOrientation
        End Get
        Set(ByVal value As Integer)
            intOrientation = value
        End Set
    End Property

    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property PlanningElements() As Integer
        Get
            Return intPlanningElements
        End Get
        Set(ByVal value As Integer)
            intPlanningElements = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub New(ByVal report As Integer)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Dim strReport As String = String.Empty

        Me.Report = report
        rbtnNewWordDoc.Checked = True
        rbtnEndDoc.Checked = True

        If Me.Report <> Reports.IndicatorList Then TabControlOptions.TabPages.Remove(TabPageIndicatorOptions)
        If Me.Report <> Reports.Planning Then TabControlOptions.TabPages.Remove(TabPagePlanningOptions)
        Select Case Me.Report
            Case Reports.Logframe
                strReport = LANG_LogicalFramework
            Case Reports.IndicatorList
                strReport = LANG_ReportListOfIndicators

                cmbPrintLevel.Items.Add(My.Settings.setStruct1)
                cmbPrintLevel.Items.Add(My.Settings.setStruct2)
                cmbPrintLevel.Items.Add(My.Settings.setStruct3)
                cmbPrintLevel.Items.Add(My.Settings.setStruct4)
                cmbPrintLevel.Items.Add(LANG_All)
                cmbPrintLevel.SelectedIndex = 2
                chkPrintPurpose.Checked = My.Settings.setPrintIndicatorsPrintPurposes
                chkPrintOutput.Checked = My.Settings.setPrintIndicatorsPrintOutputs
                chkPrintResponse.Checked = My.Settings.setPrintIndicatorsPrintOptionValues
                chkPrintRange.Checked = My.Settings.setPrintIndicatorsPrintValueRanges
                chkPrintTarget.Checked = My.Settings.setPrintIndicatorsPrintTargets
            Case Reports.Pmf
                strReport = LANG_ReportPmf
            Case Reports.RisksTable
                strReport = LANG_ReportRiskRegister
            Case Reports.ResourcesTable
                strReport = LANG_ReportResourcesTable
            Case Reports.Planning
                strReport = LANG_Planning

                With Me.cmbShowElements
                    .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    .DropDownStyle = ComboBoxStyle.DropDownList
                    .Items.AddRange(LIST_ShowElements)
                    .SelectedIndex = My.Settings.setPlanningElementsView
                End With
            Case Reports.PartnerList
                strReport = LANG_ReportPartnerList
            Case Reports.TargetGroupIdForm
                strReport = LANG_ReportTargetGroupIdForm
        End Select
        rbtnNewWordDoc.Text = String.Format(LANG_ExportToNewWordDoc, strReport)
        rbtnWordDoc.Text = String.Format(LANG_ExportToExistingWordDoc, strReport)
        gbInsertOptions.Text = String.Format(LANG_InsertAt, strReport)
        rbtnLandscape.Checked = True
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        'Dim selBookmark As String = cmbBookmark.SelectedItem.ToString

        If rbtnBookmark.Checked = True And cmbBookmark.SelectedItem Is Nothing Then
            MsgBox(LANG_SelectBookmark, MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()

        Using objWord As New WordIO
            Dim strBookmark As String = cmbBookmark.Text
            Dim intLocation As Integer

            If rbtnStartDoc.Checked = True Then
                intLocation = WordIO.InsertLocations.AtBeginning
            ElseIf rbtnEndDoc.Checked = True Then
                intLocation = WordIO.InsertLocations.AtEnd
            Else
                intLocation = WordIO.InsertLocations.AtBookmark
            End If

            Select Case Me.Report
                Case Reports.Logframe
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportLogframeToNewDocument(Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportLogframeToDocument(strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportLogframeToBookmark(strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If
                Case Reports.Planning
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportPlanningToNewDocument(Me.PlanningElements, Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportPlanningToDocument(Me.PlanningElements, strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportPlanningToBookmark(Me.PlanningElements, strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If
                Case Reports.Pmf
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportPmfToNewDocument(Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportPmfToDocument(strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportPmfToBookmark(strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If
                Case Reports.RisksTable
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportRisksTableToNewDocument(Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportRisksTableToDocument(strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportRisksTableToBookmark(strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If
                Case Reports.AssumptionsTable
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportAssumptionsTableToNewDocument(Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportAssumptionsTableToDocument(strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportAssumptionsTableToBookmark(strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If

                    'Case Reports.IndicatorList
                    '    If rbtnNewWordDoc.Checked = True Then
                    '        objWord.ExportIndicatorsToNewDocument(Me.Section, Me.Orientation)
                    '    Else
                    '        If strBookmark = String.Empty Then
                    '            objWord.ExportIndicatorsToDocument(Me.Section, strFilePath, intLocation, Me.Orientation)
                    '        Else
                    '            objWord.ExportIndicatorsToBookmark(Me.Section, strFilePath, strBookmark, Me.Orientation)
                    '        End If
                    '    End If
                Case Reports.PartnerList
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportPartnerListToNewDocument(Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportPartnerListToDocument(strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportPartnerListToBookmark(strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If
                Case Reports.TargetGroupIdForm
                    If rbtnNewWordDoc.Checked = True Then
                        objWord.ExportTargetGroupIdFormToNewDocument(Me.Orientation)
                    Else
                        If strBookmark = String.Empty Then
                            objWord.ExportTargetGroupIdFormToDocument(strFilePath, intLocation, Me.Orientation)
                        Else
                            objWord.ExportTargetGroupIdFormToBookmark(strFilePath, strBookmark, Me.Orientation)
                        End If
                    End If

                    'Case Reports.ResourcesTable
                    '    If rbtnNewWordDoc.Checked = True Then
                    '        objWord.ExportResourcesTableToNewDocument(Me.Orientation)
                    '    Else
                    '        If strBookmark = String.Empty Then
                    '            objWord.ExportResourcesTableToDocument(strFilePath, intLocation, Me.Orientation)
                    '        Else
                    '            objWord.ExportResourcesTableToBookmark(strFilePath, strBookmark, Me.Orientation)
                    '        End If
                    '    End If
            End Select
        End Using

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
#End Region

#Region "Word export options"
    Private Sub rbtnNewWordDoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnNewWordDoc.CheckedChanged
        If rbtnNewWordDoc.Checked = True Then PanelExportOptions.Enabled = True Else 
        PanelExportOptions.Enabled = False
    End Sub

    Private Sub rbtnWordDoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnWordDoc.CheckedChanged
        If rbtnWordDoc.Checked = True Then
            PanelExportOptions.Enabled = True
        Else
            PanelExportOptions.Enabled = False
        End If
        If tbFileName.Text = String.Empty Then
            btnExport.Enabled = False
            gbInsertOptions.Enabled = False
        Else
            btnExport.Enabled = True
            gbInsertOptions.Enabled = True
        End If
    End Sub

    Private Sub btnLoadFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadFile.Click
        Dim objWord As New WordIO
        strFilePath = objWord.GetDocumentName()

        If strFilePath <> String.Empty Then
            Dim selFileInfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(strFilePath)
            tbFileName.Text = selFileInfo.Name
            gbInsertOptions.Enabled = True
            btnExport.Enabled = True

            Dim BookMarksList As ArrayList = objWord.GetBookMarksList(strFilePath)

            cmbBookmark.Items.Clear()
            If BookMarksList.Count > 0 Then
                cmbBookmark.Items.AddRange(BookMarksList.ToArray)
                cmbBookmark.Enabled = True
                rbtnBookmark.Enabled = True
            Else
                cmbBookmark.Enabled = False
                rbtnBookmark.Enabled = False
            End If
        End If
    End Sub

    Private Sub cmbBookmark_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbBookmark.SelectedIndexChanged
        rbtnBookmark.Checked = True
    End Sub
#End Region

#Region "Page orientation"
    Private Sub rbtnLandscape_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnLandscape.CheckedChanged
        If rbtnLandscape.Checked = True Then Orientation = OrientationValues.Landscape Else Orientation = OrientationValues.Portrait
    End Sub
#End Region

#Region "Indicator export options"
    Private Sub cmbPrintLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPrintLevel.SelectedIndexChanged
        intSection = cmbPrintLevel.SelectedIndex
    End Sub
#End Region

#Region "Planning export options"
    Private Sub cmbShowElements_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbShowElements.SelectedIndexChanged
        Me.PlanningElements = cmbShowElements.SelectedIndex
    End Sub
#End Region

    
End Class
