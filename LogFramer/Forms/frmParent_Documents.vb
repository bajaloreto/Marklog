Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO

Partial Public Class frmParent
#Region "New project"
    Private Sub RibbonButtonNewProject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonNewProject.Click
        NewProject()
    End Sub

    Private Sub NewProject()
        AddNewProject()
        AddNewProjectButton(CurrentProjectForm.Name)
        Initialize_Panels()
        Initialize_Clipboard()
    End Sub

    Public Sub UpdateProjectButton()
        Dim strDocName As String = CurrentProjectForm.Name

        If strDocName = String.Empty Then
            Dim intNumber As Integer = MdiChildren.Count + 1
            strDocName = LANG_Project & " " & intNumber.ToString
        End If
        CurrentProjectButton.Text = strDocName
    End Sub

    Private Sub AddNewProject()
        Dim frmNewProject As New frmProject
        Dim intNumber As Integer = MdiChildren.Count + 1
        Dim strName As String = String.Format("{0} {1}", LANG_Project, intNumber)

        CurrentProjectForm = frmNewProject
        With frmNewProject
            .MdiParent = Me
            .Name = strName
            .Text = strName

            .AddNewProject()
            .Show()
        End With
        ReloadSplitUndoRedoButtons()
        SetRibbonLayOutTabsVisibility()
    End Sub

    Private Sub ProjectInitialise(ByVal selProjectForm As frmProject)
        With selProjectForm
            .ProjectInitialise()

            If Me.MdiChildren.Count > 0 Then
                Dim FirstForm As frmProject = Me.MdiChildren(0)
                .SplitContainerUtilities.Panel1Collapsed = FirstForm.SplitContainerUtilities.Panel1Collapsed
            Else
                .SplitContainerUtilities.Panel1Collapsed = True
            End If
            .TabControlProject.SelectTab(CurrentMainTab)

            'Planning buttons
            SetPlanningPeriodViewButtons(.dgvPlanning.PeriodView)
            SetPlanningElementsViewButtons()
            RibbonButtonPlanningShowActivityLinks.Checked = My.Settings.setPlanningActivityLinks
            RibbonButtonPlanningShowKeyMomentLinks.Checked = My.Settings.setPlanningKeyMomentLinks

            'Budget buttons
            SetBudgetButtons()

            'Ribbon tabs Lay-out
            SetRibbonLayOutTabsVisibility()
        End With
    End Sub

    Private Sub AddNewProjectButton(ByVal strDocName As String)
        With RibbonPanelProjects
            Dim NewButton As New RibbonButton(strDocName)
            NewButton.SmallImage = FaciliDev.LogFramer.My.Resources.FileProject
            NewButton.Image = FaciliDev.LogFramer.My.Resources.FileProject_large
            NewButton.CheckOnClick = True
            AddHandler NewButton.Click, AddressOf RibbonButtonProject_Click
            .Items.Add(NewButton)

            'uncheck other project buttons
            For Each selButton As RibbonButton In RibbonPanelProjects.Items
                selButton.Checked = False
            Next
            'check the current project button
            NewButton.Checked = True
        End With

        RibbonLF.ResumeUpdating(True)
    End Sub

    Private Sub RibbonButtonProject_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For Each selButton As RibbonButton In RibbonPanelProjects.Items
            selButton.Checked = False
        Next
        CurrentProjectButton = CType(sender, RibbonButton)
        CurrentProjectButton.Checked = True

        For Each selForm As Form In Me.MdiChildren
            If selForm.Name = CurrentProjectButton.Text Then
                CurrentProjectForm = selForm
                selForm.Activate()
                ReloadSplitUndoRedoButtons()
                Return
            End If
        Next
    End Sub
#End Region

#Region "Open project"
    Private Sub RibbonButtonOpenProject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonOpenProject.Click
        OpenProject()
    End Sub

    Private Function CancelIfNotSaved() As Boolean

        If CurrentProjectForm.Name = String.Empty Or CurrentUndoList.Count > 0 Then
            Dim strMsg As String = LANG_SaveChanges

            Dim reply = MsgBox(strMsg, MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question + MsgBoxStyle.DefaultButton1, )

            If reply = MsgBoxResult.Yes Then
                Me.SaveProject(True)
                Return False
            ElseIf reply = MsgBoxResult.No Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Public Sub OpenProject(Optional ByVal strDocPath As String = "")
        Dim getInfo As System.IO.DirectoryInfo
        Dim strDocName As String = String.Empty
        Dim boolClose As Boolean
        Dim selLogframe As LogFrame = New LogFrame

        If MdiChildren.Count = 1 And CurrentProjectForm.IsDefaultProjectName And CurrentProjectForm.UndoList.Count = 0 Then boolClose = True

        If strDocPath = String.Empty Then
            'Ask user to select file
            strDocPath = OpenProject_SelectFile(strDocPath)
            If strDocPath = String.Empty Then Return
        Else
            'Verify if document is opened
            'If that is the case, do not open it again
            If CheckIfOpened(strDocPath) = True Then Return
        End If

        Try
            getInfo = My.Computer.FileSystem.GetDirectoryInfo(strDocPath)
            strDocName = getInfo.Name

            StatusLabelGeneral.Text = LANG_OpeningLogframeDocument
            StatusLabelGeneral.Invalidate()

            Dim serializer As New LogframeSerializer(GetType(LogFrame))
            Dim LogframerFile As New FileStream(strDocPath, FileMode.Open, FileAccess.Read, FileShare.None)
            selLogframe = CType(serializer.Deserialize(LogframerFile), LogFrame)

            LogframerFile.Close()
            LogframerFile.Dispose()
            LogframerFile = Nothing
            serializer = Nothing

            'als versie < 2 gebruik XmlDocument om organisaties op te halen
            If Val(selLogframe.LogFramerVersion) < 2 Then OpenProject_PreviousVersions(strDocPath, selLogframe)

        Catch ex As Exception
            boolClose = False
            Dim strMsg As String = ERR_OpenProject & ex.Message
            Dim strTitle As String = ERRTITLE_OpenProject

            MsgBox(strMsg, vbOKOnly, strTitle)

            Exit Sub
        End Try

        OpenProject_CreateProjectForm(selLogframe, strDocName, strDocPath)

        AddNewProjectButton(CurrentProjectForm.Name)
        AddToRecentFilesList()
        Initialize_Panels()
        Initialize_Clipboard()
        ReloadSplitUndoRedoButtons()

        If boolClose = True Then
            MdiChildren(0).Close()
            RibbonPanelProjects.Items.RemoveAt(0)
            RibbonLF.ResumeUpdating(True)
        End If

        StatusLabelGeneral.Text = LANG_Ready
    End Sub

    Private Sub OpenProject_CreateProjectForm(ByVal selLogframe As LogFrame, ByVal strDocName As String, ByVal strDocPath As String)
        Dim frmNewProject As New frmProject

        CurrentProjectForm = frmNewProject
        With frmNewProject
            .ProjectLogframe = selLogframe
            .MdiParent = Me
            .Name = strDocName
            .Text = strDocName
            .DocName = strDocName
            .DocPath = strDocPath

            With .ProjectLogframe
                .CreateSectionTitles()
                .GetCustomMeasureUnits()
            End With

            ProjectInitialise(frmNewProject)
            .Show()
            .dgvLogframe.Reload()
            .dgvPlanning.Reload()
        End With
    End Sub

    Private Sub OpenProject_PreviousVersions(ByVal strDocPath As String, ByRef selLogframe As LogFrame)
        Dim objXmlDoc As New XmlDocument
        objXmlDoc.Load(strDocPath)

        'transform organisations to projectpartners
        Dim objOrganisations As XmlNodeList
        objOrganisations = objXmlDoc.SelectNodes("LogFrame/Partners/Organisation")

        For Each OrganisationNode As XmlNode In objOrganisations
            Dim objMemStream As New MemoryStream()
            Dim objStreamWriter As New StreamWriter(objMemStream)
            objStreamWriter.Write(OrganisationNode.OuterXml)
            objStreamWriter.Flush()

            objMemStream.Position = 0

            Dim ser As New XmlSerializer(GetType(Organisation))
            Dim selOrganisation As Organisation = TryCast(ser.Deserialize(objMemStream), Organisation)

            If selOrganisation IsNot Nothing Then
                Dim NewPartner As New ProjectPartner
                Dim RoleNode As XmlNode = OrganisationNode.SelectSingleNode("descendant::Role")
                If RoleNode IsNot Nothing Then _
                    NewPartner.Role = RoleNode.InnerText
                NewPartner.Organisation = selOrganisation
                selLogframe.ProjectPartners.Add(NewPartner)
            End If
        Next
    End Sub

    Public Function CheckIfOpened(ByVal strDocPath As String) As Boolean
        For Each selForm As frmProject In Me.MdiChildren
            If selForm.DocPath = strDocPath Then Return True
        Next

        Return False
    End Function

    Public Function OpenProject_SelectFile(ByVal strDocPath As String) As String
        Dim ofdOpen As New OpenFileDialog()

        Dim strMyDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim strParentDir As String
        Dim getInfo As System.IO.DirectoryInfo


        strDocPath = CurrentProjectForm.DocPath

        ofdOpen.Filter = LANG_FileDialogFilter
        ofdOpen.FilterIndex = 1
        ofdOpen.DefaultExt = "lfr"
        If strDocPath = String.Empty Then
            ofdOpen.InitialDirectory = strMyDocuments
        Else
            getInfo = My.Computer.FileSystem.GetDirectoryInfo(strDocPath)
            strParentDir = getInfo.Parent.ToString
            If strParentDir = String.Empty Then strParentDir = strMyDocuments
            ofdOpen.InitialDirectory = strParentDir
        End If

        If ofdOpen.ShowDialog <> DialogResult.OK Then Return String.Empty
        strDocPath = ofdOpen.FileName

        Return strDocPath
    End Function
#End Region

#Region "Recent files list"
    Private Sub AddToRecentFilesList()
        Dim strList As System.Collections.Specialized.StringCollection = My.Settings.setRecentFilesList
        Dim strDocPath As String = CurrentProjectForm.DocPath

        If strList IsNot Nothing Then
            'remove doubles
            Do Until strList.Contains(strDocPath) = False
                strList.Remove(strDocPath)
            Loop
        Else
            strList = New System.Collections.Specialized.StringCollection
        End If
        'put the file path at the top of the list
        strList.Insert(0, strDocPath)

        'verify if the files haven't been moved
        Dim strPath As String
        For i = strList.Count - 1 To 0 Step -1
            strPath = strList(i)
            If File.Exists(strPath) = False Then strList.Remove(strPath)
        Next

        'limit the number of files in the list
        If strList.Count > My.Settings.setRecentFilesMax Then
            Do Until strList.Count <= My.Settings.setRecentFilesMax
                strList.RemoveAt(strList.Count - 1)
            Loop
        End If

        My.Settings.setRecentFilesList = strList
        UpdateRecentFilesList()
    End Sub

    Private Sub UpdateRecentFilesList()
        With RibbonButtonRecentFiles
            .DropDownItems.Clear()

            If My.Settings.setRecentFilesList IsNot Nothing Then
                For Each strPath As String In My.Settings.setRecentFilesList
                    Dim getInfo As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(strPath)
                    Dim strDocName As String = getInfo.Name
                    Dim NewButton As New RibbonButton
                    With NewButton
                        .MaxSizeMode = RibbonElementSizeMode.Medium
                        .SmallImage = FaciliDev.LogFramer.My.Resources.FileProject
                        .Text = strDocName
                        AddHandler NewButton.Click, AddressOf RecentFileName_Click
                    End With

                    .DropDownItems.Add(NewButton)
                Next
            End If
        End With
    End Sub

    Private Sub RecentFileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strPath As String
        With RibbonButtonRecentFiles
            Dim selMenuItem As RibbonButton = CType(sender, RibbonButton)
            Dim intIndex As Integer = .DropDownItems.IndexOf(selMenuItem)
            strPath = My.Settings.setRecentFilesList(intIndex)
        End With

        OpenProject(strPath)
    End Sub
#End Region

#Region "Save project"
    Private Sub RibbonButtonSaveProject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSaveProject.Click
        Me.SaveProject(False)
    End Sub

    Private Sub RibbonButtonSaveProjectAs_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonSaveProjectAs.Click
        Me.SaveProject(True)
    End Sub

    Public Sub SaveProject(ByVal boolSaveAs As Boolean)
        Dim strDocName As String = CurrentProjectForm.DocName
        Dim strDocPath As String = CurrentProjectForm.DocPath
        Dim boolNew As Boolean

        If SaveProject_OlderVersion() = True Then
            CurrentLogFrame.LogFramerVersion = "2.0"
            boolSaveAs = True
        End If

        If strDocName = String.Empty Then boolSaveAs = True

        If boolSaveAs = True Then
            If strDocName = String.Empty Then
                With CurrentLogFrame
                    If .ShortTitle <> String.Empty Then
                        strDocName = .ShortTitle
                    ElseIf .ProjectTitle <> String.Empty Then
                        strDocName = .ProjectTitle
                    End If
                End With
            End If
            SaveProject_ChooseLocation(strDocName, strDocPath)
            boolNew = True
        End If
        If strDocPath = String.Empty Then Exit Sub

        Try
            StatusLabelGeneral.Text = LANG_Saving
            Dim serializer As New XmlSerializer(GetType(LogFrame))
            Using MyDoc As New FileStream(strDocPath, FileMode.Create, FileAccess.Write)
                serializer.Serialize(MyDoc, CurrentLogFrame)
            End Using
            serializer = Nothing

            If boolNew = True Then AddToRecentFilesList()
            If boolSaveAs = True Then
                With CurrentProjectForm
                    .Name = strDocName
                    .Text = strDocName
                    .DocName = strDocName
                    .DocPath = strDocPath
                End With
            End If
            StatusLabelGeneral.Text = LANG_Saved
        Catch ex As IOException
            Dim strMsg As String = String.Format(ERR_SaveProject, ex.InnerException.ToString, ex.Message)
            Dim strTitle As String = ERRTITLE_SaveProject

            MsgBox(strMsg, MsgBoxStyle.OkOnly, strTitle)
        End Try
    End Sub

    Private Sub SaveProject_ChooseLocation(ByRef strDocName, ByRef strDocPath)
        Dim sfdSave As New SaveFileDialog()
        Dim getInfo As System.IO.DirectoryInfo
        Dim strMyDocuments As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim strParentDir As String

        sfdSave.Filter = LANG_FileDialogFilter
        sfdSave.FilterIndex = 1
        sfdSave.DefaultExt = "lfr"
        If strDocName <> "" Then sfdSave.FileName = strDocName
        If strDocPath = String.Empty Then
            sfdSave.InitialDirectory = strMyDocuments
        Else
            getInfo = My.Computer.FileSystem.GetDirectoryInfo(strDocPath)
            strParentDir = getInfo.Parent.ToString
            If strParentDir = String.Empty Then strParentDir = strMyDocuments
            sfdSave.InitialDirectory = strParentDir
        End If

        If sfdSave.ShowDialog = DialogResult.OK Then
            strDocPath = sfdSave.FileName
            getInfo = My.Computer.FileSystem.GetDirectoryInfo(strDocPath)
            strDocName = getInfo.Name
            With CurrentProjectForm
                .Name = strDocName
                .DocName = strDocName
                .DocPath = strDocPath
            End With
            UpdateProjectButton()
        End If
    End Sub

    Private Function SaveProject_OlderVersion() As Boolean
        Select Case CurrentLogFrame.LogFramerVersion
            Case "1.0", "1.1", "1.2", "1.3"
                MsgBox(LANG_SaveAsVersion2, MsgBoxStyle.OkOnly)

                Return True
            Case Else
                Return False
        End Select

    End Function
#End Region

#Region "Close project"
    Private Sub RibbonButtonCloseProject_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RibbonButtonCloseProject.Click
        CloseProject()
    End Sub

    Private Sub CloseProject()
        If Me.ActiveMdiChild IsNot Nothing Then
            Dim strName As String = ActiveMdiChild.Name

            'if last project is the default project, then don't close
            If Me.MdiChildren.Count = 1 Then
                With CurrentProjectForm
                    If .IsDefaultProjectName = True And .UndoList.Count = 0 Then
                        Return
                    End If
                End With
            End If

            'else remove the form's button
            For Each selButton As RibbonButton In RibbonPanelProjects.Items
                If selButton.Text = strName Then
                    RibbonPanelProjects.Items.Remove(selButton)
                    RibbonLF.ResumeUpdating(True)
                    Exit For
                End If
            Next

            'close the form
            Me.ActiveMdiChild.Close()

            'if there are no more forms open, show a new project
            If MdiChildren.Count = 0 Then
                AddNewProject()
                AddNewProjectButton(CurrentProjectForm.Name)
                Initialize_Panels()
                Initialize_Clipboard()
            End If
        End If
    End Sub

    Private Sub Quit()
        Application.Exit()
    End Sub
#End Region

End Class
