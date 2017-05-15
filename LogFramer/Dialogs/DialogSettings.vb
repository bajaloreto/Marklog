Imports System.Windows.Forms

Public Class DialogSettings
    Private bindSettings As New BindingSource
    Private bindLogFrame As New BindingSource
    Private fntOldDefaultFont As Font
    Private intOldDefaultIntType As Integer
    Private strOldDefaultCurrency As String
    Private boolOldRepeatPurposes As Boolean
    Private boolOldRepeatOutputs As Boolean
    Private boolOldWarnLinkedObjDelete As Boolean

    Public Sub New()
        InitializeComponent()

        bindLogFrame.DataSource = CurrentLogFrame
        bindSettings.DataSource = My.Settings

        Dim lstCurrencyCodes1 As List(Of IdValuePair) = LoadCurrencyCodesList()
        Dim lstCurrencyCodes2 As List(Of IdValuePair) = LoadCurrencyCodesList()

        fntOldDefaultFont = New Font(My.Settings.setDefaultFont, My.Settings.setDefaultFont.Style)
        intOldDefaultIntType = My.Settings.setDefaultIndicatorType
        strOldDefaultCurrency = My.Settings.setDefaultCurrency
        boolOldRepeatPurposes = My.Settings.setRepeatPurposes
        boolOldRepeatOutputs = My.Settings.setRepeatOutputs
        boolOldWarnLinkedObjDelete = My.Settings.setWarnLinkedObjectDelete

        Dim colInstalledFonts As System.Drawing.Text.InstalledFontCollection = New System.Drawing.Text.InstalledFontCollection

        'Project logic
        dgvProjectLogic.LoadColumns(CurrentLogFrame.SchemeIndex)
        tbIndicatorName.DataBindings.Add("Text", bindLogFrame, "IndicatorName")
        tbVerificationSourceName.DataBindings.Add("Text", bindLogFrame, "VerificationSourceName")
        tbResourceName.DataBindings.Add("Text", bindLogFrame, "ResourceName")
        tbBudgetName.DataBindings.Add("Text", bindLogFrame, "BudgetName")
        tbAssumptionName.DataBindings.Add("Text", bindLogFrame, "AssumptionName")

        TabControlSettings.SelectTab(0)

        'Current logframe
        tbSortNumberFont.DataBindings.Add("Text", bindLogFrame, "SortNumberFont")
        tbSortNumberDivider.DataBindings.Add("Text", bindLogFrame, "SortNumberDivider")
        chkSortNumberTerminateWithDivider.DataBindings.Add("Checked", bindLogFrame, "SortNumberTerminateWithDivider")
        chkSortNumberRepeatParent.DataBindings.Add("Checked", bindLogFrame, "SortNumberRepeatParent")
        SortNumberExample()

        tbSectionTitleFont.DataBindings.Add("Text", bindLogFrame, "SectionTitleFont")

        If CurrentLogFrame.SectionTitleFontColor <> Nothing Then _
            tbSectionTitleFontColor.ForeColor = CurrentLogFrame.SectionTitleFontColor
        tbSectionTitleFontColor.DataBindings.Add("Text", bindLogFrame, "SectionTitleFontColor")

        tbSectionColorTop.DataBindings.Add("Text", bindLogFrame, "SectionColorTop")
        If CurrentLogFrame.SectionColorTop <> Nothing Then _
            tbSectionColorTop.BackColor = CurrentLogFrame.SectionColorTop

        tbSectionColorBottom.DataBindings.Add("Text", bindLogFrame, "SectionColorBottom")
        If CurrentLogFrame.SectionColorBottom <> Nothing Then _
            tbSectionColorBottom.BackColor = CurrentLogFrame.SectionColorBottom

        tbDetailsFont.DataBindings.Add("Text", bindLogFrame, "DetailsFont")

        With cmbCurrency
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataSource = lstCurrencyCodes2
            .DisplayMember = "Value"
            .ValueMember = "Id"
            .DataBindings.Add("SelectedValue", bindLogFrame, "CurrencyCode")
        End With

        'Default settings
        tbDefaultFont.DataBindings.Add("Text", bindSettings, "setDefaultFont")
        tbDefaultFontSortNumbers.DataBindings.Add("Text", bindSettings, "setDefaultFontSortNumbers")
        chkWarnLinkedObjDelete.DataBindings.Add("Checked", bindSettings, "setWarnLinkedObjectDelete")

        tbDefaultFontSections.DataBindings.Add("Text", bindSettings, "setDefaultFontSections")

        If My.Settings.setDefaultFontSectionColor <> Nothing Then _
            tbDefaultFontSectionColor.ForeColor = My.Settings.setDefaultFontSectionColor
        tbDefaultFontSectionColor.DataBindings.Add("Text", bindSettings, "setDefaultFontSectionColor")

        tbDefaultColorSectionTop.DataBindings.Add("Text", bindSettings, "setDefaultColorSectionTop")
        If My.Settings.setDefaultColorSectionTop <> Nothing Then _
            tbDefaultColorSectionTop.BackColor = My.Settings.setDefaultColorSectionTop

        tbDefaultColorSectionBottom.DataBindings.Add("Text", bindSettings, "setDefaultColorSectionBottom")
        If My.Settings.setDefaultColorSectionBottom <> Nothing Then _
            tbDefaultColorSectionBottom.BackColor = My.Settings.setDefaultColorSectionBottom

        chkRepeatPurposes.DataBindings.Add("Checked", bindSettings, "setRepeatPurposes")
        chkRepeatOutputs.DataBindings.Add("Checked", bindSettings, "setRepeatOutputs")

        tbDefaultFontUI.DataBindings.Add("Text", bindSettings, "setDefaultFontUI")

        Dim QuestionTypes As List(Of StructuredComboBoxItem) = LoadQuestionTypesComboList()
        With cmbDefaultIndType
            .DataSource = QuestionTypes
            .ValueMember = "Type"
            .DisplayMember = "Description"
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataBindings.Add("SelectedValue", bindSettings, "setDefaultIndicatorType")
        End With

        With cmbDefaultCurrency
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataSource = lstCurrencyCodes1
            .DisplayMember = "Value"
            .ValueMember = "Id"
            .DataBindings.Add("SelectedValue", bindSettings, "setDefaultCurrency")
        End With
        
        ntbRecentFilesMax.DataBindings.Add("Text", bindSettings, "setRecentFilesMax")
        ntbTimerAutoSave.DataBindings.Add("Text", bindSettings, "setTimerAutoSave")
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        If dgvProjectLogic.SelectedScheme IsNot Nothing Then
            Dim selScheme As ProjectLogicScheme = dgvProjectLogic.SelectedScheme

            With CurrentLogFrame
                .SchemeIndex = dgvProjectLogic.SelectedColumns(0).Index

                If selScheme IsNot Nothing Then
                    .StructName = New String() {selScheme.GoalName, selScheme.PurposeName, selScheme.OutputName, selScheme.ActivityName}
                    .StructNamePlural = New String() {selScheme.GoalNamePlural, selScheme.PurposeNamePlural, selScheme.OutputNamePlural, selScheme.ActivityNamePlural}
                End If

                If .IsEmpty = True Then
                    .SetSectionTitleFont(My.Settings.setDefaultFontSections, My.Settings.setDefaultFontSectionColor)
                    .SectionColorTop = My.Settings.setDefaultColorSectionTop
                    .SectionColorBottom = My.Settings.setDefaultColorSectionBottom
                End If
                .CreateSectionTitles()
            End With

            If Me.chkDefaultScheme.Checked = True Then
                With My.Settings
                    .setProjectLogicIndex = dgvProjectLogic.SelectedColumns(0).Index

                    If selScheme IsNot Nothing Then
                        .setStruct1 = selScheme.GoalNamePlural
                        .setStruct1sing = selScheme.GoalName
                        .setStruct2 = selScheme.PurposeNamePlural
                        .setStruct2sing = selScheme.PurposeName
                        .setStruct3 = selScheme.OutputNamePlural
                        .setStruct3sing = selScheme.OutputName
                        .setStruct4 = selScheme.ActivityNamePlural
                        .setStruct4sing = selScheme.ActivityName
                    End If
                End With
            End If
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
        My.Settings.setDefaultFont = New Font(fntOldDefaultFont, fntOldDefaultFont.Style)
        My.Settings.setDefaultIndicatorType = intOldDefaultIntType
        My.Settings.setDefaultCurrency = strOldDefaultCurrency
        My.Settings.setRepeatPurposes = boolOldRepeatPurposes
        My.Settings.setRepeatOutputs = boolOldRepeatOutputs
        My.Settings.setWarnLinkedObjectDelete = boolOldWarnLinkedObjDelete
    End Sub

#Region "Current logframe"
    Private Sub tbSortNumberFont_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbSortNumberFont.Click
        Dim fntDialog As New FontDialog
        With fntDialog
            .Font = CurrentLogFrame.SortNumberFont
            .FontMustExist = True
            .ShowColor = False
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                CurrentLogFrame.setSortNumberFont(fntDialog.Font)
                tbSortNumberFont.DataBindings(0).ReadValue()
            End If
        End With
    End Sub

    Private Sub tbSortNumberDivider_Validated(sender As Object, e As System.EventArgs) Handles tbSortNumberDivider.Validated
        SortNumberExample()
    End Sub

    Private Sub tbSortNumberDivider_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles tbSortNumberDivider.Validating
        Dim strDivider As String = tbSortNumberDivider.Text

        If strDivider.Length > 1 Then
            strDivider = strDivider.Substring(0, 1)
            tbSortNumberDivider.Text = strDivider
        End If
    End Sub

    Private Sub chkSortNumberTerminateWithDivider_Validated(sender As Object, e As System.EventArgs) Handles chkSortNumberTerminateWithDivider.Validated
        SortNumberExample()
    End Sub

    Private Sub chkSortNumberRepeatParent_Validated(sender As Object, e As System.EventArgs) Handles chkSortNumberRepeatParent.Validated
        SortNumberExample()
    End Sub

    Private Sub SortNumberExample()
        Dim strSortNumber As String = "1{0}1{0}1{0}1{1}"
        Dim strDivider As String = CurrentLogFrame.SortNumberDivider
        Dim strEndDivider As String = String.Empty

        If CurrentLogFrame.SortNumberRepeatParent = False Then
            strSortNumber = "1{1}"
        End If
        If CurrentLogFrame.SortNumberTerminateWithDivider Then
            strEndDivider = strDivider
        End If

        lblSortNumberExample.Text = String.Format(strSortNumber, strDivider, strEndDivider)
    End Sub

    Private Sub tbDetailsFont_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim fntDialog As New FontDialog
        With fntDialog
            .Font = CurrentLogFrame.DetailsFont
            .FontMustExist = True
            .ShowColor = False
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                CurrentLogFrame.SetDetailsFont(fntDialog.Font)
                tbDetailsFont.DataBindings(0).ReadValue()
            End If
        End With
    End Sub

    Private Sub tbSectionTitleFont_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbSectionTitleFont.Click
        Dim fntDialog As New FontDialog
        With fntDialog
            .Font = CurrentLogFrame.SectionTitleFont
            .FontMustExist = True
            .ShowColor = True
            .Color = CurrentLogFrame.SectionTitleFontColor
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                CurrentLogFrame.SetSectionTitleFont(.Font, .Color)
                tbSectionTitleFont.DataBindings(0).ReadValue()
                tbSectionTitleFontColor.ForeColor = .Color
                tbSectionTitleFontColor.DataBindings(0).ReadValue()
            End If
        End With
    End Sub

    Private Sub tbSectionTitleFontColor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbSectionTitleFontColor.Click
        Dim colDialog As New ColorDialog
        With colDialog
            .Color = CurrentLogFrame.SectionTitleFontColor
            .AllowFullOpen = True
            .AnyColor = True
            .CustomColors = New Integer() {Color.LightGray.ToArgb, Color.Gray.ToArgb}
            If .ShowDialog() = vbOK Then
                CurrentLogFrame.SectionTitleFontColor = .Color
                tbSectionTitleFontColor.DataBindings(0).ReadValue()
                tbSectionTitleFontColor.ForeColor = .Color
            End If
        End With
    End Sub

    Private Sub tbSectionColorTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbSectionColorTop.Click
        Dim colDialog As New ColorDialog
        With colDialog
            .Color = CurrentLogFrame.SectionColorTop
            .AllowFullOpen = True
            .AnyColor = True
            .CustomColors = New Integer() {Color.LightGray.ToArgb, Color.Gray.ToArgb}
            If .ShowDialog() = vbOK Then
                CurrentLogFrame.SectionColorTop = .Color
                tbSectionColorTop.DataBindings(0).ReadValue()
                tbSectionColorTop.BackColor = .Color
            End If
        End With
    End Sub

    Private Sub tbSectionColorBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbSectionColorBottom.Click
        Dim colDialog As New ColorDialog
        With colDialog
            .Color = CurrentLogFrame.SectionColorBottom
            .AllowFullOpen = True
            .AnyColor = True
            .CustomColors = New Integer() {Color.LightGray.ToArgb, Color.Gray.ToArgb}
            If .ShowDialog() = vbOK Then
                CurrentLogFrame.SectionColorBottom = .Color
                tbSectionColorBottom.DataBindings(0).ReadValue()
                tbSectionColorBottom.BackColor = .Color
            End If
        End With
    End Sub

    Private Sub btnEqualColorLf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqualColorLf.Click
        My.Settings.setDefaultColorSectionBottom = My.Settings.setDefaultColorSectionTop
        tbDefaultColorSectionBottom.DataBindings(0).ReadValue()
        tbDefaultColorSectionBottom.BackColor = My.Settings.setDefaultColorSectionBottom
    End Sub
#End Region

#Region "Default settings"
    Private Sub tbDefaultFont_Click(sender As Object, e As System.EventArgs) Handles tbDefaultFont.Click
        Dim fntDialog As New FontDialog

        With fntDialog
            .Font = My.Settings.setDefaultFont
            .FontMustExist = True
            .ShowColor = False
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                My.Settings.setDefaultFont = fntDialog.Font
                tbDefaultFont.DataBindings(0).ReadValue()
            End If
        End With
    End Sub

    Private Sub tbDefaultFontSortNumbers_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDefaultFontSortNumbers.Click
        Dim fntDialog As New FontDialog
        With fntDialog
            .Font = My.Settings.setDefaultFontSortNumbers
            .FontMustExist = True
            .ShowColor = False
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                My.Settings.setDefaultFontSortNumbers = fntDialog.Font
                tbDefaultFontSortNumbers.DataBindings(0).ReadValue()
            End If
        End With

    End Sub

    Private Sub tbDefaultFontUI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDefaultFontUI.Click
        Dim fntDialog As New FontDialog
        With fntDialog
            .Font = My.Settings.setDefaultFontUI
            .FontMustExist = True
            .ShowColor = False
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                My.Settings.setDefaultFontUI = fntDialog.Font
                tbDefaultFontUI.DataBindings(0).ReadValue()
            End If
        End With
    End Sub

    Private Sub tbDefaultFontSections_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDefaultFontSections.Click
        Dim fntDialog As New FontDialog
        With fntDialog
            .Font = My.Settings.setDefaultFontSections
            .FontMustExist = True
            .ShowColor = True
            .Color = My.Settings.setDefaultFontSectionColor
            .MinSize = 6
            .MaxSize = 18
            .AllowSimulations = True

            If .ShowDialog = vbOK Then
                My.Settings.setDefaultFontSections = .Font
                My.Settings.setDefaultFontSectionColor = .Color
                tbDefaultFontSections.DataBindings(0).ReadValue()
                tbDefaultFontSectionColor.ForeColor = My.Settings.setDefaultFontSectionColor
                tbDefaultFontSectionColor.DataBindings(0).ReadValue()

            End If
        End With
    End Sub

    Private Sub tbDefaultFontSectionColor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDefaultFontSectionColor.Click
        Dim colDialog As New ColorDialog
        With colDialog
            .Color = My.Settings.setDefaultFontSectionColor
            .AllowFullOpen = True
            .AnyColor = True
            .CustomColors = New Integer() {Color.LightGray.ToArgb, Color.Gray.ToArgb}
            If .ShowDialog() = vbOK Then
                My.Settings.setDefaultFontSectionColor = .Color
                tbDefaultFontSectionColor.DataBindings(0).ReadValue()
                tbDefaultFontSectionColor.ForeColor = My.Settings.setDefaultFontSectionColor
            End If
        End With
    End Sub

    Private Sub tbDefaultColorSectionTop_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDefaultColorSectionTop.Click
        Dim colDialog As New ColorDialog
        With colDialog
            .Color = My.Settings.setDefaultColorSectionTop
            .AllowFullOpen = True
            .AnyColor = True
            .CustomColors = New Integer() {Color.LightGray.ToArgb, Color.Gray.ToArgb}
            If .ShowDialog() = vbOK Then
                My.Settings.setDefaultColorSectionTop = .Color
                tbDefaultColorSectionTop.DataBindings(0).ReadValue()
                tbDefaultColorSectionTop.BackColor = My.Settings.setDefaultColorSectionTop
            End If
        End With
    End Sub

    Private Sub tbDefaultColorSectionBottom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDefaultColorSectionBottom.Click
        Dim colDialog As New ColorDialog
        With colDialog
            .Color = My.Settings.setDefaultColorSectionBottom
            .AllowFullOpen = True
            .AnyColor = True
            .CustomColors = New Integer() {Color.LightGray.ToArgb, Color.Gray.ToArgb}
            If .ShowDialog() = vbOK Then
                My.Settings.setDefaultColorSectionBottom = .Color
                tbDefaultColorSectionBottom.DataBindings(0).ReadValue()
                tbDefaultColorSectionBottom.BackColor = My.Settings.setDefaultColorSectionBottom
            End If
        End With
    End Sub

    Private Sub btnEqualColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqualColor.Click
        My.Settings.setDefaultColorSectionBottom = My.Settings.setDefaultColorSectionTop
        tbDefaultColorSectionBottom.DataBindings(0).ReadValue()
        tbDefaultColorSectionBottom.BackColor = My.Settings.setDefaultColorSectionBottom
    End Sub

    Private Sub cmbDefaultIndType_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cmbDefaultIndType.SelectedValueChanged
        If cmbDefaultIndType.DataBindings.Count < 1 Then Exit Sub

        If cmbDefaultIndType.SelectedItem Is Nothing Then Exit Sub

        Dim intQuestionType As Integer = CType(cmbDefaultIndType.SelectedItem, StructuredComboBoxItem).Type

        With My.Settings
            If intQuestionType <> .setDefaultIndicatorType Then
                If intQuestionType = CONST_QuestionTypeTitle Then
                    Dim intIndex As Integer = cmbDefaultIndType.SelectedIndex + 1
                    Dim NextItem As StructuredComboBoxItem = TryCast(cmbDefaultIndType.Items(intIndex), StructuredComboBoxItem)

                    If NextItem IsNot Nothing Then
                        intQuestionType = NextItem.Type
                    Else
                        intQuestionType = 0
                    End If
                    .setDefaultIndicatorType = intQuestionType
                    cmbDefaultIndType.DataBindings(0).ReadValue()
                Else
                    cmbDefaultIndType.DataBindings(0).WriteValue()
                End If
            End If
        End With
    End Sub

    Private Sub ntbRecentFilesMax_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ntbRecentFilesMax.Validating
        If ntbRecentFilesMax.SingleValue > 10 Then ntbRecentFilesMax.Text = "10"
    End Sub

    Private Sub ntbTimerAutoSave_Validated(sender As Object, e As System.EventArgs) Handles ntbTimerAutoSave.Validated
        If My.Settings.setTimerAutoSave > 0 Then
            With frmParent.TimerAutoSave
                .Stop()
                .Interval = My.Settings.setTimerAutoSave * 60000
                .Start()
            End With

        End If
    End Sub
#End Region
    
    
    
    
    
    
    
End Class
