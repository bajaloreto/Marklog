Imports System.Windows.Forms

Public Class DialogExportToExcel
    Private objCurrentExportSettings As Object

    Public Enum Reports
        Logframe
        Budget
        MonitoringTool
        Pmf
        RiskRegister
    End Enum

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Property CurrentExportSettings As Object
        Get
            Return objCurrentExportSettings
        End Get
        Set(value As Object)
            objCurrentExportSettings = value
        End Set
    End Property


    Public Sub New(ByVal report As Integer)
        InitializeComponent()

        Dim intSettingsHeight As Integer

        lblWarning.Text = String.Empty

        Select Case report
            Case Reports.Logframe
                Dim objExportLogframeSettings As New PrintSettingsLogframe

                With objExportLogframeSettings
                    .ShowIndicatorColumn = CurrentProjectForm.dgvLogframe.ShowIndicatorColumn
                    .ShowVerificationSourceColumn = CurrentProjectForm.dgvLogframe.ShowVerificationSourceColumn
                    .ShowAssumptionColumn = CurrentProjectForm.dgvLogframe.ShowAssumptionColumn
                    .ShowGoals = CurrentProjectForm.dgvLogframe.ShowGoals
                    .ShowPurposes = CurrentProjectForm.dgvLogframe.ShowPurposes
                    .ShowOutputs = CurrentProjectForm.dgvLogframe.ShowOutputs
                    .ShowActivities = CurrentProjectForm.dgvLogframe.ShowActivities
                    .ShowResourcesBudget = CurrentProjectForm.dgvLogframe.ShowResourcesBudget
                    intSettingsHeight = .Height
                    .Dock = DockStyle.Fill
                End With
                CurrentExportSettings = objExportLogframeSettings

                PanelSettings.Controls.Add(objExportLogframeSettings)
            Case Reports.Budget
                Dim objExportBudgetSettings As New PrintSettingsBudget

                With objExportBudgetSettings
                    intSettingsHeight = .Height
                    .GroupBoxShow.Visible = False
                    .GroupBoxWidth.Visible = False
                    .Dock = DockStyle.Fill
                End With
                CurrentExportSettings = objExportBudgetSettings

                PanelSettings.Controls.Add(objExportBudgetSettings)
            Case Reports.MonitoringTool
                Dim objExportIndicatorSettings As New PrintSettingsIndicators

                With objExportIndicatorSettings
                    intSettingsHeight = .Height
                    .GroupBoxMeasurement.Visible = False
                    .GroupBoxInclude.Visible = False
                    .Dock = DockStyle.Fill
                End With
                CurrentExportSettings = objExportIndicatorSettings

                PanelSettings.Controls.Add(objExportIndicatorSettings)
                lblWarning.Text = LANG_ExportToExcelWarning
            Case Reports.Pmf
                Dim objExportPmfSettings As New PrintSettingsPmf

                With objExportPmfSettings
                    intSettingsHeight = .Height
                    .GroupBoxWidth.Visible = False
                    .Dock = DockStyle.Fill
                End With
                CurrentExportSettings = objExportPmfSettings

                PanelSettings.Controls.Add(objExportPmfSettings)
            Case Reports.RiskRegister
                Dim objExportRiskRegisterSettings As New ucPrintSettingsRiskRegister

                With objExportRiskRegisterSettings
                    intSettingsHeight = .Height
                    .GroupBoxWidth.Visible = False
                    .Dock = DockStyle.Fill
                End With
                CurrentExportSettings = objExportRiskRegisterSettings

                PanelSettings.Controls.Add(objExportRiskRegisterSettings)
        End Select

        If PanelSettings.Controls.Count > 0 Then
            Dim intHeightDiff As Integer = intSettingsHeight - PanelSettings.Height
            Me.Height += intHeightDiff
            PanelSettings.Height = intSettingsHeight
            lblWarning.Location = New Point(lblWarning.Left, lblWarning.Top + intHeightDiff)
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
