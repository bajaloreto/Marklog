Imports System.Drawing.Printing

Public Class ucPrintSettingsRiskRegister
    Private intRiskCategory As Integer
    Private intPagesWide As Integer
    Public boolLoad As Boolean

    Public Event PrintRiskRegisterSetupChanged(ByVal sender As Object, ByVal e As PrintRiskRegisterSetupChangedEventArgs)

    Public Enum RiskCategories As Integer
        NotDefined = 0
        Operational = 1
        Financial = 2
        Objectives = 3
        Reputation = 4
        Other = 5
        All = 6
    End Enum

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        With cmbRiskCategory
            .Items.AddRange(LIST_RiskCategories)
            .Items.Add(LANG_All)
            .SelectedIndex = My.Settings.setPrintRiskRegisterRiskCategory
        End With

        nudPagesWide.Value = My.Settings.setPrintPmfPagesWide

        boolLoad = False
    End Sub

    Public Property RiskCategory() As Integer
        Get
            Return intRiskCategory
        End Get
        Set(ByVal value As Integer)
            intRiskCategory = value
            My.Settings.setPrintRiskRegisterRiskCategory = value
        End Set
    End Property

    Public Property PagesWide As Integer
        Get
            Return intPagesWide
        End Get
        Set(ByVal value As Integer)
            intPagesWide = value
        End Set
    End Property

    Private Sub cmbPrintLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbRiskCategory.SelectedIndexChanged
        RiskCategory = cmbRiskCategory.SelectedIndex
        RaiseEvent PrintRiskRegisterSetupChanged(Me, New PrintRiskRegisterSetupChangedEventArgs(RiskCategory, PagesWide))
    End Sub

    Private Sub nudPagesWide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudPagesWide.ValueChanged
        Me.PagesWide = nudPagesWide.Value

        If boolLoad = False Then
            My.Settings.setPrintPmfPagesWide = Me.PagesWide
            RaiseEvent PrintRiskRegisterSetupChanged(Me, New PrintRiskRegisterSetupChangedEventArgs(RiskCategory, PagesWide))
        End If
    End Sub
End Class

Public Class PrintRiskRegisterSetupChangedEventArgs
    Inherits EventArgs

    Public Property RiskCategory As Integer
    Public Property PagesWide As Integer

    Public Sub New(ByVal intRiskCategory As Integer, ByVal intPagesWide As Integer)
        MyBase.New()

        RiskCategory = intRiskCategory
        PagesWide = intPagesWide
    End Sub
End Class
