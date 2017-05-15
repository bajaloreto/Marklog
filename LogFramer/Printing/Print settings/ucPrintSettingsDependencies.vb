Imports System.Drawing.Printing

Public Class ucPrintSettingsDependencies
    Private intPrintSection As Integer
    Private boolShowDeliverables As Boolean
    Private intPagesWide As Integer
    Public boolLoad As Boolean

    Public Event PrintDependenciesSetupChanged(ByVal sender As Object, ByVal e As PrintDependenciesSetupChangedEventArgs)

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        With cmbPrintSection
            .Items.Add(My.Settings.setStruct1)
            .Items.Add(My.Settings.setStruct2)
            .Items.Add(My.Settings.setStruct3)
            .Items.Add(My.Settings.setStruct4)
            .Items.Add(LANG_All)
            .SelectedIndex = My.Settings.setPrintDependenciesSection
        End With

        chkShowDeliverables.Checked = My.Settings.setPrintDependenciesShowDeliverables

        nudPagesWide.Value = My.Settings.setPrintDependenciesPagesWide

        boolLoad = False
    End Sub

    Public Property PrintSection() As Integer
        Get
            Return intPrintSection
        End Get
        Set(ByVal value As Integer)
            intPrintSection = value
            My.Settings.setPrintDependenciesSection = value
        End Set
    End Property

    Public Property ShowDeliverables As Boolean
        Get
            Return boolShowDeliverables
        End Get
        Set(ByVal value As Boolean)
            boolShowDeliverables = value
            My.Settings.setPrintDependenciesShowDeliverables = value
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

    Private Sub cmbPrintLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPrintSection.SelectedIndexChanged
        PrintSection = cmbPrintSection.SelectedIndex
        RaiseEvent PrintDependenciesSetupChanged(Me, New PrintDependenciesSetupChangedEventArgs(PrintSection, ShowDeliverables, PagesWide))
    End Sub

    Private Sub chkShowDeliverables_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDeliverables.CheckedChanged
        ShowDeliverables = chkShowDeliverables.Checked
        RaiseEvent PrintDependenciesSetupChanged(Me, New PrintDependenciesSetupChangedEventArgs(PrintSection, ShowDeliverables, PagesWide))
    End Sub

    Private Sub nudPagesWide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudPagesWide.ValueChanged
        Me.PagesWide = nudPagesWide.Value

        If boolLoad = False Then
            My.Settings.setPrintDependenciesPagesWide = Me.PagesWide
            RaiseEvent PrintDependenciesSetupChanged(Me, New PrintDependenciesSetupChangedEventArgs(PrintSection, ShowDeliverables, PagesWide))
        End If
    End Sub
End Class

Public Class PrintDependenciesSetupChangedEventArgs
    Inherits EventArgs

    Public Property PrintSection As Integer
    Public Property ShowDeliverables As Boolean
    Public Property PagesWide As Integer

    Public Sub New(ByVal intPrintSection As Integer, ByVal boolShowDeliverables As Boolean, ByVal intPagesWide As Integer)
        MyBase.New()

        PrintSection = intPrintSection
        ShowDeliverables = boolShowDeliverables
        PagesWide = intPagesWide
    End Sub
End Class
