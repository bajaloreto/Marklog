Imports System.Drawing.Printing

Public Class ucPrintSettingsAssumptions
    Private intPrintSection As Integer
    Private intPagesWide As Integer
    Public boolLoad As Boolean

    Public Event PrintAssumptionsSetupChanged(ByVal sender As Object, ByVal e As PrintAssumptionsSetupChangedEventArgs)

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        With cmbPrintSection
            .Items.Add(My.Settings.setStruct1)
            .Items.Add(My.Settings.setStruct2)
            .Items.Add(My.Settings.setStruct3)
            .Items.Add(My.Settings.setStruct4)
            .Items.Add(LANG_All)
            .SelectedIndex = My.Settings.setPrintAssumptionsSection
        End With

        nudPagesWide.Value = My.Settings.setPrintAssumptionsPagesWide

        boolLoad = False
    End Sub

    Public Property PrintSection() As Integer
        Get
            Return intPrintSection
        End Get
        Set(ByVal value As Integer)
            intPrintSection = value
            My.Settings.setPrintAssumptionsSection = value
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
        RaiseEvent PrintAssumptionsSetupChanged(Me, New PrintAssumptionsSetupChangedEventArgs(PrintSection, PagesWide))
    End Sub

    Private Sub nudPagesWide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudPagesWide.ValueChanged
        Me.PagesWide = nudPagesWide.Value

        If boolLoad = False Then
            My.Settings.setPrintAssumptionsPagesWide = Me.PagesWide
            RaiseEvent PrintAssumptionsSetupChanged(Me, New PrintAssumptionsSetupChangedEventArgs(PrintSection, PagesWide))
        End If
    End Sub
End Class

Public Class PrintAssumptionsSetupChangedEventArgs
    Inherits EventArgs

    Public Property PrintSection As Integer
    Public Property PagesWide As Integer

    Public Sub New(ByVal intPrintSection As Integer, ByVal intPagesWide As Integer)
        MyBase.New()

        PrintSection = intPrintSection
        PagesWide = intPagesWide
    End Sub
End Class
