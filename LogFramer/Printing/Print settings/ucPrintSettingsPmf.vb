Imports System.Drawing.Printing

Public Class PrintSettingsPmf
    Private intPrintSection As Integer
    Private boolShowTargetRowTitles As Boolean
    Private intPagesWide As Integer
    Public boolLoad As Boolean

    Public Event PrintPmfSetupChanged(ByVal sender As Object, ByVal e As PrintPmfSetupChangedEventArgs)

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        With cmbPrintSection
            .Items.Add(My.Settings.setStruct1)
            .Items.Add(My.Settings.setStruct2)
            .Items.Add(My.Settings.setStruct3)
            .Items.Add(My.Settings.setStruct4)
            .Items.Add(LANG_All)
            .SelectedIndex = My.Settings.setPrintPmfSection
        End With
        chkShowTargetRowTitles.Checked = My.Settings.setPrintPmfShowTargetRowTitles
        nudPagesWide.Value = My.Settings.setPrintPmfPagesWide

        boolLoad = False
    End Sub

    Public Property PrintSection() As Integer
        Get
            Return intPrintSection
        End Get
        Set(ByVal value As Integer)
            intPrintSection = value
            My.Settings.setPrintPmfSection = value
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

    Public Property ShowTargetRowTitles As Boolean
        Get
            Return boolShowTargetRowTitles
        End Get
        Set(ByVal value As Boolean)
            boolShowTargetRowTitles = value
        End Set
    End Property

    Public Enum PrintSections As Integer
        notselected = -1
        printgoals = 0
        printpurposes = 1
        printoutputs = 2
        printactivities = 3
        printall = 4
    End Enum

    Private Sub cmbPrintSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPrintSection.SelectedIndexChanged
        PrintSection = cmbPrintSection.SelectedIndex

        RaiseEvent PrintPmfSetupChanged(Me, New PrintPmfSetupChangedEventArgs(PrintSection, PagesWide, ShowTargetRowTitles))
    End Sub

    Private Sub chkShowTargetRowTitles_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowTargetRowTitles.CheckedChanged
        Me.ShowTargetRowTitles = chkShowTargetRowTitles.Checked

        If boolLoad = False Then
            My.Settings.setPrintPmfShowTargetRowTitles = Me.ShowTargetRowTitles
            RaiseEvent PrintPmfSetupChanged(Me, New PrintPmfSetupChangedEventArgs(PrintSection, PagesWide, ShowTargetRowTitles))
        End If
    End Sub

    Private Sub nudPagesWide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudPagesWide.ValueChanged
        Me.PagesWide = nudPagesWide.Value

        If boolLoad = False Then
            My.Settings.setPrintPmfPagesWide = Me.PagesWide
            RaiseEvent PrintPmfSetupChanged(Me, New PrintPmfSetupChangedEventArgs(PrintSection, PagesWide, ShowTargetRowTitles))
        End If
    End Sub
End Class

Public Class PrintPmfSetupChangedEventArgs
    Inherits EventArgs

    Public Property PrintSection As Integer
    Public Property PagesWide As Integer
    Public Property ShowTargetRowTitles As Boolean

    Public Sub New(ByVal intPrintSection As Integer, ByVal intPagesWide As Integer, ByVal boolShowTargetRowTitles As Boolean)
        MyBase.New()

        PrintSection = intPrintSection
        PagesWide = intPagesWide
        ShowTargetRowTitles = boolShowTargetRowTitles
    End Sub
End Class
