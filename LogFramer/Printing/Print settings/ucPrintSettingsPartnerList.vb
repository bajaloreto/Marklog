Imports System.Drawing.Printing

Public Class ucPrintSettingsPartnerList
    Private boolLoad As Boolean

    Public Event PrintPartnerListSetupChanged(ByVal sender As Object, ByVal e As PrintPartnerListSetupChangedEventArgs)

    Public Sub New()
        InitializeComponent()

        boolLoad = True
        Dim strAll As String

        Select Case UserLanguage
            Case "fr"
                strAll = "- tout -"
            Case Else
                strAll = "- all -"
        End Select

        With cmbSelectPartner
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            .Items.AddRange(CurrentLogFrame.PartnerNamesList)
            .Items.RemoveAt(0)
            .Items.Add(strAll)
            If My.Settings.setPrintPartnersIndex <= .Items.Count - 1 Then _
                .SelectedIndex = My.Settings.setPrintPartnersIndex
        End With

        If My.Settings.setPrintPartnersBorders = True Then
            chkPrintBorders.CheckState = CheckState.Checked
        Else
            chkPrintBorders.CheckState = CheckState.Unchecked
        End If

        If My.Settings.setPrintPartnersFill = True Then
            chkFillCells.CheckState = CheckState.Checked
        Else
            chkFillCells.CheckState = CheckState.Unchecked
        End If
        boolLoad = False
    End Sub

    Public Property PartnerIndex() As Integer
        Get
            Return My.Settings.setPrintPartnersIndex
        End Get
        Set(ByVal value As Integer)
            My.Settings.setPrintPartnersIndex = value
        End Set
    End Property

    Public Property PrintBorders As Boolean
        Get
            Return My.Settings.setPrintPartnersBorders
        End Get
        Set(ByVal value As Boolean)
            My.Settings.setPrintPartnersBorders = value
        End Set
    End Property

    Public Property FillCells As Boolean
        Get
            Return My.Settings.setPrintPartnersFill
        End Get
        Set(ByVal value As Boolean)
            My.Settings.setPrintPartnersFill = value
        End Set
    End Property

    Private Sub cmbSelectPartner_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectPartner.SelectedIndexChanged
        If boolLoad = False Then
            PartnerIndex = cmbSelectPartner.SelectedIndex
            RaiseEvent PrintPartnerListSetupChanged(Me, New PrintPartnerListSetupChangedEventArgs(PartnerIndex, PrintBorders, FillCells))
        End If
    End Sub

    Private Sub chkBorders_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPrintBorders.CheckedChanged
        PrintBorders = chkPrintBorders.Checked

        RaiseEvent PrintPartnerListSetupChanged(Me, New PrintPartnerListSetupChangedEventArgs(PartnerIndex, PrintBorders, FillCells))
    End Sub

    Private Sub chkFilled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFillCells.CheckedChanged
        FillCells = chkFillCells.Checked

        RaiseEvent PrintPartnerListSetupChanged(Me, New PrintPartnerListSetupChangedEventArgs(PartnerIndex, PrintBorders, FillCells))
    End Sub
End Class

Public Class PrintPartnerListSetupChangedEventArgs
    Inherits EventArgs

    Public Property PartnerIndex As Integer
    Public Property PrintBorders As Boolean
    Public Property FillCells As Boolean

    Public Sub New(ByVal intPartnerIndex As Integer, ByVal boolPrintBorders As Boolean, ByVal boolFillCells As Boolean)
        MyBase.New()

        PartnerIndex = intPartnerIndex
        PrintBorders = boolPrintBorders
        FillCells = boolFillCells
    End Sub
End Class
