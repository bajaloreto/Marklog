Imports System.Drawing.Printing

Public Class ucPrintSettingsTargetGroupIdForm
    Private strTargetGroupName As String

    Public Event PrintTargetGroupIdFormSetupChanged(ByVal sender As Object, ByVal e As PrintTargetGroupIdFormSetupChangedEventArgs)

    Public Property TargetGroupName() As String
        Get
            Return strTargetGroupName
        End Get
        Set(ByVal value As String)
            strTargetGroupName = value
            My.Settings.setPrintTargetgroupName = value
        End Set
    End Property

    Public Sub New()
        InitializeComponent()

        Dim strAll As String

        Select Case UserLanguage
            Case "fr"
                strAll = "- tout -"
                GroupBoxPurposes.Text = My.Settings.setStruct2sing
            Case Else
                strAll = "- all -"
                GroupBoxPurposes.Text &= My.Settings.setStruct2sing
        End Select

        With cmbSelectPurpose
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
            For i = 0 To CurrentLogFrame.Purposes.Count - 1
                .Items.Add(CurrentLogFrame.Purposes(i).Text)
            Next
            .Items.Add(strAll)
            .SelectedIndex = .Items.Count - 1
        End With

        With cmbSelectTargetGroup
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .DropDownStyle = ComboBoxStyle.DropDownList
        End With

        If My.Settings.setPrintTargetgroupIdBorders = True Then
            chkPrintBorders.CheckState = CheckState.Checked
        Else
            chkPrintBorders.CheckState = CheckState.Unchecked
        End If

        If My.Settings.setPrintTargetgroupIdFill = True Then
            chkFillCells.CheckState = CheckState.Checked
        Else
            chkFillCells.CheckState = CheckState.Unchecked
        End If
    End Sub

    Public ReadOnly Property PartnerIndex() As Integer
        Get
            Return cmbSelectTargetGroup.SelectedIndex
        End Get
    End Property

    Public Property PrintBorders As Boolean
        Get
            Return My.Settings.setPrintTargetgroupIdBorders
        End Get
        Set(ByVal value As Boolean)
            My.Settings.setPrintTargetgroupIdBorders = value
        End Set
    End Property

    Public Property FillCells As Boolean
        Get
            Return My.Settings.setPrintTargetgroupIdFill
        End Get
        Set(ByVal value As Boolean)
            My.Settings.setPrintTargetgroupIdFill = value
        End Set
    End Property

    Private Sub cmbSelectTargetGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectTargetGroup.SelectedIndexChanged
        Me.TargetGroupName = cmbSelectTargetGroup.SelectedItem.ToString
        RaiseEvent PrintTargetGroupIdFormSetupChanged(Me, New PrintTargetGroupIdFormSetupChangedEventArgs(TargetGroupName, PrintBorders, FillCells))
    End Sub

    Private Sub cmbSelectPurpose_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectPurpose.SelectedIndexChanged
        Dim intIndex As Integer = cmbSelectPurpose.SelectedIndex
        Dim selPurpose As Purpose
        Dim strAll As String

        Select Case UserLanguage
            Case "fr"
                strAll = "- tout -"
            Case Else
                strAll = "- all -"
        End Select

        With cmbSelectTargetGroup
            .Items.Clear()
            If intIndex < CurrentLogframe.Purposes.Count - 1 Then
                selPurpose = CurrentLogframe.Purposes(cmbSelectPurpose.SelectedIndex)

                For i = 0 To selPurpose.TargetGroups.Count - 1
                    .Items.Add(selPurpose.TargetGroups(i).Name)
                Next
                .Items.Add(strAll)
            Else
                For j = 0 To CurrentLogframe.Purposes.Count - 1
                    selPurpose = CurrentLogframe.Purposes(j)

                    For i = 0 To selPurpose.TargetGroups.Count - 1
                        .Items.Add(selPurpose.TargetGroups(i).Name)
                    Next
                Next
                .Items.Add(strAll)
            End If
            .SelectedIndex = .Items.Count - 1
        End With
    End Sub

    Private Sub chkBorders_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPrintBorders.CheckedChanged
        PrintBorders = chkPrintBorders.Checked

        RaiseEvent PrintTargetGroupIdFormSetupChanged(Me, New PrintTargetGroupIdFormSetupChangedEventArgs(TargetGroupName, PrintBorders, FillCells))
    End Sub

    Private Sub chkFilled_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFillCells.CheckedChanged
        FillCells = chkFillCells.Checked

        RaiseEvent PrintTargetGroupIdFormSetupChanged(Me, New PrintTargetGroupIdFormSetupChangedEventArgs(TargetGroupName, PrintBorders, FillCells))
    End Sub
End Class

Public Class PrintTargetGroupIdFormSetupChangedEventArgs
    Inherits EventArgs

    Public Property TargetGroupName As String
    Public Property PrintBorders As Boolean
    Public Property FillCells As Boolean

    Public Sub New(ByVal strTargetGroupName As String, ByVal boolPrintBorders As Boolean, ByVal boolFillCells As Boolean)
        MyBase.New()

        TargetGroupName = strTargetGroupName
        PrintBorders = boolPrintBorders
        FillCells = boolFillCells
    End Sub
End Class
