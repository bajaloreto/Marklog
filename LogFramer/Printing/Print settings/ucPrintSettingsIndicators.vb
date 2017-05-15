Imports System.Drawing.Printing

Public Class PrintSettingsIndicators
    Private intPrintSection As Integer
    Private objTargetGroupGuid As Guid
    Private intMeasurement As Integer
    Private boolPrintPurposes As Boolean
    Private boolPrintOutputs As Boolean
    Private boolPrintActivities As Boolean
    Private boolPrintOptionValues As Boolean
    Private boolPrintRanges As Boolean
    Private boolPrintTargets As Boolean
    Private boolLoad As Boolean

    Public Event PrintIndicatorsSetupChanged(ByVal sender As Object, ByVal e As PrintIndicatorsSetupChangedEventArgs)

#Region "Properties"
    Public Enum PrintSections As Integer
        notselected = 0
        printgoals = 1
        printpurposes = 2
        printoutputs = 3
        printactivities = 4
        printall = 5
    End Enum

    Public Property PrintSection() As Integer
        Get
            Return intPrintSection
        End Get
        Set(ByVal value As Integer)
            intPrintSection = value
            My.Settings.setPrintIndicatorsSection = value
        End Set
    End Property

    Public Property TargetGroupGuid() As Guid
        Get
            Return objTargetGroupGuid
        End Get
        Set(ByVal value As Guid)
            objTargetGroupGuid = value
        End Set
    End Property

    Private Property Measurement As Integer
        Get
            Return intMeasurement
        End Get
        Set(ByVal value As Integer)
            intMeasurement = value
        End Set
    End Property

    Private Property PrintPurposes As Boolean
        Get
            Return boolPrintPurposes
        End Get
        Set(ByVal value As Boolean)
            boolPrintPurposes = value
        End Set
    End Property

    Private Property PrintOutputs As Boolean
        Get
            Return boolPrintOutputs
        End Get
        Set(ByVal value As Boolean)
            boolPrintOutputs = value
        End Set
    End Property

    Private Property PrintActivities As Boolean
        Get
            Return boolPrintActivities
        End Get
        Set(ByVal value As Boolean)
            boolPrintActivities = value
        End Set
    End Property

    Private Property PrintOptionValues As Boolean
        Get
            Return boolPrintOptionValues
        End Get
        Set(ByVal value As Boolean)
            boolPrintOptionValues = value
        End Set
    End Property

    Private Property PrintRanges As Boolean
        Get
            Return boolPrintRanges
        End Get
        Set(ByVal value As Boolean)
            boolPrintRanges = value
        End Set
    End Property

    Private Property PrintTargets As Boolean
        Get
            Return boolPrintTargets
        End Get
        Set(ByVal value As Boolean)
            boolPrintTargets = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        InitializeComponent()

        boolLoad = True
        With cmbPrintSection
            .Items.Add(My.Settings.setStruct1)
            .Items.Add(My.Settings.setStruct2)
            .Items.Add(My.Settings.setStruct3)
            .Items.Add(My.Settings.setStruct4)
            .Items.Add(LANG_All)
            .SelectedIndex = My.Settings.setPrintIndicatorsSection - 1
        End With

        Dim ListTargetGroups As Dictionary(Of Guid, String) = CurrentLogFrame.GetTargetGroupsList
        With cmbTargetGroupGuid
            .AutoCompleteMode = AutoCompleteMode.None
            .DropDownStyle = ComboBoxStyle.DropDownList
            .DataSource = New BindingSource(ListTargetGroups, Nothing)
            .ValueMember = "Key"
            .DisplayMember = "Value"
        End With

        Select Case My.Settings.setPrintIndicatorsMeasurement
            Case 0
                rbtnSingleMeasurement.Checked = True
            Case Else
                rbtnAllMeasurements.Checked = True
        End Select

        chkPrintPurposes.Checked = My.Settings.setPrintIndicatorsPrintPurposes
        chkPrintOutputs.Checked = My.Settings.setPrintIndicatorsPrintOutputs
        chkPrintActivities.Checked = My.Settings.setPrintIndicatorsPrintActivities
        chkPrintOptionValues.Checked = My.Settings.setPrintIndicatorsPrintOptionValues
        chkPrintRanges.Checked = My.Settings.setPrintIndicatorsPrintValueRanges
        chkPrintTargets.Checked = My.Settings.setPrintIndicatorsPrintTargets
        boolLoad = False

        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub cmbPrintLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPrintSection.SelectedIndexChanged
        PrintSection = cmbPrintSection.SelectedIndex + 1

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub cmbTargetGroupGuid_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTargetGroupGuid.SelectedValueChanged
        If cmbTargetGroupGuid.SelectedValue.GetType = GetType(Guid) Then
            Dim selGuid As Guid = cmbTargetGroupGuid.SelectedValue
            Me.TargetGroupGuid = selGuid
            My.Settings.setPrintIndicatorsTargetgroupGuid = selGuid

            If boolLoad = False Then _
            RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                                PrintOptionValues, PrintRanges, PrintTargets))
        End If
    End Sub

    Private Sub rbtnSingleMeasurement_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnSingleMeasurement.MouseUp
        If rbtnSingleMeasurement.Checked = True Then
            Me.Measurement = 0
            My.Settings.setPrintIndicatorsMeasurement = 0
        Else
            Me.Measurement = 1
            My.Settings.setPrintIndicatorsMeasurement = 1
        End If

        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub rbtnAllMeasurements_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles rbtnAllMeasurements.MouseUp
        If rbtnAllMeasurements.Checked = True Then
            Me.Measurement = 1
            My.Settings.setPrintIndicatorsMeasurement = 1
        Else
            Me.Measurement = 0
            My.Settings.setPrintIndicatorsMeasurement = 0
        End If

        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub chkPrintPurposes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrintPurposes.CheckedChanged
        Me.PrintPurposes = chkPrintPurposes.Checked
        My.Settings.setPrintIndicatorsPrintPurposes = chkPrintPurposes.Checked

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub chkPrintOutputs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrintOutputs.CheckedChanged
        Me.PrintOutputs = chkPrintOutputs.Checked
        My.Settings.setPrintIndicatorsPrintOutputs = chkPrintOutputs.Checked

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub chkPrintActivities_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPrintActivities.CheckedChanged
        Me.PrintActivities = chkPrintActivities.Checked
        My.Settings.setPrintIndicatorsPrintActivities = chkPrintActivities.Checked

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub chkPrintOptionValues_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrintOptionValues.CheckedChanged
        Me.PrintOptionValues = chkPrintOptionValues.Checked
        My.Settings.setPrintIndicatorsPrintOptionValues = chkPrintOptionValues.Checked

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub chkPrintRanges_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrintRanges.CheckedChanged
        Me.PrintRanges = chkPrintRanges.Checked
        My.Settings.setPrintIndicatorsPrintValueRanges = chkPrintRanges.Checked

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub

    Private Sub chkPrintTargets_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrintTargets.CheckedChanged
        Me.PrintTargets = chkPrintTargets.Checked
        My.Settings.setPrintIndicatorsPrintTargets = chkPrintTargets.Checked

        If boolLoad = False Then _
        RaiseEvent PrintIndicatorsSetupChanged(Me, New PrintIndicatorsSetupChangedEventArgs(PrintSection, TargetGroupGuid, Measurement, PrintPurposes, PrintOutputs, PrintActivities, _
                                                                                            PrintOptionValues, PrintRanges, PrintTargets))
    End Sub
#End Region
End Class

Public Class PrintIndicatorsSetupChangedEventArgs
    Inherits EventArgs

    Public Property PrintSection As Integer
    Public Property TargetGroupGuid As Guid
    Public Property Measurement As Integer
    Public Property PrintPurposes As Boolean
    Public Property PrintOutputs As Boolean
    Public Property PrintActivities As Boolean
    Public Property PrintOptionValues As Boolean
    Public Property PrintValueRanges As Boolean
    Public Property PrintTargets As Boolean

    Public Sub New(ByVal printsection As Integer, ByVal targetgroupguid As Guid, ByVal measurement As Integer, _
                   ByVal printpurposes As Boolean, ByVal printoutputs As Boolean, ByVal printactivities As Boolean, ByVal printoptionvalues As Boolean, _
                   ByVal printvalueranges As Boolean, ByVal printtargets As Boolean)
        MyBase.New()

        Me.PrintSection = printsection
        Me.TargetGroupGuid = targetgroupguid
        Me.Measurement = measurement
        Me.PrintPurposes = printpurposes
        Me.PrintOutputs = printoutputs
        Me.PrintActivities = printactivities
        Me.PrintOptionValues = printoptionvalues
        Me.PrintValueRanges = printvalueranges
        Me.PrintTargets = printtargets
    End Sub
End Class
