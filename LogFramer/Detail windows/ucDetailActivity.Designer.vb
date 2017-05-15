<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailActivity
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailActivity))
        Me.lblExactEndDate = New System.Windows.Forms.Label()
        Me.lblExactStartDate = New System.Windows.Forms.Label()
        Me.lblActivity = New System.Windows.Forms.Label()
        Me.TabControlActivity = New System.Windows.Forms.TabControl()
        Me.TabPageDuration = New System.Windows.Forms.TabPage()
        Me.lblEndDateMainActivity = New System.Windows.Forms.Label()
        Me.dtbEndDateMainActivity = New FaciliDev.LogFramer.DateTextBox()
        Me.lblStartDateMainActivity = New System.Windows.Forms.Label()
        Me.dtbStartDateMainActivity = New FaciliDev.LogFramer.DateTextBox()
        Me.chkDurationUntilEnd = New FaciliDev.LogFramer.CheckBoxLF()
        Me.ntbDuration = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.cmbDurationUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.lblDuration = New System.Windows.Forms.Label()
        Me.dtbStartDate = New System.Windows.Forms.DateTimePicker()
        Me.cmbGuidReferenceMoment = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbPeriodDirection = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbPeriodUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.ntbPeriod = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.rbtnReferenceDate = New System.Windows.Forms.RadioButton()
        Me.rbtnExactDate = New System.Windows.Forms.RadioButton()
        Me.lblStart = New System.Windows.Forms.Label()
        Me.TabPageOrganisation = New System.Windows.Forms.TabPage()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblOrganisation = New System.Windows.Forms.Label()
        Me.cmbType = New FaciliDev.LogFramer.ComboBoxSelectValue()
        Me.cmbLocation = New FaciliDev.LogFramer.ComboBoxText()
        Me.cmbOrganisation = New FaciliDev.LogFramer.ComboBoxText()
        Me.TabPagePreparation = New System.Windows.Forms.TabPage()
        Me.lblStartDateFollowUp = New System.Windows.Forms.Label()
        Me.dtbStartDateFollowUp = New FaciliDev.LogFramer.DateTextBox()
        Me.lblEndDatePreparation = New System.Windows.Forms.Label()
        Me.dtbEndDatePreparation = New FaciliDev.LogFramer.DateTextBox()
        Me.lblEndDateFollowUp = New System.Windows.Forms.Label()
        Me.lblStartDatePreparation = New System.Windows.Forms.Label()
        Me.lblFollowUp = New System.Windows.Forms.Label()
        Me.lblPreparation = New System.Windows.Forms.Label()
        Me.dtbEndDateFollowUp = New FaciliDev.LogFramer.DateTextBox()
        Me.chkFollowUpRepeat = New FaciliDev.LogFramer.CheckBoxLF()
        Me.chkPreparationRepeat = New FaciliDev.LogFramer.CheckBoxLF()
        Me.dtbStartDatePreparation = New FaciliDev.LogFramer.DateTextBox()
        Me.chkFollowUpUntilEnd = New FaciliDev.LogFramer.CheckBoxLF()
        Me.chkPreparationFromStart = New FaciliDev.LogFramer.CheckBoxLF()
        Me.ntbFollowUp = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.ntbPreparation = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.cmbFollowUpUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.cmbPreparationUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.TabPageRepeat = New System.Windows.Forms.TabPage()
        Me.lblRepeatDates = New System.Windows.Forms.Label()
        Me.lvRepeatDates = New System.Windows.Forms.ListView()
        Me.lblTimes = New System.Windows.Forms.Label()
        Me.lblNumberRepeats = New System.Windows.Forms.Label()
        Me.lblRepeat = New System.Windows.Forms.Label()
        Me.ntbRepeatTimes = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.chkRepeatUntilEnd = New FaciliDev.LogFramer.CheckBoxLF()
        Me.ntbRepeatEvery = New FaciliDev.LogFramer.NumericTextBoxLF()
        Me.cmbRepeatUnit = New FaciliDev.LogFramer.ComboBoxSelectIndex()
        Me.TabPageChildActivities = New System.Windows.Forms.TabPage()
        Me.lvDetailProcesses = New FaciliDev.LogFramer.ListViewProcesses()
        Me.dtbExactEndDate = New FaciliDev.LogFramer.DateTextBox()
        Me.tbActivity = New FaciliDev.LogFramer.TextBoxLF()
        Me.dtbExactStartDate = New FaciliDev.LogFramer.DateTextBox()
        Me.TabControlActivity.SuspendLayout()
        Me.TabPageDuration.SuspendLayout()
        Me.TabPageOrganisation.SuspendLayout()
        Me.TabPagePreparation.SuspendLayout()
        Me.TabPageRepeat.SuspendLayout()
        Me.TabPageChildActivities.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblExactEndDate
        '
        resources.ApplyResources(Me.lblExactEndDate, "lblExactEndDate")
        Me.lblExactEndDate.Name = "lblExactEndDate"
        '
        'lblExactStartDate
        '
        resources.ApplyResources(Me.lblExactStartDate, "lblExactStartDate")
        Me.lblExactStartDate.Name = "lblExactStartDate"
        '
        'lblActivity
        '
        resources.ApplyResources(Me.lblActivity, "lblActivity")
        Me.lblActivity.Name = "lblActivity"
        '
        'TabControlActivity
        '
        resources.ApplyResources(Me.TabControlActivity, "TabControlActivity")
        Me.TabControlActivity.Controls.Add(Me.TabPageDuration)
        Me.TabControlActivity.Controls.Add(Me.TabPageOrganisation)
        Me.TabControlActivity.Controls.Add(Me.TabPagePreparation)
        Me.TabControlActivity.Controls.Add(Me.TabPageRepeat)
        Me.TabControlActivity.Controls.Add(Me.TabPageChildActivities)
        Me.TabControlActivity.Name = "TabControlActivity"
        Me.TabControlActivity.SelectedIndex = 0
        '
        'TabPageDuration
        '
        resources.ApplyResources(Me.TabPageDuration, "TabPageDuration")
        Me.TabPageDuration.Controls.Add(Me.lblEndDateMainActivity)
        Me.TabPageDuration.Controls.Add(Me.dtbEndDateMainActivity)
        Me.TabPageDuration.Controls.Add(Me.lblStartDateMainActivity)
        Me.TabPageDuration.Controls.Add(Me.dtbStartDateMainActivity)
        Me.TabPageDuration.Controls.Add(Me.chkDurationUntilEnd)
        Me.TabPageDuration.Controls.Add(Me.ntbDuration)
        Me.TabPageDuration.Controls.Add(Me.cmbDurationUnit)
        Me.TabPageDuration.Controls.Add(Me.lblDuration)
        Me.TabPageDuration.Controls.Add(Me.dtbStartDate)
        Me.TabPageDuration.Controls.Add(Me.cmbGuidReferenceMoment)
        Me.TabPageDuration.Controls.Add(Me.cmbPeriodDirection)
        Me.TabPageDuration.Controls.Add(Me.cmbPeriodUnit)
        Me.TabPageDuration.Controls.Add(Me.ntbPeriod)
        Me.TabPageDuration.Controls.Add(Me.rbtnReferenceDate)
        Me.TabPageDuration.Controls.Add(Me.rbtnExactDate)
        Me.TabPageDuration.Controls.Add(Me.lblStart)
        Me.TabPageDuration.Name = "TabPageDuration"
        Me.TabPageDuration.UseVisualStyleBackColor = True
        '
        'lblEndDateMainActivity
        '
        resources.ApplyResources(Me.lblEndDateMainActivity, "lblEndDateMainActivity")
        Me.lblEndDateMainActivity.Name = "lblEndDateMainActivity"
        '
        'dtbEndDateMainActivity
        '
        resources.ApplyResources(Me.dtbEndDateMainActivity, "dtbEndDateMainActivity")
        Me.dtbEndDateMainActivity.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbEndDateMainActivity.ForeColor = System.Drawing.Color.Black
        Me.dtbEndDateMainActivity.Name = "dtbEndDateMainActivity"
        Me.dtbEndDateMainActivity.ReadOnly = True
        Me.dtbEndDateMainActivity.TabStop = False
        '
        'lblStartDateMainActivity
        '
        resources.ApplyResources(Me.lblStartDateMainActivity, "lblStartDateMainActivity")
        Me.lblStartDateMainActivity.Name = "lblStartDateMainActivity"
        '
        'dtbStartDateMainActivity
        '
        resources.ApplyResources(Me.dtbStartDateMainActivity, "dtbStartDateMainActivity")
        Me.dtbStartDateMainActivity.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbStartDateMainActivity.ForeColor = System.Drawing.Color.Black
        Me.dtbStartDateMainActivity.Name = "dtbStartDateMainActivity"
        Me.dtbStartDateMainActivity.ReadOnly = True
        Me.dtbStartDateMainActivity.TabStop = False
        '
        'chkDurationUntilEnd
        '
        resources.ApplyResources(Me.chkDurationUntilEnd, "chkDurationUntilEnd")
        Me.chkDurationUntilEnd.Name = "chkDurationUntilEnd"
        Me.chkDurationUntilEnd.Type = 0
        Me.chkDurationUntilEnd.UseVisualStyleBackColor = True
        '
        'ntbDuration
        '
        resources.ApplyResources(Me.ntbDuration, "ntbDuration")
        Me.ntbDuration.AllowSpace = True
        Me.ntbDuration.IsCurrency = False
        Me.ntbDuration.IsPercentage = False
        Me.ntbDuration.Name = "ntbDuration"
        '
        'cmbDurationUnit
        '
        resources.ApplyResources(Me.cmbDurationUnit, "cmbDurationUnit")
        Me.cmbDurationUnit.FormattingEnabled = True
        Me.cmbDurationUnit.Name = "cmbDurationUnit"
        '
        'lblDuration
        '
        resources.ApplyResources(Me.lblDuration, "lblDuration")
        Me.lblDuration.Name = "lblDuration"
        '
        'dtbStartDate
        '
        resources.ApplyResources(Me.dtbStartDate, "dtbStartDate")
        Me.dtbStartDate.Name = "dtbStartDate"
        '
        'cmbGuidReferenceMoment
        '
        resources.ApplyResources(Me.cmbGuidReferenceMoment, "cmbGuidReferenceMoment")
        Me.cmbGuidReferenceMoment.FormattingEnabled = True
        Me.cmbGuidReferenceMoment.Name = "cmbGuidReferenceMoment"
        '
        'cmbPeriodDirection
        '
        resources.ApplyResources(Me.cmbPeriodDirection, "cmbPeriodDirection")
        Me.cmbPeriodDirection.FormattingEnabled = True
        Me.cmbPeriodDirection.Name = "cmbPeriodDirection"
        '
        'cmbPeriodUnit
        '
        resources.ApplyResources(Me.cmbPeriodUnit, "cmbPeriodUnit")
        Me.cmbPeriodUnit.FormattingEnabled = True
        Me.cmbPeriodUnit.Name = "cmbPeriodUnit"
        '
        'ntbPeriod
        '
        resources.ApplyResources(Me.ntbPeriod, "ntbPeriod")
        Me.ntbPeriod.AllowSpace = True
        Me.ntbPeriod.IsCurrency = False
        Me.ntbPeriod.IsPercentage = False
        Me.ntbPeriod.Name = "ntbPeriod"
        '
        'rbtnReferenceDate
        '
        resources.ApplyResources(Me.rbtnReferenceDate, "rbtnReferenceDate")
        Me.rbtnReferenceDate.Checked = True
        Me.rbtnReferenceDate.Name = "rbtnReferenceDate"
        Me.rbtnReferenceDate.TabStop = True
        Me.rbtnReferenceDate.UseVisualStyleBackColor = True
        '
        'rbtnExactDate
        '
        resources.ApplyResources(Me.rbtnExactDate, "rbtnExactDate")
        Me.rbtnExactDate.Checked = True
        Me.rbtnExactDate.Name = "rbtnExactDate"
        Me.rbtnExactDate.UseVisualStyleBackColor = True
        '
        'lblStart
        '
        resources.ApplyResources(Me.lblStart, "lblStart")
        Me.lblStart.Name = "lblStart"
        '
        'TabPageOrganisation
        '
        resources.ApplyResources(Me.TabPageOrganisation, "TabPageOrganisation")
        Me.TabPageOrganisation.Controls.Add(Me.lblType)
        Me.TabPageOrganisation.Controls.Add(Me.lblLocation)
        Me.TabPageOrganisation.Controls.Add(Me.lblOrganisation)
        Me.TabPageOrganisation.Controls.Add(Me.cmbType)
        Me.TabPageOrganisation.Controls.Add(Me.cmbLocation)
        Me.TabPageOrganisation.Controls.Add(Me.cmbOrganisation)
        Me.TabPageOrganisation.Name = "TabPageOrganisation"
        Me.TabPageOrganisation.UseVisualStyleBackColor = True
        '
        'lblType
        '
        resources.ApplyResources(Me.lblType, "lblType")
        Me.lblType.Name = "lblType"
        '
        'lblLocation
        '
        resources.ApplyResources(Me.lblLocation, "lblLocation")
        Me.lblLocation.Name = "lblLocation"
        '
        'lblOrganisation
        '
        resources.ApplyResources(Me.lblOrganisation, "lblOrganisation")
        Me.lblOrganisation.Name = "lblOrganisation"
        '
        'cmbType
        '
        resources.ApplyResources(Me.cmbType, "cmbType")
        Me.cmbType.FormattingEnabled = True
        Me.cmbType.Name = "cmbType"
        '
        'cmbLocation
        '
        resources.ApplyResources(Me.cmbLocation, "cmbLocation")
        Me.cmbLocation.FormattingEnabled = True
        Me.cmbLocation.Name = "cmbLocation"
        '
        'cmbOrganisation
        '
        resources.ApplyResources(Me.cmbOrganisation, "cmbOrganisation")
        Me.cmbOrganisation.FormattingEnabled = True
        Me.cmbOrganisation.Name = "cmbOrganisation"
        '
        'TabPagePreparation
        '
        resources.ApplyResources(Me.TabPagePreparation, "TabPagePreparation")
        Me.TabPagePreparation.Controls.Add(Me.lblStartDateFollowUp)
        Me.TabPagePreparation.Controls.Add(Me.dtbStartDateFollowUp)
        Me.TabPagePreparation.Controls.Add(Me.lblEndDatePreparation)
        Me.TabPagePreparation.Controls.Add(Me.dtbEndDatePreparation)
        Me.TabPagePreparation.Controls.Add(Me.lblEndDateFollowUp)
        Me.TabPagePreparation.Controls.Add(Me.lblStartDatePreparation)
        Me.TabPagePreparation.Controls.Add(Me.lblFollowUp)
        Me.TabPagePreparation.Controls.Add(Me.lblPreparation)
        Me.TabPagePreparation.Controls.Add(Me.dtbEndDateFollowUp)
        Me.TabPagePreparation.Controls.Add(Me.chkFollowUpRepeat)
        Me.TabPagePreparation.Controls.Add(Me.chkPreparationRepeat)
        Me.TabPagePreparation.Controls.Add(Me.dtbStartDatePreparation)
        Me.TabPagePreparation.Controls.Add(Me.chkFollowUpUntilEnd)
        Me.TabPagePreparation.Controls.Add(Me.chkPreparationFromStart)
        Me.TabPagePreparation.Controls.Add(Me.ntbFollowUp)
        Me.TabPagePreparation.Controls.Add(Me.ntbPreparation)
        Me.TabPagePreparation.Controls.Add(Me.cmbFollowUpUnit)
        Me.TabPagePreparation.Controls.Add(Me.cmbPreparationUnit)
        Me.TabPagePreparation.Name = "TabPagePreparation"
        Me.TabPagePreparation.UseVisualStyleBackColor = True
        '
        'lblStartDateFollowUp
        '
        resources.ApplyResources(Me.lblStartDateFollowUp, "lblStartDateFollowUp")
        Me.lblStartDateFollowUp.Name = "lblStartDateFollowUp"
        '
        'dtbStartDateFollowUp
        '
        resources.ApplyResources(Me.dtbStartDateFollowUp, "dtbStartDateFollowUp")
        Me.dtbStartDateFollowUp.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbStartDateFollowUp.ForeColor = System.Drawing.Color.Black
        Me.dtbStartDateFollowUp.Name = "dtbStartDateFollowUp"
        Me.dtbStartDateFollowUp.ReadOnly = True
        Me.dtbStartDateFollowUp.TabStop = False
        '
        'lblEndDatePreparation
        '
        resources.ApplyResources(Me.lblEndDatePreparation, "lblEndDatePreparation")
        Me.lblEndDatePreparation.Name = "lblEndDatePreparation"
        '
        'dtbEndDatePreparation
        '
        resources.ApplyResources(Me.dtbEndDatePreparation, "dtbEndDatePreparation")
        Me.dtbEndDatePreparation.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbEndDatePreparation.ForeColor = System.Drawing.Color.Black
        Me.dtbEndDatePreparation.Name = "dtbEndDatePreparation"
        Me.dtbEndDatePreparation.ReadOnly = True
        Me.dtbEndDatePreparation.TabStop = False
        '
        'lblEndDateFollowUp
        '
        resources.ApplyResources(Me.lblEndDateFollowUp, "lblEndDateFollowUp")
        Me.lblEndDateFollowUp.Name = "lblEndDateFollowUp"
        '
        'lblStartDatePreparation
        '
        resources.ApplyResources(Me.lblStartDatePreparation, "lblStartDatePreparation")
        Me.lblStartDatePreparation.Name = "lblStartDatePreparation"
        '
        'lblFollowUp
        '
        resources.ApplyResources(Me.lblFollowUp, "lblFollowUp")
        Me.lblFollowUp.Name = "lblFollowUp"
        '
        'lblPreparation
        '
        resources.ApplyResources(Me.lblPreparation, "lblPreparation")
        Me.lblPreparation.Name = "lblPreparation"
        '
        'dtbEndDateFollowUp
        '
        resources.ApplyResources(Me.dtbEndDateFollowUp, "dtbEndDateFollowUp")
        Me.dtbEndDateFollowUp.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbEndDateFollowUp.ForeColor = System.Drawing.Color.Black
        Me.dtbEndDateFollowUp.Name = "dtbEndDateFollowUp"
        Me.dtbEndDateFollowUp.ReadOnly = True
        Me.dtbEndDateFollowUp.TabStop = False
        '
        'chkFollowUpRepeat
        '
        resources.ApplyResources(Me.chkFollowUpRepeat, "chkFollowUpRepeat")
        Me.chkFollowUpRepeat.Name = "chkFollowUpRepeat"
        Me.chkFollowUpRepeat.Type = 0
        Me.chkFollowUpRepeat.UseVisualStyleBackColor = True
        '
        'chkPreparationRepeat
        '
        resources.ApplyResources(Me.chkPreparationRepeat, "chkPreparationRepeat")
        Me.chkPreparationRepeat.Name = "chkPreparationRepeat"
        Me.chkPreparationRepeat.Type = 0
        Me.chkPreparationRepeat.UseVisualStyleBackColor = True
        '
        'dtbStartDatePreparation
        '
        resources.ApplyResources(Me.dtbStartDatePreparation, "dtbStartDatePreparation")
        Me.dtbStartDatePreparation.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbStartDatePreparation.ForeColor = System.Drawing.Color.Black
        Me.dtbStartDatePreparation.Name = "dtbStartDatePreparation"
        Me.dtbStartDatePreparation.ReadOnly = True
        Me.dtbStartDatePreparation.TabStop = False
        '
        'chkFollowUpUntilEnd
        '
        resources.ApplyResources(Me.chkFollowUpUntilEnd, "chkFollowUpUntilEnd")
        Me.chkFollowUpUntilEnd.Name = "chkFollowUpUntilEnd"
        Me.chkFollowUpUntilEnd.Type = 0
        Me.chkFollowUpUntilEnd.UseVisualStyleBackColor = True
        '
        'chkPreparationFromStart
        '
        resources.ApplyResources(Me.chkPreparationFromStart, "chkPreparationFromStart")
        Me.chkPreparationFromStart.Name = "chkPreparationFromStart"
        Me.chkPreparationFromStart.Type = 0
        Me.chkPreparationFromStart.UseVisualStyleBackColor = True
        '
        'ntbFollowUp
        '
        resources.ApplyResources(Me.ntbFollowUp, "ntbFollowUp")
        Me.ntbFollowUp.AllowSpace = True
        Me.ntbFollowUp.IsCurrency = False
        Me.ntbFollowUp.IsPercentage = False
        Me.ntbFollowUp.Name = "ntbFollowUp"
        '
        'ntbPreparation
        '
        resources.ApplyResources(Me.ntbPreparation, "ntbPreparation")
        Me.ntbPreparation.AllowSpace = True
        Me.ntbPreparation.IsCurrency = False
        Me.ntbPreparation.IsPercentage = False
        Me.ntbPreparation.Name = "ntbPreparation"
        '
        'cmbFollowUpUnit
        '
        resources.ApplyResources(Me.cmbFollowUpUnit, "cmbFollowUpUnit")
        Me.cmbFollowUpUnit.FormattingEnabled = True
        Me.cmbFollowUpUnit.Name = "cmbFollowUpUnit"
        '
        'cmbPreparationUnit
        '
        resources.ApplyResources(Me.cmbPreparationUnit, "cmbPreparationUnit")
        Me.cmbPreparationUnit.FormattingEnabled = True
        Me.cmbPreparationUnit.Name = "cmbPreparationUnit"
        '
        'TabPageRepeat
        '
        resources.ApplyResources(Me.TabPageRepeat, "TabPageRepeat")
        Me.TabPageRepeat.Controls.Add(Me.lblRepeatDates)
        Me.TabPageRepeat.Controls.Add(Me.lvRepeatDates)
        Me.TabPageRepeat.Controls.Add(Me.lblTimes)
        Me.TabPageRepeat.Controls.Add(Me.lblNumberRepeats)
        Me.TabPageRepeat.Controls.Add(Me.lblRepeat)
        Me.TabPageRepeat.Controls.Add(Me.ntbRepeatTimes)
        Me.TabPageRepeat.Controls.Add(Me.chkRepeatUntilEnd)
        Me.TabPageRepeat.Controls.Add(Me.ntbRepeatEvery)
        Me.TabPageRepeat.Controls.Add(Me.cmbRepeatUnit)
        Me.TabPageRepeat.Name = "TabPageRepeat"
        Me.TabPageRepeat.UseVisualStyleBackColor = True
        '
        'lblRepeatDates
        '
        resources.ApplyResources(Me.lblRepeatDates, "lblRepeatDates")
        Me.lblRepeatDates.Name = "lblRepeatDates"
        '
        'lvRepeatDates
        '
        resources.ApplyResources(Me.lvRepeatDates, "lvRepeatDates")
        Me.lvRepeatDates.Name = "lvRepeatDates"
        Me.lvRepeatDates.UseCompatibleStateImageBehavior = False
        '
        'lblTimes
        '
        resources.ApplyResources(Me.lblTimes, "lblTimes")
        Me.lblTimes.Name = "lblTimes"
        '
        'lblNumberRepeats
        '
        resources.ApplyResources(Me.lblNumberRepeats, "lblNumberRepeats")
        Me.lblNumberRepeats.Name = "lblNumberRepeats"
        '
        'lblRepeat
        '
        resources.ApplyResources(Me.lblRepeat, "lblRepeat")
        Me.lblRepeat.Name = "lblRepeat"
        '
        'ntbRepeatTimes
        '
        resources.ApplyResources(Me.ntbRepeatTimes, "ntbRepeatTimes")
        Me.ntbRepeatTimes.AllowSpace = True
        Me.ntbRepeatTimes.IsCurrency = False
        Me.ntbRepeatTimes.IsPercentage = False
        Me.ntbRepeatTimes.Name = "ntbRepeatTimes"
        '
        'chkRepeatUntilEnd
        '
        resources.ApplyResources(Me.chkRepeatUntilEnd, "chkRepeatUntilEnd")
        Me.chkRepeatUntilEnd.Name = "chkRepeatUntilEnd"
        Me.chkRepeatUntilEnd.Type = 0
        Me.chkRepeatUntilEnd.UseVisualStyleBackColor = True
        '
        'ntbRepeatEvery
        '
        resources.ApplyResources(Me.ntbRepeatEvery, "ntbRepeatEvery")
        Me.ntbRepeatEvery.AllowSpace = True
        Me.ntbRepeatEvery.IsCurrency = False
        Me.ntbRepeatEvery.IsPercentage = False
        Me.ntbRepeatEvery.Name = "ntbRepeatEvery"
        '
        'cmbRepeatUnit
        '
        resources.ApplyResources(Me.cmbRepeatUnit, "cmbRepeatUnit")
        Me.cmbRepeatUnit.FormattingEnabled = True
        Me.cmbRepeatUnit.Name = "cmbRepeatUnit"
        '
        'TabPageChildActivities
        '
        resources.ApplyResources(Me.TabPageChildActivities, "TabPageChildActivities")
        Me.TabPageChildActivities.Controls.Add(Me.lvDetailProcesses)
        Me.TabPageChildActivities.Name = "TabPageChildActivities"
        Me.TabPageChildActivities.UseVisualStyleBackColor = True
        '
        'lvDetailProcesses
        '
        resources.ApplyResources(Me.lvDetailProcesses, "lvDetailProcesses")
        Me.lvDetailProcesses.Activities = Nothing
        Me.lvDetailProcesses.FullRowSelect = True
        Me.lvDetailProcesses.Name = "lvDetailProcesses"
        Me.lvDetailProcesses.OutputText = Nothing
        Me.lvDetailProcesses.SortColumnIndex = -1
        Me.lvDetailProcesses.UseCompatibleStateImageBehavior = False
        Me.lvDetailProcesses.View = System.Windows.Forms.View.Details
        '
        'dtbExactEndDate
        '
        resources.ApplyResources(Me.dtbExactEndDate, "dtbExactEndDate")
        Me.dtbExactEndDate.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbExactEndDate.ForeColor = System.Drawing.Color.Black
        Me.dtbExactEndDate.Name = "dtbExactEndDate"
        Me.dtbExactEndDate.ReadOnly = True
        Me.dtbExactEndDate.TabStop = False
        '
        'tbActivity
        '
        resources.ApplyResources(Me.tbActivity, "tbActivity")
        Me.tbActivity.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbActivity.ForeColor = System.Drawing.Color.Blue
        Me.tbActivity.Name = "tbActivity"
        Me.tbActivity.ReadOnly = True
        Me.tbActivity.TabStop = False
        '
        'dtbExactStartDate
        '
        resources.ApplyResources(Me.dtbExactStartDate, "dtbExactStartDate")
        Me.dtbExactStartDate.DateValue = New Date(2013, 5, 30, 0, 0, 0, 0)
        Me.dtbExactStartDate.ForeColor = System.Drawing.Color.Black
        Me.dtbExactStartDate.Name = "dtbExactStartDate"
        Me.dtbExactStartDate.ReadOnly = True
        Me.dtbExactStartDate.TabStop = False
        '
        'DetailActivity
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.lblExactEndDate)
        Me.Controls.Add(Me.TabControlActivity)
        Me.Controls.Add(Me.dtbExactEndDate)
        Me.Controls.Add(Me.tbActivity)
        Me.Controls.Add(Me.lblExactStartDate)
        Me.Controls.Add(Me.lblActivity)
        Me.Controls.Add(Me.dtbExactStartDate)
        Me.Name = "DetailActivity"
        Me.TabControlActivity.ResumeLayout(False)
        Me.TabPageDuration.ResumeLayout(False)
        Me.TabPageDuration.PerformLayout()
        Me.TabPageOrganisation.ResumeLayout(False)
        Me.TabPageOrganisation.PerformLayout()
        Me.TabPagePreparation.ResumeLayout(False)
        Me.TabPagePreparation.PerformLayout()
        Me.TabPageRepeat.ResumeLayout(False)
        Me.TabPageRepeat.PerformLayout()
        Me.TabPageChildActivities.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents tbActivity As FaciliDev.LogFramer.TextBoxLF
    Friend WithEvents TabControlActivity As System.Windows.Forms.TabControl
    Friend WithEvents TabPageRepeat As System.Windows.Forms.TabPage
    Friend WithEvents TabPageDuration As System.Windows.Forms.TabPage
    Friend WithEvents lblExactEndDate As System.Windows.Forms.Label
    Friend WithEvents dtbExactEndDate As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents lblExactStartDate As System.Windows.Forms.Label
    Friend WithEvents dtbExactStartDate As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents TabPagePreparation As System.Windows.Forms.TabPage
    Friend WithEvents chkFollowUpRepeat As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents chkPreparationRepeat As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents chkFollowUpUntilEnd As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents chkPreparationFromStart As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents ntbFollowUp As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents ntbPreparation As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents cmbFollowUpUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblFollowUp As System.Windows.Forms.Label
    Friend WithEvents cmbPreparationUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblPreparation As System.Windows.Forms.Label
    Friend WithEvents lblRepeatDates As System.Windows.Forms.Label
    Friend WithEvents lvRepeatDates As System.Windows.Forms.ListView
    Friend WithEvents lblTimes As System.Windows.Forms.Label
    Friend WithEvents ntbRepeatTimes As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents lblNumberRepeats As System.Windows.Forms.Label
    Friend WithEvents chkRepeatUntilEnd As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents ntbRepeatEvery As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents cmbRepeatUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblRepeat As System.Windows.Forms.Label
    Friend WithEvents TabPageChildActivities As System.Windows.Forms.TabPage
    Friend WithEvents lvDetailProcesses As FaciliDev.LogFramer.ListViewProcesses
    Friend WithEvents lblEndDateFollowUp As System.Windows.Forms.Label
    Friend WithEvents dtbEndDateFollowUp As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents lblStartDatePreparation As System.Windows.Forms.Label
    Friend WithEvents dtbStartDatePreparation As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents TabPageOrganisation As System.Windows.Forms.TabPage
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cmbType As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbLocation As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblOrganisation As System.Windows.Forms.Label
    Friend WithEvents cmbOrganisation As FaciliDev.LogFramer.ComboBoxText
    Friend WithEvents lblEndDateMainActivity As System.Windows.Forms.Label
    Friend WithEvents dtbEndDateMainActivity As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents lblStartDateMainActivity As System.Windows.Forms.Label
    Friend WithEvents dtbStartDateMainActivity As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents chkDurationUntilEnd As FaciliDev.LogFramer.CheckBoxLF
    Friend WithEvents ntbDuration As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents cmbDurationUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents dtbStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbGuidReferenceMoment As FaciliDev.LogFramer.ComboBoxSelectValue
    Friend WithEvents cmbPeriodDirection As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents cmbPeriodUnit As FaciliDev.LogFramer.ComboBoxSelectIndex
    Friend WithEvents ntbPeriod As FaciliDev.LogFramer.NumericTextBoxLF
    Friend WithEvents rbtnReferenceDate As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnExactDate As System.Windows.Forms.RadioButton
    Friend WithEvents lblStart As System.Windows.Forms.Label
    Friend WithEvents lblStartDateFollowUp As System.Windows.Forms.Label
    Friend WithEvents dtbStartDateFollowUp As FaciliDev.LogFramer.DateTextBox
    Friend WithEvents lblEndDatePreparation As System.Windows.Forms.Label
    Friend WithEvents dtbEndDatePreparation As FaciliDev.LogFramer.DateTextBox

End Class
