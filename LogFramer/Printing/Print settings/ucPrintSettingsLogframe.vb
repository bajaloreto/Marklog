Imports System.Drawing.Printing

Public Class PrintSettingsLogframe
    Private boolFireEvent As Boolean = True
    Private boolShowStructColumn, boolShowIndColumn, boolShowVerColumn, boolShowAsmcolumn As Boolean
    Private boolShowGoals, boolShowPurposes, boolShowOutputs, boolShowActivities As Boolean
    Private boolShowResourcesBudget As Boolean

    Public Event ColumnSelectionChanged()
    Public Event SectionSelectionChanged()
    Public Event ResourceBudgetChanged()

    Public Property ShowIndicatorColumn As Boolean
        Get
            Return boolShowIndColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowIndColumn = value
            boolFireEvent = False
            chkShowInd.Checked = value
            boolFireEvent = True
        End Set
    End Property

    Public Property ShowVerificationSourceColumn As Boolean
        Get
            Return boolShowVerColumn
        End Get
        Set(ByVal value As Boolean)
            boolShowVerColumn = value
            boolFireEvent = False
            chkShowVer.Checked = value
            boolFireEvent = True
        End Set
    End Property

    Public Property ShowAssumptionColumn As Boolean
        Get
            Return boolShowAsmcolumn
        End Get
        Set(ByVal value As Boolean)
            boolShowAsmcolumn = value
            boolFireEvent = False
            chkShowAsm.Checked = value
            boolFireEvent = True
        End Set
    End Property

    Public Property ShowGoals() As Boolean
        Get
            Return boolShowGoals
        End Get
        Set(ByVal value As Boolean)
            boolShowGoals = value
            boolFireEvent = False
            chkShowGoals.Checked = value

            EnsureSectionIsVisible()

            boolFireEvent = True
        End Set
    End Property

    Public Property ShowPurposes() As Boolean
        Get
            Return boolShowPurposes
        End Get
        Set(ByVal value As Boolean)
            boolShowPurposes = value
            boolFireEvent = False
            chkShowPurposes.Checked = value

            EnsureSectionIsVisible()

            boolFireEvent = True
        End Set
    End Property

    Public Property ShowOutputs() As Boolean
        Get
            Return boolShowOutputs
        End Get
        Set(ByVal value As Boolean)
            boolShowOutputs = value
            boolFireEvent = False
            chkShowOutputs.Checked = value

            EnsureSectionIsVisible()

            boolFireEvent = True
        End Set
    End Property

    Public Property ShowActivities() As Boolean
        Get
            Return boolShowActivities
        End Get
        Set(ByVal value As Boolean)
            boolShowActivities = value
            boolFireEvent = False
            chkShowActivities.Checked = value

            EnsureSectionIsVisible()

            boolFireEvent = True
        End Set
    End Property

    Public Property ShowResourcesBudget As Boolean
        Get
            Return boolShowResourcesBudget
        End Get
        Set(ByVal value As Boolean)
            boolShowResourcesBudget = value
            boolFireEvent = False

            If value = True Then
                rbtnPrintResourcesBudget.Checked = True
            Else
                rbtnPrintIndicators.Checked = True
            End If
            boolFireEvent = True
        End Set
    End Property

    Private Sub chkShowInd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowInd.CheckedChanged
        If boolFireEvent = True Then
            ShowIndicatorColumn = chkShowInd.Checked
            If ShowIndicatorColumn = False Then
                ShowVerificationSourceColumn = False
                chkShowVer.Checked = False
            End If
            Me.Invalidate()
            RaiseEvent ColumnSelectionChanged()
        End If
    End Sub

    Private Sub chkShowVer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowVer.CheckedChanged
        If boolFireEvent = True Then
            ShowVerificationSourceColumn = chkShowVer.Checked
            If ShowVerificationSourceColumn = True Then
                ShowIndicatorColumn = True
                chkShowInd.Checked = True
            End If
            Me.Invalidate()
            RaiseEvent ColumnSelectionChanged()
        End If
    End Sub

    Private Sub chkShowAsm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAsm.CheckedChanged
        If boolFireEvent = True Then
            ShowAssumptionColumn = chkShowAsm.Checked
            Me.Invalidate()
            RaiseEvent ColumnSelectionChanged()
        End If
    End Sub

    Private Sub rbtnPrintResourcesBudget_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPrintResourcesBudget.CheckedChanged
        If boolFireEvent = True Then
            ShowResourcesBudget = rbtnPrintResourcesBudget.Checked
            Me.Invalidate()
            RaiseEvent ResourceBudgetChanged()
        End If
    End Sub

    Private Sub chkShowGoals_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkShowGoals.CheckedChanged
        If boolFireEvent = True Then
            ShowGoals = chkShowGoals.Checked
            Me.Invalidate()
            RaiseEvent SectionSelectionChanged()
        End If
    End Sub

    Private Sub chkShowPurposes_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkShowPurposes.CheckedChanged
        If boolFireEvent = True Then
            ShowPurposes = chkShowPurposes.Checked
            Me.Invalidate()
            RaiseEvent SectionSelectionChanged()
        End If
    End Sub

    Private Sub chkShowOutputs_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkShowOutputs.CheckedChanged
        If boolFireEvent = True Then
            ShowOutputs = chkShowOutputs.Checked
            Me.Invalidate()
            RaiseEvent SectionSelectionChanged()
        End If
    End Sub

    Private Sub chkShowActivities_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkShowActivities.CheckedChanged
        If boolFireEvent = True Then
            ShowActivities = chkShowActivities.Checked
            Me.Invalidate()
            RaiseEvent SectionSelectionChanged()
        End If
    End Sub

    Private Sub rbtnPrintIndicators_CheckedChanged(sender As Object, e As System.EventArgs) Handles rbtnPrintIndicators.CheckedChanged
        Dim boolChecked As Boolean

        If boolFireEvent = True Then
            If rbtnPrintIndicators.Checked = True Then boolChecked = False Else boolChecked = True

            ShowResourcesBudget = boolChecked
            Me.Invalidate()
            RaiseEvent ResourceBudgetChanged()
        End If
    End Sub

    Private Sub EnsureSectionIsVisible()
        If ShowGoals = False And ShowPurposes = False And ShowOutputs = False And ShowActivities = False Then
            ShowPurposes = True

            RaiseEvent SectionSelectionChanged()
        End If
    End Sub
End Class
