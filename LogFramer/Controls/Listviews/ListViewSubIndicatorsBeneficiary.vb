Public Class ListViewSubIndicatorsBeneficiary
    Inherits ListViewSubIndicators

    Private objTargetGroup As TargetGroup

    Public Shadows Event Updated()
    Public Shadows Event IndicatorModified()

#Region "Properties"
    Public Property TargetGroup As TargetGroup
        Get
            Return objTargetGroup
        End Get
        Set(ByVal value As TargetGroup)
            objTargetGroup = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub New()
        View = View.Details
        FullRowSelect = True
        OwnerDraw = True
    End Sub

    Public Sub New(ByVal indicator As Indicator, ByVal targetdeadlinessection As TargetDeadlinesSection, ByVal targetgroup As TargetGroup)
        View = View.Details
        FullRowSelect = True
        OwnerDraw = True

        Me.ParentIndicator = indicator
        Me.TargetDeadlinesSection = targetdeadlinessection
        Me.TargetGroup = targetgroup

        LoadColumns()
    End Sub

    Public Overloads Sub LoadColumns()
        Dim strHeaderText As String = String.Empty
        Columns.Clear()

        If ParentIndicator IsNot Nothing Then
            Columns.Add(LANG_Indicator, 300, HorizontalAlignment.Left)
            Columns.Add(LANG_ResponseType, 150, HorizontalAlignment.Left)
            Columns.Add(LANG_TargetGroup, 120, HorizontalAlignment.Left)
            Columns.Add(LANG_RegistrationOption, 120, HorizontalAlignment.Left)

            strHeaderText = LANG_Baseline
            Columns.Add(String.Format("{0}: {1}", strHeaderText, LANG_ScoreValueBeneficiary), 150, HorizontalAlignment.Right)
            Columns.Add(String.Format("{0}: {1}", strHeaderText, LANG_PopulationTargetText), 150, HorizontalAlignment.Right)
            Columns.Add(String.Format("{0}: {1}", strHeaderText, LANG_ScoreValueTargetGroup), 150, HorizontalAlignment.Right)

            If TargetDeadlinesSection.TargetDeadlines IsNot Nothing Then
                For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
                    'set column header text
                    strHeaderText = String.Empty
                    strHeaderText = String.Format("{0} - {1}", LANG_Target, SetColumnHeaderText(strHeaderText, selTargetDeadline))

                    'add columns
                    Columns.Add(String.Format("{0}: {1}", strHeaderText, LANG_ScoreValueBeneficiary), 150, HorizontalAlignment.Right)
                    Columns.Add(String.Format("{0}: {1}", strHeaderText, LANG_PopulationTargetText), 150, HorizontalAlignment.Right)
                    Columns.Add(String.Format("{0}: {1}", strHeaderText, LANG_ScoreValueTargetGroup), 150, HorizontalAlignment.Right)
                Next
            End If
        End If
    End Sub

    Protected Overrides Sub LoadItems_SubIndicators()
        Dim objPopulationTargets As PopulationTargets = Nothing
        Dim selPopulationTarget As PopulationTarget

        For Each selIndicator As Indicator In Me.ParentIndicator.Indicators
            Dim newItem As New ListViewItem(selIndicator.Text)
            Dim selTargetGroup As TargetGroup = CurrentLogFrame.Purposes.GetTargetGroupByGuid(selIndicator.TargetGroupGuid)
            Dim strTargetGroupName As String = String.Empty
            Dim dblMaxScoreBeneficiary As Double = selIndicator.ResponseClasses.GetMaximumScore(selIndicator.AddClassValuesForTotal)

            'values: set number of decimals and unit
            If selTargetGroup IsNot Nothing Then strTargetGroupName = selTargetGroup.Name

            LoadItems_SetUnits(selIndicator)

            With newItem
                'General information
                .Name = selIndicator.Guid.ToString
                .SubItems.Add(selIndicator.QuestionTypeName)
                .SubItems.Add(strTargetGroupName)
                .SubItems.Add(selIndicator.RegistrationOption)

                'Baseline
                .SubItems.Add(selIndicator.GetBaselineFormattedScore)
                Select Case selIndicator.QuestionType
                    Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                        .SubItems.Add(String.Empty)
                    Case Else
                        .SubItems.Add(DisplayAsUnit(selIndicator.Baseline.PopulationPercentage, intNrDecimals, "%"))
                End Select
                .SubItems.Add(selIndicator.GetPopulationBaselineFormattedScore)

                'Targets
                objPopulationTargets = selIndicator.GetPopulationTotalPercentage()

                For i = 0 To ParentIndicator.PopulationTargets.Count - 1
                    selPopulationTarget = selIndicator.PopulationTargets(i)

                    .SubItems.Add(selIndicator.GetTargetFormattedScore(i))
                    .SubItems.Add(DisplayAsUnit(selPopulationTarget.Percentage, intNrDecimals, "%"))
                    .SubItems.Add(selIndicator.GetPopulationTargetFormattedScore(i))
                Next
            End With
            Me.Items.Add(newItem)
        Next
    End Sub

    Protected Overrides Sub LoadItems_Totals()
        Dim strTotal As String = String.Format("{0} ({1})", LANG_Total, LIST_AggregationOptions(ParentIndicator.AggregateVertical).ToLower)
        Dim TotalItem As New ListViewItem(strTotal)

        LoadItems_SetUnits(ParentIndicator)

        With TotalItem
            .Name = "Total"
            .SubItems.Add(String.Empty)
            .SubItems.Add(String.Empty)
            .SubItems.Add(String.Empty)

            'Baseline
            .SubItems.Add(String.Empty)
            .SubItems.Add(String.Empty)
            .SubItems.Add(ParentIndicator.GetPopulationBaselineFormattedScore)

            'Targets
            For i = 0 To ParentIndicator.Targets.Count - 1
                .SubItems.Add(String.Empty)
                .SubItems.Add(String.Empty)
                .SubItems.Add(ParentIndicator.GetPopulationTargetFormattedScore(i))
            Next
        End With

        Me.Items.Add(TotalItem)
    End Sub
#End Region

#Region "Draw"
    Protected Overrides Sub OnDrawColumnHeader(ByVal e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs)
        MyBase.OnDrawColumnHeader(e)
        If e.ColumnIndex > 3 Then
            Dim graphics As Graphics = e.Graphics
            Dim rCell As Rectangle = e.Bounds
            rCell.X += 4
            rCell.Y += 5
            rCell.Width -= 4
            rCell.Height -= 10
            e.DrawBackground()
            If ColorBackGround(e.ColumnIndex) = True Then
                graphics.FillRectangle(brBlue, e.Bounds)
            End If

            Dim sfSubItem As New StringFormat
            sfSubItem.Alignment = StringAlignment.Center
            sfSubItem.LineAlignment = StringAlignment.Near
            graphics.DrawString(e.Header.Text, Me.Font, Brushes.Black, rCell, sfSubItem)
        Else
            e.DrawDefault = True
        End If
    End Sub

    Private Function ColorBackGround(ByVal intColumnIndex As Integer) As Boolean
        intColumnIndex -= 1
        Dim intDivider As Integer = 3
        Dim intNormalised As Integer

        'If ParentIndicator.QuestionType = Indicator.QuestionTypes.PercentageValue Then
        '    intColumnIndex -= 1
        '    intDivider = 2
        'End If
        intNormalised = Math.Floor(intColumnIndex / intDivider)

        If Decimal.Remainder(intNormalised, 2) = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    Protected Overrides Sub OnDrawSubItem(ByVal e As System.Windows.Forms.DrawListViewSubItemEventArgs)
        Dim graphics As Graphics = e.Graphics
        Dim sfSubItem As New StringFormat

        sfSubItem.LineAlignment = StringAlignment.Center
        If e.ColumnIndex = 0 Then
            sfSubItem.Alignment = StringAlignment.Near
        Else
            sfSubItem.Alignment = StringAlignment.Far
        End If

        Dim rCell As Rectangle = e.Bounds
        rCell.X += 4
        rCell.Y += 2
        rCell.Width -= 8
        rCell.Height -= 4

        If e.Item.Name = "Total" Then
            graphics.FillRectangle(SystemBrushes.ControlDark, e.Bounds)
            graphics.DrawString(e.SubItem.Text, Me.Font, Brushes.Black, rCell, sfSubItem)
        ElseIf e.ColumnIndex > 3 Then
            e.DrawBackground()
            If ColorBackGround(e.ColumnIndex) = True Then
                graphics.FillRectangle(brBlue, e.Bounds)
            End If
            graphics.DrawString(e.SubItem.Text, Me.Font, Brushes.Black, rCell, sfSubItem)
        Else
            e.DrawDefault = True
        End If
    End Sub
#End Region
End Class

