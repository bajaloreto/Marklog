Public Class ListViewSubIndicators
    Inherits ListViewSubIndicatorsBase

    Public Shadows Event Updated()
    Public Shadows Event IndicatorModified()

    Protected objTargetDeadlinesSection As TargetDeadlinesSection
    Protected strValueName As String = String.Empty
    Protected strUnit As String = String.Empty
    Protected intNrDecimals As Integer


#Region "Properties"
    Public Property TargetDeadlinesSection As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesSection
        End Get
        Set(value As TargetDeadlinesSection)
            objTargetDeadlinesSection = value
        End Set
    End Property
#End Region

#Region "Initialise"
    Public Sub New()
        View = View.Details
        FullRowSelect = True
        OwnerDraw = True
    End Sub

    Public Sub New(ByVal indicator As Indicator, ByVal targetdeadlinessection As TargetDeadlinesSection)
        View = View.Details
        FullRowSelect = True
        OwnerDraw = True

        Me.ParentIndicator = indicator
        Me.TargetDeadlinesSection = targetdeadlinessection

        LoadColumns()
    End Sub

    Public Overloads Sub LoadColumns()

        Columns.Clear()

        If ParentIndicator IsNot Nothing Then
            Columns.Add(LANG_Indicator, 300, HorizontalAlignment.Left)
            Columns.Add(LANG_ResponseType, 150, HorizontalAlignment.Left)
            Columns.Add(LANG_TargetGroup, 120, HorizontalAlignment.Left)
            Columns.Add(LANG_RegistrationOption, 120, HorizontalAlignment.Left)

            Columns.Add(LANG_Baseline, 150, HorizontalAlignment.Right)

            If TargetDeadlinesSection.TargetDeadlines IsNot Nothing Then
                For Each selTargetDeadline As TargetDeadline In TargetDeadlinesSection.TargetDeadlines
                    Dim strDeadline As String = String.Empty
                    'set column header text
                    strDeadline = SetColumnHeaderText(strDeadline, selTargetDeadline)

                    'add columns
                    Columns.Add(strDeadline, 150, HorizontalAlignment.Right)
                Next
            End If
        End If
    End Sub

    Protected Function SetColumnHeaderText(ByVal strDeadline As String, ByVal selTargetDeadline As TargetDeadline) As String
        Select Case TargetDeadlinesSection.Repetition
            Case TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                strDeadline &= selTargetDeadline.ExactDeadline.ToString("yyyy")
            Case TargetDeadlinesSection.RepetitionOptions.TwiceYear, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.MonthlyTarget
                strDeadline &= selTargetDeadline.ExactDeadline.ToString("MMM yyyy")
            Case TargetDeadlinesSection.RepetitionOptions.UserSelect
                strDeadline &= selTargetDeadline.ExactDeadline.ToShortDateString
        End Select

        Return strDeadline
    End Function

    Public Overrides Sub LoadItems()
        Dim objTargetScoresTotal As Targets = Nothing


        Me.Items.Clear()
        If ParentIndicator IsNot Nothing Then
            'sub indicators
            LoadItems_SubIndicators()

            'totals
            If ParentIndicator.Indicators.Count > 0 Then
                LoadItems_Totals()

                For i = 0 To 2
                    Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                Next
                For i = 4 To Columns.Count - 1

                    Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                    Columns(i).Width *= 1.4
                    If Columns(i).Width < 100 Then Columns(i).Width = 100
                    Select Case ParentIndicator.QuestionType
                        Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                            If ParentIndicator.Registration = Indicator.RegistrationOptions.BeneficiaryLevel Then _
                                Columns(5).Width = 0
                    End Select
                Next
            Else
                For i = 1 To 3
                    Columns(i).Width = 0
                Next
            End If

            RaiseEvent Updated()
        End If
    End Sub

    Protected Overrides Sub LoadItems_SubIndicators()

        For Each selIndicator As Indicator In Me.ParentIndicator.Indicators
            Dim newItem As New ListViewItem(selIndicator.Text)
            Dim selTargetGroup As TargetGroup = CurrentLogFrame.Purposes.GetTargetGroupByGuid(selIndicator.TargetGroupGuid)
            Dim strTargetGroupName As String = String.Empty

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

                'Targets
                For i = 0 To selIndicator.Targets.Count - 1
                    .SubItems.Add(selIndicator.GetTargetFormattedScore(i))
                Next
            End With
            Me.Items.Add(newItem)
        Next
    End Sub

    Protected Overridable Sub LoadItems_Totals()
        Dim strTotal As String = String.Format("{0} ({1})", LANG_Total, LIST_AggregationOptions(ParentIndicator.AggregateVertical).ToLower)
        Dim TotalItem As New ListViewItem(strTotal)

        With TotalItem
            'General information
            .Name = "Total"
            .SubItems.Add(String.Empty)
            .SubItems.Add(String.Empty)
            .SubItems.Add(String.Empty)

            'Baseline
            .SubItems.Add(ParentIndicator.GetBaselineFormattedScore)

            'Targets
            For i = 0 To ParentIndicator.Targets.Count - 1
                .SubItems.Add(ParentIndicator.GetTargetFormattedScore(i))
            Next
        End With

        Me.Items.Add(TotalItem)
    End Sub

    Protected Sub LoadItems_SetUnits(ByVal selIndicator As Indicator)
        Select Case selIndicator.QuestionType
            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                If selIndicator.ValuesDetail IsNot Nothing Then
                    With selIndicator.ValuesDetail
                        strValueName = .ValueName

                        intNrDecimals = .NrDecimals
                        If selIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                            strUnit = .Unit
                        Else
                            strUnit = String.Empty
                        End If
                    End With
                End If
        End Select
    End Sub
#End Region

#Region "Methods & Events"
    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        PopUpIndicatorDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If Me.ParentIndicator.Indicators.Count > 0 AndAlso Me.SelectedIndicators.Length > 0 AndAlso Me.SelectedIndicators(0) IsNot Nothing Then
            PopUpIndicatorDetails(Me.SelectedIndicators(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If Me.ParentIndicator.Indicators.Count > 0 AndAlso Me.SelectedIndicators.Length > 0 Then
            If MsgBox(LANG_RemoveIndicator, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selIndicator As Indicator = Me.SelectedIndicators(0)
                'CurrentUndoList.DeleteOperation(selIndicator, Me.ParentIndicator.Indicators, , , True)
                Me.ParentIndicator.Indicators.Remove(selIndicator)
                Me.LoadItems()

                RaiseEvent IndicatorModified()
            End If
        End If
    End Sub

    Private Sub PopUpIndicatorDetails(ByVal selIndicator As Indicator)
        Dim boolNew As Boolean

        If selIndicator Is Nothing Then
            boolNew = True
            selIndicator = New Indicator()

            Me.ParentIndicator.Indicators.Add(selIndicator) 'AddToProcess

            'CurrentUndoList.InsertNewOperation(selIndicator, Me.ParentIndicator.Indicators, , , True)
        End If

        Dim dialogIndicator As New DialogIndicator(selIndicator)
        Dim OldParentIndicatorGuid As Guid = selIndicator.ParentIndicatorGuid

        If dialogIndicator.ShowDialog() = DialogResult.OK Then
            Me.LoadItems()
            RaiseEvent IndicatorModified()
        End If
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selIndicator As Indicator In SelectedIndicators
            'CurrentUndoList.CutOperation(selIndicator, Me.ParentIndicator.Indicators, Me.ParentIndicator.Indicators.IndexOf(selIndicator), , True)

            Me.ParentIndicator.Indicators.Remove(selIndicator)
        Next

        LoadItems()
        RaiseEvent IndicatorModified()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selIndicator As Indicator In SelectedIndicators
            Dim NewItem As New ClipboardItem(CopyGroup, selIndicator, Indicator.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selIndicator As Indicator

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Indicator)
                    selIndicator = CType(selItem.Item, Indicator)
                    Dim NewIndicator As New Indicator

                    Using copier As New ObjectCopy
                        NewIndicator = copier.CopyObject(selIndicator)
                    End Using

                    Me.ParentIndicator.Indicators.Add(NewIndicator)

                    'CurrentUndoList.PasteOperation(NewIndicator, Me.ParentIndicator.Indicators, 0, , True)
            End Select
        Next

        Me.LoadItems()
        RaiseEvent IndicatorModified()
    End Sub
#End Region

#Region "Draw"
    Protected Overrides Sub OnDrawColumnHeader(e As System.Windows.Forms.DrawListViewColumnHeaderEventArgs)
        MyBase.OnDrawColumnHeader(e)
        
        e.DrawDefault = True
    End Sub

    Protected Overrides Sub OnDrawSubItem(e As System.Windows.Forms.DrawListViewSubItemEventArgs)
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
        Else
            e.DrawDefault = True
        End If
    End Sub
#End Region
End Class

