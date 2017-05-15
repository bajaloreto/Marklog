Public Class IndicatorTargetsImages
    Private objCurrentIndicator As Indicator
    Private objTargetDeadlinesSection As TargetDeadlinesSection
    Private objTargetGroup As TargetGroup

#Region "Properties"
    Public Property CurrentIndicator As Indicator
        Get
            Return objCurrentIndicator
        End Get
        Set(ByVal value As Indicator)
            objCurrentIndicator = value
            If objCurrentIndicator IsNot Nothing Then objCurrentIndicator.CalculateTargetsWithFormula()
        End Set
    End Property

    Public Property TargetDeadlinesSection As TargetDeadlinesSection
        Get
            Return objTargetDeadlinesSection
        End Get
        Set(ByVal value As TargetDeadlinesSection)
            objTargetDeadlinesSection = value
        End Set
    End Property

    Public Property TargetGroup As TargetGroup
        Get
            Return objTargetGroup
        End Get
        Set(ByVal value As TargetGroup)
            objTargetGroup = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New()
        InitializeComponent()
    End Sub

    Public Overridable Sub LoadImages()
        Dim intTargetIndex As Integer
        Dim selTarget As Target

        If CurrentIndicator IsNot Nothing Then
            With CurrentIndicator
                Dim BaselineImage As New ucImage(CurrentIndicator.Baseline.AudioVisualDetail)
                BaselineImage.Dock = DockStyle.Fill

                TabPageBaseline.Controls.Add(BaselineImage)

                For intTargetIndex = 0 To .Targets.Count - 1
                    selTarget = .Targets(intTargetIndex)

                    Dim NewTabPage As New TabPage()
                    Dim TargetImage As New ucImage(selTarget.AudioVisualDetail)
                    TargetImage.Dock = DockStyle.Fill

                    NewTabPage.Controls.Add(TargetImage)
                    TabControlTargets.TabPages.Add(NewTabPage)
                Next
            End With
            SetTabHeaders()
        End If
    End Sub

    Public Sub SetTabHeaders()
        Dim strBaselineScore As String = 0.ToString
        Dim strHeaderText As String
        Dim strDate As String
        Dim strScore As String = String.Empty
        Dim selTargetDeadline As TargetDeadline

        TabPageBaseline.Text = LANG_Baseline

        For intTargetIndex = 0 To CurrentIndicator.Targets.Count - 1
            selTargetDeadline = TargetDeadlinesSection.TargetDeadlines(intTargetIndex)

            'deadline
            Select Case TargetDeadlinesSection.Repetition
                Case TargetDeadlinesSection.RepetitionOptions.MonthlyTarget, TargetDeadlinesSection.RepetitionOptions.QuarterlyTarget, TargetDeadlinesSection.RepetitionOptions.TwiceYear
                    strDate = selTargetDeadline.ExactDeadline.ToString("MMM-yyyy")
                Case TargetDeadlinesSection.RepetitionOptions.SingleTarget, TargetDeadlinesSection.RepetitionOptions.YearlyTarget
                    strDate = selTargetDeadline.ExactDeadline.ToString("yyyy")
                Case Else
                    strDate = selTargetDeadline.ExactDeadline.ToShortDateString
            End Select

            'set header text
            strHeaderText = String.Format("{0} {1}", LANG_Target, strDate)

            TabControlTargets.TabPages(intTargetIndex + 1).Text = strHeaderText
        Next
    End Sub
#End Region
End Class
