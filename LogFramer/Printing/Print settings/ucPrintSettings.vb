Imports System.Drawing.Printing

Public Class PrintSettings
    Private objReportSetup As New ReportSetup
    Private boolFireEvent As Boolean = True

    Public Event ReportSetupChanged(ByVal sender As Object, ByVal e As ReportSetupChangedEventArgs)
    Public Event PrintButtonClicked()
    Public Event SelectPrinterButtonClicked()

    Public Property ReportSetup() As ReportSetup
        Get
            Return objReportSetup
        End Get
        Set(ByVal value As ReportSetup)
            objReportSetup = value

            CheckButtons()
        End Set
    End Property

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
    End Sub

    Private Sub rbtnLandscape_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnLandscape.CheckedChanged
        If boolFireEvent = True Then
            Me.ReportSetup.PrintAsLandScape = rbtnLandscape.Checked
            RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
        End If
    End Sub

    Private Sub rbtnThin_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnThin.CheckedChanged
        If boolFireEvent = True Then
            If rbtnThin.Checked = True Then
                Me.ReportSetup.PrintMargins = New Margins(ToInch(10), ToInch(10), ToInch(10), ToInch(10))
                RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
            End If
        End If
    End Sub

    Private Sub rbtnNormal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnNormal.CheckedChanged
        If boolFireEvent = True Then
            If rbtnNormal.Checked = True Then
                Me.ReportSetup.PrintMargins = New Margins(ToInch(18), ToInch(18), ToInch(20), ToInch(20))
                RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
            End If
        End If
    End Sub

    Private Sub rbtnWide_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnWide.CheckedChanged
        If boolFireEvent = True Then
            If rbtnWide.Checked = True Then
                Me.ReportSetup.PrintMargins = New Margins(ToInch(25), ToInch(25), ToInch(25), ToInch(25))
                RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
            End If
        End If
    End Sub

    Private Sub btnPageMargins_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPageMargins.Click
        Dim dlgPageSetup As New PageSetupDialog
        Dim objPageSettings As New PageSettings
        With objPageSettings
            .Color = ReportSetup.PrintInColor
            .Landscape = ReportSetup.PrintAsLandScape
            .Margins = ReportSetup.PrintMargins
        End With

        With dlgPageSetup
            .PageSettings = objPageSettings
            .EnableMetric = True
            If .ShowDialog() = DialogResult.OK Then
                With .PageSettings
                    ReportSetup.PrintInColor = .Color
                    ReportSetup.PrintAsLandScape = .Landscape
                    ReportSetup.PrintMargins = .Margins
                End With
                CheckButtons()
                RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
            End If
        End With
    End Sub

    Private Sub CheckButtons()
        boolFireEvent = False

        With Me.ReportSetup
            Me.rbtnLandscape.Checked = ReportSetup.PrintAsLandScape
            If ReportSetup.PrintAsLandScape = True Then Me.rbtnPortrait.Checked = False Else Me.rbtnPortrait.Checked = True
        End With

        With Me.ReportSetup.PrintMargins
            If .Top = .Right And .Top = .Bottom Then
                If .Top >= ToInch(8) And .Top <= ToInch(12) Then
                    rbtnThin.Checked = True
                ElseIf .Top >= ToInch(23) And .Top <= ToInch(27) Then
                    rbtnWide.Checked = True
                End If
            Else
                If .Top >= ToInch(18) And .Top <= ToInch(22) And _
                    .Right >= ToInch(16) And .Right <= ToInch(20) And _
                    .Bottom >= ToInch(18) And .Bottom <= ToInch(22) And _
                    .Left >= ToInch(16) And .Left <= ToInch(20) Then
                    rbtnNormal.Checked = True
                End If
            End If
        End With
        boolFireEvent = True

    End Sub

    Private Function ToInch(ByVal mm As Single) As Single
        Dim sngHundrethsOfInch As Single = mm / 25.4 * 100

        Return sngHundrethsOfInch
    End Function

    Private Sub btnHeaderText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHeaderText.Click
        Using dialogHeader As New DialogPrintHeader
            With dialogHeader
                .Text = LANG_PrintHeader
                .chkCopyToAllReports.Text = LANG_CopyHeader

                .rtbLeft.Rtf = ReportSetup.HeaderLeft
                .rtbMiddle.Rtf = ReportSetup.HeaderMiddle
                .rtbRight.Rtf = ReportSetup.HeaderRight

                If dialogHeader.ShowDialog = DialogResult.OK Then
                    If .chkCopyToAllReports.Checked = True Then
                        ChangeAllReportHeaders(.rtbLeft.Rtf, .rtbMiddle.Rtf, .rtbRight.Rtf)
                    Else
                        ReportSetup.HeaderLeft = .rtbLeft.Rtf
                        ReportSetup.HeaderMiddle = .rtbMiddle.Rtf
                        ReportSetup.HeaderRight = .rtbRight.Rtf
                    End If

                    RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
                End If
            End With
        End Using
    End Sub

    Private Sub ChangeAllReportHeaders(ByVal LeftRtf As String, ByVal MiddleRtf As String, ByVal RightRtf As String)
        With CurrentLogFrame.ReportSetupIndicators
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupLogframe
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupPartnerList
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupPlanning
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupPmf
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupResources
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupRiskRegister
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupAssumptions
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupDependencies
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupTargetGroupIdForm
            .HeaderLeft = LeftRtf
            .HeaderMiddle = MiddleRtf
            .HeaderRight = RightRtf
        End With
    End Sub

    Private Sub btnFooterText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFooterText.Click
        Using dialogFooter As New DialogPrintHeader
            With dialogFooter
                .Text = LANG_PrintFooter
                .chkCopyToAllReports.Text = LANG_CopyFooter

                .rtbLeft.Rtf = ReportSetup.FooterLeft
                .rtbMiddle.Rtf = ReportSetup.FooterMiddle
                .rtbRight.Rtf = ReportSetup.FooterRight

                If .ShowDialog = DialogResult.OK Then
                    If .chkCopyToAllReports.Checked = True Then
                        ChangeAllReportFooters(.rtbLeft.Rtf, .rtbMiddle.Rtf, .rtbRight.Rtf)
                    Else
                        ReportSetup.FooterLeft = .rtbLeft.Rtf
                        ReportSetup.FooterMiddle = .rtbMiddle.Rtf
                        ReportSetup.FooterRight = .rtbRight.Rtf
                    End If

                    RaiseEvent ReportSetupChanged(Me, New ReportSetupChangedEventArgs(Me.ReportSetup))
                End If
            End With
        End Using
    End Sub

    Private Sub ChangeAllReportFooters(ByVal LeftRtf As String, ByVal MiddleRtf As String, ByVal RightRtf As String)
        With CurrentLogFrame.ReportSetupIndicators
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupLogframe
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupPartnerList
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupPlanning
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupPmf
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupResources
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupRiskRegister
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupAssumptions
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupDependencies
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
        With CurrentLogFrame.ReportSetupTargetGroupIdForm
            .FooterLeft = LeftRtf
            .FooterMiddle = MiddleRtf
            .FooterRight = RightRtf
        End With
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        RaiseEvent PrintButtonClicked()
    End Sub

    Private Sub btnPrinter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectPrinter.Click
        RaiseEvent SelectPrinterButtonClicked()
    End Sub

    Private Sub ProgressBarDocument_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProgressBarDocument.VisibleChanged

    End Sub
End Class

Public Class ReportSetupChangedEventArgs
    Inherits EventArgs

    Public Property ReportSetup As ReportSetup

    Public Sub New(ByVal objReportSetup As ReportSetup)
        MyBase.New()

        Me.ReportSetup = objReportSetup
    End Sub
End Class
