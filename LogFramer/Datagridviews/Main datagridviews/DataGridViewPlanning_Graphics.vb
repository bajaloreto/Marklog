Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization

Partial Public Class DataGridViewPlanning

#Region "Custom painting - general"
    Private Function GetCoordinateX(ByVal intCellX As Integer, ByVal datStart As Date) As Integer
        Dim intX As Integer
        intX = intCellX + (datStart.Subtract(datPeriodFrom).Days * sngScale)
        Return intX
    End Function

    Private Function GetCoordinateY(ByVal intCellY As Integer, ByVal intCellHeight As Integer, Optional ByVal boolPreparation As Boolean = False) As Integer
        Dim intY As Integer

        If boolPreparation = False Then
            intY = intCellY + ((intCellHeight - CONST_BarHeight) / 2)
        Else
            intY = intCellY + ((intCellHeight - CONST_PreparationHeight) / 2)
        End If

        Return intY
    End Function

    Private Function GetActivityWidth(ByVal datStart As Date, ByVal datEnd As Date) As Integer
        Dim intWidth As Integer
        intWidth = ((datEnd.Subtract(datStart).Days) + 1) * sngScale
        Return intWidth
    End Function

    Private Function GetFollowUpWidth(ByVal datEndMainActivity As Date, ByVal datEndFollowUp As Date) As Integer
        Dim intWidth As Integer
        intWidth = ((datEndFollowUp.Subtract(datEndMainActivity).Days)) * sngScale
        Return intWidth
    End Function

    Private Function CoordinateToDate(ByVal intCellX As Integer, ByVal intX As Integer) As Date
        Dim selDate As Date

        selDate = datPeriodFrom.AddDays((intX - intCellX) / sngScale)
        Return selDate
    End Function

    Protected Overrides Sub OnCellPainting(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If Me.Grid.Count > 0 Then
            Dim CellGraphics As Graphics = e.Graphics
            Dim rCell As Rectangle = e.CellBounds
            Dim intIndex As Integer = e.RowIndex
            Dim selGridRow As PlanningGridRow = Grid(intIndex)
            Dim boolAlignRight As Boolean

            If e.ColumnIndex = 2 Or e.ColumnIndex = 3 Then boolAlignRight = True

            If e.ColumnIndex < CONST_PlanningColumnIndex And intIndex >= 0 Then
                'text and dates columns
                If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    CellGraphics.FillRectangle(Brushes.Lavender, rCell)
                    CellPainting_NormalText(e.FormattedValue, True, boolAlignRight, rCell, CellGraphics)
                    e.Paint(rCell, DataGridViewPaintParts.Border)
                    e.Handled = True
                ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
                    If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectActivities Then
                        CellGraphics.FillRectangle(New SolidBrush(SystemColors.Highlight), rCell)
                    Else
                        e.PaintBackground(rCell, False)
                    End If

                    If e.ColumnIndex = 1 Then
                        CellPainting_RichText(selGridRow, rCell, CellGraphics)
                    Else
                        CellPainting_NormalText(e.FormattedValue, False, boolAlignRight, rCell, CellGraphics)
                    End If
                    e.Paint(rCell, DataGridViewPaintParts.Border)
                    e.Handled = True
                End If
            ElseIf e.ColumnIndex = CONST_PlanningColumnIndex Then
                'column with Gantt chart
                If intIndex = -1 Then
                    'header
                    CellPainting_PlanningHeader(rCell.X, rCell.Height + 1, CellGraphics)
                ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity Or selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    'rows
                    PaintPlanning_Background(e)

                    If selGridRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                        CellPainting_PlanningKeyMoment(selGridRow, rCell, CellGraphics)
                        If selGridRow.LinksLoaded = False Then CellPainting_PlanningKeyMomentLinks(selGridRow, rCell, CellGraphics)
                    ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.Activity Then
                        CellPainting_PlanningActivity(selGridRow, rCell, CellGraphics)
                        If selGridRow.LinksLoaded = False Then CellPainting_PlanningActivityLinks(selGridRow, rCell, CellGraphics)
                    End If

                    CellPainting_PlanningVerticalKeyMomentLinks(selGridRow, rCell, CellGraphics)
                    CellPainting_PlanningVerticalActivityLinks(selGridRow, rCell, CellGraphics)
                    CellGraphics.DrawLine(penBorder, rCell.Left, rCell.Bottom - 1, rCell.Right, rCell.Bottom - 1)

                    If rDragRectangle <> Rectangle.Empty Then
                        PaintPlanning_Drag(selGridRow, rCell, CellGraphics)
                    End If
                End If

                e.Handled = True
            End If
            'MyBase.OnCellPainting(e)
        End If
    End Sub

    Private Sub CellPainting_RichText(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim rImage As Rectangle

        If selGridRow.StructImage IsNot Nothing Then
            rImage = New Rectangle(rCell.X, rCell.Y, selGridRow.StructImage.Width, selGridRow.StructImage.Height)
            CellGraphics.DrawImage(selGridRow.StructImage, rImage)
        End If
    End Sub

    Private Sub CellPainting_NormalText(ByVal strText As String, ByVal boolKeyMoment As Boolean, ByVal boolAlignRight As Boolean, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim fntText As Font = DefaultCellStyle.Font
        Dim brText As New SolidBrush(Color.Black)
        Dim objFormat As New StringFormat

        If Me.TextSelectionIndex = TextSelectionValues.SelectAll Or Me.TextSelectionIndex = TextSelectionValues.SelectActivities Then
            brText = New SolidBrush(SystemColors.HighlightText)
        End If

        rCell.X += CONST_Padding
        rCell.Y += CONST_Padding
        rCell.Width -= CONST_HorizontalPadding
        rCell.Height -= CONST_VerticalPadding

        If boolKeyMoment = True Then
            brText = New SolidBrush(Color.DarkBlue)
            fntText = New Font(Me.DefaultCellStyle.Font, FontStyle.Bold)
        End If
        If boolAlignRight = True Then
            objFormat.Alignment = StringAlignment.Far
        Else
            objFormat.Alignment = StringAlignment.Near
        End If

        CellGraphics.DrawString(strText, fntText, brText, rCell, objFormat)
    End Sub
#End Region

#Region "Custom painting - Gantt chart"
    Private Sub CellPainting_PlanningKeyMoment(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
        If selKeyMoment Is Nothing Then Exit Sub

        Dim strStart As String = String.Empty
        Dim intKeyMomentX As Integer = GetCoordinateX(rCell.X, selKeyMoment.ExactDateKeyMoment)
        Dim intKeyMomentY As Integer = rCell.Y + ((rCell.Height - CONST_BarHeight) / 2)
        Dim intKeyMomentHeight As Integer = CONST_BarHeight
        Dim fntText As Font = Me.Font
        Dim brText As Brush = Brushes.Black

        If Me.PeriodView = PeriodViews.Week Or Me.PeriodView = PeriodViews.Day Then
            strStart = selKeyMoment.ExactDateKeyMoment.ToString("d-MMM")
        End If

        PaintPlanning_PaintKeyMoment(CellGraphics, rCell, intKeyMomentX, intKeyMomentY, intKeyMomentHeight, fntText, brText, strStart)
    End Sub

    Private Sub CellPainting_PlanningActivity(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
        If selActivity Is Nothing Then Exit Sub

        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim intType As Integer
        Dim intActivityX As Integer, intActivityY As Integer
        Dim intActivityWidth As Integer, intActivityHeight As Integer
        Dim intPreparationX As Integer, intPreparationY As Integer
        Dim intPreparationWidth As Integer, intPreparationHeight As Integer
        Dim intFollowUpX As Integer, intFollowUpY As Integer
        Dim intFollowUpWidth As Integer, intFollowUpHeight As Integer
        Dim strText As String = String.Empty
        Dim fntText As Font = Me.Font
        Dim brText As Brush = Brushes.Black
        Dim selActivityDetail As ActivityDetail = selActivity.ActivityDetail

        If selActivity.IsProcess = False Then
            intType = selActivity.ActivityDetail.Type

            'activity
            intActivityX = GetCoordinateX(rCell.X, selActivityDetail.StartDateMainActivity)
            intActivityY = GetCoordinateY(rCell.Y, rCell.Height)
            intActivityWidth = GetActivityWidth(selGridRow.StartDate, selGridRow.EndDate)
            intActivityHeight = CONST_BarHeight

            'preparation
            If selActivityDetail.StartDatePreparation <> selActivityDetail.StartDateMainActivity Then
                intPreparationX = GetCoordinateX(rCell.X, selActivityDetail.StartDatePreparation)
                intPreparationY = GetCoordinateY(rCell.Y, rCell.Height, True)
                intPreparationWidth = intActivityX - intPreparationX + 1
                intPreparationHeight = CONST_PreparationHeight

                If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                    strStart = selActivityDetail.StartDatePreparation.ToString("d-MMM")
                    strEnd = String.Empty
                End If
                PaintPlanning_PaintActivity(CellGraphics, rCell, intType, intPreparationX, intPreparationY, intPreparationWidth, intPreparationHeight, fntText, brText, strStart, strEnd)
            End If

            'follow-up
            If selActivityDetail.EndDateFollowUp <> selActivityDetail.EndDateMainActivity Then
                intFollowUpX = intActivityX + intActivityWidth
                intFollowUpY = GetCoordinateY(rCell.Y, rCell.Height, True)
                intFollowUpWidth = GetFollowUpWidth(selActivityDetail.EndDateMainActivity, selActivityDetail.EndDateFollowUp)
                If intFollowUpX + intFollowUpWidth > rCell.Right Then intFollowUpWidth = rCell.Right - intFollowUpX
                intFollowUpHeight = CONST_FollowUpHeight

                If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                    strStart = String.Empty
                    strEnd = selActivityDetail.EndDateFollowUp.ToString("d-MMM")
                End If
                PaintPlanning_PaintActivity(CellGraphics, rCell, intType, intFollowUpX, intFollowUpY, intFollowUpWidth, intFollowUpHeight, fntText, brText, strStart, strEnd)
            End If

            'activity
            If Me.PeriodView = PeriodViews.Week Then
                strStart = selActivityDetail.StartDateMainActivity.ToString("d-MMM")
                strEnd = selActivityDetail.EndDateMainActivity.ToString("d-MMM")
            ElseIf Me.PeriodView = PeriodViews.Day Then
                strStart = selActivityDetail.StartDateMainActivity.ToString("d-MMM")
                If String.IsNullOrEmpty(selActivityDetail.Location) = False Then strStart &= " - " & selActivityDetail.Location
                If String.IsNullOrEmpty(selActivityDetail.Organisation) = False Then strStart &= " - " & selActivityDetail.Organisation
                strEnd = selActivityDetail.EndDateMainActivity.ToString("d-MMM")
            End If
            PaintPlanning_PaintActivity(CellGraphics, rCell, intType, intActivityX, intActivityY, intActivityWidth, intActivityHeight, fntText, brText, strStart, strEnd)

            'repeats
            If selActivityDetail.RepeatStartDates.Count > 0 Then
                Dim selStartDate As Date
                Dim selEndDate As Date
                For i = 0 To selActivityDetail.RepeatStartDates.Count - 1
                    selStartDate = selActivityDetail.RepeatStartDates(i)
                    selEndDate = selActivityDetail.RepeatEndDates(i)
                    intActivityX = GetCoordinateX(rCell.X, selStartDate)
                    intActivityY = GetCoordinateY(rCell.Y, rCell.Height)
                    intActivityHeight = CONST_BarHeight

                    If selActivityDetail.PreparationRepeat = True Then
                        intPreparationX = intActivityX - intPreparationWidth
                        intPreparationY = GetCoordinateY(rCell.Y, rCell.Height, True)

                        If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                            strStart = selStartDate.AddDays((selActivityDetail.StartDateMainActivity.Subtract(selActivityDetail.StartDatePreparation).Days) * -1).ToString("d-MMM")
                            strEnd = String.Empty
                        End If
                        PaintPlanning_PaintActivity(CellGraphics, rCell, intType, intPreparationX, intPreparationY, intPreparationWidth, intPreparationHeight, fntText, brText, strStart, strEnd)
                    End If

                    If selActivityDetail.FollowUpRepeat = True Then
                        intFollowUpX = intActivityX + intActivityWidth
                        intFollowUpY = GetCoordinateY(rCell.Y, rCell.Height, True)

                        If Me.PeriodView = DataGridViewPlanning.PeriodViews.Week Or Me.PeriodView = DataGridViewPlanning.PeriodViews.Day Then
                            strStart = String.Empty
                            strEnd = selEndDate.AddDays((selActivityDetail.EndDateFollowUp.Subtract(selActivityDetail.EndDateMainActivity).Days)).ToString("d-MMM")
                        End If
                        PaintPlanning_PaintActivity(CellGraphics, rCell, intType, intFollowUpX, intFollowUpY, intFollowUpWidth, intFollowUpHeight, fntText, brText, strStart, strEnd)
                    End If

                    If Me.PeriodView = PeriodViews.Week Then
                        strStart = selStartDate.ToString("d-MMM")
                        strEnd = selEndDate.ToString("d-MMM")
                    ElseIf Me.PeriodView = PeriodViews.Day Then
                        strStart = selStartDate.ToString("d-MMM")
                        If String.IsNullOrEmpty(selActivityDetail.Location) = False Then strStart &= " - " & selActivityDetail.Location
                        If String.IsNullOrEmpty(selActivityDetail.Organisation) = False Then strStart &= " - " & selActivityDetail.Organisation
                        strEnd = selEndDate.ToString("d-MMM")
                    End If
                    PaintPlanning_PaintActivity(CellGraphics, rCell, intType, intActivityX, intActivityY, intActivityWidth, intActivityHeight, fntText, brText, strStart, strEnd)
                Next
            End If
        Else
            CellPainting_PlanningProcess(selGridRow, rCell, CellGraphics)
        End If
    End Sub

    Private Sub CellPainting_PlanningProcess(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
        Dim strStart As String = String.Empty, strEnd As String = String.Empty
        Dim intActivityX As Integer, intActivityY As Integer
        Dim intActivityWidth As Integer, intActivityHeight As Integer

        With selActivity
            intActivityX = GetCoordinateX(rCell.X, .ExactStartDate)
            intActivityY = GetCoordinateY(rCell.Y, rCell.Height)
            intActivityWidth = GetActivityWidth(selActivity.ExactStartDate, .ExactEndDate)
            intActivityHeight = CONST_PreparationHeight

            If Me.PeriodView = PeriodViews.Week Then
                strStart = .ExactStartDate.ToString("d-MMM")
                strEnd = .ExactEndDate.ToString("d-MMM")
            ElseIf Me.PeriodView = PeriodViews.Day Then
                strStart = .ExactStartDate.ToString("d-MMM")
                With .ActivityDetail
                    If String.IsNullOrEmpty(.Location) = False Then strStart = String.Format("{0} - {1}", strStart, .Location)
                    If String.IsNullOrEmpty(.Organisation) = False Then strStart = String.Format("{0} - {1}", strStart, .Organisation)
                End With
                strEnd = .ExactEndDate.ToString("d-MMM")
            End If
        End With
        PaintPlanning_PaintProcess(CellGraphics, rCell, intActivityX, intActivityY, intActivityWidth, intActivityHeight, strStart, strEnd)
    End Sub

    Private Sub CellPainting_PlanningHeader(ByVal intCellX As Integer, ByVal intRowHeight As Integer, ByVal CellGraphics As Graphics)
        If datPeriodFrom > Date.MinValue And datPeriodEnd > Date.MinValue Then
            'background
            Dim intX As Integer
            Dim objDfi As DateTimeFormatInfo = DateTimeFormatInfo.CurrentInfo
            Dim objCal As Calendar = objDfi.Calendar
            Dim spanPeriod As TimeSpan = datPeriodEnd - datPeriodFrom
            Dim selDate As Date = datPeriodFrom
            Dim intDaysInMonth As Integer
            Dim rPeriod As New Rectangle(intCellX, 0, spanPeriod.Days * sngScale, intRowHeight)
            Dim strMonth As String
            Dim fntYear As New Font(Me.Font, FontStyle.Bold)
            Dim fntMonth As New Font(Me.Font.FontFamily, Me.Font.SizeInPoints - 2)
            Dim fntDetail As New Font(Me.Font.FontFamily, Me.Font.SizeInPoints - 2)
            Dim intMonthWidth, intMonthX As Integer
            Dim intNrOfDays As Integer
            Dim brWeekEnd As New SolidBrush(Color.DarkSlateGray)

            CellGraphics.FillRectangle(SystemBrushes.ControlLight, rPeriod)
            For i = 0 To spanPeriod.TotalDays
                selDate = datPeriodFrom.AddDays(i)
                intDaysInMonth = Date.DaysInMonth(selDate.Year, selDate.Month)
                intX = GetCoordinateX(intCellX, selDate)
                'intX = intCellX + (selDate.Subtract(datPeriodStart).Days * sngScale)

                If selDate.Month = 1 And selDate.Day = 1 Then
                    CellGraphics.DrawLine(penBlack1, intX, 0, intX, intRowHeight)
                    CellGraphics.DrawString(selDate.ToString("yyyy"), fntYear, Brushes.Black, intX, 0)
                End If
                If selDate.Day = 1 Then
                    Select Case PeriodView
                        Case PeriodViews.Month
                            strMonth = selDate.ToString("MMM")
                            CellGraphics.DrawLine(penBlack1, intX, CInt(intRowHeight / 2), intX, intRowHeight)
                            CellGraphics.DrawString(strMonth, fntMonth, Brushes.Black, intX, CInt(intRowHeight / 2))
                        Case PeriodViews.Week, PeriodViews.Day
                            intNrOfDays = intDaysInMonth * sngScale
                            If selDate.Month / 2 = Int(selDate.Month / 2) Then
                                CellGraphics.FillRectangle(SystemBrushes.ControlDark, New Rectangle(intX, 0, intNrOfDays, intRowHeight))
                            End If
                            strMonth = selDate.ToString("MMMM \'yy")

                            fntMonth = New Font(fntMonth, FontStyle.Bold)
                            intMonthWidth = CellGraphics.MeasureString(strMonth, fntMonth).Width
                            intMonthX = intX + Int((intNrOfDays / 2) - (intMonthWidth / 2))

                            CellGraphics.DrawString(strMonth, fntMonth, Brushes.Black, intMonthX, 2)
                    End Select
                End If
                If selDate.DayOfWeek = DayOfWeek.Monday And Me.PeriodView = PeriodViews.Week Then
                    Dim intNrWeek As Integer = objCal.GetWeekOfYear(selDate, objDfi.CalendarWeekRule, objDfi.FirstDayOfWeek)
                    CellGraphics.DrawLine(penBlack1, intX, CInt(intRowHeight / 2), intX, intRowHeight)
                    CellGraphics.DrawString(intNrWeek.ToString(), fntDetail, Brushes.Black, intX, CInt(intRowHeight / 2))
                End If
                If Me.PeriodView = PeriodViews.Day Then
                    If selDate.DayOfWeek = DayOfWeek.Saturday Or selDate.DayOfWeek = DayOfWeek.Sunday Then
                        CellGraphics.FillRectangle(brWeekEnd, intX, CInt(intRowHeight / 2), sngScale, CInt(intRowHeight / 2))
                    End If
                    CellGraphics.DrawLine(penBlack1, intX, CInt(intRowHeight / 2), intX, intRowHeight)
                    CellGraphics.DrawString(selDate.Day.ToString, fntDetail, Brushes.Black, intX, CInt(intRowHeight / 2))
                End If
            Next

            Dim penDark As New Pen(SystemBrushes.ControlDark, 1)
            CellGraphics.DrawLine(penDark, intCellX, intRowHeight, spanPeriod.Days * sngScale, intRowHeight)
        End If
    End Sub

    Private Sub PaintPlanning_Background(ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        Dim objDfi As DateTimeFormatInfo = DateTimeFormatInfo.CurrentInfo
        Dim objCal As Calendar = objDfi.Calendar
        Dim brBackGround As Brush = Brushes.White
        Dim brDuration As Brush = Brushes.Gainsboro
        Dim brWeekEnd As New SolidBrush(Me.DarkerColor(Color.Gainsboro))
        Dim intX, intY As Integer, intWidth As Integer, intHeight As Integer = e.CellBounds.Height
        Dim spanPeriod As TimeSpan = datPeriodEnd - datPeriodFrom
        Dim selDate As Date = datPeriodFrom
        Dim rBackGround As New Rectangle(e.CellBounds.X, e.CellBounds.Y, spanPeriod.Days * sngScale, intHeight)
        Dim penGrid As New Pen(SystemColors.ControlDark, 1)

        'If e.State And DataGridViewElementStates.Selected Then
        '    brBackGround = SystemBrushes.Highlight
        '    brDuration = SystemBrushes.ControlDark
        'End If

        e.Graphics.FillRectangle(brBackGround, rBackGround)

        intX = GetCoordinateX(e.CellBounds.X, datStartDate)
        intY = e.CellBounds.Y
        intWidth = GetActivityWidth(datStartDate, datEndDate)
        Dim rDuration As New Rectangle(intX, intY, intWidth, intHeight)
        e.Graphics.FillRectangle(brDuration, rDuration)

        For i = 0 To spanPeriod.TotalDays
            selDate = datPeriodFrom.AddDays(i)
            intX = GetCoordinateX(e.CellBounds.X, selDate)

            If selDate.Day = 1 Then
                If selDate.Month = 1 Then
                    e.Graphics.DrawLine(New Pen(SystemColors.ControlDark, 1), intX, intY, intX, intY + intHeight)
                Else
                    Select Case PeriodView
                        Case PeriodViews.Week
                            Dim penMonth As New Pen(SystemColors.ControlDarkDark, 2)
                            penMonth.DashStyle = DashStyle.Dash
                            e.Graphics.DrawLine(penMonth, intX, intY, intX, intY + intHeight)
                        Case Else
                            e.Graphics.DrawLine(penGrid, intX, intY, intX, intY + intHeight)
                    End Select
                End If
            End If
            If selDate.DayOfWeek = DayOfWeek.Monday And Me.PeriodView = PeriodViews.Week Then
                Dim intNrWeek As Integer = objCal.GetWeekOfYear(selDate, objDfi.CalendarWeekRule, objDfi.FirstDayOfWeek)
                e.Graphics.DrawLine(penGrid, intX, intY, intX, intY + intHeight)
            End If
            If Me.PeriodView = PeriodViews.Day Then
                If selDate.DayOfWeek = DayOfWeek.Saturday Or selDate.DayOfWeek = DayOfWeek.Sunday Then
                    e.Graphics.FillRectangle(brWeekEnd, intX, intY, sngScale, intHeight)
                End If
                e.Graphics.DrawLine(penGrid, intX, intY, intX, intY + intHeight)
            End If
        Next
    End Sub

    Private Sub PaintPlanning_Drag(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        If DragActivity IsNot Nothing Then
            Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)
            If selActivity IsNot Nothing And selActivity Is DragActivity Then
                Dim ActivityColor As Color = Me.GetActivityColor(selActivity.ActivityDetail.Type)
                Dim myHatchBrush As New HatchBrush(HatchStyle.SmallConfetti, Color.DarkGray, ActivityColor)
                Dim strDate As String = datDragMoment.ToShortDateString
                Dim intDateWidth As Integer = CellGraphics.MeasureString(strDate, DefaultFont).Width

                CellGraphics.FillRectangle(myHatchBrush, rDragRectangle)
                CellGraphics.DrawRectangle(New Pen(Brushes.DarkGray, 2), rDragRectangle)

                If boolDragActivity = True Then
                    Dim datDragActivityStart, datDragActivityEnd As Date
                    Dim intDaysMoved As Integer = datDragMoment.Subtract(datInitialDragMoment).Days

                    With selActivity.ActivityDetail
                        datDragActivityStart = .StartDateMainActivity.AddDays(intDaysMoved)
                        datDragActivityEnd = .EndDateMainActivity.AddDays(intDaysMoved)
                    End With

                    strDate = datDragActivityStart.ToShortDateString
                    intDateWidth = CellGraphics.MeasureString(strDate, DefaultFont).Width
                    CellGraphics.DrawString(strDate, DefaultFont, Brushes.Black, rDragRectangle.X - intDateWidth, rCell.Y)

                    strDate = datDragActivityEnd.ToShortDateString
                    intDateWidth = CellGraphics.MeasureString(strDate, DefaultFont).Width
                    CellGraphics.DrawString(strDate, DefaultFont, Brushes.Black, rDragRectangle.Right, rCell.Y)
                ElseIf boolDragActivityStart = True Or boolDragPreparationStart = True Then
                    CellGraphics.DrawString(strDate, DefaultFont, Brushes.Black, rDragRectangle.X - intDateWidth, rCell.Y)
                ElseIf boolDragActivityEnd = True Or boolDragFollowUpEnd = True Then
                    CellGraphics.DrawString(strDate, DefaultFont, Brushes.Black, rDragRectangle.Right, rCell.Y)
                End If

            End If
        ElseIf DragKeyMoment IsNot Nothing Then
            Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
            If selKeyMoment IsNot Nothing And selKeyMoment Is DragKeyMoment Then
                Dim KeyMomentColor As Color = Color.DarkRed
                Dim myHatchBrush As New HatchBrush(HatchStyle.SmallConfetti, Color.DarkGray, KeyMomentColor)
                Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line, PathPointType.Line}
                Dim Points(3) As Point
                Dim strDate As String = datDragMoment.ToShortDateString
                Dim intDateWidth As Integer = CellGraphics.MeasureString(strDate, DefaultFont).Width
                Dim intAdd As Integer = CONST_BarHeight / 2

                Points(0) = New Point(rDragRectangle.Left + intAdd, rDragRectangle.Top)
                Points(1) = New Point(rDragRectangle.Right, rDragRectangle.Top + intAdd)
                Points(2) = New Point(rDragRectangle.Left + intAdd, rDragRectangle.Bottom)
                Points(3) = New Point(rDragRectangle.Left, rDragRectangle.Top + intAdd)

                Dim objPath As New GraphicsPath(Points, PathPoints)
                CellGraphics.FillPath(myHatchBrush, objPath)
                CellGraphics.DrawPath(New Pen(Brushes.DarkGray, 2), objPath)

                CellGraphics.DrawString(strDate, DefaultFont, Brushes.Black, rDragRectangle.X - intDateWidth, rCell.Y)
            End If
        End If
    End Sub

    Private Function GetActivityColor(ByVal intType As Integer) As Color
        Dim ActivityColor As Color = Color.DarkSeaGreen

        Select Case intType
            Case ActivityDetail.Types.Other
                ActivityColor = Color.DarkSeaGreen
            Case ActivityDetail.Types.Planning, ActivityDetail.Types.ProjectProposal, ActivityDetail.Types.DonorSigned
                ActivityColor = Color.Plum
            Case ActivityDetail.Types.Evaluation, ActivityDetail.Types.Audit, ActivityDetail.Types.Monitoring, ActivityDetail.Types.Report
                ActivityColor = Color.Tomato
            Case ActivityDetail.Types.Identification, ActivityDetail.Types.Selection, ActivityDetail.Types.Registration
                ActivityColor = Color.Orange
            Case ActivityDetail.Types.HiringStaff, ActivityDetail.Types.Procurement, ActivityDetail.Types.Logistics, ActivityDetail.Types.Debriefing
                ActivityColor = Color.CadetBlue
            Case ActivityDetail.Types.Meeting
                ActivityColor = Color.Blue
            Case ActivityDetail.Types.Travel
                ActivityColor = Color.MediumBlue
            Case ActivityDetail.Types.Activity, ActivityDetail.Types.Construction, ActivityDetail.Types.Distribution, _
                ActivityDetail.Types.MedicalTreatment, ActivityDetail.Types.HumanitarianAssistance, ActivityDetail.Types.PeaceBuilding
                ActivityColor = Color.LightGreen
            Case ActivityDetail.Types.Research, ActivityDetail.Types.Training, ActivityDetail.Types.CapacityDevelopment, ActivityDetail.Types.Awareness
                ActivityColor = Color.Yellow
            Case Else
                ActivityColor = Color.PaleGreen
        End Select

        Return ActivityColor
    End Function

    Private Sub PaintPlanning_PaintKeyMoment(ByVal CellGraphics As Graphics, ByVal rCell As Rectangle, _
                                             ByVal intKeyMomentX As Integer, ByVal intKeyMomentY As Integer, ByVal intKeyMomentHeight As Integer, _
                                             ByVal fntText As Font, ByVal brText As Brush, ByVal strStartDate As String)

        Dim brKeyMoment As LinearGradientBrush = MakeGradientBrush(Color.DarkRed, intKeyMomentX, intKeyMomentY, intKeyMomentHeight)
        Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line, PathPointType.Line}
        Dim Points(3) As Point
        Dim intAdd As Integer = CONST_BarHeight / 2

        Points(0) = New Point(intKeyMomentX, intKeyMomentY)
        Points(1) = New Point(intKeyMomentX + intAdd, intKeyMomentY + intAdd)
        Points(2) = New Point(intKeyMomentX, intKeyMomentY + CONST_BarHeight)
        Points(3) = New Point(intKeyMomentX - intAdd, intKeyMomentY + intAdd)
        CellGraphics.FillPath(brKeyMoment, New GraphicsPath(Points, PathPoints))

        Dim intTextX As Integer = intKeyMomentX + 2
        Dim intTextY As Integer = rCell.Y
        Dim intStartDateWidth As Integer

        If String.IsNullOrEmpty(strStartDate) = False Then
            intStartDateWidth = CellGraphics.MeasureString(strStartDate, fntText).Width
            CellGraphics.DrawString(strStartDate, fntText, brText, intTextX, intTextY)
        End If

    End Sub

    Private Sub PaintPlanning_PaintProcess(ByVal CellGraphics As Graphics, ByVal rCell As Rectangle, _
                                           ByVal intActivityX As Integer, ByVal intActivityY As Integer, ByVal intActivityWidth As Integer, ByVal intActivityHeight As Integer, _
                                           ByVal strStart As String, ByVal strEnd As String)

        Dim rProcess As New Rectangle(intActivityX - 4, intActivityY, intActivityWidth + 8, intActivityHeight)
        Dim intStartDateWidth As Integer, intEndDateWidth As Integer
        Dim intTextX As Integer, intTextY As Integer
        Dim fntText As Font = Me.Font
        Dim brText As Brush = Brushes.Black

        CellGraphics.FillRectangle(Brushes.Black, rProcess)

        Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line}
        Dim Points(2) As Point
        Points(0) = New Point(rProcess.Left, rProcess.Bottom)
        Points(1) = New Point(rProcess.Left + 4, rProcess.Bottom + 8)
        Points(2) = New Point(rProcess.Left + 8, rProcess.Bottom)
        CellGraphics.FillPath(Brushes.Black, New GraphicsPath(Points, PathPoints))

        Points(0) = New Point(rProcess.Right, rProcess.Bottom)
        Points(1) = New Point(rProcess.Right - 4, rProcess.Bottom + 8)
        Points(2) = New Point(rProcess.Right - 8, rProcess.Bottom)
        CellGraphics.FillPath(Brushes.Black, New GraphicsPath(Points, PathPoints))

        If String.IsNullOrEmpty(strStart) = False Then
            intStartDateWidth = CellGraphics.MeasureString(strStart, fntText).Width
            intTextX = intActivityX
            intTextY = rCell.Y
            CellGraphics.DrawString(strStart, fntText, brText, intTextX, intTextY)
        End If
        If String.IsNullOrEmpty(strEnd) = False Then
            intEndDateWidth = CellGraphics.MeasureString(strEnd, fntText).Width
            Dim intTextHeight As Integer = CellGraphics.MeasureString(strEnd, fntText).Height
            intTextX = intActivityX + intActivityWidth
            If intActivityWidth > intStartDateWidth + intEndDateWidth Then intTextX -= intEndDateWidth
            intTextY = rCell.Bottom - intTextHeight

            CellGraphics.DrawString(strEnd, fntText, brText, intTextX, intTextY)
        End If
    End Sub

    Private Sub PaintPlanning_PaintActivity(ByVal graphics As Graphics, ByVal rCell As Rectangle, ByVal intType As Integer, _
                                            ByVal intActivityX As Integer, ByVal intActivityY As Integer, ByVal intActivityWidth As Integer, ByVal intActivityHeight As Integer, _
                                            ByVal fntText As Font, ByVal brText As Brush, ByVal strStartDate As String, ByVal strEndDate As String)
        Dim rActivity As New Rectangle(intActivityX, intActivityY, intActivityWidth, intActivityHeight)
        Dim intStartDateWidth As Integer, intEndDateWidth As Integer
        Dim intTextX As Integer, intTextY As Integer
        Dim ActivityColor As Color = Me.GetActivityColor(intType)
        Dim brActivity As LinearGradientBrush = MakeGradientBrush(ActivityColor, intActivityX, intActivityY, intActivityHeight)

        graphics.FillRectangle(brActivity, rActivity)

        If String.IsNullOrEmpty(strStartDate) = False Then
            intStartDateWidth = graphics.MeasureString(strStartDate, fntText).Width
            intTextX = intActivityX
            intTextY = rCell.Y
            graphics.DrawString(strStartDate, fntText, brText, intTextX, intTextY)
        End If
        If String.IsNullOrEmpty(strEndDate) = False Then
            intEndDateWidth = graphics.MeasureString(strEndDate, fntText).Width
            Dim intTextHeight As Integer = graphics.MeasureString(strEndDate, fntText).Height
            intTextX = intActivityX + intActivityWidth
            If intActivityWidth > intStartDateWidth + intEndDateWidth Then intTextX -= intEndDateWidth
            intTextY = rCell.Bottom - intTextHeight

            graphics.DrawString(strEndDate, fntText, brText, intTextX, intTextY)
        End If
    End Sub

    Private Function LighterBrush(ByVal OldBrush As SolidBrush) As SolidBrush
        Dim intRed As Integer = OldBrush.Color.R + 50
        Dim intGreen As Integer = OldBrush.Color.G + 50
        Dim intBlue As Integer = OldBrush.Color.B + 50
        If intRed > 255 Then intRed = 255
        If intGreen > 255 Then intGreen = 255
        If intBlue > 255 Then intBlue = 255

        Dim NewBrush As New SolidBrush(Color.FromArgb(255, intRed, intGreen, intBlue))
        Return NewBrush
    End Function

    Public Function DarkerColor(ByVal OldColor As Color) As Color
        Dim intRed As Integer = OldColor.R - 50
        Dim intGreen As Integer = OldColor.G - 50
        Dim intBlue As Integer = OldColor.B - 50
        If intRed < 0 Then intRed = 0
        If intGreen < 0 Then intGreen = 0
        If intBlue < 0 Then intBlue = 0

        Dim NewColor As Color = Color.FromArgb(255, intRed, intGreen, intBlue)
        Return NewColor
    End Function

    Public Function MakeGradientBrush(ByVal BaseColor As Color, ByVal intX As Integer, ByVal intY As Integer, ByVal intHeight As Integer) As LinearGradientBrush
        Dim colLight As Color = BaseColor
        Dim colDark As Color = DarkerColor(colLight)
        If intX < 0 Then intX = 0

        Dim NewBrush As New LinearGradientBrush(New Point(intX, intY), New Point(intX, intY + intHeight), colLight, colDark)
        Return NewBrush
    End Function
#End Region

#Region "Custom painting - Links between activities/key moments"
    Private Sub CellPainting_PlanningKeyMomentLinks(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        If selGridRow.KeyMoment Is Nothing Then Exit Sub

        If selGridRow.OutgoingLinksIndices.Count > 0 Then
            CellPainting_PlanningKeyMomentLinks_Outgoing(selGridRow, rCell, CellGraphics)
        End If

        If selGridRow.IncomingLinkIndices.Count > 0 Then
            CellPainting_PlanningKeyMomentLinks_Incoming(selGridRow, rCell, CellGraphics)
        End If
    End Sub

    Private Sub CellPainting_PlanningKeyMomentLinks_Outgoing(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
        Dim intIndex As Integer = Grid.IndexOf(selGridRow)
        Dim intLinkX As Integer = GetCoordinateX(rCell.X, selKeyMoment.ExactDateKeyMoment) + (CONST_BarHeight / 2)
        Dim intLinkY As Integer = rCell.Y + (rCell.Height / 2)
        Dim intTargetX As Integer
        Dim LinkStart As Point = New Point(intLinkX, intLinkY)

        For i = 0 To selGridRow.OutgoingLinksIndices.Count - 1
            Dim intTargetIndex As Integer = selGridRow.OutgoingLinksIndices(i)
            Dim TargetRow As PlanningGridRow = Me.Grid(intTargetIndex)

            If TargetRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                Dim selTargetKeyMoment As KeyMoment = TargetRow.KeyMoment

                intTargetX = GetCoordinateX(rCell.X, selTargetKeyMoment.ExactDateKeyMoment) 'use the current cell to calculate since that is certainly visible


            ElseIf TargetRow.RowType = PlanningGridRow.RowTypes.Activity Then
                Dim selTargetActivity As Activity = TryCast(TargetRow.Struct, Activity)

                intTargetX = GetCoordinateX(rCell.X, selTargetActivity.ExactStartDate)
            Else
                Exit Sub
            End If

            If Me.ShowKeyMomentLinks = True Then CellPainting_PlanningKeyMomentLinks_Outgoing_Draw(intIndex, intTargetIndex, intTargetX, LinkStart, rCell, CellGraphics)
        Next
    End Sub

    Private Sub CellPainting_PlanningKeyMomentLinks_Outgoing_Draw(ByVal intIndex As Integer, ByVal intTargetIndex As Integer, ByVal intTargetX As Integer, ByVal LinkStart As Point, _
                                                                  ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim FirstElbow, SecondElbow, ThirdElbow, EndPoint As New Point

        If intTargetX >= LinkStart.X + CONST_DistanceToKeyMoment Then
            'if the target key moment/activity is not directly below the key moment, draw a simple line with one right angle
            FirstElbow = New Point(intTargetX, LinkStart.Y)

            If intTargetIndex > intIndex Then
                EndPoint = New Point(FirstElbow.X, rCell.Bottom)
            Else
                EndPoint = New Point(FirstElbow.X, rCell.Top)
            End If
            CellGraphics.DrawLine(penLightBlue2, LinkStart, FirstElbow)
            CellGraphics.DrawLine(penLightBlue2, FirstElbow, EndPoint)
            CellGraphics.DrawLine(penDarkBlue1, LinkStart, FirstElbow)
            CellGraphics.DrawLine(penDarkBlue1, FirstElbow, EndPoint)
        Else
            'if the target key moment/activity starts directly below the key moment, draw a half loop
            FirstElbow = New Point(LinkStart.X + CONST_DistanceToKeyMoment, LinkStart.Y)
            If intTargetIndex > intIndex Then
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y + 10)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, rCell.Bottom)
            Else
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y - 10)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, rCell.Top)
            End If

            CellGraphics.DrawLine(penLightBlue2, LinkStart, FirstElbow)
            CellGraphics.DrawLine(penLightBlue2, FirstElbow, SecondElbow)
            CellGraphics.DrawLine(penLightBlue2, SecondElbow, ThirdElbow)
            CellGraphics.DrawLine(penLightBlue2, ThirdElbow, EndPoint)

            CellGraphics.DrawLine(penDarkBlue1, LinkStart, FirstElbow)
            CellGraphics.DrawLine(penDarkBlue1, FirstElbow, SecondElbow)
            CellGraphics.DrawLine(penDarkBlue1, SecondElbow, ThirdElbow)
            CellGraphics.DrawLine(penDarkBlue1, ThirdElbow, EndPoint)
        End If

        CellPainting_PlanningKeyMomentLinks_VerticalLinks(intIndex, intTargetIndex, EndPoint, rCell)
    End Sub

    Private Sub CellPainting_PlanningKeyMomentLinks_VerticalLinks(ByVal intIndex As Integer, ByVal intTargetIndex As Integer, ByVal EndPoint As Point, ByVal rCell As Rectangle)
        'set the co-ordinates of the vertical lines that need to be drawn in the rows between the source and target rows
        Dim selDate As Date = CoordinateToDate(rCell.X, EndPoint.X)

        If intTargetIndex > intIndex Then
            For j = intIndex + 1 To intTargetIndex - 1
                Grid(j).KeyMomentLinksAsDate.Add(selDate)
            Next
        Else
            For j = intTargetIndex + 1 To intIndex - 1
                Grid(j).KeyMomentLinksAsDate.Add(selDate)
            Next
        End If
    End Sub

    Private Sub CellPainting_PlanningKeyMomentLinks_Incoming(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selKeyMoment As KeyMoment = selGridRow.KeyMoment
        Dim intLinkX As Integer
        Dim intLinkY As Integer = GetCoordinateY(rCell.Y, rCell.Height)
        Dim intIndex As Integer = Grid.IndexOf(selGridRow)

        For i = 0 To selGridRow.IncomingLinkIndices.Count - 1
            Dim intSourceIndex As Integer = selGridRow.IncomingLinkIndices(i)
            Dim SourceRow As PlanningGridRow = Me.Grid(intSourceIndex)
            Dim selSourceKeyMoment As KeyMoment = SourceRow.KeyMoment
            Dim selSourceActivity As Activity = TryCast(SourceRow.Struct, Activity)

            If SourceRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                intLinkX = GetCoordinateX(rCell.X, selKeyMoment.ExactDateKeyMoment)
            ElseIf SourceRow.RowType = PlanningGridRow.RowTypes.Activity Then
                intLinkX = GetCoordinateX(rCell.X, selSourceActivity.ExactStartDate)
            Else
                Exit Sub
            End If

            Dim LinkEnd As New Point(intLinkX, intLinkY)

            If Me.ShowKeyMomentLinks = True Then CellPainting_PlanningKeyMomentLinks_Incoming_Draw(intIndex, intSourceIndex, LinkEnd, rCell, CellGraphics)
        Next
    End Sub

    Private Sub CellPainting_PlanningKeyMomentLinks_Incoming_Draw(ByVal intIndex As Integer, ByVal intSourceIndex As Integer, ByVal LinkEnd As Point, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim EntryPoint As New Point
        Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line}
        Dim PointsLight(2) As Point
        Dim PointsDark(2) As Point

        If intSourceIndex < intIndex Then
            EntryPoint = New Point(LinkEnd.X, rCell.Top)
            PointsLight(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsLight(1) = New Point(LinkEnd.X - 4, LinkEnd.Y - 8)
            PointsLight(2) = New Point(LinkEnd.X, LinkEnd.Y - 8)
            PointsDark(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsDark(1) = New Point(LinkEnd.X + 4, LinkEnd.Y - 8)
            PointsDark(2) = New Point(LinkEnd.X, LinkEnd.Y - 8)
        Else

            'key moments are ordered by date, so only an arrow pointing downwards is needed
            LinkEnd.Y += CONST_BarHeight
            LinkEnd = New Point(LinkEnd.X, LinkEnd.Y)
            EntryPoint = New Point(LinkEnd.X, rCell.Bottom)
            PointsLight(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsLight(1) = New Point(LinkEnd.X + 4, LinkEnd.Y + 8)
            PointsLight(2) = New Point(LinkEnd.X, LinkEnd.Y + 8)
            PointsDark(0) = New Point(LinkEnd.X, LinkEnd.Y)
            PointsDark(1) = New Point(LinkEnd.X - 4, LinkEnd.Y + 8)
            PointsDark(2) = New Point(LinkEnd.X, LinkEnd.Y + 8)
        End If

        CellGraphics.DrawLine(penLightBlue2, EntryPoint, LinkEnd)
        CellGraphics.DrawLine(penDarkBlue1, EntryPoint, LinkEnd)
        CellGraphics.FillPath(Brushes.DeepSkyBlue, New GraphicsPath(PointsLight, PathPoints))
        CellGraphics.FillPath(Brushes.RoyalBlue, New GraphicsPath(PointsDark, PathPoints))
    End Sub

    Private Sub CellPainting_PlanningActivityLinks(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
        If selActivity Is Nothing Then Exit Sub

        If selGridRow.OutgoingLinksIndices.Count > 0 Then
            CellPainting_PlanningActivityLinks_Outgoing(selGridRow, rCell, CellGraphics)
        End If

        If selGridRow.IncomingLinkIndices.Count > 0 Then
            CellPainting_PlanningActivityLinks_Incoming(selGridRow, rCell, CellGraphics)
        End If
    End Sub

    Private Sub CellPainting_PlanningActivityLinks_Outgoing(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
        Dim intIndex As Integer = Grid.IndexOf(selGridRow)
        Dim intLinkX As Integer = GetCoordinateX(rCell.X, selActivity.ExactEndDate)
        Dim intLinkY As Integer = GetCoordinateY(rCell.Y, rCell.Height)
        Dim intTargetX As Integer
        Dim LinkStart As New Point(intLinkX, intLinkY)

        For i = 0 To selGridRow.OutgoingLinksIndices.Count - 1
            Dim intTargetIndex As Integer = selGridRow.OutgoingLinksIndices(i)
            Dim TargetRow As PlanningGridRow = Me.Grid(intTargetIndex)

            If intTargetIndex > intIndex Then
                LinkStart.Y = intLinkY + (CONST_PreparationHeight * 1.5) '+ CONST_BarHeight - 2
            Else
                LinkStart.Y = intLinkY + 2
            End If

            If TargetRow.RowType = PlanningGridRow.RowTypes.Activity Then
                Dim selTargetActivity As Activity = TryCast(TargetRow.Struct, Activity)

                intTargetX = GetCoordinateX(rCell.X, selTargetActivity.ExactStartDate) 'use the current cell to calculate since that is certainly visible
            ElseIf TargetRow.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                Dim selTargetKeyMoment As KeyMoment = TargetRow.KeyMoment

                intTargetX = GetCoordinateX(rCell.X, selTargetKeyMoment.ExactDateKeyMoment) 'use the current cell to calculate since that is certainly visible
            Else
                Exit Sub
            End If

            If Me.ShowActivityLinks = True Then CellPainting_PlanningActivityLinks_Outgoing_Draw(selActivity, intIndex, intTargetIndex, intTargetX, LinkStart, rCell, CellGraphics)
        Next
    End Sub

    Private Sub CellPainting_PlanningActivityLinks_Outgoing_Draw(ByVal selActivity As Activity, ByVal intIndex As Integer, ByVal intTargetIndex As Integer, _
                                                                 ByVal intTargetX As Integer, ByVal LinkStart As Point, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim FirstElbow, SecondElbow, ThirdElbow, EndPoint As New Point

        If intTargetX >= LinkStart.X + CONST_DistanceToActivity Then
            'if the target key moment/activity is not directly below or above the activity
            If selActivity.ActivityDetail.PeriodDirection = ActivityDetail.DirectionValues.Before Then
                'target activity is above source activity
                'draw small line up, then horizontal line to begin of target activity, then up to border
                If intTargetIndex > intIndex Then
                    FirstElbow = New Point(LinkStart.X, LinkStart.Y + 6)
                    SecondElbow = New Point(intTargetX, FirstElbow.Y)
                    EndPoint = New Point(SecondElbow.X, rCell.Bottom)
                Else
                    FirstElbow = New Point(LinkStart.X, LinkStart.Y - 6)
                    SecondElbow = New Point(intTargetX, FirstElbow.Y)
                    EndPoint = New Point(SecondElbow.X, rCell.Top)
                End If

                CellGraphics.DrawLine(penBlack2, LinkStart, FirstElbow)
                CellGraphics.DrawLine(penBlack2, FirstElbow, SecondElbow)
                CellGraphics.DrawLine(penBlack2, SecondElbow, EndPoint)
            Else
                'if the target key moment/activity is not directly below the key moment, draw a simple line with one right angle
                FirstElbow = New Point(intTargetX, LinkStart.Y)
                If intTargetIndex > intIndex Then
                    EndPoint = New Point(FirstElbow.X, rCell.Bottom)
                Else
                    EndPoint = New Point(FirstElbow.X, rCell.Top)
                End If

                CellGraphics.DrawLine(penBlack2, LinkStart, FirstElbow)
                CellGraphics.DrawLine(penBlack2, FirstElbow, EndPoint)
            End If

        Else
            'target activity is directly above/below source activity
            'draw half loop
            FirstElbow = New Point(LinkStart.X + CONST_DistanceToActivity, LinkStart.Y)
            If intTargetIndex > intIndex Then
                'loop below
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y + 6)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, rCell.Bottom)
            Else
                'loop above
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y - 6)
                ThirdElbow = New Point(intTargetX, SecondElbow.Y)
                EndPoint = New Point(intTargetX, rCell.Top)
            End If

            CellGraphics.DrawLine(penBlack2, LinkStart, FirstElbow)
            CellGraphics.DrawLine(penBlack2, FirstElbow, SecondElbow)
            CellGraphics.DrawLine(penBlack2, SecondElbow, ThirdElbow)
            CellGraphics.DrawLine(penBlack2, ThirdElbow, EndPoint)
        End If

        CellPainting_PlanningActivityLinks_VerticalLinks(intIndex, intTargetIndex, EndPoint, rCell)
    End Sub

    Private Sub CellPainting_PlanningActivityLinks_VerticalLinks(ByVal intIndex As Integer, ByVal intTargetIndex As Integer, ByVal EndPoint As Point, ByVal rCell As Rectangle)
        'set the co-ordinates of the vertical lines that need to be drawn in the rows between the source and target rows
        Dim selDate As Date = CoordinateToDate(rCell.X, EndPoint.X)

        If intTargetIndex > intIndex Then
            For j = intIndex + 1 To intTargetIndex - 1
                Grid(j).ActivityLinksAsDate.Add(selDate)
            Next
        Else
            For j = intTargetIndex + 1 To intIndex - 1
                Grid(j).ActivityLinksAsDate.Add(selDate)
            Next
        End If
    End Sub

    Private Sub CellPainting_PlanningActivityLinks_Incoming(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
        Dim intIndex As Integer = Grid.IndexOf(selGridRow)
        Dim intLinkX As Integer = GetCoordinateX(rCell.X, selActivity.ExactStartDate)
        Dim intLinkY As Integer = GetCoordinateY(rCell.Y, rCell.Height)
        Dim LinkEnd As New Point(intLinkX, intLinkY)
        Dim EntryPoint As New Point
        Dim PathPoints() As Byte = {PathPointType.Line, PathPointType.Line, PathPointType.Line}

        For i = 0 To selGridRow.IncomingLinkIndices.Count - 1
            Dim intSourceIndex As Integer = selGridRow.IncomingLinkIndices(i)
            Dim SourceRow As PlanningGridRow = Me.Grid(intSourceIndex)

            If SourceRow.RowType = PlanningGridRow.RowTypes.Activity And Me.ShowActivityLinks = True Then
                Dim Points(2) As Point
                Dim selSourceActivity As Activity = TryCast(SourceRow.Struct, Activity)
                If selSourceActivity Is Nothing Then Exit Sub

                If intSourceIndex < intIndex Then
                    EntryPoint = New Point(intLinkX, rCell.Top)
                    Points(0) = New Point(intLinkX, intLinkY)
                    Points(1) = New Point(intLinkX + 4, intLinkY - 8)
                    Points(2) = New Point(intLinkX - 4, intLinkY - 8)
                Else
                    intLinkY += CONST_BarHeight
                    LinkEnd = New Point(intLinkX, intLinkY)
                    EntryPoint = New Point(intLinkX, rCell.Bottom)
                    Points(0) = New Point(intLinkX, intLinkY)
                    Points(1) = New Point(intLinkX + 4, intLinkY + 8)
                    Points(2) = New Point(intLinkX - 4, intLinkY + 8)
                End If

                CellGraphics.DrawLine(penBlack2, EntryPoint, LinkEnd)
                CellGraphics.FillPath(Brushes.Black, New GraphicsPath(Points, PathPoints))
            ElseIf SourceRow.RowType = PlanningGridRow.RowTypes.KeyMoment And Me.ShowKeyMomentLinks = True Then
                Dim PointsLight(2) As Point
                Dim PointsDark(2) As Point

                If intSourceIndex < intIndex Then
                    EntryPoint = New Point(intLinkX, rCell.Top)
                    PointsLight(0) = New Point(intLinkX, intLinkY)
                    PointsLight(1) = New Point(intLinkX - 4, intLinkY - 8)
                    PointsLight(2) = New Point(intLinkX, intLinkY - 8)
                    PointsDark(0) = New Point(intLinkX, intLinkY)
                    PointsDark(1) = New Point(intLinkX + 4, intLinkY - 8)
                    PointsDark(2) = New Point(intLinkX, intLinkY - 8)
                Else
                    intLinkY += CONST_BarHeight
                    LinkEnd = New Point(intLinkX, intLinkY)
                    EntryPoint = New Point(intLinkX, rCell.Bottom)
                    PointsLight(0) = New Point(intLinkX, intLinkY)
                    PointsLight(1) = New Point(intLinkX + 4, intLinkY + 8)
                    PointsLight(2) = New Point(intLinkX, intLinkY + 8)
                    PointsDark(0) = New Point(intLinkX, intLinkY)
                    PointsDark(1) = New Point(intLinkX - 4, intLinkY + 8)
                    PointsDark(2) = New Point(intLinkX, intLinkY + 8)
                End If

                CellGraphics.DrawLine(penLightBlue2, EntryPoint, LinkEnd)
                CellGraphics.DrawLine(penDarkBlue1, EntryPoint, LinkEnd)
                CellGraphics.FillPath(Brushes.DeepSkyBlue, New GraphicsPath(PointsLight, PathPoints))
                CellGraphics.FillPath(Brushes.RoyalBlue, New GraphicsPath(PointsDark, PathPoints))
            End If
        Next
    End Sub

    Private Sub CellPainting_PlanningVerticalKeyMomentLinks(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim intLinkX As Integer

        If selGridRow.KeyMomentLinksAsDate.Count > 0 And Me.ShowKeyMomentLinks = True Then
            For i = 0 To selGridRow.KeyMomentLinksAsDate.Count - 1
                intLinkX = GetCoordinateX(rCell.X, selGridRow.KeyMomentLinksAsDate(i))
                CellGraphics.DrawLine(penLightBlue2, intLinkX, rCell.Top, intLinkX, rCell.Bottom)
                CellGraphics.DrawLine(penDarkBlue1, intLinkX, rCell.Top, intLinkX, rCell.Bottom)
            Next
        End If
    End Sub

    Private Sub CellPainting_PlanningVerticalActivityLinks(ByVal selGridRow As PlanningGridRow, ByVal rCell As Rectangle, ByVal CellGraphics As Graphics)
        Dim intLinkX As Integer

        If selGridRow.ActivityLinksAsDate.Count > 0 And Me.ShowActivityLinks = True Then
            For i = 0 To selGridRow.ActivityLinksAsDate.Count - 1
                intLinkX = GetCoordinateX(rCell.X, selGridRow.ActivityLinksAsDate(i))
                CellGraphics.DrawLine(penBlack2, intLinkX, rCell.Top, intLinkX, rCell.Bottom)
            Next
        End If
    End Sub
#End Region

#Region "Custom paint - Post paint"
    Protected Overrides Sub OnRowPostPaint(ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs)
        Dim RowGraphics As Graphics = e.Graphics
        Dim selGridRow As PlanningGridRow = Grid(e.RowIndex)

        If selGridRow IsNot Nothing Then
            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.RepeatPurpose, PlanningGridRow.RowTypes.RepeatOutput
                    'repeated purposes and outputs
                    RowPostPaint_PaintRepeatedRows(RowGraphics, e.RowBounds, selGridRow, e.RowIndex)
                Case Else
                    If boolLinkActivityStart = True Then
                        RowPostPaint_LinkActivityStart(RowGraphics)
                    ElseIf boolLinkActivityEnd = True Then
                        RowPostPaint_LinkActivityEnd(RowGraphics)
                    Else
                        'draw focus rectangle
                        DrawSelectionRectangle(RowGraphics)
                    End If
            End Select
        End If
    End Sub

    Private Sub RowPostPaint_PaintRepeatedRows(ByVal RowGraphics As Graphics, ByVal rRowBounds As Rectangle, ByVal selGridRow As PlanningGridRow, ByVal intRowIndex As Integer)
        'draw repeated Purposes/Outputs
        Dim rStructSort As Rectangle, rStructRTF As Rectangle
        Dim boolRepeat As Boolean = True
        Dim brBackGround As Brush
        Dim intX = 1, intY As Integer
        Dim fntRepeat As Font = New Font(CurrentLogFrame.SortNumberFont, FontStyle.Bold)
        Dim brRepeat As SolidBrush = New SolidBrush(Color.Black)
        Dim objStringFormat As New StringFormat()

        objStringFormat.Alignment = StringAlignment.Near
        objStringFormat.LineAlignment = StringAlignment.Near

        If Me.RowHeadersVisible = True Then intX = Me.RowHeadersWidth
        intY = rRowBounds.Top

        Dim rRow As New Rectangle(intX, intY, Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - HorizontalScrollingOffset, rRowBounds.Height)

        If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose Then brBackGround = Brushes.CornflowerBlue Else brBackGround = Brushes.LightSteelBlue
        RowGraphics.FillRectangle(brBackGround, rRow)

        'draw number
        rStructSort = GetCellDisplayRectangle(colSortNumber.Index, intRowIndex, False)
        rStructSort.X += 2

        RowGraphics.DrawString(selGridRow.SortNumber, fntRepeat, brRepeat, rStructSort, objStringFormat)

        'draw text
        rStructRTF = GetCellDisplayRectangle(colRTF.Index, intRowIndex, False)
        Dim strText As String = selGridRow.Struct.Text
        Dim rRepeat As New Rectangle(rStructRTF.X, rStructRTF.Y, _
                                     Columns.GetColumnsWidth(DataGridViewElementStates.Visible) - rStructSort.Width - HorizontalScrollingOffset + 1, rRowBounds.Height)
        Dim intHeight As Integer = RowGraphics.MeasureString(strText, CurrentLogFrame.SectionTitleFont, rRepeat.Width, objStringFormat).Height
        rRepeat.Height = intHeight
        Me.Rows(intRowIndex).Height = intHeight
        RowGraphics.DrawString(strText, CurrentLogFrame.SectionTitleFont, brRepeat, rRepeat, objStringFormat)
        RowGraphics.DrawRectangle(penBorder, rRow)
    End Sub

    Private Sub RowPostPaint_LinkActivityStart(ByVal RowGraphics As Graphics)
        Dim LinkStart, LinkEnd, FirstElbow, SecondElbow, ThirdElbow As New Point
        Dim penLink As New Pen(Color.Blue, 2)

        If boolLinkHoverOverActivity = True Then 'Or boolLinkHoverOverKeyMoment = True 
            penLink.DashStyle = DashStyle.Solid
        Else
            penLink.DashStyle = DashStyle.Dot
        End If


        LinkEnd = ptLinkEnd
        If LinkEnd.X <= ptLinkStart.X - CONST_DistanceToActivity Then
            'if the target key moment/activity is not directly below the key moment, draw a simple line with one right angle

            If ptLinkStart.Y > ptLinkEnd.Y Then
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + 2)
            Else
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + CONST_BarHeight - 2)
            End If
            FirstElbow = New Point(LinkEnd.X, LinkStart.Y)

            RowGraphics.DrawLine(penLink, LinkStart, FirstElbow)
            RowGraphics.DrawLine(penLink, FirstElbow, LinkEnd)
        Else
            'target activity is directly above/below source activity
            'draw half loop

            If ptLinkStart.Y > ptLinkEnd.Y Then
                'loop above
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + 2)
                FirstElbow = New Point(LinkStart.X - CONST_DistanceToActivity, LinkStart.Y)
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y - 6)
                ThirdElbow = New Point(LinkEnd.X, SecondElbow.Y)
            Else
                'loop below
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + CONST_BarHeight - 2)
                FirstElbow = New Point(LinkStart.X - CONST_DistanceToActivity, LinkStart.Y)
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y + 6)
                ThirdElbow = New Point(LinkEnd.X, SecondElbow.Y)
            End If

            RowGraphics.DrawLine(penLink, LinkStart, FirstElbow)
            RowGraphics.DrawLine(penLink, FirstElbow, SecondElbow)
            RowGraphics.DrawLine(penLink, SecondElbow, ThirdElbow)
            RowGraphics.DrawLine(penLink, ThirdElbow, LinkEnd)
        End If
        Dim rCircle As New Rectangle(LinkEnd.X - 4, LinkEnd.Y - 4, 8, 8)
        RowGraphics.FillEllipse(Brushes.Blue, rCircle)
    End Sub

    Private Sub RowPostPaint_LinkActivityEnd(ByVal RowGraphics As Graphics)
        Dim LinkStart, LinkEnd, FirstElbow, SecondElbow, ThirdElbow As New Point
        Dim penLink As New Pen(Color.Blue, 2)

        If boolLinkHoverOverActivity = True Or boolLinkHoverOverKeyMoment = True Then
            penLink.DashStyle = DashStyle.Solid
        Else
            penLink.DashStyle = DashStyle.Dot
        End If


        LinkEnd = ptLinkEnd
        If LinkEnd.X >= ptLinkStart.X + CONST_DistanceToActivity Then
            'if the target key moment/activity is not directly below the key moment, draw a simple line with one right angle

            If ptLinkStart.Y > ptLinkEnd.Y Then
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + 2)
            Else
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + CONST_BarHeight - 2)
            End If
            FirstElbow = New Point(LinkEnd.X, LinkStart.Y)

            RowGraphics.DrawLine(penLink, LinkStart, FirstElbow)
            RowGraphics.DrawLine(penLink, FirstElbow, LinkEnd)
        Else
            'target activity is directly above/below source activity
            'draw half loop

            If ptLinkStart.Y > ptLinkEnd.Y Then
                'loop above
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + 2)
                FirstElbow = New Point(LinkStart.X + CONST_DistanceToActivity, LinkStart.Y)
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y - 6)
                ThirdElbow = New Point(LinkEnd.X, SecondElbow.Y)
            Else
                'loop below
                LinkStart = New Point(ptLinkStart.X, ptLinkStart.Y + CONST_BarHeight - 2)
                FirstElbow = New Point(LinkStart.X + CONST_DistanceToActivity, LinkStart.Y)
                SecondElbow = New Point(FirstElbow.X, LinkStart.Y + 6)
                ThirdElbow = New Point(LinkEnd.X, SecondElbow.Y)
            End If

            RowGraphics.DrawLine(penLink, LinkStart, FirstElbow)
            RowGraphics.DrawLine(penLink, FirstElbow, SecondElbow)
            RowGraphics.DrawLine(penLink, SecondElbow, ThirdElbow)
            RowGraphics.DrawLine(penLink, ThirdElbow, LinkEnd)
        End If
        Dim rCircle As New Rectangle(LinkEnd.X - 4, LinkEnd.Y - 4, 8, 8)
        RowGraphics.FillEllipse(Brushes.Blue, rCircle)
    End Sub
#End Region

#Region "Selection rectangle"
    Private Sub InvalidateSelectionRectangle()
        With SelectionRectangle
            Dim rSelection As New Rectangle(.Rectangle.X - 2, .Rectangle.Y - 2, .Rectangle.Width + 4, .Rectangle.Height + 4)
            Me.Invalidate(rSelection)
        End With
    End Sub

    Private Sub DrawSelectionRectangle(ByVal graph As Graphics)
        SetSelectionRectangle()

        With SelectionRectangle
            Dim IndexLastDisplayedRow As Integer = Me.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            If .FirstRowIndex > IndexLastDisplayedRow And .LastRowIndex > IndexLastDisplayedRow Then Exit Sub

            Dim IndexFirstDisplayedRow As Integer = Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
            If .FirstRowIndex < IndexFirstDisplayedRow And .LastRowIndex < IndexFirstDisplayedRow Then Exit Sub


            graph.DrawRectangle(penSelection, .Rectangle)
        End With
    End Sub

    Public Sub SetSelectionRectangle()
        Dim rTopLeftCell As New Rectangle
        Dim rBottomRightCell As New Rectangle
        Dim FirstRowIndexSelection As Integer

        If EditingControl IsNot Nothing Then
            With SelectionRectangle
                .FirstRowIndex = CurrentCell.RowIndex
                .LastRowIndex = CurrentCell.RowIndex
                .FirstColumnIndex = CurrentCell.ColumnIndex - 1
                .LastColumnIndex = CurrentCell.ColumnIndex
            End With
        ElseIf SelectedCells.Count > 0 Then
            SetSelectionRectangleGridArea()
        Else
            Exit Sub
        End If

        With SelectionRectangle
            If .FirstRowIndex < 0 Or .LastRowIndex < 0 Then Exit Sub
            FirstRowIndexSelection = .FirstRowIndex

            rTopLeftCell = GetCellDisplayRectangle(.FirstColumnIndex, .FirstRowIndex, False)


            If Me.Rows(.LastRowIndex).Displayed = False And .LastRowIndex >= Me.Rows.GetFirstRow(DataGridViewElementStates.Displayed) Then
                .LastRowIndex = Me.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            End If
            rBottomRightCell = GetCellDisplayRectangle(.LastColumnIndex, .LastRowIndex, True)

            Dim intLeft As Integer = GetColumnDisplayRectangle(.FirstColumnIndex, True).X
            Dim intTop As Integer = rTopLeftCell.Y
            Dim intFirstColumnWidth As Integer = GetColumnDisplayRectangle(.FirstColumnIndex, True).Width
            Dim intFirstRowHeight As Integer = rTopLeftCell.Height
            Dim intRight As Integer = GetColumnDisplayRectangle(.LastColumnIndex, True).X
            Dim intBottom As Integer = rBottomRightCell.Y
            Dim intLastColumnWidth As Integer = GetColumnDisplayRectangle(.LastColumnIndex, True).Width
            Dim intLastRowHeight As Integer = rBottomRightCell.Height

            If .FirstRowIndex = .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intFirstColumnWidth
                .Height = intFirstRowHeight
            ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intFirstColumnWidth
                .Height = intBottom + intLastRowHeight - intTop
            ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex = .LastColumnIndex Then
                .X = intRight
                .Y = intBottom
                .Width = intLastColumnWidth
                .Height = intTop + intFirstRowHeight - intBottom
            ElseIf .FirstRowIndex = .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intRight + intLastColumnWidth - intLeft
                .Height = intFirstRowHeight
            ElseIf .FirstRowIndex = .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                .X = intRight
                .Y = intBottom
                .Width = intLeft + intFirstColumnWidth - intRight
                .Height = intLastRowHeight

            ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                .X = intLeft
                .Y = intTop
                .Width = intRight + intLastColumnWidth - intLeft
                .Height = intBottom + intLastRowHeight - intTop
            ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex < .LastColumnIndex Then
                .X = intLeft
                .Y = intBottom
                .Width = intRight + intLastColumnWidth - intLeft
                .Height = intTop + intFirstRowHeight - intBottom
            ElseIf .FirstRowIndex < .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                .X = intRight
                .Y = intTop
                .Width = intLeft + intFirstColumnWidth - intRight
                .Height = intBottom + intLastRowHeight - intTop
            ElseIf .FirstRowIndex > .LastRowIndex And .FirstColumnIndex > .LastColumnIndex Then
                .X = intRight
                .Y = intBottom
                .Width = intLeft + intFirstColumnWidth - intRight
                .Height = intTop + intFirstRowHeight - intBottom
            End If

            .X += 1
            .Y += 1
            .Width -= 3
            .Height -= 3

            If .Rectangle.Bottom = -1 And .LastRowIndex >= Me.FirstDisplayedScrollingRowIndex Then .Height = Me.Bottom - .Y - 1
        End With
    End Sub

    Public Sub SetSelectionRectangleGridArea()
        Dim Vdir As Integer
        Dim intSelSize As Integer = SelectedCells.Count - 1
        If intSelSize < 0 Then intSelSize = 0

        If SelectedCells(0).RowIndex >= SelectedCells(intSelSize).RowIndex Then Vdir = 1 Else Vdir = -1

        SelectionRectangle.FirstRowIndex = CurrentCell.RowIndex
        SelectionRectangle.FirstColumnIndex = CurrentCell.ColumnIndex
        SelectionRectangle.LastRowIndex = CurrentCell.RowIndex
        SelectionRectangle.LastColumnIndex = CurrentCell.ColumnIndex

        If SelectedCells.Count > 1 Then
            For i = 0 To SelectedCells.Count - 1
                If SelectedCells(i).RowIndex < SelectionRectangle.FirstRowIndex Then SelectionRectangle.FirstRowIndex = SelectedCells(i).RowIndex
                If SelectedCells(i).RowIndex > SelectionRectangle.LastRowIndex Then SelectionRectangle.LastRowIndex = SelectedCells(i).RowIndex
                SelectionRectangle.FirstColumnIndex = 0
                SelectionRectangle.LastColumnIndex = CONST_PlanningColumnIndex
            Next

            SetSelectionRectangleGridArea_Modify(Vdir)
        End If
    End Sub

    Public Sub SetSelectionRectangleGridArea_Modify(ByVal Vdir As Integer)
        Dim intRowType As Integer
        Dim boolRest As Boolean
        Dim i As Integer

        With SelectionRectangle
            'do not select title rows
            If Me.Grid(.FirstRowIndex).RowType = PlanningGridRow.RowTypes.RepeatPurpose Then
                .FirstRowIndex += 1
                If .LastRowIndex > .FirstRowIndex Then .LastRowIndex = .FirstRowIndex
            End If
            If Me.Grid(.FirstRowIndex).RowType = PlanningGridRow.RowTypes.RepeatOutput Then
                .FirstRowIndex += 1
                If .LastRowIndex > .FirstRowIndex Then .LastRowIndex = .FirstRowIndex
            End If

            'if selection is larger than section, limit selection to that section
            If .LastRowIndex > .FirstRowIndex Then

                If Vdir > 0 Then
                    intRowType = Me.Grid(.FirstRowIndex).RowType
                    For i = .FirstRowIndex + 1 To .LastRowIndex
                        If Me.Grid(i).RowType <> intRowType And boolRest = False Then
                            boolRest = True
                            .LastRowIndex = i - 1
                        End If
                    Next
                Else
                    intRowType = Me.Grid(.LastRowIndex).RowType
                    For i = .LastRowIndex - 1 To .FirstRowIndex Step -1
                        If Me.Grid(i).RowType <> intRowType And boolRest = False Then
                            boolRest = True
                            .FirstRowIndex = i + 1
                        End If
                    Next
                End If
            End If
            If .LastRowIndex < 0 Then .LastRowIndex = 0
        End With
    End Sub
#End Region
End Class
