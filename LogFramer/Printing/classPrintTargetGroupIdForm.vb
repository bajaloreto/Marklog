Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintTargetGroupIdForm
    Inherits ReportBaseClass

    Private objPrintList As New PrintTargetGroupIdFormRows
    Private objTargetGroups As New TargetGroups

    Private boolColumnsWidthSet As Boolean
    Private rPropertyRectangle, rWhiteSpaceRectangle As Rectangle

    Private boolOffPage As Boolean
    Private intPropertyIndex As Integer
    Private intTargetGroupIndex As Integer

    Private boolPrintBorders As Boolean = True
    Private boolFillCells As Boolean

    Private fntNormal As New Font(CurrentLogFrame.DetailsFont, FontStyle.Regular)
    Private fntNormalBold As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)

    Private Const CONST_PropertyColumnWidth As Integer = 200
    Private Const CONST_Spacing = 25
    Private Const CONST_CheckBoxDimension = 18

    Public Event LinePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)

    Public Sub New()

    End Sub

    Public Sub New(ByVal targetgroup As TargetGroup)
        Me.TargetGroups.Clear()
        Me.TargetGroups.Add(targetgroup)

        Me.ReportSetup = CurrentLogFrame.ReportSetupTargetGroupIdForm

        CreateList()
    End Sub

    Public Sub New(ByVal targetgroups As TargetGroups)
        Me.TargetGroups = targetgroups

        Me.ReportSetup = CurrentLogFrame.ReportSetupTargetGroupIdForm
    End Sub

#Region "Properties"
    Public Property TargetGroups() As TargetGroups
        Get
            Return objTargetGroups
        End Get
        Set(ByVal value As TargetGroups)
            objTargetGroups = value
        End Set
    End Property

    Public Property PrintList() As PrintTargetGroupIdFormRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintTargetGroupIdFormRows)
            objPrintList = value
        End Set
    End Property

    Public Property PrintBorders As Boolean
        Get
            Return boolPrintBorders
        End Get
        Set(ByVal value As Boolean)
            boolPrintBorders = value
        End Set
    End Property

    Public Property FillCells As Boolean
        Get
            Return boolFillCells
        End Get
        Set(ByVal value As Boolean)
            boolFillCells = value
        End Set
    End Property

    Private Property PropertyRectangle As Rectangle
        Get
            Return rPropertyRectangle
        End Get
        Set(ByVal value As Rectangle)
            rPropertyRectangle = value
        End Set
    End Property

    Private Property WhiteSpaceRectangle As Rectangle
        Get
            Return rWhiteSpaceRectangle
        End Get
        Set(ByVal value As Rectangle)
            rWhiteSpaceRectangle = value
        End Set
    End Property
#End Region

#Region "Create list of targetgroup properties"
    Public Sub CreateList()
        Dim strText As String
        Dim newRow As PrintTargetGroupIdFormRow

        PrintList.Clear()

        For Each selTargetgroup As TargetGroup In Me.TargetGroups
            PrintList.Add(New PrintTargetGroupIdFormRow(selTargetgroup.Name, True))
            For Each selProperty As TargetGroupInformation In selTargetgroup.TargetGroupInformations
                strText = ToLabel(selProperty.Name)

                newRow = New PrintTargetGroupIdFormRow(strText, selProperty.Type)
                PrintList.Add(newRow)

                Select Case selProperty.Type
                    Case TargetGroupInformation.PropertyTypes.List
                        For Each selValue As ChecklistOption In selProperty.CheckListOptions
                            PrintList.Add(New PrintTargetGroupIdFormRow(String.Empty, selProperty.Type, selValue.OptionName))
                        Next
                    Case TargetGroupInformation.PropertyTypes.Number
                        If String.IsNullOrEmpty(selProperty.Unit) = False Then
                            newRow.PropertyValue = selProperty.Unit
                        End If
                    Case TargetGroupInformation.PropertyTypes.YesNo
                        PrintList.Add(New PrintTargetGroupIdFormRow(String.Empty, selProperty.Type, LANG_Yes))
                        PrintList.Add(New PrintTargetGroupIdFormRow(String.Empty, selProperty.Type, LANG_No))
                End Select
            Next
        Next
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth
        Dim intTextColumnWidth As Integer = intAvailableWidth - CONST_PropertyColumnWidth

        PropertyRectangle = New Rectangle(LeftMargin, ContentTop, CONST_PropertyColumnWidth, ContentHeight)
        WhiteSpaceRectangle = New Rectangle(PropertyRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight)
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintTargetGroupIdFormRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > 0 Then selPrintListRow.RowHeight = intRowHeight Else selPrintListRow.RowHeight = NewCellHeight()
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim selRow As PrintTargetGroupIdFormRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer, intCellHeight, intValueHeight As Integer
        Dim intWidth As Integer

        If PageGraph IsNot Nothing Then
            If String.IsNullOrEmpty(selRow.PropertyName) = False Then
                If selRow.IsTitle = True Then
                    intCellHeight = PageGraph.MeasureString(selRow.PropertyName, fntNormalBold, ContentWidth - CONST_HorizontalPadding).Height
                    If RowIndex > 0 And selRow.IsTitle Then intCellHeight += CONST_Spacing
                Else
                    intCellHeight = PageGraph.MeasureString(selRow.PropertyName, fntNormalBold, PropertyRectangle.Width - CONST_HorizontalPadding).Height
                End If
            End If

            If String.IsNullOrEmpty(selRow.PropertyValue) = False Then
                intWidth = WhiteSpaceRectangle.Width - CONST_HorizontalPadding

                If selRow.PropertyType = TargetGroupInformation.PropertyTypes.List Then intWidth -= CONST_CheckBoxDimension - 6
                intValueHeight = PageGraph.MeasureString(selRow.PropertyValue, fntNormal, intWidth).Height

                If selRow.PropertyType = TargetGroupInformation.PropertyTypes.List And intValueHeight < CONST_CheckBoxDimension Then
                    intValueHeight = CONST_CheckBoxDimension
                End If

                If intValueHeight > intCellHeight Then intCellHeight = intValueHeight
            End If

            If intCellHeight > intRowHeight Then intRowHeight = intCellHeight


            intRowHeight += CONST_VerticalPadding

            Return intRowHeight
        Else
            Return 0
        End If
    End Function
#End Region

#Region "General methods"
    Private Function GetTotalPages() As Integer
        Dim intTotalHeight As Integer
        Dim decPages As Decimal
        Dim intAvailableHeight As Integer = Me.ContentHeight

        For Each selRow As PrintTargetGroupIdFormRow In PrintList
            intTotalHeight += selRow.RowHeight
        Next

        decPages = intTotalHeight / intAvailableHeight
        decPages = Math.Ceiling(decPages)
        'decPages *= HorPages

        Return decPages
    End Function
#End Region

#Region "Print page"
    Protected Overrides Sub OnBeginPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnBeginPrint(e)

        boolColumnsWidthSet = False
        PrintRectangles.Clear()
        LastRowY = ContentTop

        CreateList()
    End Sub

    Protected Overrides Sub OnQueryPageSettings(ByVal e As System.Drawing.Printing.QueryPageSettingsEventArgs)
        MyBase.OnQueryPageSettings(e)
    End Sub

    Protected Overrides Sub OnEndPrint(ByVal e As System.Drawing.Printing.PrintEventArgs)
        MyBase.OnEndPrint(e)
    End Sub

    Protected Overrides Sub OnPrintPage(ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        MyBase.OnPrintPage(e)

        PageGraph = e.Graphics

        If boolColumnsWidthSet = False Then
            SetColumnsWidth()
            ResetRowHeights()
            Me.TotalPages = GetTotalPages()

            boolColumnsWidthSet = True
        End If

        Dim intRowCount As Integer = PrintList.Count
        Dim selRow As PrintTargetGroupIdFormRow
        Dim intMinHeight As Integer

        LastRowY = ContentTop

        'Print Headers
        PrintHeader()

        If PrintList.Count > 0 Then
            Do
                selRow = PrintList(RowIndex)
                If selRow Is Nothing Then Exit Do

                PrintPage_PrintPropertyName(selRow)
                PrintPage_PrintWhiteSpace(selRow)

                LastRowY += selRow.RowHeight
                RowIndex += 1

                If RowIndex > PrintList.Count - 1 Then Exit Do

                intMinHeight = LastRowY + PrintList(RowIndex).RowHeight

                RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(RowIndex, PrintList.Count))
            Loop While intMinHeight < Me.ContentBottom
        Else
            RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(0, 0))
        End If

        PrintFooter()

        If PrintList.Count = 0 Then RaiseEvent LinePrinted(Me, New LinePrintedEventArgs(0, 0))

        If RowIndex < intRowCount Then
            LastRowY = ContentTop
            PageNumber += 1
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    Private Sub PrintPage_PrintPropertyName(ByVal selRow As PrintTargetGroupIdFormRow)
        Dim rCell As Rectangle, rText As Rectangle

        If selRow.PropertyName <> String.Empty Then
            If selRow.IsTitle = True Then
                Dim intCorrection As Integer

                If PrintList.IndexOf(selRow) > 0 Then
                    intCorrection = CONST_Spacing
                End If

                rCell = New Rectangle(LeftMargin, LastRowY + intCorrection, ContentWidth, selRow.RowHeight - intCorrection)
            Else
                rCell = New Rectangle(LeftMargin, LastRowY, PropertyRectangle.Width, selRow.RowHeight)
            End If

            rText = GetTextRectangle(rCell)

            PrintBackground(selRow.IsTitle, rCell)
            If selRow.IsTitle = False Then
                PageGraph.DrawString(selRow.PropertyName, fntNormal, Brushes.Black, rText)
            Else
                PageGraph.DrawString(selRow.PropertyName, fntNormalBold, Brushes.Black, rText)
            End If
            If PrintBorders = True Then PageGraph.DrawRectangle(penBlack1, rCell)
        End If
    End Sub

    Private Sub PrintPage_PrintWhiteSpace(ByVal selRow As PrintTargetGroupIdFormRow)
        Dim intHeight As Integer = selRow.RowHeight
        Dim intWidth As Integer
        Dim rCell As Rectangle, rText, rCheck As Rectangle


        Select Case selRow.PropertyType
            Case TargetGroupInformation.PropertyTypes.DateType
                intWidth = PageGraph.MeasureString("/", fntNormal).Width

                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, 20, intHeight)
                PrintWhiteSpace_Cells(rCell)

                rText = New Rectangle(rCell.Right, rCell.Top, intWidth, intHeight)
                PageGraph.DrawString("/", fntNormal, Brushes.Black, rText)

                rCell = New Rectangle(rText.Right, rText.Top, 20, intHeight)
                PrintWhiteSpace_Cells(rCell)

                rText = New Rectangle(rCell.Right, rCell.Top, intWidth, intHeight)
                PageGraph.DrawString("/", fntNormal, Brushes.Black, rText)

                rCell = New Rectangle(rText.Right, rText.Top, 40, intHeight)
                PrintWhiteSpace_Cells(rCell)

            Case TargetGroupInformation.PropertyTypes.List
                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, WhiteSpaceRectangle.Width, intHeight)
                PrintWhiteSpace_Cells(rCell, True)

                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, CONST_CheckBoxDimension, CONST_CheckBoxDimension)
                rCheck = GetTextRectangle(rCell)
                If selRow.PropertyValue <> String.Empty Then PageGraph.DrawRectangle(penBlack1, rCheck)

                intWidth = WhiteSpaceRectangle.Width - CONST_CheckBoxDimension - 6
                rText = New Rectangle(rCell.Right + 6, rCell.Top, intWidth, intHeight)
                PageGraph.DrawString(selRow.PropertyValue, fntNormal, Brushes.Black, rText)

            Case TargetGroupInformation.PropertyTypes.Number
                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, 100, intHeight)
                PrintWhiteSpace_Cells(rCell)

                If String.IsNullOrEmpty(selRow.PropertyValue) = False Then
                    intWidth = Me.ContentWidth - LeftMargin - rCell.Right
                    rText = New Rectangle(rCell.Right, LastRowY, intWidth, intHeight)
                    PageGraph.DrawString(selRow.PropertyValue, fntNormal, Brushes.Black, rText)
                End If

            Case TargetGroupInformation.PropertyTypes.Text
                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, WhiteSpaceRectangle.Width, intHeight)
                PrintWhiteSpace_Cells(rCell)

            Case TargetGroupInformation.PropertyTypes.YesNo
                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, WhiteSpaceRectangle.Width, intHeight)
                PrintBackground(False, rCell)

                intWidth = selRow.RowHeight
                rCell = New Rectangle(WhiteSpaceRectangle.X, LastRowY, intWidth, selRow.RowHeight)
                rCheck = GetTextRectangle(rCell)
                PageGraph.DrawRectangle(penBlack1, rCheck)

                intWidth = ContentWidth - LeftMargin - rCell.Right - 6
                rText = New Rectangle(rCell.Right + 6, rCell.Top, intWidth, intHeight)
                PageGraph.DrawString(selRow.PropertyValue, fntNormal, Brushes.Black, rText)
        End Select
    End Sub

    Private Sub PrintWhiteSpace_Cells(ByVal rCell As Rectangle, Optional ByVal boolNoDottedLine As Boolean = False)
        PrintBackground(False, rCell)

        If PageGraph IsNot Nothing Then
            If Me.PrintBorders = True Then
                PageGraph.DrawRectangle(penBlack1, rCell)
            Else
                If FillCells = False And boolNoDottedLine = False Then
                    Dim penGrayDotted As New Pen(Color.DarkGray, 1)
                    penGrayDotted.DashStyle = DashStyle.Dot

                    rCell.Height -= CONST_VerticalPadding
                    PageGraph.DrawLine(penGrayDotted, rCell.Left, rCell.Bottom, rCell.Right, rCell.Bottom)
                End If
            End If
        End If
    End Sub

    Private Sub PrintBackground(ByVal boolIsTitle As Boolean, ByVal rCell As Rectangle)
        If PageGraph IsNot Nothing Then
            If FillCells = True Then
                Dim brBackGround As Brush
                If Int(RowIndex / 2) = RowIndex / 2 Then
                    brBackGround = Brushes.LightGray
                Else
                    brBackGround = Brushes.Gainsboro
                End If
                PageGraph.FillRectangle(brBackGround, rCell)
            Else
                If boolIsTitle = True Then
                    PageGraph.FillRectangle(Brushes.LightGray, rCell)
                    PageGraph.DrawRectangle(penBlack1, rCell)
                End If
            End If
        End If
    End Sub
#End Region
End Class


