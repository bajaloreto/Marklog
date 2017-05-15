Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing

Public Class PrintPartnerList
    Inherits ReportBaseClass

    Private objPrintList As New PrintPartnerRows
    Private objProjectPartners As New ProjectPartners

    Private boolColumnsWidthSet As Boolean
    Private rPropertyRectangle, rSecondaryPropertyRectangle, rTextRectangle As Rectangle
    
    Private boolPrintBorders As Boolean = True
    Private boolFillCells As Boolean

    Private fntNormal As New Font(CurrentLogFrame.DetailsFont, FontStyle.Regular)
    Private fntNormalBold As New Font(CurrentLogFrame.DetailsFont, FontStyle.Bold)

    Private Const CONST_PropertyWidth As Integer = 200
    Private Const CONST_SecondaryPropertyWidth As Integer = 175
    Private Const CONST_Spacing = 25

    Public Event LinePrinted(ByVal sender As Object, ByVal e As LinePrintedEventArgs)

    Public Sub New()

    End Sub

    Public Sub New(ByVal projectpartners As ProjectPartners)
        Me.ProjectPartners.Clear()
        Me.ProjectPartners = projectpartners

        Me.ReportSetup = CurrentLogFrame.ReportSetupPartnerList
        '
        CreateList()
    End Sub

    Public Sub New(ByVal projectpartner As ProjectPartner)
        Me.ProjectPartners.Clear()
        Me.ProjectPartners.Add(projectpartner)

        Me.ReportSetup = CurrentLogFrame.ReportSetupPartnerList
    End Sub

#Region "Properties"
    Public Enum ObjectTypes
        Organisation = 0
        Address = 1
        TelephoneNumber = 2
        Email = 3
        Website = 4
        Contact = 5
        BankAccount = 6
    End Enum

    Public Property ProjectPartners() As ProjectPartners
        Get
            Return objProjectPartners
        End Get
        Set(ByVal value As ProjectPartners)
            objProjectPartners = value
        End Set
    End Property

    Public Property PrintList() As PrintPartnerRows
        Get
            Return objPrintList
        End Get
        Set(ByVal value As PrintPartnerRows)
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

    Private Property SecondaryPropertyRectangle As Rectangle
        Get
            Return rSecondaryPropertyRectangle
        End Get
        Set(ByVal value As Rectangle)
            rSecondaryPropertyRectangle = value
        End Set
    End Property

    Private Property TextRectangle As Rectangle
        Get
            Return rTextRectangle
        End Get
        Set(ByVal value As Rectangle)
            rTextRectangle = value
        End Set
    End Property
#End Region

#Region "Create list of partners"
    Public Sub CreateList()
        PrintList.Clear()

        For Each selProjectPartner As ProjectPartner In Me.ProjectPartners
            Dim selOrganisation As Organisation = selProjectPartner.Organisation
            PrintList.Add(New PrintPartnerRow(selOrganisation.Name, ToLabel(LANG_Organisation), 0, True))
            PrintList.Add(New PrintPartnerRow(selOrganisation.Acronym, ToLabel(LANG_Acronym), 0))
            PrintList.Add(New PrintPartnerRow(selOrganisation.TypeName, ToLabel(LANG_Type), 0))
            PrintList.Add(New PrintPartnerRow(selProjectPartner.RoleName, ToLabel(LANG_Role), 0, True))

            If selOrganisation.Addresses.Count > 0 Then _
                PrintList.Add(New PrintPartnerRow(String.Empty, LANG_Addresses, 0))
            For Each selAddress As Address In selOrganisation.Addresses
                PrintList.Add(New PrintPartnerRow(selAddress.TypeName, ToLabel(LANG_AddressType), 1))
                PrintList.Add(New PrintPartnerRow(selAddress.FullStreet, ToLabel(LANG_Address), 1))
                PrintList.Add(New PrintPartnerRow(selAddress.FullTown, ToLabel(LANG_Place), 1))
                PrintList.Add(New PrintPartnerRow(selAddress.Country, ToLabel(LANG_Country), 1))
            Next

            If selOrganisation.TelephoneNumbers.Count > 0 Then _
                PrintList.Add(New PrintPartnerRow(String.Empty, LANG_TelephoneNumbers, 0))
            For Each selNumber As TelephoneNumber In selOrganisation.TelephoneNumbers
                PrintList.Add(New PrintPartnerRow(selNumber.TypeName, ToLabel(LANG_NumberType), 2))
                PrintList.Add(New PrintPartnerRow(selNumber.Number, ToLabel(LANG_TelephoneNumber), 2))
            Next

            If selOrganisation.Emails.Count > 0 Then _
                PrintList.Add(New PrintPartnerRow(String.Empty, LANG_EmailAddresses, 0))
            For Each selEmail As Email In selOrganisation.Emails
                PrintList.Add(New PrintPartnerRow(selEmail.TypeName, ToLabel(LANG_EmailType), 3))
                PrintList.Add(New PrintPartnerRow(selEmail.Email, ToLabel(LANG_EmailAddress), 3))
            Next

            If selOrganisation.WebSites.Count > 0 Then _
                PrintList.Add(New PrintPartnerRow(String.Empty, LANG_Websites, 0))
            For Each selWebsite As Website In selOrganisation.WebSites
                PrintList.Add(New PrintPartnerRow(selWebsite.TypeName, ToLabel(LANG_WebsiteType), 4))
                PrintList.Add(New PrintPartnerRow(selWebsite.WebsiteName, ToLabel(LANG_WebsiteName), 4))
                PrintList.Add(New PrintPartnerRow(selWebsite.WebsiteUrl, ToLabel(LANG_URL), 4))
            Next

            If selOrganisation.Contacts.Count > 0 Then _
                PrintList.Add(New PrintPartnerRow(String.Empty, LANG_ContactPersons, 0))
            For Each selContact As Contact In selOrganisation.Contacts
                PrintList.Add(New PrintPartnerRow(selContact.FullName, ToLabel(LANG_Name), 5))
                PrintList.Add(New PrintPartnerRow(selContact.JobTitle, ToLabel(LANG_JobTitle), 5))
                PrintList.Add(New PrintPartnerRow(selContact.Role, ToLabel(LANG_Role), 5))
            Next

            If selOrganisation.BankAccount <> String.Empty Then
                PrintList.Add(New PrintPartnerRow(String.Empty, LANG_BankAccounts, 0))
                PrintList.Add(New PrintPartnerRow(selOrganisation.BankAccount, ToLabel(LANG_BankAccount), 6))
                PrintList.Add(New PrintPartnerRow(selOrganisation.Swift, ToLabel(LANG_BIC), 6))
            End If
        Next
    End Sub
#End Region

#Region "Set column widths"
    Private Sub SetColumnsWidth()
        Dim intAvailableWidth As Integer = Me.ContentWidth
        Dim intTextColumnWidth As Integer = intAvailableWidth - CONST_PropertyWidth

        PropertyRectangle = New Rectangle(LeftMargin, ContentTop, CONST_PropertyWidth, ContentHeight)
        SecondaryPropertyRectangle = New Rectangle(LeftMargin + CONST_PropertyWidth - CONST_SecondaryPropertyWidth, ContentTop, CONST_SecondaryPropertyWidth, ContentHeight)
        TextRectangle = New Rectangle(PropertyRectangle.Right, ContentTop, intTextColumnWidth, ContentHeight)
    End Sub
#End Region

#Region "Row heights"
    Private Sub SetRowHeight(ByVal RowIndex As Integer)
        Dim selPrintListRow As PrintPartnerRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer = CalculateRowHeight(RowIndex)

        If intRowHeight > 0 Then selPrintListRow.RowHeight = intRowHeight Else selPrintListRow.RowHeight = NewCellHeight()
    End Sub

    Private Sub ResetRowHeights()
        For i = 0 To PrintList.Count - 1
            SetRowHeight(i)
        Next
    End Sub

    Private Function CalculateRowHeight(ByVal RowIndex As Integer) As Integer
        Dim selRow As PrintPartnerRow = Me.PrintList(RowIndex)
        Dim intRowHeight As Integer, intCellHeight As Integer
        Dim intColumnWidth As Integer

        If PageGraph IsNot Nothing Then
            If String.IsNullOrEmpty(selRow.PropertyName) = False Then
                If selRow.PropertyType > ObjectTypes.Organisation Then intColumnWidth = SecondaryPropertyRectangle.Width Else intColumnWidth = PropertyRectangle.Width
                intCellHeight = PageGraph.MeasureString(selRow.PropertyName, fntNormalBold, intColumnWidth - (CONST_HorizontalPadding * 2)).Height
                If intCellHeight > intRowHeight Then intRowHeight = intCellHeight
            End If

            If String.IsNullOrEmpty(selRow.PropertyValue) = False Then
                If selRow.PropertyType <> ObjectTypes.Organisation Then
                    intCellHeight = PageGraph.MeasureString(selRow.PropertyValue, fntNormal, TextRectangle.Width - (CONST_HorizontalPadding * 2)).Height
                Else
                    intCellHeight = PageGraph.MeasureString(selRow.PropertyValue, fntNormalBold, TextRectangle.Width - (CONST_HorizontalPadding * 2)).Height
                End If
                If intCellHeight > intRowHeight Then intRowHeight = intCellHeight
            End If

            intRowHeight += (CONST_VerticalPadding * 2)
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

        For Each selRow As PrintPartnerRow In PrintList
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
        Dim selRow As PrintPartnerRow = PrintList(RowIndex)
        Dim intMinHeight As Integer

        LastRowY = ContentTop

        'Print Headers
        PrintHeader()

        If PrintList.Count > 0 Then
            Do
                selRow = PrintList(RowIndex)
                If selRow Is Nothing Then Exit Do

                PrintPage_PrintPropertyName(selRow)
                PrintPropertyValue(selRow)

                LastRowY += selRow.RowHeight
                RowIndex += 1

                If RowIndex > PrintList.Count - 1 Then Exit Do
                If RowIndex > 0 And PrintList(RowIndex).PropertyName.StartsWith(LANG_Organisation) Then LastRowY += CONST_Spacing

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

    Private Sub PrintPage_PrintPropertyName(ByVal selRow As PrintPartnerRow)
        Dim rCell As Rectangle, rText As Rectangle

        rCell = New Rectangle(PropertyRectangle.Left, LastRowY, PropertyRectangle.Width, selRow.RowHeight)

        PrintBackground(rCell)

        If selRow.PropertyType > ObjectTypes.Organisation Then
            rText = New Rectangle(SecondaryPropertyRectangle.Left, LastRowY + CONST_VerticalPadding, SecondaryPropertyRectangle.Width - CONST_HorizontalPadding, selRow.RowHeight)
            PageGraph.DrawString(selRow.PropertyName, fntNormal, Brushes.Black, rText)
        Else
            rText = New Rectangle(PropertyRectangle.Left + CONST_HorizontalPadding, LastRowY + CONST_VerticalPadding, PropertyRectangle.Width - (CONST_HorizontalPadding * 2), selRow.RowHeight)
            PageGraph.DrawString(selRow.PropertyName, fntNormalBold, Brushes.Black, rText)
        End If

        If PrintBorders = True Then PageGraph.DrawRectangle(penBlack1, rCell)
    End Sub

    Private Sub PrintPropertyValue(ByVal selRow As PrintPartnerRow)
        Dim rCell As Rectangle, rText As Rectangle

        rCell = New Rectangle(TextRectangle.Left, LastRowY, TextRectangle.Width, selRow.RowHeight)
        rText = New Rectangle(rCell.X + CONST_HorizontalPadding, rCell.Y + CONST_VerticalPadding, rCell.Width - (CONST_HorizontalPadding * 2), selRow.RowHeight)

        PrintBackground(rCell)

        If selRow.Bold = False Then
            PageGraph.DrawString(selRow.PropertyValue, fntNormal, Brushes.Black, rText)
        Else
            PageGraph.DrawString(selRow.PropertyValue, fntNormalBold, Brushes.Black, rText)
        End If

        If PrintBorders = True Then PageGraph.DrawRectangle(penBlack1, rCell)
    End Sub

    Private Sub PrintBackground(ByVal rCell As Rectangle)
        If PageGraph IsNot Nothing And FillCells = True Then
            Dim brBackGround As Brush
            If Int(RowIndex / 2) = RowIndex / 2 Then
                brBackGround = Brushes.LightGray
            Else
                brBackGround = Brushes.Gainsboro
            End If
            PageGraph.FillRectangle(brBackGround, rCell)
        End If
    End Sub
#End Region
End Class


