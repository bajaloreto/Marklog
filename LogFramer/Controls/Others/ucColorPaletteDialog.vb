Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class ucColorPaletteDialog
    Private panel(40) As Panel
    Private color() As Color = {System.Drawing.Color.FromArgb(0, 0, 0), System.Drawing.Color.FromArgb(153, 51, 0), _
                                System.Drawing.Color.FromArgb(51, 51, 0), System.Drawing.Color.FromArgb(0, 51, 0), _
                                System.Drawing.Color.FromArgb(0, 51, 102), System.Drawing.Color.FromArgb(0, 0, 128), _
                                System.Drawing.Color.FromArgb(51, 51, 153), System.Drawing.Color.FromArgb(51, 51, 51), _
                                System.Drawing.Color.FromArgb(128, 0, 0), System.Drawing.Color.FromArgb(255, 102, 0), _
                                System.Drawing.Color.FromArgb(128, 128, 0), System.Drawing.Color.FromArgb(0, 128, 0), _
                                System.Drawing.Color.FromArgb(0, 128, 128), System.Drawing.Color.FromArgb(0, 0, 255), _
                                System.Drawing.Color.FromArgb(102, 102, 153), System.Drawing.Color.FromArgb(128, 128, 128), _
                                System.Drawing.Color.FromArgb(255, 0, 0), System.Drawing.Color.FromArgb(255, 153, 0), _
                                System.Drawing.Color.FromArgb(153, 204, 0), System.Drawing.Color.FromArgb(51, 153, 102), _
                                System.Drawing.Color.FromArgb(51, 204, 204), System.Drawing.Color.FromArgb(51, 102, 255), _
                                System.Drawing.Color.FromArgb(128, 0, 128), System.Drawing.Color.FromArgb(153, 153, 153), _
                                System.Drawing.Color.FromArgb(255, 0, 255), System.Drawing.Color.FromArgb(255, 204, 0), _
                                System.Drawing.Color.FromArgb(255, 255, 0), System.Drawing.Color.FromArgb(0, 255, 0), _
                                System.Drawing.Color.FromArgb(0, 255, 255), System.Drawing.Color.FromArgb(0, 204, 255), _
                                System.Drawing.Color.FromArgb(153, 51, 102), System.Drawing.Color.FromArgb(192, 192, 192), _
                                System.Drawing.Color.FromArgb(255, 153, 204), System.Drawing.Color.FromArgb(255, 204, 153), _
                                System.Drawing.Color.FromArgb(255, 255, 153), System.Drawing.Color.FromArgb(204, 255, 204), _
                                System.Drawing.Color.FromArgb(204, 255, 255), System.Drawing.Color.FromArgb(153, 204, 255), _
                                System.Drawing.Color.FromArgb(204, 153, 255), System.Drawing.Color.FromArgb(255, 255, 255)}

    Private markerColor() As Color = {System.Drawing.Color.FromArgb(255, 255, 255), _
                                      System.Drawing.Color.FromArgb(255, 0, 0), System.Drawing.Color.FromArgb(255, 153, 0), _
                                      System.Drawing.Color.FromArgb(153, 204, 0), System.Drawing.Color.FromArgb(51, 153, 102), _
                                      System.Drawing.Color.FromArgb(51, 204, 204), System.Drawing.Color.FromArgb(51, 102, 255), _
                                      System.Drawing.Color.FromArgb(255, 0, 255), System.Drawing.Color.FromArgb(255, 204, 0), _
                                      System.Drawing.Color.FromArgb(255, 255, 0), System.Drawing.Color.FromArgb(0, 255, 0), _
                                      System.Drawing.Color.FromArgb(0, 255, 255), System.Drawing.Color.FromArgb(0, 204, 255), _
                                      System.Drawing.Color.FromArgb(255, 153, 204), System.Drawing.Color.FromArgb(255, 204, 153), _
                                      System.Drawing.Color.FromArgb(255, 255, 153), System.Drawing.Color.FromArgb(204, 255, 204), _
                                      System.Drawing.Color.FromArgb(204, 255, 255), System.Drawing.Color.FromArgb(153, 204, 255)}

    Private bytPaletteType As Byte
    Private buttonMoreColors As New Button()
    Private buttonCancel As New Button()
    Private selectedColor
    Private colorTextValue
    Private colorMarkerValue

    Public Property Type() As Byte
        Get
            Return bytPaletteType
        End Get
        Set(ByVal value As Byte)
            bytPaletteType = value
        End Set
    End Property

    Public ReadOnly Property ColorChoice() As Color
        Get
            Return selectedColor
        End Get
    End Property

    Public Property ColorText() As Color
        Get
            Return colorTextValue
        End Get
        Set(ByVal value As Color)
            colorTextValue = value
        End Set
    End Property

    Public Property ColorMarker() As Color
        Get
            Return colorMarkerValue
        End Get
        Set(ByVal value As Color)
            colorMarkerValue = value
        End Set
    End Property

    Public Enum PaletteType As Byte
        normal_colors = 0
        marker_colors = 1
    End Enum

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
    End Sub

    Public Sub ColorPaletteDialog(ByVal x As Integer, ByVal y As Integer)

        If bytPaletteType = PaletteType.normal_colors Then
            Size = New Size(170, 160)
        Else
            Size = New Size(125, 145)
        End If

        FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        MinimizeBox = False
        MaximizeBox = False
        ControlBox = False
        ShowInTaskbar = False
        CenterToScreen()
        Location = New Point(x, y)
        BuildPalette(bytPaletteType)

        With buttonMoreColors
            .Text = LANG_MoreColours
            .FlatStyle = FlatStyle.Popup

            If bytPaletteType = PaletteType.normal_colors Then
                .Size = New Size(146, 22)
                .Location = New Point(5, 100)
            Else
                .Size = New Size(109, 22)
                .Location = New Point(5, 82)
            End If

            AddHandler buttonMoreColors.Click, AddressOf buttonMoreColors_Click
        End With

        Controls.Add(buttonMoreColors)

        With buttonCancel
            .Size = New Size(5, 5)
            .Location = New Point(-10, -10)
            .TabIndex = 0

            AddHandler buttonCancel.Click, AddressOf buttonCancel_Click
        End With

        Controls.Add(buttonCancel)

        Me.CancelButton = buttonCancel
    End Sub

    Sub BuildPalette(ByVal bytPaletteType As Byte)
        Dim pwidth As Integer = 13
        Dim pheight As Integer = 13
        Dim pdistance As Byte = 6
        Dim border As Byte = 5
        Dim PanelsPerRow As Byte, ppr As Byte
        Dim clrFill() As Color
        Dim x As Integer = border, y As Integer = border
        Dim toolTip As ToolTip = New ToolTip()
        Dim strToolTip() As String
        Dim i As Integer, max As Integer

        If bytPaletteType = PaletteType.normal_colors Then
            max = 39 '40-1
            ReDim strToolTip(max)
            strToolTip = LIST_ColorName
            ReDim clrFill(max)
            clrFill = color
            PanelsPerRow = 7 '8-1
        Else
            max = 18 '18-1
            ReDim strToolTip(max)
            strToolTip = LIST_MarkerColorName
            ReDim clrFill(max)
            clrFill = markerColor
            PanelsPerRow = 5 '6-1
        End If
        For i = 0 To max
            If bytPaletteType = PaletteType.marker_colors And i = 0 Then
                pwidth = 109
                ppr = 0
            Else
                pwidth = 13
                ppr = PanelsPerRow
            End If
            panel(i) = New Panel()
            panel(i).Height = pheight
            panel(i).Width = pwidth
            panel(i).Location = New Point(x, y)
            toolTip.SetToolTip(panel(i), strToolTip(i))
            Me.Controls.Add(panel(i))
            'If bytPaletteType = PaletteType.marker_colors And i = 0 Then
            '    x = pwidth + pdistance
            '    pwidth = 13

            'End If

            If (x < (ppr * (pwidth + pdistance))) Then
                x += pwidth + pdistance
            Else
                x = border
                y += pheight + pdistance
            End If
            panel(i).BackColor = clrFill(i)

            AddHandler panel(i).MouseEnter, AddressOf OnMouseEnterPanel
            AddHandler panel(i).MouseLeave, AddressOf OnMouseLeavePanel
            AddHandler panel(i).MouseDown, AddressOf OnMouseDownPanel
            AddHandler panel(i).MouseUp, AddressOf OnMouseUpPanel
            AddHandler panel(i).Paint, AddressOf OnPanelPaint
        Next
    End Sub

    Private Sub buttonMoreColors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim cdMoreColors As New ColorDialog
        cdMoreColors.FullOpen = True
        cdMoreColors.ShowDialog()
        selectedColor = cdMoreColors.Color
        DialogResult = Windows.Forms.DialogResult.OK
        cdMoreColors.Dispose()
        Close()
    End Sub

    Private Sub buttonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Close()
    End Sub

    Sub OnMouseEnterPanel(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DrawPanel(sender, 1)
    End Sub

    Sub OnMouseLeavePanel(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DrawPanel(sender, 0)
    End Sub

    Sub OnMouseDownPanel(ByVal sender As System.Object, ByVal e As System.EventArgs)
        DrawPanel(sender, 2)
    End Sub

    Sub OnMouseUpPanel(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim panel As Panel
        panel = sender
        selectedColor = panel.BackColor
        DialogResult = Windows.Forms.DialogResult.OK
        Close()
    End Sub

    Sub DrawPanel(ByVal sender As Object, ByVal state As Byte)
        Dim panel As Panel
        panel = sender
        Dim g As Graphics = panel.CreateGraphics()
        Dim graph As Graphics = Me.CreateGraphics()
        Dim penPanel As Pen, penShadow As Pen, penOutline As Pen
        Select Case state
            Case 1 'mouse over
                penPanel = New Pen(SystemColors.ControlDarkDark, 1)
                penShadow = New Pen(SystemColors.ControlDark, 3)
                penOutline = New Pen(SystemColors.ControlDarkDark, 1)
            Case 2 'clicked
                penPanel = New Pen(SystemColors.ControlDarkDark, 1)
                penShadow = New Pen(SystemColors.ControlDark, 3)
                penOutline = New Pen(SystemColors.ControlDarkDark, 1)
            Case Else 'neutral
                penPanel = New Pen(SystemColors.ControlDarkDark, 1)
                penShadow = New Pen(SystemColors.Control, 3)
                penOutline = New Pen(SystemColors.Control, 1)
        End Select
        Dim pPanel As Point = New Point(panel.Left, panel.Top)
        Dim intWidthPanel As Integer = panel.Width
        Dim intHeightPanel As Integer = panel.Height
        Dim pShadow As Point = New Point(panel.Left - 2, panel.Top - 2)
        Dim intWidthShadow As Integer = panel.Width + 3
        Dim intHeightShadow As Integer = panel.Height + 3
        Dim pOutline As Point = New Point(panel.Left - 3, panel.Top - 3)
        Dim intWidthOutline As Integer = panel.Width + 5
        Dim intHeightOutline As Integer = panel.Height + 5

        graph.DrawRectangle(penShadow, pShadow.X, pShadow.Y, intWidthShadow, intHeightShadow)
        graph.DrawRectangle(penOutline, pOutline.X, pOutline.Y, intWidthOutline, intHeightOutline)
        graph.DrawRectangle(penPanel, pPanel.X, pPanel.Y, intWidthPanel, intHeightPanel)
    End Sub

    Sub OnPanelPaint(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim panel As Panel
        panel = sender

        DrawPanel(sender, 0)
        If bytPaletteType = PaletteType.normal_colors And Not (colorTextValue Is Nothing) Then
            If panel.BackColor.ToArgb = colorTextValue.toArgb Then DrawPanel(panel, 2)
        End If
        If bytPaletteType = PaletteType.marker_colors And Not (colorMarkerValue Is Nothing) Then
            If panel.BackColor.ToArgb = colorMarkerValue.toArgb Then DrawPanel(panel, 2)
        End If
    End Sub
End Class