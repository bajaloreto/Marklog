Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.ComponentModel

Friend Enum ZoomMode
    ActualSize = 0
    Custom = 4
    FullPage = 1
    PageWidth = 2
    TwoPages = 3
End Enum

Friend Class CustomPrintPreview
    Inherits UserControl

    'Based on the CoolPrintView by Bernardo Castilho
    'Published on www.codeproject.com on 4 Oct 2013
    'http://www.codeproject.com/Articles/38758/An-Enhanced-PrintPreviewDialog

    Private brBackBrush As Brush
    Private boolCancel As Boolean
    Private objPrintDocument As PrintDocument
    Private _himm2pix As PointF = New PointF(-1.0!, -1.0!)
    Private objPageImageList As PageImageList = New PageImageList
    Private ptLast As Point
    Private boolRendering As Boolean
    Private intStartPage As Integer
    Private dblZoom As Double
    Private intZoomMode As ZoomMode
    Private Shadows Const MARGIN As Integer = 4

    ' Events
    Public Event PageCountChanged As EventHandler
    Public Event StartPageChanged As EventHandler
    Public Event ZoomModeChanged As EventHandler

    ' Methods
    Public Sub New()
        InitializeComponent()
        Me.BackColor = SystemColors.AppWorkspace
        Me.ZoomMode = ZoomMode.FullPage
        Me.StartPage = 0
        MyBase.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
    End Sub

    Private Sub _doc_EndPrint(ByVal sender As Object, ByVal e As PrintEventArgs)
        Me.SyncPageImages(True)
    End Sub

    Private Sub _doc_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Me.SyncPageImages(False)
        If Me.boolCancel Then
            e.Cancel = True
        End If
    End Sub

    Public Sub Cancel()
        Me.boolCancel = True
    End Sub

    Private Function GetImage(ByVal PageIndex As Integer) As Image
        If (PageIndex > -1) AndAlso (PageIndex < Me.PageCount) Then
            Return Me.objPageImageList.Item(PageIndex)
        Else
            Return Nothing
        End If
    End Function

    Private Function GetImageRectangle(ByVal img As Image) As Rectangle

        Dim sz As Size = Me.GetImageSizeInPixels(img)
        Dim rc As New Rectangle(0, 0, sz.Width, sz.Height)

        Dim rcCli As Rectangle = MyBase.ClientRectangle
        Dim zoomX As Double, zoomY As Double

        Select Case Me.intZoomMode
            Case ZoomMode.ActualSize
                Me.dblZoom = 1
            Case ZoomMode.FullPage
                zoomX = IIf((rc.Width > 0), (CDbl(rcCli.Width) / CDbl(rc.Width)), 0)
                zoomY = IIf((rc.Height > 0), (CDbl(rcCli.Height) / CDbl(rc.Height)), 0)
                Me.dblZoom = Math.Min(zoomX, zoomY)
            Case ZoomMode.PageWidth
                Me.dblZoom = IIf((rc.Width > 0), (CDbl(rcCli.Width) / CDbl(rc.Width)), 0)

            Case ZoomMode.TwoPages
                rc.Width = (rc.Width * 2)
                zoomX = IIf((rc.Width > 0), (CDbl(rcCli.Width) / CDbl(rc.Width)), 0)
                zoomY = IIf((rc.Height > 0), (CDbl(rcCli.Height) / CDbl(rc.Height)), 0)
                Me.dblZoom = Math.Min(zoomX, zoomY)
            Case Else
                '
        End Select
        rc = GetImageRectangle_Calculate(rc, rcCli)
        
        Return rc
    End Function

    Private Function GetImageRectangle_Calculate(ByVal rc As Rectangle, ByVal rcCli As Rectangle) As Rectangle
        rc.Width = CInt((rc.Width * Me.dblZoom))
        rc.Height = CInt((rc.Height * Me.dblZoom))
        Dim dx As Integer = ((rcCli.Width - rc.Width) / 2)
        If (dx > 0) Then
            rc.X = (rc.X + dx)
        End If
        Dim dy As Integer = ((rcCli.Height - rc.Height) / 2)
        If (dy > 0) Then
            rc.Y = (rc.Y + dy)
        End If
        rc.Inflate(-4, -4)
        If (Me.intZoomMode = ZoomMode.TwoPages) Then
            rc.Inflate(-2, 0)
        End If
        Return rc
    End Function

    Private Function GetImageSizeInPixels(ByVal img As Image) As Size
        Dim szf As SizeF = img.PhysicalDimension
        If TypeOf img Is Metafile Then
            If (Me._himm2pix.X < 0.0!) Then
                Using g As Graphics = MyBase.CreateGraphics
                    Me._himm2pix.X = (g.DpiX / 2540.0!)
                    Me._himm2pix.Y = (g.DpiY / 2540.0!)
                End Using
            End If
            szf.Width = (szf.Width * Me._himm2pix.X)
            szf.Height = (szf.Height * Me._himm2pix.Y)
        End If
        Return Size.Truncate(szf)
    End Function

    Protected Overrides Function IsInputKey(ByVal keyData As Keys) As Boolean
        Select Case keyData
            Case Keys.Prior, Keys.Next, Keys.End, Keys.Home, Keys.Left, Keys.Up, Keys.Right, Keys.Down
                Return True
        End Select
        Return MyBase.IsInputKey(keyData)
    End Function

    Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.Handled Then Return

        Select Case e.KeyCode

            ' arrow keys scroll or browse, depending on ZoomMode
            Case Keys.Left, Keys.Up, Keys.Right, Keys.Down

                ' browse
                If ZoomMode = ZoomMode.FullPage OrElse ZoomMode = ZoomMode.TwoPages Then
                    Select Case e.KeyCode
                        Case Keys.Left, Keys.Up
                            StartPage -= 1

                        Case Keys.Right, Keys.Down
                            StartPage += 1
                    End Select
                End If

                ' scroll
                Dim pt As Point = AutoScrollPosition
                Select Case e.KeyCode
                    Case Keys.Left
                        pt.X += 20
                    Case Keys.Right
                        pt.X -= 20
                    Case Keys.Up
                        pt.Y += 20
                    Case Keys.Down
                        pt.Y -= 20
                End Select
                AutoScrollPosition = New Point(-pt.X, -pt.Y)

                ' page up/down browse pages
            Case Keys.PageUp
                StartPage -= 1
            Case Keys.PageDown
                StartPage += 1

                ' home/end 
            Case Keys.Home
                AutoScrollPosition = Point.Empty
                StartPage = 0
            Case Keys.End
                AutoScrollPosition = Point.Empty
                StartPage = PageCount - 1

            Case Else
                Return
        End Select

        ' if we got here, the event was handled
        e.Handled = True
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If ((e.Button = MouseButtons.Left) AndAlso (MyBase.AutoScrollMinSize <> Size.Empty)) Then
            Me.Cursor = Cursors.NoMove2D
            Me.ptLast = New Point(e.X, e.Y)
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        If (Me.Cursor Is Cursors.NoMove2D) Then
            Dim dx As Integer = (e.X - Me.ptLast.X)
            Dim dy As Integer = (e.Y - Me.ptLast.Y)
            If ((dx <> 0) OrElse (dy <> 0)) Then
                Dim pt As Point = MyBase.AutoScrollPosition
                MyBase.AutoScrollPosition = New Point(-(pt.X + dx), -(pt.Y + dy))
                Me.ptLast = New Point(e.X, e.Y)
            End If
        End If
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If ((e.Button = MouseButtons.Left) AndAlso (Me.Cursor Is Cursors.NoMove2D)) Then
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Protected Sub OnPageCountChanged(ByVal e As EventArgs)
        RaiseEvent PageCountChanged(Me, e)
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim img As Image = Me.GetImage(Me.StartPage)
        If (Not img Is Nothing) Then
            Dim rc As Rectangle = Me.GetImageRectangle(img)
            If ((rc.Width > 2) AndAlso (rc.Height > 2)) Then
                rc.Offset(MyBase.AutoScrollPosition)
                If (Me.intZoomMode <> ZoomMode.TwoPages) Then
                    Me.RenderPage(e.Graphics, img, rc)
                Else
                    rc.Width = ((rc.Width - 4) / 2)
                    Me.RenderPage(e.Graphics, img, rc)
                    img = Me.GetImage((Me.StartPage + 1))
                    If (Not img Is Nothing) Then
                        rc = Me.GetImageRectangle(img)
                        rc.Width = ((rc.Width - 4) / 2)
                        rc.Offset((rc.Width + 4), 0)
                        Me.RenderPage(e.Graphics, img, rc)
                    End If
                End If
            End If
        End If
        e.Graphics.FillRectangle(Me.brBackBrush, MyBase.ClientRectangle)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)

    End Sub

    Protected Overrides Sub OnSizeChanged(ByVal e As EventArgs)
        Me.UpdateScrollBars()
        MyBase.OnSizeChanged(e)
    End Sub

    Protected Sub OnStartPageChanged(ByVal e As EventArgs)
        RaiseEvent StartPageChanged(Me, e)
    End Sub

    Protected Sub OnZoomModeChanged(ByVal e As EventArgs)
        RaiseEvent ZoomModeChanged(Me, e)
    End Sub

    Public Sub Print()
        Dim ps As PrinterSettings = Me.objPrintDocument.PrinterSettings
        Dim first As Integer = (ps.MinimumPage - 1)
        Dim last As Integer = (ps.MaximumPage - 1)

        Select Case ps.PrintRange
            Case PrintRange.AllPages
                Me.Document.Print()
                Return
            Case PrintRange.Selection
                first = Me.StartPage
                last = Me.StartPage
                If (Me.ZoomMode = ZoomMode.TwoPages) Then
                    last = Math.Min(CInt((first + 1)), CInt((Me.PageCount - 1)))
                End If
                Exit Select
            Case PrintRange.SomePages
                first = (ps.FromPage - 1)
                last = (ps.ToPage - 1)
                Exit Select
            Case PrintRange.CurrentPage
                first = Me.StartPage
                last = Me.StartPage
                Exit Select
        End Select

        Dim dp = New DocumentPrinter(Me, first, last)
        dp.Print()
    End Sub

    Public Sub RefreshPreview()
        If (Not Me.objPrintDocument Is Nothing) Then
            Me.objPageImageList.Clear()
            Dim savePC As PrintController = Me.objPrintDocument.PrintController
            Try
                Me.boolCancel = False
                Me.boolRendering = True
                Me.objPrintDocument.PrintController = New PreviewPrintController
                AddHandler Me.objPrintDocument.PrintPage, New PrintPageEventHandler(AddressOf Me._doc_PrintPage)
                AddHandler Me.objPrintDocument.EndPrint, New PrintEventHandler(AddressOf Me._doc_EndPrint)
                Me.objPrintDocument.Print()
                'Catch ex As Exception
                '    MsgBox(ex.InnerException.ToString & vbLf & ex.Message)

            Finally
                Me.boolCancel = False
                Me.boolRendering = False
                RemoveHandler Me.objPrintDocument.PrintPage, New PrintPageEventHandler(AddressOf Me._doc_PrintPage)
                RemoveHandler Me.objPrintDocument.EndPrint, New PrintEventHandler(AddressOf Me._doc_EndPrint)
                Me.objPrintDocument.PrintController = savePC
            End Try
        End If
        Me.OnPageCountChanged(EventArgs.Empty)
        Me.UpdatePreview()
        Me.UpdateScrollBars()
    End Sub

    Private Sub RenderPage(ByVal g As Graphics, ByVal img As Image, ByVal rc As Rectangle)
        rc.Offset(1, 1)
        g.DrawRectangle(Pens.Black, rc)
        rc.Offset(-1, -1)
        g.FillRectangle(Brushes.White, rc)
        g.DrawImage(img, rc)
        g.DrawRectangle(Pens.Black, rc)
        rc.Width += 1
        rc.Height += 1
        g.ExcludeClip(rc)
        rc.Offset(1, 1)
        g.ExcludeClip(rc)
    End Sub

    Private Sub SyncPageImages(ByVal lastPageReady As Boolean)
        Dim pvController As PreviewPrintController = TryCast(Me.objPrintDocument.PrintController, PreviewPrintController)

        If pvController IsNot Nothing Then
            Dim pageInfo As PreviewPageInfo() = pvController.GetPreviewPageInfo
            Dim count As Integer = IIf(lastPageReady, pageInfo.Length, (pageInfo.Length - 1))
            Dim i As Integer
            For i = Me.objPageImageList.Count To count - 1
                Dim img As Image = pageInfo(i).Image
                Me.objPageImageList.Add(img)
                Me.OnPageCountChanged(EventArgs.Empty)
                If (Me.StartPage < 0) Then
                    Me.StartPage = 0
                End If
                If ((i = Me.StartPage) OrElse (i = (Me.StartPage + 1))) Then
                    Me.Refresh()
                End If
                Application.DoEvents()
            Next i
        End If
    End Sub

    Private Sub UpdatePreview()
        If (Me.intStartPage < 0) Then
            Me.intStartPage = 0
        End If
        If (Me.intStartPage > (Me.PageCount - 1)) Then
            Me.intStartPage = (Me.PageCount - 1)
        End If
        MyBase.Invalidate()
    End Sub

    Private Sub UpdateScrollBars()
        Dim rc As Rectangle = Rectangle.Empty
        Dim img As Image = Me.GetImage(Me.StartPage)
        If (Not img Is Nothing) Then
            rc = Me.GetImageRectangle(img)
        End If
        Dim scrollSize As New Size(0, 0)
        Select Case Me.intZoomMode
            Case ZoomMode.ActualSize, ZoomMode.Custom
                scrollSize = New Size((rc.Width + 8), (rc.Height + 8))
                Exit Select
            Case ZoomMode.PageWidth
                scrollSize = New Size(0, (rc.Height + 8))
                Exit Select
        End Select
        If (scrollSize <> MyBase.AutoScrollMinSize) Then
            MyBase.AutoScrollMinSize = scrollSize
        End If
        Me.UpdatePreview()
    End Sub


    ' Properties
    <DefaultValue(GetType(Color), "AppWorkspace")> _
    Public Overrides Property BackColor() As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(ByVal value As Color)
            MyBase.BackColor = value
            Me.brBackBrush = New SolidBrush(value)
        End Set
    End Property

    Public Property Document() As PrintDocument
        Get
            Return Me.objPrintDocument
        End Get
        Set(ByVal value As PrintDocument)
            If (Not value Is Me.objPrintDocument) Then
                Me.objPrintDocument = value
                Me.RefreshPreview()
            End If
        End Set
    End Property

    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden), Browsable(False)> _
    Public ReadOnly Property IsRendering() As Boolean
        Get
            Return Me.boolRendering
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public ReadOnly Property PageCount() As Integer
        Get
            Return Me.objPageImageList.Count
        End Get
    End Property

    <Browsable(False)> _
    Public ReadOnly Property PageImages() As PageImageList
        Get
            Return Me.objPageImageList
        End Get
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Property StartPage() As Integer
        Get
            Return Me.intStartPage
        End Get
        Set(ByVal value As Integer)
            If (value > (Me.PageCount - 1)) Then
                value = (Me.PageCount - 1)
            End If
            If (value < 0) Then
                value = 0
            End If
            If (value <> Me.intStartPage) Then
                Me.intStartPage = value
                Me.UpdateScrollBars()
                Me.OnStartPageChanged(EventArgs.Empty)
            End If
        End Set
    End Property

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Property Zoom() As Double
        Get
            Return Me.dblZoom
        End Get
        Set(ByVal value As Double)
            If ((value <> Me.dblZoom) OrElse (Me.ZoomMode <> ZoomMode.Custom)) Then
                Me.ZoomMode = ZoomMode.Custom
                Me.dblZoom = value
                Me.UpdateScrollBars()
                Me.OnZoomModeChanged(EventArgs.Empty)
            End If
        End Set
    End Property

    <DefaultValue(1)> _
    Public Property ZoomMode() As ZoomMode
        Get
            Return Me.intZoomMode
        End Get
        Set(ByVal value As ZoomMode)
            If (value <> Me.intZoomMode) Then
                Me.intZoomMode = value
                Me.UpdateScrollBars()
                Me.OnZoomModeChanged(EventArgs.Empty)
            End If
        End Set
    End Property
    

    ' Nested Types
    Friend Class DocumentPrinter
        Inherits PrintDocument
        ' Methods
        Public Sub New(ByVal preview As CustomPrintPreview, ByVal first As Integer, ByVal last As Integer)
            Me._first = first
            Me._last = last
            Me._imgList = preview.PageImages
            MyBase.DefaultPageSettings = preview.Document.DefaultPageSettings
            MyBase.PrinterSettings = preview.Document.PrinterSettings
        End Sub

        Protected Overrides Sub OnBeginPrint(ByVal e As PrintEventArgs)
            Me._index = Me._first
        End Sub

        Protected Overrides Sub OnPrintPage(ByVal e As PrintPageEventArgs)
            e.Graphics.PageUnit = GraphicsUnit.Display
            e.Graphics.DrawImage(Me._imgList.Item(Me._index), e.PageBounds)
            Me._index = Me._index + 1
            e.HasMorePages = (Me._index <= Me._last)
        End Sub


        ' Fields
        Private _first As Integer
        Private _imgList As PageImageList
        Private _index As Integer
        Private _last As Integer
    End Class
End Class
