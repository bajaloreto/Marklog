Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D

Public Class RichTextManager
    Inherits RichTextBox

    Private Const WM_USER As Int32 = &H400&
    Private Const EM_FORMATRANGE As Int32 = WM_USER + 57
    Private Const inch As Single = 14.4
    Private strClippedRichTextTop, strClippedRichTextBottom As String
    Private strClippedTextTop, strClippedTextBottom As String

    Public Enum TextCases
        Normal = 0
        Uppercase = 1
        Lowercase = 2
        Sentence = 3
        FirstLetter = 4
    End Enum

    Public Sub New()
        ScrollBars = RichTextBoxScrollBars.None
        Multiline = True
        WordWrap = True

        'BackColor = Color.Blue
    End Sub

#Region "Properties"
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure STRUCT_RECT
        Public left As Int32
        Public top As Int32
        Public right As Int32
        Public bottom As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure STRUCT_CHARRANGE
        Public cpMin As Int32
        Public cpMax As Int32
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure STRUCT_FORMATRANGE
        Public hdc As IntPtr
        Public hdcTarget As IntPtr
        Public rc As STRUCT_RECT
        Public rcPage As STRUCT_RECT
        Public chrg As STRUCT_CHARRANGE
    End Structure

    <DllImport("user32.dll")> _
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, _
                                    ByVal msg As Int32, _
                                    ByVal wParam As Int32, _
                                    ByVal lParam As IntPtr) As Int32
    End Function

    Public Property ClippedRichTextTop As String
        Get
            Return strClippedRichTextTop
        End Get
        Set(ByVal value As String)
            strClippedRichTextTop = value
        End Set
    End Property

    Public Property ClippedRichTextBottom As String
        Get
            Return strClippedRichTextBottom
        End Get
        Set(ByVal value As String)
            strClippedRichTextBottom = value
        End Set
    End Property

    Public Property ClippedTextTop As String
        Get
            Return strClippedTextTop
        End Get
        Set(ByVal value As String)
            strClippedTextTop = value
        End Set
    End Property

    Public Property ClippedTextBottom As String
        Get
            Return strClippedTextBottom
        End Get
        Set(ByVal value As String)
            strClippedTextBottom = value
        End Set
    End Property
#End Region

#Region "General methods"
    Private Sub SetGraphicsQuality(ByRef graph As Graphics)
        With graph
            .CompositingQuality = CompositingQuality.HighQuality
            .InterpolationMode = InterpolationMode.HighQualityBicubic
            .PixelOffsetMode = PixelOffsetMode.HighQuality
            .SmoothingMode = SmoothingMode.HighQuality

            .TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        End With
    End Sub

    Private Sub SetColor(ByVal colForeGround As Color, ByVal colBackGround As Color, ByVal boolSelected As Boolean)
        If colBackGround = Nothing Then
            colBackGround = Color.White
        End If

        If boolSelected = True Then
            colBackGround = SystemColors.Highlight
            colForeGround = SystemColors.HighlightText
        End If

        Me.BackColor = colBackGround

        If colForeGround <> Nothing Then
            SelectAll()
            SelectionColor = colForeGround
            Me.Select(Me.TextLength - 1, 0)
        End If
    End Sub

    Private Sub SetIndent(ByVal intIndent As Integer)
        If intIndent > 0 Then
            intIndent *= CONST_StructIndentMultiplier
            SelectAll()
            SelectionIndent = intIndent
            Me.Select(Me.TextLength - 1, 0)
        End If
    End Sub

    Private Sub SetAlignement(ByVal objAlignment As Windows.Forms.HorizontalAlignment)
        SelectAll()
        SelectionAlignment = objAlignment
        Me.Select(Me.TextLength - 1, 0)
    End Sub
#End Region

#Region "Rich text"
    Public Function RichTextToBitmap() As Bitmap
        If Me.Width = 0 Or Me.Height = 0 Then Return Nothing

        Dim bmp As New Bitmap(Me.Width, Me.Height, Imaging.PixelFormat.Format32bppPArgb)

        'bmp.SetResolution(192, 192)

        Using graph As Graphics = Graphics.FromImage(bmp)
            SetGraphicsQuality(graph)

            Dim hDC As IntPtr = graph.GetHdc
            Dim fmtRange As STRUCT_FORMATRANGE
            Dim rect As STRUCT_RECT
            Dim fromAPI As Integer

            rect.top = 0
            rect.left = 0
            rect.bottom = Int(bmp.Height + (bmp.Height * (bmp.VerticalResolution / 100)) * inch)
            rect.right = Int(bmp.Width + (bmp.Width * (bmp.HorizontalResolution / 100)) * inch)

            fmtRange.chrg.cpMin = 0
            fmtRange.chrg.cpMax = -1
            fmtRange.hdc = hDC
            fmtRange.hdcTarget = hDC
            fmtRange.rc = rect
            fmtRange.rcPage = rect

            Dim wParam As Integer = 1

            Dim lParam As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(fmtRange))
            Marshal.StructureToPtr(fmtRange, lParam, False)

            fromAPI = SendMessage(Me.Handle, EM_FORMATRANGE, wParam, lParam)

            Marshal.FreeCoTaskMem(lParam)

            fromAPI = SendMessage(Me.Handle, EM_FORMATRANGE, wParam, New IntPtr(0))

            graph.ReleaseHdc(hDC)

        End Using

        Return bmp

    End Function

    Public Function RichTextToBitmap(ByVal intColumnWidth As Integer, ByVal strRtf As String) As Bitmap
        Me.Text = String.Empty

        Me.MaximumSize = New Size(intColumnWidth, 0)
        Me.Width = intColumnWidth

        Me.Rtf = strRtf

        Return Me.RichTextToBitmap()
    End Function

    Public Function RichTextWithPaddingToBitmap(ByVal intColumnWidth As Integer, ByVal strRtf As String, ByVal boolSelected As Boolean, _
                                                ByVal intIndent As Integer, ByVal objAlignment As Windows.Forms.HorizontalAlignment, _
                                                Optional ByVal colForeGround As Color = Nothing, Optional ByVal colBackGround As Color = Nothing) As Bitmap
        Me.Text = String.Empty
        Me.MaximumSize = New Size(intColumnWidth - CONST_HorizontalPadding, 0)
        Me.Width = intColumnWidth - CONST_HorizontalPadding
        Me.Rtf = strRtf

        SetColor(colForeGround, colBackGround, boolSelected)
        SetIndent(intIndent)
        SetAlignement(objAlignment)

        Return RichTextWithPaddingToBitmap_Draw()
    End Function

    Public Function RichTextWithPaddingToBitmap(ByVal intColumnWidth As Integer, ByVal strRtf As String, ByVal boolSelected As Boolean, Optional ByVal intIndent As Integer = 0) As Bitmap
        Me.Text = String.Empty
        Me.MaximumSize = New Size(intColumnWidth - CONST_HorizontalPadding, 0)
        Me.Width = intColumnWidth - CONST_HorizontalPadding
        Me.Rtf = strRtf

        SetColor(Nothing, Nothing, boolSelected)
        If intIndent > 0 Then SetIndent(intIndent)

        Return RichTextWithPaddingToBitmap_Draw()
    End Function

    Private Function RichTextWithPaddingToBitmap_Draw() As Bitmap
        Dim bmpText As Bitmap = Me.RichTextToBitmap()

        If bmpText IsNot Nothing Then
            Dim rText As New Rectangle(0, 0, bmpText.Width, bmpText.Height)
            Dim rPadded As New Rectangle(0, 0, bmpText.Width + CONST_HorizontalPadding, bmpText.Height + CONST_VerticalPadding)
            Dim bmpPadded As New Bitmap(rPadded.Width, rPadded.Height)
            Dim gPadded As Graphics = Graphics.FromImage(bmpPadded)

            gPadded.DrawImage(bmpText, CONST_Padding, CONST_Padding, rText, GraphicsUnit.Pixel)
            gPadded.Dispose()

            Return bmpPadded
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region "Text"
    Public Function TextToBitmap(ByVal intColumnWidth As Integer, ByVal strText As String, Optional ByVal font As Font = Nothing) As Bitmap
        Me.Text = String.Empty

        If font Is Nothing Then
            font = CurrentLogFrame.DetailsFont
        End If

        Me.Font = font

        Me.MaximumSize = New Size(intColumnWidth, 0)
        Me.Width = intColumnWidth
        Me.Text = strText

        Return Me.RichTextToBitmap()
    End Function

    Public Function TextWithPaddingToBitmap(ByVal intColumnWidth As Integer, ByVal strText As String, Optional ByVal font As Font = Nothing) As Bitmap
        Me.Text = String.Empty
        Me.MaximumSize = New Size(intColumnWidth - CONST_HorizontalPadding, 0)
        Me.Width = intColumnWidth - CONST_HorizontalPadding

        If font Is Nothing Then
            font = CurrentLogFrame.DetailsFont
        End If
        Me.Font = font
        Me.Text = strText

        Dim bmpText As Bitmap = Me.RichTextToBitmap()
        Dim rText As New Rectangle(0, 0, bmpText.Width, bmpText.Height)
        Dim rPadded As New Rectangle(0, 0, bmpText.Width + CONST_HorizontalPadding, bmpText.Height + CONST_VerticalPadding)
        Dim bmpPadded As New Bitmap(rPadded.Width, rPadded.Height)
        Dim gPadded As Graphics = Graphics.FromImage(bmpPadded)

        gPadded.DrawImage(bmpText, CONST_Padding, CONST_Padding, rText, GraphicsUnit.Pixel)
        gPadded.Dispose()

        Return bmpPadded
    End Function

    Public Function EmptyTextWithPaddingToBitmap(ByVal intColumnWidth As Integer, ByVal strItemName As String, ByVal strSort As String, ByVal boolSelected As Boolean)
        Dim strText As String

        Me.MaximumSize = New Size(intColumnWidth - CONST_HorizontalPadding, 0)
        Me.Width = intColumnWidth - CONST_HorizontalPadding

        strText = String.Format("{{{0} {1}}}", strItemName.ToLower, strSort)
        Me.Text = strText
        SelectAll()
        SelectionFont = CurrentText.Font
        SelectionColor = Color.DarkBlue
        Me.Select(Me.TextLength - 1, 0)

        Dim bmpEmpty As Bitmap = Me.RichTextWithPaddingToBitmap(intColumnWidth, Me.Rtf, boolSelected)

        Return bmpEmpty
    End Function
#End Region

#Region "Change text"
    Public Function ChangeFont(ByVal strRtf As String) As String
        If String.IsNullOrEmpty(strRtf) = False Then
            Me.Rtf = strRtf

            With CurrentText
                Dim strFontFamily As String = .Font.FontFamily.Name
                Dim intFontSize As Integer = .Font.Size
                Dim boolBold As Boolean = .Font.Bold, boolItalic As Boolean = .Font.Italic
                Dim boolUnderline As Boolean = .Font.Underline, boolStrike As Boolean = .Font.Strikeout
                Dim fstNewFontStyle As FontStyle

                fstNewFontStyle = FontStyle.Regular
                Me.SelectAll()
                If strFontFamily = String.Empty Then strFontFamily = Me.SelectionFont.Name
                If intFontSize > 0 Then
                    If Me.SelectionCharOffset <> 0 Then intFontSize *= 0.6
                Else
                    intFontSize = Me.SelectionFont.SizeInPoints
                End If

                If boolBold = True Then fstNewFontStyle = FontStyle.Bold
                If boolItalic = True Then fstNewFontStyle += FontStyle.Italic
                If boolUnderline = True Then fstNewFontStyle += FontStyle.Underline
                If boolStrike = True Then fstNewFontStyle += FontStyle.Strikeout
                Try
                    Me.SelectionFont = New Font(strFontFamily, intFontSize, fstNewFontStyle, GraphicsUnit.Point)
                Catch ex2 As ArgumentException
                    Try
                        Me.SelectionFont = New Font(strFontFamily, intFontSize, Me.SelectionFont.Style, GraphicsUnit.Point)
                    Catch ex3 As ArgumentException
                        MsgBox(ERR_FontStyleNotSupported, MsgBoxStyle.OkOnly)
                        Return strRtf
                    End Try
                End Try

                Return Me.Rtf
            End With
        Else
            Return Nothing
        End If
    End Function

    Public Function ChangeFontCase(ByVal strRtf As String, ByVal intCase As Integer) As String
        If String.IsNullOrEmpty(strRtf) = False Then
            Select Case intCase
                Case TextCases.Uppercase
                    Me.SelectedText = Me.SelectedText.ToUpper
                Case TextCases.Lowercase
                    Me.SelectedText = Me.SelectedText.ToLower
                Case TextCases.FirstLetter
                    Me.SelectedText = StrConv(Me.SelectedText, VbStrConv.ProperCase)
                Case TextCases.Sentence
                    Dim sentenceStart As Boolean = True
                    Dim strSentence() As Char = CStr(Me.SelectedText).ToLower
                    For i = 0 To Len(strSentence) - 1
                        If Char.IsLetterOrDigit(strSentence(i)) Then
                            If sentenceStart Then
                                strSentence(i) = CStr(strSentence(i)).ToUpper
                                sentenceStart = False
                            End If
                        ElseIf InStr(".!?" & vbCrLf, strSentence(i)) Then
                            sentenceStart = True
                        End If
                    Next
                    Me.SelectedText = strSentence
            End Select
            Return Me.Rtf
        Else
            Return Nothing
        End If
    End Function

    Public Function ChangeFontColor(ByVal strRtf As String, ByVal selColor As Color) As String
        If String.IsNullOrEmpty(strRtf) = False Then
            Me.Rtf = strRtf
            Me.SelectAll()
            Me.SelectionColor = selColor
            Return Me.Rtf
        Else
            Return Nothing
        End If
    End Function

    Public Function ChangeMarkerColor(ByVal strRtf As String, ByVal selColor As Color) As String
        If String.IsNullOrEmpty(strRtf) = False Then
            Me.Rtf = strRtf
            Me.SelectAll()
            Me.SelectionBackColor = selColor
            Return Me.Rtf
        Else
            Return Nothing
        End If
    End Function

    Public Function ChangeAlignment(ByVal strRtf As String, ByVal intAlignment As Integer) As String
        If String.IsNullOrEmpty(strRtf) = False Then
            Me.Rtf = strRtf
            Me.SelectAll()
            Me.SelectionAlignment = intAlignment
            Return Me.Rtf
        Else
            Return Nothing
        End If
    End Function

    Public Function ChangeLeftIndent(ByVal strRtf As String, ByVal intIncrement As Integer) As String
        If String.IsNullOrEmpty(strRtf) = False Then
            Me.Rtf = strRtf
            Me.SelectAll()
            Me.SelectionIndent = intIncrement
            Return Me.Rtf
        Else
            Return Nothing
        End If
    End Function
#End Region

#Region "Analyse text"
    Public Function GetFirstLineSpacing(ByVal intColumnWidth As Integer, ByVal strRtf As String) As Single
        Dim sngLineSpacing As Single
        Dim ptLine As New Point
        Dim lstLineIndices As New ArrayList
        Dim intSplit As Integer = -1
        Dim intIndex As Integer

        Me.Text = String.Empty
        Me.MaximumSize = New Size(intColumnWidth - CONST_HorizontalPadding, 0)
        Me.Width = intColumnWidth - CONST_HorizontalPadding
        Me.Rtf = strRtf

        For intIndex = 0 To TextLength - 1
            Dim ptLocation As Point = GetPositionFromCharIndex(intIndex)

            If ptLocation.Y > ptLine.Y Then
                ptLine.Y = ptLocation.Y

                If ptLine.Y > 0 Then
                    intIndex -= 1
                    Exit For
                End If
            End If
        Next

        If intIndex < 0 Then intIndex = TextLength - 1

        For i = 0 To intIndex
            Me.Select(i, 0)
            If Me.SelectionFont.GetHeight > sngLineSpacing Then _
                sngLineSpacing = Me.SelectionFont.GetHeight
        Next

        Return sngLineSpacing
    End Function
#End Region

#Region "Print clipped text"
    Public Function PrintClippedText(ByVal intColumnWidth As Integer, ByVal strText As String, ByVal intContentBottom As Integer) As Bitmap
        Me.Text = String.Empty
        Me.MaximumSize = New Size(intColumnWidth, 0)
        Me.Width = intColumnWidth
        Me.Text = strText

        Return PrintClippedText_CreateBitmap(intColumnWidth, intContentBottom, False)
    End Function

    Public Function PrintClippedRichText(ByVal intColumnWidth As Integer, ByVal strRtf As String, ByVal intContentBottom As Integer) As Bitmap
        Me.Text = String.Empty
        Me.MaximumSize = New Size(intColumnWidth, 0)
        Me.Width = intColumnWidth
        Me.Rtf = strRtf

        Return PrintClippedText_CreateBitmap(intColumnWidth, intContentBottom, True)
    End Function

    Private Function PrintClippedText_CreateBitmap(ByVal intColumnWidth As Integer, ByVal intContentBottom As Integer, ByVal boolRichText As Boolean) As Bitmap
        Dim bmClip As Bitmap = Nothing
        Dim ptLine As New Point
        Dim lstLineIndices As New ArrayList
        Dim intSplit As Integer = -1

        For intIndex = 0 To TextLength - 1
            Dim ptLocation As Point = GetPositionFromCharIndex(intIndex)

            If ptLocation.Y > ptLine.Y Then
                ptLine.Y = ptLocation.Y

                lstLineIndices.Add(intIndex)
                If ptLocation.Y > intContentBottom Then
                    intSplit = lstLineIndices(lstLineIndices.Count - 2)
                    bmClip = PrintClippedText_Split(intColumnWidth, intSplit, boolRichText)

                    Exit For
                End If
            End If
        Next

        If intSplit < 0 And lstLineIndices.Count > 0 Then
            intSplit = lstLineIndices(lstLineIndices.Count - 1)

            bmClip = PrintClippedText_Split(intColumnWidth, intSplit, boolRichText)
        End If

        Return bmClip
    End Function

    Private Function PrintClippedText_Split(ByVal intColumnWidth As Integer, ByVal intSplit As Integer, ByVal boolRichText As Boolean) As Bitmap
        Me.Select(0, intSplit)
        If boolRichText = True Then
            strClippedTextTop = SelectedRtf
        Else
            strClippedTextTop = SelectedText
        End If


        Me.Select(intSplit, TextLength - intSplit)
        If boolRichText = True Then
            strClippedTextBottom = SelectedRtf
        Else
            strClippedTextBottom = SelectedText
        End If

        SelectedText = String.Empty

        Dim bmClip As Bitmap = RichTextWithPaddingToBitmap(intColumnWidth, Me.Rtf, False)
        'Dim strFileName As String = String.Format("D:/{0} Splittest.bmp", Now.ToFileTime)

        'bmClip.Save(strFileName)
        Return bmClip
    End Function

    Private Sub RichTextManager_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs) Handles Me.ContentsResized
        Me.Height = e.NewRectangle.Height
    End Sub
#End Region

#Region "Not used"
    'Public Function RtbToBitmap(ByVal rtb As Control) As Bitmap
    '    Return RtbToBitmap(rtb, rtb.ClientRectangle.Width, rtb.ClientRectangle.Height)
    'End Function

    'Public Function RtbToBitmap(ByVal rtb As Control, ByVal width As Integer, ByVal height As Integer) As Bitmap

    '    Dim bmp As New Bitmap(width, height)

    '    Using gr As Graphics = Graphics.FromImage(bmp)
    '        Dim hDC As IntPtr = gr.GetHdc
    '        Dim fmtRange As STRUCT_FORMATRANGE
    '        Dim rect As STRUCT_RECT
    '        Dim fromAPI As Integer

    '        rect.bottom = Int(bmp.Height + (bmp.Height * (bmp.HorizontalResolution / 100)) * inch)
    '        rect.right = Int(bmp.Width + (bmp.Width * (bmp.VerticalResolution / 100)) * inch)

    '        fmtRange.chrg.cpMin = 0
    '        fmtRange.chrg.cpMax = -1
    '        fmtRange.hdc = hDC
    '        fmtRange.hdcTarget = hDC
    '        fmtRange.rc = rect
    '        fmtRange.rcPage = rect

    '        Dim wParam As Integer = 1

    '        Dim lParam As IntPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(fmtRange))
    '        Marshal.StructureToPtr(fmtRange, lParam, False)

    '        fromAPI = SendMessage(rtb.Handle, EM_FORMATRANGE, wParam, lParam)

    '        Marshal.FreeCoTaskMem(lParam)

    '        fromAPI = SendMessage(rtb.Handle, EM_FORMATRANGE, wParam, New IntPtr(0))

    '        gr.ReleaseHdc(hDC)

    '    End Using

    '    Return bmp

    'End Function
#End Region
End Class
