Public Class classSelectedText
    Private selFont As Font
    Private intSelectionStart As Integer, intSelectionLength As Integer
    Private intCharOffSet As Integer
    Private intHorizontalAlignment As Integer
    Private boolFontComplex, boolFontSizeComplex As Boolean
    Private boolBoldComplex, boolItalicComplex, boolUnderlineComplex, boolStrikeoutComplex As Boolean
    Private boolCharOffSetComplex As Boolean

    Public Event FontNameChanged()
    Public Event FontSizeChanged()
    Public Event FontBoldChanged()
    Public Event FontItalicChanged()
    Public Event FontUnderlineChanged()
    Public Event FontStrikeoutChanged()
    Public Event FontCharOffSetChanged()
    Public Event ParagraphAlignmentChanged()

    Public Enum TextOffSetValues
        SuperScript = 1
        SubScript = -1
    End Enum

#Region "Font"
    Public Property Font() As Font
        Get
            Return selFont
        End Get
        Set(value As Font)
            Dim OldFont As Font = Nothing
            If selFont IsNot Nothing Then OldFont = selFont
            selFont = value

            If OldFont Is Nothing Then
                RaiseEvent FontNameChanged()
                RaiseEvent FontSizeChanged()
                If selFont.Bold Then RaiseEvent FontBoldChanged()
                If selFont.Italic Then RaiseEvent FontItalicChanged()
                If selFont.Underline Then RaiseEvent FontUnderlineChanged()
                If selFont.Strikeout Then RaiseEvent FontStrikeoutChanged()
            Else
                If selFont IsNot Nothing Then
                    If OldFont.Name <> selFont.Name Then RaiseEvent FontNameChanged()
                    If OldFont.Size <> selFont.Size Then RaiseEvent FontSizeChanged()
                    If OldFont.Bold <> selFont.Bold Then RaiseEvent FontBoldChanged()
                    If OldFont.Italic <> selFont.Italic Then RaiseEvent FontItalicChanged()
                    If OldFont.Underline <> selFont.Underline Then RaiseEvent FontUnderlineChanged()
                    If OldFont.Strikeout <> selFont.Strikeout Then RaiseEvent FontStrikeoutChanged()
                End If
            End If
        End Set
    End Property

    Public Property CharOffSet() As Integer
        Get
            Return intCharOffSet
        End Get
        Set(value As Integer)
            intCharOffSet = value

            RaiseEvent FontCharOffSetChanged()
        End Set
    End Property

    Public ReadOnly Property SuperScript As Boolean
        Get
            If CharOffSet > 0 Then Return True Else Return False
        End Get
    End Property

    Public ReadOnly Property SubScript As Boolean
        Get
            If CharOffSet < 0 Then Return True Else Return False
        End Get
    End Property

    Public Property FontComplex As Boolean
        Get
            Return boolFontComplex
        End Get
        Set(value As Boolean)
            boolFontComplex = value
        End Set
    End Property

    Public Property FontSizeComplex As Boolean
        Get
            Return boolFontSizeComplex
        End Get
        Set(value As Boolean)
            boolFontSizeComplex = value
        End Set
    End Property

    Public Property BoldComplex As Boolean
        Get
            Return boolBoldComplex
        End Get
        Set(value As Boolean)
            boolBoldComplex = value
        End Set
    End Property

    Public Property ItalicComplex As Boolean
        Get
            Return boolItalicComplex
        End Get
        Set(value As Boolean)
            boolItalicComplex = value
        End Set
    End Property

    Public Property UnderlineComplex As Boolean
        Get
            Return boolUnderlineComplex
        End Get
        Set(value As Boolean)
            boolUnderlineComplex = value
        End Set
    End Property

    Public Property StrikeoutComplex As Boolean
        Get
            Return boolStrikeoutComplex
        End Get
        Set(value As Boolean)
            boolStrikeoutComplex = value
        End Set
    End Property

    Public Property CharOffSetComplex As Boolean
        Get
            Return boolCharOffSetComplex
        End Get
        Set(value As Boolean)
            boolCharOffSetComplex = value
        End Set
    End Property
#End Region

#Region "Paragraph"
    Public Property SelectionStart() As Integer
        Get
            Return intSelectionStart
        End Get
        Set(ByVal value As Integer)
            intSelectionStart = value
        End Set
    End Property

    Public Property SelectionLength() As Integer
        Get
            Return intSelectionLength
        End Get
        Set(ByVal value As Integer)
            intSelectionLength = value
        End Set
    End Property

    Public Property HorizontalAlignment As Integer
        Get
            Return intHorizontalAlignment
        End Get
        Set(value As Integer)
            intHorizontalAlignment = value

            RaiseEvent ParagraphAlignmentChanged()
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub InitializeFont(ByVal font As Font)
        selFont = font
        RaiseEvent FontNameChanged()
        RaiseEvent FontSizeChanged()
        RaiseEvent FontBoldChanged()
        RaiseEvent FontItalicChanged()
        RaiseEvent FontUnderlineChanged()
        RaiseEvent FontStrikeoutChanged()
    End Sub

    Public Sub ChangeFontName(ByVal strFontName As String)
        Me.Font = New Font(strFontName, Me.Font.Size, Me.Font.Style)
    End Sub

    Public Sub ChangeFontSize(ByVal intFontSize As Integer)
        Me.Font = New Font(Me.Font.FontFamily, intFontSize, Me.Font.Style)
    End Sub

    Public Sub ChangeStyle(Optional ByVal boolBold As Boolean = False, Optional ByVal boolItalic As Boolean = False, _
                           Optional ByVal boolUnderline As Boolean = False, Optional ByVal boolStrikeOut As Boolean = False)
        Dim NewFontStyle As New FontStyle
        NewFontStyle = FontStyle.Regular

        If boolBold = True Then NewFontStyle = FontStyle.Bold
        If boolItalic = True Then NewFontStyle += FontStyle.Italic
        If boolUnderline = True Then NewFontStyle += FontStyle.Underline
        If boolStrikeOut = True Then NewFontStyle += FontStyle.Strikeout

        Try
            Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size, NewFontStyle, GraphicsUnit.Point)
        Catch exNoSuchStyle As ArgumentException
            Try
                Me.Font = New Font(Me.Font.FontFamily, Me.Font.Size, Me.Font.Style, GraphicsUnit.Point)
            Catch exFontError As ArgumentException
                Dim strMsg As String = ERR_FontStyleNotSupported

                MsgBox(strMsg)
            End Try
        End Try
    End Sub

    Public Sub ChangeCharOffset(ByVal intUpDown As Integer)
        Dim intSize As Integer

        If intUpDown <> 0 Then
            Me.CharOffSet = Me.Font.Size * intUpDown * 0.4
            intSize = Me.Font.Size * 0.6
            ChangeFontSize(intSize)

            RaiseEvent FontCharOffSetChanged()
        End If
    End Sub

    Public Sub Reset()
        Font = My.Settings.setDefaultFont
        FontComplex = False
        FontSizeComplex = False
        BoldComplex = False
        ItalicComplex = False
        UnderlineComplex = False
        StrikeoutComplex = False
        SelectionStart = 0
        SelectionLength = 0
        HorizontalAlignment = Windows.Forms.HorizontalAlignment.Left

        CharOffSet = 0
    End Sub
#End Region
End Class
