Module modGrid

    Public Function NewCellHeight() As Integer
        Dim intNewCellHeight As Integer

        Dim fnt As New Font(My.Settings.setDefaultFont.FontFamily.Name, My.Settings.setDefaultFont.SizeInPoints)
        intNewCellHeight = fnt.Height + 9 'The height, in pixels, of the row. The default is the height of the default font plus 9 pixels.


        Return intNewCellHeight
    End Function

End Module



