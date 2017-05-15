Imports System.IO

Public Class ucImage
    Friend WithEvents AVBindingSource As New BindingSource
    Friend WithEvents ImagesBindingSource As New BindingSource

    Private objAudioVisualDetail As AudioVisualDetail
    Private objImageURLs As New List(Of Uri)

    Public Sub New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal audiovisualdetail As AudioVisualDetail)
        InitializeComponent()

        Me.AudioVisualDetail = audiovisualdetail

        LoadItems()
    End Sub

    Public Property AudioVisualDetail As AudioVisualDetail
        Get
            Return objAudioVisualDetail
        End Get
        Set(value As AudioVisualDetail)
            objAudioVisualDetail = value
        End Set
    End Property

    Public Property ImageURLs As List(Of Uri)
        Get
            Return objImageURLs
        End Get
        Set(value As List(Of Uri))
            objImageURLs = value
        End Set
    End Property


    Public Sub LoadItems()

        If AudioVisualDetail IsNot Nothing Then
            AVBindingSource.DataSource = AudioVisualDetail
            ImagesBindingSource.DataSource = ImageURLs
            bnImages.BindingSource = ImagesBindingSource

            tbURL.DataBindings.Add("Text", AVBindingSource, "URL")
            tbFileName.DataBindings.Add("Text", ImagesBindingSource, "LocalPath")
            tbDescription.DataBindings.Add("Text", AVBindingSource, "Description")
            pbImage.DataBindings.Add("ImageLocation", ImagesBindingSource, "LocalPath")

            LoadImages()
        End If
    End Sub

    Public Sub LoadImages()
        Dim strUrl As String

        If String.IsNullOrEmpty(AudioVisualDetail.URL) = False Then
            ImageURLs.Clear()
            strUrl = AudioVisualDetail.URL

            If strUrl.EndsWith("\") Then
                LoadImagesFromFolder(strUrl)
            Else
                LoadImageFromFile(strUrl)
            End If

            ImagesBindingSource.ResetBindings(False)
        End If
    End Sub

    Private Sub LoadImagesFromFolder(ByVal strUrl)
        Dim FileDirectory As New IO.DirectoryInfo(strUrl)
        Dim FilePng As IO.FileInfo() = FileDirectory.GetFiles("*.png")
        Dim FileJpg As IO.FileInfo() = FileDirectory.GetFiles("*.jpg")
        Dim FileGif As IO.FileInfo() = FileDirectory.GetFiles("*.gif")
        Dim FileBmp As IO.FileInfo() = FileDirectory.GetFiles("*.bmp")

        For Each File As IO.FileInfo In FilePng
            Dim NewUri As New Uri(File.FullName)
            ImageURLs.Add(NewUri)
        Next
        For Each File As IO.FileInfo In FileJpg
            ImageURLs.Add(New Uri(File.FullName))
        Next
        For Each File As IO.FileInfo In FileGif
            ImageURLs.Add(New Uri(File.FullName))
        Next
        For Each File As IO.FileInfo In FileBmp
            ImageURLs.Add(New Uri(File.FullName))
        Next
    End Sub

    Private Sub LoadImageFromFile(strUrl)
        If File.Exists(strUrl) Then
            ImageURLs.Add(New Uri(strUrl))
        End If
    End Sub

    Private Sub tbURL_Validated(sender As Object, e As System.EventArgs) Handles tbURL.Validated
        LoadImages()
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As System.EventArgs) Handles btnBrowse.Click
        Dim DialogImageSelect As New DialogImageSelect

        With DialogImageSelect
            If .ShowDialog = DialogResult.OK Then
                If .rbtnSingleFile.Checked = True Then
                    Dim FileDialog As New OpenFileDialog
                    Dim strFilter As String = String.Format("{0} (*.jpg, *.bmp, *.gif, *.tif, *.png)|*.jpg; *.bmp; *.tif; *.gif; *.png|{1} (*.*)|*.*", LANG_Images, LANG_AllFiles)
                    With FileDialog
                        .Title = LANG_SelectImageTitle
                        .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                        .Multiselect = False
                        .Filter = strFilter
                        .FilterIndex = 0

                        If .ShowDialog() = DialogResult.OK And String.IsNullOrEmpty(.FileName) = False Then
                            AudioVisualDetail.URL = .FileName
                            tbURL.DataBindings(0).ReadValue()
                        End If
                    End With
                Else
                    Dim Folderdialog As New FolderBrowserDialog

                    With Folderdialog
                        .Description = LANG_SelectImageFolderTitle
                        .RootFolder = Environment.SpecialFolder.Desktop
                        .ShowNewFolderButton = False
                        .SelectedPath = Environment.SpecialFolder.MyDocuments
                        If .ShowDialog() = DialogResult.OK And String.IsNullOrEmpty(.SelectedPath) = False Then
                            AudioVisualDetail.URL = .SelectedPath & "\"
                            tbURL.DataBindings(0).ReadValue()
                        End If
                    End With
                End If

                LoadImages()
            End If
        End With
    End Sub

    Private Sub pbImage_DoubleClick(sender As Object, e As System.EventArgs) Handles pbImage.DoubleClick
        Dim strUrl As String = pbImage.ImageLocation

        Try
            Process.Start(strUrl)
        Catch ex As Exception

        End Try

    End Sub
End Class
