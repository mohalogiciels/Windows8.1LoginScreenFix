Imports System.Globalization
Imports System.IO
Imports System.IO.Compression

Public Class Launcher

    Dim FilesExtractToPath As String = My.Application.Info.DirectoryPath & "\winbluelsppfix"
    Dim UICulturesArray As String() = {"en", "es-ES", "fr-FR", "pt-PT"}
    Dim UICulture As String = CultureInfo.CurrentUICulture.Name.Replace("-", "_")

    Private Sub Launcher_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' If UI language is English or not in list, set strings to universal English without region specified
        If CultureInfo.CurrentUICulture.Name.StartsWith("en") Or Not UICulturesArray.Contains(CultureInfo.CurrentUICulture.Name) Then
            UICulture = "en"
        End If

        ' Set status bar text to localised string
        StatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBar_" & UICulture)
    End Sub

    Private Sub Launcher_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Refresh()
        If ExtractZipFileTo(My.Resources.WindowsBlueLoginScreenFixZipFile, FilesExtractToPath) = True Then
            LaunchProgram(FilesExtractToPath & "\winbluelsppfix.exe")
            Me.Close()
        Else
            Me.Close()
        End If
    End Sub

    Private Function ExtractZipFileTo(ByVal ZipFileToExtract As Byte(), ByVal PathToFolder As String) As Boolean
        Try
            Using ZipFileToOpen As New MemoryStream(ZipFileToExtract)
                Using ZipFileAsArchive As New ZipArchive(ZipFileToOpen, ZipArchiveMode.Read)
                    ZipFileAsArchive.ExtractToDirectory(PathToFolder)
                End Using
            End Using
        Catch ex As Exception
            Dim ErrorMessage As DialogResult = MessageBox.Show(My.Resources.ResourceManager.GetString("ExtractingErrorPart1_" & UICulture) & ex.Message & My.Resources.ResourceManager.GetString("ExtractingErrorPart2_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OKCancel, MessageBoxIcon.Error)
            If ErrorMessage = System.Windows.Forms.DialogResult.OK Then
                Dim NewPath As New FolderBrowserDialog
                If NewPath.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    FilesExtractToPath = NewPath.SelectedPath & "\winbluelsppfix"
                    ExtractZipFileTo(ZipFileToExtract, FilesExtractToPath)
                Else
                    MessageBox.Show(My.Resources.ResourceManager.GetString("NoPathChosen_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                    Exit Function
                End If
            Else
                Return False
                Exit Function
            End If
        End Try
        Return True
    End Function

    Private Sub LaunchProgram(ByVal PathToProgram As String)
        If FileIO.FileSystem.FileExists(PathToProgram) Then
            Using StartProgram As New Process
                With StartProgram
                    .StartInfo.FileName = PathToProgram
                    .Start()
                End With
            End Using
        Else
            MessageBox.Show(My.Resources.ResourceManager.GetString("FilesCouldNotBeExtracted_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class
