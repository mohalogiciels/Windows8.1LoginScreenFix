Imports System.Globalization
Imports System.IO
Imports System.Security.Principal
Imports Microsoft.Win32

Public Class MainProgram
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        InfoWindow.ShowDialog()
    End Sub

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set language of .NET Framework to English
        If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Then
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-GB")
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US")
        Else
            CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.GetCultureInfo("en-US")
        End If

        ' Check if 64-bit program started when system is 64-bit
        Dim Is64BitSystem As Boolean = Environment.Is64BitOperatingSystem
        Dim Is64BitProcess As Boolean = Environment.Is64BitProcess
        If Is64BitSystem <> Is64BitProcess Then
            MessageBox.Show("The 64-bit program could not be loaded!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if OS is Windows 8.1
        Dim WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
        If WindowsVersion.GetValue("CurrentVersion") <> "6.3" Then
            MessageBox.Show("This program is intended to only run on Windows 8.1!", "Not supported operating system", MessageBoxButtons.OK, MessageBoxIcon.Error)
            WindowsVersion.Close()
            Me.Close()
        End If
        WindowsVersion.Close()
    End Sub

    Private Sub ApplyFixButton_Click(sender As Object, e As EventArgs) Handles ApplyFixButton.Click
        ' Run fix in sub FixMeNow
        MainProgramToolStripStatusLabel.Text = "Fix in progress..."
        FixMeNow()
    End Sub

    Private Sub FixMeNow()
        Dim UserProfileSid As String = WindowsIdentity.GetCurrent.User.ToString
        Dim AccountPictureRegistryKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users\" & UserProfileSid, False)
        If AccountPictureRegistryKey IsNot Nothing Then
            If AccountPictureRegistryKey.GetValueNames.Length > 0 Then
                If AccountPictureRegistryKey.GetValue("Image200") Is Nothing Then
                    ' Get profile picture path
                    Dim AccountPicturePath As String = AccountPictureRegistryKey.GetValue(AccountPictureRegistryKey.GetValueNames(0))
                    Dim AccountPicturePathAsList As List(Of String) = Split(AccountPicturePath, "\").ToList
                    Dim AccountPicturePathImageNumber As String = String.Empty
                    For Each ch As Char In Split(AccountPicturePathAsList.Last, "Image").Last
                        If Char.IsDigit(ch) Then
                            AccountPicturePathImageNumber &= ch
                        End If
                    Next
                    ' Replace Image....jpg with Image200.jpg in string
                    AccountPicturePathAsList(AccountPicturePathAsList.Count - 1) = AccountPicturePathAsList.Last.Replace("Image" & AccountPicturePathImageNumber & ".jpg", "Image200.jpg")
                    Dim AccountPictureImage200Path As String = String.Join("\", AccountPicturePathAsList)
                    ' Check if picture exists
                    If FileIO.FileSystem.FileExists(AccountPictureImage200Path) Then
                        ' Grant permission for adding registry value
                        '' Create a file needed for executing regini.exe
                        FileIO.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\regini.txt", My.Resources.permission, False)
                        Dim ReginiTextFile As String() = File.ReadAllLines(My.Application.Info.DirectoryPath & "\regini.txt")
                        ReginiTextFile(0) = ReginiTextFile(0).Replace("UserProfileSid", UserProfileSid)
                        File.WriteAllLines(My.Application.Info.DirectoryPath & "\regini.txt", ReginiTextFile)
                        Dim ReginiProcess As New Process
                        With ReginiProcess
                            .StartInfo.FileName = WinDir & "\System32\regini.exe"
                            .StartInfo.Arguments = """" & My.Application.Info.DirectoryPath & "\regini.txt"""
                            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                            .Start()
                            .WaitForExit()
                        End With
                        ' Add value Image200
                        Try
                            AccountPictureRegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users\" & UserProfileSid, True)
                            AccountPictureRegistryKey.SetValue("Image200", AccountPictureImage200Path, RegistryValueKind.String)
                            AccountPictureRegistryKey.Close()
                            ' Delete file regini.txt
                            If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\regini.txt") Then
                                FileIO.FileSystem.DeleteFile(My.Application.Info.DirectoryPath & "\regini.txt")
                            End If
                        Catch ex As Exception
                            MessageBox.Show("An error occured during fixing: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            MainProgramToolStripStatusLabel.Text = "An error occured!"
                            Me.Close()
                        End Try
                        MainProgramToolStripStatusLabel.Text = "Finished!"
                        ' Ask for log off
                        If MessageBox.Show("The fix has been successfully applied! You need to log off and back on again to finish the modification. Do you want to log off now?", "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.Yes Then
                            Dim LogoffProcess As New Process
                            With LogoffProcess
                                .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                                .StartInfo.Arguments = "/l"
                                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                .Start()
                            End With
                            Me.Close()
                        ElseIf System.Windows.Forms.DialogResult.No Then
                            Me.Close()
                        End If
                    Else
                        MessageBox.Show("Profile picture in " & AccountPictureImage200Path & " could not be found on this system! Please try to set your profile picture and run this program again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        MainProgramToolStripStatusLabel.Text = "An error occured!"
                    End If
                Else
                    MessageBox.Show("Fix has already been applied, or missing registry key has already been set. If you still see no profile picture in the logon screen, try restarting your computer first, or set your profile picture again and run this program.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    MainProgramToolStripStatusLabel.Text = "An error occured!"
                End If
            Else
                MessageBox.Show("No profile picture has been set, or registry entries are broken!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                MainProgramToolStripStatusLabel.Text = "An error occured!"
            End If
        Else
            MessageBox.Show("Registry entry for your account could not be found or is broken!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MainProgramToolStripStatusLabel.Text = "An error occured!"
        End If
    End Sub
End Class
