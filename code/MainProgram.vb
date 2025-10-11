Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Drawing.Text
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class MainProgram
    Dim AccountPictureUsersRegistryKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users", False)
    Dim CurrentUserSid As String = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI", False).GetValue("SelectedUserSID")
    Dim IsCurrentUserMissingProfilePicture As Boolean
    Dim LanguageAsString As String = String.Empty
    Dim StatusChecked As Boolean = False
    Dim UsersWithMissingProfilePictureSidList As New List(Of String)
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Change program language to display language if matches available languages, otherwise English
        If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
            ChangeLanguageTo("en")
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
            If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\es-ES\winbluelsppfix.resources.dll") Then
                ChangeLanguageTo("es-ES")
            Else
                ChangeLanguageTo("en")
            End If
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
            If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\fr\winbluelsppfix.resources.dll") Then
                ChangeLanguageTo("fr-FR")
            Else
                ChangeLanguageTo("en")
            End If
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "PTG" Then
            If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\pt-PT\winbluelsppfix.resources.dll") Then
                ChangeLanguageTo("pt-PT")
            Else
                ChangeLanguageTo("en")
            End If
        Else
            ChangeLanguageTo("en")
        End If

        ' Check if 64-bit program started when system is 64-bit
        If Environment.Is64BitOperatingSystem <> Environment.Is64BitProcess Then
            MessageBox.Show(My.Resources.ResourceManager.GetString("IsNot64BitProcess" & LanguageAsString), My.Resources.ResourceManager.GetString("Error" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if OS is Windows 8.1
        Using WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
            If WindowsVersion.GetValue("CurrentBuild") <> "9600" Then
                MessageBox.Show(My.Resources.ResourceManager.GetString("IsNotWin8" & LanguageAsString), My.Resources.ResourceManager.GetString("Error" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
            End If
        End Using

        ' Set font of VersionLabel from resources if not exist on system
        InitialiseVersionLabelFont()
    End Sub

    Private Sub LanguageToolStripMenuItem_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles LanguageToolStripMenuItem.DropDownItemClicked
        If Not DirectCast(e.ClickedItem, ToolStripMenuItem).Checked Then
            If e.ClickedItem.Text = "English" Then
                ChangeLanguageTo("en")
            ElseIf e.ClickedItem.Text = "español (España)" Then
                ChangeLanguageTo("es-ES")
            ElseIf e.ClickedItem.Text = "français (France)" Then
                ChangeLanguageTo("fr-FR")
            ElseIf e.ClickedItem.Text = "português (Portugal)" Then
                ChangeLanguageTo("pt-PT")
            End If
        End If
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        InfoWindow.ShowDialog()
    End Sub

    Private Sub ContactToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactToolStripMenuItem.Click
        ContactMe.ShowDialog()
    End Sub

    Private Sub CheckStatus_Click(sender As Object, e As EventArgs) Handles CheckStatusButton.Click, CheckStatusToolStripMenuItem.Click
        ' Check status for profile pictures
        MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("CheckingStatus" & LanguageAsString)
        CheckStatus()
    End Sub

    Private Sub ApplyFix_Click(sender As Object, e As EventArgs) Handles ApplyFixButton.Click, ApplyFixToolStripMenuItem.Click
        ' Run fix in sub FixMeNow
        MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("FixInProgress" & LanguageAsString)

        ' Try to fix, and show exception when error occurs
        Try
            If FixAllUsersCheckBox.Checked Then
                FixMeNow(True)
            Else
                FixMeNow(False)
            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.ResourceManager.GetString("ExceptionError" & LanguageAsString) & ex.Message, My.Resources.ErrorEnglish, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    Private Sub InfoLabel_Click(sender As Object, e As EventArgs) Handles InfoLabel.Click
        MessageBox.Show(My.Resources.ResourceManager.GetString("InformationFixAllUsers" & LanguageAsString), My.Resources.ResourceManager.GetString("Info" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub InitialiseVersionLabelFont()
        Dim FontDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
        If Not FileIO.FileSystem.FileExists(FontDirectory & "\SHOWG.TTF") Then
            Using ProgramFonts As New PrivateFontCollection
                Dim FontBuffer As IntPtr = Marshal.AllocCoTaskMem(My.Resources.ShowcardFontFile.Length)
                Marshal.Copy(My.Resources.ShowcardFontFile, 0, FontBuffer, My.Resources.ShowcardFontFile.Length)
                ProgramFonts.AddMemoryFont(FontBuffer, My.Resources.ShowcardFontFile.Length)
                With VersionLabel
                    .UseCompatibleTextRendering = True
                    .Font = New Font(ProgramFonts.Families(0), 10.2, FontStyle.Regular)
                End With
                Marshal.FreeCoTaskMem(FontBuffer)
            End Using
        End If
    End Sub

    Private Sub ChangeLanguageTo(ByVal ChangeToLanguage As String)
        If ChangeToLanguage <> "en" And FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & ChangeToLanguage & "\winbluelsppfix.resources.dll") Then
            My.Application.ChangeUICulture(ChangeToLanguage)
            InitialiseLanguageChange()
            LanguageAsString = StrConv(CultureInfo.GetCultureInfo(ChangeToLanguage).NativeName.Split(" (").First, VbStrConv.ProperCase)
            DirectCast(LanguageToolStripMenuItem.DropDownItems.Find(LanguageAsString & "ToolStripMenuItem", True).First, ToolStripMenuItem).Checked = True
        ElseIf ChangeToLanguage = "en" Then
            My.Application.ChangeUICulture(ChangeToLanguage)
            InitialiseLanguageChange()
            LanguageAsString = StrConv(CultureInfo.GetCultureInfo(ChangeToLanguage).NativeName, VbStrConv.ProperCase)
            EnglishToolStripMenuItem.Checked = True
        Else
            MessageBox.Show(My.Resources.ResourceManager.GetString("LocalisationFileNotFoundPart1" & LanguageAsString) & My.Application.Info.DirectoryPath & My.Resources.ResourceManager.GetString("LocalisationFileNotFoundPart2" & LanguageAsString), My.Resources.ResourceManager.GetString("Error" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub InitialiseLanguageChange()
        Me.Controls.Clear()
        InitializeComponent()
        InitialiseVersionLabelFont()
    End Sub

    Private Sub CheckStatus()
        If StatusChecked = False Then
            ' Get user SIDs from registry first
            If AccountPictureUsersRegistryKey IsNot Nothing Then
                Dim UsersSidList As String() = AccountPictureUsersRegistryKey.GetSubKeyNames
                Dim UsersWithProfilePictureSidList As New List(Of String)
                For Each UserSid As String In UsersSidList
                    If AccountPictureUsersRegistryKey.OpenSubKey(UserSid).GetValueNames.Length <> 0 Then
                        UsersWithProfilePictureSidList.Add(UserSid)
                    End If
                Next
                ' Check which users have no profile picture on login screen
                If UsersWithProfilePictureSidList.Count <> 0 Then
                    For Each UserSid As String In UsersWithProfilePictureSidList
                        If AccountPictureUsersRegistryKey.OpenSubKey(UserSid).GetValue("Image200") Is Nothing Then
                            UsersWithMissingProfilePictureSidList.Add(UserSid)
                            If UserSid = CurrentUserSid Then
                                IsCurrentUserMissingProfilePicture = True
                            End If
                        End If
                    Next
                    StatusChecked = True
                Else
                    MessageBox.Show(My.Resources.ResourceManager.GetString("NoUserProfilePicture" & LanguageAsString), My.Resources.ResourceManager.GetString("Warning" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError" & LanguageAsString)
                    Exit Sub
                End If
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("UserNameListError" & LanguageAsString), My.Resources.ResourceManager.GetString("Error" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Error)
                MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError" & LanguageAsString)
            End If
        End If

        ' If checked and missing profile pictures, enable apply fix button
        If UsersWithMissingProfilePictureSidList.Count <> 0 Then
            '' Show MessageBox with how many users have missing profile picture
            Dim IsCurrentUserMissingProfilePictureString As String = String.Empty
            If IsCurrentUserMissingProfilePicture = True Then
                IsCurrentUserMissingProfilePictureString = My.Resources.ResourceManager.GetString("Yes" & LanguageAsString)
            Else
                IsCurrentUserMissingProfilePictureString = My.Resources.ResourceManager.GetString("No" & LanguageAsString)
            End If
            MessageBox.Show(My.Resources.ResourceManager.GetString("UsersWithMissingProfilePictures" & LanguageAsString) & UsersWithMissingProfilePictureSidList.Count & vbCrLf & My.Resources.ResourceManager.GetString("CurrentUserMissingProfilePicture" & LanguageAsString) & IsCurrentUserMissingProfilePicture, My.Resources.ResourceManager.GetString("Info" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Information)
            '' Preparations before fix
            CheckStatusButton.Enabled = False
            CheckStatusToolStripMenuItem.Enabled = False
            ApplyFixButton.Enabled = True
            ApplyFixToolStripMenuItem.Enabled = True
            Me.AcceptButton = ApplyFixButton
            '' Check if one or more users and if current user missing profile picture
            If IsCurrentUserMissingProfilePicture = True And UsersWithMissingProfilePictureSidList.Count = 1 Then
                FixAllUsersCheckBox.Checked = False
                FixAllUsersCheckBox.Enabled = False
                InfoLabel.Enabled = False
            ElseIf IsCurrentUserMissingProfilePicture = True And UsersWithMissingProfilePictureSidList.Count > 1 Then
                FixAllUsersCheckBox.Enabled = True
                InfoLabel.Enabled = True
            ElseIf IsCurrentUserMissingProfilePicture = False Then
                FixAllUsersCheckBox.Enabled = False
                InfoLabel.Enabled = False
            End If
            '' Set status bar to ready
            MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("Ready" & LanguageAsString)
        Else
            MessageBox.Show(My.Resources.ResourceManager.GetString("FixAlreadyApplied" & LanguageAsString), My.Resources.ResourceManager.GetString("Warning" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarAlreadyApplied" & LanguageAsString)
        End If
    End Sub

    Private Sub FixMeNow(ByVal FixAllUsers As Boolean)
        If FixAllUsers = True Then
            ' Get profile picture path
            For Each UserSid As String In UsersWithMissingProfilePictureSidList
                Dim UserSidRegistryKey As RegistryKey = AccountPictureUsersRegistryKey.OpenSubKey(UserSid, False)
                Dim UserSidRegistryKeyPermissions As RegistrySecurity = UserSidRegistryKey.GetAccessControl
                Dim UserImageNumberValue As String = UserSidRegistryKey.GetValueNames(0)
                Dim UserProfilePicturePath As String = UserSidRegistryKey.GetValue(UserImageNumberValue)
                Dim UserImage200PicturePath As String = String.Empty
                '' Replace Image....jpg with Image200.jpg in string
                UserImage200PicturePath = UserProfilePicturePath.Replace(UserImageNumberValue, "Image200")
                '' Grant permission for adding registry value
                Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).Value
                Dim EveryoneAccessRule As RegistryAccessRule = New RegistryAccessRule(EveryoneSidAsUserName, RegistryRights.FullControl, AccessControlType.Allow)
                Dim EveryoneReadAccessRule As RegistryAccessRule = New RegistryAccessRule(EveryoneSidAsUserName, RegistryRights.ReadKey, InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow)
                UserSidRegistryKeyPermissions.AddAccessRule(EveryoneAccessRule)
                '' Add value Image200
                If FileIO.FileSystem.FileExists(UserImage200PicturePath) Then
                    '' Change permissions of registry key -> Everyone:Full Control
                    UserSidRegistryKey = AccountPictureUsersRegistryKey.OpenSubKey(UserSid, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.ChangePermissions)
                    UserSidRegistryKey.SetAccessControl(UserSidRegistryKeyPermissions)
                    '' Open registry key as writable now
                    UserSidRegistryKey = AccountPictureUsersRegistryKey.OpenSubKey(UserSid, True)
                    UserSidRegistryKey.SetValue("Image200", UserImage200PicturePath)
                    '' Set permissions back to original
                    UserSidRegistryKeyPermissions.RemoveAccessRule(EveryoneAccessRule)
                    UserSidRegistryKeyPermissions.AddAccessRule(EveryoneReadAccessRule)
                    UserSidRegistryKey.SetAccessControl(UserSidRegistryKeyPermissions)
                    UserSidRegistryKey.Close()
                Else
                    Dim UserName As String = New SecurityIdentifier(UserSid).Translate(GetType(NTAccount)).Value
                    MessageBox.Show(My.Resources.ResourceManager.GetString("ProfilePictureNotFoundPart1" & LanguageAsString) & UserImage200PicturePath & My.Resources.ResourceManager.GetString("ProfilePictureNotFoundPart2" & LanguageAsString) & UserName & My.Resources.ResourceManager.GetString("ProfilePictureNotFoundPart3" & LanguageAsString), My.Resources.ResourceManager.GetString("Error" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError" & LanguageAsString)
                End If
            Next
            AccountPictureUsersRegistryKey.Close()
        Else
            Dim UserSidRegistryKey As RegistryKey = AccountPictureUsersRegistryKey.OpenSubKey(CurrentUserSid, False)
            Dim UserSidRegistryKeyPermissions As RegistrySecurity = UserSidRegistryKey.GetAccessControl
            Dim UserImageNumberValue As String = UserSidRegistryKey.GetValueNames(0)
            Dim UserProfilePicturePath As String = UserSidRegistryKey.GetValue(UserImageNumberValue)
            Dim UserImage200PicturePath As String = String.Empty
            '' Replace Image....jpg with Image200.jpg in string
            UserImage200PicturePath = UserProfilePicturePath.Replace(UserImageNumberValue, "Image200")
            '' Grant permission for adding registry value
            Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).Value
            Dim EveryoneAccessRule As RegistryAccessRule = New RegistryAccessRule(EveryoneSidAsUserName, RegistryRights.FullControl, InheritanceFlags.None, PropagationFlags.None, AccessControlType.Allow)
            Dim EveryoneReadAccessRule As RegistryAccessRule = New RegistryAccessRule(EveryoneSidAsUserName, RegistryRights.ReadKey, InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow)
            UserSidRegistryKeyPermissions.AddAccessRule(EveryoneAccessRule)
            '' Add value Image200
            If FileIO.FileSystem.FileExists(UserImage200PicturePath) Then
                '' Change permissions of registry key -> Everyone:Full Control
                UserSidRegistryKey = AccountPictureUsersRegistryKey.OpenSubKey(CurrentUserSid, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.ChangePermissions)
                UserSidRegistryKey.SetAccessControl(UserSidRegistryKeyPermissions)
                '' Open registry key as writable now
                UserSidRegistryKey = AccountPictureUsersRegistryKey.OpenSubKey(CurrentUserSid, True)
                UserSidRegistryKey.SetValue("Image200", UserImage200PicturePath)
                '' Set permissions back to original
                UserSidRegistryKeyPermissions.RemoveAccessRule(EveryoneAccessRule)
                UserSidRegistryKeyPermissions.AddAccessRule(EveryoneReadAccessRule)
                UserSidRegistryKey.SetAccessControl(UserSidRegistryKeyPermissions)
                UserSidRegistryKey.Close()
                AccountPictureUsersRegistryKey.Close()
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("ProfilePictureCurrentUserNotFoundPart1" & LanguageAsString) & UserImage200PicturePath & My.Resources.ResourceManager.GetString("ProfilePictureCurrentUserNotFoundPart2" & LanguageAsString), My.Resources.ResourceManager.GetString("Error" & LanguageAsString), MessageBoxButtons.OK, MessageBoxIcon.Error)
                MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError" & LanguageAsString)
            End If
        End If

        ' If fix has been applied
        MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarDone" & LanguageAsString)

        ' Ask for log off
        Dim LogOffMessage As New DialogResult
        LogOffMessage = MessageBox.Show(My.Resources.ResourceManager.GetString("LogOffMessage" & LanguageAsString), My.Resources.ResourceManager.GetString("Success" & LanguageAsString), MessageBoxButtons.YesNo, MessageBoxIcon.Information)

        ' Log off if answer is "Yes", otherwise just close program
        If LogOffMessage = System.Windows.Forms.DialogResult.Yes Then
            Dim LogoffProcess As New Process
            With LogoffProcess
                .StartInfo.FileName = WinDir & "\System32\shutdown.exe"
                .StartInfo.Arguments = "/l"
                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                .Start()
            End With
            Me.Close()
        ElseIf LogOffMessage = System.Windows.Forms.DialogResult.No Then
            Me.Close()
        End If
    End Sub
End Class
