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
    Dim StatusChecked As Boolean = False
    Dim UICulturesArray As String() = {"en", "es-ES", "fr-FR", "pt-PT"}
    Dim UICultureString As String = CultureInfo.CurrentUICulture.Name
    Dim UICulture As String = UICultureString.Replace("-", "_")
    Dim UsersWithMissingProfilePictureSidList As New List(Of String)
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' If UI language is English or not in list, set strings to universal English without region specified
        If CultureInfo.CurrentUICulture.Name.StartsWith("en") Or Not UICulturesArray.Contains(CultureInfo.CurrentUICulture.Name) Then
            UICultureString = "en"
            UICulture = "en"
        End If

        ' Check if 64-bit program started when system is 64-bit
        If Environment.Is64BitOperatingSystem <> Environment.Is64BitProcess Then
            MessageBox.Show(My.Resources.ResourceManager.GetString("IsNot64BitProcess_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End If

        ' Check if OS is Windows 8.1
        Using WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
            If WindowsVersion.GetValue("CurrentBuild") <> "9600" Then
                MessageBox.Show(My.Resources.ResourceManager.GetString("IsNotWin8_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
            End If
        End Using

        ' Set font of VersionLabel from resources if not exist on system
        InitialiseVersionLabelFont()
    End Sub

    Private Sub MainProgram_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' Check entry in language menu
        DirectCast(LanguageToolStripMenuItem.DropDownItems.Find(StrConv(CultureInfo.GetCultureInfo(UICultureString).NativeName.Split(" (").First, VbStrConv.ProperCase) & "ToolStripMenuItem", False).First, ToolStripMenuItem).Checked = True
    End Sub

    Private Sub LanguageToolStripMenuItem_DropDownItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles LanguageToolStripMenuItem.DropDownItemClicked
        If Not DirectCast(e.ClickedItem, ToolStripMenuItem).Checked Then
            Me.TopMost = True
            If e.ClickedItem.Text = "English" Then
                ChangeLanguageTo("en", e.ClickedItem.Text)
            ElseIf e.ClickedItem.Text = "español (España)" Then
                ChangeLanguageTo("es-ES", e.ClickedItem.Text)
            ElseIf e.ClickedItem.Text = "français (France)" Then
                ChangeLanguageTo("fr-FR", e.ClickedItem.Text)
            ElseIf e.ClickedItem.Text = "português (Portugal)" Then
                ChangeLanguageTo("pt-PT", e.ClickedItem.Text)
            End If
            Me.TopMost = False
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
        MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("CheckingStatus_" & UICulture)
        CheckStatus()
    End Sub

    Private Sub ApplyFix_Click(sender As Object, e As EventArgs) Handles ApplyFixButton.Click, ApplyFixToolStripMenuItem.Click
        ' Try to fix, and show exception when error occurs
        MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("FixInProgress_" & UICulture)
        Try
            '' Run fix in sub FixMeNow(ByVal FixAllUsers As Boolean)
            If FixAllUsersCheckBox.Checked Then
                FixMeNow(True)
            Else
                FixMeNow(False)
            End If
        Catch ex As Exception
            MessageBox.Show(My.Resources.ResourceManager.GetString("ExceptionError_" & UICulture) & ex.Message, My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    Private Sub InfoLabel_Click(sender As Object, e As EventArgs) Handles InfoLabel.Click
        MessageBox.Show(My.Resources.ResourceManager.GetString("InformationFixAllUsers_" & UICulture), My.Resources.ResourceManager.GetString("Info_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Information)
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

    Private Sub ChangeLanguageTo(ByVal ChangeToLanguage As String, ByVal LanguageAsString As String)
        If ChangeToLanguage <> "en" And FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & ChangeToLanguage & "\winbluelsppfix.resources.dll") Then
            My.Application.ChangeUICulture(ChangeToLanguage)
            InitialiseLanguageChange()
            UICultureString = ChangeToLanguage
            UICulture = UICultureString.Replace("-", "_")
            DirectCast(LanguageToolStripMenuItem.DropDownItems.Find(StrConv(LanguageAsString.Split(" (").First, VbStrConv.ProperCase) & "ToolStripMenuItem", False).First, ToolStripMenuItem).Checked = True
        ElseIf ChangeToLanguage = "en" Then
            My.Application.ChangeUICulture(ChangeToLanguage)
            InitialiseLanguageChange()
            UICultureString = "en"
            UICulture = "en"
            EnglishToolStripMenuItem.Checked = True
        Else
            MessageBox.Show(My.Resources.ResourceManager.GetString("LocalisationFileNotFoundPart1_" & UICulture) & My.Application.Info.DirectoryPath & My.Resources.ResourceManager.GetString("LocalisationFileNotFoundPart2_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                    MessageBox.Show(My.Resources.ResourceManager.GetString("NoUserProfilePicture_" & UICulture), My.Resources.ResourceManager.GetString("Warning_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError_" & UICulture)
                    Exit Sub
                End If
            Else
                MessageBox.Show(My.Resources.ResourceManager.GetString("UserNameListError_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
                MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError_" & UICulture)
            End If
        End If

        ' If checked and missing profile pictures, enable apply fix button
        If UsersWithMissingProfilePictureSidList.Count <> 0 Then
            '' Show MessageBox with how many users have missing profile picture
            Dim IsCurrentUserMissingProfilePictureString As String = String.Empty
            If IsCurrentUserMissingProfilePicture = True Then
                IsCurrentUserMissingProfilePictureString = My.Resources.ResourceManager.GetString("Yes_" & UICulture)
            Else
                IsCurrentUserMissingProfilePictureString = My.Resources.ResourceManager.GetString("No_" & UICulture)
            End If
            MessageBox.Show(My.Resources.ResourceManager.GetString("UsersWithMissingProfilePictures_" & UICulture) & UsersWithMissingProfilePictureSidList.Count & vbCrLf & My.Resources.ResourceManager.GetString("CurrentUserMissingProfilePicture_" & UICulture) & IsCurrentUserMissingProfilePictureString, My.Resources.ResourceManager.GetString("Info_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("Ready_" & UICulture)
        Else
            MessageBox.Show(My.Resources.ResourceManager.GetString("FixAlreadyApplied_" & UICulture), My.Resources.ResourceManager.GetString("Warning_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarAlreadyApplied_" & UICulture)
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
                    MessageBox.Show(My.Resources.ResourceManager.GetString("ProfilePictureNotFoundPart1_" & UICulture) & UserImage200PicturePath & My.Resources.ResourceManager.GetString("ProfilePictureNotFoundPart2_" & UICulture) & UserName & My.Resources.ResourceManager.GetString("ProfilePictureNotFoundPart3_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError_" & UICulture)
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
                MessageBox.Show(My.Resources.ResourceManager.GetString("ProfilePictureCurrentUserNotFoundPart1_" & UICulture) & UserImage200PicturePath & My.Resources.ResourceManager.GetString("ProfilePictureCurrentUserNotFoundPart2_" & UICulture), My.Resources.ResourceManager.GetString("Error_" & UICulture), MessageBoxButtons.OK, MessageBoxIcon.Error)
                MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarError_" & UICulture)
            End If
        End If

        ' If fix has been applied, show in status bar
        MainProgramToolStripStatusLabel.Text = My.Resources.ResourceManager.GetString("StatusBarDone_" & UICulture)

        ' Ask for log off
        Dim LogOffMessage As New DialogResult
        LogOffMessage = MessageBox.Show(My.Resources.ResourceManager.GetString("LogOffMessage_" & UICulture), My.Resources.ResourceManager.GetString("Success_" & UICulture), MessageBoxButtons.YesNo, MessageBoxIcon.Information)

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
