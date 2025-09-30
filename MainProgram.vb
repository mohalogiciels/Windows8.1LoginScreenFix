Imports Microsoft.Win32
Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class MainProgram
    Dim AccountPictureUsersRegistryKey As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\AccountPicture\Users", False)
    Dim CurrentUserSid As String = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI", False).GetValue("SelectedUserSID")
    Dim IsCurrentUserMissingProfilePicture As Boolean
    Dim Language As String = String.Empty
    Dim StatusChecked As Boolean = False
    Dim UsersWithMissingProfilePictureSidList As New List(Of String)
    Dim WinDir As String = Environment.GetFolderPath(Environment.SpecialFolder.Windows)

    Private Sub ChangeLanguageTo(ByVal ChangeToLanguage As String)
        If ChangeToLanguage = "English" Then
            Language = "English"
            My.Application.ChangeCulture("en")
            My.Application.ChangeUICulture("en")
            Me.Controls.Clear()
            InitializeComponent()
            EnglishToolStripMenuItem.Checked = True
            EnglishToolStripMenuItem.Enabled = False
        ElseIf ChangeToLanguage = "Español (España)" Then
            If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\es-ES\winbluelsppfix.resources.dll") Then
                Language = "Español (España)"
                My.Application.ChangeCulture("es-ES")
                My.Application.ChangeUICulture("es-ES")
                Me.Controls.Clear()
                InitializeComponent()
                EspanolToolStripMenuItem.Checked = True
                EspanolToolStripMenuItem.Enabled = False
            Else
                MessageBox.Show("¡La ruta """ & My.Application.Info.DirectoryPath & "\es-ES\winbluelsppfix.resources.dll"" no se encuentra en su sistema! ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        ElseIf ChangeToLanguage = "Français" Then
            If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\fr\winbluelsppfix.resources.dll") Then
                Language = "Français"
                My.Application.ChangeCulture("fr")
                My.Application.ChangeUICulture("fr")
                Me.Controls.Clear()
                InitializeComponent()
                FrancaisToolStripMenuItem.Checked = True
                FrancaisToolStripMenuItem.Enabled = False
            Else
                MessageBox.Show("Le chemin """ & My.Application.Info.DirectoryPath & "\fr\winbluelsppfix.resources.dll"" est introuvable sur votre système !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Sub MainProgram_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set default language to display language, otherwise English
        If CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENG" Or CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ENU" Then
            ChangeLanguageTo("English")
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "ESN" Then
            ChangeLanguageTo("Español (España)")
        ElseIf CultureInfo.CurrentUICulture.ThreeLetterWindowsLanguageName = "FRA" Then
            ChangeLanguageTo("Français")
        Else
            ChangeLanguageTo("English")
        End If

        ' Check if 64-bit program started when system is 64-bit
        Dim Is64BitSystem As Boolean = Environment.Is64BitOperatingSystem
        Dim Is64BitProcess As Boolean = Environment.Is64BitProcess
        If Is64BitSystem <> Is64BitProcess Then
            If Language = "English" Then
                MessageBox.Show("The 64-bit program could not be loaded!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Language = "Español (España)" Then
                MessageBox.Show("¡No se pudo iniciar el programa de 64 bits!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Language = "Français" Then
                MessageBox.Show("Ce logiciel 64 bits ne peut être pas lancé sur cet ordinateur !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End If

        ' Check if OS is Windows 8.1
        Dim WindowsVersion As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion", False)
        If WindowsVersion.GetValue("CurrentBuild") <> "9600" Then
            If Language = "English" Then
                MessageBox.Show("This program is intended to only run on Windows 8.1!", "Not supported operating system", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Language = "Español (España)" Then
                MessageBox.Show("¡Este programa solo está diseñado para ejecutarse en Windows 8.1!", "Sistema operativo no compatible", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Language = "Français" Then
                MessageBox.Show("Ce logiciel est conçu pour lancer uniquement sous Windows 8.1 !", "Système d’exploitation non compatible", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            WindowsVersion.Close()
            Me.Close()
        End If
        WindowsVersion.Close()
    End Sub

    Private Sub LanguageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnglishToolStripMenuItem.Click, EspanolToolStripMenuItem.Click, FrancaisToolStripMenuItem.Click
        ' Set language of program and .NET Framework
        Dim SelectedLanguage = CType(sender, ToolStripMenuItem)
        ChangeLanguageTo(SelectedLanguage.Text)
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        InfoWindow.ShowDialog()
    End Sub

    Private Sub CheckStatus_Click(sender As Object, e As EventArgs) Handles CheckStatusButton.Click, CheckStatusToolStripMenuItem.Click
        ' Check status for profile pictures
        If Language = "English" Then
            MainProgramToolStripStatusLabel.Text = "Checking status..."
        ElseIf Language = "Español (España)" Then
            MainProgramToolStripStatusLabel.Text = "Comprobando estado ahora..."
        ElseIf Language = "Français" Then
            MainProgramToolStripStatusLabel.Text = "Vérification de l’état en cours..."
        End If
        CheckStatus()
    End Sub

    Private Sub ApplyFix_Click(sender As Object, e As EventArgs) Handles ApplyFixButton.Click, ApplyFixToolStripMenuItem.Click
        ' Run fix in sub FixMeNow
        If Language = "English" Then
            MainProgramToolStripStatusLabel.Text = "Fix in progress..."
        ElseIf Language = "Español (España)" Then
            MainProgramToolStripStatusLabel.Text = "Corrección en curso..."
        ElseIf Language = "Français" Then
            MainProgramToolStripStatusLabel.Text = "Correctif en cours..."
        End If

        ' Try to fix, and show exception when error occurs
        Try
            If FixAllUsersCheckBox.Checked Then
                FixMeNow(True)
            Else
                FixMeNow(False)
            End If
        Catch ex As Exception
            If Language = "English" Then
                MessageBox.Show("An error occured during fixing: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Language = "Español (España)" Then
                MessageBox.Show("Se ha producido un error durante aplicar la corrección: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Language = "Français" Then
                MessageBox.Show("L’erreur est survenue pendant appliquer le correctif : " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            Me.Close()
        End Try
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
                    If Language = "English" Then
                        MessageBox.Show("No user has a profile picture set, or registry entries can’t be found!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        MainProgramToolStripStatusLabel.Text = "An error occured!"
                    ElseIf Language = "Español (España)" Then
                        MessageBox.Show("¡Ningún usuario tiene una foto de perfil configurada o no se pueden encontrar las entradas del registro!")
                        MainProgramToolStripStatusLabel.Text = "¡Se ha producido un error!"
                    ElseIf Language = "Français" Then
                        MessageBox.Show("Aucun utilisateur n’a défini une photo de profil, ou les entrées de registre sont introuvables !", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        MainProgramToolStripStatusLabel.Text = "Il y a eu une erreur !"
                    End If
                    Exit Sub
                End If
            Else
                If Language = "English" Then
                    MessageBox.Show("User name list could not be opened or found on this system!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = "An error occured!"
                ElseIf Language = "Español (España)" Then
                    MessageBox.Show("¡No se pudo abrir o encontrar la lista de nombres de usuario en este sistema!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = "¡Se ha producido un error!"
                ElseIf Language = "Français" Then
                    MessageBox.Show("La liste des noms d’utilisateurs n’a pas pu être ouverte ou trouvée sur cet ordinateur !", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = "Il y a eu une erreur !"
                End If
            End If
        End If

        ' If checked and missing profile pictures, enable apply fix button
        If UsersWithMissingProfilePictureSidList.Count <> 0 Then
            '' Show MessageBox with how many users have missing profile picture
            Dim IsCurrentUserMissingProfilePictureString As String = String.Empty
            If Language = "English" Then
                MessageBox.Show("Users with missing profile pictures found: " & UsersWithMissingProfilePictureSidList.Count & vbCrLf & "Current user missing profile picture: " & IsCurrentUserMissingProfilePicture, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf Language = "Español (España)" Then
                If IsCurrentUserMissingProfilePicture = True Then
                    IsCurrentUserMissingProfilePictureString = "Sí"
                Else
                    IsCurrentUserMissingProfilePictureString = "No"
                End If
                MessageBox.Show("Usuarios con fotos de perfil faltantes encontrados: " & UsersWithMissingProfilePictureSidList.Count & vbCrLf & "El usuario actual no tiene foto de perfil: " & IsCurrentUserMissingProfilePictureString, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf Language = "Français" Then
                If IsCurrentUserMissingProfilePicture = True Then
                    IsCurrentUserMissingProfilePictureString = "Oui"
                Else
                    IsCurrentUserMissingProfilePictureString = "Non"
                End If
                MessageBox.Show("Utilisateurs avec photo de profil manquante : " & UsersWithMissingProfilePictureSidList.Count & vbCrLf & "Utilisateur actuel avec photo de profil manquante : " & IsCurrentUserMissingProfilePictureString, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
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
            If Language = "English" Then
                MainProgramToolStripStatusLabel.Text = "Ready"
            ElseIf Language = "Español (España)" Then
                MainProgramToolStripStatusLabel.Text = "Listo"
            ElseIf Language = "Français" Then
                MainProgramToolStripStatusLabel.Text = "Prêt"
            End If
        Else
            If Language = "English" Then
                MessageBox.Show("It looks like that the fix has already been applied for all users! If you still see one or more profile pictures missing, set the profile picture and apply this fix again.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                MainProgramToolStripStatusLabel.Text = "Info: fix has already been applied!"
            ElseIf Language = "Español (España)" Then
                MessageBox.Show("¡Parece que la solución ya se ha aplicado para todos los usuarios! Si sigues viendo que faltan una o más fotos de perfil, configura la foto de perfil y vuelve a aplicar esta solución.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                MainProgramToolStripStatusLabel.Text = "Información: ¡La corrección ya se ha aplicado!"
            ElseIf Language = "Français" Then
                MessageBox.Show("Le correctif a peut-être déjà été appliqué pour tous les utilisateurs ! Si vous regardez qu’une ou plusieurs photos de profil sont manquantes, veuillez changer la photo de profil et appliquez à nouveau ce correctif.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                MainProgramToolStripStatusLabel.Text = "Information : correctif a déjà été appliqué !"
            End If
        End If
    End Sub

    Private Sub InfoLabel_Click(sender As Object, e As EventArgs) Handles InfoLabel.Click
        If Language = "English" Then
            MessageBox.Show("If you check this option, the program will fix missing profile pictures for all users where the registry entry is missing. Uncheck this option to only fix it for the current user.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf Language = "Español (España)" Then
            MessageBox.Show("Si marca esta opción, el programa corregirá las imágenes de perfil que faltan para todos los usuarios en los que falte la entrada del registro. Desmarque esta opción para corregirlo solo para el usuario actual.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf Language = "Français" Then
            MessageBox.Show("Si vous sélectionnez cette option, le logiciel corrigera les photos de profil manquantes pour tous les utilisateurs dont l’entrée de registre est manquante. Veuillez désélectionner cette option pour ne corriger que celle de l’utilisateur actuel.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
                Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).ToString
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
                    Dim UserName As String = New SecurityIdentifier(UserSid).Translate(GetType(NTAccount)).ToString
                    If Language = "English" Then
                        MessageBox.Show("Profile picture in """ & UserImage200PicturePath & """ could not be found on this system! Please try to set the profile picture for user """ & UserName & """ and run the fix again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        MainProgramToolStripStatusLabel.Text = "An error occured!"
                    ElseIf Language = "Español (España)" Then
                        MessageBox.Show("¡No se ha encontrado la imagen de perfil en la ruta """ & UserImage200PicturePath & """ en este sistema! Intente establecer la imagen de perfil para el usuario """ & UserName & """ y vuelva a ejecutar la corrección.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        MainProgramToolStripStatusLabel.Text = "¡Se ha producido un error!"
                    ElseIf Language = "Français" Then
                        MessageBox.Show("La photo de profil dans le chemin """ & UserImage200PicturePath & """ est introuvable sur cet ordinateur ! Veuillez essayer de changer la photo de profil pour l’utilisateur """ & UserName & """ et relancer la correction.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        MainProgramToolStripStatusLabel.Text = "Il y a eu une erreur !"
                    End If
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
            Dim EveryoneSidAsUserName As String = New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing).Translate(GetType(NTAccount)).ToString
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
                If Language = "English" Then
                    MessageBox.Show("Profile picture in """ & UserImage200PicturePath & """ could not be found on this system! Please try to set your profile picture and run the fix again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = "An error occured!"
                ElseIf Language = "Español (España)" Then
                    MessageBox.Show("¡No se ha encontrado la imagen de perfil en """ & UserImage200PicturePath & """ en este sistema! Intenta configurar tu imagen de perfil y vuelve a ejecutar la corrección.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = "¡Se ha producido un error!"
                ElseIf Language = "Français" Then
                    MessageBox.Show("La photo de profil dans le chemin """ & UserImage200PicturePath & """ est introuvable sur cet ordinateur ! Veuillez essayer de mettre votre photo de profil et relancez la correction.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    MainProgramToolStripStatusLabel.Text = "Il y a eu une erreur !"
                End If
            End If
        End If

        ' If fix has been applied
        If Language = "English" Then
            MainProgramToolStripStatusLabel.Text = "Done"
        ElseIf Language = "Español (España)" Then
            MainProgramToolStripStatusLabel.Text = "Listo"
        ElseIf Language = "Français" Then
            MainProgramToolStripStatusLabel.Text = "Terminé"
        End If

        ' Ask for log off
        Dim LogOffMessage As New DialogResult
        If Language = "English" Then
            LogOffMessage = MessageBox.Show("The fix has been successfully applied! You need to log off and back on again to finish the modification. Do you want to log off now?", "Success", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        ElseIf Language = "Español (España)" Then
            LogOffMessage = MessageBox.Show("¡La corrección se ha aplicado correctamente! Debe cerrar sesión y volver a iniciarla para completar la modificación. ¿Desea cerrar sesión ahora?", "Éxito", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        ElseIf Language = "Français" Then
            LogOffMessage = MessageBox.Show("Le correctif a été appliqué avec succès ! Veuillez vous déconnecter puis reconnecter pour finaliser le correctif. Voulez-vous vous déconnecter maintenant ?", "Succès", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        End If

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
