# Windows 8.1 Profile Picture Login Screen Fix
<sup>made with Visual Basic .NET in Visual Studio 2012 in Windows 8.1</sup>

<img width="100" alt="icon" src="https://github.com/user-attachments/assets/302758c9-194d-4929-b553-d800bc5e11f0" />

This tool fixes the missing profile picture(s) on Windows 8.1’s login screen

## Download
➡️ **Download <code>winbluelsppfix.zip</code> now: <a href="https://github.com/win8-png/winbluelsppfix/releases/latest">here</a>** ⬅️

* The program works on both 32-bit and 64-bit systems and recognises the platform automatically.
* Languages: **English**, **Español (España)**, **Français**

## Description
As you may recognise, on a fully updated Windows 8.1, your profile picture won’t show on the login screen. That means if you start your PC, or want to log in after logging off, even when you just locked your PC, there will be only generic user icon(s) instead of your profile picture, despite having them set and seeing them when the user is logged in. On a multi-user environment, you won’t see any profile picture too. This is what the login screen will look like, even if this user has a profile picture set.

<img width="300" alt="login_screen_no_profile_picture" src="https://github.com/user-attachments/assets/1ea9e660-867f-4837-9357-180098e1aa1c" />

### What should I do now?
That’s where this tool comes in handy, which has been designed and written by me for you nice people out there 😄

It’ll add a missing registry key called “Image200” where the path to the image file for the profile picture on the login screen is being stored. The file itself exist, that means the system creates one, but the link to this file is missing.

This program gives the choice to fix it for your account only, or all users at the same time. It’ll recognise how many and which users have a missing profile picture on the login screen, so that the fix will be applied only to those users.

### Et voilà !
After the fix: **Look at that, much better!** 

<img width="300" alt="login_screen_fixed" src="https://github.com/user-attachments/assets/7dc7c0bd-a284-4007-8d15-20e745172eba" />

## System requirements
* Windows 8.1 32-bit or 64-bit, fully updated
* .NET Framework 4.5.2 or higher<sup>1</sup>
* Administrator privileges

<sub><sup>1</sup> a fully updated Windows 8.1 already contains the mandatory version of .NET Framework</sub>

## Images

<img width="200" alt="winbluelsppfix_screenshot" src="https://github.com/user-attachments/assets/177525fe-755f-49e1-b9da-5716664ad070" />

Version 2.0 of this program

## Something is wrong? ##
* Something doesn’t work right, or the translation is wrong? You are welcome to write in the issues tab! Help is very appreciated 😄
