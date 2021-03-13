__Build Status:__ [![Build status](https://build.powershell.org/guestAuth/app/rest/builds/buildType:(id:PolicyFileEditor_PublishStatusToGitHub)/statusIcon)](https://build.powershell.org/project.html?projectId=PolicyFileEditor&tab=projectOverview&guest=1)

# PolicyFileEditor
PowerShell functions and DSC resource wrappers around the TJX.PolFileEditor.PolFile .NET class.

This is for modifying registry.pol files (Administrative Templates) of local GPOs.  The .NET class code and examples of the original usage can be found at https://gallery.technet.microsoft.com/Read-or-modify-Registrypol-778fed6e .

It was written when I was still very new to both C# and PowerShell, and is pretty ugly / painful to use.  The new functions make this less of a problem, and the DSC resource wrapper around the functions will give us some capability to manage user-specific settings via DSC (something that's come up in discussions on a mailing list recently.)

## Quick start

This example shows you how to install PolicyFileEditor from the gallery and use it to set a mandatory screen saver timout with logon:

```powershell
Write-host "Trusting PS Gallery"
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

Write-Host "Installing PolicyFileEditor V3"
Install-Module -Name PolicyFileEditor -RequiredVersion 3.0.0 -Scope CurrentUser

$UserDir = "$env:windir\system32\GroupPolicy\User\registry.pol"

Write-Host "Setting `Password protect the screen saver` to on"
$RegPath = 'Software\Policies\Microsoft\Windows\Control Panel\Desktop'
$RegName = 'ScreenSaverIsSecure'
$RegData = '1'
$RegType = 'String'
Set-PolicyFileEntry -Path $UserDir -Key $RegPath -ValueName $RegName -Data $RegData -Type $RegType

Write-Host "Setting `Screen saver timeout` to 5m"

$RegPath = 'Software\Policies\Microsoft\Windows\Control Panel\Desktop'
$RegName = 'ScreenSaveTimeOut'
$RegData = '300'
$RegType = 'String'

Set-PolicyFileEntry -Path $UserDir -Key $RegPath -ValueName $RegName -Data $RegData -Type $RegType

# apply the new policy immediately
gpupdate.exe /force
```
