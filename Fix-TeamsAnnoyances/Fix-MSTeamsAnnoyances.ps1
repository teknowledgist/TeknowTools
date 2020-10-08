# This prevents Teams from auto-installing when each user logs in
#   and establishes a Start Menu shortcut to a custom app that 
#   will either install Teams (if not installed for this user) or 
#   start Teams.  It's what MS should have done.
# Custom App compiler at https://github.com/MScholtes/PS2EXE

$ErrorActionPreference = 'Stop'
Start-Transcript -Path "$env:windir\Logs\TeamsFix.log"

$TeamsInstaller = "${env:ProgramFiles(x86)}\Teams Installer\Teams.exe"
$AltStarter = "$env:ProgramData\Organization\Scripts\Start-Teams.exe"

# Create a settings file (if needed) to prevent lauch upon install.
if (Test-Path $TeamsInstaller) {
   $SetupJSON = ($TeamsInstaller -replace 'Teams.exe','setup.json')
   if (-not (Test-Path $SetupJSON)) {
      '{"noAutoStart":"true"}' | Out-File -FilePath $SetupJSON
   }
} else {
   Throw 'Machine-wide install of Teams not found.'
}

# Remove the registry entry to auto-run the Teams installer on login.
if (${Env:ProgramFiles(x86)}) { $X = '\WOW6432Node' } 
$RegPath = "HKLM:\SOFTWARE$X\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $RegPath -Name 'TeamsMachineInstaller' -Force -ErrorAction Continue

# Copy custom app for installing and/or starting teams to a common location
$ScriptDir = "$(Split-Path -parent $Script:MyInvocation.MyCommand.Path)"
$EXE2copy = Get-ChildItem -Path $ScriptDir -Filter '*.exe'
([io.directoryinfo](split-path $AltStarter)).create()
Copy-Item -Path $EXE2copy.FullName -Destination $AltStarter


# Create a shortcut to be pinned to the Start Menu that appears
#   to be Teams, but actuall runs the script
$ShortcutPath = 'C:\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk'
$Shortcut = (New-Object -ComObject 'WScript.Shell').CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $AltStarter
#$Shortcut.IconLocation = 'C:\Program Files (x86)\Teams Installer\Teams.exe,0'
$Shortcut.WorkingDirectory = '%LOCALAPPDATA%\Microsoft\Teams'
$Shortcut.Description = 'Microsoft Teams'
#$Shortcut.windowstyle = 7  # Minimized
$Shortcut.Save()

Stop-Transcript
