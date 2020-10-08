$ErrorActionPreference = 'Stop'

if (${Env:ProgramFiles(x86)}) { 
   $PF32 = ${Env:ProgramFiles(x86)} } 
else { $PF32 = $Env:ProgramFiles }
$TeamsInstaller = Join-Path $PF32 'Teams Installer\Teams.exe'

$ShortcutPath = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Microsoft Teams.lnk'

if (Test-Path $TeamsInstaller) {
   # "Update.exe" is how Teams actually normally starts
   $TeamsUpdater = Join-Path $env:LOCALAPPDATA 'Microsoft\Teams\Update.exe'
   if (-not (Test-Path $TeamsUpdater)) {
      # No Teams Update = Teams not installed for this user
      # Lock the starter shortcut so the install doesn't replace it.  This has an 
      #   added benefit of not dropping an icon on the desktop.
      $FileStream = [System.IO.File]::Open($ShortcutPath,'Open','Write')
      try {
         # Install Teams for this user
         Start-Process -FilePath $TeamsInstaller -ArgumentList '--CheckInstall','--source=default' -Wait 
      } catch {
         (New-Object -ComObject Wscript.Shell).popup($Error[0],0,'Install error!',48)
         Return
      } finally {
         # unlock the shortcut (i.e. good housekeeping)
         $FileStream.Close()
         $FileStream.Dispose()
      }
   }
   # Start Teams "normally"
   Start-Process -FilePath $TeamsUpdater -ArgumentList '--processStart "Teams.exe"'
} else {
   (New-Object -ComObject Wscript.Shell).popup('Microsoft Teams is not installed!',0,'Unexpected error!',48)
}
