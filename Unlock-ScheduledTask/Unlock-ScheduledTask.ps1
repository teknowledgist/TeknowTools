<#
.SYNOPSIS
   Updates the security descriptor for a scheduled task allowing any user to run it. 
.DESCRIPTION
   Earlier versions of Windows apparently used file permissions on 
   C:\Windows\System32\Tasks files to manage security.  Windows 10 now controls
   the ability to run tasks using the Security Descriptor value on tasks under 
   HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree 
   (even though the file permissions still appear to control visibility). 
   This script will grant Users read and execute permissions to the task. 
   Because the registry key is protected, this is intended to be run as SYSTEM, 
   but an admin user can successfully run it too. (It's just messier).
.PARAMETER Taskname   
   The tasks subfolder and name of a scheduled task.  Required.

.EXAMPLE
   Unlock-ScheduledTask.ps1 -Taskname "My task"  
   Unlock-ScheduledTask.ps1 -Taskname "Microsoft\Windows\Defrag\ScheduledDefrag"  

.NOTES
   Inspired/explained by Dave K. (aka MotoX80) on the MS Technet forums:
   https://social.technet.microsoft.com/Forums/windows/en-US/6b9b7ac3-41cd-419e-ac25-c15c45766c8e/scheduled-task-that-any-user-can-run

   This script/information is free: you can redistribute 
   it and/or modify it under the terms of the GNU General Public License 
   as published by the Free Software Foundation, either version 2 of the 
   License, or (at your option) any later version.

   This script is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   The GNU General Public License can be found at <http://www.gnu.org/licenses/>.
#>
Param([Parameter(Mandatory=$true, position=0)][string]$TaskName)

$KeyPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\$TaskName"

$Key = Get-Item $KeyPath -ErrorAction SilentlyContinue

if ($Key) {
   $SecDescHelp = New-Object System.Management.ManagementClass Win32_SecurityDescriptorHelper 
   $oldBinSD = (Get-ItemProperty $Key.name.replace('HKEY_LOCAL_MACHINE','HKLM:')).SD
   $oldSDDL = $SecDescHelp.BinarySDToSDDL( $oldBinSD ) 

   $AdtlSDDL = '(A;ID;0x1301bf;;;BU)'    # add BUILTIN\users read and execute
   #$AdtlSDDL = '(A;ID;0x1301bf;;;AU)'    # add Authenticated users read and execute

   if ($oldSDDL.SDDL -match $AdtlSDDL) {
      Return     # already allowed
   } else { 
      $newSDDL = $oldSDDL.SDDL,$AdtlSDDL -join ''
   }
   $newBinSD = $SecDescHelp.SDDLToBinarySD( $newSDDL )
   [string]$binSDDLStr =  ([BitConverter]::ToString($newBinSD['BinarySD'])).replace('-','') 

   # Only the SYSTEM account can update this registry value
   if ($env:username -eq "$env:computername`$") {
      Set-ItemProperty -Path $KeyPath -Name 'SD' -Value $binSDDLStr -Force
   } else {
      # Not running as SYSTEM, so create a scheduled task and run that as SYSTEM
      $updateTaskName = 'Set-A-Task-Free'
      # A .bat is required because the string is too long to send straight to SchTasks.exe
      "reg.exe add `"$($Key.name)`" /f /v SD /t REG_BINARY /d $binSDDLStr" | 
            Out-File -FilePath "$env:Temp\$updateTaskName.bat" -Encoding ascii -Force
      $Tmrw = (get-date).AddMinutes(-1).AddDays(1)
      # The task will delete itself at the end of the day and will never auto-trigger.
      & SCHTASKS.EXE /CREATE /F /TN "$updateTaskName" /SC DAILY /ST $Tmrw.tostring('HH:mm') /ED $Tmrw.tostring('MM/dd/yyyy') /Z /TR "cmd.exe /c `"$env:Temp\$updateTaskName.bat`"" /RU system 
      # Thus, it needs to be run here.
      & SCHTASKS.EXE /RUN /TN "$updateTaskName"
   }
}

