<#
.SYNOPSIS
   A process and interface for auto-rebooting after updates.
.DESCRIPTION
   When run as a startup script, this will create or modify 
   scheduled tasks to call the script to run again at a later time.
   The scheduled (or standalone) runs will check if Windows is in a pending 
   reboot state.  If so, it will start an escalating process (eventually
   including disabling the network) to encourage the current user to 
   acknowledge the necessity of a reboot.  User acknowledgement causes the 
   script to set a deadline for automatic reboot if the user does not do 
   so themselves.  If no user is logged in, it will automatically reboot 
   to keep the machine fully patched.
.PARAMETER Set
   Usefull for testing script functions.  (This is also used internally 
   for the management -- i.e. disabling and enabling -- of network 
   interfaces.  This internal use requires that this be the first
   parameter.)
   Most values will cause the script to run as if a reboot is pending 
   EXCEPT instead of a reboot, a notification dialog will appear explaining 
   that a reboot would have occurred.
   Special values have specific effects:
      'Start':  (Requires elevation/Admin rights) Runs as if a startup 
            script (i.e. no pending reboot).
            This re-enables NICs and will also pass on any non-default 
            parameters (except these special test parameters) to
            the copy of the script called in the scheduled tasks.  This
            can be combined with other text.   E.G. A value of 
            'starttest' will setup the reboot warning system, and 
            every subsequent time the reboot check runs, it will behave 
            as if there is a pending reboot, but not actually reboot.
      'Clear':  (Requires elevation/Admin rights) All scheduled tasks 
            and created scripts are removed and the NICs are re-enabled.  
            Except for logs the machine should be as if this script were 
            never run -- I.E. an "uninstall".
      'Late':  Disables NICs as if "Patch Tuesday" has passed.
   Default Value: Null
.PARAMETER RebootTime
   The time of day (in 24-hr time) auto-reboots will occur on the weekday defined
   in RebootDOW.
   Default value: 17:00
.PARAMETER RebootDOW
   The day of the week auto-reboots will take place at the time defined
   in the RebootTime.  NOTE: This will be overridden by the default requirement
   to reboot (if pending) prior to "Patch Tuesday".  See the "Relaxed" switch
   to exclude "Patch Tuesday" requirements.
   Default value: Friday
.PARAMETER OrgName
   The short name of the organization.  This is used in several places including
   for separating files and tasks from other applications and system processes.  
   For example, the script and log files will be in the "\ProgramData\OrgName\Reboot" 
   folder, there will be a "\OrgName" task folder, and the HTA window title will
   be "OrgName Security Reboot Notice".
   Default Value: ITServices
.PARAMETER VBScriptName
   Name to use for the Visual Basic Script, polyglot script file that will be
   created by this script and called by the scheduled task.  A polyglot is 
   necessary to run PowerShell code from a scheduled task **without showing 
   a console window**.
   Default Value: CheckAutoReboot.vbs
.PARAMETER TaskName
   The prefix to the name of the scheduled task that will run this script 
   (with modifications) as the user at a later time.  The username is 
   appended to the prefix given here.
   Default Value: 'Check for Pending Reboot - '
.PARAMETER MinLead
   Minimum hours before auto-reboot point user must acknowledge warning or 
   the reboot will be delayed by a week.
   Default Value:  54
.PARAMETER Period
   The default number of hours between checks for pending reboots and/or
   warnings that a reboot is pending/scheduled.  As time expires, warnings
   will occur more frequently.
   Default Value:  4
.PARAMETER Purl
   The url to view the policy named in the Policy parameter.  This will
   open in the user's default browser.
   Default Value:  http://www.school.edu/autorebootpolicy.php
.PARAMETER Policy
   The name or brief description of the policy that requires a computer to
   be up-to-date.  This will follow the text:
         "This computer must reboot to comply with the"
   and be the text for the Purl to allow users to see the policy.
   Default Value:  "Enterprise security requirements."
.PARAMETER Address
   The email address for help or questions about the reboot notice.
   Default Value:  support@school.edu
.PARAMETER Level
   Defines the degree of aggressiveness used to protect a machine 
   on/after Microsoft's "Patch Tuesday" ("PT") when a pending reboot
   was noticed prior to PT.  The four (single letter) options are:
     (N)one:  "Patch Tuesday" is treated as any other day -- i.e. ignored.
     (L)ow:   Network interfaces are disabled IF the notice is not
                yet acknowledged.  Interfaces will re-enable on reboot.
     (M)id:   Network interfaces are disabled.  Will re-enable on reboot.
     (H)igh:  Mid + reboot deadlines that would normally fall after PT 
                are moved back to PT even if that does not provide the 
                normal MinLead warning.
   Note: In some unusual situations, a "High" level could cause a machine
   to reboot immediately upon wake or acknowledgement.  Only network
   interfaces that can ping the domain of the email address will
   be (re-)disabled during every subsequent notification event.
   Default Value: Low
.PARAMETER BlockFile
   The path of the file that, if exists, will prevent a machine from rebooting or
   even checking for a pending reboot.  This should be a location where changes
   require elevated administrator rights.
   Default Value: <SystemDrive>\NoReboot -- e.g. "C:\NoReboot"
.NOTES
   Script version: 3.8

   Copyright 2015-2018 Erich Hammer

   This script/information is free: you can redistribute it and/or modify 
   it under the terms of the GNU General Public License as published by 
   the Free Software Foundation, either version 2 of the License, or (at 
   your option) any later version.

   This script is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   The GNU General Public License can be found at <http://www.gnu.org/licenses/>.
#>
[CmdletBinding()]
Param(
   # Triggers specific functions an/or mimics a pending reboot.
   # This is first to make it easy to pass arguments via the vbscript.
   [ValidateScript({
      if ($_ -imatch '(start|clear)') {
         $principal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
         if ($principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) { $true }
         else {Throw "Testing 'Start' or 'Clear' functions require administrative elevation."}
      } else { $true }
   })]
   [string]$Global:Set = $null,

   # Time of day for auto-reboots to occur
   [ValidateScript({
      If ($_ -match '^([01]\d|2[0-3]):?([0-5]\d)$') { $true }
      else {Throw "`n'$_' is not a time in HH:mm format."}
   })]
   [string]$RebootTime = '17:00',

   # Day of week for auto-reboots to occur
   [ValidateScript({
      $Days = ('Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa')
      If ($Days -icontains ($_[0..1] -join '')) { $true } 
      else {Throw "`n'$_' is not a day of the week or an abbreviation for one."}
   })]
   [string]$RebootDOW = 'Friday',

   # Short name of the organization used for sub-paths
   [string]$OrgName = 'ITServices',

   # Name given to the polyglot, VBS file started by scheduled tasks
   [string]$VBScriptName = 'CheckAutoReboot.vbs',

   # The prefix to the user-owned task that will be scheduled by this script
   [string]$Global:TaskName = 'Check for Pending Reboot - ',

   # Minimum hours before auto-reboot point user must acknowledge warning or it will be delayed
   [int]$MinLead = 54,

   # Default hours between checks/warnings.
   [int]$Period = 4,

   # The URL to the policy
   [string]$Purl = 'http://www.school.edu/autorebootpolicy.php',

   # The name of the security policy requiring reboots after updates
   [string]$Policy = 'Enterprise security requirements',

   # The email address for IT support
   [string]$Address = 'support@school.edu',

   # Level of aggressiveness on "Patch Tuesday"
   [ValidateScript({
      If ($_ -imatch '^[nlmh]') { $true } 
      else {Throw "'$_' does not match a the first letter of 'none', 'low', 'mid', or 'high'."}
   })]
   [string]$Level = 'Low',

   # If present, this machine should not auto-reboot
   [string]$BlockFile = (Join-Path -Path $env:SystemDrive -ChildPath 'NoReboot')
)

#=====================
#region Functions
#=====================
Function Test-RebootPending
{ 
<# 
.SYNOPSIS 
   Tests the pending reboot status of the local computer. 
 
.DESCRIPTION 
   This function will query the registry and determine if the system is pending a reboot. 
   Checks:
   CBServicing = Component Based Servicing (Windows Vista/2008+) 
   WindowsUpdate = Windows Update / Auto Update (Windows 2003+) 
   CCMClientSDK = SCCM 2012 Clients only (DetermineIfRebootPending method) otherwise $null value 
   PendFileRename = PendingFileRenameOperations (Windows 2003+) 
   
.LINK 
    Component-Based Servicing: 
    http://technet.microsoft.com/en-us/library/cc756291(v=WS.10).aspx 
   
    PendingFileRename/Auto Update: 
    http://support.microsoft.com/kb/2723674 
    http://technet.microsoft.com/en-us/library/cc960241.aspx 
    http://blogs.msdn.com/b/hansr/archive/2006/02/17/patchreboot.aspx 
 
    SCCM 2012/CCM_ClientSDK: 
    http://msdn.microsoft.com/en-us/library/jj902723.aspx 
 
.NOTES 
    Inpired by: https://blogs.technet.microsoft.com/heyscriptingguy/2013/06/11/determine-pending-reboot-statuspowershell-style-part-2/
#> 
 
[CmdletBinding()] 
[OutputType([bool])]
param()

Begin { $StartLoc = Get-Location }
Process { 
   $Reason = @()
   
   # Query the Component Based Servicing Reg Key
   Set-Location -Path 'hklm:\SOFTWARE\Microsoft\Windows\CurrentVersion'
   if ((get-item -Path '.\Component Based Servicing').getsubkeynames() -contains 'RebootPending' ) {
      $Reason += 'ComponentBasedServicing' 
   }

   # Query WUAU from the registry 
   if ((get-item -Path '.\WindowsUpdate\Auto Update').getsubkeynames() -contains 'RebootRequired' ) {
      $Reason += 'WindowsUpdate' 
   }
       
   # Query PendingFileRenameOperations from the registry
<# COMMENTED OUT FOR HIGH FREQUENCY/LOW RISK
   Set-Location -Path 'hklm:\SYSTEM\CurrentControlSet\Control\Session Manager'
   if ((get-item .).getvalue('pendingfilerenameoperations')) {
      $Reason += 'PendingFileRename' 
   }
#>
   # Determine SCCM 2012 Client Reboot Pending Status 
   $CCMClientSDK = $null 
   $CCMSplat = @{ 
      NameSpace='ROOT\ccm\ClientSDK' 
      Class='CCM_ClientUtilities' 
      Name='DetermineIfRebootPending' 
      ComputerName=$env:COMPUTERNAME 
      ErrorAction='SilentlyContinue' 
   } 
   $CCMClientSDK = Invoke-WmiMethod @CCMSplat 

   If ($CCMClientSDK.IsHardRebootPending -or $CCMClientSDK.RebootPending) { 
      $Reason += 'SCCMclient' 
   }
   
   $Reason
}#End Process 
End { Set-Location -Path $StartLoc }   

}#End Test-RebootPending Function

Function Get-DisabledNICs {
   # It is faster to first check the Settings file
   if ($Global:Settings.root.NIC.Disabled) {
      $Nodes = @($Global:Settings.root.NIC.Disabled)
      $Props = $Nodes | Get-Member -MemberType property |Select-Object -ExpandProperty name
      ForEach ($Node in $Nodes) {
         $Hash = @{}
         $Props | ForEach-Object {$Hash[$_] = $Node.$_}
         New-Object psobject -property $Hash
      }
   }
   # If no Settings file (possibly deleted by user), check logs for disabled NICs
   if (-not $Nodes) {
      # Look for the last NIC event (1473 = disabled, 1999 = re-enabled)
      #    or 'clear' (0) event.  Removing '0' will cause every reset action to
      #    re-attempt enabling a sole NIC that previously failed enabling.
      $EventFilter = @{ Logname      = 'Application'
                        ProviderName = "$Global:LogSource"
                        Id           = 0,1473,1999
                        }
      $LastDisabledNicEvent = Get-WinEvent -FilterHashtable $EventFilter -MaxEvents 1 -ErrorAction SilentlyContinue

      if ($LastDisabledNicEvent.ID -eq 1473) {
         $Data = $LastDisabledNicEvent.Message.split("`n") | Where-Object {$_ -match ':'} 
         $Props = $Data | ForEach-Object {$_.split(':')[0].trim()} |Select-Object -Unique
         ForEach ($Line in $Data) {
            $null = $Line -Match '(.+?):(.+)'
            If ($Matches[1].trim() -eq $Props[0]) {
               if ($Hash) { $DisabledNICs += New-Object psobject -property $Hash }
               $Hash = @{}
               $Hash = @{$Props[0] = $Matches[2].trim()}
            } Else {
               $Hash[$matches[1]] = $Matches[2].trim()
            }
         }
         New-Object psobject -property $Hash
      }  # END check for a disable event
   } # END if (-not $Nodes)
} #END Get-DisabledNICs

Function Disable-NICs {
<#
.SYNOPSIS 
   Disables all Internet-connected NICs.
 
.DESCRIPTION 
   This function attempts to identify all Network Interface Cards (NICs) that have 
   an Internet (not local) connection and systematically attempts to disable them.  
   It returns a custom object for each NIC containing the name, MAC address,
   description and GUID of the NICs that have been attempted to be disabled.  If 
   the disabling of a particular NIC fails, object will include an 'ERROR' 
   attribute.  If a successfully disabled NIC appears to be a USB device, the 
   object will include a 'USB' attribute.  (This is to allow 
   for warning the user because a disabled NIC cannot be re-enabled if it is
   not present during the enabling attempt.)
.PARAMETER URL
   The address (domain or IP) to test for a network connection.
   Default value is the system NCSI Host (probably www.msftncsi.com)
#>
   [CmdletBinding()]
   Param([string]$URL = ((Get-ItemProperty HKLM:\SYSTEM\CurrentControlSet\Services\NlaSvc\Parameters\Internet).ActiveWebProbeHost))

   # The correct way to work with NICs in Win8+ is to use the NetAdapter module
   if (Get-Command -noun netadapter) {
      $ConnectedNICs = @((Get-NetAdapter -Physical | Where-Object {($_.status -eq 'up') -and ($_.virtual -eq $false)}))
      foreach ($NIC in $ConnectedNICs) {
         $Online = $false
         # Confirm that the NIC has Internet (not just local) network access
         $IPs = $NIC | get-netipaddress -AddressFamily ipv4 | Select-Object -ExpandProperty IPAddress
         :ip1loop foreach ($IP in $IPs) {
            $null = Test-Connection -Source $IP -ComputerName $URL -Count 1 -ErrorAction SilentlyContinue
            if ($?) { 
               $Online = $true
               break ip1loop 
            }
         }
         if ($Online) {
            $Hash = @{
               ID          = $NIC.DeviceID
               Name        = $NIC.Name
               MAC         = $NIC.MacAddress
               Description = $NIC.InterfaceDescription
            }
            # Need to track USB so user can be warned
            if (($NIC.InterfaceDescription -imatch 'USB') -or ($NIC.PnPDeviceID -imatch 'USB')) {
               $Hash.Add('USB',$true)
            }
            try { $NIC | Disable-NetAdapter -Confirm:$false -ErrorAction Stop }
            catch { $Hash.Add('ERROR',$true) }
            Finally { New-Object psobject -Property $Hash }
         }
      }
   } else {
      # The depreciated wmi method for managing NICs is still needed for older systems.
      $filter = "ipenabled='true' AND DNSDomain IS NOT NULL"
      $ConnectedNICs = @((Get-WmiObject win32_networkadapterconfiguration -filter $filter))

      foreach ($NIC in $ConnectedNICs) {
         $Online = $false
         # Confirm that the NIC has Internet (not just local) network access
         :ip2loop foreach ($IP in $NIC.IPAddress) {
            $null = Test-Connection -Source $IP -ComputerName $URL -Count 1 -ErrorAction SilentlyContinue
            if ($?) { 
               $Online = $true
               break ip2loop 
            }
         }
         if ($Online) {
            $Hash = @{
               ID          = $NIC.SettingID
               Name        = $NIC.AdapterType
               MAC         = $NIC.MacAddress
               Description = $NIC.Description
            }
            $wmi = Get-WmiObject win32_networkadapter -filter "guid='$($NIC.SettingID)'"
            # Need to track USB so user can be warned
            if (($NIC.Description -imatch 'USB') -or ($wmi.PnPDeviceID -imatch 'USB')) {
               $Hash.Add('USB',$true)
            }
            $ErrorActionPreference = 'stop'
            try { $null = $wmi.disable() }
            catch { $Hash.Add('ERROR',$true) }
            Finally { 
               $ErrorActionPreference = 'continue'
               New-Object psobject -Property $Hash
            }
         }
      }
   }
} #END of Disable-NICs function

Function Enable-NICs {
<#
.SYNOPSIS 
   Enables NICs by GUID.
 
.DESCRIPTION 
   This function attempts to re-enable Network Interface Cards (NICs) identified
   by GUID.  If a NIC is already enabled, the function passes over it.  If 
   any NICs are not present to be enabled, or cannot be enabled for some other
   reason, the function returns a comma-separated list of their GUIDs.  Otherwise,
   the function returns nothing.
.PARAMETER GUIDList
   A comma-separated list of GUIDs for previously disabled NICs.
#>
   Param([Parameter(Mandatory=$true, position=0)][string[]]$GUIDList)

   $StillOff = @()

   # The correct way to work with NICs in Win8+ is to use the NetAdapter module
   if (Get-Command -noun netadapter) {
      foreach ($GUID in $GUIDList) {
         $NIC = Get-NetAdapter -Physical | Where-Object {($_.Virtual -eq $false) -and ($_.InterfaceGuid -eq $GUID)}
         if ($NIC) {
            if ($NIC.status -eq 'Disabled') {
               try { $NIC | Enable-NetAdapter -Confirm:$false -ErrorAction Stop }
               catch { $StillOff += $GUID }
            }
         } else { $StillOff += $GUID }
      }
   } else {
      # But the depreciated wmi method for managing NICs is still needed for older systems.
      foreach ($GUID in $GUIDList) {
         $ErrorActionPreference = 'stop'
         try { $null = (Get-WmiObject win32_networkadapter -filter "guid='$GUID'").enable() }
         catch { $StillOff += $GUID }
         Finally { $ErrorActionPreference = 'continue' }
      }
   }

   if ($StillOff) {return ,$StillOff}

} #END Enable-NICs function

Function Reset-AutoReboot 
{ 
<#
.SYNOPSIS 
   Resets the auto-reboot settings and optionally powers down or reboots. 
 
.DESCRIPTION 
   This function:
      (1) Re-enables any disabled NICs it thinks were disabled by this script.
      (2) Deletes the task created to prevent logoff.
      (3) Deletes the user-side task(s) for re-running the auto-reboot check.
      (4) Renames the auto-reboot settings file (for tracking). 
      (5) Deletes one-time .hta files in %Temp%.
      (6) Will force Shut down or reboot if called.
#>
   [CmdletBinding()]
   Param([string]$LogTime = (get-date -Format yyyy.MM.dd-HH.mm),
         [string]$Action = $null)

   # (1): Re-enable NICs
   # Do any NICs have a network connection?  (don't reboot on startup if on network)
   $NetworkConnected = $False
   if (Test-Connection $Global:PingAddress -Count 1 -Quiet) {
      $NetworkConnected = $true
   }
   # If NICs were logged as disabled, creating the log event for re-enabling
   #    should trigger the SYSTEM-level task to actually re-enable them.
   if ([array]$DisabledNICs = Get-DisabledNICs) {
      if ($Global:TaskFolder.gettask($Global:NICTaskName).enabled) {
         $EventLogArgs = @{
               EntryType = 'Warning'
               EventID   = 1969   # Trigger network restoration
               Message   = $Global:LogEvents.get_item(1969)
            }
      } else {   # If the task is disabled, then NICs can't be enabled.
         $EventLogArgs = @{
               EntryType = 'Warning'
               EventID   = 1941   # Failed to enable NICs
               Message   = $Global:LogEvents.get_item(1941) + "`n" + $($DisabledNICs |Format-List |out-string)
            }
      }
      Write-EventLog @Global:CommonLogArgs @EventLogArgs
   } 

   # (2 & 3) Delete (user-based) .hta-calling and block-logoff tasks
   #    for elevated users.  Un-elevated users may not be able to delete
   #    their own tasks (why?), so disable them.
   foreach ($Name in ($Global:TaskFolder.gettasks(1) |Select-Object -expandproperty name)) {
      if (($Name -eq $Global:NoLogoffTaskName) -or ($Name -imatch $Global:TaskName)) {
         if ($Global:RunningElevated) {
            $Global:TaskFolder.DeleteTask($Name,0)
         } else {
            $Global:TaskFolder.GetTask($Name).enabled = $False
         }
      }
   }

   # (4) Rename the settings file
   if (Test-Path $Global:SettingsPath) { 
      if (Test-Path (Join-Path (Split-Path $Global:SettingsPath) ('reboot' + $LogTime + '.log'))) {
         Rename-Item -Path $Global:SettingsPath -NewName ('reboot' + (get-date -Format yyyy.MM.dd-HH.mm.ss) + '.log') -Force
      } else {
         Rename-Item -Path $Global:SettingsPath -NewName ('reboot' + $LogTime + '.log') -Force
      }
   }

   # (5) Delete older, one-time .HTA files (keep one for troubleshooting)
   $HTAs = @(Get-ChildItem -Path ($env:TEMP + '\Reboot*.hta') | Sort-Object -Property LastWriteTime -Descending)
   if ($HTAs) {  $HTAs[1..$HTAs.count] | ForEach-Object { Remove-Item $_.Fullname} }

   # (1 redux)
   if ((($Action -eq 'Startup') -and ($DisabledNICs -and -not $NetworkConnected)) -or 
         (($Action -eq 'clear') -and ($DisabledNICs)))  {
      # Machines that are starting up with disabled NICs should be 
      #    rebooted so they can properly log into the domain.
      if ($Action -eq 'Startup') {
         $Action = 'Restart'
      }
   # If necessary, need to give the NIC Manager a few seconds to trigger and run
      $LoopCount = 0
      While ( ($LoopCount -lt 5) -and 
               ($Global:TaskFolder.gettask($Global:NICTaskName)).lastruntime -lt $Now.AddSeconds(-6)) {
         Start-Sleep -Seconds 1
         $LoopCount++
      }
      # Give the (hopefully running) NIC Manager task a few seconds to re-enable NICs
      $LoopCount = 0
      While ($LoopCount -lt 5) {
         $status = $Global:TaskFolder.gettask($Global:NICTaskName).state
         if (($status -eq 2) -or ($status -eq 4)) {   # 2 = "Queued", 4 = "Running"
            Start-Sleep -Seconds 1
            $LoopCount++
         } else {
            $LoopCount = 6
         }
      }
   }

   # (6) Force reboots or shutdown
   if ($Global:Set -and ($Global:Set -inotmatch '(start|clear)')) {
      # Only give a notice if testing
      (new-object -ComObject wscript.shell).popup($Global:TestMsg,0,'Debugging') > $null
      Return $null
   }
   if ($Action -eq 'PowerOff') { Stop-Computer -Force }
   if ($Action -eq 'Restart') { Restart-Computer -Force }

} #End Reset-AutoReboot function

Function Set-NextInterval
{
<#
.SYNOPSIS 
   Determins how long until the auto-reboot notice should return. 
 
.DESCRIPTION 
   This function compares the default time for a return notice with the remaining time
   to a shutdown and gives an increasingly shorter "snooze" time for the notice to return 
   as time is running out. For example, a default period of 4 hours leads to:
      more than 8 hours left -> 4 hour snooze
      4-8 hours left -> 2 hour snooze
      2-4 hours left -> 1 hour snooze
      1-2 hours left -> 30 minute snooze
      .5-1 hour left -> 15 minute snooze
      15-30 minutes left -> 7.5 minute snooze
      7.5-15 minutes left -> 3.75 minutes snooze
      less than 7.5 minutes left -> 1 minute snooze
#>
   [CmdletBinding()]
   Param([int]$DefaultPeriod=4,
         [int]$TotalTimeLeft=480)
   
   $Ratio = $TotalTimeLeft / ($DefaultPeriod * 60)
   If ($Ratio -ge 2) {$NextInterval = $DefaultPeriod*60 }
   ElseIf ($Ratio -ge 1) { $NextInterval = $DefaultPeriod*60/2 }
   ElseIf ($Ratio -ge .5) { $NextInterval = $DefaultPeriod*60/4 }
   ElseIf ($Ratio -ge .25) { $NextInterval = $DefaultPeriod*60/8 }
   ElseIf ($Ratio -ge .125) { $NextInterval = $DefaultPeriod*60/16 }
   ElseIf ($Ratio -ge .0625) { $NextInterval = $DefaultPeriod*60/32 }
   ElseIf ($Ratio -ge .03125) { $NextInterval = $DefaultPeriod*60/64 }
   Else {$NextInterval = 1}
   
   $NextInterval
} #End Set-NextInterval

#=====================
#endregion Functions
#=====================

# Check that this machine can reboot
if (Test-Path $BlockFile) {
   Return
}

#region Initialize strings
#    Some variables are of 'Global' scope because the invoke-command from within VBScript
#    don't always work with 'Script' scope and sub-functions can't reference them.
$ScriptVersion           = '3.8'
$LogonTaskName           = 'Initialize user notices'
$LogonDescription        = 'Starts the process of checking and notifying of the need to reboot.'
$DailyTaskName           = 'Daily pending reboot check'
$DailyDescription        = 'Reboots when no users are logged in if needed.'
$WakeTaskName            = 'Wakeup Reboot Check'
$WakeDescription         = 'Checks for pending reboot on wake if no user is logged in.'
$UserTaskName            = $Global:TaskName + $env:USERNAME
$UserDescription         = 'Periodic check for pending reboot and GUI notice.'
$Global:NoLogoffTaskName = 'No Logoff'
$NoLogoffDescription     = 'Force restart on Logoff'
$Global:NICTaskName      = 'NIC Manager for Reboot Checks'
$NICDescription          = 'Enables or disables network interfaces based on reboot status.'
$XMLfile                 = 'Shutdown.xml'
$Global:LogSource        = "$OrgName Reboot Check"
$Global:LogEvents        = @{   # DON'T CHANGE NUMBERS!
   0    = 'Reboot reminder system uninstalling.'                                             # Information
   1    = 'AutoReboot system startup initialization.'                                        # Information
   2    = "AutoReboot initialization for '$env:USERNAME' login"                              # Information
   8    = 'SYSTEM checking for pending reboot with <ToBeInserted> current user sessions.'    # Information
   13   = 'An unknown error occurred.'                                                       # Error
   42   = "Presenting initial pending reboot notice to '$env:USERNAME'."                     # Information
   60   = "Checking for pending reboot within '$env:USERNAME' logon session."                # Information
   100  = "'$env:USERNAME' agreed: Reboot deadline of <ToBeInserted>."                       # Warning
   101  = "'$env:USERNAME' requested a delay."                                               # Information
   314  = "Reminding '$env:USERNAME' of deadline for scheduled reboot."                      # Information
   666  = "'$env:USERNAME' has requested for no more reminders.`nReboot is still scheduled." # Warning
   806  = 'Pending reboot delayed past "Patch Tuesday". Triggering network disconnection.'   # Warning
   1473 = 'Network interfaces disabled until reboot.'                                        # Warning
   1941 = 'Re-enabling a network interface that was previously disabled has failed.'         # Error
   1969 = 'Triggering a restoration of networking from a disabled state.'                    # Warning
   1999 = 'Restored network interfaces from a disabled state.'                               # Information
   2001 = "'$env:USERNAME' requested a reboot/shutdown."                                     # Information
}
$Global:TestMsg = "You requested a shutdown/restart, but`nsince this is a test, nothing will happen."
$WindowTitle    = "$OrgName Security Reboot Notice"
# Balloon notice (Win7) titles display only the first 62 characters.
# Toast notices (Win8+) show "one line"
$BaloonTitle1   = "$OrgName security policy requires a reboot!"
$BaloonTitle2   = 'Reboot in approximately <ToBeInserted> hours!'
# Balloon notices (Win7) displays only the first 256 characters.  
# Toast notices (Win 8+) show "4 lines" without a specific character count.
$BaloonText     = 'Software patches requiring a reboot were installed.  Please reboot ' +
                  'soon to ensure the security of this system and the entire network.'
$Heading1       = 'REBOOT REQUIRED'
$Heading2       = 'REBOOT SCHEDULED'
$Claim          = "This computer must reboot to comply with the $Policy"
$Countdown      = 'Automatic reboot in <ToBeInserted>'
$Bid            = 'Click a button below to close this notice:'
$LaterText      = 'I accept an automated reboot on <Day>, <Date>`n' +              # where `n = newline
                     'at <Time>.  If I reboot/shutdown beforehand,`n' +            # without them,
                     'the automatic reboot will not occur.'                        # window is very wide
$LaterButton    = 'Remind me later'
$QuietText      = 'Hide these warnings until <Number>hours before automatic`n' +   # `n as above ^
                     'reboot (at <DateTime>).'
$QuietButton    = 'Hold reminders for now'
$NowText        = 'I have saved my work and am ready.'
$NowButton      = 'Reboot now'
$NowPopTitle    = 'Confirm Reboot'
$NowPopText     = 'Have you saved your work?'
$NetOffTitle    = 'Network disabled for protection!'
$NetOffText     = 'Your network connection should be restored when you reboot.'
$ShutdownCheck  = 'Shutdown instead of reboot.'
$ContactText    = 'If you have questions or concerns about this reboot process, please contact <ToBeInserted>'
$base64Image    = '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'
#endregion Initialize strings

# Set paths
$VBScriptPath = Join-Path -Path $env:ProgramData -ChildPath "$OrgName\Reboot\$VBScriptName"
$Global:SettingsPath = Join-Path -Path (Split-Path $VBScriptPath) -ChildPath $XMLfile
$Global:PingAddress = $Address.split('@')[-1]

$Now = Get-Date
# Comprehensive list of methods for obtaining last boot time: 
#   http://www.happysysadm.com/2014/07/windows-boot-time-explored-in-powershell.html
$OSInfo = Get-WmiObject -Class Win32_OperatingSystem
$LastBootTime = [Management.ManagementDateTimeConverter]::ToDateTime($OSInfo.lastbootuptime)
$TimeSinceBoot = New-TimeSpan -Start $LastBootTime -End $Now

# Check for running elevated.
$principal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
$Global:RunningElevated = $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

# Read any existing settings
$Global:Settings = $null
if (Test-Path $Global:SettingsPath) {
   $Global:Settings = [xml](get-content -Path $Global:SettingsPath -ErrorAction SilentlyContinue)
}

# Some common log settings
$Global:CommonLogArgs = @{
   LogName  = 'Application'
   Source   = $Global:LogSource
   Category = 0
}

# Enabling or disabling NICs can only be done elevated and is a 
#    priority because it is triggered by an event set by another instance
#    of the script that may be in the process of shuting down the system.
if ($Global:Set -ceq 'ENABLEorDISABLE') {
   # The 'ENABLEorDISABLE' value for $Set is probably passed as an argument
   #    by a scheduled task triggered by a log entry.  It is unlikely for
   #    this to be called by a user, so there is no check for elevated rights.
   #    If an unprivileged user tries to modify NICs, it just won't work.
   If ([Diagnostics.EventLog]::sourceexists($Global:LogSource)) { 
      $EventFilter = @{ Logname      = 'Application'
                        ProviderName = "$Global:LogSource"
                        Id           = 806,1969    # Disable and Enable events, respectively
                        }
      $LastEventID = (Get-WinEvent -FilterHashtable $EventFilter -MaxEvents 1 -ErrorAction SilentlyContinue).Id
   } else {
      # No log = script never run before = no NICs to enable and pending status is unknown
      return
   }

   # Enable NICs if requested or triggered by "enabling" log entry
   if ($LastEventID -eq 1969) {
      [array]$DisabledNICs = Get-DisabledNICs
      if ($DisabledNICs) {
         $Problems = Enable-NICs ($DisabledNICs | Select-Object -ExpandProperty ID)
         if ($Problems -is [array]) {
            $ProblemNICs = $DisabledNICs | Where-Object {$Problems -contains $_.ID}
         }
         if ($ProblemNICs) {
            $EventLogArgs = @{
               EntryType = 'Error'
               EventID   = 1941
               Message   = $Global:LogEvents.get_item(1941) + "`n" + $($ProblemNICs |Format-List |out-string)
            }
            Write-EventLog @Global:CommonLogArgs @EventLogArgs
         }
         if ($ProblemNICs.count -lt $DisabledNICs.count) {
            $Sucesses = $DisabledNICs | Where-Object {$ProblemNICs -notcontains $_}
            $EventLogArgs = @{
               EntryType = 'Information'
               EventID   = 1999
               Message   = $Global:LogEvents.get_item(1999) + "`n" + $($Sucesses |Format-List |out-string)
            }
            Write-EventLog @Global:CommonLogArgs @EventLogArgs
         }
      } else {
         $EventLogArgs = @{
            EntryType = 'Error'
            EventID   = 13   # unknown error!
            Message   = $Global:LogEvents.get_item(13) + "`n`nNo NICs to enable."
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs
      }

      Return
   } #END checking for enabling
   
   # Disable NICs on all pending reboot checks
   if ($LastEventID -eq 806) {
      [array]$DisabledNICs = Get-DisabledNICs

      [array]$NewlyDisabledNICs = Disable-NICs -URL $Address.split('@')[-1]

      # Only update (with old and new) if there are newly disabled NICs.
      if ($NewlyDisabledNICs) {
         $DisabledNICs = ($DisabledNICs + $NewlyDisabledNICs) | Select-Object -Unique

         # Disabled NICs should always be stored in the logs.
         $EventLogArgs = @{
               EntryType = 'Warning'
               EventID   = 1473   # Disabling NICs
               Message   = $Global:LogEvents.get_item(1473) + "`n" + $($DisabledNICs |Format-List |out-string)
            }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs
      }
      # Store disabled NICs in the settings file only if it already exists.
      if ($Global:Settings) {
         if ($Global:Settings.root.nic.disabled.count -lt $DisabledNICs.count) {
            $Global:Settings.root.RemoveChild($Global:Settings.root.nic)
            $NICelement = $Global:Settings.CreateElement('NIC')
            foreach ($NIC in $DisabledNICs) {
               $DisabledElement = $Global:Settings.CreateElement('Disabled')
               foreach ($property in $NIC.PSObject.Properties) {
                  $attribute = $Global:Settings.CreateAttribute($property.Name)
                  $attribute.value = $property.Value
                  $null = $DisabledElement.Attributes.Append($attribute)
               }
               $null = $NICelement.AppendChild($DisabledElement)
            }
            $null = $Global:Settings.root.AppendChild($NICelement)
            $Global:Settings.Save($Global:SettingsPath)
         }
      } #END if ($Global:Settings)

      Return
   } #END if asked to disable

   # Getting here means something is wrong.
   $EventLogArgs = @{
      EntryType = 'Error'
      EventID   = 13   # unknown error!
      Message   = $Global:LogEvents.get_item(13) + "`n`nSet: $Global:Set"
   }
   Write-EventLog @Global:CommonLogArgs @EventLogArgs
   Return
} #END if ($Global:Set -ceq 'ENABLEorDISABLE')


# Initialize the task service and folder for the rest of the script
$TaskService    = new-object -ComObject('Schedule.Service')
$TaskService.connect()

# Because the Schedule Service can't run PoSh silently, if this process
#    was initiated from a Scheduled task, it probably started as a VBScript.
# The VBScript process (wscript) is of interest to link it to a Task.
$ParentPID = (Get-WmiObject -Class Win32_Process -Filter "ProcessID = $PID").parentprocessid
if ((Get-Process -id $ParentPID).ProcessName -match 'wscript') {
   $PIDtoCheck = $ParentPID
} else { $PIDtoCheck = $PID }

$ThisTask = $TaskService.GetRunningTasks(1) | 
               Where-Object {$_.EnginePID -eq $PIDtoCheck} |
               Select-Object -ExpandProperty Path
If (-not $ThisTask) {
   # Tasks can only be seen by the user/group identified as the "Author".
   # This is an alternate method of determining which task is running, but
   #    the logging of the start of the task occurs a few moments after
   #    the task.
   Start-Sleep -Seconds 3
   # Only look for an entry with the same Process ID from the last minute.
   $XPath = @"
<QueryList>
  <Query>
    <Select>*
      [System/Provider[@Name='microsoft-windows-taskscheduler'] and
      (System/EventID=129) and
      (EventData/Data[@Name="ProcessID"]='$PIDtoCheck') and
      (System/TimeCreated[@SystemTime > '$($now.addseconds(-60).tostring('yyyy-MM-ddTHH:mm:ss.000Z'))'])]
    </Select>
  </Query>
</QueryList>
"@
   $ThisProcessEvent = Get-WinEvent -logname 'Microsoft-Windows-TaskScheduler/Operational' -FilterXPath $XPath -MaxEvents 1 -ErrorAction SilentlyContinue
   if ($ThisProcessEvent) {
      $EventXML = [xml]$ThisProcessEvent.toxml()
      $ThisTask = $eventxml.event.eventdata.data | 
                     Where-Object {$_.name -eq 'taskname'} |
                     Select-Object -ExpandProperty "#text"
   } else {
      $ThisTask = "Unknown" 
   }
}
$ErrorActionPreference = 'stop'
Try { $Global:TaskFolder = $TaskService.GetFolder($OrgName) }
Catch { 
   if ($ThisTask.split('\').count -ge 3) {
      $Global:TaskFolder = $TaskService.GetFolder($ThisTask.substring(0,$ThisTask.LastIndexOf('\')))
   } else {
      $Global:TaskFolder = $TaskService.GetFolder('\')
   }
}
Finally { $ErrorActionPreference = 'continue' }

# Uninstall when called
if ($Global:Set -imatch 'clear') {
   # Reset NICs and user-based tasks
   Reset-AutoReboot -LogTime $Now.ToString('yyyy.MM.dd-HH.mm') -Action 'clear'
   $Global:Settings = $null
   if ($Settings) { Remove-Variable 'Settings' }

   # Remove system-based tasks
   foreach ($Name in ($Global:TaskFolder.gettasks(1) |Select-Object -expandproperty name)) {
      if (($Name -eq $LogonTaskName) -or 
            ($Name -eq $DailyTaskName) -or 
            ($Name -eq $WakeTaskName) -or
            ($Name -eq $Global:NICTaskName)) {
         $Global:TaskFolder.DeleteTask($Name,0)
      }
   }
   # The NIC Manager task is tricky because it may not delete if it is still
   #    trying to re-enable NICs.  Can't wait forever, so log the error.
   if ($Global:TaskFolder.gettasks(1) | Where-Object {$_.name -eq $Global:NICTaskName}) {
      $EventLogArgs = @{
         EntryType = 'Error'
         EventID   = 13   # unknown error!
         Message   = $Global:LogEvents.get_item(13) + "`n`n$Global:NICTaskName"
      }
      Write-EventLog @Global:CommonLogArgs @EventLogArgs
   }

   # Remove the script that runs from tasks
   Remove-Item $VBScriptPath -Force -ErrorAction SilentlyContinue

   $EventLogArgs = @{
      EntryType = 'Information'
      EventID   = 0   # uninstallation of reboot reminder system
      Message   = $Global:LogEvents.get_item(0)
   }
   Write-EventLog @Global:CommonLogArgs @EventLogArgs

   Return
} #END if $Global:Set -match 'clear'

# Check for the newest existing task (which are only created at startup/install).
$TaskCreationTime = $Now.AddDays(-365)  
foreach ($Task in ($Global:TaskFolder.gettasks(1))) {
   if (($Task.Name -eq $LogonTaskName) -or 
          ($Task.Name -eq $DailyTaskName) -or 
          ($Task.Name -eq $WakeTaskName)) {
      $t = $Global:TaskFolder.gettask($Task.Name).definition.registrationInfo.date
      if ($t -gt $TaskCreationTime) { $TaskCreationTime = $t }
   }
}

$RunningAsStartupScript = $false
# If the machine has booted since the tasks were created, then reset everything.  
if ((get-date $TaskCreationTime) -lt $LastBootTime) {
   if ($Global:RunningElevated) {
      $RunningAsStartupScript = $true
   } else {
      if ($Global:TaskFolder.gettasks(1) | Where-Object {$_.Name -ieq $UserTaskName}) {
         if ((($Global:TaskFolder.gettask($UserTaskName)).lastruntime -lt $LastBootTime) -or
               (($Global:TaskFolder.gettask($UserTaskName)).nextruntime -lt $Now)) {
            # Treat it as a login script where the startup script failed
            $RunningAsLoginScript = $true
         } else {
            # There is a scheduled script to run, so let it be.
            $RunningAsLoginScript = $false
         }
      } elseif ($t) {
         # Auto-reboot is installed and some tasks exist, but not for this user
         $RunningAsLoginScript = $true
      } else { 
         (new-object -ComObject wscript.shell).popup('This must be installed with admin rights!',0,'Error!',16) > $null
         Return  # Can't be installed, so exit.
      }
   }
} elseif ($ThisTask.split('\')[-1] -eq $LogonTaskName) {
   # If script is deployed and a user logs in before the machine is rebooted
   $RunningAsLoginScript = $true
}


if ($RunningAsStartupScript -or ($Global:Set -imatch 'start')) {
   # It is useful to know if this is running as a PoSh script vs
   #    running as an invoke-command call in a polyglot.
   #    See:  https://stackoverflow.com/a/43643346
   function Get-ScriptName() { return $MyInvocation.ScriptName }
   if (-not ($ScriptName = Get-ScriptName)) {
      $ScriptName = 'Invoked from VBScript'
   } else {
      # Recompile the Task Scheduler interface as a precaution.  Reference:
      #    https://www.ctrl.blog/entry/idle-task-scheduler-powershell
      & "$env:windir\system32\wbem\mofcomp.exe" ([Environment]::GetFolderPath('System') + '\wbem\SchedProv.mof')
   }

   # As a startup script:
   #   (0) Create an Application EventLog source
   #   (1) Clean up previous reboot notices and enable disabled NICs
   #   (2) Create/update the polyglot, VBS script necessary to run without a console window
   #   (3) Create/update the logon-triggered, scheduled task
   #   (4) Create/update the daily task (for when no user is logged in) 
   #   (5) Create/update the wake-or-boot-triggered task
   #   (6) Create/update the event-triggered task to enable/disabled NICs

   # (0): The Event Logs require a defined source to allow writing to them
   If (-not ([Diagnostics.EventLog]::sourceexists($Global:LogSource))) { 
      New-EventLog -LogName Application -Source $Global:LogSource 
   }
   $EventLogArgs = @{
      EntryType = 'Information'
      EventID   = 1   # Startup initilialization
      Message   = $Global:LogEvents.get_item(1) + "`n$ScriptName`nversion: $ScriptVersion"
   }
   Write-EventLog @Global:CommonLogArgs @EventLogArgs

   # (1): Reset user notification checks and re-enable any disabled NICs
   #      Assuming this is used as a GPO startup script, it will run twice at startup
   #      (once via Scheduled task), so only reset once (to prevent collisions).
   Try {
      $null = [DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
      $SeesDomain = $true
   } Catch {
      $SeesDomain = $false
   }
   if (($SeesDomain -and ($ScriptName -ne 'Invoked from VBScript')) -or 
         (-not $SeesDomain) -or ($Global:Set -imatch 'start')) {
      Reset-AutoReboot -LogTime $LastBootTime.ToString('yyyy.MM.dd-HH.mm') -Action 'Startup'
      $Global:Settings = $null
      if ($Settings) { Remove-Variable 'Settings' }
   }

   # (2): embed all PoSh code into a VBS "shell"
   # Read itself if run as PoSh script (and not polyglot)
   if ($ScriptName -ne 'Invoked from VBScript') {
      $PoShCode = Get-Content -Path $MyInvocation.MyCommand.Definition

      # If command-line switches are used for a startup script, they
      #    should be passed along (except for a 'start') as the 
      #    defaults of the embedded script.
      $ReplacementsHash = @{
         "(^\s*\[string]\`$RebootTime =).*,"      = "`$1 '$RebootTime',"
         "(^\s*\[string]\`$RebootDOW =).*,"       = "`$1 '$RebootDOW',"
         "(^\s*\[string]\`$BlockFile =).*,"       = "`$1 '$BlockFile',"
         "(^\s*\[string]\`$OrgName =).*,"         = "`$1 '$OrgName',"
         "(^\s*\[string]\`$VBScriptName =).*,"    = "`$1 '$VBScriptName',"
         "(^\s*\[string]\`$Global:TaskName =).*," = "`$1 '$Global:TaskName',"
         "(^\s*\[int]\`$MinLead =).*,"            = "`$1 $MinLead,"
         "(^\s*\[int]\`$Period =).*,"             = "`$1 $Period,"
         "(^\s*\[string]\`$Purl =).*,"            = "`$1 '$Purl',"
         "(^\s*\[string]\`$Policy =).*,"          = "`$1 '$Policy',"
         "(^\s*\[string]\`$Address =).*,"         = "`$1 '$Address',"
         "(^\s*\[string]\`$Level =).*,"           = "`$1 '$Level',"
      }
      $Replacement = $Global:Set -replace 'start',''
      if ($Replacement -ne '') {
         $ReplacementsHash.Add("(^\s*\[string]\`$Global:Set =).*","`$1 '$Replacement'")
      }

      # This will turn all the PoSh code into VBScript comments, going
      #   through each line of code only once, and each replacement 
      #   only until all of them have been used.
      # 
      $ModCode = @()
      $count = 0
      foreach ($line in $PoShCode) {
         if ($count -lt $ReplacementsHash.count) {
            $NotMatched = $true
            :hashloop foreach ($item in $ReplacementsHash.GetEnumerator()) {
               if ($line -imatch $item.Name) {
                  $ModCode += "'" + ($line -replace $item.Name,$item.Value)
                  $NotMatched = $False
                  $count++
                  break hashloop
               }
            }
            if ($NotMatched) {
               $ModCode += "'" + $line
            }
         } else {
            $ModCode += "'" + $line
         }
      }

      $VBwrapper = @"
'  This "polyglot script" is a PowerShell script embedded in a
'  Visual Basic Script file.  All PowerShell code must both exist 
'  between the # START and # END lines and have a ' (quote mark) 
'  as the first char of the line (to exclude it from VBS parsing).
'
'  Arguments for the PowerShell script can be passed in definition
'  order by separating them with a space, quoting strings with 
'  spaces, and NOT using parameter names.  This is not well tested.
'
'  To see the PowerShell window and any errors, add " -noexit" to 
'  the start of the "PoShswitches" variable defined near the end of
'  the VBscript code *AND* change "0" to "1" in the RUN command
'  a few lines later.
' =====================================================================
'# Start PowerShell  # (Don't modify this line!)

#ReplaceThisLine

'# End PowerShell  (Don't modify this line!)

' Uncomment for testing delay to bring other windows forward
'WScript.Sleep(3000)

' Minimize impact on "No Reboot" machines.
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("<BlockFile>") Then
   WScript.Quit
End If

Dim WinScriptHost, PoShswitches, PoShcmd, arrArgs

PoShswitches = " -ExecutionPolicy ByPass -noprofile -Command " 

PoShcmd = "`$String = ([System.IO.File]::ReadAllText('" & Wscript.ScriptFullName & "')); "  'Read this file.
PoShcmd = PoShcmd & "`$String = `$String.remove("                                           'trim VBscript
PoShcmd = PoShcmd & "`$String.indexof('# End PowerShell',`$String.indexof('#END Script'))"  'from here to end
PoShcmd = PoShcmd & ").remove(0,`$String.indexof('# Start PowerShell')) "                   'and from start to PoSh.
PoShcmd = PoShcmd & "-replace '(?s)(\n)\x27','`$1';"                                        'remove VB comment marks.
'PoShcmd = PoShcmd & "`$String |Out-file `$env:TEMP\RebootPoSh.txt ; "                      '(opt) line for bug check
PoShcmd = PoShcmd & "Invoke-Command ([scriptblock]::Create( `$String )) "                   'Establish scriptblock
If WScript.Arguments.Count > 0 Then 
   ReDim arrArgs(WScript.Arguments.Count-1)
   For i = 0 To WScript.Arguments.Count-1
      arrArgs(i) = WScript.Arguments(i)
   Next
   PoShcmd = PoShcmd & "-ArgumentList " & """" & join(arrArgs, """,""") & """"                   'and add arguments.
End If

Set WinScriptHost = CreateObject("WScript.Shell") 
WinScriptHost.Run "powershell.exe" & PoShswitches & CHR(34) & PoShcmd & Chr(34), 0, TRUE 

Set WinScriptHost = Nothing 
"@
      $VBwrapper = $VBwrapper -replace '<BlockFile>',$BlockFile
      $VBpolyglot = $VBwrapper.split("`n") | 
                        ForEach-Object {
                           if ($_ -cmatch '^#ReplaceThisLine') {
                              $ModCode
                           } else {
                              $_
                           }
                        }
      $VBdir   = Split-Path -Path $VBScriptPath
      if (-not (Test-Path -path $VBdir)) {$null = New-Item -Path $VBdir -ItemType Directory}
      Out-File -InputObject $VBpolyglot -FilePath $VBScriptPath -Force
   } #END if $ScriptName

   # (3): A scheduled task triggered by user logon.
   $LogonTask_xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(get-date -Format yyyy-MM-ddTHH:mm:ss.00000)</Date>
    <Author>Interactive</Author>
    <Description>$LogonDescription</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>PT1H</Delay>
      <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <GroupId>S-1-5-32-545</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$VBScriptPath</Command>
    </Exec>
  </Actions>
</Task>
"@

   # (4): A daily scheduled task that will wake the machine (to reboot if pending and no user is logged in).
   # This was intended to run when a machine was idle for an extended ammount of time, but Microsoft has broken
   #   idle triggers in Windows 10 (see:  https://www.ctrl.blog/entry/idle-task-scheduler-powershell).  Therefore,
   #   this daily trigger is the compromise.
   $DailyTask_xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(get-date -Format yyyy-MM-ddTHH:mm:ss.00000)</Date>
    <Author>Administrators</Author>
    <Description>$DailyDescription</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$(get-date -Format yyyy-MM-ddTHH:mm:ss)</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>NT AUTHORITY\SYSTEM</UserId>
      <LogonType>S4U</LogonType>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$VBScriptPath</Command>
    </Exec>
  </Actions>
</Task>
"@

   # (5): A scheduled task for when the machine wakes (to reboot if pending before a user logs in)
   #      or when the machine boots while disconnected from the domain (i.e. no GPO to run script)
   $WakeTask_xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(get-date -Format yyyy-MM-ddTHH:mm:ss.00000)</Date>
    <Author>Administrators</Author>
    <Description>$WakeDescription</Description>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
    <BootTrigger>
      <Enabled>true</Enabled>
    </BootTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>NT AUTHORITY\SYSTEM</UserId>
      <LogonType>S4U</LogonType>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$VBScriptPath</Command>
    </Exec>
  </Actions>
</Task>
"@

   # (6): A high-priority task for when the network needs to be enabled or disabled
   $NICTask_xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(get-date -Format yyyy-MM-ddTHH:mm:ss.00000)</Date>
    <Author>Administrators</Author>
    <Description>$NICDescription</Description>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Application"&gt;&lt;Select Path="Application"&gt;*[System[Provider[@Name='$Global:LogSource'] and (Level=3) and (EventID=806 or EventID=1969)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>NT AUTHORITY\SYSTEM</UserId>
      <LogonType>S4U</LogonType>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>false</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1M</ExecutionTimeLimit>
    <Priority>2</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$VBScriptPath</Command>
      <Arguments>ENABLEorDISABLE</Arguments>
    </Exec>
  </Actions>
</Task>
"@

   # Create the task folder if necessary
   $ErrorActionPreference = 'stop'
   Try {$Global:TaskFolder = $TaskService.GetFolder($OrgName)}
   Catch { 
      $rootFolder = $TaskService.GetFolder('\') 
      if ($ThisTask -and ($ThisTask.split('\').count -ge 3)) {
         $otherFolder = $TaskService.GetFolder($ThisTask.substring(0,$ThisTask.LastIndexOf('\')))
      }
      $null = $rootFolder.CreateFolder($OrgName) 
      $Global:TaskFolder = $rootFolder.GetFolder($OrgName)
   }
   Finally { $ErrorActionPreference = 'continue' }

   # Delete possibly outdated tasks
   foreach ($Name in ($Global:TaskFolder.gettasks(1) |Select-Object -expandproperty name)) {
      if (($Name -eq $LogonTaskName) -or 
            ($Name -eq $DailyTaskName) -or 
            ($Name -eq $WakeTaskName) -or 
            ($Name -eq $Global:NICTaskName)) {
         $Global:TaskFolder.DeleteTask($Name,0)
      }
   }
   # Delete really old possible tasks from previous version
   foreach ($oldFolder in ($rootFolder,$otherFolder)) {
      if ($oldFolder) {
         foreach ($Name in ($oldFolder.gettasks(1) |Select-Object -expandproperty name)) {
            if (($Name -eq $LogonTaskName) -or 
                ($Name -eq $DailyTaskName) -or 
                ($Name -eq $WakeTaskName)) {
               $oldFolder.DeleteTask($Name,0)
            }
         }
      }
      Remove-Variable rootFolder,otherfolder -ErrorAction SilentlyContinue
   }


   $Task = $TaskService.NewTask($null)
   $task.XmlText = $LogonTask_xml
   $null = $Global:TaskFolder.RegisterTaskDefinition($LogonTaskName, $Task, 6, $null, $null, 3)

   $Task = $TaskService.NewTask($null)
   $task.XmlText = $DailyTask_xml
   $null = $Global:TaskFolder.RegisterTaskDefinition($DailyTaskName, $Task, 6, $null, $null, 3)

   $Task = $TaskService.NewTask($null)
   $task.XmlText = $WakeTask_xml
   $null = $Global:TaskFolder.RegisterTaskDefinition($WakeTaskName, $Task, 6, $null, $null, 3)

   # No need for NIC management if security level is "none"
   if ($Level[0] -ine 'n') {
      $Task = $TaskService.NewTask($null)
      $task.XmlText = $NICTask_xml
      $null = $Global:TaskFolder.RegisterTaskDefinition($Global:NICTaskName, $Task, 6, $null, $null, 3)
      # When re-enabling NICs, it is important for this script running under unprivileged
      #    user accounts to be able to see some status information about the NIC Management task.
      $TaskFile = Join-Path $env:SystemRoot "System32\Tasks\$($Global:TaskFolder.path)\$Global:NICTaskName"
      if (Test-Path $TaskFile) {
         $Acl = get-acl -Path $TaskFile
         $rule = New-Object -TypeName system.security.accesscontrol.filesystemaccessrule -ArgumentList ('Authenticated Users','Read','Allow')
         $Acl.setaccessrule($rule)
         set-acl -Path $TaskFile -AclObject $Acl
      }
   }

   # As this is running at startup, don't even check for a pending reboot
   Return

} else {   # NOT running at startup
   # Use sessions to check for logged in users
   $qwinstaOut = & "$env:windir\system32\QWinSta.exe"
   if (-not $qwinstaOut) {
      $UserSessions = 'unknown'
   } else {
      $UserSessions = $qwinstaOut[1..($qwinstaOut.count-1)] | 
                           ForEach-Object {$_.substring(19)} | 
                           Where-Object {$_ -match '^\w'}
   }

   if (-not $UserSessions) { 
      # This must be running as SYSTEM
      $EventLogArgs = @{
         EntryType = 'Information'
         EventID   = 8   # Checking for pending reboot outside of any user session..
         Message   = ($Global:LogEvents.get_item(8) -replace '<ToBeInserted>',$UserSessions.count) + "`n`nTask:  $ThisTask"
      }
      Write-EventLog @Global:CommonLogArgs @EventLogArgs

      if (Test-RebootPending) {
         # A pending reboot with nobody logged in should proceed to reboot
         Reset-AutoReboot -Action 'Restart'
         return
      } else { 
         return  # nothing to see here
      }
   } else {  # Users ARE logged in
      # Is this running as a daily or wake task?
      if (($ThisTask.split('\')[-1] -eq $WakeTaskName) -or ($ThisTask.split('\')[-1] -eq $DailyTaskName)) {
         $EventLogArgs = @{
            EntryType = 'Information'
            EventID   = 8   # Checking for pending reboot outside of any user session.
            Message   = ($Global:LogEvents.get_item(8) -replace '<ToBeInserted>',$UserSessions.count) + "`n`nTask:  $ThisTask"
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs
      } elseif ($RunningAsLoginScript) {
         $EventLogArgs = @{
            EntryType = 'Information'
            EventID   = 2   # Initialization for user login
            Message   = $Global:LogEvents.get_item(2)
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs

         # Reset user-based settings before checking for pending reboot
         Reset-AutoReboot -LogTime $LastBootTime.ToString('yyyy.MM.dd-HH.mm')
         $Global:Settings = $null
         Remove-Variable 'Settings'

      } else {
         $EventLogArgs = @{
            EntryType = 'Information'
            EventID   = 60   # Scheduled check for pending reboot within user logon session.
            Message   = $Global:LogEvents.get_item(60) + "`n`nTask:  $ThisTask"
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs
      }
   } #END Else (i.e. users are logged in)
} #END Else # not running at startup

# Why we are here.
$Reason = Test-RebootPending

if ($Global:Set) { $Reason = "Testing: '$Global:Set'" }

if ($Reason) {
   # When a reboot is pending, force reboot on logoff attempt.
   if (($Global:TaskFolder.gettasks(1) |Select-Object -expandproperty name) -notcontains $Global:NoLogoffTaskName) {
      $TaskDef = $TaskService.NewTask(0)  # Not sure what the "0" is for
      $Taskdef.RegistrationInfo.Description = $NoLogoffDescription
      $TaskDef.RegistrationInfo.Date = $Now.ToString('yyyy-MM-ddTHH:mm:ss.00000')
      $TaskDef.settings.priority = 2
      $TaskDef.settings.StartWhenAvailable = $true
      $Trigger = $Taskdef.Triggers.Create(0) 
      # The Event definition string is extracted from a sample task created in the GUI, 
      #   exported to XML, and with the html-characters corrected.
      #   Could just copy from XML and use:  .replace("&lt;","<").replace("&gt;",">")
      $Trigger.subscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">' +
                              '*[System[Provider[@Name="Microsoft-Windows-Winlogon"] and EventID=7002]]' +
                              '</Select></Query></QueryList>'
      $Action = $Taskdef.Actions.Create(0)
      $Action.Path = 'shutdown.exe'
      $Action.Arguments = '/r /f'
      # Finally, register the task
      $Global:TaskFolder.RegisterTaskDefinition($Global:NoLogoffTaskName, $Taskdef, 6, $null, $null, 3) > $null
   } else {
      $Global:TaskFolder.GetTask($Global:NoLogoffTaskName).enabled = $true
   }
}

# Don't give notice unless there is a pending reboot 
#   AND the machine hasn't rebooted in 24 hours.
if (($Reason -and ($TimeSinceBoot.TotalHours -ge 24)) -or ($Global:Set)) {

   # Want to know when the pending reboot was first identified.
   # This will be either now, or the log entry timestamp for the first user notification after 
   #    the last reboot (or initialization).
   $EventFilter = @{ Logname      = 'Application'
                     ProviderName = "$Global:LogSource"
                     Id           = 1   # Auto-reboot initialized
                     StartTime    = $LastBootTime
                     }
   $LastInitialization = Get-WinEvent -FilterHashtable $EventFilter -MaxEvents 1 -ErrorAction SilentlyContinue |
                           Select-Object -ExpandProperty TimeCreated

   if (-not $LastInitialization) {
      $LastInitialization = $LastBootTime
   }
   $EventFilter = @{ Logname      = 'Application'
                     ProviderName = "$Global:LogSource"
                     Id           = 42   # Initial notice to user
                     StartTime    = $LastInitialization
                     }
   $FlagDate = Get-WinEvent -FilterHashtable $EventFilter -MaxEvents 1 -Oldest -ErrorAction SilentlyContinue |
                        Select-Object -ExpandProperty TimeCreated
   If (-not $FlagDate) {
      $FlagDate = $Now
   }

   # Check if "Patch Tuesday" ("PT") matters.
   if ($Level[0] -ine 'n') {
      # If any PT is important, it is the one that follows the noticed 
      #    pending reboot.  PT is always in the week containing the 
      #    12th day of the month.
      $BaseDate = Get-Date -Day 12 -month $Flagdate.month   
      $PatchTuesday = $BaseDate.AddDays(2 - [int]$BaseDate.DayOfWeek)
      # Was the pending reboot noticed before or after this month's PT?
      If ($FlagDate -gt $PatchTuesday) { 
         $BaseDate = $BaseDate.AddMonths(1)
         $PatchTuesday = $BaseDate.AddDays(2 - [int]$BaseDate.DayOfWeek)
      }
      If ($Global:Set -imatch 'late') {
         $PatchTuesday = $PatchTuesday.AddMonths(-2)
      }
   } else {
      # Ignore PT - i.e. set it to a far future date
      $PatchTuesday = Get-Date ($Now.AddYears(1))
   }

   if ($Now -gt $PatchTuesday) {
      if ((($Level[0] -ieq 'l') -and -not $Global:Settings) -or 
          ($Level[0] -imatch '^[mh]') -or 
          ($Global:Set -imatch 'late')
         ) {
         $EventLogArgs = @{
            EntryType = 'Warning'
            EventID   = 806   # Entry that should trigger network disabling
            Message   = $Global:LogEvents.get_item(806)
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs
      } #END check on security level
   } #END if after Patch Tuesday

   # The "RebootPoint" (and Settings file) don't exist until a user acknowledges.
   if ($Global:Settings.Root.RebootPoint) {
      $RebootPoint = get-date -Date $Global:Settings.Root.RebootPoint
      $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes

      # Reboot immediately if the user has acknowledged and time is up
      If ($Remaining -le 0) {
         If ($Global:Settings.Root.Shutdown -eq 'Checked') {
            Reset-AutoReboot -Action 'PowerOff'
            Return
         } else {
            Reset-AutoReboot -Action 'Restart'
            Return
         }
      }
   } else {  # I.E. User has not acknowledged yet (or deleted the settings file).
      If ($Global:Settings) {
         # Something went wrong -- don't reboot.
         Reset-AutoReboot
         $Global:Settings = $null
      }

      # The *potential* time of reboot is shown in the initial GUI
      $RebootPoint = 0..6 | ForEach-Object {$(Get-Date -Date $RebootTime).AddDays($_)} | 
                           Where-Object {-not (Compare-Object $_.DayOfWeek.tostring().tolower()[0..1] $RebootDOW.tolower()[0..1])}
      $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes

      # delay the reboot for a week if the time is already too late
      if ($Remaining -lt $MinLead*60) {
         $RebootPoint = $RebootPoint.AddDays(7)
         $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes
      }
      if ($Level[0] -ieq 'h') { 
         If ($RebootPoint -gt $PatchTuesday) {
            $RebootPoint = $PatchTuesday
            $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes
         }
      }

   } #End Else (for no acknowledgment yet)

   # No point displaying notice in non-user tasks
   if (($ThisTask.split('\')[-1] -eq $WakeTaskName) -or ($ThisTask.split('\')[-1] -eq $DailyTaskName)) { Return }
   
   # Users who:
   #      have not requested to not be reminded, 
   #      or are in the final period (i.e. time is running out),
   #      or are experiencing a disabled network due to Patch Tuesday security risks
   #    will get a notification of the pending reboot status.
   if (($Global:Settings.Root.quiet -ne 'True') -or 
       ($Remaining -le ($Period*60)) -or
       ($Global:Settings.root.NIC.Disabled)) {

      if ($Global:Settings) {
         $EventLogArgs = @{
            EntryType = 'Information'
            EventID   = 314   # Reminding of deadline to which the user agreed
            Message   = $Global:LogEvents.get_item(314) + "`n`nWhy: $Reason" + "`n`nTask:  $ThisTask"
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs

         $Heading = $Heading2
         $ShowCountdown = 'inline'

         if ($Global:Settings.root.NIC.Disabled) {
            $ShowQuiet = 'none'
            $NetOff_VBbool = 'True'
         } else {
            $ShowQuiet = 'inline'
            $NetOff_VBbool = 'False'
         }
      } else {
         $EventLogArgs = @{
            EntryType = 'Information'
            EventID   = 42   # Presenting initial reboot notice
            Message   = $Global:LogEvents.get_item(42) + "`n`nWhy: $Reason" + "`n`nTask:  $ThisTask"
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs

         $Heading = $Heading1
         $ShowQuiet = $ShowCountdown = 'none'
         $NetOff_VBbool = 'False'
      }

      # The following here-string is exported to a unique HTA that will be saved in a temporary
      #   location for one-time use.  The HTA provides a GUI, and creates/updates the 
      #   XML settings file with information on the choices the user makes.  
      $HTACode = @"
   <html xmlns:IE>
   <!-- http://www.itsupportguides.com/windows-7/windows-7-shutdown-message-with-countdown-and-cancel/ -->
   <title>$WindowTitle</title>
   <script type="text/javascript">
     // Quickly move the window to the approximate location
     window.resizeTo(500,700)
     window.moveTo(screen.availWidth-500,screen.availHeight-700);
   </script>
   <HTA:APPLICATION
     BORDERSTYLE            = "Normal"
     CAPTION                = "Yes"
     CONTEXTMENU            = "No"
     INNERBORDER            = "Yes"
     MAXIMIZEBUTTON         = "No"
     MINIMIZEBUTTON         = "No"
     NAVIGABLE              = "No"
     SCROLL                 = "No"
     SCROLLFLAT             = "Yes"
     SELECTION              = "No"
     SHOWINTASKBAR          = "No"
     SINGLEINSTANCE         = "Yes"
     SYSMENU                = "No"
   >
   <head>
   <STYLE>
   html, body {
       position:fixed;
       border:0;
       top:0;
       bottom:0;
       left:0;
       right:0;
       font-family: Verdana,Arial,Helvetica,sans-serif;
       font-size: 87%;
       line-height: 1.5em;
       margin: 0;
       padding: 0;
       background-color: #FCDD00;
       text-align: center;

   #top_body {
       clear: both;
       margin: 20px 0 0;
   }
   #top_body p {
       font-style: italic;
       margin: 20px 0 0;
   }
   #countdown {
       color: #FF0000;
       font-size: 1.4em;
       font-weight: bold;
   }
   #LinkSpan { cursor: pointer; }
   img {
       width: 25%;
       height: auto;
   }
   h1 {
       color: #2B165E;
       font-weight: bold;
       text-transform: uppercase;
       white-space: nowrap;
   }
   hr {
     width: 75%;
     color: #2B165E;
     height:  1px;
   } 
   button {
       font-family: tahoma;
       font-size: 1.4em;
       font-weight: 500;
       overflow:visible;
       padding:.1em,.5em;
       margin: .3em;
   }

   </STYLE>

   <script language='vbscript'>
   dim ShutdownFilePath, strShutdown, idTimer
   ShutdownFilePath = "$Global:SettingsPath"

   Sub window_onLoad
     idTimer = window.setTimeout("close()", 0.99*1000*60*$(Set-NextInterval $Period $Remaining), "VBScript")
     window.focus()
     SetShutdown
   End Sub

   Sub Beep
     BeepSound = chr(007)
     CreateObject("WScript.Shell").Run "cmd /c @echo " & BeepSound, 0
   End Sub

   sub SetShutdown 
     if(document.getElementById("ShutdownCheck").checked) Then
       strShutdown = "Checked"
     else
       strShutdown = ""
     End If
   End Sub

   Sub MakePrettyXML(strFileName)
      Set objFSO=CreateObject("Scripting.FileSystemObject")
      Set objXMLFile = objFSO.OpenTextFile(strFileName,1,False,-2)
      strXML = objXMLFile.ReadAll
      strXML = Replace(strXML,"><",">" & vbCrLf & "<")
      objXMLFile.Close

      Set objXMLFile = objFSO.CreateTextFile(strFileName,True,False)
      objXMLFile.Write strXML
      objXMLFile.Close

      strStylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
                        "<xsl:output method=""xml"" indent=""yes""/>" & _
                        "<xsl:template match=""/"">" & _
                        "<xsl:copy-of select="".""/>" & _
                        "</xsl:template></xsl:stylesheet>"

      Set objXSL=CreateObject("Msxml2.DOMDocument")
      objXSL.loadXML strStylesheet

      Set objXML=CreateObject("Msxml2.DOMDocument")
      objXML.load strFileName
      objXML.transformNode objXSL
      objXML.save strFileName
   End Sub

   Sub RecordSettings(strChoice)
     Set fso = CreateObject("Scripting.FileSystemObject")
     Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")
     xmlDoc.setProperty "SelectionLanguage", "XPath"
  
     If strChoice = "GoAhead" Then
       iAnswer = CreateObject("Wscript.Shell").Popup( _
                 "$NowPopText", _
                 60, _
                 "$NowPopTitle", _
                 vbYesNo + vbQuestion + vbSystemModal)
       if iAnswer <> vbYes Then
         exit sub
       End If
     End If
     If ((strChoice = "Snooze") And ($NetOff_VBbool)) Then
       iAnswer = CreateObject("Wscript.Shell").Popup( _
                 "$NetOffText", _
                 60, _
                 "$NetOffTitle", _
                 vbOKOnly + vbExclamation + vbSystemModal)
     End If
  
     If (fso.FileExists(ShutdownFilePath)) Then
       xmlDoc.Async = "False"
       xmlDoc.load ShutdownFilePath
       Set objRoot = xmlDoc.documentElement
    
       Set objReason = xmlDoc.selectSingleNode("//RebootReason")
       objReason.text = "$($Reason -join ',')"

       Set Node = xmlDoc.selectSingleNode("//Acknowledgements")
         Set objStamp = xmlDoc.createElement("Stamp")
         objStamp.text = Now
         Node.appendChild objStamp
    
       Set Node = xmlDoc.selectSingleNode("//Shutdown")
       Node.text = strShutdown
    
       If strChoice = "GoAhead" Or strChoice = "Auto" Then
         Set objRecord = xmlDoc.createElement("ActNow")
         objRecord.text = "True"
         objRoot.appendChild objRecord
       ElseIf strChoice = "OKquiet" Then
         Set Node = xmlDoc.selectSingleNode("//Quiet")
         Node.text = "True"
       End If
     Else  
       NewXml = True
       xmlDoc.loadXML "<Root>" + vbcr + "</Root>"
       Set objRoot = xmlDoc.documentElement

       Set objReason = xmlDoc.createElement("RebootReason")
       objReason.text = "$($Reason -join ',')"
       objRoot.appendChild objReason

       Set objReason = xmlDoc.createElement("TriggerPoint")
       objReason.text = "$($FlagDate.ToString())"
       objRoot.appendChild objReason

       Set objReason = xmlDoc.createElement("RebootPoint")
       objReason.text = "$($RebootPoint.ToString())"
       objRoot.appendChild objReason

       Set objRecord = xmlDoc.createElement("Acknowledgements")
       objRoot.appendChild objRecord
         Set objStamp = xmlDoc.createElement("Stamp")
         objStamp.text = Now
         Set objMark = xmlDoc.createAttribute("Mark")
         objMark.Value = 0
         objStamp.attributes.setNamedItem(objMark)
         objRecord.appendChild objStamp

       Set objRecord = xmlDoc.createElement("Shutdown")
       objRecord.text = strShutdown
       objRoot.appendChild objRecord

       If strChoice = "Snooze" Then
         Set objRecord = xmlDoc.createElement("Quiet")
         objRecord.text = "False"
       ElseIf strChoice = "OKquiet" Then
         Set objRecord = xmlDoc.createElement("Quiet")
         objRecord.text = "True"
       Else 
         Set objRecord = xmlDoc.createElement("ActNow")
         objRecord.text = "True"
       End If
       objRoot.appendChild objRecord

       Set objIntro = xmlDoc.createProcessingInstruction ("xml","version='1.0'")  
       xmlDoc.insertBefore objIntro,xmlDoc.childNodes(0)  
     End If 
     xmlDoc.save ShutdownFilePath

     MakePrettyXML ShutdownFilePath

     close()
   End Sub

   </script>
   </head>
   <body>
   <div id="Content" style="zoom:1">
     <p><img src="data:image/png;base64,$base64Image"/>
     </p>

     <span id="Notice"><h1>$Heading</h1></span>
     <div id="top_body">
       <p><b>$($Claim -replace $Policy,'')<br />
         <a>
           <span id="LinkSpan" onClick="OpenURL()">
             <u>$Policy</u>.
           </span>
         </a></b>
       </p>
       <span style="display:$ShowCountdown">
          <p>$($Countdown -replace '<.*>','')</p>
          <div id="countdown">Unknown. Error.</div><br/>
       </span>
     </div>
     <p>
       $Bid<br/>
       <hr style="height:3px" />
       <span ID="TooLate">
         $($LaterText -replace '`n','<br />' -replace '<Day>',$RebootDOW -replace '<Date>',(get-date -Date $RebootPoint -UFormat %x) -replace '<Time>',(get-date -Date $RebootPoint -Format t) )
         <button type='button' id='SnoozeButton' onclick='RecordSettings("Snooze")' autofocus>$LaterButton</button><br />
         <hr />
         <span ID="NoReminder" style="display:$ShowQuiet">
           $($QuietText -replace '`n','<br />' -replace '<Number>','<span id="FinalPeriod"></span>&nbsp;' -replace '<DateTime>',(get-date -Date $RebootPoint.addhours(-$Period) -Format f) )<br />
           <button type='button' id='OKButton' onclick='RecordSettings("OKquiet")'>$QuietButton</button><br />
           <hr />
         </span>
       </span>
       $NowText<br />
       <button type='button' id='GoButton' onclick='RecordSettings("GoAhead")'>$NowButton</button><br />
          <label>
            <input type="checkbox" name='ShutdownCheck' onclick="SetShutdown()" $($Global:Settings.Root.Shutdown)/>
            $ShutdownCheck
          </label>
     <hr style="height:3px" />
     $($ContactText -replace '<ToBeInserted>','')
       <a href="mailto:$($Address)?Subject=$($Policy -replace ' ','%20')" target="_top">$Address</a>.
     </p>
   </div>
   <script type="text/javascript">
   // Some variables set by calling script
   var defperiod = $($Period*60);  // how many minutes between checks by default
   var mins = $($Remaining.tostring('f0'));      // how many minutes until reboot (change this as required)
   var period = $(Set-NextInterval $Period $Remaining);    // how many minutes between notices
   var howMany = Math.round(mins * 60);    // total time in seconds

   var innerWidth = document.body.offsetWidth;
   var innerHeight = document.body.offsetHeight;
   var CladWidth = 500 - innerWidth;
   var CladHeight = 700 - innerHeight;

   if (howMany <= 60) {
     document.getElementById('TooLate').style.display="none";
   } else if (howMany < 60*defperiod) {
     document.getElementById('NoReminder').style.display="none";
   }
   var WinWidth = document.getElementById('Notice').offsetWidth*1.05 + CladWidth;
   var WinHeight = document.getElementById('Content').offsetHeight*1.00 + CladHeight;
   window.resizeTo(WinWidth,WinHeight);
   window.moveTo(screen.availWidth - WinWidth,screen.availHeight-WinHeight);

   if (period > 120) {
     document.getElementById('SnoozeButton').value = "Remind me again in " + Math.ceil(period/60) + " hours";
   } else {
     document.getElementById('SnoozeButton').value = "Remind me again in " + Math.ceil(period) + " minutes";
   }
   document.getElementById('FinalPeriod').innerText = Math.ceil(defperiod/60);


   beep();

   function OpenURL() {
     var shell = new ActiveXObject("WScript.Shell");
     shell.run("$Purl",0);
   }

   // JavaScript Number prototype Property
   // http://www.w3schools.com/jsref/jsref_prototype_num.asp
   Number.prototype.toMinutesAndSeconds = function() {
     Hrs = Math.floor(this/3600);
     Mins = Math.floor(this/60)-Hrs*60;
     Secs = this-(Hrs*60*60)-(Mins*60);
     return ((Hrs>1)?Hrs+" hours ":"")+((Hrs==1)?Hrs+" hour ":"")+
               ((Mins>1)?Mins+" minutes ":"")+((Mins==1)?Mins+" minute ":"")+
               (((Secs)>=10)?Secs+" seconds":"0"+Secs+" seconds");
   }

   function display(seconds, output) {
     // update screen with remaining time
     output.innerHTML = (--seconds).toMinutesAndSeconds();
     if(seconds > 0) {
       if (seconds <= 60) {
         beep();  
         document.getElementById('TooLate').style.display="none";
         window.focus();
       } else if (seconds < 60*defperiod) {
         document.getElementById('NoReminder').style.display="none";
       }
       // Recursive call after 1 second
       window.setTimeout(function(){display(seconds, output)}, 1000);
     }
     if (seconds <= 0) {
       RecordSettings("Auto");
     }
   }

   // Call recursive function on start supplying initial time and
   //   countdown <div> element
   display(howMany, document.getElementById("countdown"));

   </script>
   </body>
   </html>
"@
      
      # Create and store the path to a one-time-use HTA file.
      $HTApath = $env:TEMP + '\Reboot' + (Get-Date -Format 'yyyyMMddHHmmss') + '.hta'
      # Create the file.
      $HTACode > $HTApath
      # Start the file.

      # Before opening the HTA, create a script block that will run in parallel to the 
      #    HTA notice window and display balloon notices (and eventually minimize all 
      #    other windows to focus on the notice).
      $BalloonNoticeScriptBlock = {
         function Show-BalloonTip {
            [CmdletBinding(SupportsShouldProcess = $true)]
            param (
               [Parameter(Mandatory=$true)][string]$Text,
               [Parameter(Mandatory=$true)][string]$Title,
               [ValidateSet('None', 'Info', 'Warning', 'Error')][string]$BalloonIcon = 'Info',
               [string]$NoticeIcon = (Get-Process -id $pid | Select-Object -ExpandProperty Path),
               [int]$Timeout = 10000
            )
            Add-Type -AssemblyName System.Drawing

            # This will allow for referencing any icon that can be seen inside a binary file.
            #    https://social.technet.microsoft.com/Forums/exchange/en-US/16444c7a-ad61-44a7-8c6f-b8d619381a27/using-icons-in-powershell-scripts?forum=winserverpowershell
            $code = @'
               using System;
               using System.Drawing;
               using System.Runtime.InteropServices;

               namespace System {
                  public class IconExtractor {
                     public static Icon Extract(string file, int number, bool largeIcon) {
                     IntPtr large;
                     IntPtr small;
                     ExtractIconEx(file, number, out large, out small, 1);
                     try { return Icon.FromHandle(largeIcon ? large : small); }
                     catch { return null; }
                     }
                     [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
                     private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
                  }
               }
'@
            Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing
            $Icon = [IconExtractor]::Extract('shell32.dll', 77, $true)  # 77 is the Warning triangle

            if ([version](Get-WmiObject win32_operatingsystem).version -lt [version]'6.2') {   # Pre Windows 8 uses balloons
               Add-Type -AssemblyName System.Windows.Forms
               if ($Global:balloon -eq $null) {
                  $Global:balloon = New-Object -TypeName System.Windows.Forms.NotifyIcon
               }
               $balloon.Icon            = $Icon
               $balloon.BalloonTipIcon  = $BalloonIcon
               $balloon.BalloonTipText  = $Text
               $balloon.BalloonTipTitle = $Title
               $balloon.Text            = $Title
               $balloon.Visible         = $true

               $balloon.ShowBalloonTip($Timeout)

               $null = Register-ObjectEvent -InputObject $balloon -EventName BalloonTipClicked -Action {
                              $balloon.Dispose()
                              Unregister-Event -SourceIdentifier $EventSubscriber.SourceIdentifier
                              Remove-Job -Id $EventSubscriber.Action
                           }
            } else {   # Win8 and Win10 use Toast notifications
               Add-Type -AssemblyName Windows.UI

               if (-not (Test-Path "$env:TEMP\WarningIcon.png")) {
                  $Icon.ToBitmap().Save("$env:TEMP\WarningIcon.png", [System.Drawing.Imaging.ImageFormat]::png)
               }

               # reference:  https://stackoverflow.com/questions/46814858
               $AppID = ((Get-StartApps -Name '*PowerShell*') | Select-Object -First 1).AppId
               $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
               $Template = [Windows.UI.Notifications.ToastTemplateType]::ToastText02
               [xml]$ToastTemplate = ([Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent($Template).GetXml())

               [xml]$ToastTemplate = @"
               <toast scenario='Alarm'>
                 <visual>
                   <binding template="ToastText02">
                     <text id='1'>$Title</text>
                     <text id='2'>$Text</text>
                     <image placement="appLogoOverride" src="file:///$env:TEMP\WarningIcon.png"/>
                   </binding>
                 </visual>
                 <actions>
                   <action activationType="system" arguments="dismiss" content=""/>
                 </actions>
               </toast>
"@
               $ToastXml = New-Object -TypeName Windows.Data.Xml.Dom.XmlDocument
               $ToastXml.LoadXml($ToastTemplate.OuterXml)

               $notify = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppID)
               $notify.Show($ToastXml)
            } # END (for Win8+ notifications)
         } # END function Show-BalloonTip

         # This allows for checking if the front most window is the HTA
         Add-Type  -TypeDefinition @'
         using System;
         using System.Runtime.InteropServices;
         using System.Text;
         public class UserWindows {
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
               public static extern int GetWindowText(IntPtr hwnd,StringBuilder lpString, int cch);
            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
               public static extern IntPtr GetForegroundWindow();
            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
               public static extern Int32 GetWindowTextLength(IntPtr hWnd);
         }
'@

         # Only display notice ballons if the Notification Window is NOT put in front
         $ForgroundWindow = [UserWindows]::GetForegroundWindow()
         $FGWTitleLength = [UserWindows]::GetWindowTextLength($ForgroundWindow)
         $StringBuilder = New-Object -TypeName text.stringbuilder -ArgumentList ($FGWTitleLength + 1)
         $null = [UserWindows]::GetWindowText($ForgroundWindow,$StringBuilder,$StringBuilder.Capacity)
         while ($StringBuilder.ToString() -notmatch $HTATitle) {
            Show-BalloonTip -Text $BalloonText -Title $BalloonTitle -BalloonIcon Warning
            Start-Sleep -Seconds $NapTime
            $ForgroundWindow = [UserWindows]::GetForegroundWindow()
            $FGWTitleLength = [UserWindows]::GetWindowTextLength($ForgroundWindow)
            $StringBuilder = New-Object -TypeName text.stringbuilder -ArgumentList ($FGWTitleLength + 1)
            $null = [UserWindows]::GetWindowText($ForgroundWindow,$StringBuilder,$StringBuilder.Capacity)
         }

         if ($Global:balloon) { $Global:balloon.Dispose() }
      } # END $BalloonNoticeScriptBlock

      If ($Global:Settings) {
         $bTitle = $BaloonTitle2 -replace '<.*>',$([math]::Round($Remaining/60))
      } else {
         $bTitle = $BaloonTitle1
      }
      # Display a balloon notice 10 times/default cycle time
      if (($PauseTime = $Period*60*60/10) -lt 30) {$PauseTime = 30}  # number of seconds between balloons

      $runspace = [runspacefactory]::CreateRunspace()
      $runspace.open()
      $ps = [powershell]::create()
      $ps.runspace = $runspace
      $runspace.sessionstateproxy.setvariable('BalloonText',$BaloonText)
      $runspace.sessionstateproxy.setvariable('BalloonTitle',$bTitle)
      $runspace.sessionstateproxy.setvariable('HTATitle',$WindowTitle)
      $runspace.sessionstateproxy.setvariable('NapTime',$PauseTime)
      $ps.AddScript($BalloonNoticeScriptBlock)
      $ps.BeginInvoke()

      # If the user has not acknowledged for nearly the entire minimum lead time or 
      #    remaining time is running out, force user attention to the notice window.
      if (((New-TimeSpan -Start $FlagDate -End (Get-Date)).TotalHours -ge ($MinLead - $Period/2)) -or
            ($Remaining -le $Period*60)) {
         (New-Object -ComObject Shell.Application).minimizeall()
      }

      # Now open the HTA
      start-process -FilePath $HTApath -Wait

      $runspace.Close()
      $ps.Dispose()

      # Allow other users to manipulate the settings file in case of an unclean-reboot
      if (Test-Path $Global:SettingsPath) {
         $Acl = get-acl -Path $Global:SettingsPath
         $rule = New-Object -TypeName system.security.accesscontrol.filesystemaccessrule -ArgumentList ('Authenticated Users','Modify','Allow')
         $Acl.setaccessrule($rule)
         set-acl -Path $Global:SettingsPath -AclObject $Acl

         # No (previous) Settings = No user Acknowledgement
         # Thus, if the Settings file now exists, the user has just acknowledged.
         if (-Not $Global:Settings) {
            $Priors = 0
         } else {   # otherwise count priors to check for RE-acknowledgment
            $Priors = @($Global:Settings.root.Acknowledgements.stamp).count
         }
         # re-read settings in case they were modified by an HTA button choice
         $Global:Settings = [xml](get-content -Path $Global:SettingsPath)

         $EventLogArgs = @{
            EntryType = 'Warning'
            EventID   = 100   # Deadline established = User acknowledged
            Message   = $Global:LogEvents.get_item(100) -replace '<ToBeInserted>',$RebootPoint
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs

      } #end if (test-path...
      
   } #end if (not quiet or time is running out) {show HTA}
   
   # Recalculate how much time remains (for long-displayed HTA)
   $Remaining = (New-TimeSpan -Start (get-date) -End $RebootPoint).TotalMinutes

   # When the command comes down from the GUI, restart/poweroff
   If ($Global:Settings.Root.ActNow) {
      $EventLogArgs = @{
         EntryType = 'Information'
         EventID   = 2001   # User requested immediate reboot/shutdown
         Message   = $Global:LogEvents.get_item(2001)
      }
      Write-EventLog @Global:CommonLogArgs @EventLogArgs

      If ($Global:Settings.Root.Shutdown -eq 'Checked') {
         Reset-AutoReboot -Action 'PowerOff'
         Return
      } else {
         Reset-AutoReboot -Action 'Restart'
         Return
      }
   }

   # Settings are stamped every time the user clicks a button
   if (@($Global:Settings.root.Acknowledgements.stamp).count -gt $Priors) {
      $EventLogArgs = @{
         EntryType = 'Information'
         EventID   = 101   # User "snoozed" the reboot
         Message   = $Global:LogEvents.get_item(101) + "`n`nDeadline: $(get-date -Date $RebootPoint)"
      }
      Write-EventLog @Global:CommonLogArgs @EventLogArgs

      # Log when the user requested quiet, but not once the warnings return (at the end)
      if (($Global:Settings.Root.quiet -eq 'True') -and -not ($Remaining -le ($Period*60))) {
         $EventLogArgs = @{
            EntryType = 'Warning'
            EventID   = 666   # User requested no more reminders
            Message   = $Global:LogEvents.get_item(666)
         }
         Write-EventLog @Global:CommonLogArgs @EventLogArgs
      }
   }

   # Give the longest "snooze" available to a user who has asked 
   # to not be bothered and to any user who just clicked a button.
   if (($Global:Settings.Root.quiet -eq 'True') -or 
         (@($Global:Settings.root.Acknowledgements.stamp).count -gt $Priors)) {
      $NextInterval = Set-NextInterval $Period $Remaining
   } else { # The user didn't see the notice, so bring it back soon
      $NextInterval = 1
   }
   
} else { #End If($Reason and 24hrs since reboot)   
   # Clean up any crud that may be left behind
   If ($Global:Settings) {
      Reset-AutoReboot -LogTime $LastBootTime.ToString('yyyy.MM.dd-HH.mm')
   }
   $NextInterval = $Period*60
} #End Else

# If somehow to this point when running as SYSTEM, exit.
if ($Global:RunningElevated) { Return }

# Create or update a task to re-run this script at a later time.
if ($Global:TaskFolder.gettasks(1) | Where-Object {$_.Name -ieq $UserTaskName}) {
   $Global:TaskFolder.GetTask($UserTaskName).enabled = $true
   $TaskDef = $Global:TaskFolder.GetTask($UserTaskName).definition
   # Adjust the scheduled task to re-run the script at a later time.
   $TaskDef.Triggers | ForEach-Object {
      $_.StartBoundary = get-date -Date (get-date).AddMinutes($NextInterval) -Format 'yyyy\-MM\-dd\THH:mm:ss'
   }
   $TaskDef.Actions | ForEach-Object {$_.path = $VBScriptPath}
} else {
   $TaskDef = $TaskService.NewTask(0)  # Not sure what the "0" is for
   $Taskdef.RegistrationInfo.Description = $UserDescription
   $TaskDef.RegistrationInfo.Date = $Now.tostring('yyyy\-MM\-dd\THH:mm:ss.00000')
   $Taskdef.RegistrationInfo.Author = "$env:USERDOMAIN\$env:USERNAME"
   $TaskDef.settings.priority = 5
   $TaskDef.Settings.MultipleInstances = 3
   $TaskDef.settings.StartWhenAvailable = $true
   # Create a trigger to run after the next time interval
   $Trigger = $Taskdef.Triggers.Create(1)
   $Trigger.Id = 'NextCheck'
   $Trigger.StartBoundary = get-date -Date (get-date).AddMinutes($NextInterval) -Format 'yyyy\-MM\-dd\THH:mm:ss'
   $Trigger.Enabled = $true
   # Run this script again when triggered.
   $Action = $Taskdef.Actions.Create(0)
   $Action.Path = $VBScriptPath
}

# Wake to reboot only if user acknowleged notice
If ($Global:Settings) { 
   $Taskdef.Settings.WakeToRun=$True 
} Else { 
   $Taskdef.Settings.WakeToRun=$False 
}

# Finally, register the task
$Global:TaskFolder.RegisterTaskDefinition($UserTaskName, $Taskdef, 6, $null, $null, 3) > $null

Return   
#END Script
