'# CheckAutoReboot - An interface to auto-reboot after updates
'# Copyright 2015 Erich Hammer
'# This script/information is free: you can redistribute it and/or modify
'# it under the terms of the GNU General Public License as published by
'# the Free Software Foundation, either version 2 of the License, or
'# (at your option) any later version.
'#
'# This script is distributed in the hope that it will be useful,
'# but WITHOUT ANY WARRANTY; without even the implied warranty of
'# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'# GNU General Public License for more details.
'#
'# The GNU General Public License can be found at <http://www.gnu.org/licenses/>.
' =====================================================================
'  This "polyglot script" allows a PowerShell script to be embedded 
'  in a .vbs file.  The purpose for this is to run PowerShell code 
'  from a scheduled task **without showing a console window**. 
'
'  All PowerShell code must both exist between the # START and # END 
'  lines and have a ' (quote mark) as the first char of the line (to 
'  exclude it from VBS parsing).
'
'  To see PoSh window and any errors, add " -noexit" to the start of 
'  the PoShswitches variable AND change "0" to "1" in the run command
'  toward the end of this entire file (both in the last 12 lines)
'  .
' =====================================================================
'# Start PowerShell  # (Don't modify this line!)
'
'<#
'.SYNOPSIS
'   A process and interface for auto-rebooting after updates.
'.DESCRIPTION
'   This script will check if Windows is in a pending reboot state.
'   If so, it will start a process of notifying the current user and setting
'   a deadline for automatic reboot.  When run, it will create or modify 
'   scheduled tasks to call the script to run again at a later time.
'.PARAMETER RebootTime
'   The time of day (in 24-hr time) auto-reboots will occur on the weekday defined
'   in RebootDOW.
'   Default value: 17:00
'.PARAMETER RebootDOW
'   The day of the week auto-reboots will take place at the time defined
'   in the RebootTime.
'   Default value: Friday
'.PARAMETER BlockFile
'   The path of the file that, if exists, will prevent a machine from rebooting or
'   even checking for a pending reboot.  This should be a location where changes
'   require elevated administrator rights.
'   Default Value: <SystemDrive>\NoReboot -- e.g. "C:\NoReboot"
'.PARAMETER OrgName
'   The short name of the organization.  This is used in several places including
'   for separating files and tasks from other applications and system processes.  
'   For example, the script and log files will be in the "\ProgramData\OrgName\Reboot" 
'   folder, there will be a "\OrgName" task folder, and the HTA window title will
'   be "OrgName Security Reboot Notice".
'   Default Value: ITServices
'.PARAMETER DefaultName
'   Name to use if the script cannot discover its own name.  Reason:  This script
'   needs to be able to create a scheduled task which will call the script again.
'   Since PowerShell scripts called from a scheduled task cannot hide the console
'   window, this PowerShell code may be embedded in a polyglot .vbs script and 
'   will not be able to discover its name.  It must know it's name in order to 
'   configure a scheduled task to run the script again at a later time.
'   Default Value: CheckAutoReboot.vbs
'.PARAMETER TaskName
'   The name of the scheduled task that will run this script (and be modified
'   by it) at a later time.
'   Default Value: 'Check for Pending Reboot - <username>'
'.PARAMETER MinLead
'   Minimum hours before auto-reboot point user must acknowledge warning or 
'   the reboot will be delayed by a week.
'   Default Value:  54
'.PARAMETER Period
'   The default number of hours between checks for pending reboots and/or
'   warnings that a reboot is pending/scheduled.  As time expires, warnings
'   will occur more frequently.
'   Default Value:  4
'.PARAMETER Policy
'   The name or brief description of the policy that requires a computer to
'   be up-to-date.  This will follow the text:
'         "This computer must reboot to comply with the"
'   and be the text for the Purl to allow users to see the policy.
'   Default Value:  "Enterprise security requirements"
'.PARAMETER Purl
'   The url to view the policy named in the Policy parameter.  This will
'   open in the user's default browser.
'   Default Value:  http://www.school.edu/autorebootpolicy.php
'.PARAMETER Address
'   The email address for help or questions about the reboot notice.
'   Default Value:  support@school.edu
'.NOTES
'   Copyright 2015-2017 Erich Hammer
'
'   This script/information is free: you can redistribute it and/or
'   modify it under the terms of the GNU General Public License as
'   published by the Free Software Foundation, either version 2 of the 
'   License, or (at your option) any later version.
'
'   This script is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   The GNU General Public License can be found at <http://www.gnu.org/licenses/>.
'#>
'[CmdletBinding()]
'Param(
'   # Time of day for auto-reboots to occur
'   [ValidateScript({
'      If ($_ -match '^([01]\d|2[0-3]):?([0-5]\d)$') { $true }
'      else {Throw "`n'$_' is not a time in HH:mm format."}
'   })]
'   [string]$RebootTime = '17:00',
'
'   # Day of week for auto-reboots to occur
'   [ValidateScript({
'      $Days = ('Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa')
'      If ($Days -icontains ($_[0..1] -join '')) { $true } 
'      else {Throw "`n'$_' is not a day of the week or an abbreviation for one."}
'   })]
'   [string]$RebootDOW = 'Friday',
'
'   # If present, this machine should not auto-reboot
'   [string]$BlockFile = (Join-Path $env:SystemDrive 'NoReboot'),
'
'   # Short name of the organization used for sub-paths
'   [string]$OrgName = 'ITServices',
'
'   # Use this name if can't discover own name
'   [string]$DefaultName = 'CheckAutoReboot.vbs',
'
'   # The task that will be scheduled by this script
'   [string]$TaskName = 'Check for Pending Reboot - ' + $env:USERNAME,
'
'   # Minimum hours before auto-reboot point user must acknowledge warning or it will be delayed
'   [int]$MinLead = 54,
'
'   # Default hours between checks/warnings.
'   [int]$Period = 4,
'
'   # The name of the security policy requiring reboots after updates
'   [string]$Policy = 'Enterprise security requirements',
'
'   # The URL to the policy
'   [string]$Purl = 'http://www.school.edu/autorebootpolicy.php',
'
'   # The email address for IT support
'   [string]$Address = 'support@school.edu'
')
'
'#=====================
'#region Functions
'#=====================
'Function Test-RebootPending
'{ 
'<# 
'.SYNOPSIS 
'   Tests the pending reboot status of the local computer. 
' 
'.DESCRIPTION 
'   This function will query the registry and determine if the system is pending a reboot. 
'   Checks:
'   CBServicing = Component Based Servicing (Windows Vista/2008+) 
'   WindowsUpdate = Windows Update / Auto Update (Windows 2003+) 
'   CCMClientSDK = SCCM 2012 Clients only (DetermineIfRebootPending method) otherwise $null value 
'   PendFileRename = PendingFileRenameOperations (Windows 2003+) 
'   
'.LINK 
'    Component-Based Servicing: 
'    http://technet.microsoft.com/en-us/library/cc756291(v=WS.10).aspx 
'   
'    PendingFileRename/Auto Update: 
'    http://support.microsoft.com/kb/2723674 
'    http://technet.microsoft.com/en-us/library/cc960241.aspx 
'    http://blogs.msdn.com/b/hansr/archive/2006/02/17/patchreboot.aspx 
' 
'    SCCM 2012/CCM_ClientSDK: 
'    http://msdn.microsoft.com/en-us/library/jj902723.aspx 
' 
'.NOTES 
'    Inpired by: http://blogs.technet.com/b/heyscriptingguy/archive/2013/06/11/determine-pending-reboot-status-powershell-style-part-2.aspx
'#> 
' 
'[CmdletBinding()] 
'[OutputType([bool])]
'param()
'
'Begin { $StartLoc = Get-Location }
'Process { 
'   $Reason = @()
'   
'   # Query the Component Based Servicing Reg Key
'   Set-Location -Path 'hklm:\SOFTWARE\Microsoft\Windows\CurrentVersion'
'   if ((get-item '.\Component Based Servicing').getsubkeynames() -contains 'RebootPending' ) {
'      $Reason += 'ComponentBasedServicing' 
'   }
'
'   # Query WUAU from the registry 
'   if ((get-item '.\WindowsUpdate\Auto Update').getsubkeynames() -contains 'RebootRequired' ) {
'      $Reason += 'WindowsUpdate' 
'   }
'       
'   # Query PendingFileRenameOperations from the registry - REMOVED FOR HIGH FREQUENCY/LOW RISK
'<#   Set-Location -Path 'hklm:\SYSTEM\CurrentControlSet\Control\Session Manager'
'   if ((get-item .).getvalue('pendingfilerenameoperations')) {
'      $Reason += 'PendingFileRename' 
'   }
'#>
'   # Determine SCCM 2012 Client Reboot Pending Status 
'   $CCMClientSDK = $null 
'   $CCMSplat = @{ 
'      NameSpace='ROOT\ccm\ClientSDK' 
'      Class='CCM_ClientUtilities' 
'      Name='DetermineIfRebootPending' 
'      ComputerName=$env:COMPUTERNAME 
'      ErrorAction='SilentlyContinue' 
'   } 
'   $CCMClientSDK = Invoke-WmiMethod @CCMSplat 
'
'   If ($CCMClientSDK.IsHardRebootPending -or $CCMClientSDK.RebootPending) { 
'      $Reason += 'SCCMclient' 
'   }
'   
'   $Reason
'}#End Process 
'End { Set-Location -Path $StartLoc }   
'
'}#End Test-RebootPending Function
'
'Function Reset-AutoReboot 
'{ 
'<#
'.SYNOPSIS 
'   Resets the auto-reboot settings and optionally powers down or reboots. 
' 
'.DESCRIPTION 
'   This function:
'      - Deletes the auto-reboot settings file. 
'      - Deletes the task (possibly) created to prevent logoff.
'      - Deletes the task for re-running the auto-reboot check.
'      - Forces Shut down or reboot of computer (depending on settings).
'#>
'   [CmdletBinding(DefaultParametersetName='Clear')] 
'   Param([Parameter(Mandatory=$true, position=0)][string]$SettingsFile,
'         [Parameter(Mandatory=$true, position=1)][string]$TaskName,
'         [Parameter(position=2)][string]$TaskFolderName = '\',
'         [Parameter(ParameterSetName='Clear')][string]$LogTime = (get-date -Format yyyy.MM.dd-HH.mm) ,
'         [Parameter(ParameterSetName='Restart')][switch]$Restart = $False,
'         [Parameter(ParameterSetName='PowerOff')][switch]$PowerOff = $False)
'
'   # Keep the settings files and the last .HTA for troubleshooting reference
'   if (Test-Path $SettingsFile) { Rename-Item $SettingsFile ('reboot' + $LogTime + '.log')}
'   $HTAs = Get-ChildItem ($env:TEMP + '\Reboot*.hta') | Sort-Object LastWriteTime -Descending
'   $HTAs[1..$HTAs.count] | % { Remove-Item $_.Fullname}
'   
'   # Delete leftover tasks in case the user rebooted on their own
'   $TaskService = New-Object -ComObject Schedule.Service
'   $TaskService.connect()                     # connect to the local computer (default)
'   $ErrorActionPreference = 'stop'
'   Try {
'      $TaskFolder = $TaskService.GetFolder($TaskFolderName)
'   } Catch {
'      # Fall back to the root folder on error
'      $TaskFolder = $TaskService.GetFolder('\')
'   } Finally { 
'      $ErrorActionPreference = 'continue'
'      if (($TaskFolder.gettasks(1) |Select-Object -expandproperty name) -icontains 'No Logoff') {
'         $TaskFolder.DeleteTask('No Logoff',0)
'      }
'      if (($TaskFolder.gettasks(1) |Select-Object -expandproperty name) -icontains $TaskName) {
'         $TaskFolder.DeleteTask($TaskName,0)
'   }
'   }
'
'#   if ($PowerOff) { Stop-Computer -Force }    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Commented for testing
'#   if ($Restart) { Restart-Computer -Force }  # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Commented for testing
'   if ($PowerOff -or $Restart) {               # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Uncommented for testing -- Put <# at front for production
'   (new-object -ComObject wscript.shell).popup(
'         "You requested a shutdown/restart, but`nsince this is a test, nothing will happen.",0,'Bug Testing'
'         ) > $null
'   }
'#>
'} #End Reset-AutoReboot function
'
'Function Set-NextInterval
'{
'<#
'.SYNOPSIS 
'   Determins how long until the auto-reboot notice should return. 
' 
'.DESCRIPTION 
'   This function compares the default time for a return notice with the remaining time
'   to a shutdown and gives an increasingly shorter "snooze" time for the notice to return 
'   as time is running out. For example, a default period of 4 hours leads to:
'      more than 8 hours left -> 4 hour snooze
'      4-8 hours left -> 2 hour snooze
'      2-4 hours left -> 1 hour snooze
'      1-2 hours left -> 30 minute snooze
'      .5-1 hour left -> 15 minute snooze
'      15-30 minutes left -> 7.5 minute snooze
'      7.5-15 minutes left -> 3.75 minutes snooze
'      less than 7.5 minutes left -> 1 minute snooze
'#>
'   [CmdletBinding()]
'   Param([int]$DefaultPeriod=4,
'         [int]$TotalTimeLeft=8)
'   
'   $Ratio = $TotalTimeLeft / ($DefaultPeriod * 60)
'   If ($Ratio -ge 2) {$NextInterval = $DefaultPeriod*60 }
'   ElseIf ($Ratio -ge 1) { $NextInterval = $DefaultPeriod*60/2 }
'   ElseIf ($Ratio -ge .5) { $NextInterval = $DefaultPeriod*60/4 }
'   ElseIf ($Ratio -ge .25) { $NextInterval = $DefaultPeriod*60/8 }
'   ElseIf ($Ratio -ge .125) { $NextInterval = $DefaultPeriod*60/16 }
'   ElseIf ($Ratio -ge .0625) { $NextInterval = $DefaultPeriod*60/32 }
'   ElseIf ($Ratio -ge .03125) { $NextInterval = $DefaultPeriod*60/64 }
'   Else {$NextInterval = 1}
'   
'   $NextInterval
'}
'
'#=====================
'#endregion Functions
'#=====================
'
'
'# Double-check that this machine can reboot
'if (Test-Path $BlockFile) {
'   Return
'}
'
'# Define the xml settings file path
'$AppSettings = Join-path $env:ProgramData (Join-path $OrgName 'Reboot\Shutdown.xml')
'
'if (Test-Path $AppSettings) {
'   $Settings = [xml](get-content $AppSettings)
'}
'
'# To set the triggers for the scheduled task, need to know the name of the task.  
'# If the script name is known, can search the tasks for the calling task
'$ScriptPath = $MyInvocation.MyCommand.path
'
'If (-not $ScriptPath) {
'   # If this is a VBS-to-PoSh polyglot script, no name will be returned.
'   # Set it to a default which can be changed in the VBS with a line like:
'   #     PoShcmd = PoShcmd & "-replace 'Check-AutoReboot.vbs','" & Wscript.ScriptFullName & "' "
'   $ScriptPath = Join-Path $env:ProgramData "$OrgName\Reboot\$DefaultName"
'}
'
'$Now = Get-Date
'# Comprehensive list of methods for obtaining last boot time: 
'#   http://www.happysysadm.com/2014/07/windows-boot-time-explored-in-powershell.html
'$OSInfo = Get-WmiObject Win32_OperatingSystem
'$LastBootTime = [Management.ManagementDateTimeConverter]::ToDateTime($OSInfo.lastbootuptime)      
'
'# Don't do much unless there is a pending reboot 
'#   AND the machine hasn't rebooted in 24 hours.
'#if (($Reason = Test-RebootPending) -and   # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Commented for testing
'#      ((New-TimeSpan -Start $LastBootTime -End $Now).TotalHours -ge 24)) {  # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Commented for testing
'if ($true) {                                  # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  Uncommented for testing

'   # Want to know when the pending reboot was first identified (for more active notifications)
'   # This will be either the time this instance started, or the timestamp of the 'No Logoff' task.
'   $TaskService = new-object -ComObject Schedule.Service
'   $TaskService.connect() 
'   $ErrorActionPreference = 'stop'
'   Try {
'      $TaskFolder = $TaskService.GetFolder($OrgName)
'      $TaskDef = $TaskFolder.GetTask('No Logoff')
'      $FlagDate = get-date $TaskDef.Definition.RegistrationInfo.date
'   } Catch {
'      $FlagDate = $Now
'   } Finally { 
'      $ErrorActionPreference = 'continue'
'   }
'
'   if ($Settings) {
'      # If a user manually restarted, the settings file from previous pending reboot will
'      #   remain.  To stop a possible auto-reboot with insufficient warning, the last boot
'      #   time must be compared with when the previous reboot warning was acknowledged.
'      # Gather the first time the pending reboot was acknowledged
'      $FirstAckTime = Get-Date ($Settings.Root.Acknowledgements.Stamp | 
'                        Where-Object {$_.mark -eq 0} |Select-Object -ExpandProperty '#text')
'
'      If ($LastBootTime -gt $FirstAckTime) {
'         # Machine restarted NOT via the script, so reset the settings.
'         Reset-AutoReboot $AppSettings $TaskName $OrgName -LogTime $LastBootTime.ToString('yyyy.MM.dd-HH.mm')
'         $Settings = $null
'      }
'   }
'
'   # The settings file should be deleted if a reboot has occurred.  As a precaution, also
'   # check that there is a set reboot point/time to guard against erroneous, spontaneous reboots.
'   if ($Settings -and $Settings.Root.RebootPoint) {
'      $RebootPoint = get-date $Settings.Root.RebootPoint
'      $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes
'      # For situations in which the time is well past the RebootPoint, give the user a one-minute warning
'      If ($Remaining -lt -10) {
'         $Remaining = 1
'      } ElseIf ($Remaining -le 0) {
'         # reboot immediately if the user has acknowledged and time is up
'         If ($Settings.Root.Shutdown -eq 'Checked') {
'            Reset-AutoReboot $AppSettings $TaskName $OrgName -PowerOff
'            Return
'         } else {
'            Reset-AutoReboot $AppSettings $TaskName $OrgName -Restart
'            Return
'         }
'      }
'   } else {  # No settings and/or no reboot point
'      If ($Settings) {
'         # No set reboot point = something went wrong -- don't reboot.
'         Reset-AutoReboot $AppSettings $TaskName $OrgName 
'         $Settings = $null
'      }
'      # Otherwise, just no settings file = User has not acknowledged pending reboot.
'      $RebootPoint = 0..6 | % {$(Get-Date $RebootTime).AddDays($_)} | 
'                           Where-Object {-not (Compare-Object $_.DayOfWeek.tostring().tolower()[0..1] $RebootDOW.tolower()[0..1])}
'      $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes
'      # delay the reboot for a week if the time is already too late      
'      if ($Remaining -lt $MinLead*60) {
'         $RebootPoint = $RebootPoint.AddDays(7)
'         $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes
'      }
'      # When a reboot is pending, we want to stop users from logging off by redirecting them to a reboot.
'      $ErrorActionPreference = 'stop'
'      Try {
'         $TaskFolder = $TaskService.GetFolder($OrgName)
'      } Catch {
'         $null = $TaskService.GetFolder('\').CreateFolder($OrgName) 
'         $TaskFolder = $TaskService.GetFolder($OrgName)
'      } Finally { 
'         $ErrorActionPreference = 'continue'
'      }
'
'      if (($TaskFolder.gettasks(1) |Select-Object -expandproperty name) -notcontains 'No Logoff') {
'         $FlagDate = $Now
'         $TaskDef = $TaskService.NewTask(0)  # Not sure what the "0" is for
'         $Taskdef.RegistrationInfo.Description = 'Force restart on Logoff'
'         $TaskDef.RegistrationInfo.Date = $Now.ToString('yyyy-MM-ddTHH:mm:ss.00000')
'         $TaskDef.settings.priority = 2
'         $TaskDef.settings.StartWhenAvailable = $true
'         $Trigger = $Taskdef.Triggers.Create(0) 
'         # The Event definition string is extracted from a sample task created in the GUI, 
'         #   exported to XML, and with the html-characters corrected.
'         #   Could just copy from XML and use:  .replace("&lt;","<").replace("&gt;",">")
'         $Trigger.subscription = '<QueryList><Query Id="0" Path="System"><Select Path="System">' +
'                                 '*[System[Provider[@Name="Microsoft-Windows-Winlogon"] and EventID=7002]]' +
'                                 '</Select></Query></QueryList>'
'         $Action = $Taskdef.Actions.Create(0)
'         $Action.Path = 'shutdown.exe'
'         $Action.Arguments = '/r /f'
'         # Finally, register the task
'         $TaskFolder.RegisterTaskDefinition('No Logoff', $Taskdef, 6, $null, $null, 3) > $null
'      }
'   } #End Else {  # No settings file and/or reboot point...
'   
'   # Don't notify users who have requested to not be reminded 
'   #   **unless** they are in the final period (i.e. time is running out).
'   if (($Settings.Root.quiet -ne 'True') -or ($Remaining -le ($Period*60))) {
'
'      if ($Settings) {
'         $Verb = 'SCHEDULED'
'         $ShowMore = 'inline'
'      } else {
'         $Verb = 'REQUIRED'
'         $ShowMore = 'none'
'      }
'
'      # The following here-string is exported to a unique HTA that will be saved in a temporary
'      #   location for one-time use.  The HTA provides a GUI, and creates/updates the 
'      #   XML settings file with information on the choices the user makes.  There are 
'      #   several insertions into the here-string which customize the HTA for the instance
'      #   to be run.  The insertions define the following:
'      #     The default time ($Period) between pending reboot checks.
'      #     A length of time for the HTA to run before closing to be replaced by another.
'      #     The number of minutes until auto-reboot occurs.
'      #     The folder/path to read/save the XML settings file.
'      #     Whether the user has previously selected to auto-shutdown vs reboot.
'      #     The reason(s) why is a reboot is pending (from the Test-RebootPending function).
'      #     The DateTime the script (NOT the user) recognized the pending reboot status.
'      #     The word ($Verb) used in the primary phrase "Reboot Required/Scheduled".
'      #     A CSS keyword ($ShowMore) to control the visibility of some parts of the HTA.
'      #     The date/time the warning will return even if the user asks to no be bothered.
'      $HTACode = @"
'   <html xmlns:IE>
'   <!-- http://www.itsupportguides.com/windows-7/windows-7-shutdown-message-with-countdown-and-cancel/ -->
'   <title>$OrgName Security Reboot Notice</title>
'   <script type="text/javascript">
'     // Quickly move the window to the approximate location
'     window.resizeTo(500,700)
'     window.moveTo(screen.availWidth-500,screen.availHeight-700);
'   </script>
'   <HTA:APPLICATION
'     BORDERSTYLE            = "Normal"
'     CAPTION                = "Yes"
'     CONTEXTMENU            = "No"
'     INNERBORDER            = "Yes"
'     MAXIMIZEBUTTON         = "No"
'     MINIMIZEBUTTON         = "No"
'     NAVIGABLE              = "No"
'     SCROLL                 = "No"
'     SCROLLFLAT             = "Yes"
'     SELECTION              = "No"
'     SHOWINTASKBAR          = "No"
'     SINGLEINSTANCE         = "Yes"
'     SYSMENU                = "No"
'   >
'   <head>
'   <STYLE>
'   html, body {
'       position:fixed;
'       border:0;
'       top:0;
'       bottom:0;
'       left:0;
'       right:0;
'       font-family: Verdana,Arial,Helvetica,sans-serif;
'       font-size: 87%;
'       line-height: 1.5em;
'       margin: 0;
'       padding: 0;
'       background-color: #FCDD00;
'       text-align: center;
'
'   #top_body {
'       clear: both;
'       margin: 20px 0 0;
'   }
'   #top_body p {
'       font-style: italic;
'       margin: 20px 0 0;
'   }
'   #countdown {
'       color: #FF0000;
'       font-size: 1.4em;
'       font-weight: bold;
'   }
'   #LinkSpan { cursor: pointer; }
'   img {
'       width: 25%;
'       height: auto;
'   }
'   h1 {
'       color: #2B165E;
'       font-weight: bold;
'       text-transform: uppercase;
'       white-space: nowrap;
'   }
'   hr {
'   width: 25%
'   } 
'   button {
'       font-family: tahoma;
'       font-size: 1.4em;
'       font-weight: 500;
'       overflow:visible;
'       padding:.1em,.5em;
'       margin: .3em;
'   }
'
'   </STYLE>
'
'   <script language='vbscript'>
'   dim ShutdownFilePath, strShutdown, idTimer
'   ShutdownFilePath = "$AppSettings"
'
'   Sub window_onLoad
'     idTimer = window.setTimeout("close()", 0.99*1000*60*$(Set-NextInterval $Period $Remaining), "VBScript")
'     window.focus()
'     SetShutdown
'   End Sub
'
'   Sub Beep
'     BeepSound = chr(007)
'     CreateObject("WScript.Shell").Run "cmd /c @echo " & BeepSound, 0
'   End Sub
'
'   sub SetShutdown 
'     if(document.getElementById("ShutdownCheck").checked) Then
'       strShutdown = "Checked"
'     else
'       strShutdown = ""
'     End If
'   End Sub
'
'   sub ReplaceString(strFilename, strSearch, strReplace) 
'      'usage: cscript replace.vbs Filename "StringToFind" "stringToReplace"
' 
'      Dim fso,objFile,oldContent,newContent
'      Set fso=CreateObject("Scripting.FileSystemObject")
'
'      'Read file
'      set objFile=fso.OpenTextFile(strFilename,1)
'      oldContent=objFile.ReadAll
' 
'      'Write file
'      newContent=replace(oldContent,strSearch,strReplace,1,-1,0)
'      set objFile=fso.OpenTextFile(strFilename,2)
'      objFile.Write newContent
'      objFile.Close 
'   End Sub
'
'   Sub RecordSettings(strChoice)
'     Set fso = CreateObject("Scripting.FileSystemObject")
'     Set xmlDoc = CreateObject("Microsoft.XMLDOM")
'  
'     If strChoice = "GoAhead" Then
'       iAnswer = CreateObject("Wscript.Shell").Popup( _
'                 "Have you saved your work?", _
'                 60, _
'                 "Confirm Reboot", _
'                 vbYesNo + vbExclamation + vbSystemModal)
'       if iAnswer <> vbYes Then
'         exit sub
'       End If
'     End If
'  
'     If (fso.FileExists(ShutdownFilePath)) Then
'       xmlDoc.Async = "False"
'       xmlDoc.load ShutdownFilePath
'       Set objRoot = xmlDoc.documentElement
'    
'       Set objReason = xmlDoc.selectSingleNode("//RebootReason")
'       objReason.text = "$($Reason -join ',')"
'
'       Set Node = xmlDoc.selectSingleNode("//Acknowledgements")
'         Set objStamp = xmlDoc.createElement("Stamp")
'         objStamp.text = Now
'         Node.appendChild objStamp
'    
'       Set Node = xmlDoc.selectSingleNode("//Shutdown")
'       Node.text = strShutdown
'    
'       If strChoice = "GoAhead" Or strChoice = "Auto" Then
'         Set objRecord = xmlDoc.createElement("ActNow")
'         objRecord.text = "True"
'         objRoot.appendChild objRecord
'       ElseIf strChoice = "OKquiet" Then
'         Set Node = xmlDoc.selectSingleNode("//Quiet")
'         Node.text = "True"
'       End If
'     Else  
'       NewXml = True
'       Set objRoot = xmlDoc.createElement("Root")  
'       xmlDoc.appendChild objRoot  
'
'       Set objReason = xmlDoc.createElement("RebootReason")
'       objReason.text = "$($Reason -join ',')"
'       objRoot.appendChild objReason
'
'       Set objReason = xmlDoc.createElement("TriggerPoint")
'       objReason.text = "$($FlagDate.ToString())"
'       objRoot.appendChild objReason
'
'       Set objRecord = xmlDoc.createElement("Acknowledgements")
'       objRoot.appendChild objRecord
'         Set objStamp = xmlDoc.createElement("Stamp")
'         objStamp.text = Now
'         Set objMark = xmlDoc.createAttribute("Mark")
'         objMark.Value = 0
'         objStamp.attributes.setNamedItem(objMark)
'         objRecord.appendChild objStamp
'
'       Set objRecord = xmlDoc.createElement("Shutdown")
'       objRecord.text = strShutdown
'       objRoot.appendChild objRecord
'
'       If strChoice = "Snooze" Then
'         Set objRecord = xmlDoc.createElement("Quiet")
'         objRecord.text = "False"
'       ElseIf strChoice = "OKquiet" Then
'         Set objRecord = xmlDoc.createElement("Quiet")
'         objRecord.text = "True"
'       Else 
'         Set objRecord = xmlDoc.createElement("ActNow")
'         objRecord.text = "True"
'       End If
'       objRoot.appendChild objRecord
'
'       Set objIntro = xmlDoc.createProcessingInstruction ("xml","version='1.0'")  
'       xmlDoc.insertBefore objIntro,xmlDoc.childNodes(0)  
'     End If 
'     xmlDoc.save ShutdownFilePath
'
'     If NewXml Then
'       ReplaceString ShutdownFilePath, "><", (">" + vbcr + "<")
'     Else
'       ReplaceString ShutdownFilePath, "Stamp><Stamp", ("Stamp>" + vbcr + vbTab + vbTab + "<Stamp")
'     End If
'
'     close()
'   End Sub
'
'   </script>
'   </head>
'   <body>
'   <div id="Content" style="zoom:1">
'     <p><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAlgAAAJzCAYAAADTDW0pAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAumgAALpoBcGdoOgAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAACAASURBVHic7N13mCRVucfx78xmYBNLhiVJTpIUkChIEBERcwBMgNcEesWcs+g1J0QRCSKgKBKVHARFlyBBskhOy5I2787cP87usmFCd09VvRW+n+d5H5ZlpufHVHf126dOndOFpKabAN3TokPUS89k4MHoFJLidEcHkCRJqhsbLEmSpIzZYEmSJGXMBkuSJCljNliSJEkZs8GSJEnKmA2WJElSxmywJEmSMmaDJUmSlDEbLEmSpIzZYEmSJGXMBkuSJCljNliSJEkZs8GSJEnKmA2WJElSxmywJEmSMmaDJUmSlDEbLEmSpIzZYEmSJGXMBkuSJCljNliSJEkZs8GSJEnKmA2WJElSxmywJEmSMmaDJUmSlDEbLEmSpIzZYEmSJGXMBkuSJCljNliSJEkZs8GSJEnKmA2WJElSxmywJEmSMmaDJUmSlDEbLEmSpIzZYEmSJGXMBkuSJCljNliSJEkZs8GSJEnKmA2WJElSxmywJEmSMmaDJUmSlDEbLEmSpIzZYEmSJGXMBkuSJCljNliSJEkZs8GSJEnKmA2WJElSxmywJEmSMmaDJUmSlDEbLEmSpIzZYEmSJGVseHQAKWdjgNWB8Qtq4XO+F5i6WM0ISSfVzyRgIun1NgHoIr3engaeAZ5Y8E+p1mywVBfDge2BbYHtgM2BdYFVW/z+R4G7gDuAKcB1wM3A3KyDSjWxCvBS0utta2BD0mtu+Ra+dwZwN+k1dwvp9XYd8GQeQaUIXdEBpCFYHTgY2BfYAxib8eM/D1wCXAicAzyU8eOXxQTonhYdol56JgMPRqfIWDewM3Ag8ArgxWT/HvIv0uvtAuAqYH7Gjy9J6sdo4B3ApcA80qWHImo+qdl6F619Qq+SCdDda2VZrBV9UDO0EfBt4AGKe731kj7Q/B+wVf7/i5LUXGsC3yRdQijyJN9XTSO94ayb5/9wgWywbLD68krgz0AP8a+5K4DXAcNy/T+WpAZZE/ghMIv4k/zSNQf4ObB2bv/3xbDBssFa3L7A34h/ffVVdwCHYKMlSR0rc2O1dM0CjgVWyOU3kT8bLBssgI1J85+iX0+t1B3A27HRkqSWVamxWroeBF6f/a8kdzZYzW6wRgJfBmYT/xpqt27HRkuSBrQK8A3S7dvRJ+2h1hmktYCqwgaruQ3WZsA/iX/NDLXuBY7ApYckaZGFI1YziT9JZ1kPALtk+HvKkw1WMxust1OPDzSL1+3A23BES1KDrUo9G6vFazbw3qx+YTmywWpWgzUM+C7xr488a2Gj5XZwkhpjOeAzwLPEn4SLqh9S7hO9DVZzGqxRwO+Jf00UVTcAe2Xym5OkkuomLRBa9GKFZanTSJOJy8gGqxkN1gqkxXKjXwsRdS5pvplUmDJ/qlZ97EXa3+9XlPONpwhvJo0clLXJUr2NBv4I7BkdJMirgJuAn9H6/qTSkNhgKU+bAecBF5M2g226A4BTcAKuijUcOB0vlQ0HjiRtMP1p0nQFKTc2WMrDqqRPijcB+wdnKZs3AD+NDqFG+T/SBs1KxgJfIS1Wehi+D0qqgOVInwybNIG90/pwh7/jPDgHq75zsN5D/HO97HU9zb10qhx1RQdQbRxAultu3eAcVTGfNLr3l+ggpAZrWnSIeumZTFrZP9JWwHWkOwc1uD8CRwH3RwdRPTg0qqFaAzgJOAebq3YMI83HWj06iGppFHAyNlftOAj4N/AFvBlFGbDBUqdGAMcAd5J2t1f7VibdWelIsrL2VdIIltqzHPB50l3PVdmJQSVlg6VO7EI6AR0LLB+cper2Bd4VHUK1sgXwoegQFbcFcCVpdH6V4CyqKBsstWMl4ATSiWfL4Cx1ciyexJWNLtJdqiOig9RAF2l0/jbg3TjSrDbZYKkVXaS7kW4H3oknmqytCHwjOoRq4UC8tJW1ScAvgKvwsqvaYIOlwWwBXA0cTzrRKB+HAptHh1CldQNfig5RYzuTpkZ8m7TtkDQgGyz1ZxjwceCfwMuCszTBMOBr0SFUaQfhCEvehgP/C/wL2CM2isrOBkt9WR+4lHTZytu8i3Mgbimkzh0dHaBB1iOdI4/D0Sz1wwZLi+smrTB+C7BbcJamKtMK76qOrYFdo0M0TBdwBGkl+J2Ds6iEbLC00PrAZcB3gDHBWZrszcBq0SFUOe+IDtBgGwJXAN8CRgdnUYnYYGnhp7CbcNSqDEYCb4sOoUrpJm0irjjDgI+SRv+9i1OADVbTrU3aC895BOVyWHQAVcoepC2rFO9FpCsBX8XtdhrPBqu53g3cDLwiOoiWsSUu2aDW7RcdQEsYDnwK+AfetNJoNljNMx44k7Rw3rjgLOrfAdEBVBn7RAdQn7YCriPt2erizA1kg9UsOwA3AK+PDqJBvSo6gCphEq59VWYjSFthnUfa3F0NYoPVDF3AUaQ9BNcLzqLW7AgsFx1Cpbc9jo5UwStJE+AdbWwQG6z6Wwk4B/geTrqskhHAdtEhVHo+R6pjFeBC0gLOw4KzqAA2WPX2ctLyC15uqqYdowOo9DaLDqC2dJG2ILsY7/ysPRusehoGfAG4CF/EVeadhBrM+tEB1JE98MNv7dlg1c+awCXA53EYuuo2jg6g0ls3OoA6tnD6xrdIUwJUMzZY9bI3cCOwe3QQZWLD6AAqtS68M63qukgrwF+FzXLt2GDVxxHA+aRPRaqHFfGTrfq3PGlRS1XfDsA/ceHnWrHBqr4xwCmk7W482dZLFzbM6p8LBdfLJOAC0pI6qgEbrGqbTBpadnPg+poYHUCl5fm7foaTltT5NTA6OIuGyBdode1M2obBdXDqzVFJ9Wd+dADl5lDgGmCd6CDqnA1WNb2PtGP7atFBlDsbLPVnXnQA5Wob4G+kD9OqIBusahkFHA/8GCc/N0VPdACV1jSgNzqEcrUacCnpJiZVjA1WdawM/Bl4T3QQFeqZ6AAqrXnA09EhlLuRpJuYjsPtzirFBqsatgFuwPWtmujZ6AAqtUejA6gwR5AWkV41OohaY4NVfvsCV5BWaFezzMURCg3s7ugAKtQuwLW4y0Ml2GCV2ztIWymMDc6hGPfjnWIa2J3RAVS49UhN1m7RQTQwG6xy6iJt1vwrnMzeZP+JDqDSuyU6gEJMJM3JfX10EPXPBqt8hgM/I23WrGbzzVOD+Vt0AIUZDZxB+jCuErLBKpcVgLPxllwlU6IDqPTuAKZGh1CYLtKH8e/j+3npeEDKYw3Stjf7RwdRadhgaTC9wJXRIRTuQ6TRLLfXKREbrHLYnDRpcevoICqNR4Hbo0OoEs6LDqBSeB1ph4+Vo4MoscGKtztwNbB2dBCVykW4Srdacz6u+K9kR9KVkPWjg8gGK9p+wAXAhOggKp0LogOoMh4hrZUnQVoj6xpgy+ggTWeDFecA4A/AmOggKp0ZpPXPpFadFB1ApbIqqeneITpIk9lgxXgLqblyQqL6cg7wfHQIVcrvcFslLWnhWlk7RwdpKhus4r0LOJm03pXUlxOjA6hyngd+ER1CpTOe1GTtFR2kiWywivVe4HhgWHQQldZdwF+iQ6iSvkvav1Ja3PKkO00Pig7SNDZYxfkY8FP8nWtg38M7wtSZB4FTokOolEYBp+PWOoXqig7QEB8HvhEdQqX3MLABMLPgnzsBuqcV/DNrrmcyqeEp2pqkUVBvnlFf5gPvwWkIhXA0JV9dwLewuVJrvkzxzZXq5SHgJ9EhVFrDgF/idmyFcAQrX98nbWEgDeYO0ro1EXNoHMHKXNgIFsBY4DZgraCfr/LrBT5CmpKgnDiClZ+vY3Ol1r0fJygrG88BR0WHUKl1kW6KOCY6SJ3ZYOXj88AnokOoMk4DLokOoVo5C/htdAiV3jfxcmFuvESYvaNJnwykVjwEvBiYGpjBS4SZC71EuNB44AZgveAcKrde4HDS3CxlyBGsbH0Amyu1rgc4jNjmSvX1DPA2YHZ0EJVaF/AzXMIhczZY2TmMNKldatUX8dKg8nUtXgLS4IYDvwFeHR2kTmywsvFW4AT8fap1Z5KWZZDydhJwbHQIld4I4AzgFdFB6sKGYOgOBn6Nv0u17lrgHaS5D1IRPoFzbDS40cAfgV2ig9SBTcHQ7E+6A8yNm9WqG0nPmxnRQdQovcCRpPOVNJCFexe+NDpI1dlgdW434PfAyOggqowbgX2Bp6ODqJHmA4cAv4gOotIbB1xAusNZHbLB6sxmpGHU0dFBVBlXAnsAjwfnULPNJ016/1p0EJXeisCfcZmPjtlgtW914HxgYnQQVcZpwH6k2+alaL3Ap0mjWe59qYGsShrJmhQdpIpssNozlnRtep3oIKqE+cBHSXeZ+kamsjkF2BX4T3QQldrGwNnAmOggVWOD1boRpFvrt4kOokq4D9gT+L/gHNJAppDm2fw8OohKbWfgdGBYdJAqscFqTRdwHGmCsjSQHtKqyFuS5l1JZfcc6Q7DA0kfDKS+vBr4TnSIKrHBas0XgHdGh1DpXQfsCPwP8HxwFqld5wCbA1/BS9rq24eA/40OURU2WIN7N/C56BAqtdtJ86x2Av4RnEUaihnAZ4ENgJ8Ac2LjqISOBd4YHaIKbLAG9krS5R6pL3cBhwJbkO4U7ImNI2XmYeD9wEak+Vk2Wlqom7R7yW7RQcrOBqt/25L2ZXKVdi3tLtLm3psBJ5PuFpTq6L+k+Vk2Wlrcwi11No0OUmY2WH1blfTkWSE6iEpl8cbqJGBebBypMDZaWtpE0hpZq0YHKSsbrGWNIN2OOjk6iErjP6Q3FxsrNd3CRmtD4AfArNg4CrYOcBYwKjpIGdlgLetHwO7RIVQK/yWNWG1I+tRuYyUl9wNHkS4RnYLzD5vsZcD3o0OUkQ3Wkt5H2qdLzfY88EXSm8dJOMdK6s99pC13XgpcFhtFgY4E3hsdomy6ogOUyM7ApcDI6CAKMw84gbQsx2PBWYo0AbqnRYeol57JwIPRKQK8Avgu6c5aNctcYB/g8uAcpeEIVrI26TqyzVVzXUzaBulImtVcSVnyddRcI4DfAetFBykLGyxYjnTH4CrRQRRiCvByYG/gluAsUh3MI81Z3Aj4Bk6Eb5JJpCbLjaGxweoCfoEbODfR08DRwA44pC3l4Vngk6TLhRcGZ1FxtiWtD9j4KUhNb7A+AbwlOoQK1Ut68W9MuvPFCexSvu4h7YpxIPBAcBYV43XAMdEhojW5wdqHtKmpmuNWYA/S9jaPx0aRGuccYEvgh/jBpgm+BuwXHSJSUxus1UmjGE39/2+a6cDHSZeCrwzOIjXZM8CHSJfm3Ri93oYBvwHWDc4RpokNxjDgVJzU3hRnk1ZgP5Z0G7GkeFOAHUkbSj8XnEX5mQj8lnSHYeM0scH6LOmuMdXb06RbxQ8irTotqVx6gJ+QLhteEpxF+dkB+Hp0iAhNa7B2Bz4THUK5u4B059LPo4NIGtR/ScukHEnaRUH18xHgNdEhitakBmtl0vXgYdFBlJtnSCfp/YGHgrNIal0v6QPRVrhsSh11Ab8ibQ7dGE1psLpJk9rXiA6i3PyFdKnBUSupuv4D7En6oDQ9OIuy1bj5WE1psD4G7BsdQrl4jnQy3g/X2JHqYOFo1vbAdcFZlK0dgS9GhyhKE1Za3QG4igZ1zQ0yBXgzcHd0kIpzs+fMNXaz56wNJ82b/SzNGRCoux7SwrN/iQ6St7o/YScCp2NzVTe9wA+Al2FzJdXZPOALpCsQj8ZGUUYaM2Wn7g3WCTRsUl0DTCVtuXEUMCc4i6RiXAxshxPg62IV4BRqftNZnRusw0hrIKk+rgK2Bs6NDiKpcA8DryCNaLnVTvW9nLQfcG3VdQ7WmsDNpEuEqr5e0v5lH8XV2PPgHKzMOQcrZ7uTduRYMzqIhmQeaeL7lOggeajjCFYXcCI2V3XxOOlT61HYXElKriDtLXpFdBANyXDSVJ5R0UHyUMcG6/2kN2RV342ku0AvjQ4iqXSeIJ3rvxkdREOyFfD56BB5qNslwvWBm4AVooNoyH4LvBuYER2kAbxEmDkvERbsbcDxwJjoIOpID+my79XRQbJUpxGsbtKlQZuraptPmvj4FmyuJLXmVGAX3Ni9qmr5/l2nButjwK7RITQkT5EWoHPIX1K7rgd2Av4WHUQdeRHw1egQWapLg7UF6dZdVdfNwEuBi6KDSKqsh4E9SBOnVT0fJO1FWQt1aLBGACdR07sQGuJc0qrs90QHkVR5s0nzNz9FWuJF1dFFao7HRQfJQh0arM+QbtdVNR0HvBZ4PjqIpFr5OnAI7vhQNesA34kOkYWqN1ibUvOVYGusF/g08F7SYnOSlLVTSfM6n4kOora8G9g/OsRQVbnB6gJ+BIyMDqK2zQEOBb4WHURS7V0K7Ix3GFbNT6n4XYVVbrDeSY0mwzXIc8BrSBt9SlIRbiVtyXJDdBC1bG3gc9EhhqKqC42uBNwOTIoOorY8SBr2vTk6iJbgQqOZc6HRkhoH/B53+6iKecB2wL+ig3SiqiNY38LmqmpuJ61RY3MlKcqzwAHA2dFB1JLhpBuhKtmrVDH0bsBh0SHUlttIl3P9RC8p2mzgdaTlfVR+OwJHRIfoRNUarJHAz6jupc0mmkJqih+JDiJJC8wH3gX8MjqIWvJNYI3oEO2qWoP1CdLSDKqGq0kjV1Ojg0jSUuYDhwPfiw6iQY0Dvh0dol1VarA2AD4ZHUItu5y0/syzwTkkqT+9wIeBL0cH0aDeQsXWxqpSg/VTYHR0CLXkPNILwdXZJVXB53DR6ir4MbB8dIhWVaXBegPeVlsVvydtfTMzOogkteGbVHzdpQZYlwpdyarCZPHRpLvQ1osOokGdQ7o7Z250ELXFdbAy5zpYFfYN4OPRIdSvOcAWwF3RQQZThRGso7G5qoKLgTdicyWp2j5BBSdUN8hI4NjoEK0o+wjWKsCdwPjoIBrQNcC+OOeqqhzBypwjWBXXRZr3e2R0EPVrb9IH+9Iq+wjWV7C5Kru/A/thcyWpPnqB9wGnRgdRv74NDIsOMZAyN1ibkTZ0VnndRLpb8LnoIJKUsR7SriGnRwdRn15MyXd1KXOD9T3SPkQqp1tJd3Y+FR1EknIyHzgU+HN0EPXpK8AK0SH6U9YG6wDS9VWV03+BfYAno4NIUs7mkJYKuj46iJaxOiVev6yMk9yHky49bRYdRH16BtgVuDk6iDLjJPfMOcm9hlYm3dCzQXQQLWEWsAnpg3+plHEE6/3YXJXVHNI6VzZXkprmCdL2X09EB9ESRgNfjw7Rl7I1WBOBz0aHUJ96gEOAS6KDSFKQu0kfMmdFB9ES3gzsGB1iaWVrsI4BJkWHUJ+OAc6IDiFJwa4iTXzviQ6iRbqA71KyaU9larBWAj4QHUJ9+hnwnegQklQSZwIfjQ6hJewIHBQdYnFlarA+AYyNDqFlnEWaFydJesF3gZ9Eh9ASvkiJ+pqyBFkd+J/oEFrGjTgULkn9OQq4LDqEFtmStCduKZSlwfoUsFx0CC1hKnAwMD06iCSV1DzSGln3RgfRIl+iJIuUl6HBmgwcHh1CS5gLvB74T3QQSSo5P4yWy4bA26JDQDkarM8Ao6JDaAlHAZdHh5CkiriJNJ2iNzqIAPgCMDI6RHSDtS7wjuAMWtKJwE+jQ0hSxZwFfDM6hIDUW4RvBB3dYH2BEnSZWuQa4L3RISSpoj4NnBsdQkAJro5FNlgbAW8P/Pla0oOkeQSzo4NIUkUt3PHC+avx1iZ4fndkg/UFYFjgz9cL5pJubX0sOogkVdzTpK1b5kQHEZ8CxkT98KgGa31KtFaF+DRwbXQISaqJ64BPRocQqwPvi/rhUQ3WMTh6VRYXAN+ODiFJNfNd4JzoEOIYYHTED45osFahBLP7BcBDeGuxJOWhl/Red19wjqZblaCeI6LB+hCB10S1SA+puXoyOogk1dQ00nysudFBGu5jBFw1K7rBWh6XASiLzwOXRoeQpJr7O+l8qzjrAwcV/UOLbrAOByYV/DO1rIuBr0WHkKSG+CbpvKs4Hyv6BxbZYI0APlzgz1PfppIuDfZEB5GkhugB3gU8Ex2kwV4K7F7kDyyywXoLaeEvxXof8Eh0CElqmAeAo6NDNFyho1hFNVhdOHpVBqcBZ0SHkKSGOpG0Z6Fi7A9sXdQPK6rBKvR/Sn16BPhgdAhJarj3Ao9Hh2iwwgZ7imqwjino56h/h5PmX0mS4jwBHBkdosEKm65URIP1YgqeWKZlHA+cFx1CkgTAH0lTNlS8EcBRRfygIhqssH2ABKRVhP83OoQkaQkfwBuOohwOjM/7h+TdYI0H3pbzz1D/ekm3Bj8XHUSStISn8K7CKGOBQ/L+IXk3WIeRVm9XjBOAy6JDSJL6dAZwbnSIhnofaYWD3OTdYB2R8+Orf1OBT0SHkCQN6APA9OgQDbQpsFuePyDPBmtPYPMcH18D+whu5CxJZfdf4KvRIRrqf/J88DwbrFyDa0BXAidHh5AkteRbwE3RIRroYGCNvB48rwZrdeA1OT22BjaHtJBdb3QQSVJL5pEuFXreLtYI0o1gucirwTqCFFzF+zrw7+gQkqS2XA38MjpEAx0BDM/jgfOYQT8c+A+wVg6PrYHdSVrYdVZ0EFXKBOieFh2iXnomAw9Gp1DlrEg6j0+KDtIwBwFnZ/2geYxgvRqbqygfxOZKkqrqKeBL0SEaKJc543k0WO/N4TE1uPOAv0SHkCQNyU+B26NDNMw+wAZZP2jWDdaawF4ZP6YGNxf4aHQISdKQeT4vXhc5DA5l3WC9DRiW8WNqcH7ikaT6OA/4c3SIhjmEjG/Oy7rBenvGj6fBTcNr9pJUNx8hLd+gYqxCulSYmSwbrG2ALTN8PLXmi6RtcSRJ9XEbLttQtEw3gM6ywcp9Z2ot4x7S5UFJUv18FngmOkSDvAaYkNWDZdVgDQfektFjqXUfIa3cLkmqnyeAb0eHaJDRwOuyerCsGqy9gdUyeiy15krgT9EhJEm5+j7wZHSIBsnsalxWDdahGT2OWvf56ACSpNw9R9oMWsXYDVgviwfKosEaBxyYweOodRcDl0eHkCQV4ofAw9EhGqILeGsWD5RFg/UGYLkMHket+1x0AElSYWYCx0aHaJDDyGCv5iwaLO8eLNb5wLXRISRJhToONxAvyobAS4b6IENtsNYAdh1qCLWsF0evJKmJZgFfjQ7RIEOeWz7UBus1GTyGWnc2MCU6hCQpxAnAfdEhGuJNpCWoOjbU5ujgIX6/WtcLfCE6hCQpzByci1WUlYDdh/IAQ2mwJg71h6stfwBuig4hSQp1IvB4dIiGeO1QvnkoDdaBZLzztAbkar6SpJmkZRuUv9cyhLsJh9JgDamzU1suxzsHJUnJj4Hno0M0wBrADp1+c6cN1nKk7XFUDK+5S5IWmgb8IjpEQ3Q8mNRpg/VKXFy0KDcDF0aHkCSVyv8Bc6NDNEDHN/N12mB5ebA4x5LuIJQkaaEHgdOjQzTABsAWnXxjJw3WCGD/Tn6Y2vYAvoAkSX3zA3gxOhrF6qTB2pO0RIPy9x0cApYk9e1m4LLoEA3Q0VW7ThosLw8W4xmcxChJGthPogM0wNbA+u1+UycNlpcHi3Ei3oYrSRrY2cBD0SEaoO3BpXYbrE2Bye3+ELWtl7RzuiRJA5mHVzuK8Jp2v6HdBmvfdn+AOnIZ8O/oEJKkSjgO5+vmbSdgfDvf0G6D5eKixfhpdABJUmU8AvwpOkTNDQf2aOcb2mmwRgK7tfPg6sgjpGvqkiS1yg/m+WtrkKmdBmsXYIX2sqgDP8ehXklSey4Bbo0OUXP7tPPF7TRYXh7Mn5MVJUmdOiE6QM1tCKzX6he302C11bmpI+eStj+QJKldvyF9UFd+Wu6FWm2wViIttKV8nRgdQJJUWY8CF0WHqLmWr+a12mDt08bXqjNTgQuiQ0iSKu3k6AA1txfpjsJBtdo0Of8qf6cBc6JDSJIq7Y+krdaUjwnAS1r5Qhus8vBThyRpqGYCZ0WHqLmW5mG10mBtCKw5tCwaxJ3AddEhJEm14Af2fLU06NRKg/WyIQbR4E6KDiBJqo0rgPujQ9TYDsDYwb7IBiteD3BKdAhJUm30AKdGh6ix4aQma0CtNFi7DD2LBnAF8N/oEJKkWvlddICa23mwLxiswZoIbJJNFvXjt9EBJEm1cz1wX3SIGhv06t5gDdZOLXyNOteDO6BLkvLh3YT52REYNtAXDNY8DToEpiH5K2nlXUmSsvaH6AA1Ng7YYqAvGKzBcoJ7vnzyS5Lycg3wSHSIGhtwEGqgBmsE8NJss2gpZ0cHkCTVltNQ8jXgINRADdbWwHLZZtFibgDujQ4hSao1r5Tkp+MRLOdf5csnvSQpb5cC06JD1NS6DLDTzUAN1k6ZR9Hi/hgdQJJUe3OBC6ND1Fi/vdJADVZLu0WrI3cDN0eHkCQ1wl+iA9RYv1f7+muwxpOGvpQPP01IkopyUXSAGmt7BGtroCufLAL+HB1AktQYDwG3Roeoqa1IexMuY6AGS/mYA1weHUKS1CiOYuVjDLBRX/+hvwbrxfllabxrgOejQ0iSGsUGKz999kw2WMVzsqEkqWiXA7OjQ9RUyw3WcGCzfLM0mg2WJKloM0j73yp7fU6r6qvB2hQYnW+WxnqStIK7JElF8zJhPrbp6y/7arCc4J6fi0h7Q0mSVLTLogPU1CrAqkv/ZV8N1lb5Z2msS6IDSJIa6wZgZnSImlpmcMoRrGJdFR1AktRYc4B/RIeoqWUmujuCVZwngbuiQ0iSGs2J7vkYtMFalXQtUdm7FuiNDiFJarRrogPU1KCXCDcsKEgT+alBkhTND/v52Ii0qvsiNljFscGSJEWbCtweHaKGhgMbLP4XSzdYG6A8zAb+GR1CkiT8wJ+XJQapbLCKcQMwKzqEJEk4DysvAzZYXiLMx9XRASRJWuD66AA1NeAlwhcVGKRJXHdEklQW/yatiaVs9TuCtSowrtgsjeH+g/s9sAAAIABJREFUg5KkspgD3BYdoob6HcHy8mA+pgP3RIeQJGkxN0UHqKE1gOUX/sviDZYT3PNxE27wLEkqFxus7HWx2FQrG6z83RgdQJKkpfjelI9FvZQNVv78lCBJKpsbcUX3PCyabmWDlb9/RQeQJGkp04AHo0PUUJ8N1poBQepuPnBzdAhJkvrgFZbsLXOJcDiwSkyWWrubdBehJEllc0d0gBpaf+EfFjZYq7LsoqMauluiA0iS1I+7owPU0Gos6KcWNlVeHszHXdEBJEnqhw1W9kYAK8ELDdYacVlqzSevJKmsHATIx+rwQoO1emCQOrPBkiSV1QPArOgQNbQG2GDlzU8HkqSy6gHujQ5RQ0uMYHmJMHszgEeiQ0iSNACvtGRviREsG6zs3Y2r5EqSys0rLdlbYgRrtcAgdeWnAklS2flelb0lRrBcpiF7fiqQJJXdQ9EBamjRCNaiNRuUqfuiA0iSNAjnCmdv0QjWJFzFPQ8PRweQJGkQvldlb1WgqxsYH52kpvxUIEkqu8eB+dEhamYksJINVn68ri1JKrt5pCZL2ZrYDYyLTlFD8/EJK0mqBq+4ZG+8I1j5eJz0qUCSpLJzHlb2bLBy4pNVklQVjmBlb5wNVj58skqSqsL3rOxNsMHKhxPcJUlV8XR0gBoa1w2MjU5RQ09GB5AkqUXPRAeoofHdwIToFDXkpwFJUlU8Gx2ghpyDlRM/DUiSqsL3rOyNdx2sfPhpQJJUFTZY2bPByokNliSpKmywsje+GxgVnaKGnIMlSaoKBwWyN64bGB6dooZ8skqSqsIRrOyN7QaGRaeoIRssSVJVzADmRoeomRHdwIjoFDXkJUJJUpXYYGVrmJcI8zE9OoAkSW2wwcrWcBus7M0HeqJDSJLUhvnRAWrGBisH86IDSJLUJt+7suUlwhz4JJUkVY3vXdlyBCsHPkklSVXje1e2bLBy4JNUklQ1zsHKlg1WDmywJElV43tXtmywcuCTVJJUNb53ZWt4N9AVnaJmfJJKkqqmOzpAzQzrBuZEp6gZr2NLkqpmVHSAuukGZkWHqJmR0QEkSWqTDVbGuoHZ0SFqxiepJKlqfO/K1nwvEWbPJ6kkqWpGRweomdmOYGXPBkuSVDVOb8mWDVYORuKdmZKk6nDJpuzZYOWgCxgRHUKSpBZ55SV7Nlg58Vq2JKkqbLCyN9tlGvJhgyVJqorlowPUkCNYOfHTgCSpKlaMDlBDs4djg5WHCcAD0SGkFs0Gfh4domaejw4gtWGl6AA1ZIOVEz8NqEpmQs+R0SEkhfE9K3uzu/GTVh4mRgeQJKlFE6ID1NDsbuCp6BQ15KcBSVJVTIoOUEOzu4Fp0SlqyAZLklQVvmdl72lHsPLhJUJJUlXYYGVvqg1WPnyySpKqwves7E2zwcqH17MlSVXhMg3Zs8HKiZcIJUlVsVZ0gBp6ygYrH6tHB5AkqQVd2GDlwTlYOVk7OoAkSS1YBbd3y8NT3cBzwNzoJDUzFhgXHUKSpEFMjg5QU091L/iDa2FlzyetJKnsfK/Kx6IGy8uE2fNJK0kqO9+rsjcPeNYGKz/Ow5IklZ0T3LP3NNC7sMF6JDJJTfmpQJJUdg4GZG8qwMIG64HAIHVlgyVJKjvfq7Jng5Uzn7SSpLLbKDpADT0INlh5Wi86gCRJA1gJt8nJgw1WztYBxkSHkCSpH5tGB6ip+8EGK0/dwIbRISRJ6ocNVj4egBcarEdwNfc8bBIdQJKkfvgelY8lGqwe4OG4LLXlpwNJUlnZYOVjiQZr0V8oUxtHB5AkqR82WNmbAzwONlh588krSSqjMaSbsZSth0hXBW2wcrYxS/6OJUkqg43w/SkPi3opG6x8LYfbEEiSymfr6AA11WeDdX9AkCbYLDqAJElL2TY6QE09uPAPizdYdwUEaYLtogNIkrSUl0QHqKlFg1WLN1h341pYebDBkiSVyTDgxdEhauq/C/+weIM1F7i3+Cy1t310AEmSFrMZaY6wsnf7wj8sfQfBvwsO0gRrAqtHh5AkaQE/+OdjNnDfwn9ZusG6HeXBJ7MkqSycupKPO4H5C//FBqsYPpklSWXhh/58LHEV0AarGD6ZJUllMBzYKjpETS3RQzkHqxg2WJKkMtiKtE2OsnfH4v+ydIP1LPBwcVkaY1VgcnQISVLj7R4doMYGHMFa5guUmZ2jA0iSGm+P6AA11Uua5L6IDVZx/NQgSYrUjR/283I/8Pzif2GDVRwbLElSpK2ASdEhamqZ3qmvBuvWAoI00SakuViSJEXwg35+WmqwppCuJSpbXcCu0SEkSY21R3SAGmupwXoGuCf/LI20R3QASVIjdeOH/DzdvPRf9NVgAfwz5yBNtVt0AElSI22B86/yMh+4cem/7K/BmpJvlsbaAlgpOoQkqXFeHh2gxv4NTF/6Lx3BKpbzsCRJEfaPDlBjffZM/TVY1wM9+WVptH2iA0iSGmUFvIMwT201WM8Cd+WXpdFeRRrJkiSpCHsDo6JD1Fif06r6a7D6/QYN2WRgy+gQkqTGeFV0gBqbB9zU13+wwYrhk12SVIQuYL/oEDV2KzCzr/8wUIPlRPf82GBJkoqwLbBmdIga67dXGqjBup60toOytyMu1yBJyp8f6PPV79W+gRqs54E7s88iYBiwb3QISVLt2WDlq6MRLIC/ZxxEL/BJL0nK02rA9tEhamwufWyRs9BgDdYV2WbRYvYFhkeHkCTV1usZ/H1enbsFmNXffxzsF39ltlm0mBVx6wJJUn7eHB2g5q4e6D8O1mDdC9yfXRYt5U3RASRJtTQZ2Ck6RM0NeJWvlaFDR7HyczAwMjqEJKl23oiXB/PUC1w10Be08st3HlZ+JpK2MJAkKUteIcnXrcDjA32BDVY8XwSSpCytj3cP5m3Q3qiVBusu4KGhZ1E/DgLGRIeQJNXGm0hb5Cg/mTRY4DysPI3FfaIkSdnxyki+emmhL2q1wbp8SFE0GF8MkqQsbAa8ODpEzd0OPDbYF7XaYDkPK18HACtEh5AkVd67owM0wOWtfFGrDdYdwMMdR9FglifdUitJUqdGAm+PDtEALQ06tbNGxoDrPWjI/NQhSRqKA4FVokPUXC85NFh/6SyLWvQy0rVzSZI68Z7oAA1wB/BoK1/YToN1PqlzU34cxZIkdWJtXLi6CJe1+oXtNFiPAlPaz6I2HAaMig4hSaqcd+HWOEU4v9UvbPdgnNfm16s9k0jX0CVJalU38I7oEA0wE7i01S+2wSofr6FLktqxN7BOdIgGuAyY0eoXt9tgTaHFyV3q2CvwhSJJat37owM0RFuDTO02WD3ABW1+j9rTDRweHUKSVAkbAq+KDtEQLc+/gs4mxHmZMH//AywXHUKSVHpH4+T2ItwC3NfON3RyUC4C5nTwfWrdisDbokNIkkptInBodIiGOLfdb+ikwXoWV3UvwtFAV3QISVJpHYH72Bal7at3nQ4repkwf5vhonGSpL4Nx8ntRZkG/K3db7LBKrcPRweQJJXSG4DJ0SEa4gJgXrvf1GmDdSdpPx7la19g0+gQkqTSOSo6QIN0tHrCUO48OGMI36vWdOGLSJK0pF2AHaJDNMR8Apan2py0+bOVb00HVmrxmEiS6u9C4t+bmlKXtHhMljGUEaxbSetCKF/Lke4olCRpB9L0ERXjt51+41AXJzt9iN+v1nyQtDaWJKnZPhsdoEHmAmd1+s1DbbBOG+L3qzXjgA9Fh5AkhdoG2D86RINcBEzt9JuH2mDdA1w/xMdQa44GJkSHkCSF+TwuQF2kIV2ly2L/Ii8TFmM86VKhJKl5NgdeHR2iQWYDZ0eHWBvoIX6mfxNqKjC2tcMiSaqRM4l/D2pS/aG1w9K/LEaw7gf+nsHjaHArAh+IDiFJKtQWwMHRIRpmyGt9ZtFggZcJi/QR3NxTkprka2T3fq3BzQDOGeqDZNlg9WT0WBrYSqQmS5JUf7vh3KuinQs8Hx1icZcTf820KfUssGpLR0WSVFVdpCk40e85TatMLsdmOeR4YoaPpYGNBT4XHUKSlKs3Ai+NDtEwzxGw9+BgxgDTiO88m1JzgU1bOjKSpKoZAdxF/HtN0+oXrRycVmQ5gjWTIezZo7YNB74cHUKSlIv3ARtEh2igX0YH6M/2xHefTaudWzoykqSqGAs8Rvz7S9Pq360cnFZlfdvnP4EbM35MDewb0QEkSZn6JLBKdIgGOj7LBxuW5YMtMAI3oyzS2sBNwO3RQSRJQ7YBcBJpGoiKMwd4B2kNrNKaAEwnfqivSXU3MLqVgyNJKrVziX9PaWJlvmB6HivDPk0Ge/ioLS8CPhEdQpI0JK8FXhUdoqFKO7l9aS8nvhttWs0CNmrl4EiSSmcMcC/x7yVNrPvJYcpUXnsbXU66bKXijAJ+EB1CktSRTwPrRYdoqF8C87N+0DwmuS80Ftgrx8fXsjYg3cV5R3QQSVLLNgBOxontEXqAdwLPRAdpx+qkWfnRQ39Nq/8Ay7VwfCRJ5XAB8e8dTa3ctsXJ6xIhwCPAmTk+vvq2LmkNFUlS+b0O2C86RIPlNrm9K68HXuAlwHU5/wwtazbwYrxUKEllNgm4BVgtOkhDPQisT9rbN3N5jmAB/AP4a84/Q8saBZxIvnPsJElD8wNsriL9iJyaq6K8jvhrrE2to1s4PpKk4h1A/HtEk2s6aQSx0oYB9xD/y2xiTQc2HPwQSZIKNJ50eSr6PaLJ9eNBj9IQ5X2JENLaErn/j6hPy5EuFRZxnCVJrfkBsGZ0iAbrBX4YHSIrY0lb6ER3rE2t9w9+iCRJBXgV8e8JTa9zBj1KFfM94n+pTa3nSfsVSpLijAceIP49oem152AHqmpeBMwj/hfb1LqY/JflkCT172Ti3wuaXjcNepQq6g/E/3KbXB8a/BBJknJwCPHvARa8Y5DjVFm7Ef/LbXLNArYe9ChJkrK0Pmmvu+j3gKbXY8DoQY5VZoq+u+xKXNk90ijgN7hXoSQVZQRwGjAuOoj4KWmgobZcXC2+jhv0KEmSsvB14s/5FswAVh3kWNXCP4j/ZTe93jToUZIkDcXueHNXWep7gxyr2ngt8b/sptc0YJ3BDpQkqSMrAQ8Rf6630mXBwhd2jVrh+4/Av4J+tpIJwEm4IbQkZa0b+BWwRnQQAfBLUrNbqMg31yeBNwb+fKURrC7gsuggklQjnwGOjA4hAOaQeo1no4MUqYs0ihU9dNj06gEOHuRYSZJaszfOuypT/Wzgw1VfbyH+l2+lzn7TQY6VJGlg65KuzkSf061Uc4D1BjpgdTYMuJ34g2DBv3GdFknq1BhgCvHncuuF+sWAR6wBDiX+IFipzsL9CiWpEycSfw63Xqh5wIYDHbAmGAbcQfzBsFJ9ZODDJUlayoeJP3dbS9avBzxiDfIO4g+GlWousMdAB0uStMgepPNm9LnbeqHmARsPcMwaZRhwC/EHxUr1GA2eGChJLdoYmEr8Odtask4d6KA1kXsUlqtuAyYOeMQkqbkmAXcSf662lqw5OPeqT5cQf3CsF+pyYORAB0ySGmgkaYHm6HO0tWz9YIDj1mgvIS18GX2ArBfqhAGPmCQ1SxdwMvHnZmvZehZYtf9DV6yy7UP3MGnByy2ig2iRbUhDrldHB5GkEvgS8MHoEOrTl4ELokOU2Xqkna+jO2HrheoB3jrQQZOkBngzXmUpaz0ELN//odNC3yf+YFlL1kxgp4EOmiTV2F7AbOLPxVbf9Z7+D50WtzLwNPEHzFqynsA9CyU1zw7Ac8Sfg62+61bKN+WJ7ugA/XgCODY6hJaxEnARrpElqTm2AM4HVogOon59HJgfHaJKxgD3E98ZW8vW3cDq/R86SaqFF5Fuvoo+51r91xX9Hj0N6F3EHzyr77oRmND/oZOkSlsDuIf4c63Vf/UAL+3vAGpg3cB1xB9Eq+/6K961Ial+JuH2bVWo0/s7gGrNdqSNG6MPpNV3XQyM6vfoSVK1LEf68Bh9brUGrhnAun0fQrXjZ8QfTKv/OhMY3u/Rk6RqGAtcSfw51Rq8PtXPMVSbVgQeJ/6AWv3XmcCI/g6gJJXc8ri/YFXqLrxykqnDiT+o1sB1Lj7pJVXPeOBa4s+hVmv1yr4PozrVjS+AKtR52GRJqo7xwN+IP3dardWZfR9GDdW2OOG9CnU+MLqfYyhJZTEB+Dvx50yrtZqOE9tz9RPiD7I1eF2ATZak8pqAywBVrT7R55FUZibihPeq1AWkFfklqUxWIy2WHH2OtFqv24GRfR1MZcsV3qtT15IW7ZOkMlgPuJP4c6PVXu3X18FU9rpwIbgq1U2kbSckKdJ2wGPEnxOt9uqMvg6m8rMJMJP4A2+1Vv8BNurzSEpS/vYCniX+XGi1V88Ca/VxPJWzjxN/8K3WayqwU59HUpLy81r8QF7Vem8fx1MFGIa32Fatnsdr6ZKK835gPvHnPqv9uow0JUhBNgdmEf9EsFqv2cBb+zqYkpSRbuDrxJ/vrM7qeWD9ZY5qhQyLDpCBJ0gHY8/oIGrZMOBg0hIOl5KOnyRlZXngt6Qt1lRN/wv8OTqEYDjwD+I7bqv9OgNYbtlDKkkdWRP4J/HnNqvz+itpBFIlsRXp0lP0E8Nqv24AJi97SCWpLdsADxB/TrM6r+nAhksfWMX7AvFPDquzegjYfpkjKkmteQPpzTn6XGYNrY5e+sCqHIYDU4h/glid1Uyc/C6pPV2kJXt6iD+HWUOra6nH3PDa2haYS/wTxeqseoDP4fV3SYObAPyR+POWNfSaCWyMSu9zxD9ZrKHV+biHoaT+bQPcQ/y5ysqmjkGVMAy4kvgnjDW0ug94KZK0pMNxZfY61ZV4abBS1gKeJP6JYw2t5pLmV0jSaOB44s9LVnY1DVgHVc7BxD95rGzqLGA8kppqQ+Am4s9FVrb1JlRZxxH/BLKyqduAzZDUNG8CniH+HGRlWz9HlTYa+BfxTyQrm5oJHIUbgEpNMA4/JNe1bsNdPGphC2AG8U8oK7v6M7AGkupqR+Au4s81VvY1C9ga1cZRxD+prGzrceA1SKqT4aRdOeYRf46x8qkPoVrpAv5E/BPLyr5OAlZAUtVtghs1170uwCketbQy8DDxTzAr+7oDeAmSqqgLeB/uJVj3epj0Pqya2guHnutac4FvAGOQVBUbAZcTf/6w8q35wN6o9j5G/JPNyq/uBHZHUpmNAD6JK7I3pb6CGqELOJ34J5yVX/WQ5matiKSyeTHwD+LPE1YxdRFuhdMoKwC3EP/Es/KtR4DXI6kMxpAu4ztNozl1H7ASapyNcXXgptTvgNWRFGUf4G7izwVWcTUD17tqtANJl5Oin4hW/vU8aX2dUUgqymTS5fro179VfL0DNd5XiH8iWsXVncCrkJSn5UgfaJzE3sz6PhLQDZxP/BPSKrYuAjZFUtZeDfyH+Ne4FVPXACORFpiI8wOaWHNIn7TGImmoNgEuJP51bcXVI7hPrPqwDW4K3dR6AHgn3kosdWIt4HjSYr/Rr2UrrmYDL0Pqx5tx0nuT61bgtbhXltSKScC38IOpler9SIP4DPFPVCu2/g7siaS+LA98HJhG/GvVKkcdj9SinxH/hLXi6yJgOyRB2t7mCNKmvdGvTas8dSEwHKlFI4FLiX/iWvHVA5wGbI7UTCOBdwP3Ev96tMpVNwHjkNo0DriZ+CewVY7qAc4BdkBqhlGkEav7iX/9WeWrh4G1kTq0HvAo8U9kq1x1NWmtH6mOVgCOAh4i/rVmlbNm4IdNZeAlwHTin9BW+Wpho+Vdh6qDSaTV16cS/9qyylvzgYOQMvJGXL7B6r+uB96OqxermjYCfgg8R/xrySp/fRgpY58i/oltlbseA75B2uBWKrMu4BXAGcA84l87VjXK5RiUm18Q/wS3yl+zgVNxjoLKZyzwQeAO4l8nVrXqfFyOQTkaBpxF/BPdqk79HXgb6Y4sKcrGpH03nyH+NWFVr6bgnq0qwEjgAuKf8Fa1ahpwHC5cquKMAd5AWjDXOaRWp3UnsCpSQZYDriL+iW9Vs24lbTWyMlL2tiM1805at4ZaDwDroLZ5a/nQjAcuB7YOzqHqmgP8BTgJ+ANpsrHUiTWAQ0grrm8YnEX18ASwG3B7dJAqssEautVJI1kvig6iynsE+D1wJml9rZ7YOKqANYDXAa8HdgG6Y+OoRp4GXg7cGB2kqmywsrE2qclyywBlZSrpjp0zSfP9HNnSQisDryTNrdoP7+pS9maQnltXRQepMhus7GwOXEFaCVnK0qOkka3fkU5482PjKMBk4GBSU7UTjlQpP3OAA4E/RwepOhusbL0EuARvZVV+nifN+zsHOI+0V5zqZxhpbuergQOAbfF8rfzNJy0pc3p0kDrwBZu9PUlvfKOjg6j2eklb9FxAupx4HY5uVdmapEt/ryStsD4uNo4aphc4grSYtjJgg5WPvYGzSWvQSEWZSlrv6ArSpcTbSCdNldMk0sT03UgN1VaxcdRwHwa+Fx2iTmyw8rMbaSRrheggaqxnSaNaFwN/XfDnOaGJmm01YFdSU7UzsA3OpVK8XlJz9f3oIHVjg5WvPYBzgeWDc0iQFp28ZkFNWVCPhiaqr9GkEaltgR1JjdX6oYmkZfUCHwJ+FB2kjmyw8rcLaX6ME99VRtNIlxKnLFZeWmzPSNLCntstVtvj3pMqN5urnNlgFcMmS1XyOHATcAdpBec7F/z5AZrdeI0nbZi8sDYiLc+yMemuP6kqeoD3AsdHB6kzG6zivIx0t5d3BqmqZpCarYUN112kpuvhBf+cGRctE8NIG9pOJq2Q/iJSE7URsAludqt66CHdLfjL6CB1Z4NVrB1Ii7eNjw4i5eApUrN1P2l9roX1FGnbjWlL/bMIY4CJwISl/rmwkVpzQa294O9cFV111kPaq/LE4ByNYINVvJeQmqyJ0UGkQL280GxNI41+zVrw32YAsxf8eTp93/k4sZ8/TyDdVDJxQTkPSkrmA+8ETo4O0hQ2WDG2Ay4EVooOIkmqvfnAocBvooM0iQ1WnI1JI1nrRAeRJNXWLNL2N2dFB2kaG6xYq5PuLtw6OogkqXaeJ20SflF0kCaywYo3AfgTaSFCSZKy8CiwP3BDdJCmcpuGeE+T9i78XXQQSVIt3Ev60G5zFcjF8cphPun6+GqkCfCSJHViCrAnaW06BbLBKo9e0ubQkPYwlCSpHZcBryStPadgNljlcznwJLAfzpGTJLXmD6QJ7dOjgyixwSqnf5C2IzkQj5EkaWA/BN4FzI0Oohc4QlJuuwG/xwVJJUnL6gE+BXwzOoiWZYNVfuuTlnHYPDqIJKk0pgOHkC4NqoRssKphLHAq8OroIJKkcA8BryHdMaiScn5PNcwBziBtXLtLcBZJUpwbgL2AO6KDaGA2WNXRC1xM+uSyHx47SWqaM0k3P02NDqLBeYmwmnYmXXdfOTqIJCl3vcCxwCcX/FkVYINVXS8iTX7fLDqIJCk3s4B3A7+JDqL22GBV2wTgt8C+0UEkSZlzMnuFOY+n2mYBp5GO467YMEtSXVwL7E1adFoV5BtyfewPnAysGB1EkjQkPwc+SLqDXBVlg1Uvk0nLOewYHUSS1LbngMOB06ODaOi8RFgvzwKnAOOAHYKzSJJa929gH+Dy4BzKiA1W/cwHLgTuJk1+HxkbR5I0iFNIk9kfjg6i7HiJsN42AX6H+xhKUhnNBj4OfD86iLJng1V/KwDHA2+ODiJJWuR+4A3AddFBlA8vEdbfHOD3pK0V9gSGx8aRpMY7m7Tl2T3RQZQfR7CaZTPStf5tooNIUgPNJG138wPc8qb2HMFqlieAE4AeYDdssCWpKDcDryRtcaYG8A22ufYETiStnSVJykcP8CPgGFw4tFFssJptPOmF//boIJJUQ/8FDgOuiA6i4nVHB1CoZ4BDgDcC04KzSFKdnEma72pz1VCOYGmhtYGTgN2jg0hShT0DvB84NTqIYjnJXQs9Q9oseiawKy7nIEntuoS03c1fo4MoniNY6ssGpN3cXx4dRJIqYDrwZeBbpEntkiNY6tNTpMuFDwN7AKNC00hSeZ0P7L/gn65tpUWc5K7+9JJGsTYB/hCcRZLK5nHSHYKvIm17Iy3BS4Rq1RuAHwMrRweRpGBnAu8DnowOovJyBEutOhPYmDSqJUlNdB9pD8E3YnOlQTiCpU7sB/wMWCc6iCQVoAf4BfC/wPPBWVQRTnJXJ+4mnWzGANvj80hSfd0EHAQch1vdqA2OYGmoNgK+R9rEVJLqYhrwRdLc03nBWVRBNljKyquB7wPrRQeRpCHoIa3C/lHSnYJSR7y0o6zcSZoAPwfYARgRG0eS2nY18Frgp6TFQ6WOOYKlPKwJfB14Oz7HJJXfw8AnSduFuVioMuGbn/K0O/ADYKvoIJLUh7mk0arPAM8FZ1HNeIlQefov6W7DJ0mXDcfExpGkRc4BDgR+g3cHKgeOYKkoY0krH396wZ8lKcI/SJcDL4kOonqzwVLRViYt1nc0biItqTh3AJ8FfofzrFQAGyxFWQf4FPBuvFQtKT8PAl8GTsD1rFQgGyxF2xz4PGkzaUnKylPAsaQbbWYGZ1ED2WCpLF5GWtpht+ggkiptOvAj4BvA08FZ1GA2WCqbg4AvAVtGB5FUKbOB44GvAI8FZ5FssFRaryDNm9gxOoikUpsN/Jp0vngwOIu0iA2Wym4X4OPAAdFBJJXKc8CvSJcCHwnOIi3DBktVsQ1p7ZrX4/NWarIngR+TNpefFpxF6pdvVKqaLYFjgLfi8g5SkzwGfBf4ITAjOIs0KBssVdX6wFHAkbhgqVRn9wHfA44DZsVGkVpng6WqW5u0Bc97gEnBWSRl51rSZcDf4wKhqiAbLNXFKOBNwEeAFwdnkdSZOcDZpBGra4KzSENig6U62o50+fAtwPDgLJIG9yhpqYUfAg8FZ5EyYYOlOludNEfr/cBKwVkkLWsKaSub04Bil2/jAAAC90lEQVS5wVmkTNlgqQm8fCiVh5cB1Qg2WGqanYF3Am8ExgZnkZrkNtLCoCcBjwdnkXJng6WmGg28GjgC2AtfC1IeniWNVp0EXAL0xsaRiuObigSTSQuXHgmsF5xFqroe0hILJwGnAtNj40gxbLCkF3QDewKHAq8DlouNI1XKQ8ApwPHAPcFZpHA2WFLfViSNar0F2AlfK1JfniNdAvw1cClp9EoSvmlIrViLNKL1BlKz1R0bRwo1g9RMnUlaZd1LgFIfbLCk9thsqYlsqqQ22WBJnVsTeD1wAPByYFhsHClTizdVv1vw75JaZIMlZWNN0sjWgcCuwMjYOFJHHgMuAM4C/gLMjo0jVZcNlpS95YCXkdbZeg2wTmwcqV89wA3AxcC5pJXVnaguZcAGS8rf+sArSA3X3qSte6QoTwKXkZqqP5E2WpaUMRssqViLj24dBKwdG0cN4CiVFMAGS4q1KWnO1q7Abthwaejmkxqqq4ArF/xzamgiqYFssKRyWYO0IfUuC/65Lb5ONbB5wE2kEaq/AlcD00ITSfLELZXcKsAOvNB0vRQYEZpI0aYDN5IaqYVN1czQRJKWYYMlVcs40qjW4rUxLnhaV9NJo1PXL1a3kkatJJWYDZZUfSuQmqzNge0W1PZ4t2LVPAf8C5iyWN1OmlMlqWJssKR6GgNsRRrh2grYiNSErRkZSgDMBe4lNU93kCakTwHuBnoDc0nKkA2W1CyjgA2AzUjrc61PGvnaChgbmKuOppEaqdtIl/XuXVC3ArMCc0kqgA2WpIUmk0a6NgLWI93RuPaCf64JjI6LVkpPAw8BDwAPL/jn3cCdpJGpZ+KiSYpmgyWpVSuTGq21FtTiDdhawCRgItW/y3EGafTpMZZtoB5aUPf/f3t3rAIgCEVh+B9aAoN6/4cUgqKaGhSug7jU+H8gHAT3g8O9uPxY0oAFS9LfErBSytbW5N7dXN8swNTJiShsbT6Ap+aTWEp8ESMLbqIE7ZQfp1zPKLvgWNJnL4J1k85laZUtAAAAAElFTkSuQmCC"/>
'     </p>
'
'     <span id="Notice"><h1>REBOOT $Verb</h1></span>
'     <div id="top_body">
'       <p><b>This computer must reboot to comply with the<br />
'         <a>
'           <span id="LinkSpan" onClick="OpenURL()">
'             <u>$Policy</u>.
'           </span>
'         </a></b>
'       </p>
'       <span style="display:$ShowMore">
'          <p>Automatic reboot in </p>
'          <div id="countdown">Unknown. Error.</div><br/>
'          <label>
'            <input type="checkbox" name='ShutdownCheck' onclick="SetShutdown()" $($Settings.Root.Shutdown)/>
'            Shutdown instead of reboot.
'          </label>
'       </span>
'     </div>
'     <p>
'       Click a button below to close this notice:<br/>
'       <hr style="Width: 75%" />
'       <span ID="TooLate">
'         <button type='button' id='SnoozeButton' onclick='RecordSettings("Snooze")' autofocus>Remind me later.</button><br />
'         <hr />
'         <span ID="NoReminder" style="display:$ShowMore">
'           <button type='button' id='OKButton' onclick='RecordSettings("OKquiet")'>OK. I will shutdown later.</button><br />
'           This dialog will return when only <span id="FinalPeriod"></span>&nbsp;hours remain.<br />
'           (at $(get-date $RebootPoint.addhours(-$Period) -Format f))
'           <hr />
'         </span>
'       </span>
'       <button type='button' id='GoButton' onclick='RecordSettings("GoAhead")'>I've saved my work.  Reboot now.</button><br />
'     </p>
'     <p>If you have questions or concerns about this reboot process, please contact 
'       <a href="mailto:$($Address)?Subject=Automatic%20Reboot" target="_top">$Address</a>.
'     </p>
'   </div>
'   <script type="text/javascript">
'   // Some variables set by calling script
'   var defperiod = $($Period*60);  // how many minutes between checks by default
'   var mins = $($Remaining.tostring('f0'));      // how many minutes until reboot (change this as required)
'   var period = $(Set-NextInterval $Period $Remaining);    // how many minutes between notices
'   var howMany = Math.round(mins * 60);    // total time in seconds
'
'   var innerWidth = document.body.offsetWidth;
'   var innerHeight = document.body.offsetHeight;
'   var CladWidth = 500 - innerWidth;
'   var CladHeight = 700 - innerHeight;
'
'   if (howMany <= 60) {
'     document.getElementById('TooLate').style.display="none";
'   } else if (howMany < 60*defperiod) {
'     document.getElementById('NoReminder').style.display="none";
'   }
'   var WinWidth = document.getElementById('Notice').offsetWidth*1.05 + CladWidth;
'   var WinHeight = document.getElementById('Content').offsetHeight*1.05 + CladHeight;
'   window.resizeTo(WinWidth,WinHeight);
'   window.moveTo(screen.availWidth - WinWidth,screen.availHeight-WinHeight);
'
'   if (period > 120) {
'     document.getElementById('SnoozeButton').value = "Remind me in " + Math.ceil(period/60) + " hours";
'   } else {
'     document.getElementById('SnoozeButton').value = "Remind me in " + Math.ceil(period) + " minutes";
'   }
'   document.getElementById('FinalPeriod').innerText = Math.ceil(defperiod/60);
'
'
'   beep();
'
'   function OpenURL() {
'     var shell = new ActiveXObject("WScript.Shell");
'     shell.run("$Purl",0);
'   }
'
'   // JavaScript Number prototype Property
'   // http://www.w3schools.com/jsref/jsref_prototype_num.asp
'   Number.prototype.toMinutesAndSeconds = function() {
'     Hrs = Math.floor(this/3600);
'     Mins = Math.floor(this/60)-Hrs*60;
'     Secs = this-(Hrs*60*60)-(Mins*60);
'     return ((Hrs>1)?Hrs+" hours ":"")+((Hrs==1)?Hrs+" hour ":"")+
'               ((Mins>1)?Mins+" minutes ":"")+((Mins==1)?Mins+" minute ":"")+
'               (((Secs)>=10)?Secs+" seconds":"0"+Secs+" seconds");
'   }
'
'   function display(seconds, output) {
'     // update screen with remaining time
'     output.innerHTML = (--seconds).toMinutesAndSeconds();
'     if(seconds > 0) {
'       if (seconds <= 60) {
'         beep();  
'         document.getElementById('TooLate').style.display="none";
'         window.focus();
'       } else if (seconds < 60*defperiod) {
'         document.getElementById('NoReminder').style.display="none";
'       }
'       // Recursive call after 1 second
'       window.setTimeout(function(){display(seconds, output)}, 1000);
'     }
'     if (seconds <= 0) {
'       RecordSettings("Auto");
'     }
'   }
'
'   // Call recursive function on start supplying initial time and
'   //   countdown <div> element
'   display(howMany, document.getElementById("countdown"));
'
'   </script>
'   </body>
'   </html>
'"@
'      
'      # Create and store the path to a one-time-use HTA file.
'      $HTApath = $env:TEMP + '\Reboot' + (Get-Date -Format 'yyyyMMddHHmmss') + '.hta'
'      # Create the file.
'      $HTACode > $HTApath
'      # Start the file.
'
'      # Before opening the HTA, create a script block that will run in parallel to the 
'      #    HTA notice window and display balloon notices (and eventually minimize all 
'      #    other windows to focus on the notice).
'      $BalloonNoticeScriptBlock = {
'         function Show-BalloonTip {
'            [CmdletBinding(SupportsShouldProcess = $true)]
'            param (
'               [Parameter(Mandatory=$true)][string]$Text,
'               [Parameter(Mandatory=$true)][string]$Title,
'               [ValidateSet('None', 'Info', 'Warning', 'Error')][string]$BalloonIcon = 'Info',
'               [string]$NoticeIcon = (Get-Process -id $pid | Select-Object -ExpandProperty Path),
'               [int]$Timeout = 10000
'            )
'
'            Add-Type -AssemblyName System.Windows.Forms
'
'            # This will allow for referencing any icon that can be seen inside a binary file.
'            #    https://social.technet.microsoft.com/Forums/exchange/en-US/16444c7a-ad61-44a7-8c6f-b8d619381a27/using-icons-in-powershell-scripts?forum=winserverpowershell
'            $code = @'
'               using System;
'               using System.Drawing;
'               using System.Runtime.InteropServices;
'
'               namespace System {
'                  public class IconExtractor {
'                     public static Icon Extract(string file, int number, bool largeIcon) {
'                     IntPtr large;
'                     IntPtr small;
'                     ExtractIconEx(file, number, out large, out small, 1);
'                     try { return Icon.FromHandle(largeIcon ? large : small); }
'                     catch { return null; }
'                     }
'                     [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
'                     private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
'                  }
'               }
''@
'            Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing
'
'            if ($script:balloon -eq $null) {
'               $script:balloon = New-Object System.Windows.Forms.NotifyIcon
'            }
'
'            $balloon.Icon            = [System.IconExtractor]::Extract('shell32.dll', 77, $true)  # 77 is the Warning triangle.  See note above
'            $balloon.BalloonTipIcon  = $BalloonIcon
'            $balloon.BalloonTipText  = $Text
'            $balloon.BalloonTipTitle = $Title
'            $balloon.Visible         = $true
'
'            $balloon.ShowBalloonTip($Timeout)
'
'            $null = Register-ObjectEvent -InputObject $balloon -EventName BalloonTipClicked -Action {
'                           $balloon.Dispose()
'                           Unregister-Event $EventSubscriber.SourceIdentifier
'                           Remove-Job $EventSubscriber.Action
'                        }
'         } # END function Show-BalloonTip
'
'         # This allows for checking if the front most window is the HTA
'         Add-Type  @'
'         using System;
'         using System.Runtime.InteropServices;
'         using System.Text;
'         public class UserWindows {
'            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
'               public static extern int GetWindowText(IntPtr hwnd,StringBuilder lpString, int cch);
'            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
'               public static extern IntPtr GetForegroundWindow();
'            [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
'               public static extern Int32 GetWindowTextLength(IntPtr hWnd);
'         }
''@
'
'         # Only display notice ballons if the Notification Window is NOT put in front
'         $ForgroundWindow = [UserWindows]::GetForegroundWindow()
'         $FGWTitleLength = [UserWindows]::GetWindowTextLength($ForgroundWindow)
'         $StringBuilder = New-Object text.stringbuilder -ArgumentList ($FGWTitleLength + 1)
'         $null = [UserWindows]::GetWindowText($ForgroundWindow,$StringBuilder,$StringBuilder.Capacity)
'         while ($StringBuilder.ToString() -notmatch $HTARegEx) {
'            Show-BalloonTip -Text $BalloonText -Title $BalloonTitle -BalloonIcon Warning
'            Start-Sleep -Seconds $NapTime
'            $ForgroundWindow = [UserWindows]::GetForegroundWindow()
'            $FGWTitleLength = [UserWindows]::GetWindowTextLength($ForgroundWindow)
'            $StringBuilder = New-Object text.stringbuilder -ArgumentList ($FGWTitleLength + 1)
'            $null = [UserWindows]::GetWindowText($ForgroundWindow,$StringBuilder,$StringBuilder.Capacity)
'         }
'
'         $script:balloon.Dispose()
'         Remove-Variable -Scope script -Name balloon
'      } # END $BalloonNoticeScriptBlock
'
'      # Balloon notice titles display only the first 62 characters.
'      If ($Settings) {
'         $bTitle = 'This computer will reboot in approximately ' + [math]::Round($Remaining/60) + ' hours!'
'      } else {
'         $bTitle = "This computer needs to reboot by $OrgName policy."
'      }
'      # Balloon notice text displays only the first 256 characters.
'      $bText = 'Software patches that require a reboot were recently initiated.  Until this computer ' +
'               'reboots, it will be vulnerable to malicious attacks or instability.  Please save your ' +
'               'work and reboot to ensure the security of this system and the entire network.'
'      # Display a balloon notice 10 times/default cycle time
'      $PauseTime = $Period*60*60/10  # number of seconds between balloons
'
'      $runspace = [runspacefactory]::CreateRunspace()
'      $runspace.open()
'      $ps = [powershell]::create()
'      $ps.runspace = $runspace
'      $runspace.sessionstateproxy.setvariable('BalloonText',$bText)
'      $runspace.sessionstateproxy.setvariable('BalloonTitle',$bTitle)
'      $runspace.sessionstateproxy.setvariable('HTARegEx',"$OrgName.*Reboot Notice")
'      $runspace.sessionstateproxy.setvariable('NapTime',$PauseTime)
'      $ps.AddScript($BalloonNoticeScriptBlock)
'      $ps.BeginInvoke()
'
'      # If the user has not acknowledged for nearly the entire minimum lead time or 
'      #    remaining time is running out, force user attention to the notice window.
'      if (((New-TimeSpan -Start $FlagDate -End (Get-Date)).TotalHours -ge ($MinLead - $Period/2)) -or
'            ($Remaining -le $Period*60)) {
'         (New-Object -ComObject Shell.Application).minimizeall()
'      }
'
'      # Now open the HTA
'      Start-Process $HTApath -Wait
'      
'      $runspace.Close()
'      $ps.Dispose()
'
'      # Allow other users to manipulate the acknowledgement file in case of an unclean-reboot
'      if (Test-Path $AppSettings) {
'         $Acl = get-acl $AppSettings
'         $rule = New-Object  system.security.accesscontrol.filesystemaccessrule('Authenticated Users','Modify','Allow')
'         $Acl.setaccessrule($rule)
'         set-acl $AppSettings $Acl
'
'         # re-read settings in case the user modified them
'         $Settings = [xml](get-content $AppSettings)
'      } #end if (test-path...
'      
'   } #end if (lastboot > 24 hrs and not quiet or time is running out) {show HTA}
'   
'   $Now = Get-Date
'   # Recalculate how much time remains (for long-displayed HTA)
'   $Remaining = (New-TimeSpan -Start $Now -End $RebootPoint).TotalMinutes
'
'   # The first time a user acknowledges, the reboot point needs to be set.
'   If ($Settings -and -not $Settings.Root.RebootPoint) {
'      # delay the reboot for a week if the acknowledgement is too late
'      if ($Remaining -lt $MinLead*60) {
'         $RebootPoint = $RebootPoint.AddDays(7)
'      }
'      $NewNode = $Settings.CreateElement('RebootPoint')
'      $NewNode.InnerText = $RebootPoint
'      $null = $Settings.Root.AppendChild($NewNode)
'      $Settings.Save($AppSettings)
'   }
'
'   # When the command comes down from the GUI, restart/poweroff
'   If ($Settings.Root.ActNow) {
'      If ($Settings.Root.Shutdown -eq 'Checked') {
'         Reset-AutoReboot $AppSettings $TaskName $OrgName -PowerOff
'         Return
'      } else {
'         Reset-AutoReboot $AppSettings $TaskName $OrgName -Restart
'         Return
'      }
'   }
'
'   if ($Settings) {
'      $LastStamp = Get-Date $Settings.Root.Acknowledgements.LastChild.'#text'
'      # Give the longest "snooze" available to a user who has asked 
'      # to not be bothered.  Also to a user who just now acknowledged the
'      # notice (rather than acknowledged previously but missed this notice).
'      if (($Settings.Root.quiet -eq 'True') -or 
'            ((New-TimeSpan -Start $LastStamp -End $Now).TotalMinutes -lt 1)) {
'         $NextInterval = Set-NextInterval $Period $Remaining
'      } else {
'         $NextInterval = 1
'      }
'   } else {
'      # If the HTA closes itself after X hours (or is killed), schedule to start again soon
'      $NextInterval = 1
'   }
'   
'} else { #End If(Is-RebootPending)   
'   # Clean up any crud that may be left behind
'   If ($Settings) {
'      Reset-AutoReboot $AppSettings $TaskName $OrgName -LogTime $LastBootTime.ToString('yyyy.MM.dd-HH.mm')
'   }
'   $NextInterval = $Period*60
'} #End Else
'   
'# Create or update a task to re-run this script at a later time.
'$TaskService = new-object -ComObject Schedule.Service
'$TaskService.connect()                     # connect to the local computer (default)
'$ErrorActionPreference = 'stop'
'Try {
'   $TaskFolder = $TaskService.GetFolder($OrgName)
'   
'} Catch {
'   $null = $TaskService.GetFolder('\').CreateFolder($OrgName) 
'   $TaskFolder = $TaskService.GetFolder($OrgName)
'} Finally { 
'   $ErrorActionPreference = 'continue'
'}
'if (($TaskFolder.gettasks(1) |Select-Object -expandproperty name) -icontains $TaskName) {
'   $TaskDef = $TaskFolder.GetTask($TaskName).definition
'   # Adjust the scheduled task to re-run the script at a later time.
'   $TaskDef.Triggers | % {
'      $_.StartBoundary = get-date (get-date).AddMinutes($NextInterval) -f 'yyyy\-MM\-dd\THH:mm:ss'
'   }
'   $TaskDef.Actions | % {$_.path = $ScriptPath}
'} else {
'   $TaskDef = $TaskService.NewTask(0)  # Not sure what the "0" is for
'   $Taskdef.RegistrationInfo.Description = 'Periodic check for pending reboot and GUI notice.'
'   $TaskDef.RegistrationInfo.Date = $FlagDate.tostring('yyyy\-MM\-dd\THH:mm:ss.00000')
'   $TaskDef.settings.priority = 2
'   $TaskDef.Settings.MultipleInstances = 3
'   $TaskDef.settings.StartWhenAvailable = $true
'   If ($Settings) { $Taskdef.Settings.WakeToRun=$True }   
'   # Create a trigger to run after the next time interval
'   $Trigger = $Taskdef.Triggers.Create(1)
'   $Trigger.Id = 'NextCheck'
'   $Trigger.StartBoundary = get-date (get-date).AddMinutes($NextInterval) -f 'yyyy\-MM\-dd\THH:mm:ss'
'   $Trigger.Enabled = $true
'   # Run this script again when triggered.
'   $Action = $Taskdef.Actions.Create(0)
'   $Action.Path = $ScriptPath
'}
'
'# Wake to reboot only if user acknowleged notice
'If ($Settings) { 
'   $Taskdef.Settings.WakeToRun=$True 
'} Else { 
'   $Taskdef.Settings.WakeToRun=$False 
'}
'
'# Finally, register the task
'$TaskFolder.RegisterTaskDefinition($TaskName, $Taskdef, 6, $null, $null, 3) > $null
'
'Return   
'
'
'# End PowerShell  (Don't modify this line!)

' Uncomment for testing delay to bring other windows forward
'WScript.Sleep(3000)

' Minimize impact on "No Reboot" machines.
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("c:\NoReboot") Then
  WScript.Quit
End If

Dim WinScriptHost, PoShswitches, PoShcmd 
PoShswitches = " -ExecutionPolicy ByPass -noprofile -Command " 

PoShcmd = "Invoke-Expression ($("
PoShcmd = PoShcmd & "$($Temp = ([System.IO.File]::ReadAllText('" & Wscript.ScriptFullName & "')); "
PoShcmd = PoShcmd & "$Temp.remove(($Temp.indexof('# End PowerShell')+1))"
PoShcmd = PoShcmd & ".remove(0,$Temp.indexof('# Start PowerShell')) "
PoShcmd = PoShcmd & "-replace '(?s)(\n)\x27','$1'"
PoShcmd = PoShcmd & "-replace 'CheckAutoReboot.vbs','" & Wscript.ScriptName & "')))"

Set WinScriptHost = CreateObject("WScript.Shell") 
WinScriptHost.Run "powershell.exe" & PoShswitches & CHR(34) & PoShcmd & Chr(34), 0, TRUE 

Set WinScriptHost = Nothing 