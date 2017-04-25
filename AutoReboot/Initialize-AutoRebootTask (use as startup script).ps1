# InitializeAutoReboot - Pairs with CheckAutoReboot.vbs as startup script
# Copyright 2015 Erich Hammer
# This script/information is free: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
#
# This script is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# The GNU General Public License can be found at <http://www.gnu.org/licenses/>.

$ScriptDir  = "$env:ProgramData\Institution\Reboot\"
$ScriptName = 'CheckAutoReboot.vbs'
$Author     = 'Domain\Administrator'

if (!(Test-Path -path $ScriptDir)) {New-Item $ScriptDir -Type Directory | Out-Null}
Copy-Item "\\server\DeployPoint`$\AutoReboot\CheckAutoReboot.vbs" -Destination $ScriptDir

$TaskName = 'Initialize auto-reboot checks'
$TaskService = New-Object -ComObject('Schedule.Service')
$TaskService.connect()                     # connect to the local computer (default)

if ((@($TaskService.getfolder('\').gettasks(1)) |select -expandproperty name) -icontains $TaskName) {
   $TaskService.getfolder('\').DeleteTask($TaskName,0)
}

$task_xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2015-05-19T15:56:23.89838</Date>
    <Author>$Author</Author>
    <Description>Runs a script once to start the process of periodic testing for the need to reboot</Description>
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
      <Command>$ScriptDir$ScriptName</Command>
    </Exec>
  </Actions>
</Task>
"@

$Task = $TaskService.NewTask($null)
$task.XmlText = $task_xml

$TaskService.getfolder('\').RegisterTaskDefinition($TaskName, $Task, 6, $null, $null, 3) > $null
