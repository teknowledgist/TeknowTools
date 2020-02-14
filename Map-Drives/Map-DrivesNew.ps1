$Version = '2.2'
<#
.SYNOPSIS
   A startup/login script for mapping multiple shares.
.DESCRIPTION
   When run as a startup script, this will create or modify a
   scheduled task to call the script to run again when a user logs in.
   When run as a logon script (always elevated), this will create a 
   scheduled task to immediately run the script again without elevation.
   When run from a scheduled task, this will compare the AD security 
   groups of the user with a all AD security groups matching specified
   properties that identify them as defining access to file shares.
   The server and share are identified from other properties of the
   security groups and the set of defined drive letters are used
   to map the shares.  
   
   Users can define shares they don't want to be mapped (even ones 
   not mapped by this script) using a "DoNotMap" file.  The can also
   allow blocked mappings on specific computers using a "KeepMapped" 
   file.  For example, don't map "home" when logging onto most
   computers (e.g. in a classroom), but do map "home" when logging
   onto the primary office computer.
.NOTES
   Copyright 2015-2019 Erich Hammer

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

#**************************************************
#region          Customizable part
<#**************************************************
      Define properties of custom mapping types:
****************************************************
   Name:          The name of the type or set of available shares
   Letters:       Agreed upon letters available for this type of share
   Server:        The server hosting the shares
   GroupsPath:    The FQDN for the OU of the Groups that define share access.
   FilterOn:      The property of the groups that identifies a group that
                  defines share access.  Defaults to "Description", but 
                  "Name" and "folderPathname" are good possibilities too.
   FilterString:  A string contained in the FilterOn property that defines 
                  the group as one to use for mapping
   PathProperty:  The property of the groups that defines the share path.
                  If this contains the "\\server\" substring, that will
                  be stripped before combining with the "Server" item above.
   PathPrefix:    Regular Expressions to strip from the the path property
       and        to end up with the entire "\\server\share" string or
   PathPostfix:   just the "share" portion (to be combined with the
                  Server item above).
   Trailers:      Group CN names to map after other shares.  This is to 
                  prevent shares with regularly changeing membership causing
                  users to get different drive letters over time.
#>
$Localization = @(
@{
   Name         = 'Department'
   Letters      = 'O','P'
   Server       = 'storage.domain.org'
   GroupsPath   = 'OU=Departments,OU=Groups,DC=domain,DC=org'
   FilterOn     = 'Description'
   FilterString = '*\\*'
   PathProperty = 'Description'
   PathPrefix   = '(Map|Read Only|\?):'
   Trailers     = 'npFile - DeansOffice',
                  'npFile - ApprenticeProg'
}
@{
   Name         = 'Research'
   Letters      = 'R','S','Q'
   Server       = 'research.domain.org'
   GroupsPath   = 'OU=Research_Labs,OU=Groups,DC=domain,DC=org'
   FilterOn     = 'Description'
   FilterString = '*\\*'
   PathProperty = 'Description'
   PathPrefix   = '(Map|Read Only):'
   Trailers     = 'npFile - SpecialLab'
}
@{
   Name         = 'CoreFacility'
   Letters      = 'I','J','K'
   Server       = 'data.domain.org'
   GroupsPath   = 'OU=CoreFacilities,OU=Groups,DC=domain,DC=org'
   FilterOn     = 'Description'
   FilterString = '*\\*'
   PathProperty = 'Description'
   PathPrefix   = '(Map|Read Only):'
}
)
# End of custom mapping types
#****************************************************

#****************************************************
#    User control of share mapping
#****************************************************
# If a user does not want a share(s) automatically mapped, they can place the
# \\server\share name(s) in this file (one per line).  The only location that
# makes sense to store this to be accessible whenever the user logs in is 
# the User's "HomeDirectory" 
$NoMapFile = 'DoNotMap.txt'
# If a user would like the "protected" shares to be mapped on specific computers
# (e.g. their office computer), place the \\server\share name in the following 
# file in the %APPDATA% location of the user profile.
$OKmapFile = 'KeepMapped.txt'

#****************************************************
#    Other variables to customize
#****************************************************
# The short name of the organization.  This is used in several places including
#   for separating files and tasks from other applications and system processes.  
$OrgName = $env:USERDOMAIN

# Members of this group will see verbose messages and a pause before closing the shell.
$VerboseGroup = 'Test - TroubleshootScript'

# A log file will be (over)written to the %APPDATA% location of the user profile.
$LogFile = 'DriveMapping.log'

#endregion 
#**************************************************

#**************************************************
#region       Define Functions
#**************************************************

Function Write-Log { 
   Param(
      [Parameter(Mandatory = $true, Position = 0)][string]$path,
      [Parameter(Mandatory = $true, Position = 1)][string[]]$info,
      [Switch]$ToScreen
      )
   if (($VerbosePreference -ne 'SilentlyContinue') -and $ToScreen) { 
      Write-Verbose ($info -join "`n") 
   }
   Add-content $path "($(Get-Date))  $($info -join `"`n`")"
} # End of Write-Log function

Function Is-Elevated {
  # This will determine if this script is running under elevated credentials.
  # Elevation is not a "normal" state for a script unless it is a Group 
  # Policy logon script.  For Windows Vista and newer (assuming UAC is 
  # enabled), a check for an administrator role will return true only if 
  # the script is running elevated.  Limited User Accounts (LUAs) require
  # checking the current integrity level (Some additional information:
  # is at: http://msdn.microsoft.com/en-us/library/bb625963.aspx)
  # Windows XP and earlier does not have UAC, so there is no elevation.
  
   If ([int](Get-WmiObject Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
      # Operating system is Vista or newer
      $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
      $principal = New-Object System.Security.Principal.WindowsPrincipal( $identity )
      $admin = [Security.Principal.WindowsBuiltInRole]::Administrator
      # One-line method:
      #([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

      if ($principal.IsInRole( $admin ) -or ([bool]((& "$env:windir\system32\whoami.exe" /all) -match 'S-1-16-8448'))) {
         # script is running elevated
         return $true
      } else {
       return $false
      }
   } else {
      # Operating system is XP or lower
      return $false
    }
} # end Is-Elevated


Function Create-ScheduledMappingTask {
   # Makes a scheduled task to run a script either:
   #    upon user login -- when running as startup script
   #    immediately     -- when running as logon script
   # This is necessary due to UAC.  See here:
   #    https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-vista/cc766208(v=ws.10)#group-policy-scripts-can-fail-due-to-user-account-control

   Param([Parameter(Mandatory = $true, Position = 0)][string]$TaskName,
      [Parameter(Mandatory = $true, Position = 3)][string]$ScriptFile
   )

   $Arguments = '-executionPolicy Bypass' + 
                ' -noprofile' +
                " -File `"$ScriptFile`""
   # Keep this around in case there is a future need to remain in powershell
   #$Arguments = "-noexit " + $Arguments

   if ("$env:computername`$" -eq $env:username) {  # running as SYSTEM
      # The principals can see/run the script.  Want that to be "users"...
      $Principals = '<GroupId>S-1-5-32-545</GroupId>' +
                    '<RunLevel>LeastPrivilege</RunLevel>'
      $Author = 'Interactive'
      # The task should run on login.
      $Trigger = '<LogonTrigger>' +
                   '<Enabled>true</Enabled>' +
                   '<Delay>PT10S</Delay>' +
                   '<ExecutionTimeLimit>PT5M</ExecutionTimeLimit>' +
                 '</LogonTrigger>'
   } else {
      # or the current user.
      $Principals = "<UserId>$env:USERDOMAIN\$env:USERNAME</UserId>" +
                    '<LogonType>S4U</LogonType>'
      $Author = "$env:USERDOMAIN\$env:USERNAME"
      # The user is logging in (and elevated), so run this immediately.
      $Trigger = '<RegistrationTrigger>' +
                   '<Delay>PT10S</Delay>' +
                   '<Enabled>true</Enabled>' +
                 '</RegistrationTrigger>'
   }

   $LogonTask_xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>$(get-date -Format yyyy-MM-ddTHH:mm:ss.00000)</Date>
    <Author>$Author</Author>
    <Description>Map defined shares to allotted letters ($Version)</Description>
  </RegistrationInfo>
  <Triggers>
    $Trigger
  </Triggers>
  <Principals>
    <Principal id="Author">
      $Principals
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>StopExisting</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>false</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
    <Priority>4</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>$Arguments</Arguments>
    </Exec>
  </Actions>
</Task>
"@

   $Task = $TaskService.NewTask($null)
   $task.XmlText = $LogonTask_xml
   try {
      $result = $TaskFolder.RegisterTaskDefinition($TaskName, $Task, 6, $null, $null, 3)
   } catch {
      $result = $_.Exception.Message
   }
   if ("$env:computername`$" -eq $env:username) {  # running as SYSTEM
      # Want users to have the ability to manually (re)run the task.
      #   (This may not work in Windows 10.)
      $TaskFile = Join-Path $env:SystemRoot "System32\Tasks\$($TaskFolder.path)\$TaskName"
      if (Test-Path $TaskFile) {
         $Acl = get-acl -Path $TaskFile
         $rule = New-Object -TypeName system.security.accesscontrol.filesystemaccessrule -ArgumentList ('Authenticated Users','ReadAndExecute','Allow')
         $Acl.setaccessrule($rule)
         set-acl -Path $TaskFile -AclObject $Acl
      }
   }
   return $result
}


#**************************************************
#endregion    End of function definitions
#**************************************************

$host.ui.RawUI.WindowTitle = 'Mapping drives.  DO NOT CLOSE!'
$Host.UI.RawUI.WindowSize = @{Width = 80; Height = $Host.UI.RawUI.WindowSize.Height}
$Host.UI.RawUI.BufferSize = @{Width = 80; Height = $Host.UI.RawUI.BufferSize.Height}
Write-Host '** Standard drive mapping script **' -ForegroundColor Green
Write-Host "Mapping file shares... This window should close soon.`n" 
if (Test-Path "Filesystem::$env:HomeShare") {
      $msg = "`n****`nIf you would like to control to which computers any of your `n" +
             "shares will map, contact your technical coordinator.`n****`n"
      Write-Host $msg -backgroundcolor yellow -foregroundcolor black
}

$RunningAsSystem = ("$env:computername`$" -eq $env:username)

# Initialize the task service and folder
$TaskService = new-object -ComObject('Schedule.Service')
$TaskService.connect()

$ErrorActionPreference = 'stop'
Try {$TaskFolder = $TaskService.GetFolder($OrgName)}
Catch { 
   $rootFolder = $TaskService.GetFolder('\') 
   if ($RunningAsSystem) {
      $null = $rootFolder.CreateFolder($OrgName) 
      $TaskFolder = $rootFolder.GetFolder($OrgName)
   } else {
      $TaskFolder = $rootFolder
   }
}
Finally { $ErrorActionPreference = 'continue' }

# For finding the path of this script regardless of PoSh version
#   see: https://stackoverflow.com/questions/817198
if ($PSCommandPath -eq $null) { 
   $PSCommandPath = $script:MyInvocation.MyCommand.definition
}

if ($RunningAsSystem) {
   # Start new log for each boot
   $LogFilePath = Join-Path $env:windir "Logs\$LogFile"
   If (Test-Path $LogFilePath) {
      $FileObj = Get-ChildItem $LogFilePath
      $PrevLogFilePath = (join-path $FileObj.directoryname $FileObj.basename) + '.previous' + $FileObj.extension
      if (Test-Path $PrevLogFilePath) {
         Remove-Item $PrevLogFilePath -Force
      }
      Rename-Item $LogFilePath $PrevLogFilePath -Force
   }

   $ScriptsDirectory = Join-Path $env:ProgramData "$OrgName\Scripts"
   if ((-not (Test-Path $ScriptsDirectory)) -or 
       (-not (Test-Path (Split-Path $ScriptsDirectory)))) {
      $info = New-Item (Join-Path $env:ProgramData "$OrgName\Scripts") -ItemType Directory -Force
      Write-Log $LogFilePath $info
   }
   try {Copy-Item -Path $PSCommandPath -Destination $ScriptsDirectory -Force -ErrorAction Stop}
   catch { Write-Log $LogFilePath $_.Exception.Message }
   
   $ScriptPath = Join-Path $ScriptsDirectory (Split-Path $PSCommandPath -Leaf)

   $AllUsersTaskName = 'Map Network Shares - All users'
   $xml = [xml]$TaskFolder.GetTask($AllUsersTaskName).xml
   if ($xml.Task.RegistrationInfo.description -match '\([0-9.]+\)$') {
      $taskversion = [version]$Matches[0].trim('()')
   }
   if ($taskversion -lt [version]$Version) {
      $info = Create-ScheduledMappingTask -TaskName $AllUsersTaskName -ScriptFile $ScriptPath
      Write-Log $LogFilePath $info
   }
   Return
}

# No point in running for local accounts
if ($env:USERDOMAIN -eq $env:computername) {
   return
}

$LogFilePath = Join-Path $env:APPDATA $LogFile

# If there is a previous log file...
If ( Test-Path $LogFilePath ) {
   $LastLine = (Get-Content $LogFilePath)[-1]
   # and the script had an error or reached its conclusion...
   If (($LastLine.length -gt 256) -or ($LastLine -eq '--')) {
      # rename it to allow for the next log
      $FileObj = Get-ChildItem $LogFilePath
      $PrevLogFilePath = (join-path $FileObj.directoryname $FileObj.basename) + '.previous' + $FileObj.extension
      if (Test-Path $PrevLogFilePath) {
         Remove-Item $PrevLogFilePath -Force
      }
      Rename-Item $LogFilePath $PrevLogFilePath -Force
   }
}

if (Is-Elevated) {
   if ($Global:TaskFolder.gettasks(1) | Where-Object {$_.Name -match '^Map Network Shares'}) {
      Return
   }

   $LogFileHeader =  "*** Log of drive mapping script run as logon script ***`r`n" +
                     "$(' '*22)*******************************************************`r`n" +
                     "$(' '*22)This is running elevated.`r`n" +
                     "$(' '*22)Creating scheduled task to re-run as unelevated user.`r`n`r`n"
   Write-log $LogFilePath $LogFileHeader

   $TempScriptPath = (Join-Path $env:TEMP 'TempMapDrives.ps1')
   try { Copy-Item -Path $PSCommandPath -Destination $TempScriptPath -Force -ErrorAction Stop }
   catch { Write-Log $LogFilePath $_.Exception.Message }

   $info = Create-ScheduledMappingTask -TaskName "Map Network Shares - $env:username" -ScriptFile $TempScriptPath
   Write-Log $LogFilePath $info
   Return
}

$LogFileHeader =  "*** Log of drive mapping script run as scheduled task ***`r`n" +
                  "$(' '*22)*********************************************************`r`n"
Write-log $LogFilePath $LogFileHeader

# Create a type for easier referencing
Add-Type -TypeDefinition @'
  public struct MapType {
    public string Name;
    public string[] Letters;
    public string Server;
    public string GroupsPath;
    public string FilterOn;
    public string FilterString;
    public string PathProperty;
    public string PathPrefix;
    public string PathPostfix;
    public string[] Trailers;
  }
'@

# Requires PoShv3:
# Update-TypeData -TypeName MapType -DefaultDisplayPropertySet name,letters,Server,FilterOn,pathproperty

$MapTypes = @()
Foreach ($Item in $Localization) {
   $MapTypes += New-Object MapType -Property $Item
}

# Test network connection
Write-Log $LogFilePath "Testing network connection to primary server for '$($MapTypes[0].Name)' shares." -ToScreen
$maxtries = 2
For ($i=1; $i -le $maxtries; $i++) {
   if (-not (Test-Connection $MapTypes[0].server -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
      Write-Log $LogFilePath "Server could not be reached on attempt $i."
      if ($i -lt $maxtries) {
         $msg = "`nServer or network currently unavailable.`n" +
                "Another attempt will occur in one minute to allow:`n" +
                "  * an in-progress, on campus WiFi connection to complete`n" +
                "  * an (optional) log into VPN.`n" +
                "(To cancel, press Ctrl-C.)`n"
         Write-Host $msg -foregroundcolor red
         Start-Sleep 60
      }
   } else {
      Write-Host "`nServer \\$($MapTypes[0].server) is available."
      Write-Log $LogFilePath "Server \\$($MapTypes[0].server) successfully pinged."
      $i = 10
      $Connected = $true
   }
}

if (-not $Connected) {
   $msg = "`nNo connection to the primary file server was found.`n" +
            "`tDrive mapping will not be attempted."
   Write-Host $msg
   Write-Log $LogFilePath "Server \\$($MapTypes[0].server) is NOT available.  Exiting."
   Write-Log $LogFilePath "End of script`n--"
   if ($VerbosePreference -ne 'SilentlyContinue') {
      Write-Verbose 'Press any key to exit...'
      $Null = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
   }
   return
}

# Build a list of shares that should NOT be mapped for this user on this computer.
$NoList = @()
if ($env:homeshare) {
   $NoMapFilePath = Join-path $env:homeshare $NoMapFile
   if (Test-Path "Filesystem::$NoMapFilePath") {
      # First, collect requests for not mapping
      $NoList = @(Get-Content ($NoMapFilePath) | Where-Object {$_ -like '\\*'})

      $OKmapFilePath = Join-path $env:APPDATA $OKmapFile
      if ($NoList -and (Test-Path $OKmapFilePath)) {
         $OKList = Get-Content ($OKmapFilePath) | Where-Object {$_ -like '\\*'}
         if ($OKList) {
            $NoList = @(Compare-Object $NoList $OKList | 
                        Where-Object {$_.SideIndicator -eq '<='} | 
                        Select-Object -ExpandProperty inputobject)
         }
      }
      if ($NoList) {
         $msg = "`n$($NoList.count) share(s) should not be mapped on this " +
                  "computer as requested in:`n`t '$($NoMapFilePath)'`n"
         Write-Host $msg -backgroundcolor yellow -foregroundcolor black
         Write-Log $LogFilePath "Do not map (or unmap) $($NoList -join(', '))."
      }
      Start-Sleep 3
   }
}

# Get the entire list (direct and indirect) groups of which the user is a member
$UsersGroups = & "$env:windir\system32\whoami.exe" /groups /fo csv | convertfrom-csv | ForEach-Object { $_.'group name'.split('\')[1]}

#Check if debugging is neccessary
if ($UsersGroups -contains $VerboseGroup) {
   $VerbosePreference = 'Continue'
   Write-Log $LogFilePath 'Verbose output enabled.' -ToScreen
}
$msg = "$($ENV:Username) is a direct or indirect member " +
       "of $($UsersGroups.count) AD and local groups in total."
Write-log $LogFilePath $msg -ToScreen

# Initialize the ADO connection
$adoCommand = New-Object -comObject 'ADODB.Command'
$adoConnection = New-Object -comobject 'ADODB.Connection'
$adoConnection.Provider = 'ADsDSOObject'
$adoConnection.Open('Active Directory Provider')
$adoCommand.ActiveConnection = $adoConnection

Foreach ($MapType in $MapTypes) {
   Write-Host "`nAttempting to map any '$($MapType.Name)' shares accessible to $env:username..." -Foregroundcolor Cyan
   # Help with "key" SQL condition (from http://www.rlmueller.net/ADOSearchTips.htm):
   #     Conditions can be nested using parenthesis. 
   #     The "*" wildcard character is OK. However, it cannot be used with 
   #     Distinguished Name attributes (attributes of data type DN), such as 
   #        distinguishedName, memberOf, directReports, and managedBy
   #     Escape the following caracters with the backslash escape character, "\", 
   #     and the 2 digit ASCII hex equivalent of the character:
   #        (  becomes  \28
   #        )  becomes  \29
   #        \  becomes  \5C
   $MapType.FilterString = $MapType.FilterString -replace '\\','\5C' -replace '\(','\28' -replace '\)','\29'

   # Get list of groups for mapping shares
   $SQLstring = "SELECT Name, $($MapType.PathProperty) " +
               "FROM 'LDAP://$($MapType.GroupsPath)' " +
               "WHERE objectCategory='group' AND $($MapType.FilterOn)='$($MapType.FilterString)' "+
               'ORDER By Name'
   $adoCommand.CommandText = $sqlstring
   $adoRecordSet = $adoCommand.Execute()

   Write-Log $LogFilePath "Groups discovered of type '$($MapType.Name)':  $($adoRecordSet.RecordCount)" -ToScreen

   if ($adoRecordSet.RecordCount -eq 0) {
      # Move to the next MapType
      Continue
   }

   $Properties = @('Name')
   if ($Properties -notcontains $MapType.PathProperty) {
      $Properties += $MapType.PathProperty
   }

   $FileShareGroups = @()
   do {
      $Name = $adoRecordSet.Fields.Item('Name') | Select-Object -ExpandProperty value
      $PProp = $adoRecordSet.Fields.Item($MapType.PathProperty) | Select-Object -ExpandProperty value

      if ($PProp) {
         # May as well store the path info from creation.
         # First strip any potential prefix or postfix,
         $Path = $PProp -replace "^$($MapType.PathPrefix)",'' -replace "$($MapType.PathPostfix)`$",''
         
         if ($Path.Trim()) {
            # A path starting with "\\" is likely complete (and might use a different server) so
            #   it should be left alone.  The possible "paths" are "server\share", "\share" or "share".
            if ($Path -notlike '\\*') {
               if ($Path -notlike "$($MapType.Server)*") {
                  $Path = Join-path "\\$($MapType.Server)" $Path
               } else {
                  $Path = "\\$Path"
               }
            }

            $Obj = New-Object -TypeName psobject
            $Obj | Add-Member -MemberType NoteProperty -Name 'Name' -Value $Name
            $Obj | Add-Member -MemberType NoteProperty -Name 'MapPath' -Value $Path

            $FileShareGroups += $Obj
         }
      }
      $adoRecordset.MoveNext()
   } Until ($adoRecordset.EOF)
   $adoRecordset.Close()

   Write-Log $LogFilePath "Groups of type '$($MapType.Name)' with a potential path:  $($FileShareGroups.Count)" -ToScreen

   if ($FileShareGroups.Count -eq 0) {
      # Move to the next MapType
      Continue
   }

   $MatchingGroups = @($FileShareGroups | Where-Object {$UsersGroups -contains $_.Name})

   $msg = "User's membership in groups of type '$($MapType.Name)':  $($MatchingGroups.count)"
   Write-Log $LogFilePath $msg -ToScreen

   if (-not $MatchingGroups) {
      Continue
   } else {
      # The path could be a full, folder path (i.e. \\server\share\folder) identifying the
      #   folder the group can access.  Reduce it to just \\server\share for mapping.
      $MatchingGroups | 
         ForEach-Object { $_.MapPath = $_.MapPath -replace '(\\\\[^\\]+\\[^\\]+).*','$1' }
   }

   if ($MapType.Trailers) {
      # First isolate the trailing groups
      $AftGroups = $MatchingGroups | 
                     Where-Object {$MapType.Trailers -contains $_.Name}
      # Then sort them in the order of the list of trailers
      [array]$AftMaps = foreach ($Trailer in $MapType.Trailers) { 
                     $AftGroups | Where-Object {$_.name -eq $Trailer} |
                                 Select-Object -ExpandProperty MapPath
                 }

      if ($AftMaps) {
         [array]$ForeMaps = $MatchingGroups | 
                           Where-Object {$AftMaps -notcontains $_.MapPath} |
                           Select-Object -ExpandProperty MapPath -Unique
         # See here: https://github.com/PowerShell/PowerShell/issues/6131
         $SharesToMap = $ForeMaps + $AftMaps
      } else {
         $SharesToMap = $MatchingGroups | Select-Object -ExpandProperty MapPath -Unique
      }
   } else {
      $SharesToMap = $MatchingGroups | Select-Object -ExpandProperty MapPath -Unique
   }

   $msg = "Unique '$($MapType.Name)' shares user can map:  $($SharesToMap.count)"
   Write-Log $LogFilePath $msg -ToScreen


   if (-not $SharesToMap) {
      Write-Host 'None to map on this computer.'
      # Move to the next MapType
      Continue
   } else {
      # Remove shares the users prefers to not map on this computer.
      if ($NoList) {
         $SharesToMap = @($SharesToMap | Where-Object {$NoList -notcontains $_})
      }
   }

   if (-not $SharesToMap) {
      Write-Host 'None to map on this computer.'
      # Move to the next MapType
      Continue
   }
   
   $i = 0
   Foreach ($Share in $SharesToMap) {
      if ($i -lt $MapType.Letters.count) {
         $drive = $MapType.Letters[$i] + ':'
         # First, (try to) remove any drives that have been already mapped to this drive letter.
         $QueryString = "SELECT * FROM Win32_NetworkConnection WHERE LocalName = '" + $drive + "'"
         if ($Existing = Get-WmiObject -Query $QueryString ) {
            try {
               Write-Log $LogFilePath "Disconnecting current mapping to '$drive':  $($Existing.Name)." -ToScreen
               (New-Object -ComObject WScript.Network).RemoveNetworkDrive($drive,$true,$true)
            } catch {
               Write-Log $LogFilePath "ERROR:  $_" -ToScreen
            }
         }
         # Now, (try to) map the share.
         try {
            (New-Object -ComObject WScript.Network).MapNetworkDrive($drive, $Share)  # add $true for persistence
            Write-Host "Share '$Share' mapped to the $drive drive."
            Write-Log $LogFilePath "Successfully mapped '$Share' to $drive."
            $i++
         } catch {
            Write-Warning "Mapping '$($Share)' to $drive failed."
            Write-Log $LogFilePath "   ERROR mapping '$Share' to '$drive' `n`t$_"
         }
      } else {
         Write-Warning "All allocated drives in use.  Cannot map '$Share'"
         $msg = "All $($MapType.Letters.count) drive letters already mapped to $($MapType.Name) shares."
         Write-Log $LogFilePath $msg
      }
   }
}
$adoConnection.Close()

Write-Host "`n"

# The $NoMap shares are not limited to those mapped by this script.
Foreach ($UNCpath in $NoList) {
   # Look to see if the shares are mapped, and then un-map them
   $QueryString = "Select * From Win32_LogicalDisk Where DriveType = 4 AND ProviderName = '" + $UNCpath.replace('\','\\') + "'"
   if ($Existing = Get-WmiObject -Query $QueryString ) {
      try {
         (New-Object -ComObject WScript.Network).RemoveNetworkDrive($Existing.DeviceID,$true,$true)
         $msg = "Disconnecting '$($Existing.ProviderName)' on $($Existing.DeviceID) as requested in:`n`t '$NoMapFilePath'."
         Write-Log $LogFilePath $msg -ToScreen
      } catch {
         Write-Log $LogFilePath "ERROR disconnecting '$($Existing.ProviderName)' on $($Existing.DeviceID): `n`t $_" -ToScreen
      }
   }
}

Write-Log $LogFilePath "End of script`n--"

if ($VerbosePreference -ne 'SilentlyContinue') {
   Write-Host "`n"
  Write-Verbose 'Press any key to exit...'
  $Null = $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}