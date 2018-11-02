<#
.SYNOPSIS
   Collects login/logout times from the security log for tracking usage.
.Description
   This script will extract interactive logins and user initiated logouts as 
   recorded in the system security log.  It is intended to be run as a startup
   script (with system privileges), but it will run under (elevated) admin
   privileges.  The purpose is to build a set of simple, CSV files to be used 
   in billing for the time people have used the machine. 
   
   Each time the script runs, there is a potential for 5 files to be created 
   in a local archive plus 5 duplicate files in a shared location.  One file 
   will hold all logon/off records found in the security log at the time the
   script is run.  One file will simply have the time stamp of the last time 
   the script was run.  Three files will hold the logon/off records of the 
   current month, last month and the previous-to-last month.  The latter two 
   will overwrite previous logs only for those months it is certain the current 
   set contains the entire month in question.
   
   This script does not attempt to determine the amount of time (or even if)
   a user has locked or slept the computer.  It simply calculates the difference
   between user logon and logoff (or other user-session-ending) events.  As 
   such, this script assumes that only one user account is logged in at a time. 
   Therefore, Fast-User-Switching MUST BE TURNED OFF for the best accuracy.
.Parameter LogSharePath
   The folder path to the set of logs that are to be created.  The intent is to
   store the logs in a shared location that a non-technical user can access.
   This can be a network share if it is made accessable to the system account.
   (For security, it should not be write-accessible by users.) A set of logs 
   will be kept locally in %windir%\Logs\Usage regardless of whether this is 
   set or not.
   Default value: <none>
.Parameter LDAPpath
   Specifies the distinguished name of the path used for the Global Catalog 
   that contains user information.
   Default Value: The DN of the forest of the entity running the script.
.Example
   .\Log-usage.ps1 \\FileServer\UsageLogs\
   This will create a set of logs both in the local archive and in a folder
   with the name of the computer it is run from on the "UsageLogs" share on 
   the server, \\FileServer.  For example, on a computer named "Station1",
   the folder, "\\FileServer\UsageLogs\Station1" would be created and contain
   at least a "LatestRunTime.txt" and a "Current.csv" file.  There may be
   three additional csv files named after the current month, previous month,
   and the month before that.
.Notes
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
Param(
   #The folder path to the location of the usage log
   [Parameter(Position=0)]   
   [string]$LogSharePath = $null,
   
   [Parameter(Position=1)]   
   [string]$LDAPpath = ([DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().domains | 
                           Where-Object {$_.name -eq $_.forest}).getdirectoryentry().distinguishedname[0]
)

#**************************************************
#region     Start of script processing.
#**************************************************
$XP = ([Environment]::OSVersion.Version.major -le 5)

if ($LogSharePath) {
   $LogSharePath = Join-path $LogSharePath $env:COMPUTERNAME
   if (!(Test-Path -path $LogSharePath)) {
      $null = New-Item -Path $LogSharePath -ItemType directory
   }
   if (!(Test-Path -path $LogSharePath)) {
      $LogSharePath = $null
   }
}
$LocalLogPath = Join-Path $env:windir 'Logs\Usage'
if (!(Test-Path -path $LocalLogPath)) {
   $null = New-Item -Path $LocalLogPath -ItemType directory
}

Add-Type -TypeDefinition @'
   public class UseRecord {
      public string NetID;
      public string FullName;
      public string eMail;
      public System.DateTime LoginTime;
      public System.DateTime LogoutTime;
      public int Duration;
      public string Verification;
      public int Confidence;
   }
'@

# To be used for collecting user info
$Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = "GC://$LDAPpath"
$null = $Searcher.PropertiesToLoad.Add('displayName')
$null = $Searcher.PropertiesToLoad.Add('mail')
$null = $Searcher.PropertiesToLoad.Add('sAMAccountName')

# Logon process and all event types explained here:
#    https://www.ultimatewindowssecurity.com/blog/default.aspx?p=26180f8b-42a6-49a2-949d-ac44494353cb
#    https://www.ultimatewindowssecurity.com/securitylog/book/page.aspx?spid=chapter5
# A blog post discussing the difficulty in definitively tracking usage times:
#    http://blogs.msdn.com/b/ericfitz/archive/2008/08/20/tracking-user-logon-activity-using-logon-events.aspx
if ($XP) {
   $Entries = Get-EventLog -LogName security -InstanceId 528,538,551,513,512 | 
               Where-Object { ($_.InstanceID -eq 528 -and
                           $_.message -match 'Logon Type:\s+2' -and 
                           $_.message -match 'Source Network Address:\s+127.0.0.1') -or
                        ($_.InstanceID -eq 538 -and
                           $_.message -match 'Logon Type:\s+2') -or
                       ($_.InstanceID -eq 551) -or
                       ($_.InstanceID -eq 513) -or
                       ($_.InstanceID -eq 512)
                     }
   $LogIndexes = @{'Logon' = @{'EventID' = 528;'uSID'=4;'NetID'=5;'Domain'=6;'LogonID'=7;'Type'=8;'LogonGUID'=12};
                  'Begin_Logoff' = @{'EventID' = 551;'uSID'=0;'NetID'=1;'Domain'=2;'LogonID'=3};
                  'Logoff' = @{'EventID' = 538;'uSID'=0;'NetID'=1;'Domain'=2;'LogonID'=3;'Type'=4};
                  'Stop_Logging' = @{'EventID' = x};
                  'Shutdown' = @{'EventID' = 513};
                  'Startup' = @{'EventID' = 512}
                  }
} else {
   # Win10 identifies the process differently
   $OSVersion = [version](Get-WmiObject Win32_OperatingSystem | Select-Object -ExpandProperty Version)
   if ($OSVersion.Major -eq 10) { $LogonProcess = 'svchost.exe' } 
   else { $LogonProcess = 'winlogon.exe' }

   # X-path filter for collecting the security events that best define a logon session
   $XMLfilter = @"
   <QueryList>
      <Query Id="0" Path="Security">
         <Select Path="Security">
            *[System[(EventID=4624)] 
               and 
               EventData[(Data[@Name='LogonType'] = '2')]
               and
               EventData[(Data[@Name='ProcessName'] = 'C:\Windows\System32\$LogonProcess')]
               and 
               (EventData[(Data[@Name='TargetDomainName'] = '$env:USERDOMAIN')] 
                  or 
                  EventData[(Data[@Name='TargetDomainName'] = '$env:COMPUTERNAME')])]
            or 
            *[System[(EventID=4634)] 
               and
               EventData[(Data[@Name='LogonType'] = '2')]
               and 
               (EventData[(Data[@Name='TargetDomainName'] = '$env:USERDOMAIN')] 
                  or 
                  EventData[(Data[@Name='TargetDomainName'] = '$env:COMPUTERNAME')])]
            or
            *[System[(EventID=4647)]]
            or
            *[System[(EventID=1100)]]
            or
            *[System[(EventID=4608)]]
            or
            *[System[(EventID=4609)]]
         </Select>
      </Query>
   </QueryList>
"@
   
   $Entries = Get-WinEvent -FilterXml $XMLfilter |Sort-Object TimeCreated
   $LogIndexes = @{'Logon' = @{'EventID' = 4624;'uSID'=4;'NetID'=5;'Domain'=6;'LogonID'=7;'Type'=8;'LogonGUID'=12};
                  'Begin_Logoff' = @{'EventID' = 4647;'uSID'=0;'NetID'=1;'Domain'=2;'LogonID'=3};
                  'Logoff' = @{'EventID' = 4634;'uSID'=0;'NetID'=1;'Domain'=2;'LogonID'=3;'Type'=4};
                  'Stop_Logging' = @{'EventID' = 1100};
                  'Shutdown' = @{'EventID' = 4609};
                  'Startup' = @{'EventID' = 4608}
                  }
}

$Records = @()
$InTimes = @()
:iloop for ($i=0; $i -lt ($Entries.length - 1); $i++) {
   if ($Entries[$i].id -eq $LogIndexes.Logon.EventID) {
      $InID = $Entries[$i].properties[$LogIndexes.Logon.LogonID].value
      :jloop for ($j=$i+1; $j -lt $Entries.length; $j++) {
         # skip non-logout items as they can't be definitively matched
         if (-not (($Entries[$j].id -eq $LogIndexes.Begin_Logoff.EventID) -or
                     ($Entries[$j].id -eq $LogIndexes.Logoff.EventID))) {
            continue jloop
         }
         if ($Entries[$j].id -eq $LogIndexes.Begin_Logoff.EventID) {
            $Off = 'Begin_Logoff'
         } else {$Off = 'Logoff'}
         
         # look for matching, user-initiated logouts
         if ($InID -eq $Entries[$j].properties[$LogIndexes.$Off.LogonID].value) {
            $Session = New-Object -TypeName UseRecord -Property @{
                        NetID = $Entries[$i].properties[$LogIndexes.Logon.NetID].value.tostring();
                        FullName = '';
                        eMail = '';
                        LoginTime = $Entries[$i].TimeCreated;
                        LogoutTime = $Entries[$j].TimeCreated;
                        Duration = ('{0:N2}' -f (New-TimeSpan -Start $Entries[$i].TimeCreated -End $Entries[$j].TimeCreated).TotalMinutes);
                        Verification = '';
                        Confidence = 0;
                     }
            If ($Entries[$i].properties[$LogIndexes.Logon.Domain].value -ne $env:COMPUTERNAME) {
               # For domain user logons
               if ($XP) {
                  $Searcher.Filter = "(SamAccountName=$($Entries[$i].ReplacementStrings[$LogIndexes.Logon.NetID].value.tostring()))"
               } else {
                  $Searcher.Filter = "(objectSid=$($Entries[$i].properties[$LogIndexes.Logon.uSID].value.tostring()))"
               }
               $user = $Searcher.Findone()
               $Session.FullName = $user.properties.displayname[0]
               $Session.eMail = $(if ($user.properties.mail) {$user.properties.mail[0]} else {'none'})
            } Else {
               # For local user logons
               $Session.FullName = (Get-WmiObject win32_useraccount -filter "LocalAccount='$true' and Name='$($Session.NetID)'").FullName
            }
            
            # Slightly different information for different logoffs
            If ($Off -eq 'Begin_Logoff') {
               $Session.Verification = 'User initiated logoff - normal use'
               $Session.Confidence = 10
            } else {
               $Session.Verification = 'Session logoff completed - service initiated logoff?'
               $Session.Confidence = 9
            }
            
            $Records += $Session
            $InTimes += $Entries[$i].TimeCreated
            $Entries[$i].message = $Entries[$j].message = 'used'
            continue iloop
         } # end IF (matched logonID)
      } # end jloop
   } # end IF (logon)
} # end iloop

# Only look through the remaining entries
$RemEntries = $Entries | Where-Object {$_.message -ne 'used'}

for ($i=0; $i -lt ($RemEntries.length - 1); $i++) {
   If ($RemEntries[$i].id -eq $LogIndexes.Logon.EventID) {
      # Usually a domain logon create a pair simultaneous logon events.
      if ($InTimes -contains $RemEntries[$i].TimeCreated) {
         # Need to ignore any logons paired with verified session times 
         continue
      } else {
         # and use only one of a pair without a verified logoff.
         $InTimes += $RemEntries[$i].TimeCreated
      }
      If ($RemEntries[$i].properties[$LogIndexes.Logon.Domain].value -ne $env:COMPUTERNAME) {
         $Searcher.Filter = "(objectSid=$($RemEntries[$i].properties[$LogIndexes.Logon.uSID].value.tostring()))"
         $user = $Searcher.Findone()
         $Session = New-Object -TypeName UseRecord -Property @{
                     NetID = $user.properties.samaccountname;
                     FullName = $user.properties.displayname[0];
                     eMail = $user.properties.mail[0];
                     LoginTime = $RemEntries[$i].TimeCreated;
                     LogoutTime = 0;
                     Duration = 0;
                     Verification = $null;
                     Confidence = 0;
                  }
      } Else {
         $Session = New-Object -TypeName UseRecord -Property @{
                     NetID = $RemEntries[$i].properties[$LogIndexes.Logon.NetID].value.tostring();
                     FullName = (Get-WmiObject win32_useraccount -filter "LocalAccount='$true' and Name='$($Session.NetID)'").FullName;
                     eMail = ''
                     LoginTime = $RemEntries[$i].TimeCreated;
                     LogoutTime = 0;
                     Duration = 0;
                     Verification = $null;
                     Confidence = 0;
                  }
      }

      # The next collected entry is the best option for identifying when the logon session ended.
      $loop = $false
      do {
         Switch ($RemEntries[$i+1].id) {
            $LogIndexes.Begin_Logoff.EventID {
               $Session.LogoutTime = $RemEntries[$i+1].TimeCreated
               $Session.Duration = '{0:N2}' -f (New-TimeSpan -Start $Session.LoginTime -End $Session.LogoutTime).TotalMinutes
               if ($RemEntries[$i].properties[$LogIndexes.Logon.NetID].value -eq 
                           $RemEntries[$i+1].properties[$LogIndexes.Begin_Logoff.NetID].value) {
                  $Session.Verification = 'Matched user initiated logoff - mismatched session ID.'
                  $Session.Confidence = 6
               } else {
                  $Session.Verification = 'Unmatched user logoff - unknown error.'
                  $Session.Confidence = 0
               }
               $loop = $false}
            $LogIndexes.Logoff.EventID {
               $Session.LogoutTime = $RemEntries[$i+1].TimeCreated
               $Session.Duration = '{0:N2}' -f (New-TimeSpan -Start $Session.LoginTime -End $Session.LogoutTime).TotalMinutes
               if ($RemEntries[$i].properties[$LogIndexes.Logon.NetID].value -eq 
                           $RemEntries[$i+1].properties[$LogIndexes.Logoff.NetID].value) {
                  $Session.Verification = 'Matched user logoff completed - mismatched session ID.'
                  $Session.Confidence = 5
               } else {
                  $Session.Verification = 'Unmatched user logoff completed - unknown error.'
                  $Session.Confidence = 0
               }
               $loop = $false}
            $LogIndexes.Stop_Logging.EventID {
               $Session.LogoutTime = $RemEntries[$i+1].TimeCreated
               $Session.Duration = '{0:N2}' -f (New-TimeSpan -Start $Session.LoginTime -End $Session.LogoutTime).TotalMinutes
               $Session.Verification = 'System logging stopped - funny business?'
               $Session.Confidence = 5
               $loop = $false}
            $LogIndexes.Shutdown.EventID {
               $Session.LogoutTime = $RemEntries[$i+1].TimeCreated
               $Session.Duration = '{0:N2}' -f (New-TimeSpan -Start $Session.LoginTime -End $Session.LogoutTime).TotalMinutes
               $Session.Verification = 'System shutdown - instability from application crash?'
               $Session.Confidence = 5
               $loop = $false}
            $LogIndexes.Startup.EventID {
               $Session.LogoutTime = $RemEntries[$i+1].TimeCreated
               $Session.Verification = 'System startup - loss of power/system crash?'
               $Session.Confidence = 2
               $loop = $false}
            $LogIndexes.Logon.EventID {
               # Check for the end of the log
               if ($i+1 -ge ($RemEntries.length - 1)) {
                  $Session = $null
                  $loop = $false
                  break # out of switch
               }
               if ($Session.NetID -eq $RemEntries[$i+1].properties[$LogIndexes.Logon.NetID].value) {
                  # handles repeated same-user logon entries (maybe locked/unlocked?)
                  $i++
                  $loop = $true
               } else {
                  $Session.Verification = 'Other user logon - Unknown error.'
                  $Session.Confidence = 0
                  $loop = $false
               }
               } # End of "Logon" SWITCH item
         } # End of Switch statement
      } while ($loop)
      
      if ($Session) {
         [array]$Records += $Session
      }
   } # End of Session building

} # End of for loop

$Records = $Records | Sort-Object LoginTime

$now = Get-Date
$MonthsBack = New-Object System.Collections.Generic.List[System.Object]
$MonthsBackLog = New-Object System.Collections.Generic.List[System.Object]
$Count = 0
for ($i=0; $i -lt 3; $i++) {
   $MonthsBack.add(((get-date).addmonths(-$i)).tostring('yyyy-MM'))
   $MonthsBackLog.add(($Records | 
      Where-Object {($_.LoginTime.tostring('yyyy-MM') -eq $now.addmonths(-$i).tostring('yyyy-MM'))}
         ))
   $Count += $MonthsBackLog[$i].count
}
$MBMBtest = ($Count -lt $Records.Count)

# Create a log for the previous months if they don't exist.
for ($i=2; $i -ge 0; $i--) {
   if (-not (Test-Path (Join-Path $LocalLogPath ($MonthsBack[$i] + ".csv")))) {
      if ($MonthsBackLog[$i]) {
         $MonthsBackLog[$i] | Export-Csv (Join-Path $LocalLogPath ($MonthsBack[$i] + ".csv")) -NoTypeInformation
         if ($LogSharePath) {
            $MonthsBackLog[$i] | Export-Csv (Join-Path $LogSharePath ($MonthsBack[$i] + ".csv")) -NoTypeInformation
         }
      } elseif ($i -ne 0) {  # Don't save empty record for the current month
         "No logins found for this month." | Out-File (Join-Path $LocalLogPath ($MonthsBack[$i] + ".csv"))
         if ($LogSharePath) {
         "No logins found for this month." | Out-File (Join-Path $LogSharePath ($MonthsBack[$i] + ".csv"))
         }
      }
   } elseif ($MonthsBackLog[$i]) {
      # Need to determine if a full or partial month was previously recorded.  
      if ((-not (Test-Path (Join-Path $LocalLogPath ($MonthsBack[$i-1] + ".csv")))) -or ($i -eq 0)) {
         if ($MBMBtest -or $MonthsBackLog[$i+1]) {
            $MonthsBackLog[$i] | Export-Csv -Path (Join-Path $LocalLogPath ($MonthsBack[$i] + ".csv")) -NoTypeInformation -Force
            if ($LogSharePath) {
               $MonthsBackLog[$i] | Export-Csv -Path (Join-Path $LogSharePath ($MonthsBack[$i] + ".csv")) -NoTypeInformation -Force
            }
         } else {
            # We know the end of the month info, but the beginning is unsure. 
            # The existing file has more/all of the beginning. Merging gives
            # the most information available.
            $selectList = @('NetID','FullName','eMail','verification',
                              @{Name='LoginTime';Expression={[datetime]$_.LoginTime}},
                              @{Name='LogoutTime';Expression={[datetime]$_.LogoutTime}},
                              @{Name='Duration';Expression={[int]$_.Duration}},
                              @{Name='Confidence';Expression={[int]$_.Confidence}}
                           )
            $Start = import-csv (Join-Path $LocalLogPath ($MonthsBack[$i] + ".csv")) | Select-Object $selectList
            $Full = ($Start + $MonthsBackLog[$i]) | Sort-Object LoginTime -Unique
            $Full | Export-Csv -Path (Join-Path $LocalLogPath ($MonthsBack[$i] + ".csv")) -NoTypeInformation -Force
            if ($LogSharePath) {
               $Full | Export-Csv -Path (Join-Path $LogSharePath ($MonthsBack[$i] + ".csv")) -NoTypeInformation -Force
            }
         }
      }
   }
}

# Create a rolling log of all the current log records to cover edge cases
$Records | Export-Csv -Path (Join-Path $LocalLogPath 'Current.csv') -NoTypeInformation -Force
if ($LogSharePath) {
   $Records | Export-Csv -Path (Join-Path $LogSharePath 'Current.csv') -NoTypeInformation -Force
}

# Create a time stamp for when this script last ran as a check
Out-File -FilePath (Join-Path $LocalLogPath 'LatestRunTime.txt') -InputObject $now -Force
if ($LogSharePath) {
   Out-File -FilePath (Join-Path $LogSharePath 'LatestRunTime.txt') -InputObject $now -Force
}


