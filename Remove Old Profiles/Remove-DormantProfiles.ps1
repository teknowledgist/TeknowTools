<#
.SYNOPSIS
   Removes user profiles based on age with exceptions.
.DESCRIPTION
   This script is intended to be run on startup and will remove user 
   profiles based on age.  Unlike using the popular GPO option to remove 
   dormant profiles, this script can exclude removing profiles for local 
   (or other) accounts.  The need is that vendor supplied/supported 
   equipment or other special setups may have local support accounts 
   with documentation or repair applications saved within the profile,
   and the GPO is all-or-nothing based on age only.  
.PARAMETER MaxAge
   The maximum number of days since a profile was last used before it is
   considered dormant and will be deleted.
   Default value: 90
.PARAMETER ExcludeLocal
   Exclude local profiles from being deleted.  "True" means they will
   NOT be deleted even if dormant.
   Default value: True
.PARAMETER ExcludeOther
   Other profiles (by name -- comma-separated) to be excluded from 
   being deleted.  
   Default value: ''
.Example
   .\Remove-DormantProfiles.ps1 -MaxAge 5 -ExcludeOthers 'Tom','Sue'
   This will remove all user profiles that have no activity in the 
   last 5 or more days except for any profiles for local accounts or
   for the "tom" or "sue" profiles.
.Example 
   .\Remove-DormantProfiles.ps1 -MaxAge 180 -ExcludeLocal $False
   This will remove all user profiles that have no activity in the 
   last 180 or more days including any local account profiles.
.NOTES
   Copyright 2023 Erich Hammer

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
Param(
   [Parameter(Position = 1)][int]$MaxAge = 90,
   [switch]$ExcludeLocal = $true,
   [string[]]$ExcludeOthers = @()
)

Function Write-Log { 
   [CmdletBinding()]
   param(
      [ValidateNotNullOrEmpty()]
         [string]$Path = "$env:Temp\DormantProfileRemoval.log",
      [ValidateSet('Information','Warning','Error')]
         [string]$Status = 'Information',
      [ValidateNotNullOrEmpty()]
         [string[]]$Message = '(No log message submitted.)'
   )
   $Content = [ordered]@{
         DateTime = (Get-Date -uformat %Y-%m-%dT%T)
         Severity = $Status
         Message  = $Message -join "`r`n`t`t"
   }
   [PSCustomObject]$Content | Export-Csv -Path $Path -Append -NoTypeInformation
} # End of Write-Log function


if ($ExcludeLocal) {
   # Find all local accounts (admin/vendor accounts to be kept are local)
   $LocalAccountNames = Get-CimInstance -Class Win32_UserAccount -Namespace 'root\cimv2' -Filter "LocalAccount='$True'" | 
                           Select-Object -ExpandProperty caption | ForEach-Object {$_.split('\')[-1]}
} else { $LocalAccountNames = $null}

# Find all normal, user-profiles.  
$Profiles = Get-CimInstance -class Win32_UserProfile | Where-Object {!($_.special)}
$ProfileNames = $Profiles | Select-Object -ExpandProperty localpath | ForEach-Object {$_.split('\')[-1]}

# Only non-local (i.e. Active Directory) user profiles
$TargetProfiles = $ProfileNames | Where-Object {
                                    ($LocalAccountNames -notcontains $_) -and
                                    ($ExcludeOthers -notcontains $_)
                                 }


# Registry keys are what the "profile deletion" GPO uses to identify time of last use.
$RegProfiles = foreach ($p in (Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList")) {
   $Uname = (Get-ItemProperty "HKLM:\$p" -Name profileimagepath).profileimagepath.split('\')[-1]
   
   # Only check profiles that aren't local
   if ($TargetProfiles -notcontains $Uname) {
      continue
   }

   Remove-Variable -Force UTH,UTL -ErrorAction SilentlyContinue
   $LTH = '{0:X8}' -f (Get-ItemProperty -Path $p.PSPath -Name LocalProfileLoadTimeHigh -ErrorAction SilentlyContinue).LocalProfileLoadTimeHigh
   $LTL = '{0:X8}' -f (Get-ItemProperty -Path $p.PSPath -Name LocalProfileLoadTimeLow -ErrorAction SilentlyContinue).LocalProfileLoadTimeLow
   $UTH = '{0:X8}' -f (Get-ItemProperty -Path $p.PSPath -Name LocalProfileUnloadTimeHigh -ErrorAction SilentlyContinue).LocalProfileUnloadTimeHigh
   $UTL = '{0:X8}' -f (Get-ItemProperty -Path $p.PSPath -Name LocalProfileUnloadTimeLow -ErrorAction SilentlyContinue).LocalProfileUnloadTimeLow
   # Load and Unload times equate to logon/logoff times
   if ($UTH -and $UTL) {
      # Using unload (i.e. logoff) time is most accurate
      [pscustomobject][ordered]@{
         User       = $Uname
         SID        = $p.PSChildName  # Not really needed
         LastTime   = [datetime]::FromFileTime("0x$UTH$UTL")
      }
   } elseif ($LTH -and $LTL) {
      # If logoff is missing (e.g. power outage), then collect load (i.e. logon) time.
      [pscustomobject][ordered]@{
         User       = $Uname
         SID        = $p.PSChildName  # Not really needed
         LastTime   = [datetime]::FromFileTime("0x$LTH$LTL")
      }
   } else {
      # If no information on logon/logoff, then leave profile alone
      continue
   }
} 

$Removals = 0
Foreach ($Candidate in $RegProfiles) {
   if ($Candidate.LastTime -lt (Get-Date).AddDays(-$MaxAge)) {
      try {
         $Profiles | Where-Object { $_.LocalPath.split('\')[-1] -eq $Candidate.user } | 
                     Remove-CimInstance -ErrorAction Stop
         $Removals++
         Write-Log -Message "Removed user profile '$($Candidate.user)'."
      } catch {
         Write-Log -Status Error -Message "Error attempting to remove profile '$($Candidate.user)'."
      }
   }
}

if ($Removals -lt 1) {
   Write-Log -Message "No eligible profiles to remove."
}



