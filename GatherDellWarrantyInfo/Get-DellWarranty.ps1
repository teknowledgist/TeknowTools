<#
.SYNOPSIS
   A startup script for gathering and storing Dell Warranty info.
.DESCRIPTION
   When run as a startup script, this will query Dell's Warranty
   API (if the machine is a Dell) and store the XML returned as
   well as extract the "Ship Date" and store it in a registry key.
.PARAMETER APIkey
   The security key assigned by Dell to the TechDirect admin of
   the organization for accessing the Dell Warranty API.  
   ** This cannot be shared outside of the organization. **
   Default Value: ''
.PARAMETER OrgName
   The short name of the organization.  This is used in the 
   file and registry paths for storing the warranty information.
   The XML information returned by Dell is stored in:
      "$env:ProgramData\$OrgName\Information\DellWarranty.xml"
   The "ShipDate" and "WarrantyEndDate" registry values are in:
      "HKLM:\SYSTEM\$OrgName"
   Default Value: ITServices
.NOTES
   Script version: 1.0

   Copyright 2018 Erich Hammer

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
   [string]$APIKey     = '',
   [string]$Orgname    = 'ITServices'
)

#=====================
#region Functions
#=====================
function Get-XMLDates {
   param([Parameter(Mandatory=$true)][xml]$XML)

   $ns = New-Object -TypeName System.Xml.XmlNamespaceManager -ArgumentList ($XML.NameTable)
   $ns.AddNamespace('TempNS', $XML.DocumentElement.NamespaceURI)
   $Result = 0
   Try {
      $Date = $xml.SelectSingleNode('//TempNS:ShipDate',$ns).'#text' 
      if (-not ([DateTime]::TryParse($Date,[ref]$Result))) {
         Throw
      }
      $ShipDate = $Result.ToString('yyyy.MM.dd')
   } Catch {
      $ShipDate = 'XML_ERROR'
   }

   Try {
      $End = $xml.selectNodes('//TempNS:AssetEntitlement',$ns) |
                  Where-Object {$_.serviceleveldescription -match '(onsite|diagnosis)'} |
                  Sort-Object EndDate |
                  Select-Object -ExpandProperty EndDate -last 1
      if (-not ([DateTime]::TryParse($End,[ref]$Result))) {
         Throw
      }
      $WarrantyEnd = $Result.ToString('yyyy.MM.dd')
   } Catch {
         $WarrantyEnd = 'XML_ERROR'
   }

   # Return an object with these properties
   New-Object -TypeName PSObject -Property @{
      ShipDate        = $ShipDate
      WarrantyEndDate = $WarrantyEnd
   }
} #END Get-XMLDates function

#=====================
#endregion Functions
#=====================

$FilePath = "$env:ProgramData\$OrgName\Information\DellWarranty.xml"
$RegPath = "HKLM:\SYSTEM\$OrgName"

if (Test-path $FilePath) {
   if (Test-Path $RegPath) {
      $ShipDate = (Get-ItemProperty -Path $RegPath -Name 'ShipDate' -ErrorAction SilentlyContinue).ShipDate
      if ($ShipDate -match '^\d\d\d\d\.\d\d\.\d\d') {
         Return    # XML file exists and registry value is in the expected format
      }
   } 
   
   # A file exists, but the registry value is missing/bad
   $XML = [xml](Get-Content $FilePath)
   $Dates = Get-XMLDates $XML
   if (($Dates.ShipDate -notmatch 'ERROR') -and ($Dates.WarrantyEndDate -notmatch 'ERROR')) {
      if (-not (Test-Path $RegPath)) {
         $null = New-Item -Path 'HKLM:\SYSTEM' -Name $Orgname -Force
      }
      $null = New-ItemProperty -Path $RegPath -Name 'ShipDate' -Value $Dates.ShipDate -force
      $null = New-ItemProperty -Path $RegPath -Name 'WarrantyEndDate' -Value $Dates.WarrantyEndDate -force
   
      # Verified XML info and attempted to write to Registry, so we're done.
      Return
   }
}

# XML file is either missing or invalid

$BIOS = Get-WmiObject -Class win32_bios

If ($BIOS.manufacturer -imatch 'dell') {
   # Don't bother looking if not online
   $Online = test-connection -ComputerName www.dell.com -Count 1 -Quiet
   If (-not $Online)  { 
      $ShipDate = 'OFFLINE' 
      $WarrantyEndDate = 'OFFLINE'
   } else {
      $ServiceTag = $BIOS.serialnumber
    
      # Sandbox (Doesn't seem to work any longer.):
      # $DellURL = "https://sandbox.api.dell.com/support/assetinfo/v4/getassetwarranty/$ServiceTag"
      # Production:
      $DellURL = "https://api.dell.com/support/assetinfo/v4/getassetwarranty/$ServiceTag"
    
      Try {
         $web = New-Object Net.WebClient
         
         $web.Headers.Add('Accept','Application/xml')   # There is a 'Application/json' option too
         $web.Headers.Add('APIKey',$APIKey)
        
         $XML = [xml]$web.DownloadString($DellURL)
      } catch { 
         $ShipDate = 'API_ERROR' 
         $WarrantyEndDate = 'API_ERROR'
      }
      if ($XML) {
         $Dates = Get-XMLDates $XML
         $ShipDate = $Dates.ShipDate 
         $WarrantyEndDate = $Dates.WarrantyEndDate

         # Create the file location if needed
         if (-not (Test-Path (Split-Path $FilePath))) {
            $Folders = (Split-Path $FilePath).split('\')
            for ($i=0; $i -lt $Folders.count; $i++) {
               $SubPath = $Folders[0..$i] -join '\'
               if (-not (Test-Path $SubPath)) {
                  $null = New-Item -Path $SubPath.TrimEnd($Folders[$i]) -Name $Folders[$i] -ItemType Directory
               }
            }
         }
         $XML.Save($FilePath)
      } else {
         # Can't imagine a situation where this occurs, but just to be safe
         $ShipDate = 'WEB_ERROR' 
         $WarrantyEndDate = 'WEB_ERROR'
      }
   }
} else {
   $ShipDate = (Get-ItemProperty -Path $RegPath -Name 'ShipDate' -ErrorAction SilentlyContinue).ShipDate
   if (-not $ShipDate) {
      # Use a much less accurate BIOS date for non-Dells
      if ($BIOS.releasedate) {
         $BIOSdate = ([Management.ManagementDateTimeConverter]::ToDateTime($BIOS.releasedate))
         $ShipDate = get-date $BIOSdate -format 'yyyy.MM.dd'
      }

      # May as well identify virtual systems
      $ComputerSystem = Get-WmiObject -Class Win32_ComputerSystem
      if ($ComputerSystem.model -match 'virtual') {
         $WarrantyEndDate = $ComputerSystem.model.split()[0]
      } else {
         $WarrantyEndDate = 'UNKNOWN'
      }
   }
}

# An empty $WarrantyEndDate should mean it is a Non-Dell and 
#    has already written the BIOS date to the registry.  
#    There is no reason to write it every time.
if ($WarrantyEndDate) {
   if (-not (Test-Path $RegPath)) {
      $null = New-Item -Path 'HKLM:\SYSTEM' -Name $Orgname -Force
   }
   $null = New-ItemProperty -Path $RegPath -Name 'ShipDate' -Value $ShipDate -Force
   $null = New-ItemProperty -Path $RegPath -Name 'WarrantyEndDate' -Value $WarrantyEndDate -Force
}

