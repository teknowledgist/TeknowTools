function Get-DellWarrantyInfo {
<#
.SYNOPSIS
   A script to return Dell purchase date and warranty end.
.DESCRIPTION
   When run, this will query Dell's v5 Warranty API for a list of 
   provided Service Tags and return the corresponding "Ship Date"
   and the "End Date" for any "Onsite" or "Diagnosis" entitlements.
.PARAMETER Tags
   A list of Service Tags or a string of comma-separated Service
   Tags or a combination of these. (The alias, "ServiceTags" as 
   the parameter name will also work.)
.PARAMETER ApiKey
   The Client ID assigned by Dell to the TechDirect admin of
   the organization for accessing the Dell Warranty API.  
   ** This cannot be shared outside of the organization. **
   (The alias, "ClientID" as the parameter name will also work.)
.PARAMETER KeySecret
   The Client Secret assigned by Dell to the TechDirect admin of
   the organization for accessing the Dell Warranty API.  
   ** This cannot be shared outside of the organization. **
   (The alias, "ClientSecret" as the parameter name will also work.)
.PARAMETER Xml
   Collect data in XML format from the API rather in the default,
   JSON format.  This doesn't change the output of this function,
   but it is included for testing and modification of the script.
.NOTES
   Script version: 1.0

   Copyright 2020 Erich Hammer

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
      [Parameter(Mandatory = $true)] [alias('ServiceTags')][string[]]$Tags,
      [Parameter(Mandatory = $true)] [alias('ClientID')][string]$ApiKey,
      [Parameter(Mandatory = $true)] [alias('ClientSecret')][string]$KeySecret,
      [Parameter(Mandatory = $false)] [switch]$xml
   ) 

   $Bytes = [Text.Encoding]::ASCII.GetBytes("$ApiKey`:$KeySecret")
   $EncodedOAuth = [Convert]::ToBase64String($Bytes)
   [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
   $TokenArgs = @{
      Method  = 'Post'
      URI     = 'https://apigtwb2c.us.dell.com/auth/oauth/v2/token'
      Body    = 'grant_type=client_credentials'
      Headers = @{ 
         authorization = "Basic $EncodedOAuth"
      }
   }

   Try { 
      # Access tokens have a one-hour lifetime, so we may as well note the time
      $AccessToken = @{
         Token = (Invoke-RESTMethod @TokenArgs).access_token
         Stamp = Get-Date
      }
   }
   Catch {
      Write-Error $Error[0]
      BREAK
   }

   # I don't know why you would want to use xml over json, but both options work
   if ($xml) { $Type = 'xml' } else {$Type = 'json'}
   
   # Tags could be pre-joined with commas or independent (or both).
   #   This will separate all tags.
   [array]$TagsList = ($Tags -join ',').split(',')
   
   Do {
      # Query parameters are limited to 100 service tags at a time
      $TagSet = $TagsList[0..99]
      $QueryArgs = @{
         Method      = 'GET'
         URI         = 'https://apigtwb2c.us.dell.com/PROD/sbil/eapi/v5/asset-entitlements'
         ContentType = "application/$Type"
         Body        = @{
            servicetags = $TagSet -join ','
            #            Method      = 'GET'
         }
         Headers     = @{ 
            Accept        = "application/$Type"
            Authorization = "Bearer $($AccessToken.Token)"
         }
      }
      $Response = Invoke-RestMethod @QueryArgs

      if ($Type -eq 'xml') {
         foreach ($ST in $TagSet) {
            $ShipString = ($Response | select-xml -xpath "//item[contains(serviceTag,'$ST')]/shipDate").tostring()
            $ShipDate = Get-Date $ShipString -Format 'yyyy.MM.dd'
            
            $node = ($response | select-xml -xpath "//item[contains(serviceTag,'$ST')]").node
            $EndString = $node.entitlements.entitlement |
                              Where-Object {$_.serviceleveldescription -match 'onsite|diagnosis'} |
                              Select-Object -ExpandProperty endDate
            $WarrantyEnd = Get-Date $EndString -Format 'yyyy.MM.dd'

            New-Object -TypeName PSObject -Property @{
               ServiceTag      = $ST
               ShipDate        = $ShipDate
               WarrantyEndDate = $WarrantyEnd
            }
         }
      } else {
         foreach ($Record in $Response) {
            $ServiceTag = $Record.servicetag

            $ShipDate = get-date $Record.ShipDate -Format 'yyyy.MM.dd'
            $EndString = $Record.entitlements | 
                              Where-Object {$_.serviceleveldescription -match 'onsite|diagnosis'} | 
                              Select-Object -ExpandProperty enddate
            $WarrantyEnd = Get-Date $EndString -Format 'yyyy.MM.dd'

            New-Object -TypeName PSObject -Property @{
               ServiceTag      = $ServiceTag
               ShipDate        = $ShipDate
               WarrantyEndDate = $WarrantyEnd
            }
         }
      }
      
      # Purge the first 100 tags in preparation for the next 100
      $TagsList = $TagsList[100..$TagsList.Length]
   } while ($TagsList)

}

