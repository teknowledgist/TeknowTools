<#
.SYNOPSIS
   Sets a local password based on a human-computable rubric.
.Description
   This script is intended to be called in a startup script for setting the 
   password of a local administrator account.  (For best practice, this script 
   should be called from a Group Policy Object that is filtered to only allow 
   domain computers and OU/Domain admins read access.)  The goal is to set a 
   password that can be built from a specified set of attributes generally 
   known or accessible to technicians who support the computer.  The idea is 
   to adequately secure local administrator accounts for domain workstations 
   while allowing technical support to determine the password while having 
   only immediate access to the local computer.  To determine the password, 
   an attacker with Active Directory account credentials and knowledge of 
   Group Policy, Active Directory structure and Scripting would still need to:
      1) Be an admin of a domain machine to read/parse the script (as system).
      2) Acquire physical information about a computer by either:
         a) Making a remote WMI call as a domain admin.
         b) Having physical access to the computer.
   Thus, discovery of local admin credentials and the ability to broadly 
   attack workstations accross the domain is severely restricted while the
   ability of technicians to work on machines as local admins is relatively
   unhampered.
.Parameter UserName
   The LogonName of the user whose password will be set.  If not given or is 
   "password", NO password will be set, but will be output to the console only.
   Default value: password
.Parameter Pattern
   Defines the set and order of attributes to be compiled into a password. The
   valid set of (case-insensitive) items to choose from are as follows:
      B: representing the computer Brand or Manufacturer.
      C: representing the next special character in the -character parameter.
      D: representing the next (comma-separated) item in the -datestring parameter.
      M: representing the computer Model.
      O: representing the next -OUAttribute of the computer's Organizational Unit.
      S: representing the next substring of the serial number.
   Default value: SCD
.Parameter BrandNameLimit
   Specifies the number of (lower case) characters from the front of the string
   returned by the computer as the Brand (or "Manufacturer") to be included in the 
   password when "B" is included in the -Pattern parameter.  Special characters or 
   spaces are stripped before counting.  For example, the first six characters used
   from "Dell, Inc." would be "dellin".
   Default Value: 4
.Parameter Character
   Lists the special characters to be included in the password when "C" is included
   in the -Pattern parameter.  Strings of characters will be looped through.  Thus,
   specifying the string "!#" will use "!" for every odd instance of "C" and "#" 
   for every even instance. Surround strings of special characters with single
   quotes (') to be safe and test as some combinations of special characters will
   cause unexpected results.
   Default Value: !
.Parameter DateString
   Specifies a custom date string or strings to be included in the password when "D"
   is included in the -Pattern parameter.  Only combinations of day, month and year
   (d,M,y) as defined in http://technet.microsoft.com/en-us/library/ee692801.aspx
   are accepted, but if multiple strings are provided (separated by comma), they
   will be used in sequence and looped for each instance of "D".
   Default Value: MMMyy
.Parameter ModelStub
   Specifies which end of the Model string will be used for each "M" included in
   the -Pattern parameter.  Only values of "First" or "Last" or combinations of
   those separated by commas are accepted.  The stubs will be looped for multiple 
   instances of "M".
   Default Value: Last
.Parameter ModelLimit
   Specifies the number of (lower case) characters from the computer Model (as
   recognized by Windows) to be included in the password when "B" is included in 
   the -Pattern parameter.  Special characters or spaces are stripped before 
   counting.  Values are looped for multiple instances of "M".
   Default Value: 4
.Parameter OUAttribute
   Specifies the AD attribute of the Organizational Unit of which the computer 
   is a direct member which be used when "O" is included in the -Pattern 
   parameter.  If multiple strings are provided (separated by comma), they
   will be used in sequence and looped for each instance of "O".  If the 
   attribute cannot be found or is empty then the default password will be 
   set.  If the computer is not in a domain (and thus has no OU) or the domain
   cannot be reached (e.g. off-network or access denied), then the value 
   processed will be the domain or workgroup of the computer.  
   Default Value: "Description"
.Parameter OULimit
   Specifies the number of characters from the front of the string returned
   as the OUAttribute to be included in the password when "O" is included in 
   the -Pattern parameter.  Special characters or spaces are stripped before 
   counting, but *_upper/lower case is preserved_*.  Values are looped for 
   multiple instances of "O".
   Default Value: 4
.Parameter SerialStub
   Specifies which end of the serial number string will be used for each "S" 
   included in the -Pattern parameter.  Only values of "First" or "Last" or 
   combinations of those separated by commas are accepted. The stubs will be 
   looped for multiple instances of "S".
   Default Value: First
.Parameter SerialLimit
   Specifies the number of (lower case) characters from the computer serial
   number (as recognized by Windows) to be included in the password when "S" 
   is included in the -Pattern parameter.  Values are looped for multiple 
   instances of "S".
   Default Value: 4
.Parameter Default
   Specifies the default password to be set if any attribute of the -Pattern 
   parameter contributes zero length to the constructed password.  If this
   value is set to "" or $null, the existing password will remain unchanged
   if an empty pattern item is encountered.  If this value is set to "random"
   the password will be set to a random, 12 character string if an empty
   pattern item is encountered.
   Default Value: ""
.Parameter Random
   Specifies the number of random characters to which the password should 
   be set.  This can only be used in conjunction with the username parameter.
   Default Value: 12
.Example
   .\Set-LocalPassword.ps1
   This will return an example password set to the default parameters (first 4 
   characters of the serial number followed by an exclaimation point [!] 
   followed by the three letter abbreviation of the current month and the 
   two digit year).  
   For an Dell, Optiplex 990 computer with serial number "62z0dv1" in an OU
   "PsychLab" (with OU description, "-->HIPPA!") running the script in 
   March, 2016, the string returned would be "62z0!mar16".
.Example 
   .\Set-LocalPassword.ps1 Administrator -Pattern ocs -OUAttribute invalid
   This will not change the local "Administrator" account password for 
   the example computer because there is no "invalid" OU attribute.
.Example
   .\Set-LocalPassword.ps1 Boss -Pattern dbsd -DateString yy,MMMM -SerialStub Last
   This will set the local "Boss" account password for the example computer
   to "16dell0dv1march".
.Example
   .\Set-LocalPassword.ps1 LocAdmin cocsc -Character '@&' -SerialLimit 10
   This will set the local "LocAdmin" account password for the example
   computer to "@HIPP&62z0dv1@".
.Example
   .\Set-LocalPassword.ps1 fixIT -Pattern oss -OUAttribute Name -SerialStub First,Last
   This will set the local "fixIT" account password for the example
   computer to "psyc62z00zdv1".  Note the double "0" from the serial number.
.Example
   .\Set-LocalPassword.ps1 Admin -default 'random'
   If run offline or with another local administrator account, this will
   set the local "Admin" account password to a random string of 12 characters.
   If run with access to a domain as a domain member, the string returned 
   would be "62z0!mar16"
.Example 
   .\Set-LocalPassword.ps1 Administrator -Random 25
   This will set the local "Administrator" account password for the Example
   computer to a 25-character, random string.
.Notes
   Copyright 2015-2017 Teknowledgist

   This script/information is free: you can redistribute it and/or
   modify it under the terms of the GNU General Public License as
   published by the Free Software Foundation, either version 2 of the 
   License, or (at your option) any later version.

   This script is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   The GNU General Public License can be found at <http://www.gnu.org/licenses/>.
#>
[CmdletBinding(DefaultParametersetName='Rubric')]
Param(
   #The LogonName of the user whose password will be set.
   [Parameter(Position=0)]   
   [ValidateNotNullOrEmpty()]
   [string]$UserName = 'password',

   # Snippet selection & order.
   [Parameter(ParameterSetName='Rubric',Position=1)]    
   [ValidateScript({
      If ($_ -imatch "^[bcdmos]+$") { $true } 
      else {Throw "`n'$_' contains letters other than those from the set [bcdmos]."}
   })]
   [string]$Pattern = 'OCS',

   # Number of characters of the Brand ("Manufacturer") string
   [Parameter(ParameterSetName='Rubric')]
   [int]$BrandNameLimit = 4,

   # symbol characters to include (reference:  http://www.asciitable.com/) 
   [Parameter(ParameterSetName='Rubric')]
   [ValidateScript({
      If ($_ -imatch "^[\x20-\x2F\x3A-\x40\x5B-\x60\x7B-\x7E]+$") { $true } 
      else {Throw "`n'$_' is not a string of only special characters."}
   })]
   [string]$Character = '!',
   
   # Custom date string(s).
   [Parameter(ParameterSetName='Rubric')]
   [ValidateScript({
      If ($_ -cmatch "^((((d\.)|d{2,4})|(y{1,4})|(M{1,4})),?)+$") { $true } 
      else {Throw "`n'$_' is not a date string made up of only d, M, or y"}
   })]
   [string[]]$DateString = @('MMMyy'),

   # End(s) of Model string to use.
   [Parameter(ParameterSetName='Rubric')]
   [ValidateScript({
      If ($_ -imatch "^(First|Last)(,(First|Last))*$") { $true } 
      else {Throw "`n$_ must be 'first' or 'last' or comma-separated combinations."}
   })]
   [string[]]$ModelStub = @('Last'),

   # Number of characters of the Model string to be used.
   [Parameter(ParameterSetName='Rubric')]
   [int[]]$ModelLimit = '4',

   # AD attributes to be used of the OU of the computer.
   [Parameter(ParameterSetName='Rubric')]
   [string[]]$OUAttribute = @('Description'),
   
   # Max lengths of substring used from OU attribute
   [Parameter(ParameterSetName='Rubric')]
   [int[]]$OULimit = '4',
   
   # First or Last characters of serial number
   [Parameter(ParameterSetName='Rubric')]
   [ValidateScript({
      If ($_ -imatch "^(First|Last)(,(First|Last))*$") { $true } 
      else {Throw "`n$_ must be 'first' or 'last' or comma-separated combinations."}
   })]
   [string[]]$SerialStub = @('Last'),

   # Number of characters of serial number separated by a comma
   [Parameter(ParameterSetName='Rubric')]
   [int[]]$SerialLimit = 4,

   # Default password to use if there is a failure
   [Parameter(ParameterSetName='Rubric')]
   [string]$Default = '',
   
[Parameter(ParameterSetName='Random',Position=1)]
  [Int]$Random = 12
)

########################################################################
# Start of script processing.
########################################################################
set-strictmode -version 2.0
$PW = ''

:top switch ($PsCmdlet.ParameterSetName) {
   'Rubric' {
# Preload (one time) all the extracted values to be used and initialize counts
      Switch -regex ($Pattern) {
         'b' { $BrandName = (Get-WmiObject Win32_ComputerSystem).Manufacturer.trim().tolower() -replace '\W',''}
         'c' { $Ccount = 0 }
         'd' { $Dcount = 0 }
         'm' { $Mcount = 0
               $Model = (Get-WmiObject Win32_ComputerSystem).Model.trim().tolower() -replace '\W','' 
         }
         's' { $Scount = 0
               $SN = (Get-WmiObject Win32_Bios).SerialNumber.tolower() -replace '\W',''
         }
         'o' { $OUcount = 0
               $Domain = (Get-WmiObject -Class win32_computersystem).domain
               $PossibleMsg = "The computer is not in a domain or the domain cannot be contacted.`n" + 
                                 "Defaulting to an OU attribute value of:   " + $(Get-WmiObject -Class win32_computersystem).domain

               # Is the computer's "domain" reachable?
               try {
                  test-connection $Domain -count 1 -erroraction stop | out-null 
               } catch {
                  write-warning $PossibleMsg
                  $OUObject = $Domain
                  break
               }
               $DomainDN = "DC=" + ($Domain.split('.') -join ",DC=")
               # In which OU is the computer?
                  $adsiSearch = [adsisearcher]('(&(objectCategory=computer)(cn=' + $env:COMPUTERNAME + '))')
                  $adsiSearch.SearchRoot = [adsi]"LDAP://$DomainDN"
               try {
                  $FullPath = $adsiSearch.FindAll()[0].path.remove(0,7)
               } catch {
                  write-warning $PossibleMsg
                  $OUObject = $Domain
                  break
               }
               $OUname = $FullPath.split(',')[1].remove(0,3)
               $OUPath = $FullPath -replace '.*?,.*?,(.*)','$1'
               # Collect OU object (with attributes)
               $adsiSearch.Filter = '(&(objectCategory=OrganizationalUnit)(name=' + $OUname + '))'
               $adsiSearch.SearchRoot = [adsi]('LDAP://' + $OUpath)
               $OUObject = $adsiSearch.FindAll()[0]
            }
      } # end of: Switch -regex ($Pattern)

      Foreach ($x in $Pattern.ToCharArray()) {
         Switch ($x) {
            'B' {
               if (-not $BrandName) {
                  $PW = $Default
                  Break top
               }
               if ($BrandNameLimit -gt $BrandName.length) {
                  $PW = [string]($PW + $BrandName)
               } else {
                  $PW = [string]($PW + $BrandName.substring(0,$BrandNameLimit))
               }
               Break
            }
            'C' { 
               $PW = [string]($PW + $Character.ToCharArray()[$Ccount % $Character.length])
               $Ccount++
               Break
            }
            'D' {
               $PW = [string]($PW + (Get-Date -Format ($DateString[$Dcount % $DateString.length])).tolower())
               $Dcount++
               Break
            }
            'M' {
               if (-not $Model) {
                  $PW = $Default
                  Break top
               }
               if ($Model.length -lt $ModelLimit[$Mcount % $ModelLimit.length]) {
                   $PW = [string]($PW + $Model)
               } else {
                  if ($ModelStub[$Mcount % $ModelStub.length] -eq 'First') {
                     $PW = [string]($PW + $Model.substring(0,$ModelLimit[$Mcount % $ModelLimit.length]))
                  } else {
                     $PW = [string]($PW + $Model.substring($Model.length - $ModelLimit[$Mcount % $ModelLimit.length]))
                  }
               }
               $Mcount++
               Break
            }
            'O' {
               if ($OUObject -ne $Domain) {
                  $OUstring = $OUObject.properties.item($OUAttribute[$OUcount % $OUAttribute.length])[0] -replace '[^1-9a-zA-Z]',''
                  if (-not $OUstring) {
                     $PW = $Default
                     Break top
                  }
               } else {
                  $OUstring = $Domain
               }
               if ($OUstring.length -gt $OULimit[$OUcount % $OULimit.length]) {
                  $PW = [string]($PW + $OUstring.substring(0,$OULimit[$OUcount % $OULimit.length]))
               } else { $PW = [string]($PW + $OUstring) }
               $OUcount++
               Break
            }
            'S' {
               if (-not $SN) {
                  $PW = $Default
                  Break top
               }
               if ($SN.length -lt $SerialLimit[$Scount % $SerialLimit.length]) {
                   $PW = [string]($PW + $SN)
               } else {
                  if ($SerialStub[$Scount % $SerialStub.length] -eq 'First') {
                     $PW = [string]($PW + $SN.substring(0,$SerialLimit[$Scount % $SerialLimit.length]))
                  } else {
                     $PW = [string]($PW + $SN.substring($SN.length - $SerialLimit[$Scount % $SerialLimit.length]))
                  }
               }
               $Scount++
               Break
            }
         } # end of: Switch ($x)
      } # end of: foreach
   } # end of : "Rubric" switch item
   'Random' {
      [Reflection.Assembly]::LoadWithPartialName('System.Web') | Out-Null
      $PW = [System.Web.Security.Membership]::GeneratePassword($Random,0)
   }
} # end of: switch ($PsCmdlet.ParameterSetName)

# If something failed in defining a password and the default calls
# for a random password, create a random string.
if ($PW -eq 'Random') {
      [Reflection.Assembly]::LoadWithPartialName('System.Web') | Out-Null
      $PW = [System.Web.Security.Membership]::GeneratePassword(12,0)
}

If ($Username -eq 'password') {
   # By default, output password instead of changing it
   $PW
} else {
   try {
      $ADSI = [ADSI] "WinNT://$env:ComputerName/$Username,User"
      $ADSI.setpassword($PW)
   }
   catch { Write-Error $Error[0] }
}