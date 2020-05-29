<#
.SYNOPSIS
   Set and log a fairly readable/memorable but complex passphrase on an AD account.
.DESCRIPTION
   The intended use of this script is to run it in a periodic scheduled task 
   targeting a generic account that is either dangerous if cracked or is shared 
   among a large group of people.  The password for these generic accounts does
   need to be logged, but often also transmitted through voice or notes.  Using
   a passphrase is easier to share and/or remember than a random string of characters.
   The passphrase will combine random words from a list with one random letter
   from each word capitalized and a random, non-alphabetic character used to 
   separate the words.  If the dictionary to be used is too short, the phrase
   will be further complicated by the replacement of characters by numerals.
   
   This idea is loosely based on concepts demonstrated in two places:
         https://www.grc.com/haystack.htm
         http://xkcd.com/936/
   Note: It is shocking how the mind finds meaning in randomly paired words,
         so it may be a good idea to remove words from the dictionary of
         potential impoliteness, violence, sexuality, morbidity or bias.
.PARAMETER NetID
   The LogonName of the user whose password will be set.  For testing purposes,
   if the NetID is "password", an example passphrase will be generated and 
   output to the console only but not logged.
   Default value: "password"
.PARAMETER Minimum
   The minimum length of the resulting password.
   Default value: 12
.PARAMETER LogPath
   Full path to the log of the passphrase.  For it to be useful, it probably 
   should be on a network share with appropriately set ACL.
   Default value: <temp directory>\<NetID>.log
.PARAMETER Maximum
   The maximum length of the resulting password.
   Default value: 1.75*Minimum
.PARAMETER TimeStamp
   Only log the passphrase.  Do not include a timestamp.
   Default value: False
.PARAMETER Overwrite
   The previous log file will be overwritten.  Without this switch, the created
   passphrase will be appended to the log.
   Default value: False
.PARAMETER Dictionary
   Path to the local text file containing a list of words to use for creating
   passphrases.  If this list is shorter than 100 words, an internal,
   default list will be used instead.  Avoid dictionaries with words
   containing non-alphabetic characters (e.g. "-").
   Default value: <script directory>\Dictionary.txt
.PARAMETER Domain
   The FQDN for the domain in which the user resides.
   Default value: The domain of the computer
.Example
   .\Set-MemorablePassphrase.ps1 -NetID tempvisitor
   This will set a 12+ character passphrase for the tempvisitor account in the 
   default domain and append it with a timestamp to "tempvisitor.log" in the temp directory.
.Example 
   .\Set-MemorablePassphrase.ps1 -NetID tempvisitor -logpath \\FS1\Logs\VPass.txt -timestamp -overwrite
   The same as above, but logged to a shared filepath containing only the current passphrase.
.Example
   .\Set-MemorablePassphrase.ps1 HDuser 20 \\FS1\Logs\Secret.txt -Dictionary .\LatinWords.txt
   This will set the password for the HDuser account in the default 
   domain to a 20+ character passphrase derived from a custom dictionary 
   file and append the result (with a timestamp) to a text file on FS1.
.NOTES
   Copyright 2020 Erich Hammer

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
   [Parameter(ValueFromPipeLine = $true, 
              ValueFromPipelineByPropertyName = $true,
              Position = 0)]
              [Alias('Name','SamAccountName')]
      [string]$NetID = 'password',
   [Parameter(Position = 1)][ValidateRange(4,100)][int]$Minimum = 12,
   [Parameter(Position = 2)][string]$LogPath = "$env:temp\$NetID.log",
   [ValidateRange(4,100)][int]$Maximum = 1.75*$Minimum,
      [switch]$TimeStamp = $false,
      [switch]$Overwrite = $false,
      [string]$Dictionary = "$(Split-Path -parent $Script:MyInvocation.MyCommand.Path)\Dictionary.txt",
      [string]$Domain = (get-wmiobject win32_computersystem).domain
)

# If no (or too short a) dictionary file, use a default list
$list = @()
if (Test-Path $Dictionary) {
   $list = Get-Content -Path $Dictionary
}
if ($list.count -lt 100) {
   # 100 common words > 4 characters
   $list = ('about,after,again,along,another,around,because,before,below,between,could,different,nothing,every,skip,' +
            'found,great,house,large,little,might,never,number,other,people,place,right,should,small,something,sound,' +
            'still,there,these,thing,think,those,thought,three,through,together,under,water,where,which,while,world,' +
            'would,write,above,across,against,almost,among,animal,answer,became,become,began,behind,being,better,' +
            'black,brought,cannot,certain,change,children,close,country,course,night,during,early,earth,enough,' +
            'example,family,father,front,given,green,ground,group,heard,himself,however,important,inside,known,later,' +
            'learn,letter,light,living,making,means,money,morning,mother').split(',')
   # For a short list, further complicate the passphrase (below)
}

$Words = @()
$GotWords = $false
While (-not $GotWords) {
   $Words += $list | Get-Random

   # Capitalize a random character in the word
   $ToCap = 0..($Words[-1].length -1) | Get-Random
   $Words[-1] = $Words[-1].insert($ToCap,$Words[-1][$ToCap].tostring().toupper()).remove($ToCap+1,1)

   # When working with a short word list, add a level of complexity 
   #  by changing one character to a number
   If ($list.count -le 500) {
      if ($Words.count -eq 1) {
         $skip = 1   # Avoid a number as the first character of the passphrase
      } else { $skip = 0 }

      $ToNum = $skip..($Words[-1].length -1) | Where-Object {$_ -ne $ToCap} | Get-Random
      $Words[-1] = $Words[-1].insert($ToNum,(0..9 | Get-Random)).remove($ToNum+1,1)
   }

   $candidate = $Words -join ''
   # When the length is annoyingly long, start again
   if ($candidate.length -gt ($Maximum - $Words.count)) {
      $Words = @()
   }
   # Need more than one word, but not more than necessary
   if (($candidate.length -ge ($Minimum - $Words.count)) -and ($Words.count -gt 1)) {
      $GotWords = $true
   }
}

# Words of the passphrase are separated with non-alphabetic characters.
# It is best to avoid certain printable and non-printable "problem
#  characters" when setting passwords.  All other standard, ASCII 
#  values should be OK.  Printable characters to avoid:
#      Technical: ampersand, less than, double quote and back slash
#      Perceptive: single quote, backtick, apostrophe
$Avoids = (34,38,39,44,60,92,96) + 65..90 + 97..122
# Also, don't avoid numerals as separators if they will be inserted in words
if ($list.count -le 500) { $Avoids += 48..57 }

# Candidates that are shy only one char double-up one word separator
if ($candidate.length -eq ($Minimum - $Words.count)) {
   $ExtraSpace = 0..($Words.count -2) | Get-Random
}

$Passphrase = ''
foreach ($i in (0..($Words.count-1))) {
   $Passphrase += $Words[$i]
   if ($Words[$i] -ne $Words[-1]) {
      do { $n = 33..126 | Get-Random } until ( $Avoids -notcontains $n )
      $Passphrase += [char]$n
      if ($i -eq $ExtraSpace) {
         do { $n = 33..126 | Get-Random } until ( $Avoids -notcontains $n )
         $Passphrase += [char]$n
      }
   }
}

If ($NetID -eq 'password') {
   # For testing, output instead of changing the password for this dummy account
   $Passphrase
} else {
   $adsi = New-Object System.directoryServices.directorySearcher([ADSI]"LDAP://$Domain")
   $adsi.filter = "(&(objectClass=Person)(samAccountName=$NetID))"

   $User = [adsi](($adsi.findall().getenumerator()).path)

   $User.psbase.invoke('SetPassword',$Passphrase)
   $User.psbase.CommitChanges()

   if ($LogPath) {
	     $LogString = $Passphrase
	     if ($TimeStamp) { $LogString = $LogString + '  -- set at ' + (get-date).tostring('g') }
	     if ($Overwrite) {Set-content -Path $LogPath -Value $LogString}
	     else {add-content -Path $LogPath -Value $LogString}
   }
}

Remove-Variable -Name Passphrase

