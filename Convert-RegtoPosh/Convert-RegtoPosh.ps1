<#
.SYNOPSIS
   To convert a registry (.reg) file into a pure PowerShell script.

.DESCRIPTION
Intended Use
   This script was produced to assist with importing registry (.reg) RegFiles where the registry
   handlers, such as reg.exe or regedit.exe, are blocked from executing. It should output a 
   .PS1 script file with the same base name as each .REG file it processes.

About
   This is modified from a code snippet GIST by rileyz who modified code by Xeнεi Ξэnвϵς. 

Code Snippet Credits
   * https://superuser.com/questions/1614623/how-to-convert-reg-RegFiles-to-powershell-set-itemproperty-commands-automatically
   * https://gist.github.com/rileyz/52e721a609a9143158180f77b9f7ea0b

Version History 
   1.00 2025-04-04
      Initial release.

.LINK
Sourcecode:  https://github.com/teknowledgist/TeknowTools/tree/master/Convert-RegtoPosh

.Parameter Path
   Optional absolute or relative path to a REG file or to a directory.  If not provided, 
   the script directory path will be used to convert all REG files in the directory.

.Parameter Recursive
   Optional switch to recursively search the provided path directory or the script path 
   for all REG files in the branch. 

.Parameter Whatif
   Optional switch for the generated PowerShell script to use -Whatif for all lines
   involving changes.  The script should only report on what would happen if the
   REG file were converted to PowerShell without the -Whatif switch.

.EXAMPLE
   Convert-RegtoPosh.ps1 FileName.reg

   Convert a single registry file located in the same directory as this script.


.EXAMPLE
   Convert-RegtoPosh C:\Path\FileName.reg
    
   Convert a single registry file with absolute path.


.EXAMPLE
   "C:\Path\FileName.reg","C:\AnotherPath\FileName.reg" | .\Convert-RegtoPosh.ps1

   Pipeline regstry RegFiles to be converted by the script.


.EXAMPLE
   Convert-RegtoPosh.ps1 -Recursive -WhatIf

   Recursively import registry RegFiles from the script root and subdirectories.
   Also, the resulting PowerShell script(s) should be "neutered" with all actions 
   set to evaluate changes without actually making any.

.Notes
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



# Start of script work ############################################################################
[CmdLetBinding()]
Param(
   [Parameter(ValueFromPipeline = $true)][string]$Path,
   [Parameter][switch]$Recursive,
   [Parameter][switch]$WhatIf
)

Begin {
   $hive = @{
      "HKEY_CLASSES_ROOT"   = "Registry::HKEY_CLASSES_ROOT" # native path
      "HKEY_CURRENT_USER"   = "HKCU:" # alias 
      "HKEY_LOCAL_MACHINE"  = "HKLM:" # alias 
      "HKEY_USERS"          = "Registry::HKEY_USERS" # native path
      "HKEY_CURRENT_CONFIG" = "Registry::HKEY_CURRENT_CONFIG" #native path
   }
   if ($WhatIf) {
      Write-Warning 'WhatIf will throw errors if preformed on registry keys or subkey values which do not exist!'
      $WhatIfArg = $true
      $NullifyOutput = ''
   } else {
      $WhatIfArg = $False
      $NullifyOutput = '$Null = '
   }
   if ($Recursive) {$Recurse = $true} else {$Recurse = $False}

   if (($Path -eq $null) -or ($Path -eq '')) {
      Write-Verbose '$Path is empty. Using $PSScriptRoot as location.'
      $Path = $PSScriptRoot
   }
    
   if (Test-Path $Path -PathType Leaf) {
      Write-Verbose '$Path was passed a single filename only.'
      $RegFiles = (Get-ChildItem "$Path" -Include "*.reg").fullname
   }
   elseif (test-path $Path -PathType Container) {
      $RegFiles = (Get-ChildItem -Path $Path -Recurse:$Recurse -Force -filter '*.reg').FullName
   }

   function Convert-RegHex ([string]$HexKey){
      $hextype = $HexKey -replace '.*(hex\([27b]\)?).*', '$1'
      $HexArray = ($HexKey -replace '^.+=hex(\([2,7,b]\))?:', '').Split(",")
      switch ($hextype) {
         'hex(2)' {
            $type = "expandstring"
            $CharArray = for ($i = 0; $i -lt $HexArray.count; $i += 2) {
               if ($HexArray[$i] -ne '00') { 
                  [char][int]('0x' + $HexArray[$i]) 
               }
            }
            $value = "'$($CharArray -join '')'"
         }
         'hex(7)' {
            $type = 'multistring'
            $StrArray = for ($i = 0; $i -lt $HexArray.count; $i += 2) {
               if ($HexArray[$i] -ne '00') { 
                  [string][char][int]('0x' + $HexArray[$i]) 
               } else { 
                  ',' 
               }
            }
            $value = ("@('" + ($StrArray[0..($StrArray.count -2)] -join "','") + "')")
         }
         'hex(b)' {
            $type = 'qword'
            $QArray = for ($i = $HexArray.count - 1; $i -ge 0; $i--) { 
               $HexArray[$i] 
            }
            $value = '0x' + ($QArray -join '').trimstart('0')
         }
         'hex' {
            $type = 'binary'
            $value = '(0x' + ($HexArray -join ',0x') + ')'
         }
      }
      [PSCustomObject]@{
         KeyType = $type
         KeyValue = $value
      }
   }
}

Process {
   foreach ($File in $RegFiles) {
      $FileContent = Get-Content $File | Where-Object { ![string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
      $Commands = @()
      $addedpath = @()
      $joinedlines = @()
      [string]$text = $null
      for ($i = 0; $i -lt $FileContent.count; $i++) {
         if ($FileContent[$i].EndsWith("\")) {  # "\" is line continuation for .REG files
            $text = $text + $FileContent[$i].trimend('\').trim()
         }
         else {
            $joinedlines += $text + $FileContent[$i]
            [string]$text = $null
         }
      }
      Write-Debug "Contents of registry file: $((Get-ChildItem $File).Name) `r`n $($joinedlines | Out-String)"
        
      foreach ($FullLine in $joinedlines) {
         if ($FullLine.StartsWith(';')) {
            Write-Debug "*** Including a comment line : $FullLine"
            $Commands += "# $($FullLine.trimstart(';'))"
         }
         elseif ($FullLine -match '^\[-?HKEY_.*]$') {
            $key = $FullLine.trim('[]')
            Write-Debug "Processing registry key: $key"
            If ($key.StartsWith('-')) {
               $key = $key.trimstart('-')
               $HivePath = $key.split('\')[0]
               $key = '"' + ($key -replace $HivePath, $hive.$HivePath) + '"'
               Write-Debug "Registry key remove detected: $key"
               $Commands += 'Remove-Item -Path {0} -Force -Recurse -Whatif:${1}' -f $key, $WhatIfArg
            } else {
               $HivePath = $key.split('\')[0]
               $key = '"' + ($key -replace $HivePath, $hive.$HivePath) + '"'
               if ($addedpath -notcontains $key) {
                  Write-Debug " Registry key create/add detected: $key"
                  $Commands += '{0}New-Item -Path {1} -Whatif:${2} -ErrorAction SilentlyContinue | Out-Null' -f $NullifyOutput, $key, $WhatIfArg
                  $addedpath += $key
               }
               else {
                  Write-Debug " Registry key was already requested. Don't add it again."
               }

            }
         }
         elseif ($FullLine -match '^(".+?"|@)=') {
            Write-Debug "Processing registry value: $FullLine"
            $delete = $false
            $value = $null
            $name = $FullLine -replace '^("(.+?)"|@).*','$1'
            if ($name -eq '@') {
               $name = '"(Default)"'
            }
            switch ($FullLine) {
               { $FullLine -match "=-" } {
                  $delete = $true
                  Write-Debug ' Registry item remove detected.'

               }
               { $FullLine -match '("|@)="' } {
                  $type = "string"
                  $value = $FullLine -replace '^(@|".+?")=','' -replace '\\\\', '\' -replace "'`"",'`"'
               }
               { $FullLine -match '=dword:' } {
                  $type = "dword"
                  $value = '0x' + $FullLine.substring($FullLine.indexof('=dword:') + 7)
               }
               { $FullLine -match '=qword:' } {
                  $type = "qword"
                  $value = '0x' + $FullLine.substring($FullLine.indexof('=qword:') + 7)
               }
               { $FullLine -match "hex(\([2,7,b]\))?:" } {
                  $converted = Convert-RegHex $FullLine
                  $type = $converted.KeyType
                  $value = $converted.KeyValue
               }
            }
            if ($delete -eq $false) {
               Write-Debug " Registry item add/set detected, the registry item type is '$type'."
               $Commands += '{0}New-ItemProperty -Path {1} -Name {2} -PropertyType {3} -Value {4} -Force -Whatif:${5}' -f $NullifyOutput, $key, $name, $type, $value, $WhatIfArg
 
            } else {
               if ($name -eq '"(Default)"') {
                  # Can't delete default item, only empty it
                  Write-Debug ' Registry key default item remove detected.'
                  $Commands += '$(Get-Item -Path {0} -Whatif:${1} ).OpenSubKey("", $true).DeleteValue("")' -f $key, $WhatIfArg
               } else {
                  $Commands += 'Remove-ItemProperty -Path {0} -Name {1} -Force -Whatif:${2}' -f $key, $Name, $WhatIfArg
               }
            }
         } else {
            Write-Debug "*** Skipping a line with unknown syntax : $FullLine"
         }
      }

      $parent = Split-Path $file -Parent
      $filename = [System.IO.Path]::GetFileNameWithoutExtension($file)
      $Commands | Out-File -FilePath "$parent\${filename}_reg.ps1" -Encoding utf8
      Write-Debug "Created file: `"$parent\${filename}_reg.ps1`""
   }
}

End {
}
#<<< End of script work >>>

