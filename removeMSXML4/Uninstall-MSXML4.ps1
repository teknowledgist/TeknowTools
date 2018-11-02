# Removes any potential MSXML v4.x applications

$machine_key     = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
$machine_key6432 = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'

[array]$keys = Get-ChildItem -Path @($machine_key6432, $machine_key) -ErrorAction SilentlyContinue

# 'Get-ItemProperty`' fails if a registry key is encoded incorrectly.
[int]$maxAttempts = $keys.Count
for ([int]$attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    [bool]$success = $false

    $keyPaths = $keys | Select-Object -ExpandProperty PSPath
    try {
      [array]$foundKey = Get-ItemProperty -Path $keyPaths -ErrorAction Stop | Where-Object { $_.DisplayName -like 'MSXML 4*' }
      $success = $true
    } catch {
      Write-Debug 'Found bad key.'
      foreach ($key in $keys){
        try {
          Get-ItemProperty $key.PsPath > $null
        } catch {
          $badKey = $key.PsPath
        }
      }
      Write-Verbose "Skipping bad key: $badKey"
      [array]$keys = $keys | Where-Object { $badKey -NotContains $_.PsPath }
   }
   
   if ($success) { break; }

   if ($attempt -ge 10) {
      # Each key searched should correspond to an installed program. 
      # To have more than a few programs with incorrectly encoded keys may 
      # be indicative of one or more corrupted registry branches.
      $attempt = $maxAttempts
      Return
   }
}

if ($foundKey -eq $null -or $foundkey.Count -eq 0) {
    Return
}

Foreach ($Key in $foundKey) {
   $Code = Split-Path $Key.pspath -leaf
   $args = "/x $Code /qn /norestart /l*v `"$env:TEMP\$($Key.DisplayName).MsiInstall.log`""

   Start-Process 'msiexec.exe' -wait -ArgumentList $args
}

# If these are not removed by the uninstallers, manually delete them.
if (Test-Path '%SystemRoot%\System32\msxml4*.dll') {
   $null = Remove-Item '%SystemRoot%\System32\msxml4*.dll' -Force
}
if (Test-Path '%SystemRoot%\SysWOW64\msxml4*.dll') {
   $null = Remove-Item '%SystemRoot%\SysWOW64\msxml4*.dll' -Force
}
