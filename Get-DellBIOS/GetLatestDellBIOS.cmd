@@echo off
::============================================::
:: This "polyglot script" allows a PowerShell ::
:: script to be embedded in a .CMD file.  Any ::
:: batch lines must begin with "@@" or "::"   ::
::============================================::
@@set POWERSHELL_BAT_ARGS=%*
@@if defined POWERSHELL_BAT_ARGS set POWERSHELL_BAT_ARGS=%POWERSHELL_BAT_ARGS:"=\"%
@@PowerShell -ExecutionPolicy ByPass -noprofile -Command Invoke-Expression $('$args=@(^&{$args} %POWERSHELL_BAT_ARGS%);'+[String]::Join([Environment]::NewLine,$((Get-Content '%~f0') -notmatch '^^@@^|^^::'))) & goto :AfterPoSh
{ 
# Start PowerShell

# Process extracted from:  https://github.com/maurice-daly/DriverAutomationTool
If ((Get-WmiObject Win32_ComputerSystem).Manufacturer -match "Dell") {
   $Desktop = [Environment]::GetFolderPath("Desktop")

   $Model = (Get-WmiObject Win32_ComputerSystem).Model.trim()

   $DownloadBase = 'https://downloads.dell.com'
   $SKUreferenceURL = 'https://downloads.dell.com/catalog/DriverPackCatalog.cab'
   $CatalogURL = 'https://downloads.dell.com/catalog/CatalogPC.cab'

   $BitsOptions = @{
      RetryInterval = '60'
      RetryTimeout = '180'
      Priority = 'Foreground'
      TransferType = 'Download'
   }

   $ReferenceCabFile = [string]($SKUreferenceURL | Split-Path -Leaf)
   $CatalogCabFile = [string]($CatalogURL | Split-Path -Leaf)
   $RefXMLFile = $ReferenceCabFile.TrimEnd('.cab') + '.xml'
   $CatalogXMLFile = $CatalogCabFile.TrimEnd('.cab') + '.xml'

   # Get the reference file connecting the model name with the SKU/SystemID
   Start-BitsTransfer -Source $SKUreferenceURL -Destination $env:TEMP @BitsOptions
   & "$env:windir\system32\expand.exe" "$env:TEMP\$ReferenceCabFile" -F:* "$env:TEMP" -R | Out-Null
   [xml]$ReferenceXML = Get-Content -Path (Join-Path -Path $env:TEMP -ChildPath $RefXMLFile) -Raw

   $SkuValue = (($ReferenceXML.driverpackmanifest.driverpackage.supportedsystems.brand.model | 
                           Where-Object {$_.Name -eq $Model}
                       ).systemID) | Select-Object -Unique -First 1

   # Now get the Catalog file which points to the latest BIOS file URL
   Start-BitsTransfer -Source $CatalogURL -Destination $env:TEMP @BitsOptions
   & "$env:windir\system32\expand.exe" "$env:TEMP\$CatalogCabFile" -F:* "$env:TEMP" -R | Out-Null
   [xml]$CatalogXML = Get-Content -Path $(Join-Path -Path $env:TEMP -ChildPath $CatalogXMLFile) -Raw

   $BiosURLstub = $CatalogXML.Manifest.SoftwareComponent | 
                     Where-Object {
                        ($_.name.display.'#cdata-section' -match 'BIOS') -and 
                        (Compare-Object $SkuValue $_.SupportedSystems.Brand.Model.SystemID -ExcludeDifferent -IncludeEqual)
                     } | Sort-Object ReleaseDate | Select-Object -First 1

   $BIOSURL = $DownloadBase + '/' + $BiosURLstub.Path.trim('/')
   $BIOSfile = $BIOSURL.split('/')[-1]

   $DownloadFile = Join-Path $Desktop "Latest BIOS - $BIOSfile"

   Write-Host "Downloading latest BIOS update file for $Model." -ForegroundColor Green 
   (New-Object System.Net.WebClient).DownloadFile($BIOSurl, $DownloadFile)

} else {
   Write-Warning "This script only works for Dell computers."
}

read-host 'Press Enter to Close'

# End PowerShell
}.Invoke($args)

#:AfterPoSh
:: Batch commands (begining with "@@" or "::") 
:: can follow this label.

