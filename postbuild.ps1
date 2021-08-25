param([string]$targetName, [string]$outputDir)

Write-Host $targetName
Write-Host $outputDir

$pluginPath = "$Env:APPDATA\Elgato\StreamDeck\Plugins\$targetName.sdPlugin"

Write-Host $pluginPath

taskkill /f /im "StreamDeck.exe"

taskkill /f /im "$targetName.exe"

Remove-Item -Path "$pluginPath" -Force -Recurse

Copy-Item -Path "$outputDir" -Destination "$pluginPath" -Force -Recurse -Container: $false -Exclude "*.pdb"

Start-Process -FilePath "C:\Program Files\Elgato\StreamDeck\StreamDeck.exe"

Remove-Item "$PSScriptRoot\$targetName.streamDeckPlugin" -Force -Recurse

Write-Host "$($outputDir)$targetName.sdPlugin"

. "$PSScriptRoot\tools\DistributionTool.exe" -b -i "$pluginPath" -o "$PSScriptRoot" -NoNewWindow