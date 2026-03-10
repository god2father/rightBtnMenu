$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$inputFile = Join-Path $scriptDir "rightBtnMenu.ps1"
$distDir = Join-Path $scriptDir "dist"
$outputFile = Join-Path $distDir "Right Click Menu Manager.exe"

if (-not (Test-Path -LiteralPath $inputFile)) {
    throw "Input script not found: $inputFile"
}

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}

Import-Module ps2exe

New-Item -ItemType Directory -Path $distDir -Force | Out-Null

Invoke-ps2exe `
    -inputFile $inputFile `
    -outputFile $outputFile `
    -STA `
    -noConsole `
    -title "Right Click Menu Manager" `
    -product "Right Click Menu Manager" `
    -company "Local Build" `
    -description "Windows right-click menu manager" `
    -version "1.0.0.0" `
    -supportOS `
    -DPIAware

Write-Host "EXE created:"
Write-Host " - $outputFile"
