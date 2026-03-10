$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$exePath = Join-Path $scriptDir "dist\Right Click Menu Manager.exe"
$powerShellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$scriptPath = Join-Path $scriptDir "rightBtnMenu.ps1"

if ((-not (Test-Path -LiteralPath $exePath)) -and (-not (Test-Path -LiteralPath $scriptPath))) {
    throw "Neither EXE nor script launcher source was found."
}

$shell = New-Object -ComObject WScript.Shell
$shortcutTargets = @(
    (Join-Path ([Environment]::GetFolderPath("Desktop")) "Right Click Menu Manager.lnk"),
    (Join-Path ([Environment]::GetFolderPath("Programs")) "Right Click Menu Manager.lnk")
)

foreach ($shortcutPath in $shortcutTargets) {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    if (Test-Path -LiteralPath $exePath) {
        $shortcut.TargetPath = $exePath
        $shortcut.Arguments = ""
    }
    else {
        if (-not (Test-Path -LiteralPath $powerShellPath)) {
            throw "PowerShell not found: $powerShellPath"
        }

        $shortcut.TargetPath = $powerShellPath
        $shortcut.Arguments = ('-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -STA -File "{0}"' -f $scriptPath)
    }
    $shortcut.WorkingDirectory = $scriptDir
    $shortcut.WindowStyle = 1
    $shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,220"
    $shortcut.Description = "Open the right-click menu manager"
    $shortcut.Save()
}

Write-Host "Shortcuts created:"
foreach ($shortcutPath in $shortcutTargets) {
    Write-Host " - $shortcutPath"
}

