Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class RightBtnMenuShellNative
{
    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
"@

$readScopeMap = [ordered]@{
    "文件" = "Registry::HKEY_CLASSES_ROOT\*\shell"
    "文件夹" = "Registry::HKEY_CLASSES_ROOT\Directory\shell"
    "文件夹空白处" = "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell"
    "桌面空白处" = "Registry::HKEY_CLASSES_ROOT\DesktopBackground\Shell"
}

$writeScopeMap = [ordered]@{
    "文件" = "HKCU:\Software\Classes\`*\shell"
    "文件夹" = "HKCU:\Software\Classes\Directory\shell"
    "文件夹空白处" = "HKCU:\Software\Classes\Directory\Background\shell"
    "桌面空白处" = "HKCU:\Software\Classes\DesktopBackground\Shell"
}

function Get-DisplayText {
    param(
        $Properties,
        [string]$KeyName
    )

    $menuText = [string]$Properties.MUIVerb
    if (-not [string]::IsNullOrWhiteSpace($menuText)) {
        return $menuText
    }

    $defaultText = [string]$Properties."(default)"
    if (-not [string]::IsNullOrWhiteSpace($defaultText)) {
        return $defaultText
    }

    return $KeyName
}

function Get-ItemSource {
    param(
        [string]$RegistryPath
    )

    if ($RegistryPath -like "Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER*") {
        return "当前用户"
    }

    if ($RegistryPath -like "Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE*") {
        return "系统"
    }

    if ($RegistryPath -like "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT*") {
        return "系统/合并视图"
    }

    return "未知"
}

function Resolve-ExecutablePath {
    param(
        [string]$PathValue
    )

    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return $null
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($PathValue.Trim())
    if ($expanded.StartsWith('"') -and $expanded.EndsWith('"')) {
        $expanded = $expanded.Trim('"')
    }

    if ([System.IO.Path]::IsPathRooted($expanded) -and (Test-Path -LiteralPath $expanded)) {
        return (Resolve-Path -LiteralPath $expanded).Path
    }

    $exeName = [System.IO.Path]::GetFileName($expanded)
    if (-not [string]::IsNullOrWhiteSpace($exeName)) {
        $appPathKeys = @(
            "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths\$exeName",
            "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\$exeName"
        )

        foreach ($appPathKey in $appPathKeys) {
            if (-not (Test-Path -LiteralPath $appPathKey)) {
                continue
            }

            $registeredPath = Resolve-ExecutablePath -PathValue (Get-Item -LiteralPath $appPathKey).GetValue("")
            if (-not [string]::IsNullOrWhiteSpace($registeredPath)) {
                return $registeredPath
            }
        }
    }

    try {
        $command = Get-Command -Name $expanded -CommandType Application -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
            return $command.Source
        }
    }
    catch {
    }

    if ($exeName -ieq "code.exe") {
        $commonVsCodePaths = @(
            "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe",
            "$env:ProgramFiles\Microsoft VS Code\Code.exe",
            "${env:ProgramFiles(x86)}\Microsoft VS Code\Code.exe"
        )

        foreach ($candidate in $commonVsCodePaths) {
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
    }

    return $expanded
}

function Get-ExecutableFromCommand {
    param(
        [string]$CommandText
    )

    if ([string]::IsNullOrWhiteSpace($CommandText)) {
        return $null
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($CommandText)
    if ($expanded -match '^\s*"([^"]+?\.exe)"') {
        return $matches[1]
    }

    if ($expanded -match '^\s*([^"\s]+?\.exe)\b') {
        return $matches[1]
    }

    return $null
}

function Get-ApplicationDisplayName {
    param(
        $Properties,
        [string]$KeyName
    )

    $friendlyName = [string]$Properties.FriendlyAppName
    if (-not [string]::IsNullOrWhiteSpace($friendlyName)) {
        return $friendlyName
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($KeyName)
}

function Get-ScopeTargetArgument {
    param(
        [string]$ScopeName
    )

    switch ($ScopeName) {
        "文件" { return "%1" }
        "文件夹" { return "%1" }
        "文件夹空白处" { return "%V" }
        "桌面空白处" { return "%V" }
        default { return "%1" }
    }
}

function Get-DisplayNameFromExecutable {
    param(
        [string]$ExePath,
        [string]$FallbackName
    )

    $resolvedExePath = Resolve-ExecutablePath -PathValue $ExePath
    if (-not [string]::IsNullOrWhiteSpace($resolvedExePath) -and (Test-Path -LiteralPath $resolvedExePath)) {
        try {
            $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($resolvedExePath)
            if (-not [string]::IsNullOrWhiteSpace($versionInfo.FileDescription)) {
                return [string]$versionInfo.FileDescription
            }
        }
        catch {
        }
    }

    return $FallbackName
}

function New-CommandTextForApplication {
    param(
        [string]$ExePath,
        [string]$ScopeName
    )

    $resolvedExePath = Resolve-ExecutablePath -PathValue $ExePath
    $targetArgument = Get-ScopeTargetArgument -ScopeName $ScopeName
    $exeName = [System.IO.Path]::GetFileName($resolvedExePath)

    if ($exeName -ieq "Code.exe") {
        return ('"{0}" --reuse-window "{1}"' -f $resolvedExePath, $targetArgument)
    }

    return ('"{0}" "{1}"' -f $resolvedExePath, $targetArgument)
}

function Split-CommandText {
    param(
        [string]$CommandText
    )

    if ([string]::IsNullOrWhiteSpace($CommandText)) {
        return $null
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($CommandText.Trim())
    if ($expanded -match '^\s*"([^"]+)"\s*(.*)$') {
        return [PSCustomObject]@{
            Executable = $matches[1]
            Arguments = [string]$matches[2]
        }
    }

    if ($expanded -match '^\s*([^\s]+)\s*(.*)$') {
        return [PSCustomObject]@{
            Executable = $matches[1]
            Arguments = [string]$matches[2]
        }
    }

    return $null
}

function Test-CommandExecutable {
    param(
        [string]$CommandText
    )

    $parts = Split-CommandText -CommandText $CommandText
    if ($null -eq $parts) {
        return [PSCustomObject]@{
            IsValid = $false
            Executable = $null
            Arguments = $null
            Message = "命令格式无法识别。"
        }
    }

    $resolvedExecutable = Resolve-ExecutablePath -PathValue $parts.Executable
    if ([string]::IsNullOrWhiteSpace($resolvedExecutable) -or -not (Test-Path -LiteralPath $resolvedExecutable)) {
        return [PSCustomObject]@{
            IsValid = $false
            Executable = $resolvedExecutable
            Arguments = [string]$parts.Arguments
            Message = "未找到可执行文件：$($parts.Executable)"
        }
    }

    return [PSCustomObject]@{
        IsValid = $true
        Executable = $resolvedExecutable
        Arguments = [string]$parts.Arguments
        Message = ""
    }
}

function Get-TestTargetPath {
    param(
        [string]$ScopeName
    )

    $tempRoot = Join-Path $env:TEMP "rightBtnMenu-test"
    if (-not (Test-Path -LiteralPath $tempRoot)) {
        New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
    }

    switch ($ScopeName) {
        "文件" {
            $filePath = Join-Path $tempRoot "sample.txt"
            if (-not (Test-Path -LiteralPath $filePath)) {
                Set-Content -LiteralPath $filePath -Value "rightBtnMenu test" -Encoding UTF8
            }
            return $filePath
        }
        "文件夹" { return $tempRoot }
        "文件夹空白处" { return $tempRoot }
        "桌面空白处" { return [Environment]::GetFolderPath("Desktop") }
        default { return $tempRoot }
    }
}

function Expand-CommandArguments {
    param(
        [string]$Arguments,
        [string]$TargetPath
    )

    $result = [string]$Arguments
    if (-not [string]::IsNullOrWhiteSpace($TargetPath)) {
        $quotedTarget = '"{0}"' -f $TargetPath
        $result = $result.Replace('"%1"', $quotedTarget)
        $result = $result.Replace('"%V"', $quotedTarget)
        $result = $result.Replace('"%L"', $quotedTarget)
        $result = $result.Replace("%1", $quotedTarget)
        $result = $result.Replace("%V", $quotedTarget)
        $result = $result.Replace("%L", $quotedTarget)
    }

    return $result
}

function Refresh-ShellMenuCache {
    [RightBtnMenuShellNative]::SHChangeNotify(0x08000000, 0x1000, [IntPtr]::Zero, [IntPtr]::Zero)
}

function Get-RegisteredApplications {
    $itemsById = @{}

    $appPathSources = @(
        @{ Path = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths"; Source = "当前用户 App Paths" },
        @{ Path = "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths"; Source = "系统 App Paths" }
    )

    foreach ($source in $appPathSources) {
        if (-not (Test-Path -LiteralPath $source.Path)) {
            continue
        }

        foreach ($key in Get-ChildItem -LiteralPath $source.Path -ErrorAction SilentlyContinue) {
            $keyItem = Get-Item -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ($null -eq $keyItem) {
                continue
            }

            $exePath = Resolve-ExecutablePath -PathValue $keyItem.GetValue("")
            if ([string]::IsNullOrWhiteSpace($exePath)) {
                continue
            }

            $resolvedExePath = Resolve-ExecutablePath -PathValue $exePath
            $id = $resolvedExePath.ToLowerInvariant()
            if (-not $itemsById.ContainsKey($id)) {
                $itemsById[$id] = [PSCustomObject]@{
                    DisplayName = Get-DisplayNameFromExecutable -ExePath $exePath -FallbackName ([System.IO.Path]::GetFileNameWithoutExtension($key.PSChildName))
                    ExePath = $resolvedExePath
                    Source = [string]$source.Source
                }
            }
        }
    }

    $applicationsPath = "Registry::HKEY_CLASSES_ROOT\Applications"
    if (Test-Path -LiteralPath $applicationsPath) {
        foreach ($key in Get-ChildItem -LiteralPath $applicationsPath -ErrorAction SilentlyContinue) {
            $commandPath = Join-Path $key.PSPath "shell\open\command"
            if (-not (Test-Path -LiteralPath $commandPath)) {
                continue
            }

            $commandText = (Get-Item -LiteralPath $commandPath -ErrorAction SilentlyContinue).GetValue("")
            $exePath = Get-ExecutableFromCommand -CommandText $commandText
            if ([string]::IsNullOrWhiteSpace($exePath)) {
                continue
            }

            $resolvedExePath = Resolve-ExecutablePath -PathValue $exePath
            $id = $resolvedExePath.ToLowerInvariant()
            if (-not $itemsById.ContainsKey($id)) {
                $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
                $itemsById[$id] = [PSCustomObject]@{
                    DisplayName = Get-DisplayNameFromExecutable -ExePath $resolvedExePath -FallbackName (Get-ApplicationDisplayName -Properties $properties -KeyName $key.PSChildName)
                    ExePath = $resolvedExePath
                    Source = "Applications"
                }
            }
        }
    }

    return $itemsById.Values | Sort-Object DisplayName, ExePath
}

function Show-RegisteredApplicationPicker {
    param(
        [string]$ScopeName
    )

    $applications = @(Get-RegisteredApplications)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "选择已注册程序"
    $dialog.Size = New-Object System.Drawing.Size(820, 540)
    $dialog.StartPosition = "CenterParent"
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "搜索"
    $searchLabel.Location = New-Object System.Drawing.Point(20, 20)
    $searchLabel.AutoSize = $true
    $dialog.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(70, 16)
    $searchBox.Size = New-Object System.Drawing.Size(520, 24)
    $dialog.Controls.Add($searchBox)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = ""
    $countLabel.Location = New-Object System.Drawing.Point(610, 20)
    $countLabel.Size = New-Object System.Drawing.Size(180, 20)
    $dialog.Controls.Add($countLabel)

    $appList = New-Object System.Windows.Forms.ListView
    $appList.Location = New-Object System.Drawing.Point(20, 55)
    $appList.Size = New-Object System.Drawing.Size(770, 395)
    $appList.View = "Details"
    $appList.FullRowSelect = $true
    $appList.GridLines = $true
    $appList.HideSelection = $false
    [void]$appList.Columns.Add("程序名称", 220)
    [void]$appList.Columns.Add("可执行文件", 410)
    [void]$appList.Columns.Add("来源", 120)
    $dialog.Controls.Add($appList)

    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Text = "选择"
    $selectButton.Location = New-Object System.Drawing.Point(520, 465)
    $selectButton.Size = New-Object System.Drawing.Size(110, 32)
    $dialog.Controls.Add($selectButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "取消"
    $cancelButton.Location = New-Object System.Drawing.Point(650, 465)
    $cancelButton.Size = New-Object System.Drawing.Size(110, 32)
    $dialog.Controls.Add($cancelButton)

    $hintLabel = New-Object System.Windows.Forms.Label
    $hintLabel.Text = "选择后会自动填充菜单名称、图标和命令；你只需要点保存。"
    $hintLabel.Location = New-Object System.Drawing.Point(20, 472)
    $hintLabel.Size = New-Object System.Drawing.Size(470, 20)
    $dialog.Controls.Add($hintLabel)

    $dialog.AcceptButton = $selectButton
    $dialog.CancelButton = $cancelButton

    function Update-ApplicationList {
        param(
            [string]$Keyword
        )

        $appList.Items.Clear()
        $needle = [string]$Keyword
        $filtered = $applications

        if (-not [string]::IsNullOrWhiteSpace($needle)) {
            $filtered = $applications | Where-Object {
                $_.DisplayName -like "*$needle*" -or
                $_.ExePath -like "*$needle*" -or
                $_.Source -like "*$needle*"
            }
        }

        foreach ($app in $filtered) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$app.DisplayName)
            [void]$item.SubItems.Add([string]$app.ExePath)
            [void]$item.SubItems.Add([string]$app.Source)
            $item.Tag = $app
            [void]$appList.Items.Add($item)
        }

        $countLabel.Text = "共 $(@($filtered).Count) 项"
        if ($appList.Items.Count -gt 0) {
            $appList.Items[0].Selected = $true
        }
    }

    $searchBox.Add_TextChanged({
        Update-ApplicationList -Keyword $searchBox.Text
    })

    $cancelButton.Add_Click({
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Close()
    })

    $selectAction = {
        if ($appList.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("请先选择一个程序。")
            return
        }

        $dialog.Tag = $appList.SelectedItems[0].Tag
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    }

    $selectButton.Add_Click($selectAction)
    $appList.Add_DoubleClick($selectAction)

    Update-ApplicationList -Keyword ""

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $selectedApp = $dialog.Tag
    if ($null -eq $selectedApp) {
        return $null
    }

    return [PSCustomObject]@{
        MenuText = [string]$selectedApp.DisplayName
        Icon = [string](Resolve-ExecutablePath -PathValue $selectedApp.ExePath)
        Command = New-CommandTextForApplication -ExePath $selectedApp.ExePath -ScopeName $ScopeName
    }
}

function Get-MenuItems {
    param(
        [string]$RegistryPath
    )

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        return @()
    }

    $items = @()
    foreach ($key in Get-ChildItem -LiteralPath $RegistryPath -ErrorAction SilentlyContinue) {
        $properties = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
        $commandPath = Join-Path $key.PSPath "command"
        $commandValue = ""

        if (Test-Path -LiteralPath $commandPath) {
            $commandValue = (Get-Item -LiteralPath $commandPath -ErrorAction SilentlyContinue).GetValue("")
        }

        $items += [PSCustomObject]@{
            KeyName = $key.PSChildName
            MenuText = Get-DisplayText -Properties $properties -KeyName $key.PSChildName
            Icon = $properties.Icon
            Command = $commandValue
            Enabled = -not ($properties.PSObject.Properties.Name -contains "LegacyDisable")
            Source = Get-ItemSource -RegistryPath $key.PSPath
            ItemPath = $key.PSPath
        }
    }

    return $items | Sort-Object MenuText, KeyName
}

function Save-MenuItem {
    param(
        [string]$RegistryRoot,
        [string]$KeyName,
        [string]$MenuText,
        [string]$Command,
        [string]$Icon
    )

    $targetPath = Join-Path $RegistryRoot $KeyName
    New-Item -Path $targetPath -Force | Out-Null
    New-ItemProperty -LiteralPath $targetPath -Name "MUIVerb" -Value $MenuText -PropertyType String -Force | Out-Null

    if ([string]::IsNullOrWhiteSpace($Icon)) {
        Remove-ItemProperty -LiteralPath $targetPath -Name "Icon" -ErrorAction SilentlyContinue
    }
    else {
        New-ItemProperty -LiteralPath $targetPath -Name "Icon" -Value $Icon -PropertyType String -Force | Out-Null
    }

    $commandPath = Join-Path $targetPath "command"
    New-Item -Path $commandPath -Force | Out-Null
    Set-Item -LiteralPath $commandPath -Value $Command -Force
}

function Remove-MenuItem {
    param(
        [string]$ItemPath
    )

    if (Test-Path -LiteralPath $ItemPath) {
        Remove-Item -LiteralPath $ItemPath -Recurse -Force
    }
}

function Set-MenuItemEnabled {
    param(
        [string]$ItemPath,
        [bool]$Enabled
    )

    if (-not (Test-Path -LiteralPath $ItemPath)) {
        return
    }

    if ($Enabled) {
        Remove-ItemProperty -LiteralPath $ItemPath -Name "LegacyDisable" -ErrorAction SilentlyContinue
    }
    else {
        New-ItemProperty -LiteralPath $ItemPath -Name "LegacyDisable" -Value "" -PropertyType String -Force | Out-Null
    }
}

function New-SafeKeyName {
    param(
        [string]$Source
    )

    $value = ($Source -replace "[^a-zA-Z0-9_-]", "_").Trim("_")
    if ([string]::IsNullOrWhiteSpace($value)) {
        return "menu_" + [DateTime]::Now.ToString("yyyyMMddHHmmss")
    }

    return $value
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "右键菜单管理"
$form.Size = New-Object System.Drawing.Size(880, 620)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$scopeLabel = New-Object System.Windows.Forms.Label
$scopeLabel.Text = "范围"
$scopeLabel.Location = New-Object System.Drawing.Point(20, 20)
$scopeLabel.AutoSize = $true
$form.Controls.Add($scopeLabel)

$scopeBox = New-Object System.Windows.Forms.ComboBox
$scopeBox.Location = New-Object System.Drawing.Point(80, 16)
$scopeBox.Size = New-Object System.Drawing.Size(240, 28)
$scopeBox.DropDownStyle = "DropDownList"
$readScopeMap.Keys | ForEach-Object { [void]$scopeBox.Items.Add([string]$_) }
if ($scopeBox.Items.Count -gt 0) {
    $scopeBox.SelectedIndex = 0
}
$form.Controls.Add($scopeBox)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "刷新"
$refreshButton.Location = New-Object System.Drawing.Point(340, 15)
$refreshButton.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($refreshButton)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 60)
$grid.Size = New-Object System.Drawing.Size(830, 260)
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.RowHeadersVisible = $false
$form.Controls.Add($grid)

$group = New-Object System.Windows.Forms.GroupBox
$group.Text = "菜单项"
$group.Location = New-Object System.Drawing.Point(20, 340)
$group.Size = New-Object System.Drawing.Size(830, 210)
$form.Controls.Add($group)

$keyLabel = New-Object System.Windows.Forms.Label
$keyLabel.Text = "注册表键名"
$keyLabel.Location = New-Object System.Drawing.Point(20, 35)
$keyLabel.AutoSize = $true
$group.Controls.Add($keyLabel)

$keyBox = New-Object System.Windows.Forms.TextBox
$keyBox.Location = New-Object System.Drawing.Point(110, 30)
$keyBox.Size = New-Object System.Drawing.Size(280, 24)
$group.Controls.Add($keyBox)

$menuLabel = New-Object System.Windows.Forms.Label
$menuLabel.Text = "菜单名称"
$menuLabel.Location = New-Object System.Drawing.Point(420, 35)
$menuLabel.AutoSize = $true
$group.Controls.Add($menuLabel)

$menuBox = New-Object System.Windows.Forms.TextBox
$menuBox.Location = New-Object System.Drawing.Point(510, 30)
$menuBox.Size = New-Object System.Drawing.Size(280, 24)
$group.Controls.Add($menuBox)

$iconLabel = New-Object System.Windows.Forms.Label
$iconLabel.Text = "图标"
$iconLabel.Location = New-Object System.Drawing.Point(20, 80)
$iconLabel.AutoSize = $true
$group.Controls.Add($iconLabel)

$iconBox = New-Object System.Windows.Forms.TextBox
$iconBox.Location = New-Object System.Drawing.Point(110, 75)
$iconBox.Size = New-Object System.Drawing.Size(680, 24)
$group.Controls.Add($iconBox)

$commandLabel = New-Object System.Windows.Forms.Label
$commandLabel.Text = "命令"
$commandLabel.Location = New-Object System.Drawing.Point(20, 125)
$commandLabel.AutoSize = $true
$group.Controls.Add($commandLabel)

$commandBox = New-Object System.Windows.Forms.TextBox
$commandBox.Location = New-Object System.Drawing.Point(110, 120)
$commandBox.Size = New-Object System.Drawing.Size(680, 24)
$group.Controls.Add($commandBox)

$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text = "当前：新增模式"
$modeLabel.Location = New-Object System.Drawing.Point(20, 165)
$modeLabel.Size = New-Object System.Drawing.Size(85, 20)
$group.Controls.Add($modeLabel)

$newButton = New-Object System.Windows.Forms.Button
$newButton.Text = "新增"
$newButton.Location = New-Object System.Drawing.Point(110, 160)
$newButton.Size = New-Object System.Drawing.Size(85, 30)
$group.Controls.Add($newButton)

$pickAppButton = New-Object System.Windows.Forms.Button
$pickAppButton.Text = "选择程序"
$pickAppButton.Location = New-Object System.Drawing.Point(205, 160)
$pickAppButton.Size = New-Object System.Drawing.Size(85, 30)
$group.Controls.Add($pickAppButton)

$testButton = New-Object System.Windows.Forms.Button
$testButton.Text = "测试命令"
$testButton.Location = New-Object System.Drawing.Point(300, 160)
$testButton.Size = New-Object System.Drawing.Size(85, 30)
$group.Controls.Add($testButton)

$addSaveButton = New-Object System.Windows.Forms.Button
$addSaveButton.Text = "新增保存"
$addSaveButton.Location = New-Object System.Drawing.Point(395, 160)
$addSaveButton.Size = New-Object System.Drawing.Size(85, 30)
$group.Controls.Add($addSaveButton)

$updateSaveButton = New-Object System.Windows.Forms.Button
$updateSaveButton.Text = "更新保存"
$updateSaveButton.Location = New-Object System.Drawing.Point(490, 160)
$updateSaveButton.Size = New-Object System.Drawing.Size(85, 30)
$updateSaveButton.Enabled = $false
$group.Controls.Add($updateSaveButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "删除"
$deleteButton.Location = New-Object System.Drawing.Point(585, 160)
$deleteButton.Size = New-Object System.Drawing.Size(85, 30)
$deleteButton.Enabled = $false
$group.Controls.Add($deleteButton)

$toggleButton = New-Object System.Windows.Forms.Button
$toggleButton.Text = "禁用"
$toggleButton.Location = New-Object System.Drawing.Point(680, 160)
$toggleButton.Size = New-Object System.Drawing.Size(85, 30)
$toggleButton.Enabled = $false
$group.Controls.Add($toggleButton)

$helpLabel = New-Object System.Windows.Forms.Label
$helpLabel.Text = '命令示例：notepad.exe "%1"   |   Code.exe --reuse-window "%V"   |   powershell.exe -NoProfile -Command "Write-Host hello"'
$helpLabel.Location = New-Object System.Drawing.Point(20, 560)
$helpLabel.Size = New-Object System.Drawing.Size(830, 20)
$form.Controls.Add($helpLabel)

function Clear-Editor {
    $keyBox.Text = ""
    $menuBox.Text = ""
    $iconBox.Text = ""
    $commandBox.Text = ""
    $grid.ClearSelection()
    Update-EditorState
}

function Load-Grid {
    $selectedScope = [string]$scopeBox.SelectedItem
    if ([string]::IsNullOrWhiteSpace($selectedScope)) {
        $grid.Columns.Clear()
        $grid.Rows.Clear()
        Update-EditorState
        return
    }

    $registryPath = $readScopeMap[$selectedScope]
    $items = @(Get-MenuItems -RegistryPath $registryPath)

    $grid.Columns.Clear()
    $grid.Rows.Clear()

    [void]$grid.Columns.Add("KeyName", "键名")
    [void]$grid.Columns.Add("MenuText", "菜单名称")
    [void]$grid.Columns.Add("Icon", "图标")
    [void]$grid.Columns.Add("Command", "命令")
    [void]$grid.Columns.Add("Enabled", "启用")
    [void]$grid.Columns.Add("Source", "来源")
    [void]$grid.Columns.Add("ItemPath", "路径")
    $grid.Columns["ItemPath"].Visible = $false

    foreach ($item in $items) {
        [void]$grid.Rows.Add(
            [string]$item.KeyName,
            [string]$item.MenuText,
            [string]$item.Icon,
            [string]$item.Command,
            $(if ($item.Enabled) { "是" } else { "否" }),
            [string]$item.Source,
            [string]$item.ItemPath
        )
    }

    Update-EditorState
}

function Update-ToggleButtonText {
    if ($grid.SelectedRows.Count -eq 0) {
        $toggleButton.Text = "禁用"
        return
    }

    $row = $grid.SelectedRows[0]
    if ([string]$row.Cells["Enabled"].Value -eq "是") {
        $toggleButton.Text = "禁用"
    }
    else {
        $toggleButton.Text = "启用"
    }
}

function Get-UserWritableItemPath {
    param(
        [string]$ScopeName,
        [string]$KeyName
    )

    if ([string]::IsNullOrWhiteSpace($ScopeName) -or [string]::IsNullOrWhiteSpace($KeyName)) {
        return $null
    }

    $registryRoot = $writeScopeMap[$ScopeName]
    if ([string]::IsNullOrWhiteSpace($registryRoot)) {
        return $null
    }

    $targetPath = Join-Path $registryRoot $KeyName
    if (Test-Path -LiteralPath $targetPath) {
        return $targetPath
    }

    return $null
}

function Update-EditorState {
    if ($grid.SelectedRows.Count -eq 0) {
        $modeLabel.Text = "当前：新增模式"
        $updateSaveButton.Enabled = $false
        $deleteButton.Enabled = $false
        $toggleButton.Enabled = $false
        Update-ToggleButtonText
        return
    }

    $selectedScope = [string]$scopeBox.SelectedItem
    $row = $grid.SelectedRows[0]
    $keyName = [string]$row.Cells["KeyName"].Value
    $userItemPath = Get-UserWritableItemPath -ScopeName $selectedScope -KeyName $keyName

    if ([string]::IsNullOrWhiteSpace($userItemPath)) {
        $modeLabel.Text = "当前：系统项只读"
        $updateSaveButton.Enabled = $false
        $deleteButton.Enabled = $false
        $toggleButton.Enabled = $false
    }
    else {
        $modeLabel.Text = "当前：编辑模式"
        $updateSaveButton.Enabled = $true
        $deleteButton.Enabled = $true
        $toggleButton.Enabled = $true
    }

    Update-ToggleButtonText
}

function Select-GridRowByKeyName {
    param(
        [string]$KeyName
    )

    if ([string]::IsNullOrWhiteSpace($KeyName)) {
        return
    }

    foreach ($row in $grid.Rows) {
        if ([string]$row.Cells["KeyName"].Value -eq $KeyName) {
            $row.Selected = $true
            break
        }
    }
}

function Get-EditorDraft {
    $menuText = $menuBox.Text.Trim()
    $commandText = $commandBox.Text.Trim()
    $keyName = $keyBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($menuText)) {
        [System.Windows.Forms.MessageBox]::Show("菜单名称不能为空。")
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($commandText)) {
        [System.Windows.Forms.MessageBox]::Show("命令不能为空。")
        return $null
    }

    $validation = Test-CommandExecutable -CommandText $commandText
    if (-not $validation.IsValid) {
        [System.Windows.Forms.MessageBox]::Show($validation.Message, "保存失败")
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($keyName)) {
        $keyName = New-SafeKeyName -Source $menuText
        $keyBox.Text = $keyName
    }

    return [PSCustomObject]@{
        KeyName = $keyName
        MenuText = $menuText
        CommandText = $commandText
        IconText = $iconBox.Text.Trim()
    }
}

$refreshButton.Add_Click({
    Load-Grid
})

$scopeBox.Add_SelectedIndexChanged({
    Clear-Editor
    Load-Grid
})

$newButton.Add_Click({
    Clear-Editor
})

$pickAppButton.Add_Click({
    $selectedScope = [string]$scopeBox.SelectedItem
    if ([string]::IsNullOrWhiteSpace($selectedScope)) {
        [System.Windows.Forms.MessageBox]::Show("请先选择右键菜单范围。")
        return
    }

    $selectedApplication = Show-RegisteredApplicationPicker -ScopeName $selectedScope
    if ($null -eq $selectedApplication) {
        return
    }

    $menuBox.Text = [string]$selectedApplication.MenuText
    $iconBox.Text = [string]$selectedApplication.Icon
    $commandBox.Text = [string]$selectedApplication.Command

    if ([string]::IsNullOrWhiteSpace($keyBox.Text)) {
        $keyBox.Text = New-SafeKeyName -Source $selectedApplication.MenuText
    }
})

$testButton.Add_Click({
    $selectedScope = [string]$scopeBox.SelectedItem
    $validation = Test-CommandExecutable -CommandText $commandBox.Text.Trim()
    if (-not $validation.IsValid) {
        [System.Windows.Forms.MessageBox]::Show($validation.Message, "测试失败")
        return
    }

    $targetPath = Get-TestTargetPath -ScopeName $selectedScope
    $argumentText = Expand-CommandArguments -Arguments $validation.Arguments -TargetPath $targetPath

    try {
        Start-Process -FilePath $validation.Executable -ArgumentList $argumentText | Out-Null
        [System.Windows.Forms.MessageBox]::Show("测试命令已启动。`r`n目标路径：$targetPath", "测试成功")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "测试失败")
    }
})

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -eq 0) {
        Update-EditorState
        return
    }

    $row = $grid.SelectedRows[0]
    $keyBox.Text = [string]$row.Cells["KeyName"].Value
    $menuBox.Text = [string]$row.Cells["MenuText"].Value
    $iconBox.Text = [string]$row.Cells["Icon"].Value
    $commandBox.Text = [string]$row.Cells["Command"].Value
    Update-EditorState
})

$addSaveButton.Add_Click({
    $selectedScope = [string]$scopeBox.SelectedItem
    $registryPath = $writeScopeMap[$selectedScope]
    $draft = Get-EditorDraft
    if ($null -eq $draft) {
        return
    }

    $targetPath = Join-Path $registryPath $draft.KeyName
    if (Test-Path -LiteralPath $targetPath) {
        [System.Windows.Forms.MessageBox]::Show('当前键名已存在。请修改键名，或改用"更新保存"。', "新增失败")
        return
    }

    Save-MenuItem -RegistryRoot $registryPath -KeyName $draft.KeyName -MenuText $draft.MenuText -Command $draft.CommandText -Icon $draft.IconText
    Refresh-ShellMenuCache
    Load-Grid
    Select-GridRowByKeyName -KeyName $draft.KeyName
    [System.Windows.Forms.MessageBox]::Show("已新增，并已刷新资源管理器右键菜单缓存。")
})

$updateSaveButton.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先在列表中选中要更新的菜单项。")
        return
    }

    $selectedScope = [string]$scopeBox.SelectedItem
    $draft = Get-EditorDraft
    if ($null -eq $draft) {
        return
    }

    $selectedRow = $grid.SelectedRows[0]
    $originalKeyName = [string]$selectedRow.Cells["KeyName"].Value
    $userItemPath = Get-UserWritableItemPath -ScopeName $selectedScope -KeyName $originalKeyName

    if ([string]::IsNullOrWhiteSpace($userItemPath)) {
        [System.Windows.Forms.MessageBox]::Show('当前选中项不是当前用户创建的菜单项，不能直接更新。请使用"新增保存"创建自己的菜单项。', "更新失败")
        return
    }

    if ($draft.KeyName -ne $originalKeyName) {
        [System.Windows.Forms.MessageBox]::Show('更新保存不支持修改注册表键名。请保持原键名，或使用"新增保存"创建新项。', "更新失败")
        return
    }

    Save-MenuItem -RegistryRoot $writeScopeMap[$selectedScope] -KeyName $draft.KeyName -MenuText $draft.MenuText -Command $draft.CommandText -Icon $draft.IconText
    Refresh-ShellMenuCache
    Load-Grid
    Select-GridRowByKeyName -KeyName $draft.KeyName
    [System.Windows.Forms.MessageBox]::Show("已更新，并已刷新资源管理器右键菜单缓存。")
})

$deleteButton.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择一个菜单项。")
        return
    }

    $selectedRow = $grid.SelectedRows[0]
    $selectedScope = [string]$scopeBox.SelectedItem
    $keyName = [string]$selectedRow.Cells["KeyName"].Value
    $userItemPath = Get-UserWritableItemPath -ScopeName $selectedScope -KeyName $keyName

    if ([string]::IsNullOrWhiteSpace($userItemPath)) {
        [System.Windows.Forms.MessageBox]::Show('当前选中项不是当前用户创建的菜单项，不能直接删除。', "删除失败")
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "确定删除菜单项 '$keyName' 吗？",
        "确认删除",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Remove-MenuItem -ItemPath $userItemPath
    Refresh-ShellMenuCache
    Clear-Editor
    Load-Grid
})

$toggleButton.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("请先选择一个菜单项。")
        return
    }

    $selectedRow = $grid.SelectedRows[0]
    $selectedScope = [string]$scopeBox.SelectedItem
    $keyName = [string]$selectedRow.Cells["KeyName"].Value
    $userItemPath = Get-UserWritableItemPath -ScopeName $selectedScope -KeyName $keyName

    if ([string]::IsNullOrWhiteSpace($userItemPath)) {
        [System.Windows.Forms.MessageBox]::Show('当前选中项不是当前用户创建的菜单项，不能直接启用或禁用。', "操作失败")
        return
    }

    $enabled = [string]$selectedRow.Cells["Enabled"].Value -eq "是"
    Set-MenuItemEnabled -ItemPath $userItemPath -Enabled (-not $enabled)
    Refresh-ShellMenuCache
    Load-Grid

    foreach ($row in $grid.Rows) {
        if ($row.Cells["KeyName"].Value -eq $keyName) {
            $row.Selected = $true
            break
        }
    }
})

Load-Grid
[void]$form.ShowDialog()
