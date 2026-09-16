@echo off
setlocal DisableDelayedExpansion
set "TG_ASSOC_SCRIPT=%~f0"
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -Command "$text = [System.IO.File]::ReadAllText($env:TG_ASSOC_SCRIPT, [System.Text.Encoding]::UTF8); & ([scriptblock]::Create(($text -split '(?m)^# POWERSHELL\r?$', 2)[1]))"
exit /b %errorlevel%

# POWERSHELL
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

try {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = '选择 Telegram.exe'
    $dialog.Filter = 'Telegram 程序 (Telegram.exe)|Telegram.exe'
    $dialog.CheckFileExists = $true

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        exit 0
    }

    $exe = $dialog.FileName
    $workdir = Split-Path -LiteralPath $exe
    $command = '"' + $exe + '" -workdir "' + $workdir + '" -- "%1"'

    foreach ($scheme in 'tg', 'tdesktop.tg') {
        $root = "HKCU:\Software\Classes\$scheme"

        New-Item -Path "$root\shell\open\command" -Force | Out-Null
        Set-ItemProperty -Path $root `
            -Name 'URL Protocol' `
            -Value '' `
            -Type String

        Set-ItemProperty -Path "$root\shell\open\command" `
            -Name '(default)' `
            -Value $command
    }

    [void][System.Windows.Forms.MessageBox]::Show(
        '关联完成', 'Telegram 链接关联',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
} catch {
    [void][System.Windows.Forms.MessageBox]::Show(
        $_.Exception.Message, '关联失败',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
