# Installs JARVIS only for the current Windows user.

$appName = 'JARVIS'
$source = Join-Path $PSScriptRoot 'jarvis.py'
$startup = [Environment]::GetFolderPath([Environment+SpecialFolder]::Startup)
$shortcutPath = Join-Path $startup "$appName.lnk"

if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    throw "JARVIS script was not found: $source"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$pythonw = Join-Path (Split-Path (Get-Command python).Source) 'pythonw.exe'
if (-not (Test-Path -LiteralPath $pythonw -PathType Leaf)) {
    throw "pythonw.exe was not found: $pythonw"
}
$shortcut.TargetPath = $pythonw
$shortcut.Arguments = "`"$source`""
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.WindowStyle = 7
$shortcut.Description = 'JARVIS voice greeting and wake-word listener at Windows sign-in'
$shortcut.Save()

Write-Host "JARVIS added to startup: $shortcutPath"
