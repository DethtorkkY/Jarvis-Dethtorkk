# Removes only the JARVIS shortcut from the current user's startup folder.

$shortcut = Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::Startup)) 'JARVIS.lnk'
if (Test-Path -LiteralPath $shortcut -PathType Leaf) {
    Remove-Item -LiteralPath $shortcut
    Write-Host 'JARVIS removed from startup.'
} else {
    Write-Host 'JARVIS shortcut was not found in startup.'
}
