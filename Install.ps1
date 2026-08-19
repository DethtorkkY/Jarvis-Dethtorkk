# Install JARVIS dependencies, download the speech model if necessary, and add it to startup.
$ErrorActionPreference = 'Stop'
$python = Join-Path (Split-Path (Get-Command python).Source) 'python.exe'
if (-not (Test-Path -LiteralPath $python)) { $python = 'python' }

& $python -m pip install --upgrade -r (Join-Path $PSScriptRoot 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Python package installation failed.' }

$modelDirectory = Join-Path $PSScriptRoot 'models\vosk-model-small-ru-0.22'
if (-not (Test-Path -LiteralPath $modelDirectory -PathType Container)) {
    $modelsDirectory = Split-Path $modelDirectory -Parent
    $archive = Join-Path $env:TEMP 'vosk-model-small-ru-0.22.zip'
    New-Item -ItemType Directory -Path $modelsDirectory -Force | Out-Null
    Invoke-WebRequest 'https://alphacephei.com/vosk/models/vosk-model-small-ru-0.22.zip' -OutFile $archive
    Expand-Archive -LiteralPath $archive -DestinationPath $modelsDirectory -Force
    Remove-Item -LiteralPath $archive
}

& (Join-Path $PSScriptRoot 'Install-Autostart.ps1')
Write-Host 'JARVIS is installed. Sign out and back in to test autostart.'
