$ErrorActionPreference = "Stop"

& "$PSScriptRoot\build_exe.ps1"

$iscc = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
if (-not (Test-Path -LiteralPath $iscc)) {
    $iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
}
if (-not (Test-Path -LiteralPath $iscc)) {
    throw "ISCC.exe (Inno Setup) was not found. Install it with: winget install --id JRSoftware.InnoSetup -e"
}

& $iscc ".\installer\setup.iss"
if ($LASTEXITCODE -ne 0) {
    throw "ISCC failed with exit code $LASTEXITCODE."
}

Write-Host "Installer build complete: dist_installer\"
