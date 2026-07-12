$ErrorActionPreference = "Stop"

$venvPython = ".\.venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $venvPython)) {
    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        throw "Python was not found. Install Python 3.10+ and add python.exe to PATH."
    }
    python -m venv .venv
}

& $venvPython -m pip install --disable-pip-version-check -r requirements.txt
& $venvPython .\tools\generate_app_icon.py
.\.venv\Scripts\pyinstaller.exe `
    --noconfirm `
    --clean `
    --windowed `
    --name ExcelPrivacyCleanerNativeCsv `
    --icon .\assets\app_icon.ico `
    --add-data "assets\app_icon.ico;assets" `
    --add-data "assets\app_icon.png;assets" `
    --paths .\src `
    .\native_main.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE."
}

Write-Host "Build complete: dist\ExcelPrivacyCleanerNativeCsv\ExcelPrivacyCleanerNativeCsv.exe"
