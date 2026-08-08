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
# Invoked as "python -m PyInstaller" rather than the .venv\Scripts\pyinstaller.exe
# launcher: that launcher fails silently (exit code 1, no output at all, even on
# --version) when the venv lives under a path containing non-ASCII characters,
# as this project's OneDrive path does ("ドキュメント").
& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name hosoPrivacyCleaner `
    --icon .\assets\app_icon.ico `
    --add-data "assets\app_icon.ico;assets" `
    --add-data "assets\app_icon.png;assets" `
    --add-data "resources\tessdata;resources\tessdata" `
    --collect-all ja_ginza `
    --collect-all spacy_legacy `
    --collect-all ginza `
    --paths .\src `
    .\native_main.py

if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller failed with exit code $LASTEXITCODE."
}

Write-Host "Build complete: dist\hosoPrivacyCleaner\hosoPrivacyCleaner.exe"
