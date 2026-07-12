@echo off
setlocal
if exist "%~dp0dist\ExcelPrivacyCleanerNativeCsv\ExcelPrivacyCleanerNativeCsv.exe" (
    start "" "%~dp0dist\ExcelPrivacyCleanerNativeCsv\ExcelPrivacyCleanerNativeCsv.exe"
    exit /b 0
)
if exist "%~dp0dist\ExcelPrivacyCleanerNativeFixed\ExcelPrivacyCleanerNativeFixed.exe" (
    start "" "%~dp0dist\ExcelPrivacyCleanerNativeFixed\ExcelPrivacyCleanerNativeFixed.exe"
    exit /b 0
)
if exist "%~dp0dist\ExcelPrivacyCleanerNative\ExcelPrivacyCleanerNative.exe" (
    start "" "%~dp0dist\ExcelPrivacyCleanerNative\ExcelPrivacyCleanerNative.exe"
    exit /b 0
)
if exist "%~dp0dist\ExcelPrivacyCleaner\ExcelPrivacyCleaner.exe" (
    start "" "%~dp0dist\ExcelPrivacyCleaner\ExcelPrivacyCleaner.exe"
) else (
    call "%~dp0run_python_app.cmd"
)
