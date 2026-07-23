@echo off
setlocal
if exist "%~dp0dist\hosoPrivacyCleaner\hosoPrivacyCleaner.exe" (
    start "" "%~dp0dist\hosoPrivacyCleaner\hosoPrivacyCleaner.exe"
    exit /b 0
)
call "%~dp0run_python_app.cmd"
