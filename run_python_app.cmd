@echo off
setlocal
set "PYTHONPATH=%~dp0src"
where python >nul 2>nul
if errorlevel 1 (
    echo Python が見つかりません。Python 3.10 以上をインストールし、python コマンドを PATH に追加してください。
    pause
    exit /b 1
)
python "%~dp0native_main.py"
