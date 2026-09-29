@echo off
rem ダブルクリックでセットアップ（PowerShell の実行ポリシーに関係なく動くようにする）
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup.ps1" %*
echo.
pause
