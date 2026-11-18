@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\build_local.ps1" %*
set "flotekaExit=%ERRORLEVEL%"
if not "%flotekaExit%"=="0" echo Budowa nie powiodla sie. Sprawdz komunikat i build.log w katalogu wyniku.
if "%~1"=="" if not "%FLOTEKA_NO_PAUSE%"=="1" pause
exit /b %flotekaExit%
