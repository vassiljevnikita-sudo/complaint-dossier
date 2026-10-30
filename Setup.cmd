@echo off
rem First command: checks this computer, creates config.json and checks your Excel list and Word templates.
rem Double-click it or run it in a terminal: Setup.cmd   (options: -Reconfigure, -Demo)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup.ps1" %*
set CD_EXIT=%ERRORLEVEL%
rem Keep the window open when the file was double-clicked, so the result stays readable.
if not defined CI powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Wait-IfDoubleClicked.ps1"
exit /b %CD_EXIT%
