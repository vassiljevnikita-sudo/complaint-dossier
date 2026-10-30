@echo off
rem Second command: generates a draft report for every complaint of the list that has none yet.
rem Double-click it or run it in a terminal: Run.cmd   (options: -Demo, -CaseId <no.>, -Regenerate, -From, -To)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-ComplaintDossier.ps1" %*
set CD_EXIT=%ERRORLEVEL%
rem Keep the window open when the file was double-clicked, so the result stays readable.
if not defined CI powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\Wait-IfDoubleClicked.ps1"
exit /b %CD_EXIT%
