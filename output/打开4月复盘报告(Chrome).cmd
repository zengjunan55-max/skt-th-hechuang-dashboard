@echo off
setlocal
set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_EXE%" set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if exist "%CHROME_EXE%" (
  start "" "%CHROME_EXE%" "file:///C:/Users/Administrator/Documents/Playground/skt-codex-analyst/output/april_review_interactive.html"
  exit /b 0
)
echo Chrome was not found in the standard install path.
echo Falling back to the system default browser for this file.
start "" "C:\Users\Administrator\Documents\Playground\skt-codex-analyst\output\april_review_interactive.html"
