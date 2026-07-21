@echo off
setlocal
set "TARGET_HTML=C:\Users\Administrator\Desktop\ai分析\ai分析\skt-codex-analyst\output\thai_april_review_management_v2.html"
set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
if exist "%CHROME_EXE%" (
  start "" "%CHROME_EXE%" "%TARGET_HTML%"
  exit /b 0
)
set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if exist "%CHROME_EXE%" (
  start "" "%CHROME_EXE%" "%TARGET_HTML%"
  exit /b 0
)
echo Chrome was not found in the standard install path.
echo Falling back to the system default browser for this file.
start "" "%TARGET_HTML%"
