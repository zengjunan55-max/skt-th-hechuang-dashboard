@echo off
setlocal
cd /d "%~dp0"

python skills\build_april_review.py --output-html output\april_review_interactive_clean.html --output-json output\april_review_interactive_clean.json
if errorlevel 1 exit /b %errorlevel%

python skills\publish_review_site.py --source-html output\april_review_interactive_clean.html --source-json output\april_review_interactive_clean.json --site-dir site
if errorlevel 1 exit /b %errorlevel%

echo.
echo The fixed-link site files have been updated in:
echo   %CD%\site
echo.
echo If GitHub Pages is enabled for this repository, commit and push to refresh the shared link.
pause
