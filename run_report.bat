@echo off
REM Daily Sales Report - one-shot refresh (for Task Scheduler or manual double-click).
REM 1) pull live sheets (read-only) if the service-account key is configured,
REM 2) build index.html + index.png + QC summary.
cd /d "%~dp0"

echo [%date% %time%] pulling live sheets (read-only)...
python -c "import gsheets; gsheets.fetch_snapshots()" || echo   live pull skipped (no creds) - using local files

echo [%date% %time%] building report...
python main.py

echo [%date% %time%] publishing to GitHub Pages...
call "%~dp0publish.bat"

echo [%date% %time%] done. Output: %~dp0output\index.png
