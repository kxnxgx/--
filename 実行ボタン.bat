@echo off
cd /d %~dp0

echo ========================================
echo FJALLRAVEN MD Analysis Tool V3
echo ========================================
echo.

python v3\main.py

echo.
echo ========================================
echo Process Finished.
echo Please check "MD Analysis Report v3" Excel file in the "v3" folder.
echo Logs are saved in "v3\process_log.txt".
echo ========================================
pause
