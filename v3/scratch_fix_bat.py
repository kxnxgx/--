bat_content = """@echo off
cd /d %~dp0\\..\\

echo ========================================
echo  FJALLRAVEN MD Analysis Tool V3
echo ========================================
echo.

python v3\\main.py

echo.
echo ========================================
echo  Completed!
echo  Please check the Excel file in the 'v3' folder.
echo  Log is saved in 'v3\\process_log.txt'.
echo ========================================
pause
"""

with open(r"c:\分析\v3\実行ボタン.bat", "w", encoding="cp932") as f:
    f.write(bat_content)
