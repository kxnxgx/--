@echo off
chcp 65001 > nul
setlocal
cd /d %~dp0\..\

echo ========================================
echo  FJALLRAVEN MD分析ツール v2 起動中...
echo ========================================
echo.
echo CSVファイルを確認しています...
echo.

python v2\create_md_db_v2.py

echo.
echo ========================================
echo  処理が完了しました。
echo  「v2」フォルダ内のExcelファイルを
echo  確認してください。
echo  ログは「v2\process_log.txt」にあります。
echo ========================================
pause
