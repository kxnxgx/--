@echo off
cd /d %~dp0\..\

echo ========================================
echo  FJALLRAVEN MD分析ツール V3 起動中...
echo ========================================
echo.

python v3\main.py

echo.
echo ========================================
echo  処理が完了しました。
echo  「v3」フォルダ内のExcelファイルを確認してください。
echo  ログは「v3\process_log.txt」にあります。
echo ========================================
pause
