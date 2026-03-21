@echo off
chcp 65001 >nul
echo 全プロセスを停止中...
taskkill /F /IM ngrok.exe >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq KiritanBot" >nul 2>&1
echo 停止完了
pause
