@echo off
chcp 65001 >nul
echo ===================================
echo  きりたんトーク LINE Bot 起動
echo ===================================
echo.

REM 既存プロセスを停止
echo [1/4] 既存プロセスを停止中...
taskkill /F /IM ngrok.exe >nul 2>&1
timeout /t 1 /nobreak >nul

REM サーバー起動
echo [2/4] LINE Bot サーバーを起動中...
start "KiritanBot" cmd /k "cd /d %~dp0 && python run_linebot.py"
timeout /t 3 /nobreak >nul

REM サーバー確認
echo [3/4] サーバー接続確認中...
curl -s http://localhost:5000/ >nul 2>&1
if %errorlevel%==0 (
    echo   → サーバー OK
) else (
    echo   → サーバー起動失敗。ログを確認してください。
    pause
    exit /b 1
)

REM ngrok起動
echo [4/4] ngrok トンネルを起動中...
start "ngrok" cmd /k "ngrok http 5000 --url=epigenous-anne-synonymical.ngrok-free.dev"
timeout /t 5 /nobreak >nul

REM URL取得
echo.
echo ===================================
echo  起動完了！
echo ===================================
echo.
echo  サーバー: http://localhost:5000/
echo.
echo  Webhook URL (固定):
echo  https://epigenous-anne-synonymical.ngrok-free.dev/callback
echo.
echo  ngrok の Web UI: http://localhost:4040
echo.
pause
