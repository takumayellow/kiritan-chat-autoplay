@echo off
chcp 65001 >nul
echo ===================================
echo  きりたんチャット GUI 起動
echo ===================================
echo.

REM VOICEVOX エンジン起動
echo [1/2] VOICEVOX エンジンを起動中...
set VOICEVOX_PATH=%USERPROFILE%\Downloads\voicevox-cpu-0.25.1-x64.nsis\vv-engine\run.exe
if exist "%VOICEVOX_PATH%" (
    start "" "%VOICEVOX_PATH%" --host 127.0.0.1 --port 50021
    timeout /t 8 /nobreak >nul
    echo   → VOICEVOX 起動中...
) else (
    echo   → VOICEVOX が見つかりません: %VOICEVOX_PATH%
    echo     パスを確認してください。
)

REM GUI起動
echo [2/2] GUI アプリを起動中...
cd /d %~dp0
python run_gui.py

pause
