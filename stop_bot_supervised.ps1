# run_bot_supervised.ps1 で動かしている Bot・ngrok・見張り役をまとめて止める。
# タスクごと止めないと見張り役が起こし直すので、先にタスクを止める。
Stop-ScheduledTask -TaskName 'KiritanLineBot' -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process | Where-Object {
    ($_.Name -eq 'powershell.exe' -and $_.CommandLine -like '*run_bot_supervised.ps1*') -or
    ($_.Name -eq 'python.exe' -and $_.CommandLine -like '*run_linebot.py*') -or
    ($_.Name -eq 'ngrok.exe' -and $_.CommandLine -like '*epigenous-anne-synonymical*')
} | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
