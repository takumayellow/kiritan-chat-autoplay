# LINE Bot と ngrok を動かし続ける。どちらかが落ちたら 10 秒後に起こし直す。
# タスクスケジューラ（KiritanLineBot、ログオン時）から wsl/run-hidden.vbs 経由で窓なしで起動する。
# 出力は logs/runtime/ に残す。止めるときは stop_bot_supervised.ps1。
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$py = Join-Path $root '.venv\Scripts\python.exe'
$ngrok = (Get-Command ngrok.exe).Source
$domain = 'epigenous-anne-synonymical.ngrok-free.dev'
$logDir = Join-Path $root 'logs\runtime'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$env:PYTHONUNBUFFERED = '1'
$env:HOST = '127.0.0.1'                 # ngrok は localhost へ転送するので、LAN には開けない

function Write-Log([string]$msg) {
    Add-Content -Path (Join-Path $logDir 'supervisor.log') -Value "$(Get-Date -Format s) $msg" -Encoding UTF8
}

function Start-Child([string]$name, [string]$exe, [string[]]$argv) {
    Write-Log "start $name"
    Start-Process -FilePath $exe -ArgumentList $argv -WorkingDirectory $root -NoNewWindow -PassThru `
        -RedirectStandardOutput (Join-Path $logDir "$name.out.log") `
        -RedirectStandardError (Join-Path $logDir "$name.err.log")
}

$children = @{}
while ($true) {
    if (-not $children.bot -or $children.bot.HasExited) {
        $children.bot = Start-Child 'bot' $py @('run_linebot.py')
        Start-Sleep -Seconds 3
    }
    if (-not $children.ngrok -or $children.ngrok.HasExited) {
        $children.ngrok = Start-Child 'ngrok' $ngrok @('http', '5000', "--url=$domain", '--log=stdout')
    }
    Start-Sleep -Seconds 10
}
