# debug_step3_find_pid.py
import psutil

print("=== SeikaSay2.exe プロセス一覧 ===")
for proc in psutil.process_iter(['pid','name','exe']):
    info = proc.info
    if info['name'] and 'SeikaSay2.exe' in info['name']:
        print(f"PID={info['pid']}, name={info['name']}, exe={info['exe']}")
