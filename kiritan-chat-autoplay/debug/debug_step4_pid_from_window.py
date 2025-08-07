# debug_step4_pid_from_window.py
import ctypes
from ctypes import wintypes
import win32gui
import psutil

# Win32 API 定義
GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId
GetWindowTextW            = ctypes.windll.user32.GetWindowTextW
GetWindowTextLengthW      = ctypes.windll.user32.GetWindowTextLengthW

def enum_cb(hwnd, result):
    length = GetWindowTextLengthW(hwnd)
    buff   = ctypes.create_unicode_buffer(length+1)
    GetWindowTextW(hwnd, buff, length+1)
    title = buff.value
    # タイトルに VOICEROID や SeikaSay2 を含むウィンドウだけ処理
    if 'VOICEROID' in title or 'SeikaSay2' in title:
        # PID を取る
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        result.append((hwnd, title, pid.value))
    return True

def main():
    windows = []
    win32gui.EnumWindows(enum_cb, windows)

    print("=== VOICEROID ウィンドウ → PID → 実ファイル path 一覧 ===")
    for hwnd, title, pid in windows:
        try:
            p = psutil.Process(pid)
            exe = p.exe()
            name= p.name()
        except Exception as e:
            exe = f"<error: {e}>"
            name= "<unknown>"
        print(f"hwnd={hwnd}  title={title!r}")
        print(f"   PID={pid}  name={name}  exe={exe}")
        print()

if __name__ == "__main__":
    main()
