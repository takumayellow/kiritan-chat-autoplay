# debug_step2_connect_fixed3.py
from pywinauto import Application, timings

def try_connect(backend: str):
    try:
        # path 指定で SeikaSay2.exe 自体に接続を試みる
        app = Application(backend=backend).connect(path="SeikaSay2.exe")
        print(f"[OK] backend={backend!r} → SeikaSay2.exe に接続成功")
        return app
    except Exception as e:
        print(f"[NG] backend={backend!r}: {type(e).__name__}")
        return None

if __name__ == "__main__":
    for bk in ("uia","win32"):
        app = try_connect(bk)
        if app:
            break
