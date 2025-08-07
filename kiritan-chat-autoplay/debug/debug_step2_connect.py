# debug_step2_connect_fixed2.py
from pywinauto import Application, timings

# 「VOICEROID+」→何文字でも→「東北きりたん」→何文字でも→「EX」
TITLE_RE = r"VOICEROID\+.*東北きりたん.*EX"

def try_connect(backend: str):
    try:
        app = timings.wait_until_passes(
            5, 0.5,
            lambda: Application(backend=backend)
                          .connect(title_re=TITLE_RE, visible_only=False)
        )
        print(f"[OK] backend={backend!r}")
    except Exception as e:
        print(f"[NG] backend={backend!r}: {type(e).__name__}")

if __name__ == "__main__":
    for bk in ("uia", "win32"):
        try_connect(bk)
