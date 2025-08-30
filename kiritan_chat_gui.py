# -*- coding: utf-8 -*-
"""
kiritan_chat_gui.py
GUI優先（タブ復帰）＋CLI再生の分離版。
- 読み上げ: SeikaSay2.exe（環境変数 SEIKA_EXE で上書き可）
- キャラ:   KIRITAN_CID（未設定は 1707 をフォールバック）
- タブ復帰: 「フレーズ編集」に select→invoke→click_input の順で確実に戻す
"""

import os, sys, time, subprocess, textwrap

# ---- OpenAI ---------------------------------------------------------------
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

OPENAI_MODEL_FALLBACKS = [
    os.getenv("OPENAI_MODEL") or "",
    "gpt-4o-mini",
    "o4-mini-high",
    "o3-mini",
    "gpt-4o",
]

def chat_once(prompt: str) -> str:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key or OpenAI is None:
        # オフライン/キー未設定でも落とさない
        return prompt  # 入力をそのまま返す（動作確認用）
    client = OpenAI(api_key=api_key)
    last_err = None
    for m in [m for m in OPENAI_MODEL_FALLBACKS if m]:
        try:
            resp = client.chat.completions.create(
                model=m,
                messages=[{"role":"system","content":"あなたは簡潔に話す日本語アシスタントです。"},
                          {"role":"user","content":prompt}],
                temperature=0.6,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            last_err = e
    return f"（応答生成に失敗: {last_err}）"

# ---- Seika CLI ------------------------------------------------------------
def find_seika_exe() -> str:
    env = os.getenv("SEIKA_EXE")
    if env and os.path.isfile(env):
        return env
    # 代表的な置き場所を探索（見つかったら即返す）
    candidates = []
    home = os.path.expanduser("~")
    roots = [
        os.path.join(home, "Downloads"),
        r"C:\Program Files",
        r"C:\Program Files (x86)",
        os.getenv("LOCALAPPDATA") or "",
        os.getenv("ProgramData") or "",
    ]
    for r in [p for p in roots if p and os.path.isdir(p)]:
        for name in ("SeikaSay2N.exe","SeikaSay2.exe"):
            p = _search_file(r, name, depth=2)
            if p: candidates.append(p)
    # 優先: SeikaSay2N.exe（HTTP版） > SeikaSay2.exe（WCF）
    for name in ("SeikaSay2N.exe","SeikaSay2.exe"):
        for c in candidates:
            if c.endswith(name):
                return c
    # 最後の苦し紛れ（配布既定パス例）
    guess = os.path.join(home, r"Downloads\assistantseika20250113a\SeikaSay2\SeikaSay2.exe")
    return guess

def _search_file(root: str, filename: str, depth=2) -> str|None:
    try:
        for base, dirs, files in os.walk(root):
            if filename in files:
                return os.path.join(base, filename)
            # 深さ制限
            level = base[len(root):].count(os.sep)
            if level >= depth:
                dirs[:] = []
    except Exception:
        pass
    return None

def clamp_speed(x: float) -> float:
    try:
        v = float(x)
    except Exception:
        v = 1.0
    return max(0.5, min(4.0, v))

def speak_with_seika(text: str, cid: str, speed: float = 1.0) -> int:
    exe = find_seika_exe()
    speed = clamp_speed(speed)
    args = [exe, "-cid", str(cid), "-nc", "-speed", f"{speed:.2f}", "-t", text]
    try:
        p = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, encoding="utf-8")
        return p.returncode
    except FileNotFoundError:
        print("SeikaSay2.exe が見つかりません。環境変数 SEIKA_EXE を設定してください。", file=sys.stderr)
        return 127

# ---- GUI（タブ復帰） ------------------------------------------------------
def ensure_phrase_tab(retry=3, interval=0.5):
    try:
        from pywinauto import Desktop
    except Exception:
        print("pywinauto が未インストールのため、タブ復帰はスキップします。 `pip install pywinauto pywin32`", file=sys.stderr)
        return
    title_re = r"^VOICEROID＋.*東北きりたん.*EX$"
    for _ in range(retry):
        try:
            win = Desktop(backend="uia").window(title_re=title_re)
            if not win.exists(timeout=1.0):
                time.sleep(interval)
                continue
            # TabItem を総当たり
            phrase = None
            for t in win.descendants(control_type="TabItem"):
                name = (t.window_text() or "").strip()
                if "フレーズ編集" in name:
                    phrase = t; break
            if phrase is None:
                time.sleep(interval); continue
            # select → invoke → click_input の順で
            ok = False
            try:
                phrase.select(); ok = True
            except Exception:
                pass
            if not ok:
                try:
                    phrase.invoke(); ok = True
                except Exception:
                    pass
            if not ok:
                try:
                    phrase.click_input()
                    ok = True
                except Exception:
                    pass
            if ok:
                return
        except Exception:
            time.sleep(interval)
    # あきらめ
    return

def focus_console():
    # 失敗しても無視
    try:
        import win32gui, win32con
        from ctypes import windll
        hwnd = windll.kernel32.GetConsoleWindow()
        if hwnd:
            win32gui.ShowWindow(hwnd, 5)
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass

# ---- main -----------------------------------------------------------------
def main():
    cid = os.getenv("KIRITAN_CID", "").strip() or "1707"  # 既定フォールバック
    speed = 1.0
    print("【GUI版】きりたん Chat（タブ復帰＋CLI再生）")
    print("環境: KIRITAN_CID =", cid, " / SEIKA_EXE =", os.getenv("SEIKA_EXE") or "(auto)")
    print("コマンド: speed X / exit（それ以外は会話）")

    # 起動時にタブ復帰
    ensure_phrase_tab()
    focus_console()

    while True:
        try:
            user = input("あなた> ").strip()
        except (EOFError, KeyboardInterrupt):
            print(); break
        if not user:
            continue
        if user.lower().startswith("speed"):
            toks = user.split()
            if len(toks) >= 2:
                speed = clamp_speed(toks[1])
                print(f"→ speed = {speed:.2f}")
            continue
        if user.lower() in ("exit","quit","bye"):
            break

        reply = chat_once(user)
        print("きりたん>", reply)
        # 読み上げ
        rc = speak_with_seika(reply, cid=cid, speed=speed)
        # 再生後にタブ復帰＋コンソール前面
        ensure_phrase_tab()
        focus_console()

    print("終了します。")

if __name__ == "__main__":
    main()
