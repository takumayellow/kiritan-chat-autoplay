# -*- coding: utf-8 -*-
"""
kiritan_chat_gui.py  (robust)
- AssistantSeika 非依存。VOICEROID＋ 東北きりたん EX を GUI 操作して読み上げ
- ウィンドウ検出を強化（複数パターン + フォールバック走査 + wrapper_object 固定）
- クリック対象ボタンの探索を強化（名前/記号/部分一致 + F5/Space フォールバック）
- ループ毎にウィンドウを再取得（タイトル変化/再起動に追従）
"""

import os, sys, time
from typing import Optional

# ---------- OpenAI ----------
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
        return prompt
    client = OpenAI(api_key=api_key)
    last_err = None
    for model in [m for m in OPENAI_MODEL_FALLBACKS if m]:
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role":"system","content":"あなたは簡潔に話す日本語アシスタントです。"},
                    {"role":"user","content":prompt}
                ],
                temperature=0.6,
            )
            return (resp.choices[0].message.content or "").strip()
        except Exception as e:
            last_err = e
    return f"（応答生成に失敗: {last_err}）"

# ---------- GUI helpers ----------
from pywinauto import timings
timings.after_clickinput_wait = 0.05
timings.Timings.window_find_timeout = 5.0

def _desktop_uia():
    from pywinauto import Desktop
    return Desktop(backend="uia")

def _wrap(spec):
    """WindowSpecification -> wrapper_object (安定化)"""
    try:
        spec.wait("exists ready", timeout=1.5)
        return spec.wrapper_object()
    except Exception:
        return None

def _find_voiceroid_window(timeout: float = 8.0):
    """
    複数パターン＋全体走査で探す。
    タイトルは環境差が大きいので緩めに。
    """
    patterns = [
        r"^VOICEROID＋.*きりたん.*EX$",     # 正式（全角＋）
        r"^VOICEROID\+.*きりたん.*EX$",     # 半角+
        r"^VOICEROID.*きりたん.*EX$",       # ゆるめ
        r"^VOICEROID.*東北.*きりたん.*$",   # ゆるめ（EXなし）
        r"VOICEROID.*きりたん",             # 最後の砦
    ]
    desk = _desktop_uia()
    t0 = time.time()
    while time.time() - t0 < timeout:
        # 1) 正規表現パターンで試行
        for pat in patterns:
            try:
                spec = desk.window(title_re=pat, top_level_only=True)
                w = _wrap(spec)
                if w:
                    return w
            except Exception:
                pass
        # 2) 全トップを走査してタイトルで判定（フォールバック）
        try:
            for spec in desk.windows():
                title = (spec.window_text() or "").strip()
                if "VOICEROID" in title and ("きりたん" in title or "Kiritan" in title or "東北" in title):
                    w = _wrap(spec)
                    if w:
                        return w
        except Exception:
            pass
        time.sleep(0.3)
    return None

def ensure_phrase_tab(win, tries=3, interval=0.4):
    for _ in range(tries):
        try:
            target = None
            for t in win.descendants(control_type="TabItem"):
                name = (t.window_text() or "").strip()
                if "フレーズ編集" in name:
                    target = t; break
            if not target:
                time.sleep(interval); continue
            ok = False
            for action in ("select", "invoke", "click_input"):
                try:
                    getattr(target, action)()
                    ok = True
                    break
                except Exception:
                    pass
            if ok:
                return
        except Exception:
            pass
        time.sleep(interval)

def set_phrase_text(win, text: str) -> bool:
    edits = [e for e in win.descendants(control_type="Edit") if e.is_enabled() and e.is_visible()]
    if not edits:
        return False

    def area(ctrl):
        r = ctrl.rectangle()
        try:
            return max(1, r.width() * r.height())
        except Exception:
            return 1

    edits.sort(key=area, reverse=True)

    # 1) ValuePattern 系（速い）
    for e in edits[:4]:
        try:
            e.set_edit_text(text)
            return True
        except Exception:
            pass

    # 2) キー送信のフォールバック
    try:
        from pywinauto.keyboard import send_keys
        win.set_focus()
        send_keys("^a{BACKSPACE}", pause=0.02)
        send_keys(text, with_spaces=True, pause=0.01)
        return True
    except Exception:
        return False

def click_play(win) -> bool:
    """ボタン名いろいろ試す + F5 / Space フォールバック"""
    candidates = ("再生", "▶", "Play", "再生(F5)", "再生 / 停止")
    try:
        btns = win.descendants(control_type="Button")
    except Exception:
        btns = []
    for b in btns:
        try:
            name = (b.window_text() or "").strip()
        except Exception:
            name = ""
        if any(c in name for c in candidates):
            try:
                b.click_input()
                return True
            except Exception:
                pass
    # フォールバック: ショートカット送信
    try:
        from pywinauto.keyboard import send_keys
        win.set_focus()
        for key in ("{F5}", " "):
            send_keys(key, pause=0.03)
            time.sleep(0.06)
            return True
    except Exception:
        pass
    return False

def focus_console():
    try:
        import win32gui
        from ctypes import windll
        hwnd = windll.kernel32.GetConsoleWindow()
        if hwnd:
            win32gui.ShowWindow(hwnd, 5)
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass

def debug_print_top_windows():
    """見つからない時の調査用：トップレベルタイトルを出す"""
    try:
        desk = _desktop_uia()
        print("== Top windows (UIA) ==")
        for w in desk.windows():
            try:
                print(" -", (w.window_text() or "").strip())
            except Exception:
                pass
    except Exception:
        pass

# ---------- main ----------
def main():
    print("【GUI版(robust)】VOICEROID を直接操作して読み上げ（AssistantSeika 非依存）")
    print("使い方: VOICEROID＋ 東北きりたん EX を起動してから、このスクリプトを実行。")
    print("コマンド: exit / quit （それ以外は会話）")

    win = _find_voiceroid_window(timeout=8.0)
    if not win:
        print("VOICEROID のウィンドウが見つかりません。タイトルが異なる可能性があります。", file=sys.stderr)
        debug_print_top_windows()
        return
    ensure_phrase_tab(win)
    focus_console()

    while True:
        try:
            user = input("あなた> ").strip()
        except (EOFError, KeyboardInterrupt):
            print(); break
        if not user:
            continue
        if user.lower() in ("exit", "quit", "bye"):
            break

        reply = chat_once(user)
        print("きりたん>", reply)

        # ループ毎にウィンドウを再取得（タイトル変化/再起動に追従）
        w2 = _find_voiceroid_window(timeout=2.0) or win
        ensure_phrase_tab(w2)

        if not set_phrase_text(w2, reply):
            print("本文エリアの検出/入力に失敗しました。VOICEROIDの画面レイアウトを確認してください。", file=sys.stderr)
        else:
            if not click_play(w2):
                print("『再生』の実行に失敗しました。ショートカット（F5/Space）も効かない可能性があります。", file=sys.stderr)

        ensure_phrase_tab(w2)
        focus_console()

    print("終了します。")

if __name__ == "__main__":
    main()
