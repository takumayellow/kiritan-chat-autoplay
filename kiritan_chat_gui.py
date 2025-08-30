# -*- coding: utf-8 -*-
"""
kiritan_chat_gui.py
AssistantSeikaに依存せず、VOICEROID＋ 東北きりたん EX のGUIを直接操作して再生する版。
- OpenAI応答生成（失敗時はユーザ入力をそのまま返すフォールバック）
- 起動時＆再生後に「フレーズ編集」タブへ確実に復帰（select→invoke→click_input）
- 大きいテキストエリア( Edit )を自動検出して内容を差し替え
- 「再生」ボタンを検出してクリック（見つからない場合は F5 → Space をフォールバック送信）
- コンソールを最前面復帰（pywin32 が無い場合は黙ってスキップ）
"""

import os, sys, time, textwrap

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
        # オフライン/キー未設定でも動作確認できるように
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
def _desktop_uia():
    from pywinauto import Desktop
    return Desktop(backend="uia")

def _find_voiceroid_window(timeout=5.0):
    """VOICEROID＋ 東北きりたん EX を探す（フル幅プラスに注意）"""
    import re
    desk = _desktop_uia()
    title_re = r"^VOICEROID＋.*東北きりたん.*EX$"
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            win = desk.window(title_re=title_re)
            if win.exists(timeout=0.7):
                return win
        except Exception:
            pass
        time.sleep(0.3)
    return None

def ensure_phrase_tab(win, tries=3, interval=0.4):
    """タブを『フレーズ編集』へ戻す（select→invoke→click_input）"""
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
    """最大面積の Edit を本文エリアとみなして差し替える"""
    edits = [e for e in win.descendants(control_type="Edit") if e.is_enabled() and e.is_visible()]
    if not edits:
        return False

    def area(ctrl):
        r = ctrl.rectangle()
        return max(1, r.width() * r.height())

    edits.sort(key=area, reverse=True)

    # 1) APIあり（UIAで実装されていれば速い）
    for e in edits[:3]:
        try:
            e.set_edit_text(text)
            return True
        except Exception:
            pass

    # 2) フォールバック：キー送信
    try:
        from pywinauto.keyboard import send_keys
        win.set_focus()
        send_keys("^a{BACKSPACE}", pause=0.02)
        # 日本語と記号を含むので with_spaces + 低速化
        send_keys(text, with_spaces=True, pause=0.01)
        return True
    except Exception:
        return False

def click_play(win) -> bool:
    """『再生』らしきボタンを押す。無ければ F5→Space を順に送る"""
    # 1) ボタン名探索
    for b in win.descendants(control_type="Button"):
        name = (b.window_text() or "").strip()
        if ("再生" in name) or ("▶" in name) or ("Play" in name):
            try:
                b.click_input()
                return True
            except Exception:
                pass
    # 2) ショートカットのフォールバック
    try:
        from pywinauto.keyboard import send_keys
        win.set_focus()
        for key in ("{F5}", " "):
            send_keys(key, pause=0.02)
            time.sleep(0.05)
            return True
    except Exception:
        pass
    return False

def focus_console():
    """コンソールを最前面へ。失敗しても無視"""
    try:
        import win32gui
        from ctypes import windll
        hwnd = windll.kernel32.GetConsoleWindow()
        if hwnd:
            win32gui.ShowWindow(hwnd, 5)
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass

# ---------- main ----------
def main():
    print("【GUI版】VOICEROID を直接操作して読み上げ（AssistantSeika 非依存）")
    print("使い方: VOICEROID＋ 東北きりたん EX を起動してから、このスクリプトを実行。")
    print("コマンド: exit / quit （それ以外は会話）")

    # 起動時：VOICEROID検出 → フレーズ編集へ整列
    win = _find_voiceroid_window(timeout=8.0)
    if not win:
        print("VOICEROID＋ 東北きりたん EX のウィンドウが見つかりません。起動してから再実行してください。", file=sys.stderr)
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

        # 入力 → 再生 → タブ復帰 → コンソール復帰
        if not set_phrase_text(win, reply):
            print("本文エリアの検出/入力に失敗しました。VOICEROIDの画面レイアウトを確認してください。", file=sys.stderr)
        else:
            if not click_play(win):
                print("『再生』の実行に失敗しました。ショートカット（F5/Space）も効かない可能性があります。", file=sys.stderr)

        ensure_phrase_tab(win)
        focus_console()

    print("終了します。")

if __name__ == "__main__":
    main()
