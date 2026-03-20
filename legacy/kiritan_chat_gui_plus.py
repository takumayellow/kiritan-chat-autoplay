# -*- coding: utf-8 -*-
"""
GUI発展版（GUI Plus）
- VOICEROID＋ 東北きりたん EX を UIA で掴み、テキストを流し込んで「再生」するだけ
- AssistantSeika / HTTP/WCF は使わない
- 毎ループでウィンドウを再取得 & 「フレーズ編集」タブに戻す
- ターミナルにも返答を逐次表示
- /reset /reload /retry /paste /clear /save /sys などの管理コマンド付き
注意: VOICEROID はユーザが先に起動しておくこと
"""

import os
import sys
import time
import re
from typing import Optional, List, Dict

from pywinauto import Desktop, keyboard
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.controls.hwndwrapper import HwndWrapper

# ==== OpenAI ====
try:
    from openai import OpenAI
except Exception as e:
    print("openai パッケージが見つかりません。`pip install openai` を実行してください。", file=sys.stderr)
    raise

# ---- 設定 ----
TITLE_RE = r"VOICEROID[＋+]\s*.*東北きりたん\s*EX(?:[^\S\r\n]*[\+\*/☆])?$"
PLAY_LABEL_RE = r"再生"
SAVE_LABEL_RE = r"音声保存"
PHRASE_TAB_LABEL = "フレーズ編集"

DEFAULT_MODELS = ["gpt-4o-mini", "o4-mini-high", "o3-mini", "gpt-4o"]
SYSTEM_PROMPT_DEFAULT = "あなたは気さくで、やさしく短めに返すアシスタントです。"

LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)

# ---- UTILS ----
def _rewrap(ctrl) -> BaseWrapper:
    """ElementInfo/Wrapper どちらでも Wrapper を返す"""
    if isinstance(ctrl, BaseWrapper):
        return ctrl
    if hasattr(ctrl, "wrapper_object"):
        return ctrl.wrapper_object()
    return ctrl

def find_voiceroid_window(timeout: float = 3.0) -> Optional[BaseWrapper]:
    """VOICEROID トップレベルウィンドウを取る（見つからなければ None）"""
    end = time.time() + timeout
    while time.time() < end:
        wins = Desktop(backend="uia").windows(title_re=TITLE_RE, control_type="Window")
        if wins:
            # 最も最近フォアグラウンドっぽいものを優先
            return _rewrap(wins[0])
        time.sleep(0.2)
    return None

def ensure_phrase_tab(win: BaseWrapper, timeout: float = 3.0) -> bool:
    """「フレーズ編集」タブへ復帰（select→invoke→click_input フォールバック）"""
    end = time.time() + timeout
    while time.time() < end:
        try:
            tabs = win.descendants(control_type="TabItem")
        except Exception:
            tabs = []
        for t in tabs:
            name = (t.window_text() or t.element_info.name or "").strip()
            if PHRASE_TAB_LABEL in name:
                w = _rewrap(t)
                try:
                    w.select()
                    return True
                except Exception:
                    pass
                try:
                    w.invoke()
                    return True
                except Exception:
                    pass
                try:
                    w.click_input()
                    return True
                except Exception:
                    pass
        time.sleep(0.2)
    return False

def _find_text_area(win: BaseWrapper):
    """入力欄（Document/Edit）候補を探す"""
    nodes = []
    try:
        nodes = win.descendants(control_type="Document")
    except Exception:
        pass
    if not nodes:
        try:
            nodes = win.descendants(control_type="Edit")
        except Exception:
            pass
    return [_rewrap(x) for x in nodes]

def set_phrase_text(win: BaseWrapper, text: str) -> bool:
    """入力欄に text を流し込む"""
    areas = _find_text_area(win)
    for a in areas:
        try:
            # UIA の edit wrapper は set_edit_text を持っていることが多い
            if hasattr(a, "set_edit_text"):
                a.set_edit_text(text)
            else:
                a.set_focus()
                keyboard.send_keys("^a{BACKSPACE}")
                time.sleep(0.05)
                keyboard.send_keys(text, with_spaces=True, pause=0.01)
            return True
        except Exception:
            continue
    return False

def click_play(win: BaseWrapper) -> bool:
    """「再生」ボタンを押す -> 失敗時 F5 → Space"""
    try:
        btns = win.descendants(control_type="Button")
    except Exception:
        btns = []
    # 名前に「再生」が含まれるボタンを可視優先で
    cand = []
    for b in btns:
        name = (b.window_text() or b.element_info.name or "").strip()
        if re.search(PLAY_LABEL_RE, name):
            cand.append(_rewrap(b))
    # 可視優先
    cand = sorted(cand, key=lambda w: 0 if w.is_visible() else 1)
    for w in cand:
        try:
            w.click_input()
            return True
        except Exception:
            continue
    # フォールバック
    try:
        win.set_focus()
    except Exception:
        pass
    try:
        keyboard.send_keys("{F5}")
        return True
    except Exception:
        pass
    try:
        keyboard.send_keys(" ")
        return True
    except Exception:
        pass
    return False

def click_save_and_type_path(win: BaseWrapper, path: str, timeout: float = 4.0) -> bool:
    """「音声保存」→ 保存ダイアログにパス入力 → Enter"""
    try:
        btns = win.descendants(control_type="Button")
    except Exception:
        btns = []
    target = None
    for b in btns:
        name = (b.window_text() or b.element_info.name or "").strip()
        if re.search(SAVE_LABEL_RE, name):
            target = _rewrap(b)
            break
    if not target:
        return False
    try:
        target.click_input()
    except Exception:
        return False

    # 保存ダイアログ（日本語/英語）にざっくり対応
    end = time.time() + timeout
    dlg = None
    title_re = r"(名前を付けて保存|保存|Save As)"
    while time.time() < end:
        ds = Desktop(backend="uia").windows(title_re=title_re, control_type="Window")
        if ds:
            dlg = _rewrap(ds[0]); break
        time.sleep(0.2)
    if not dlg:
        return False

    try:
        dlg.set_focus()
        keyboard.send_keys("^l")           # パス入力ボックスへ（エクスプローラー準拠）
        time.sleep(0.1)
        keyboard.send_keys(path, with_spaces=True)
        time.sleep(0.1)
        keyboard.send_keys("{ENTER}")
        return True
    except Exception:
        return False

# ==== OpenAI ====
def choose_model() -> str:
    env = os.environ.get("OPENAI_MODEL", "").strip()
    if env:
        return env
    return DEFAULT_MODELS[0]

def chat_once(messages: List[Dict[str, str]]) -> str:
    """モデル自動フォールバック付きで 1 回会話。逐次表示も行う。"""
    models = [os.environ.get("OPENAI_MODEL", "").strip()] if os.environ.get("OPENAI_MODEL") else DEFAULT_MODELS[:]
    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

    last_err = None
    for m in models:
        if not m: 
            continue
        print(f"[model] {m}")
        try:
            # 逐次表示: Chat Completions で chunk を受けつつ標準出力へ
            # （SDK によってストリーミング実装が変わるため、失敗したら通常モード）
            try:
                stream = client.chat.completions.create(
                    model=m, messages=messages, stream=True, temperature=0.7
                )
                buf = []
                print("assistant >", end="", flush=True)
                for chunk in stream:
                    delta = chunk.choices[0].delta.content or ""
                    if delta:
                        buf.append(delta)
                        print(delta, end="", flush=True)
                print()
                return "".join(buf).strip()
            except Exception:
                # 非ストリーミング
                resp = client.chat.completions.create(
                    model=m, messages=messages, temperature=0.7
                )
                text = (resp.choices[0].message.content or "").strip()
                print("assistant >", text)
                return text
        except Exception as e:
            last_err = e
            print(f"[warn] {m} 失敗: {e}", file=sys.stderr)
            continue
    raise RuntimeError(f"全モデルで失敗しました: {last_err}")

# ==== MAIN ====
def main():
    print("[ GUI発展版 ] VOICEROID を直接操作して読み上げ（AssistantSeika 非依存）")
    print("使い方: VOICEROID＋ 東北きりたん EX を起動してから、このスクリプトを実行。")
    print("コマンド: exit / quit（それ以外は会話）")
    print("補助コマンド: /reset /reload /retry /paste /clear /save <path> /sys <prompt>")
    print()

    # system プロンプト & 履歴
    system_prompt = os.environ.get("SYSTEM_PROMPT", SYSTEM_PROMPT_DEFAULT)
    history: List[Dict[str, str]] = [{"role": "system", "content": system_prompt}]
    last_reply: Optional[str] = None

    # 1回取得に失敗しても、毎ループで再取得する
    while True:
        win = find_voiceroid_window(timeout=3.0)
        if not win:
            print("VOICEROID＋ 東北きりたん EX のウィンドウが見つかりません。起動してから再実行してください。")
            return

        ensure_phrase_tab(win)

        user = input("あなた > ").strip()
        if not user:
            continue
        if user.lower() in ("exit", "quit"):
            print("終了します。")
            return

        # 補助コマンド
        if user.startswith("/"):
            cmd, *rest = user[1:].split(" ", 1)
            arg = rest[0].strip() if rest else ""

            if cmd == "reset":
                history = [{"role": "system", "content": system_prompt}]
                print("[reset] 履歴を消去しました。"); continue
            if cmd == "reload":
                print("[reload] ウィンドウを再取得…")
                continue  # 次ループで再取得
            if cmd == "retry":
                if last_reply:
                    if set_phrase_text(win, last_reply) and click_play(win):
                        print("[retry] 貼り付け→再生 OK")
                    else:
                        print("[retry] 実行に失敗。画面レイアウトを確認してください。", file=sys.stderr)
                else:
                    print("[retry] 直前の返答がありません。")
                continue
            if cmd == "paste":
                if last_reply and set_phrase_text(win, last_reply):
                    print("[paste] 貼り付けました。")
                else:
                    print("[paste] 失敗。")
                continue
            if cmd == "clear":
                if set_phrase_text(win, ""):
                    print("[clear] 入力欄をクリアしました。")
                else:
                    print("[clear] 失敗。")
                continue
            if cmd == "save":
                if not arg:
                    print("使い方: /save C:\\path\\to\\voice.wav"); continue
                if click_save_and_type_path(win, arg):
                    print(f"[save] {arg} に保存を試みました。")
                else:
                    print("[save] 失敗。音声保存ボタン/保存ダイアログが見つかりませんでした。", file=sys.stderr)
                continue
            if cmd == "sys":
                if arg:
                    system_prompt = arg
                    history = [{"role": "system", "content": system_prompt}]
                    print("[sys] system プロンプトを更新し、履歴を初期化しました。")
                else:
                    print("[sys] 使い方: /sys <新しいプロンプト>")
                continue

            print(f"[info] 未知のコマンドです: /{cmd}")
            continue

        # ---- 通常会話 ----
        history.append({"role": "user", "content": user})

        try:
            reply = chat_once(history[-12:])  # 直近だけ送ってトークン節約
        except Exception as e:
            print(f"[error] 生成に失敗: {e}", file=sys.stderr)
            continue

        last_reply = reply
        history.append({"role": "assistant", "content": reply})

        if not set_phrase_text(win, reply):
            print("本文エリアの検出/入力に失敗。VOICEROIDの画面レイアウトを確認してください。", file=sys.stderr)
            continue
        if not click_play(win):
            print("[ 再生 ] の実行に失敗（Space/F5 も不発）。", file=sys.stderr)
            continue

        # ログ
        try:
            from datetime import datetime
            with open(os.path.join(LOG_DIR, f"{datetime.now():%Y-%m-%d}.txt"), "a", encoding="utf-8") as f:
                f.write(f"[user] {user}\n[assistant] {reply}\n---\n")
        except Exception:
            pass

if __name__ == "__main__":
    main()
