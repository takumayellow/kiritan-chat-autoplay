# -*- coding: utf-8 -*-
"""
kiritan_chat_gui.py  (AssistantSeika 非依存 / 直GUI操作)
- VOICEROID＋ 東北きりたん EX のウィンドウを UIA で掴み、テキストを流し込んで「再生」ボタンを押すだけ
- AssistantSeika の HTTP/WCF など一切使わない
- 毎ループごとにウィンドウを再取得してタブを「フレーズ編集」に戻す
- * がタイトルに付く/付かない、全角＋/半角+ の揺れに対応
"""

import sys
import time
import re
from typing import Optional

from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.controls.hwndwrapper import HwndWrapper
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.base_wrapper import BaseWrapper


# タイトルのゆらぎ（＋/+, EX の後ろに * or ＊ が付くなど）を許容
TITLE_RE = r"VOICEROID[＋+]\s*.*東北きりたん\s*EX(?:[\*＊])?$"

def _rewrap(ctrl):
    """ElementInfo/Wrapper のどちらで来ても Wrapper を返すヘルパ"""
    if isinstance(ctrl, BaseWrapper):
        return ctrl
    if hasattr(ctrl, "wrapper_object"):
        return ctrl.wrapper_object()
    return ctrl  # 最後の保険（基本ここには来ない）

def find_voiceroid_window(timeout: float = 3.0) -> Optional[BaseWrapper]:
    """VOICEROID のトップレベル Window を取る（なければ None）"""
    end = time.time() + timeout
    while time.time() < end:
        wins = Desktop(backend="uia").windows(title_re=TITLE_RE, control_type="Window")
        # 最も最近フォアグラウンドっぽいものを優先
        if wins:
            # 念のため rewrap
            return _rewrap(wins[0])
        time.sleep(0.2)
    return None

def ensure_phrase_tab(win: BaseWrapper) -> bool:
    """タブを『フレーズ編集』に戻す（select→invoke→click_input の順でフォールバック）"""
    try:
        tabitem = win.child_window(title_re=r"^\s*フレーズ編集\s*$", control_type="TabItem")
        tabitem = _rewrap(tabitem)
    except ElementNotFoundError:
        # タブコンテナから総当たり
        try:
            for t in win.descendants(control_type="TabItem"):
                if re.search(r"フレーズ編集", t.window_text() or ""):
                    tabitem = _rewrap(t)
                    break
            else:
                return False
        except Exception:
            return False

    # フォールバック順
    for op in ("select", "invoke", "click_input"):
        try:
            getattr(tabitem, op)()
            return True
        except Exception:
            continue
    return False

def set_phrase_text(win: BaseWrapper, text: str) -> bool:
    """本文エリア（Document or Edit）にテキストを流し込む"""
    area = None
    # Document 優先
    try:
        area = win.child_window(control_type="Document")
        area = _rewrap(area)
    except ElementNotFoundError:
        pass

    if area is None:
        # Edit でも探す
        try:
            # 「本文」相当だけ拾いたいのでサイズが大きめの Edit を優先
            edits = [ _rewrap(e) for e in win.descendants(control_type="Edit") ]
            area = max(edits, key=lambda e: (e.rectangle().width() * e.rectangle().height()), default=None)
        except Exception:
            area = None

    if area is None:
        return False

    try:
        area.set_focus()
        area.type_keys("^a{BACKSPACE}", set_foreground=True)
        # 日本語をそのまま送ると IME が効くので、直接 ValuePattern を使う
        try:
            vp = area.iface_value
            vp.SetValue(text)
        except Exception:
            # だめなら paste 系で
            import pyperclip  # 未インストールでも落ちないように局所 import
            try:
                pyperclip.copy(text)
                area.type_keys("^v", set_foreground=True)
            except Exception:
                # 最後の手段（ゆっくり直接タイプ）
                area.type_keys(text, with_spaces=True, pause=0.01, set_foreground=True)
        return True
    except Exception:
        return False

def click_play(win: BaseWrapper) -> bool:
    """『再生』ボタンを押す。名称の揺れ（例: '▶ 再生'）に強めの探索を行う"""
    try:
        # まずはタイトル完全一致狙い
        btn = win.child_window(title="再生", control_type="Button")
        btn = _rewrap(btn)
        btn.click_input()
        return True
    except ElementNotFoundError:
        pass
    except Exception:
        # 見つかったがクリック失敗 → 次の探索に回す
        pass

    # 総当たり（タイトル・AutomationId・アクセシブル名に '再生' を含むもの）
    try:
        for b in win.descendants(control_type="Button"):
            try:
                txt = (b.window_text() or "") + " " + (b.automation_id() or "")
            except Exception:
                txt = (b.window_text() or "")
            if "再生" in txt:
                _rewrap(b).click_input()
                return True
    except Exception:
        pass
    return False

def focus_console():
    """実行中のコンソールを手前に戻す（必須ではないけど操作感向上）"""
    try:
        # 実行中プロセスのコンソールを探す（ザックリ）
        for w in Desktop(backend="uia").windows(control_type="Window"):
            title = (w.window_text() or "")
            if "Windows PowerShell" in title or "PowerShell" in title or "cmd.exe" in title:
                _rewrap(w).set_focus()
                break
    except Exception:
        pass

def main():
    print("[ GUI版 ] VOICEROID を直接操作して読み上げ（AssistantSeika 非依存）\n")
    print("使い方：VOICEROID＋ 東北きりたん EX を起動してから、このスクリプトを実行。")
    print("コマンド:  exit / quit  （それ以外は会話）\n")

    while True:
        try:
            user = input("あなた> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n終了します。")
            return

        if not user:
            continue
        if user.lower() in ("exit", "quit"):
            print("終了します。")
            return

        # そのまま読み上げ（OpenAI 連携は意図的に外してます）
        reply = user

        # 毎回ウィンドウ取得
        win = find_voiceroid_window(timeout=2.0)
        if not win:
            print("VOICEROID＋ 東北きりたん EX のウィンドウが見つかりません。起動してから再実行してください。")
            continue

        # タブ復帰
        if not ensure_phrase_tab(win):
            print("タブ『フレーズ編集』の選択に失敗。ウィンドウのレイアウトをご確認ください。")
            continue

        # 本文セット
        if not set_phrase_text(win, reply):
            print("本文エリアの検出/入力に失敗しました。VOICEROIDの画面レイアウトを確認してください。")
            continue

        # 再生
        if not click_play(win):
            print("[!] 『再生』の実行に失敗（Space/F5 も不発）。ボタン名称や配置をご確認ください。")
            continue

        # 終了後にタブ復帰 + コンソール前面
        ensure_phrase_tab(win)
        focus_console()

if __name__ == "__main__":
    main()
