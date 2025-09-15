# --- UTF-8 hardening (Windows/ConHost) ---------------------------------------
import os, sys, io, locale
os.environ.setdefault("PYTHONUTF8","1")
os.environ["PYTHONIOENCODING"] = "utf-8"
def _reconfig_streams():
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        else:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
        if hasattr(sys.stdin, "reconfigure"):
            sys.stdin.reconfigure(encoding="utf-8", errors="ignore")
    except Exception:
        pass
try:
    locale.setlocale(locale.LC_ALL, "")
except Exception:
    pass
# ----------------------------------------------------------------------------- 
# -*- coding: utf-8 -*-
"""
GUI版：VOICEROID＋ 東北きりたん EX のウィンドウを UIA で掴み、テキストを流し込んで「再生」ボタンを押すだけ
- AssistantSeika の HTTP/WCF などは一切使わない
- 毎ループごとにウィンドウを再取得してタブを「フレーズ編集」に戻す
- * がタイトルに付く/付かない、全角＋/半角+ の揺れに対応
- OpenAIキーがあれば簡易チャット応答、なければ入力そのまま読み上げ
"""

import os
import sys
import time
import re
from typing import Optional

from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.base_wrapper import BaseWrapper

# タイトルの揺らぎ（+/＋, EX の後ろに * が付くなど）を許容
TITLE_RE = r"VOICEROID[＋+].*東北きりたん\s*EX(?:\s*\*|\s*)$"

# -------- ユーティリティ --------

def _rewrap(ctrl):
    """ElementInfo/Wrapper のどちらで来ても Wrapper を返すヘルパ"""
    if isinstance(ctrl, BaseWrapper):
        return ctrl
    # UIAWrapper ならそのまま、ElementInfo ならラップ
    try:
        return ctrl.wrapper_object()
    except Exception:
        return ctrl  # 最後の保険（ここに来ない想定）

def find_voiceroid_window(timeout: float = 3.0) -> Optional[BaseWrapper]:
    """VOICEROID+ 東北きりたん EX のトップレベル Window を取る（なければ None）"""
    end = time.time() + timeout
    while time.time() < end:
        wins = Desktop(backend="uia").windows(title_re=TITLE_RE, control_type="Window")
        # 直近にアクティブっぽいものを優先（先頭）
        if wins:
            return _rewrap(wins[0])
        time.sleep(0.2)
    return None

def ensure_phrase_tab(win: BaseWrapper, quiet: bool = False) -> bool:
    """タブを『フレーズ編集』に確実に戻す（select→invoke→click の順でフォールバック）"""
    tab = None
    # 1) まず名前一致（ * 付きも許容 ）
    for cand in win.descendants(control_type="TabItem"):
        name = (cand.element_info.name or "").strip()
        if re.fullmatch(r"フレーズ編集\*?", name):
            tab = cand
            break
    if not tab:
        # 2) ぼやっと 'フレーズ' を含むタブも許容
        for cand in win.descendants(control_type="TabItem"):
            name = (cand.element_info.name or "").strip()
            if "フレーズ" in name:
                tab = cand
                break
    if not tab:
        if not quiet:
            print("タブ『フレーズ編集』が見つかりません。VOICEROID の画面レイアウトを確認してください。", file=sys.stderr)
        return False

    w = _rewrap(tab)
    for op in ("select", "invoke", "click_input"):
        try:
            getattr(w, op)()
            time.sleep(0.1)
            return True
        except Exception:
            continue
    if not quiet:
        print("タブ切替に失敗しました（select/invoke/click_input 全滅）。", file=sys.stderr)
    return False

def _area(ctl) -> int:
    r = ctl.rectangle()
    return max(0, (r.right - r.left) * (r.bottom - r.top))

def find_main_edit(win: BaseWrapper) -> Optional[BaseWrapper]:
    """本文 Edit（入力エリア）を推定：一番大きい Edit を採用"""
    edits = win.descendants(control_type="Edit")
    if not edits:
        return None
    edits = sorted(edits, key=_area, reverse=True)
    return _rewrap(edits[0])

def set_phrase_text(win: BaseWrapper, text: str) -> bool:
    """本文エリアへテキストを書き込む（set_edit_text → Ctrl+A→type の順で試す）"""
    edit = find_main_edit(win)
    if not edit:
        print("本文エリア(Edit)が見つかりません。", file=sys.stderr)
        return False
    try:
        edit.set_focus()
    except Exception:
        try:
            edit.click_input()
        except Exception:
            pass

    # set_edit_text が最も安定
    try:
        edit.set_edit_text(text)
        return True
    except Exception:
        pass
    # フォールバック：全選択→削除→入力
    try:
        edit.type_keys("^a{BACKSPACE}", set_foreground=True)
        # with_spaces=True でスペースもそのまま
        edit.type_keys(text, with_spaces=True, set_foreground=True)
        return True
    except Exception:
        print("本文エリアへの書き込みに失敗しました。", file=sys.stderr)
        return False

def click_play(win: BaseWrapper) -> bool:
    """『再生』ボタンを押す。見つからなければ F5/Space へフォールバック"""
    # 1) Button 群から名前一致
    btn = None
    for b in win.descendants(control_type="Button"):
        name = (b.element_info.name or "").strip()
        # 例：'再生', '▶ 再生', '再生(P)'
        if re.search(r"(▶\s*)?再生", name):
            btn = b
            break
    if btn:
        w = _rewrap(btn)
        for op in ("invoke", "click_input"):
            try:
                getattr(w, op)()
                return True
            except Exception:
                continue

    # 2) フォールバック：F5 → Space
    try:
        _rewrap(win).type_keys("{F5}")
        return True
    except Exception:
        pass
    try:
        _rewrap(win).type_keys(" ")
        return True
    except Exception:
        print("『再生』の実行に失敗しました（ボタン／F5／Space 全滅）。", file=sys.stderr)
        return False

def focus_console():
    """PowerShell／ターミナルを前面に戻す（失敗しても無視）"""
    try:
        # よく使うタイトルをざっくり網羅
        cand = None
        for w in Desktop(backend="uia").windows(control_type="Window"):
            name = (w.element_info.name or "")
            if ("PowerShell" in name) or ("Windows Terminal" in name) or ("ターミナル" in name):
                cand = w; break
        if cand:
            _rewrap(cand).set_focus()
    except Exception:
        pass

# -------- 生成（任意：OPENAI_API_KEY があれば簡易応答） --------
def gen_reply(user_text: str) -> str:
    """OPENAI_API_KEY があれば簡易チャット、なければユーザ入力をそのまま返す"""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        return user_text
    try:
        # 1.x SDK を想定
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        model = os.getenv("OPENAI_MODEL") or "gpt-4o-mini"
        sys_prompt = "あなたは東北きりたん風のシンプルで優しいアシスタントです。30〜60文字程度に収めて返答してください。"
        r = client.chat.completions.create(
            model=model,
            messages=[
                {"role":"system","content":sys_prompt},
                {"role":"user","content":user_text}
            ],
            temperature=0.6,
        )
        content = r.choices[0].message.content
        return content or user_text
    except Exception as e:
        # 失敗時はそのまま読み上げ（ログだけ出す）
        print(f"[OpenAI error] {e}", file=sys.stderr)
        return user_text

# -------- メイン --------
def main():
    print("[ GUI版 ] VOICEROID を直接操作して読み上げ（AssistantSeika 非依存）")
    print("使い方：VOICEROID＋ 東北きりたん EX を起動してから、このスクリプトを実行。")
    print("コマンド: exit / quit （それ以外は会話）\n")

    while True:
        user = input("あなた > ").strip()
        if not user:
            continue
        if user.lower() in ("exit", "quit"):
            print("終了します。")
            return

        # 1) VOICEROID ウィンドウを再取得
        win = find_voiceroid_window(timeout=2.5)
        if not win:
            print("VOICEROID＋ 東北きりたん EX のウィンドウが見つかりません。起動してから再実行してください。", file=sys.stderr)
            continue

        # 2) タブを『フレーズ編集』へ
        if not ensure_phrase_tab(win, quiet=True):
            # タブ切替に失敗しても一応続行（環境によっては F5 で鳴ることがある）
            print("※警告: 『フレーズ編集』タブへの復帰に失敗", file=sys.stderr)

        # 3) 返答を用意（APIキーがあれば簡易応答、なければエコー）
        reply = gen_reply(user)
        print(f'きりたん > {reply}')
# 4) 本文に流し込み
        if not set_phrase_text(win, reply):
            print("本文エリアの検出/入力に失敗しました。VOICEROID の画面レイアウトを確認してください。", file=sys.stderr)
            continue

        # 5) 再生
        if not click_play(win):
            # 最後にもう一度タブを整えておく
            ensure_phrase_tab(win, quiet=True)
            print("『再生』の実行に失敗しました。ショートカット（F5/Space）も効かない可能性があります。", file=sys.stderr)
            continue

        # 6) 再生後、タブ整列→PowerShell 前面へ
        ensure_phrase_tab(win, quiet=True)
        focus_console()

if __name__ == "__main__":
    main()


