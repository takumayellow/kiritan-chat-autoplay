# -*- coding: utf-8 -*-
"""
GUI発展版-音声（Voice）
- VOICEROID＋ 東北きりたん EX を UIA で制御（AssistantSeika 不要）
- text/mic/loop の会話モードを /mode で切替
- 録音は sounddevice、文字起こしは OpenAI Whisper API
- 相槌モード（/aizuchi on）で短め＆相槌多めの返答スタイルに切替
"""

import os, sys, re, time, tempfile
from typing import Optional, List, Dict
from datetime import datetime

# === UIA / 入力操作 ===
from pywinauto import Desktop, keyboard
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.findwindows import ElementNotFoundError

# === 音声入出力 ===
import sounddevice as sd
import soundfile as sf

# === OpenAI ===
try:
    from openai import OpenAI
except Exception as e:
    print("openai パッケージがありません。`pip install openai` を実行してください。", file=sys.stderr)
    raise

# ====== 設定 ======
TITLE_RE = r"VOICEROID[＋+]\s*.*東北きりたん\s*EX(?:[^\S\r\n]*[\+\*/☆])?$"
PLAY_LABEL_RE = r"再生"
SAVE_LABEL_RE = r"音声保存"
PHRASE_TAB_LABEL = "フレーズ編集"

DEFAULT_MODELS = ["gpt-4o-mini", "o4-mini-high", "o3-mini", "gpt-4o"]
SYSTEM_PROMPT_BASE = "あなたは気さくで優しく、短めに素早く返答するアシスタントです。"
SYSTEM_PROMPT_AIZUCHI = (
    "あなたは聞き上手なアシスタントです。相手の話に相槌（うん、なるほど、たしかに等）を適度に交え、"
    "文は短め・わかりやすく・端的にまとめて返答してください。"
)
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)

# ====== 小物 ======
def _wrap(ctrl) -> BaseWrapper:
    if isinstance(ctrl, BaseWrapper): return ctrl
    if hasattr(ctrl, "wrapper_object"): return ctrl.wrapper_object()
    return ctrl

def find_voiceroid_window(timeout: float = 3.0) -> Optional[BaseWrapper]:
    end = time.time() + timeout
    while time.time() < end:
        wins = Desktop(backend="uia").windows(title_re=TITLE_RE, control_type="Window")
        if wins: return _wrap(wins[0])
        time.sleep(0.2)
    return None

def ensure_phrase_tab(win: BaseWrapper, timeout: float = 3.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        try:
            tabs = win.descendants(control_type="TabItem")
        except Exception:
            tabs = []
        for t in tabs:
            name = (t.window_text() or t.element_info.name or "").strip()
            if PHRASE_TAB_LABEL in name:
                w = _wrap(t)
                for action in (lambda: w.select(), lambda: w.invoke(), lambda: w.click_input()):
                    try:
                        action()
                        return True
                    except Exception:
                        pass
        time.sleep(0.2)
    return False

def _find_text_area(win: BaseWrapper):
    nodes = []
    try: nodes = win.descendants(control_type="Document")
    except Exception: pass
    if not nodes:
        try: nodes = win.descendants(control_type="Edit")
        except Exception: pass
    return [_wrap(n) for n in nodes]

def set_phrase_text(win: BaseWrapper, text: str) -> bool:
    for a in _find_text_area(win):
        try:
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
    try:
        btns = win.descendants(control_type="Button")
    except Exception:
        btns = []
    cand = []
    for b in btns:
        name = (b.window_text() or b.element_info.name or "").strip()
        if re.search(PLAY_LABEL_RE, name):
            cand.append(_wrap(b))
    cand = sorted(cand, key=lambda w: 0 if w.is_visible() else 1)
    for w in cand:
        try:
            w.click_input()
            return True
        except Exception:
            pass
    try:
        win.set_focus()
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
    try:
        btns = win.descendants(control_type="Button")
    except Exception:
        btns = []
    target = None
    for b in btns:
        name = (b.window_text() or b.element_info.name or "").strip()
        if re.search(SAVE_LABEL_RE, name):
            target = _wrap(b); break
    if not target: return False
    try:
        target.click_input()
    except Exception:
        return False

    end = time.time() + timeout
    dlg = None
    title_re = r"(名前を付けて保存|保存|Save As)"
    while time.time() < end:
        ds = Desktop(backend="uia").windows(title_re=title_re, control_type="Window")
        if ds: dlg = _wrap(ds[0]); break
        time.sleep(0.2)
    if not dlg: return False
    try:
        dlg.set_focus()
        keyboard.send_keys("^l")
        time.sleep(0.1)
        keyboard.send_keys(path, with_spaces=True)
        time.sleep(0.1)
        keyboard.send_keys("{ENTER}")
        return True
    except Exception:
        return False

# ====== OpenAI ======
def _choose_models() -> List[str]:
    env = (os.environ.get("OPENAI_MODEL") or "").strip()
    if env: return [env]
    return DEFAULT_MODELS[:]

def chat_once(messages: List[Dict[str, str]]) -> str:
    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    last_err = None
    for m in _choose_models():
        if not m: continue
        print(f"[model] {m}")
        try:
            # streaming が失敗したら non-stream へフォールバック
            try:
                stream = client.chat.completions.create(model=m, messages=messages, stream=True, temperature=0.7)
                buf = []
                print("assistant >", end="", flush=True)
                for chunk in stream:
                    delta = chunk.choices[0].delta.content or ""
                    if delta:
                        buf.append(delta); print(delta, end="", flush=True)
                print()
                return "".join(buf).strip()
            except Exception:
                resp = client.chat.completions.create(model=m, messages=messages, temperature=0.7)
                text = (resp.choices[0].message.content or "").strip()
                print("assistant >", text)
                return text
        except Exception as e:
            last_err = e
            print(f"[warn] {m} 失敗: {e}", file=sys.stderr)
    raise RuntimeError(f"全モデル失敗: {last_err}")

def transcribe_wav(path: str) -> str:
    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    # Whisper API（whisper-1）
    with open(path, "rb") as f:
        try:
            r = client.audio.transcriptions.create(model="whisper-1", file=f, response_format="text")
            # 新SDKは text 文字列を返すこともあるので両対応
            return (getattr(r, "text", None) or r or "").strip()
        except Exception as e:
            raise RuntimeError(f"transcribe 失敗: {e}")

def record_to_wav(seconds: float = 6.0, fs: int = 16000) -> str:
    print(f"[rec] 録音 {seconds:.1f}s ...（話しかけてください）")
    sd.default.samplerate = fs
    sd.default.channels = 1
    data = sd.rec(int(seconds * fs), dtype="float32")
    sd.wait()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".wav")
    sf.write(tmp.name, data, fs)
    print("[rec] 完了")
    return tmp.name

# ====== MAIN ======
def main():
    print("[ GUI発展版-音声 ] VOICEROID を直接操作して音声会話（AssistantSeika 不要）")
    print("先に VOICEROID＋ 東北きりたん EX を起動してください。")
    print("コマンド: exit / quit")
    print("補助: /mode text|mic|loop, /time N, /aizuchi on|off, /reset /reload /retry /paste /clear /save <path> /sys <txt>")
    print()

    mode = "text"       # text / mic / loop
    rec_seconds = 6.0
    aizuchi = False

    system_prompt = SYSTEM_PROMPT_BASE
    history: List[Dict[str, str]] = [{"role": "system", "content": system_prompt}]
    last_reply: Optional[str] = None

    while True:
        # 毎ループでウィンドウを再取得して安定化
        win = find_voiceroid_window(timeout=3.0)
        if not win:
            print("VOICEROID が見つかりません。起動してから実行してください。")
            return
        ensure_phrase_tab(win)

        if mode == "text":
            user = input("あなた > ").strip()
            if not user: continue
        else:
            # mic / loop は録音→transcribe
            wav = record_to_wav(rec_seconds)
            try:
                user = transcribe_wav(wav)
            finally:
                try: os.remove(wav)
                except Exception: pass
            print(f"you (ASR)> {user}")
            if not user:
                if mode == "mic":
                    continue
                else:
                    # loopでも空認識なら次へ
                    continue

        if user.lower() in ("exit", "quit"):
            print("終了します。"); return

        # ====== 補助コマンド ======
        if user.startswith("/"):
            cmd, *rest = user[1:].split(" ", 1)
            arg = rest[0].strip() if rest else ""

            if cmd == "mode":
                if arg in ("text","mic","loop"):
                    mode = arg; print(f"[mode] => {mode}")
                else:
                    print("使い方: /mode text|mic|loop")
                continue
            if cmd == "time":
                try:
                    rec_seconds = max(1.0, float(arg))
                    print(f"[time] 録音秒数: {rec_seconds}s")
                except Exception:
                    print("使い方: /time 6  （秒を指定）")
                continue
            if cmd == "aizuchi":
                on = arg.lower() in ("on","true","1")
                aizuchi = on
                system_prompt = SYSTEM_PROMPT_AIZUCHI if aizuchi else SYSTEM_PROMPT_BASE
                history = [{"role":"system","content":system_prompt}]
                print(f"[aizuchi] {'ON' if aizuchi else 'OFF'}（返答スタイル変更）")
                continue
            if cmd == "reset":
                history = [{"role":"system","content":system_prompt}]
                print("[reset] 履歴クリア"); continue
            if cmd == "reload":
                print("[reload] 次ループでウィンドウ再取得"); continue
            if cmd == "retry":
                if last_reply and set_phrase_text(win, last_reply) and click_play(win):
                    print("[retry] 貼り付け→再生 OK")
                else:
                    print("[retry] 失敗")
                continue
            if cmd == "paste":
                if last_reply and set_phrase_text(win, last_reply):
                    print("[paste] OK")
                else:
                    print("[paste] 失敗")
                continue
            if cmd == "clear":
                if set_phrase_text(win, ""):
                    print("[clear] OK")
                else:
                    print("[clear] 失敗")
                continue
            if cmd == "save":
                if not arg:
                    print("使い方: /save C:\\path\\to\\voice.wav"); continue
                if click_save_and_type_path(win, arg):
                    print(f"[save] {arg} に保存実行")
                else:
                    print("[save] 失敗（ボタン/ダイアログ未検出）")
                continue
            if cmd == "sys":
                if arg:
                    system_prompt = arg
                    history = [{"role":"system","content":system_prompt}]
                    print("[sys] 更新 & 履歴初期化")
                else:
                    print("使い方: /sys <新しいプロンプト>")
                continue

            print(f"[info] 未知のコマンド: /{cmd}")
            continue

        # ====== 通常会話 ======
        history.append({"role":"user","content":user})
        try:
            reply = chat_once(history[-12:])
        except Exception as e:
            print(f"[error] 生成失敗: {e}", file=sys.stderr)
            if mode == "loop": continue
            else:            continue

        last_reply = reply
        history.append({"role":"assistant","content":reply})

        if not set_phrase_text(win, reply):
            print("[paste] 入力欄に貼り付け失敗。", file=sys.stderr)
            if mode == "loop": continue
            else:            continue
        if not click_play(win):
            print("[play] 失敗（Space/F5 も不発）", file=sys.stderr)

        # ログ
        try:
            with open(os.path.join(LOG_DIR, f"{datetime.now():%Y-%m-%d}.txt"), "a", encoding="utf-8") as f:
                f.write(f"[user] {user}\n[assistant] {reply}\n---\n")
        except Exception:
            pass

        # loop のときはそのまま次の録音へ
        if mode == "loop":
            continue

if __name__ == "__main__":
    main()
