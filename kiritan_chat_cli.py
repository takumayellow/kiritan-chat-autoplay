# -*- coding: utf-8 -*-
"""
きりたん会話・自動読み上げ（VOICEROID+ 東北きりたん EX）
 - 会話生成: OpenAI API（モデルは環境に応じて自動選択）
 - 読み上げ: SeikaSay2.exe の CLI (-play)
 - 起動時と再生後に、VOICEROID のタブを「フレーズ編集」に自動で戻す
 - PowerShell のフォーカスが勝手に失われないよう前面復帰

必須:
  - OpenAI API キー: 環境変数 OPENAI_API_KEY
  - SeikaSay2.exe のパス: 既定値 or 環境変数 SEIKA_EXE で上書き可
"""

import os
import sys
import time
import io
import wave
import ctypes
import subprocess
import threading
import re
try:
    import msvcrt  # type: ignore
except Exception:
    msvcrt = None
import win32gui
import win32process
from typing import Dict, List, Optional, Tuple

# 音声入出力（必要なら使う）
try:
    import speech_recognition as sr
    import sounddevice as sd
except Exception:
    sr = None
    sd = None

try:
    import pyaudio  # type: ignore
except Exception:
    pyaudio = None

# LLM
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

# UI 操作（UIA バックエンド）
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys


MODEL_PROFILES: List[Dict[str, str]] = [
    {
        "key": "1",
        "alias": "light",
        "label": "ライト雑談（軽め）",
        "model": "gpt-4o-mini",
        "prompt": (
            "あなたは東北きりたんEXです。短めで気楽な雑談トーンを意識し、"
            "相手を励ますように明るく返答してください。"
        ),
    },
    {
        "key": "2",
        "alias": "normal",
        "label": "ゆったり会話（丁寧）",
        "model": "o4-mini",
        "prompt": (
            "あなたは東北きりたんEXです。落ち着いて相手の意図をくみ取り、"
            "適度な長さで丁寧に説明してください。"
        ),
    },
    {
        "key": "3",
        "alias": "deep",
        "label": "じっくり深掘り（教授モード）",
        "model": "o4-mini-high",
        "prompt": (
            "あなたは東北きりたんEXです。背景や理由を補足しながら、"
            "洞察や提案も添えて会話してください。"
        ),
    },
]

_LAST_VOICEROID_SPEED: Optional[float] = None

SANITIZE_TABLE = str.maketrans({
    "*": "",
    "`": "",
    "_": "",
    "~": "",
    "^": "",
    "#": "",
    "|": "",
})


def sanitize_for_voice(text: str) -> str:
    """
    読み上げ時に不要な装飾記号を除去する。
    """
    cleaned = text.translate(SANITIZE_TABLE)
    cleaned = re.sub(r"[•●○◆◇■□※☆★▶▷◀◁]", "・", cleaned)
    return cleaned


def focus_voiceroid_window():
    """VOICEROID ウィンドウを前面に持ってくる（フォーカス移動）"""
    hwnd, _ = find_voiceroid_handle()
    if not hwnd:
        return
    try:
        ctypes.windll.user32.ShowWindow(hwnd, 5)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def start_phrase_tab_sentry(interval: float = 1.2) -> Tuple[threading.Event, threading.Thread]:
    """フレーズ編集タブを見張る常駐スレッドを起動する。"""
    stop_event = threading.Event()

    def _worker():
        ensure_phrase_tab(log_failure=False)
        while not stop_event.wait(interval):
            ensure_phrase_tab(log_failure=False)

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    return stop_event, thread


def stop_phrase_tab_sentry(stop_event: Optional[threading.Event], thread: Optional[threading.Thread]) -> None:
    if stop_event:
        stop_event.set()
    if thread:
        try:
            thread.join(timeout=1.5)
        except KeyboardInterrupt:
            pass


def ensure_phrase_tab_with_retry(duration: float = 1.0, interval: float = 0.2) -> bool:
    """
    音声効果タブへの自動遷移対策として、短時間リトライしながら
    フレーズ編集タブへ戻す。
    """
    end = time.time() + max(duration, 0.1)
    success = False
    while time.time() < end:
        if ensure_phrase_tab(log_failure=False):
            success = True
        time.sleep(interval)
    if not success:
        ensure_phrase_tab(log_failure=True)
    return success


def _phrase_tab_active(win) -> bool:
    specs = [
        {"title": "フレーズ編集", "control_type": "TabItem"},
        {"title": "フレーズ編集", "control_type": "Button"},
        {"title": "フレーズ編集", "control_type": "RadioButton"},
    ]
    for spec in specs:
        try:
            ctrl = win.child_window(**spec).wrapper_object()
        except Exception:
            continue
        try:
            if hasattr(ctrl, "is_selected") and ctrl.is_selected():
                return True
        except Exception:
            pass
        try:
            if hasattr(ctrl, "get_toggle_state") and ctrl.get_toggle_state() == 1:
                return True
        except Exception:
            pass
    return False


def _sound_effect_active(win) -> bool:
    specs = [
        {"title": "音声効果", "control_type": "TabItem"},
        {"title": "音声効果", "control_type": "Button"},
        {"title": "音声効果", "control_type": "RadioButton"},
    ]
    for spec in specs:
        try:
            ctrl = win.child_window(**spec).wrapper_object()
        except Exception:
            continue
        try:
            if hasattr(ctrl, "is_selected") and ctrl.is_selected():
                return True
        except Exception:
            pass
        try:
            if hasattr(ctrl, "get_toggle_state") and ctrl.get_toggle_state() == 1:
                return True
        except Exception:
            pass
    return False


def _profile_lookup() -> Dict[str, Dict[str, str]]:
    lookup: Dict[str, Dict[str, str]] = {}
    for profile in MODEL_PROFILES:
        for key in (profile["key"], profile["alias"], profile["label"].lower()):
            lookup[key.lower()] = profile
    return lookup


def choose_conversation_profile() -> Dict[str, str]:
    lookup = _profile_lookup()
    print("\n=== 会話スタイルを選びましょう ===")
    print("きりたん、今日はどんなテンポでお話ししますか？")
    for p in MODEL_PROFILES:
        print(f"  {p['key']}: {p['label']}  (model: {p['model']})")
    print("  他にも 'light' / 'normal' / 'deep' のキーワードでも選べます。")

    default = MODEL_PROFILES[0]
    while True:
        bring_powershell_front()
        choice = input(f"スタイル番号 [Enter={default['key']}]: ").strip().lower()
        if not choice:
            print(f"→ {default['label']} を選択しました。")
            return default
        if choice in lookup:
            profile = lookup[choice]
            print(f"→ {profile['label']} を選択しました。")
            return profile
        print("  ※ 1/2/3 または light/normal/deep で選んでください。")

def print_cli_usage():
    print("\n--- コマンド一覧 ---")
    print("  help           : このヘルプ")
    print("  mode text      : 入力を文字入力に戻す")
    print("  mode mic       : マイク会話（'voice' でもOK、Enter で途中停止・待ち時間は time N）")
    print("  mode loop      : PCの再生音を拾って会話")
    print("  mode dual      : テキスト→マイクの連続モード")
    print("  time N         : 録音秒数の設定（mic/loop 共通）")
    print("  speed X        : 読み上げ速度 0.5～4.0")
    print("  style          : 会話スタイルを再選択")
    print("  exit           : 終了")
    print("\nヒント: mic モードでは自動で録音。待ち時間を変えたいときは time N を話す/入力してください。")


def normalize_mode_name(raw: str) -> Optional[str]:
    if not raw:
        return None
    key = raw.strip().lower()
    aliases = {
        "voice": "mic",
        "mic": "mic",
        "microphone": "mic",
        "text": "text",
        "loop": "loop",
        "dual": "dual",
    }
    return aliases.get(key)

    default = MODEL_PROFILES[0]
    while True:
        choice = input(f"スタイル番号 [Enter={default['key']}]: ").strip().lower()
        if not choice:
            print(f"→ {default['label']} を選択しました。")
            return default
        if choice in lookup:
            profile = lookup[choice]
            print(f"→ {profile['label']} を選択しました。")
            return profile
        print("  ※ 1/2/3 または light/normal/deep で選んでください。")
# ---------------- 設定 ----------------
CID_KIRITAN = 1707            # 東北きりたんEX CID
DEFAULT_SPEED = 1.0           # 読み上げ速度（Seika側の話速に対して倍率）
DEFAULT_LISTEN = 0            # mic/loop 時の秒数（使わない場合は 0 のまま）
VOICEROID_TITLE = 'VOICEROID＋ 東北きりたん EX'  # 全角プラス（＋）に注意

# SeikaSay2.exe の既定パス（必要なら SEIKA_EXE 環境変数で上書き）
DEFAULT_SEIKA_EXE = (
    r"C:\Users\takum\Downloads\assistantseika20250113a\SeikaSay2\SeikaSay2.exe"
)

SYSTEM_PROMPT = (
    "あなたは『東北きりたんEX』です。可愛らしく親しみやすい口調で、"
    "返答の最後に会話が続くような自然な一つの質問を添えてください。"
)


# ---------------- ユーティリティ ----------------
def seika_exe_path() -> str:
    p = os.getenv("SEIKA_EXE") or DEFAULT_SEIKA_EXE
    if not os.path.exists(p):
        raise FileNotFoundError(
            f"SeikaSay2.exe が見つかりません: {p}\n"
            "環境変数 SEIKA_EXE で正しいパスを指定してください。"
        )
    return p


def bring_powershell_front():
    """PowerShell を前面に戻す（フォーカス維持）"""
    def _cb(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if "PowerShell" in title:
            try:
                ctypes.windll.user32.SetForegroundWindow(hwnd)
            except Exception:
                pass
    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass


# ---------------- VOICEROID ウィンドウ検出＆接続 ----------------
def find_voiceroid_handle() -> Tuple[Optional[int], Optional[int]]:
    """VOICEROID ウィンドウの HWND と PID を返す"""
    hwnd = win32gui.FindWindow(None, VOICEROID_TITLE)
    if not hwnd:
        return None, None
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return hwnd, pid


def connect_by_pid_hwnd(pid: int, hwnd: int):
    """
    UIA バックエンドで PID/ハンドル指定でアタッチし WindowSpecification を返す。
    32bit/64bitの差異を気にせず TabItem を列挙可能。
    """
    try:
        app = timings.wait_until_passes(
            5, 0.5,
            lambda: Application(backend='uia').connect(process=pid, visible_only=False)
        )
        return app.window(handle=hwnd)
    except Exception as e:
        print(f"✖️ VOICEROID への接続失敗: {e}")
        return None


def ensure_phrase_tab(log_failure: bool = True) -> bool:
    """
    VOICEROID のタブを『フレーズ編集』に合わせる。
    起動時と、再生のたびに呼ぶと安定。
    """
    hwnd, pid = find_voiceroid_handle()
    if not (hwnd and pid):
        if log_failure:
            print("⚠️ VOICEROID ウィンドウが見つかりません（タブ切替スキップ）")
        return False

    win = connect_by_pid_hwnd(pid, hwnd)
    if not win:
        return False

    if _phrase_tab_active(win):
        return True

    control_types = ("TabItem", "Button", "RadioButton", "ToggleButton")
    controls = []
    for ctype in control_types:
        try:
            controls.extend(win.descendants(control_type=ctype))
        except Exception:
            continue

    last_err: Optional[Exception] = None
    for ctrl in controls:
        name = (getattr(ctrl, "window_text", lambda: "")() or ctrl.element_info.name or "").strip()
        if "フレーズ編集" not in name:
            continue
        try:
            wrapper = ctrl.wrapper_object()
        except Exception:
            wrapper = ctrl

        actions = []
        for attr in ("select", "invoke", "toggle"):
            if hasattr(wrapper, attr):
                actions.append(getattr(wrapper, attr))
        actions.append(lambda w=wrapper: w.click_input())

        for action in actions:
            try:
                action()
                time.sleep(0.05)
                if _phrase_tab_active(win):
                    return True
            except Exception as e:
                last_err = e
                continue

    if _sound_effect_active(win):
        try:
            btn = win.child_window(title="フレーズ編集", control_type="Button").wrapper_object()
            btn.click_input()
            time.sleep(0.05)
            if _phrase_tab_active(win):
                return True
        except Exception:
            pass
        try:
            win.set_focus()
            send_keys("^1")
            time.sleep(0.05)
            if _phrase_tab_active(win):
                return True
        except Exception:
            pass
    if log_failure:
        if last_err:
            print(f"✖️ 『フレーズ編集』 tab 操作失敗: {last_err}")
        else:
            print("⚠️ 『フレーズ編集』タブが見つかりませんでした")
    return False


def _guard_phrase_tab(stop_event: "threading.Event", interval: float = 0.6) -> None:
    """VOICEROID が音声再生中にタブがズレないよう見張る。"""
    ensure_phrase_tab(log_failure=False)
    miss = 0
    while not stop_event.wait(interval):
        if ensure_phrase_tab(log_failure=False):
            miss = 0
            continue
        miss += 1
        if miss >= 5:
            break


# ---------------- 音声再生（SeikaSay2 CLI） ----------------
def speak(text: str, speed: float = DEFAULT_SPEED):
    """
    SeikaSay2.exe -play で非同期起動→待機。
    再生後は PowerShell を前面に戻し、VOICEROID のタブを『フレーズ編集』へ戻す。
    """
    global _LAST_VOICEROID_SPEED
    exe = seika_exe_path()
    cmd = [
        exe,
        "-cid",   str(CID_KIRITAN),
    ]
    if _LAST_VOICEROID_SPEED is None or abs(float(speed) - _LAST_VOICEROID_SPEED) > 1e-3:
        cmd += ["-speed", f"{float(speed):.2f}"]
        _LAST_VOICEROID_SPEED = float(speed)
    cmd += [
        "-play",
        "-nc",
        "-t", text,
    ]
    guard_stop: Optional[threading.Event] = None
    guard_thread: Optional[threading.Thread] = None
    try:
        focus_voiceroid_window()
        ensure_phrase_tab(log_failure=False)
        proc = subprocess.Popen(
            cmd,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        guard_stop = threading.Event()
        guard_thread = threading.Thread(
            target=_guard_phrase_tab,
            args=(guard_stop,),
            daemon=True,
        )
        guard_thread.start()
        proc.wait()
    except KeyboardInterrupt:
        try:
            proc.terminate()
        except Exception:
            pass
        print("◆ 再生を中断しました。")
    finally:
        if guard_stop is not None:
            guard_stop.set()
        if guard_thread is not None:
            guard_thread.join(timeout=1.0)
        # タブを戻す（音声効果に飛ばされる対策）
        ensure_phrase_tab_with_retry(duration=1.2, interval=0.2)
        # PowerShell を前面に
        bring_powershell_front()


# ---------------- 会話生成（OpenAI） ----------------
def create_client():
    if OpenAI is None:
        raise RuntimeError("openai ライブラリが未インストールです。`pip install openai`")
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("環境変数 OPENAI_API_KEY が未設定です。")
    return OpenAI(api_key=key)




def _collect_candidate_models(preferred: str) -> List[str]:
    ordered: List[str] = []

    def _add(name: Optional[str]):
        if name and name not in ordered:
            ordered.append(name)

    _add(os.getenv("OPENAI_MODEL"))
    _add(preferred)
    for profile in MODEL_PROFILES:
        _add(profile["model"])
    for fallback in ("gpt-4o-mini", "o4-mini", "o4-mini-high", "o3-mini", "gpt-4o"):
        _add(fallback)
    return ordered

def chat_once(client, user_text: str, preferred_model: str, system_prompt: str) -> str:
    """選んだモデルを優先しつつ、順にフォールバックして応答を得る。"""
    models = _collect_candidate_models(preferred_model)
    tried = []
    last_err: Optional[Exception] = None
    prompt = system_prompt or SYSTEM_PROMPT

    for model_name in models:
        try:
            res = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": user_text},
                ],
            )
            raw = (res.choices[0].message.content or "").strip()
            return sanitize_for_voice(raw)
        except Exception as err:
            tried.append(model_name)
            last_err = err
    raise RuntimeError(f"利用可能なモデルが見つかりません (tried={tried}) : {last_err}")


def _record_with_pyaudio(seconds: int, rate: int = 16000) -> bytes:
    if not (pyaudio and seconds > 0):
        return b""
    pa = pyaudio.PyAudio()
    stream = None
    frames = []
    chunk = 1024
    total = max(1, int(rate / chunk * seconds))
    try:
        stream = pa.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=rate,
            input=True,
            frames_per_buffer=chunk,
        )
        print(f"[mic] 録音 {seconds}s ...")
        for _ in range(total):
            try:
                frames.append(stream.read(chunk, exception_on_overflow=False))
            except Exception as e:
                print(f"[mic] 取得エラー: {e}", file=sys.stderr)
                break
            if msvcrt and msvcrt.kbhit():
                ch = msvcrt.getwch()
                if ch == "\r":
                    print("[mic] Enter で録音を終了しました。")
                    # consume trailing LF if存在
                    while msvcrt.kbhit():
                        msvcrt.getwch()
                    break
    except Exception as e:
        print(f"[mic] PyAudio 初期化失敗: {e}", file=sys.stderr)
    finally:
        if stream:
            try:
                stream.stop_stream()
                stream.close()
            except Exception:
                pass
        pa.terminate()
    audio = b"".join(frames)
    if not audio:
        print("[mic] 音声を取得できませんでした。", file=sys.stderr)
    return audio


def _transcribe_with_openai(client, audio_bytes: bytes, rate: int = 16000) -> str:
    if not (client and audio_bytes):
        return ""
    buf = io.BytesIO()
    with wave.open(buf, "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(rate)
        wf.writeframes(audio_bytes)
    buf.seek(0)
    model = os.getenv("OPENAI_TRANSCRIBE_MODEL") or "gpt-4o-mini-transcribe"
    try:
        res = client.audio.transcriptions.create(
            model=model,
            file=("mic-input.wav", buf, "audio/wav"),
            response_format="text",
        )
    except Exception as e:
        print(f"[mic] 文字起こし失敗: {e}", file=sys.stderr)
        return ""
    if isinstance(res, str):
        return res.strip()
    text = getattr(res, "text", "")
    if isinstance(text, str):
        return text.strip()
    return ""


def listen_mic(client, limit: int) -> str:
    if limit <= 0:
        return ""
    if pyaudio and client:
        try:
            audio = _record_with_pyaudio(limit)
        except KeyboardInterrupt:
            print("\n[mic] 録音をキャンセルしました。")
            raise
        if audio:
            text = _transcribe_with_openai(client, audio)
            if text:
                return text
    if not sr:
        return ""
    r = sr.Recognizer()
    try:
        with sr.Microphone() as mic:
            print(f"[mic] 発話どうぞ（最大 {limit}s）…")
            audio = r.listen(mic, phrase_time_limit=limit)
    except KeyboardInterrupt:
        print("\n[mic] 録音をキャンセルしました。")
        raise
    try:
        return r.recognize_google(audio, language="ja-JP")
    except Exception:
        return ""


def listen_loopback(limit: int) -> str:
    if not (sd and limit > 0):
        return ""
    print(f"[loop] システム音声録音（{limit}s）…")
    rec = sd.rec(int(limit * 44100), samplerate=44100, channels=2)
    sd.wait()
    try:
        data = rec.tobytes()
        recog = sr.Recognizer()
        audio = sr.AudioData(data, 44100, 2)
        return recog.recognize_google(audio, language="ja-JP")
    except Exception:
        return ""


# ---------------- メイン ----------------
def main():
    ensure_phrase_tab()

    sentry_stop: Optional[threading.Event] = None
    sentry_thread: Optional[threading.Thread] = None

    client = create_client()
    profile = choose_conversation_profile()
    current_model = profile["model"]
    system_prompt = profile["prompt"]
    print(f"[model] {current_model} ({profile['label']})")
    print("Hint: 'mode mic'（または 'voice'）でマイク会話に切替。Enter で途中停止・録音秒数は time N。")

    try:
        sentry_stop, sentry_thread = start_phrase_tab_sentry()
    except Exception:
        sentry_stop = None
        sentry_thread = None

    speed = DEFAULT_SPEED
    wait = DEFAULT_LISTEN
    mode = "dual"
    print_cli_usage()

    try:
        while True:
            if mode == "dual":
                bring_powershell_front()
                user = input("You: " ).strip()
            elif mode == "text":
                bring_powershell_front()
                user = input("You (text): " ).strip()
            elif mode == "mic":
                if wait <= 0:
                    wait = 6
                    print(f"[mic] 録音秒数が未設定だったため {wait}s に設定しました（time N で変更）。")
                user = listen_mic(client, wait)
                if user:
                    print(f"You (mic): {user}")
                else:
                    print("[mic] 音声を認識できませんでした。")
                    continue
            elif mode == "loop":
                user = listen_loopback(wait)
                if user:
                    print(f"You (loop): {user}")
                else:
                    continue
            else:
                bring_powershell_front()
                user = input("You: " ).strip()

            if not user:
                continue

            low = user.lower()

            if low in ("exit", "quit"):
                break
            if low in ("help", "/help"):
                print_cli_usage()
                continue
            if low.startswith("style"):
                profile = choose_conversation_profile()
                current_model = profile["model"]
                system_prompt = profile["prompt"]
                print(f"[model] {current_model} ({profile['label']})")
                continue
            if low.startswith("mode "):
                parts = low.split()
                if len(parts) >= 2:
                    new_mode = normalize_mode_name(parts[1])
                    if new_mode:
                        mode = new_mode
                        print(f"[mode] -> {mode}")
                        if mode == "mic":
                            if wait <= 0:
                                wait = 6
                                print(f"  -> Mic conversation mode. 録音秒数を {wait}s に設定しました（time N で変更）。")
                            else:
                                print(f"  -> Mic conversation mode. 録音秒数は {wait}s（time N で変更）。")
                            print("     以降は自動で録音します。")
                        continue
                print("Usage: mode text|mic|dual|loop")
                continue
            quick_mode = normalize_mode_name(low)
            if quick_mode:
                mode = quick_mode
                print(f"[mode] -> {mode}")
                if mode == "mic":
                    if wait <= 0:
                        wait = 6
                        print(f"  -> Mic conversation mode. 録音秒数を {wait}s に設定しました（time N で変更）。")
                    else:
                        print(f"  -> Mic conversation mode. 録音秒数は {wait}s（time N で変更）。")
                    print("     以降は自動で録音します。")
                continue
            if low.startswith("time "):
                try:
                    wait = max(0, int(low.split()[1]))
                    print(f"[listen] {wait}s")
                except Exception:
                    print("Usage: time N")
                continue
            if low.startswith("speed "):
                try:
                    speed = float(low.split()[1])
                    speed = max(0.5, min(4.0, speed))
                    print(f"[speed] {speed}x")
                except Exception:
                    print("Usage: speed X")
                continue

            reply = chat_once(client, user, current_model, system_prompt)
            print(f"Kiritan: {reply}")
            speak(reply, speed)

            if mode == "mic":
                if wait <= 0:
                    wait = 6
                    print(f"[mic] 録音秒数が未設定だったため {wait}s に設定しました（time N で変更）。")
                follow = listen_mic(client, wait)
                if follow:
                    print(f"You (mic): {follow}")
                    reply2 = chat_once(client, follow, current_model, system_prompt)
                    print(f"Kiritan: {reply2}")
                    speak(reply2, speed)
    except KeyboardInterrupt:
        print("\n(CTRL+C) 終了します。")
        bring_powershell_front()
        return
    finally:
        stop_phrase_tab_sentry(sentry_stop, sentry_thread)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n致命的エラー: {e}", file=sys.stderr)
        bring_powershell_front()
        sys.exit(1)
