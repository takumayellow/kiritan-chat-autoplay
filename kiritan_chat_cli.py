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
import ctypes
import subprocess
import win32gui
import win32process
from typing import Optional, Tuple

# 音声入出力（必要なら使う）
try:
    import speech_recognition as sr
    import sounddevice as sd
except Exception:
    sr = None
    sd = None

# LLM
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

# UI 操作（UIA バックエンド）
from pywinauto import Application, timings


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


def ensure_phrase_tab():
    """
    VOICEROID のタブを『フレーズ編集』に合わせる。
    起動時と、再生のたびに呼ぶと安定。
    """
    hwnd, pid = find_voiceroid_handle()
    if not (hwnd and pid):
        print("⚠️ VOICEROID ウィンドウが見つかりません（タブ切替スキップ）")
        return

    win = connect_by_pid_hwnd(pid, hwnd)
    if not win:
        return

    # TabItem を総当たりで取得
    try:
        items = win.descendants(control_type='TabItem')
    except Exception as e:
        print(f"✖️ TabItem 取得失敗: {e}")
        return

    # 名前一致で select → invoke → click の順に試す
    for tab in items:
        if tab.element_info.name == 'フレーズ編集':
            try:
                tab.select()
                print("◎ 『フレーズ編集』 tab: select() 成功")
                return
            except Exception:
                pass
            try:
                tab.invoke()
                print("◎ 『フレーズ編集』 tab: invoke() 成功")
                return
            except Exception:
                pass
            try:
                tab.click_input()
                print("◎ 『フレーズ編集』 tab: click_input() 成功")
                return
            except Exception as e:
                print(f"✖️ 『フレーズ編集』 tab クリック失敗: {e}")
                return

    print("⚠️ 『フレーズ編集』タブが見つかりませんでした")


# ---------------- 音声再生（SeikaSay2 CLI） ----------------
def speak(text: str, speed: float = DEFAULT_SPEED):
    """
    SeikaSay2.exe -play で非同期起動→待機。
    再生後は PowerShell を前面に戻し、VOICEROID のタブを『フレーズ編集』へ戻す。
    """
    exe = seika_exe_path()
    cmd = [
        exe,
        "-cid",   str(CID_KIRITAN),
        "-speed", f"{float(speed):.2f}",
        "-play",
        "-nc",
        "-t", text,
    ]
    try:
        proc = subprocess.Popen(
            cmd,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        proc.wait()
    except KeyboardInterrupt:
        try:
            proc.terminate()
        except Exception:
            pass
        print("◆ 再生を中断しました。")
    finally:
        # タブを戻す（音声効果に飛ばされる対策）
        ensure_phrase_tab()
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


def chat_once(client, user_text: str) -> str:
    """
    利用可能そうなモデルを順に試す（環境によって異なるため）。
    OPENAI_MODEL が設定されていれば最優先。
    """
    tried = []
    models = []
    if os.getenv("OPENAI_MODEL"):
        models.append(os.getenv("OPENAI_MODEL"))
    models += ["gpt-5","gpt-4o-mini", "o4-mini-high", "o3-mini", "gpt-4o"]

    last_err = None
    for m in models:
        try:
            res = client.chat.completions.create(
                model=m,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_text},
                ],
            )
            return (res.choices[0].message.content or "").strip()
        except Exception as e:
            tried.append(m)
            last_err = e
    raise RuntimeError(f"使用可能なモデルが見つかりません（試行: {tried}）: {last_err}")


# ---------------- 入力ヘルパ（必要なら） ----------------
def listen_mic(limit: int) -> str:
    if not (sr and limit > 0):
        return ""
    r = sr.Recognizer()
    with sr.Microphone() as mic:
        print(f"[mic] 発話どうぞ（最大 {limit}s）…")
        audio = r.listen(mic, phrase_time_limit=limit)
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
    # 起動直後にタブを『フレーズ編集』へ
    ensure_phrase_tab()

    client = create_client()
    speed = DEFAULT_SPEED
    wait = DEFAULT_LISTEN
    mode = "dual"   # dual | text | mic | loop

    print("=== きりたんEX 会話 (CLI版) ===")
    print("mode dual/text/mic/loop | time N | speed X | exit")

    while True:
        try:
            # 入力
            if mode == "dual":
                user = input("You: ").strip()
            elif mode == "text":
                user = input("You (text): ").strip()
            elif mode == "mic":
                user = listen_mic(wait)
                if user:
                    print(f"You (mic): {user}")
                else:
                    continue
            elif mode == "loop":
                user = listen_loopback(wait)
                if user:
                    print(f"You (loop): {user}")
                else:
                    continue
            else:
                user = input("You: ").strip()

            if not user:
                continue

            low = user.lower()
            if low in ("exit", "quit"):
                break
            if low.startswith("mode "):
                v = low.split()[1]
                if v in ("dual", "text", "mic", "loop"):
                    mode = v
                    print(f"→ mode = {mode}")
                continue
            if low.startswith("time "):
                try:
                    wait = max(0, int(low.split()[1]))
                    print(f"→ listen = {wait}s")
                except Exception:
                    print("time N 形式")
                continue
            if low.startswith("speed "):
                try:
                    speed = float(low.split()[1])
                    # 安全範囲にクランプ
                    speed = max(0.5, min(4.0, speed))
                    print(f"→ speed = {speed}x")
                except Exception:
                    print("speed X 形式")
                continue

            # 生成→読み上げ
            reply = chat_once(client, user)
            print(f"きりたん: {reply}")
            speak(reply, speed)

            # mic モードは続けて一往復
            if mode == "mic":
                follow = listen_mic(wait)
                if follow:
                    print(f"You (mic): {follow}")
                    reply2 = chat_once(client, follow)
                    print(f"きりたん: {reply2}")
                    speak(reply2, speed)

        except KeyboardInterrupt:
            print("\n(CTRL+C) 中断。続けます。")
            bring_powershell_front()
            continue


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n致命的エラー: {e}", file=sys.stderr)
        bring_powershell_front()
        sys.exit(1)
