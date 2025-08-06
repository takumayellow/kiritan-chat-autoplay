# kiritan-chat-autoplay.py (CLI版・完全版)
"""
主な変更点
1. GUI 自動操作は不要 → SeikaSay2.exe の -play で直接再生
2. -play 呼び出しによる invalid option:-play エラーを解消
3. フレーズ編集タブ切り替えコードを削除
4. 再生完了後に PowerShell へフォーカスを戻す
"""

import os
import subprocess
import ctypes
import win32gui
import time
import speech_recognition as sr
import sounddevice as sd
from openai import OpenAI

# ---------------- 基本設定 ----------------
DEFAULT_SPEED  = 1.0   # 読み上げ速度
DEFAULT_LISTEN = 0     # mik/loop モードの聞き取り秒数
CID_KIRITAN    = 1707  # 東北きりたんEX の CID
SEIKA_CLI      = (
    r"C:\Users\takum\Downloads\assistantseika20250113a"
    r"\SeikaSay2\SeikaSay2.exe"
)

SYSTEM_PROMPT = (
    "あなたは「東北きりたんEX」です。"
    "声優・茜屋日海夏さんの柔らかく落ち着いた声質を踏まえ、可愛らしい口調で返答してください。"
    "フレーズごとに適度な話速と抑揚をつけ、必要に応じて「明るい」「デレ」「ダルい」「怒り」「泣き」などの感情表現を交えてください。"
    "文末は「〜だよ」「〜だね」「〜かな？」等で親しみやすく締めましょう。"
    "無理に「〜だよ」を付ける必要はありません。"
)

# ---------------- OpenAI クライアント ----------------
def create_client() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise ValueError("環境変数 OPENAI_API_KEY が未設定")
    return OpenAI(api_key=key)

def chat(client: OpenAI, prompt: str) -> str:
    res = client.chat.completions.create(
        model="o3-mini",
        messages=[
            {"role":"system","content":SYSTEM_PROMPT},
            {"role":"user",  "content":prompt}
        ]
    )
    return res.choices[0].message.content.strip()

# ---------------- 入力ハンドラ ----------------
def listen_mic(limit: int) -> str:
    r = sr.Recognizer()
    with sr.Microphone() as mic:
        print(f"[mic] 発話どうぞ（最大 {limit}s）…")
        audio = r.listen(mic, phrase_time_limit=limit)
    try:
        return r.recognize_google(audio, language="ja-JP")
    except:
        return ""

def listen_loopback(limit: int) -> str:
    print(f"[loop] システム音声録音（{limit}s）…")
    rec = sd.rec(int(limit*44100), samplerate=44100, channels=2)
    sd.wait()
    try:
        data = rec.tobytes()
        return sr.Recognizer().recognize_google(
            sr.AudioData(data, 44100, 2),
            language="ja-JP"
        )
    except:
        return ""

# ---------------- 補助関数 ----------------
def bring_powershell_front():
    def _cb(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if "PowerShell" in title:
            ctypes.windll.user32.SetForegroundWindow(hwnd)
    win32gui.EnumWindows(_cb, None)

# ---------------- 再生関数（CLI -play） ----------------
import subprocess

def speak(text: str, speed: float):
    """
    SeikaSay2.exe -play を非同期に起動し、終了まで待つ。
    CTRL+C などで中断しても PowerShell が終了しないようにする。
    """
    cmd = [
        SEIKA_CLI,
        "-cid",   str(CID_KIRITAN),
        "-speed", str(speed),
        "-play",              # 再生
        "-nc",                # コンソール出力抑制
        "-t",    text
    ]
    try:
        # CREATE_NO_WINDOW フラグを付けると新しいコンソールを開かない
        proc = subprocess.Popen(
            cmd,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        proc.wait()
    except KeyboardInterrupt:
        # 再生中に Ctrl+C で中断してもスクリプトを終了しない
        proc.terminate()
        print("◆ 音声再生を中断しました。続けて入力できます。")
    finally:
        # 再生終了後に PowerShell を前面へ
        bring_powershell_front()

# ---------------- メインループ ----------------
def main():
    client = create_client()
    speed  = DEFAULT_SPEED
    wait   = DEFAULT_LISTEN
    mode   = 'dual'  # dual | text | mic | loop

    print("=== きりたんEX 会話 (CLI版) ===")
    print("mode dual/text/mic/loop | time N | speed X | exit")

    while True:
        # ── 入力処理 ──
        if mode == 'dual':
            user = input("You: ").strip()
        elif mode == 'text':
            user = input("You (text): ").strip()
        elif mode == 'mic':
            user = listen_mic(wait)
            if user:
                print(f"You (mic): {user}")
        elif mode == 'loop':
            user = listen_loopback(wait)
            if user:
                print(f"You (loop): {user}")
        else:
            user = input("You: ").strip()

        if not user:
            continue

        cmd = user.lower()
        if cmd in ("exit","quit"):
            break
        if cmd.startswith("mode "):
            new = cmd.split()[1]
            if new in ("dual","text","mic","loop"):
                mode = new
                print(f"→ mode = {mode}")
            continue
        if cmd.startswith("time "):
            try:
                wait = int(cmd.split()[1])
                print(f"→ listen = {wait}s")
            except:
                print("time N 形式")
            continue
        if cmd.startswith("speed "):
            try:
                speed = float(cmd.split()[1])
                print(f"→ speed = {speed}x")
            except:
                print("speed X 形式")
            continue

        # ── 生成＆再生 ──
        reply = chat(client, user)
        print(f"きりたん: {reply}")
        speak(reply, speed)

        # mic モードなら続けて再リッスン
        if mode == 'mic':
            follow = listen_mic(wait)
            if follow:
                print(f"You (mic): {follow}")
                reply2 = chat(client, follow)
                print(f"きりたん: {reply2}")
                speak(reply2, speed)

if __name__ == "__main__":
    main()
