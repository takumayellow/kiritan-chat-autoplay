# kiritan-chat-autoplay.py (CLI 版)
"""
主な変更点
1. VOICEROID+ GUI を一切操作せず **SeikaSay2.exe** の `-play` オプションで直接再生。
2. これにより UI Automation 待ち時間・タブ競合が消え、応答が高速化。
3. PowerShell フォーカス維持も単に `bring_powershell_front()` で完結。
"""
import os
import subprocess
import ctypes
import time
import win32gui
import speech_recognition as sr
import sounddevice as sd
from openai import OpenAI

# ---------------- 基本設定 ----------------
DEFAULT_SPEED  = 1.0  # 読み上げ速度
DEFAULT_LISTEN = 0    # mic / loop の聞き取り秒数
CID_KIRITAN    = 1707 # 東北きりたんEX CID
SEIKA_CLI      = r"C:\Users\takum\Downloads\assistantseika20250113a\SeikaSay2\SeikaSay2.exe"

SYSTEM_PROMPT = (
    "あなたは「東北きりたんEX」です。"
    "声優・茜屋日海夏さんの柔らかく落ち着いた声質を踏まえ、可愛らしい口調で返答してください。"
    "フレーズごとに適度な話速と抑揚をつけ、必要に応じて『明るい』『デレ』『ダルい』『怒り』『泣き』など感情表現を交えてください。"
    "文末は『〜だよ』『〜だね』『〜かな？』等で親しみやすく締めましょう。"
    "無理に『〜だよ』を付ける必要はありません。"
)

# ---------------- OpenAI ----------------

def create_client() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY")
    if not key:
        raise ValueError("環境変数 OPENAI_API_KEY が未設定")
    return OpenAI(api_key=key)


def chat(client: OpenAI, prompt: str) -> str:
    res = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": prompt}
        ],
    )
    return res.choices[0].message.content.strip()

# ---------------- 入力系 ----------------

def listen_mic(limit: int) -> str:
    r = sr.Recognizer()
    with sr.Microphone() as mic:
        print(f"[mic] 発話どうぞ（最大 {limit}s）…")
        audio = r.listen(mic, phrase_time_limit=limit)
    try:
        return r.recognize_google(audio, language="ja-JP")
    except Exception:
        return ""


def listen_loopback(limit: int) -> str:
    print(f"[loop] システム音声録音（{limit}s）…")
    rec = sd.rec(int(limit*44100), samplerate=44100, channels=2)
    sd.wait()
    try:
        return sr.Recognizer().recognize_google(sr.AudioData(rec.tobytes(),44100,2), language="ja-JP")
    except Exception:
        return ""

# ---------------- 補助 ----------------

def bring_powershell_front():
    def _cb(hwnd, _):
        if "PowerShell" in win32gui.GetWindowText(hwnd):
            ctypes.windll.user32.SetForegroundWindow(hwnd)
    win32gui.EnumWindows(_cb, None)

# ---------------- 再生 (CLI -play) ----------------

def speak(text: str, speed: float):
    """SeikaSay2.exe -play で直接読み上げ。GUI 操作なし・高速。"""
    cmd = [
        SEIKA_CLI,
        "-cid",   str(CID_KIRITAN),
        "-speed", str(speed),
        "-play",              # 直接再生
        "-nc",                # コンソール出力抑制
        "-t", text
    ]
    subprocess.run(cmd)  # ブロッキングで再生完了を待つ
    bring_powershell_front()

# ---------------- メイン ----------------

def main():
    client = create_client()
    speed  = DEFAULT_SPEED
    wait   = DEFAULT_LISTEN
    mode   = 'dual'  # dual | text | mic | loop

    print("=== きりたんEX 会話 (CLI版) ===")
    print("mode dual/text/mic/loop | time N | speed X | exit")

    while True:
        if mode == 'dual':
            user = input("You: ").strip()
        elif mode == 'text':
            user = input("You (text): ")
        elif mode == 'mic':
            user = listen_mic(wait)
            if user:
                print(f"You (mic): {user}")
        elif mode == 'loop':
            user = listen_loopback(wait)
            if user:
                print(f"You (loop): {user}")
        else:
            user = input("You: ")

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

        # ------ 生成 & 再生 ------
        reply = chat(client, user)
        print(f"きりたん: {reply}")
        speak(reply, speed)

        if mode == 'mic':
            follow = listen_mic(wait)
            if follow:
                print(f"You (mic): {follow}")
                reply2 = chat(client, follow)
                print(f"きりたん: {reply2}")
                speak(reply2, speed)

if __name__ == "__main__":
    main()
