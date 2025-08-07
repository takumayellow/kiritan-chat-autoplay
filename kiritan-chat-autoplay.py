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
from pywinauto import Application

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
    "フレーズごとに適度な話速と抑揚をつけ、必要に応じて「明るい」「デレ」「ダルい」「怒り」「泣き」などの感情を表現するようなセリフにしてください。"
    "文末は「〜だよ」「〜だね」「〜かな？」等で親しみやすく締めましょう。"
    "無理に「〜だよ」を付ける必要はありません。"
    "そして会話が続くように、返答の最後に相手に対して自然な質問を一つ添えてください。"
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

# VOICEROID+ のウィンドウタイトル正規表現
WIN_TITLE_RE = r".*きりたん EX.*"

def switch_to_phrase_tab():
    """VOICEROIDウィンドウを取得し、フレーズ編集タブをクリックして選択する"""
    try:
        app = Application(backend="win32").connect(title_re=WIN_TITLE_RE)
        win = app.window(title_re=WIN_TITLE_RE)
        win.set_focus()
        # フレーズ編集タブをクリック
        win.child_window(title="フレーズ編集", control_type="TabItem").click_input()
    except Exception:
        # 見つからなかった場合は無視
        pass

# ---------------- 再生関数（CLI -play） ----------------
import psutil
from pywinauto import Application
import time

def speak(text: str, speed: float):
    # 1) VOICEROID.exe の PID を探す
    target_pid = None
    for proc in psutil.process_iter(['pid','name']):
        if proc.info['name'] == 'VOICEROID.exe':
            target_pid = proc.info['pid']
            break
    if not target_pid:
        print("⚠️ VOICEROID プロセスが見つかりません")
        return

    # 2) PID 指定で接続
    try:
        app = Application(backend="win32").connect(process=target_pid)
        # タイトルを指定せず PID だけでウィンドウを取得
        win = app.top_window()
    except Exception as e:
        print(f"⚠️ VOICEROID ウィンドウへの接続に失敗: {e}")
        return

    # 3) 前面化
    win.set_focus()

    # 4) フレーズ編集タブに切り替え
    try:
        # top_level_only=False で深い階層まで検索
        btn = win.child_window(
            title="再生",
            control_type="Button",
            top_level_only=False
        )
        btn.click_input()
    except Exception:
        win.type_keys('%2')  # Alt+2
    time.sleep(0.1)

    # 5) 再生ボタンをクリック
    try:
        win.child_window(title="再生", control_type="Button").click_input()
    except Exception as e:
        print(f"⚠️ 再生ボタンのクリックに失敗: {e}")
        return

    # 6) 再生完了まで待つ
    time.sleep(0.5)

    # 7) PowerShell を前面へ
    bring_powershell_front()


# ---------------- メインループ ----------------
from pywinauto import Application, timings

def connect_by_pid_hwnd(pid, hwnd):
    """UIA バックエンドでPIDとHWNDから確実にアタッチし、WindowSpecificationを返す"""
    try:
        app = timings.wait_until_passes(
            5, 0.5,
            lambda: Application(backend='uia').connect(process=pid, handle=hwnd)
        )
        win = app.window(handle=hwnd)
        return win
    except Exception as e:
        print(f"✖️ Connection failed: {e}")
        return None

def list_and_switch_tab(win):
    """Tabコントロールを取得 → タブ一覧を出力 → 'フレーズ編集' を選択"""
    try:
        tab_ctrl = win.child_window(control_type='Tab').wrapper_object()
    except Exception as e:
        print(f"✖️ Failed to locate Tab control: {e}")
        return

    try:
        tabs = tab_ctrl.tabs()
        print(f"◎ Available tabs: {tabs}")
    except Exception as e:
        print(f"✖️ Failed to list TabItems: {e}")
        return

    try:
        tab_ctrl.select('フレーズ編集')
        print("◎ 'フレーズ編集' tab selected successfully")
    except Exception as e:
        print(f"✖️ Tab switch failed: {e}")

def main():
    # --- ここから追加 ---
    hwnd, pid = find_voiceroid_handle()
    if hwnd and pid:
        win = connect_by_pid_hwnd(pid, hwnd)
        if win:
            list_and_switch_tab(win)
    else:
        print("⚠️ VOICEROID window not found → タブ切り替えスキップ")
    # --- ここまで追加 ---
    
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
