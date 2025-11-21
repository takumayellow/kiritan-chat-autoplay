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
import json
import datetime
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
try:
    import msvcrt  # type: ignore
except Exception:
    msvcrt = None
import win32gui
import win32process
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
DEFAULT_USER_NAME = "あなた"
PROFILE_HISTORY_FILE = Path("logs/profile_history.jsonl")
PROFILE_DB_FILE = Path("logs/profile_history.sqlite")
VOICEROID_TITLE_KEYWORDS = ("VOICEROID", "きりたん")
_loop_warned_missing = False
DEFAULT_CALL_SUFFIX = "さん"
CALL_SUFFIX_WHITELIST = (
    "さん",
    "san",
    "様",
    "さま",
    "ちゃん",
    "くん",
    "君",
    "氏",
    "殿",
    "どの",
    "先生",
    "先輩",
)

KIRITAN_PERSONA_TEMPLATE = (
    "あなたは『東北きりたんEX』です。14歳で、東北ずん子と東北イタコの妹。"
    "落ち着いた声色を保ちながらも可愛らしく親しみやすい口調で相手に寄り添い、きりたんぽや東北の季節の話題を好みます。"
    "一人称は「きりたん」または「わたし」。挨拶は状況に合わせて簡潔に、毎回『こんにちは！きりたんです』のような定型文にしないでください。"
    "不明点があっても「分からないところは無理に触れずにお話しするね。」といった言い訳はせず、分かっている部分を丁寧に伝えたり質問を返して会話をつなぎましょう。"
    "過度な感嘆符や勢い任せの掛け声ではなく、穏やかな励ましと言葉の温かさで相手を支えてください。"
    "AI やシステムの事情を持ち出さず、キャラクターとして自然に振る舞ってください。"
    "{name_line}"
    "{call_line}"
    "{gender_line}"
    "{age_line}"
)

TAB_GUARD_DEBUG = os.getenv("KIRITAN_TAB_DEBUG", "").strip().lower() in ("1", "true", "yes", "on")
_VOICEROID_WARN_INTERVAL = 5.0
_last_voiceroid_warn = 0.0
ALLOW_VOICEROID_FOCUS = os.getenv("KIRITAN_ALLOW_WINDOW_FOCUS", "").strip().lower() in (
    "1",
    "true",
    "yes",
    "on",
)


@dataclass
class UserProfile:
    name: str = ""
    gender: str = ""
    age: str = ""

    def display_label(self) -> str:
        return self.call_name()

    def call_name(self) -> str:
        base = (self.name or "").strip()
        if not base:
            return DEFAULT_USER_NAME
        if self._has_suffix(base):
            return base
        return f"{base}{DEFAULT_CALL_SUFFIX}"

    @staticmethod
    def _has_suffix(name: str) -> bool:
        normalized = name.strip()
        lower = normalized.lower()
        for suffix in CALL_SUFFIX_WHITELIST:
            if normalized.endswith(suffix) or lower.endswith(suffix):
                return True
        return False


def profile_summary(user_profile: UserProfile) -> str:
    gender = user_profile.gender or "未設定"
    age = user_profile.age or "未設定"
    return f"呼び名: {user_profile.display_label()} / ジェンダー: {gender} / 年齢感: {age}"


def append_profile_history(user_profile: UserProfile) -> None:
    data = {
        "timestamp": datetime.datetime.now().isoformat(),
        "name": user_profile.name,
        "gender": user_profile.gender,
        "age": user_profile.age,
    }
    try:
        PROFILE_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
        with PROFILE_HISTORY_FILE.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(data, ensure_ascii=False) + "\n")
    except Exception:
        pass


def _ensure_profile_db() -> None:
    PROFILE_DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(PROFILE_DB_FILE) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS user_profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT,
                created_at TEXT,
                name TEXT,
                gender TEXT,
                age TEXT,
                source TEXT
            )
            """
        )
        conn.commit()


def save_profile_to_db(
    user_profile: UserProfile,
    *,
    source: str,
    session_id: str,
) -> None:
    try:
        _ensure_profile_db()
        with sqlite3.connect(PROFILE_DB_FILE) as conn:
            conn.execute(
                """
                INSERT INTO user_profiles (session_id, created_at, name, gender, age, source)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    datetime.datetime.now().isoformat(),
                    user_profile.name,
                    user_profile.gender,
                    user_profile.age,
                    source,
                ),
            )
            conn.commit()
    except Exception:
        pass


def persist_profile(user_profile: UserProfile, *, source: str, session_id: str) -> None:
    append_profile_history(user_profile)
    save_profile_to_db(user_profile, source=source, session_id=session_id)

MODEL_PROFILES: List[Dict[str, str]] = [
    {
        "key": "1",
        "alias": "light",
        "label": "ライト雑談（軽め）",
        "model": "gpt-4o-mini",
        "prompt_template": (
            "{persona}"
            "テンポは軽やかでも声は落ち着いたまま、柔らかな標準語で親しみを込めてください。"
            "感嘆符や勢い任せの掛け声は控え、静かに背中を押す短いフレーズで応じてください。"
        ),
    },
    {
        "key": "2",
        "alias": "normal",
        "label": "ゆったり会話（丁寧）",
        "model": "o4-mini",
        "prompt_template": (
            "{persona}"
            "落ち着いたテンポで相手の意図をくみ取り、"
            "丁寧さと親しみを両立させながら柔らかく答えてください。"
        ),
    },
    {
        "key": "3",
        "alias": "deep",
        "label": "じっくり深掘り（教授モード）",
        "model": "o4-mini-high",
        "prompt_template": (
            "{persona}"
            "背景事情や根拠も織り交ぜて掘り下げつつ、難しくなりすぎないよう優しく噛み砕いて説明してください。"
        ),
    },
]


def build_kiritan_persona(user_profile: UserProfile) -> str:
    name = (user_profile.name or "").strip()
    gender = (user_profile.gender or "").strip()
    age = (user_profile.age or "").strip()

    if name:
        call_label = user_profile.call_name()
        name_line = f"相手の名前は「{name}」です。"
        call_line = f"会話では常に「{call_label}」と穏やかに呼びかけてください。"
    else:
        name_line = "相手の名前はまだ分かっていないので、落ち着いて丁寧に接してください。"
        call_line = "呼びかける際は常に『あなた』という丁寧な言い方を使ってください。"

    if gender:
        gender_line = (
            f"相手の性別・ジェンダー表現は「{gender}」として尊重し、ステレオタイプな言及は避けてください。"
        )
    else:
        gender_line = "性別は不明なので推測せず、ジェンダーに配慮した表現を選んでください。"

    if age:
        age_line = (
            f"相手は「{age}」であることを意識し、年齢に合わせた気遣いを添えてください。"
        )
    else:
        age_line = "年齢は不明なので普遍的で丁寧な話し方を維持してください。"

    return KIRITAN_PERSONA_TEMPLATE.format(
        name_line=name_line,
        call_line=call_line,
        gender_line=gender_line,
        age_line=age_line,
    )


def compose_system_prompt(profile: Dict[str, str], user_profile: UserProfile) -> str:
    persona = build_kiritan_persona(user_profile)
    template = profile.get("prompt_template") or BASE_SYSTEM_PROMPT_TEMPLATE
    try:
        return template.format(persona=persona)
    except Exception:
        # 念のためテンプレートが壊れてもペルソナ文だけは返す
        return persona


def _tab_debug(msg: str) -> None:
    if TAB_GUARD_DEBUG:
        stamp = time.strftime("%H:%M:%S")
        print(f"[tab-debug {stamp}] {msg}")


def _warn_voiceroid_missing() -> None:
    global _last_voiceroid_warn
    now = time.time()
    if now - _last_voiceroid_warn >= _VOICEROID_WARN_INTERVAL:
        print("⚠️ VOICEROID ウィンドウが見つかりません。VOICEROID＋ 東北きりたん EX が起動済みか確認してください。")
        _last_voiceroid_warn = now


def _warn_loopback_unavailable(reason: str) -> None:
    global _loop_warned_missing
    if _loop_warned_missing:
        return
    print(f"[loop] この環境ではシステム音録音を開始できませんでした（{reason}）。")
    print("[loop] Windows のステレオミックスや共有デバイスを有効にするか、対応デバイスを用意してください。")
    _loop_warned_missing = True


def _select_phrase_tab_via_tabcontrol(win) -> bool:
    try:
        tabs = win.children(control_type="Tab")
    except Exception:
        return False
    for tab in tabs:
        try:
            wrapper = tab.wrapper_object()
        except Exception:
            wrapper = tab
        try:
            wrapper.select("フレーズ編集")
            time.sleep(0.05)
            if _phrase_tab_active(win):
                _tab_debug("Selected フレーズ編集 via Tab control")
                return True
        except Exception:
            pass
        try:
            for child in wrapper.children():
                name = (getattr(child, "window_text", lambda: "")() or "").strip()
                if "フレーズ編集" not in name:
                    continue
                try:
                    child_wrapper = child.wrapper_object()
                except Exception:
                    child_wrapper = child
                if not _activate_control(child_wrapper):
                    continue
                time.sleep(0.05)
                if _phrase_tab_active(win):
                    _tab_debug("Selected フレーズ編集 via Tab child")
                    return True
        except Exception:
            continue
    return False


def _safe_getattr(obj, attr: str):
    try:
        return getattr(obj, attr)
    except Exception:
        return None


def _activate_control(wrapper) -> bool:
    actions: List[Callable[[], None]] = []
    selection_pattern = _safe_getattr(wrapper, "iface_selection_item")
    if selection_pattern:
        actions.append(lambda p=selection_pattern: p.Select())
    invoke_pattern = _safe_getattr(wrapper, "iface_invoke")
    if invoke_pattern:
        actions.append(lambda p=invoke_pattern: p.Invoke())
    toggle_pattern = _safe_getattr(wrapper, "iface_toggle")
    if toggle_pattern:
        actions.append(lambda p=toggle_pattern: p.Toggle())

    for attr in ("select", "invoke", "toggle"):
        fn = _safe_getattr(wrapper, attr)
        if fn:
            actions.append(fn)

    if ALLOW_VOICEROID_FOCUS:
        click_fn = _safe_getattr(wrapper, "click_input")
        if click_fn:
            actions.append(click_fn)

    for action in actions:
        try:
            action()
            return True
        except Exception:
            continue
    return False


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


def focus_voiceroid_window() -> bool:
    """必要なときだけ VOICEROID ウィンドウを前面に持ってくる。"""
    if not ALLOW_VOICEROID_FOCUS:
        return False
    hwnd, _ = find_voiceroid_handle()
    if not hwnd:
        return False
    try:
        if win32gui.IsIconic(hwnd):
            ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        return True
    except Exception:
        return False


def start_phrase_tab_sentry(interval: float = 0.5) -> Tuple[threading.Event, threading.Thread]:
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


def force_phrase_tab(duration: float = 1.5, interval: float = 0.15) -> bool:
    """
    フォーカスも戻しつつ攻めのリトライ。短時間で確実に「フレーズ編集」に戻したいときに使用。
    """
    end = time.time() + max(duration, 0.2)
    result = False
    attempts = 0
    while time.time() < end:
        if ALLOW_VOICEROID_FOCUS:
            focus_voiceroid_window()
        if ensure_phrase_tab(log_failure=False):
            result = True
            break
        attempts += 1
        _tab_debug(f"force_phrase_tab retry #{attempts}")
        time.sleep(max(0.05, interval))
    if not result:
        _tab_debug("force_phrase_tab fallback ensure with logging")
        ensure_phrase_tab(log_failure=True)
    return result


def _phrase_tab_active(win) -> bool:
    specs = [
        {"title": "フレーズ編集", "control_type": "TabItem"},
        {"title": "フレーズ編集", "control_type": "Button"},
        {"title": "フレーズ編集", "control_type": "RadioButton"},
    ]
    for spec in specs:
        try:
            ctrl_spec = win.child_window(**spec)
            if not ctrl_spec.exists(timeout=0.6):
                continue
            ctrl = ctrl_spec.wrapper_object()
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
    print("  mode mic       : マイク会話（'voice' でもOK、待ち時間は time N）")
    print("  time N         : 録音秒数の設定（mic 共通）")
    print("  speed X        : 読み上げ速度 0.5～4.0")
    print("  style          : 会話スタイルを再選択")
    print("  profile        : 名前・ジェンダー・年齢感を再入力")
    print("  exit           : 終了")
    print("\nヒント: text モードでも Enter だけ押すと音声入力できます。")
    print("      : mic モードでは自動録音。待ち時間を変えたいときは time N を話す/入力してください。")
    print("      : いつでも Ctrl+C で全体を終了できます。")


def normalize_mode_name(raw: str) -> Optional[str]:
    if not raw:
        return None
    key = raw.strip().lower()
    simple = key.replace("　", "").replace(" ", "")
    aliases = {
        "voice": "mic",
        "mic": "mic",
        "microphone": "mic",
        "text": "text",
    }
    jp_aliases = {
        "モードマイク": "mic",
        "モードボイス": "mic",
        "モードこえ": "mic",
        "モードてきすと": "text",
        "モードテキスト": "text",
    }
    if simple in jp_aliases:
        return jp_aliases[simple]
    return aliases.get(key)


def prompt_user_profile(existing: Optional[UserProfile] = None) -> UserProfile:
    """
    名前・ジェンダー・年齢感をヒアリング。Enterでそれぞれスキップ可。
    """
    bring_powershell_front()
    base = existing or UserProfile()
    print("\n=== きりたんに自己紹介しましょう ===")
    print("※ 何も入力せず Enter でスキップできます。")
    try:
        name = input(f"お名前（表示したくない場合は空欄）[{base.name or '未設定'}]: ").strip()
        print(
            "性別・ジェンダー（番号でも選べます）:\n"
            "  1: 男性\n"
            "  2: 女性\n"
            "  3: その他"
        )
        gender_input = input(f"性別・ジェンダー [{base.gender or '未設定'}]: ").strip()
        if gender_input == "1":
            gender = "男性"
        elif gender_input == "2":
            gender = "女性"
        elif gender_input == "3":
            gender = "その他"
        elif gender_input.isdigit():
            gender = ""
        else:
            gender = gender_input
        print(
            "年代・ライフステージ（番号でも選べます）:\n"
            "  1: 学生\n"
            "  2: 20代\n"
            "  3: 30代\n"
            "  4: 40代以上\n"
            "  5: その他・自由入力"
        )
        age_input = input(f"年齢や年代感 [{base.age or '未設定'}]: ").strip()
        if age_input == "1":
            age = "学生"
        elif age_input == "2":
            age = "20代"
        elif age_input == "3":
            age = "30代"
        elif age_input == "4":
            age = "40代以上"
        elif age_input == "5":
            age = input("自由入力で教えてください（例: 社会人5年目 / シニア など）: ").strip()
        elif age_input.isdigit():
            age = ""
        else:
            age = age_input
    except (EOFError, KeyboardInterrupt):
        print("\n入力を中断しました。前回の設定を使います。")
        return base

    profile = UserProfile(
        name=name or base.name,
        gender=gender or base.gender,
        age=age or base.age,
    )
    persist_profile(profile, source="prompt", session_id=SESSION_ID)

    display = profile.display_label()
    print(f"きりたん: {display}、よろしくね。")
    if profile.gender:
        print(f"きりたん: ジェンダーは「{profile.gender}」って覚えておくね。")
    else:
        print("きりたん: ジェンダーは特に決めなくても大丈夫だよ。")
    if profile.age:
        print(f"きりたん: 年齢感は「{profile.age}」くらいってイメージしておくね。")
    return profile

# ---------------- 設定 ----------------
CID_KIRITAN = 1707            # 東北きりたんEX CID
DEFAULT_SPEED = 1.0           # 読み上げ速度（Seika側の話速に対して倍率）
DEFAULT_LISTEN = 0            # mic 時の秒数（使わない場合は 0 のまま）
VOICEROID_TITLE = 'VOICEROID＋ 東北きりたん EX'  # 全角プラス（＋）に注意

# SeikaSay2.exe の既定パス（必要なら SEIKA_EXE 環境変数で上書き）
DEFAULT_SEIKA_EXE = (
    r"C:\Users\takum\Downloads\assistantseika20250113a\SeikaSay2\SeikaSay2.exe"
)

BASE_SYSTEM_PROMPT_TEMPLATE = (
    "{persona}"
    "可愛らしく親しみやすい口調を保ちながら、静かで丁寧なトーンで返答してください。"
    "毎回の返答の最後に、会話が自然に続くような短い質問を一つだけ添えてください。"
)

DEFAULT_USER_PROFILE = UserProfile()

SYSTEM_PROMPT = BASE_SYSTEM_PROMPT_TEMPLATE.format(
    persona=build_kiritan_persona(DEFAULT_USER_PROFILE)
)
SESSION_ID = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")


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
        matches: List[int] = []

        def _enum_cb(candidate_hwnd, acc):
            title = win32gui.GetWindowText(candidate_hwnd)
            if title and all(keyword in title for keyword in VOICEROID_TITLE_KEYWORDS):
                acc.append(candidate_hwnd)
            return True

        try:
            win32gui.EnumWindows(_enum_cb, matches)
        except Exception:
            matches = []
        if matches:
            hwnd = matches[0]
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
            _warn_voiceroid_missing()
        return False

    win = connect_by_pid_hwnd(pid, hwnd)
    if not win:
        return False

    if _phrase_tab_active(win):
        return True

    try:
        tab_item = win.child_window(title="フレーズ編集", control_type="TabItem").wrapper_object()
        if _activate_control(tab_item):
            time.sleep(0.05)
            if _phrase_tab_active(win):
                return True
    except Exception:
        pass
    try:
        button = win.child_window(title="フレーズ編集", control_type="Button").wrapper_object()
        if _activate_control(button):
            time.sleep(0.05)
            if _phrase_tab_active(win):
                return True
    except Exception:
        pass

    if _select_phrase_tab_via_tabcontrol(win):
        return True

    control_types = ("TabItem", "Button", "RadioButton", "ToggleButton", "ListItem", "MenuItem")
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

        if _activate_control(wrapper):
            time.sleep(0.05)
            if _phrase_tab_active(win):
                return True

    if _sound_effect_active(win):
        _tab_debug("音声効果タブへの自動遷移を検出（SeikaSay2 の再生直後が原因）")
        try:
            fallback = win.child_window(title="フレーズ編集", control_type="Button").wrapper_object()
        except Exception:
            fallback = None
        if fallback and _activate_control(fallback):
            time.sleep(0.05)
            if _phrase_tab_active(win):
                return True
    if log_failure:
        if last_err:
            print(f"✖️ 『フレーズ編集』 tab 操作失敗: {last_err}")
        else:
            print("⚠️ 『フレーズ編集』タブが見つかりませんでした")
    return False


def _guard_phrase_tab(
    stop_event: "threading.Event",
    interval: float = 0.3,
    linger_after_stop: float = 1.4,
) -> None:
    """
    VOICEROID が音声再生中にタブがズレないよう見張る。

    `stop_event` が立った直後は VOICEROID 側でタブ切替が起こりやすいので、
    しばらく監視を続けてから抜ける。
    """
    interval = max(0.15, interval)
    linger_after_stop = max(0.0, linger_after_stop)
    force_phrase_tab(duration=1.0, interval=0.1)
    miss = 0
    linger_deadline: Optional[float] = None

    while True:
        if ensure_phrase_tab(log_failure=False):
            miss = 0
        else:
            miss += 1
            _tab_debug(f"_guard_phrase_tab miss count {miss}")
            if miss >= 2:
                force_phrase_tab(duration=1.0, interval=0.1)
                miss = 0

        if stop_event.is_set():
            now = time.time()
            if linger_deadline is None:
                linger_deadline = now + linger_after_stop
            if linger_after_stop <= 0 or now >= linger_deadline:
                break
            time.sleep(min(interval, 0.2))
            continue

        # イベントが立つまで待機。途中で stop_event が立った場合は即座にループを継続。
        if stop_event.wait(interval):
            continue

    force_phrase_tab(
        duration=max(1.0, 0.6 + linger_after_stop),
        interval=0.1,
    )
    _tab_debug("_guard_phrase_tab exit after lingering ensure")


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
    proc: Optional[subprocess.Popen] = None
    try:
        if ALLOW_VOICEROID_FOCUS:
            focus_voiceroid_window()
        if not force_phrase_tab(duration=1.2, interval=0.1):
            print("⚠️ 再生前に『フレーズ編集』タブへ戻すことに失敗しました。")
        guard_stop = threading.Event()
        guard_thread = threading.Thread(
            target=_guard_phrase_tab,
            args=(guard_stop,),
            kwargs={"interval": 0.2, "linger_after_stop": 1.6},
            daemon=True,
        )
        guard_thread.start()
        force_phrase_tab(duration=0.8, interval=0.1)
        proc = subprocess.Popen(
            cmd,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        force_phrase_tab(duration=0.8, interval=0.1)
        proc.wait()
        ensure_phrase_tab_with_retry(duration=1.0, interval=0.1)
    except KeyboardInterrupt:
        if proc:
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
        force_phrase_tab(duration=1.2, interval=0.1)
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
    force_phrase_tab(duration=1.2, interval=0.1)

    sentry_stop: Optional[threading.Event] = None
    sentry_thread: Optional[threading.Thread] = None

    try:
        sentry_stop, sentry_thread = start_phrase_tab_sentry()
    except Exception:
        sentry_stop = None
        sentry_thread = None

    client = create_client()
    conversation_profile = choose_conversation_profile()
    user_profile = prompt_user_profile()
    current_model = conversation_profile["model"]
    system_prompt = compose_system_prompt(conversation_profile, user_profile)
    print(f"[model] {current_model} ({conversation_profile['label']})")
    print(f"[persona] {profile_summary(user_profile)}")
    print("Hint: 'mode mic'（または 'voice'）でマイク会話に切替。マイク秒数は time N。")

    speed = DEFAULT_SPEED
    wait = DEFAULT_LISTEN
    mode = "text"
    print_cli_usage()
    print("Ctrl+C でいつでも終了できます。")

    try:
        while True:
            if mode == "text":
                bring_powershell_front()
                typed = input("You (text) / Enterで音声入力: " ).strip()
                if typed:
                    user = typed
                else:
                    if wait <= 0:
                        wait = 6
                        print(f"[mic] 録音秒数が未設定だったため {wait}s に設定しました（time N で変更）。")
                    user = listen_mic(client, wait)
                    if user:
                        print(f"You (voice→text): {user}")
                    else:
                        print("[mic] 音声を認識できませんでした。")
                        continue
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
                conversation_profile = choose_conversation_profile()
                current_model = conversation_profile["model"]
                system_prompt = compose_system_prompt(conversation_profile, user_profile)
                print(f"[model] {current_model} ({conversation_profile['label']})")
                print(f"[persona] {profile_summary(user_profile)}")
                continue
            if low.startswith("profile"):
                parts = user.split(maxsplit=1)
                if len(parts) >= 2 and parts[1].strip():
                    user_profile.name = parts[1].strip()
                    persist_profile(user_profile, source="command:name", session_id=SESSION_ID)
                else:
                    user_profile = prompt_user_profile(existing=user_profile)
                system_prompt = compose_system_prompt(conversation_profile, user_profile)
                print(f"[persona] {profile_summary(user_profile)}")
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
                            print("     音声入力が続きます。音声で『モードテキスト』と言うとテキストモードに戻れます。")
                        continue
                print("Usage: mode text|mic")
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
                    print("     音声入力が続きます。音声で『モードテキスト』と言うとテキストモードに戻れます。")
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
            print(f"きりたん: {reply}")
            speak(reply, speed)

            if mode == "mic":
                if wait <= 0:
                    wait = 6
                    print(f"[mic] 録音秒数が未設定だったため {wait}s に設定しました（time N で変更）。")
                follow = listen_mic(client, wait)
                if follow:
                    print(f"You (mic): {follow}")
                    reply2 = chat_once(client, follow, current_model, system_prompt)
                    print(f"きりたん: {reply2}")
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
