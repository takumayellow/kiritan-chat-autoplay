# -*- coding: utf-8 -*-
"""
kiritan_chat_gui_new.py
きりたんチャット GUI アプリ（VOICEVOX版）

技術スタック:
  - Python + tkinter（標準ライブラリ）
  - VOICEVOX HTTP API（ローカル http://localhost:50021）
  - OpenAI API

ペルソナ・プロンプト設定は kiritan_chat_cli.py から引き継ぎ。
音声合成は kiritan_voicevox.py モジュールを使用。
"""

import io
import json
import os
import queue
import sys
import threading
import tkinter as tk
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
from typing import Any, Dict, List, Optional

# 画像表示用
try:
    from PIL import Image, ImageTk
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

# .env 読み込み（python-dotenv が入っていれば）
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# VOICEVOX モジュール
from core import kiritan_voicevox as vvox

# ── 共通ロジック（core/kiritan_core.py） ─────────────────
from core.kiritan_core import (
    DEFAULT_USER_NAME,
    DEFAULT_CALL_SUFFIX,
    CALL_SUFFIX_WHITELIST,
    KIRITAN_PERSONA_TEMPLATE,
    BASE_SYSTEM_PROMPT_TEMPLATE,
    MODEL_PROFILES,
    SANITIZE_TABLE,
    UserProfile,
    build_kiritan_persona,
    compose_system_prompt,
    sanitize_for_voice,
    get_model_profile,
    create_openai_client,
    chat_with_history,
)


@dataclass
class AppSettings:
    voicevox_url: str = "http://localhost:50021"
    voicevox_speaker_id: int = 58
    speed_scale: float = 1.0
    pitch_scale: float = 0.0
    intonation_scale: float = 1.0
    volume_scale: float = 1.0
    voice_enabled: bool = True
    model_key: str = "1"
    user_profile: UserProfile = field(default_factory=UserProfile)
    chat_history_max: int = 20  # 送信する会話履歴の最大ペア数

    def to_dict(self) -> dict:
        d = asdict(self)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "AppSettings":
        profile_data = d.pop("user_profile", {})
        profile = UserProfile(**{k: v for k, v in profile_data.items() if k in UserProfile.__dataclass_fields__})
        settings = cls(**{k: v for k, v in d.items() if k in cls.__dataclass_fields__})
        settings.user_profile = profile
        return settings


# ── 設定の保存/ロード ─────────────────────────────────────
SETTINGS_FILE = Path(os.path.expanduser("~")) / ".kiritan_gui_settings.json"
HISTORY_FILE = Path(os.path.expanduser("~")) / ".kiritan_gui_history.json"


def save_settings(settings: AppSettings) -> None:
    try:
        with SETTINGS_FILE.open("w", encoding="utf-8") as f:
            json.dump(settings.to_dict(), f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_settings() -> AppSettings:
    try:
        if SETTINGS_FILE.exists():
            with SETTINGS_FILE.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return AppSettings.from_dict(data)
    except Exception:
        pass
    return AppSettings()


def save_chat_history(history: List[Dict[str, str]], path: Optional[Path] = None) -> bool:
    target = path or HISTORY_FILE
    try:
        with open(target, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_chat_history(path: Optional[Path] = None) -> List[Dict[str, str]]:
    target = path or HISTORY_FILE
    try:
        if Path(target).exists():
            with open(target, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return []


# ── ダイアログ：プロフィール編集 ──────────────────────────
class ProfileDialog(simpledialog.Dialog):
    def __init__(self, parent, profile: UserProfile):
        self.profile = profile
        self.result_profile: Optional[UserProfile] = None
        super().__init__(parent, title="プロフィール編集")

    def body(self, master):
        master.configure(bg="#ffffff")

        def _lbl(text, row):
            tk.Label(
                master, text=text, bg="#ffffff", fg=THEME["text"],
                font=(_FONT, 10),
            ).grid(row=row, column=0, sticky="w", padx=8, pady=4)

        _lbl("お名前:", 0)
        self.name_var = tk.StringVar(value=self.profile.name)
        tk.Entry(master, textvariable=self.name_var, bg="#f9fafb", fg=THEME["text"],
                 insertbackground=THEME["accent"], font=(_FONT, 10)).grid(
            row=0, column=1, sticky="ew", padx=8, pady=4)

        _lbl("性別・ジェンダー:", 1)
        self.gender_var = tk.StringVar(value=self.profile.gender)
        gender_combo = ttk.Combobox(
            master, textvariable=self.gender_var,
            values=["", "男性", "女性", "その他"],
            font=(_FONT, 10), width=20,
        )
        gender_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=4)

        _lbl("年代感:", 2)
        self.age_var = tk.StringVar(value=self.profile.age)
        age_combo = ttk.Combobox(
            master, textvariable=self.age_var,
            values=["", "学生", "20代", "30代", "40代以上", "その他"],
            font=(_FONT, 10), width=20,
        )
        age_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=4)

        master.grid_columnconfigure(1, weight=1)
        return master

    def apply(self):
        self.result_profile = UserProfile(
            name=self.name_var.get().strip(),
            gender=self.gender_var.get().strip(),
            age=self.age_var.get().strip(),
        )


# ── ダイアログ：詳細設定 ──────────────────────────────────
class SettingsDialog(simpledialog.Dialog):
    def __init__(self, parent, settings: AppSettings):
        self.settings = settings
        self.result_settings: Optional[AppSettings] = None
        super().__init__(parent, title="設定")

    def body(self, master):
        master.configure(bg="#ffffff")

        def _lbl(text, row):
            tk.Label(
                master, text=text, bg="#ffffff", fg=THEME["text"],
                font=(_FONT, 10),
            ).grid(row=row, column=0, sticky="w", padx=8, pady=4)

        # VOICEVOX URL
        _lbl("VOICEVOX URL:", 0)
        self.url_var = tk.StringVar(value=self.settings.voicevox_url)
        tk.Entry(master, textvariable=self.url_var, bg="#f9fafb", fg=THEME["text"],
                 insertbackground=THEME["accent"], font=(_FONT, 10), width=32).grid(
            row=0, column=1, sticky="ew", padx=8, pady=4)

        # Speaker ID
        _lbl("Speaker ID:", 1)
        self.speaker_var = tk.IntVar(value=self.settings.voicevox_speaker_id)
        tk.Spinbox(master, from_=0, to=999, textvariable=self.speaker_var,
                   bg="#f9fafb", fg=THEME["text"], buttonbackground=THEME["border"],
                   font=(_FONT, 10), width=8).grid(
            row=1, column=1, sticky="w", padx=8, pady=4)

        # 話速
        _lbl("話速 (0.5〜2.0):", 2)
        self.speed_var = tk.DoubleVar(value=self.settings.speed_scale)
        tk.Scale(master, from_=0.5, to=2.0, resolution=0.1, orient="horizontal",
                 variable=self.speed_var, bg="#ffffff", fg=THEME["text_sub"],
                 highlightbackground="#ffffff", troughcolor=THEME["border"],
                 activebackground=THEME["accent"], length=200).grid(
            row=2, column=1, sticky="ew", padx=8, pady=4)

        # 音高
        _lbl("音高 (-0.15〜0.15):", 3)
        self.pitch_var = tk.DoubleVar(value=self.settings.pitch_scale)
        tk.Scale(master, from_=-0.15, to=0.15, resolution=0.01, orient="horizontal",
                 variable=self.pitch_var, bg="#ffffff", fg=THEME["text_sub"],
                 highlightbackground="#ffffff", troughcolor=THEME["border"],
                 activebackground=THEME["accent"], length=200).grid(
            row=3, column=1, sticky="ew", padx=8, pady=4)

        # 抑揚
        _lbl("抑揚 (0.0〜2.0):", 4)
        self.intonation_var = tk.DoubleVar(value=self.settings.intonation_scale)
        tk.Scale(master, from_=0.0, to=2.0, resolution=0.1, orient="horizontal",
                 variable=self.intonation_var, bg="#ffffff", fg=THEME["text_sub"],
                 highlightbackground="#ffffff", troughcolor=THEME["border"],
                 activebackground=THEME["accent"], length=200).grid(
            row=4, column=1, sticky="ew", padx=8, pady=4)

        # 音量
        _lbl("音量 (0.0〜2.0):", 5)
        self.volume_var = tk.DoubleVar(value=self.settings.volume_scale)
        tk.Scale(master, from_=0.0, to=2.0, resolution=0.1, orient="horizontal",
                 variable=self.volume_var, bg="#ffffff", fg=THEME["text_sub"],
                 highlightbackground="#ffffff", troughcolor=THEME["border"],
                 activebackground=THEME["accent"], length=200).grid(
            row=5, column=1, sticky="ew", padx=8, pady=4)

        master.grid_columnconfigure(1, weight=1)
        return master

    def apply(self):
        self.result_settings = AppSettings(
            voicevox_url=self.url_var.get().strip(),
            voicevox_speaker_id=int(self.speaker_var.get()),
            speed_scale=round(self.speed_var.get(), 2),
            pitch_scale=round(self.pitch_var.get(), 3),
            intonation_scale=round(self.intonation_var.get(), 2),
            volume_scale=round(self.volume_var.get(), 2),
            voice_enabled=self.settings.voice_enabled,
            model_key=self.settings.model_key,
            user_profile=self.settings.user_profile,
            chat_history_max=self.settings.chat_history_max,
        )


# ── カラーテーマ（白ベース・モダン） ─────────────────────
_FONT = "Segoe UI"  # Windows モダンフォント

THEME = {
    "bg":           "#ffffff",   # メイン背景（白）
    "bg2":          "#f7f7f8",   # ヘッダー/ステータス背景
    "bg3":          "#f0f0f2",   # 入力欄背景
    "accent":       "#8b5cf6",   # きりたん紫（バイオレット）
    "accent_light": "#a78bfa",   # 薄め紫
    "accent_hover": "#7c3aed",   # ホバー
    "accent_bg":    "#f5f3ff",   # 紫の超薄い背景
    "kiritan_fg":   "#1e1b4b",   # きりたん発言テキスト（濃紺）
    "kiritan_bg":   "#ede9fe",   # きりたん発言背景（ラベンダー）
    "user_fg":      "#1e3a5f",   # ユーザー発言テキスト
    "user_bg":      "#e0f2fe",   # ユーザー発言背景（水色）
    "system_fg":    "#9ca3af",   # システムメッセージ
    "text":         "#1f2937",   # 通常テキスト（ほぼ黒）
    "text_sub":     "#6b7280",   # サブテキスト
    "border":       "#e5e7eb",   # ボーダー
    "input_border": "#d1d5db",   # 入力欄ボーダー
    "input_focus":  "#8b5cf6",   # 入力欄フォーカス色
    "status_ok":    "#10b981",   # 接続OK（エメラルド）
    "status_ng":    "#ef4444",   # 接続NG
    "status_wait":  "#f59e0b",   # 待機中
    "send_bg":      "#8b5cf6",   # 送信ボタン
    "send_hover":   "#7c3aed",   # 送信ボタンホバー
}


# ── メイン GUI クラス ──────────────────────────────────────
class KiritanChatGUINew:
    """
    きりたんチャット GUI アプリ（VOICEVOX版）
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.settings = load_settings()

        # 画像参照を保持（GC防止）
        self._images: Dict[str, Any] = {}
        self._load_images()

        # 会話状態
        self.client: Optional[Any] = None
        self.chat_history: List[Dict[str, str]] = []  # {"role": ..., "content": ...}
        self.is_generating = False
        self.is_speaking = False

        # コンテキスト：現在のモデルプロファイル
        self.model_profile: Dict[str, str] = self._get_model_profile(self.settings.model_key)
        self.system_prompt: str = compose_system_prompt(
            self.settings.user_profile, model_profile=self.model_profile
        )

        # キュー（バックグラウンドスレッド → UIスレッド）
        self.msg_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

        self._setup_window()
        self._setup_ui()
        self._start_queue_poll()
        self._init_client()
        self._check_voicevox()

        # 起動メッセージ
        self._add_system_message(
            "きりたんとの会話を開始しましょう！"
        )
        self._add_system_message(
            "下の入力欄にメッセージを入力して Enter で送信できます。"
        )

    # ── 画像読み込み ────────────────────────────────────
    def _load_images(self):
        if not _PIL_AVAILABLE:
            return
        assets = Path(__file__).parent.parent / "assets"

        # ヘッダーアイコン（40x40）
        icon_path = assets / "kiritan_small.png"
        if icon_path.exists():
            try:
                img = Image.open(icon_path)
                img = img.resize((40, 40), Image.LANCZOS)
                self._images["icon"] = ImageTk.PhotoImage(img)
            except Exception:
                pass

        # 立ち絵（元画像を保持、表示時にリサイズ）
        stand_path = assets / "kiritan_stand.png"
        if stand_path.exists():
            try:
                self._stand_pil = Image.open(stand_path).copy()
                # 初期サイズ（後でウィンドウに合わせてリサイズ）
                ratio = 350 / self._stand_pil.height
                new_w = int(self._stand_pil.width * ratio)
                resized = self._stand_pil.resize((new_w, 350), Image.LANCZOS)
                self._images["stand"] = ImageTk.PhotoImage(resized)
            except Exception:
                pass

    def _resize_stand(self, panel):
        """立ち絵をパネルの高さに合わせてリサイズ"""
        if not hasattr(self, "_stand_pil"):
            return
        panel_h = panel.winfo_height()
        if panel_h < 50:
            return
        target_h = max(100, panel_h - 16)
        ratio = target_h / self._stand_pil.height
        new_w = int(self._stand_pil.width * ratio)
        resized = self._stand_pil.resize((new_w, target_h), Image.LANCZOS)
        self._images["stand"] = ImageTk.PhotoImage(resized)
        self._stand_label.configure(image=self._images["stand"])

    # ── ウィンドウ設定 ────────────────────────────────────
    def _setup_window(self):
        self.root.title("きりたんチャット")
        self.root.configure(bg=THEME["bg"])
        self.root.geometry("800x700")
        self.root.minsize(560, 480)

        # アイコン（存在する場合）
        try:
            icon_path = Path(__file__).parent.parent / "assets" / "kiritan_icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            pass

    # ── UI 構築 ───────────────────────────────────────────
    def _setup_ui(self):
        self._apply_styles()
        self._setup_menubar()
        self._setup_header()
        # ステータスバーと入力エリアを先にpack（bottomから）
        # → チャットエリアが残りスペースを埋める
        self._setup_statusbar()
        self._setup_input_area()
        self._setup_chat_area()

    def _apply_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox",
                         fieldbackground=THEME["bg"],
                         background=THEME["bg"],
                         foreground=THEME["text"],
                         selectbackground=THEME["accent"],
                         selectforeground="#ffffff",
                         arrowcolor=THEME["text_sub"])
        style.configure("TCheckbutton",
                         background=THEME["bg2"],
                         foreground=THEME["text"])
        style.map("TCheckbutton",
                   background=[("active", THEME["bg2"])],
                   foreground=[("active", THEME["accent"])])

    def _setup_menubar(self):
        menubar = tk.Menu(self.root, bg=THEME["bg2"], fg=THEME["text"],
                          activebackground=THEME["accent"],
                          activeforeground="#ffffff",
                          font=(_FONT, 9))
        self.root.config(menu=menubar)

        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0, bg="#ffffff", fg=THEME["text"],
                            activebackground=THEME["accent"], activeforeground="#ffffff",
                            font=(_FONT, 9))
        file_menu.add_command(label="会話履歴を保存", command=self._save_history_as)
        file_menu.add_command(label="会話履歴を読み込み", command=self._load_history_from)
        file_menu.add_separator()
        file_menu.add_command(label="履歴をクリア", command=self._clear_history)
        file_menu.add_separator()
        file_menu.add_command(label="終了", command=self._on_close)
        menubar.add_cascade(label="ファイル", menu=file_menu)

        # 設定メニュー
        settings_menu = tk.Menu(menubar, tearoff=0, bg="#ffffff", fg=THEME["text"],
                                 activebackground=THEME["accent"], activeforeground="#ffffff",
                                 font=(_FONT, 9))
        settings_menu.add_command(label="プロフィール編集", command=self._edit_profile)
        settings_menu.add_command(label="VOICEVOX・音声設定", command=self._edit_settings)
        settings_menu.add_separator()
        settings_menu.add_command(label="VOICEVOX 接続確認", command=self._check_voicevox_manual)
        menubar.add_cascade(label="設定", menu=settings_menu)

    def _setup_header(self):
        header_frame = tk.Frame(self.root, bg=THEME["bg2"], pady=10)
        header_frame.pack(fill="x")

        # 左: アイコン + タイトル
        title_frame = tk.Frame(header_frame, bg=THEME["bg2"])
        title_frame.pack(side="left", padx=16)
        if "icon" in self._images:
            tk.Label(
                title_frame, image=self._images["icon"],
                bg=THEME["bg2"],
            ).pack(side="left", padx=(0, 8))
        tk.Label(
            title_frame, text="きりたんチャット",
            bg=THEME["bg2"], fg=THEME["text"],
            font=(_FONT, 18, "bold"),
        ).pack(side="left")
        # VOICEVOXバッジ
        badge = tk.Label(
            title_frame, text=" VOICEVOX ",
            bg=THEME["accent"], fg="#ffffff",
            font=(_FONT, 8, "bold"),
            padx=6, pady=1,
        )
        badge.pack(side="left", padx=(8, 0), pady=(4, 0))

        # 右: コントロール群
        right_frame = tk.Frame(header_frame, bg=THEME["bg2"])
        right_frame.pack(side="right", padx=16)

        # VOICEVOX 接続インジケーター
        self.voicevox_status_label = tk.Label(
            right_frame, text="●",
            bg=THEME["bg2"], fg=THEME["status_wait"],
            font=(_FONT, 10), cursor="hand2",
        )
        self.voicevox_status_label.pack(side="right", padx=(12, 0))
        self.voicevox_status_label.bind("<Button-1>", lambda e: self._check_voicevox_manual())

        # 音声ON/OFF
        self.voice_var = tk.BooleanVar(value=self.settings.voice_enabled)
        voice_cb = ttk.Checkbutton(
            right_frame, text="音声",
            variable=self.voice_var,
            command=self._on_voice_toggle,
        )
        voice_cb.pack(side="right", padx=8)

        # スタイル選択
        tk.Label(
            right_frame, text="スタイル",
            bg=THEME["bg2"], fg=THEME["text_sub"],
            font=(_FONT, 9),
        ).pack(side="right", padx=(8, 4))

        self.style_var = tk.StringVar(value=self._model_profile_label(self.model_profile))
        self.style_combo = ttk.Combobox(
            right_frame,
            textvariable=self.style_var,
            values=[self._model_profile_label(p) for p in MODEL_PROFILES],
            state="readonly", width=22,
            font=(_FONT, 9),
        )
        self.style_combo.pack(side="right")
        self.style_combo.bind("<<ComboboxSelected>>", self._on_style_change)

        # 区切り線
        tk.Frame(self.root, bg=THEME["border"], height=1).pack(fill="x")

    def _setup_chat_area(self):
        # メインコンテンツ: チャット + 立ち絵
        main_frame = tk.Frame(self.root, bg=THEME["bg"])
        main_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # 右側: 立ち絵パネル
        if "stand" in self._images:
            stand_panel = tk.Frame(main_frame, bg=THEME["accent_bg"])
            stand_panel.pack(side="right", fill="y")
            self._stand_label = tk.Label(
                stand_panel, image=self._images["stand"],
                bg=THEME["accent_bg"],
            )
            self._stand_label.pack(side="bottom", pady=(0, 8))
            # ウィンドウリサイズ時に立ち絵をフィットさせる
            def _on_resize(event, panel=stand_panel):
                self._resize_stand(panel)
            self.root.bind("<Configure>", _on_resize)
            # 区切り線
            tk.Frame(main_frame, bg=THEME["border"], width=1).pack(side="right", fill="y")

        chat_frame = tk.Frame(main_frame, bg=THEME["bg"])
        chat_frame.pack(side="left", fill="both", expand=True)

        self.chat_text = tk.Text(
            chat_frame,
            bg=THEME["bg"],
            fg=THEME["text"],
            font=(_FONT, 11),
            state="disabled",
            wrap="word",
            relief="flat",
            borderwidth=0,
            selectbackground=THEME["accent_light"],
            selectforeground="#ffffff",
            spacing1=2,
            spacing3=4,
            padx=20,
            pady=12,
            cursor="arrow",
        )
        scrollbar = ttk.Scrollbar(chat_frame, orient="vertical", command=self.chat_text.yview)
        self.chat_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.chat_text.pack(side="left", fill="both", expand=True)

        # テキストタグ設定
        self.chat_text.tag_configure(
            "kiritan_name",
            foreground=THEME["accent"],
            font=(_FONT, 10, "bold"),
            spacing1=12,
        )
        self.chat_text.tag_configure(
            "kiritan_text",
            foreground=THEME["kiritan_fg"],
            background=THEME["kiritan_bg"],
            font=(_FONT, 11),
            lmargin1=12, lmargin2=12,
            rmargin=80,
            spacing1=2, spacing3=6,
        )
        self.chat_text.tag_configure(
            "user_name",
            foreground="#2563eb",
            font=(_FONT, 10, "bold"),
            spacing1=12,
        )
        self.chat_text.tag_configure(
            "user_text",
            foreground=THEME["user_fg"],
            background=THEME["user_bg"],
            font=(_FONT, 11),
            lmargin1=80, lmargin2=80,
            rmargin=12,
            spacing1=2, spacing3=6,
        )
        self.chat_text.tag_configure(
            "system_text",
            foreground=THEME["system_fg"],
            font=(_FONT, 9),
            justify="center",
            spacing1=6, spacing3=6,
        )
        self.chat_text.tag_configure(
            "timestamp",
            foreground="#c0c0c8",
            font=(_FONT, 8),
        )

    def _setup_input_area(self):
        # 区切り線
        tk.Frame(self.root, bg=THEME["border"], height=1).pack(fill="x")

        # 入力エリア全体の背景
        input_bg = tk.Frame(self.root, bg=THEME["bg2"], pady=8)
        input_bg.pack(fill="x")

        # ── メイン入力行: テキスト + 送信ボタン ──
        input_row = tk.Frame(input_bg, bg=THEME["bg2"])
        input_row.pack(fill="x", padx=16)

        # ボーダー付き入力欄コンテナ
        input_border = tk.Frame(
            input_row, bg=THEME["input_border"],
            highlightthickness=0,
        )
        input_border.pack(side="left", fill="both", expand=True, padx=(0, 8))

        # 入力テキストウィジェット（直接配置、ネスト最小化）
        self.input_text = tk.Text(
            input_border,
            bg="#ffffff",
            fg=THEME["text"],
            font=(_FONT, 12),
            height=2,
            wrap="word",
            relief="solid",
            borderwidth=1,
            insertbackground=THEME["accent"],
            selectbackground=THEME["accent_light"],
            selectforeground="#ffffff",
            padx=12,
            pady=8,
        )
        self.input_text.pack(fill="both", expand=True, padx=1, pady=1)
        self.input_text.bind("<Return>", self._on_return_key)
        self.input_text.bind("<Shift-Return>", lambda e: None)

        # フォーカス時のビジュアルフィードバック
        def _on_focus_in(e):
            input_border.configure(bg=THEME["input_focus"])
        def _on_focus_out(e):
            input_border.configure(bg=THEME["input_border"])
        self.input_text.bind("<FocusIn>", _on_focus_in)
        self.input_text.bind("<FocusOut>", _on_focus_out)

        # 送信ボタン
        self.send_btn = tk.Button(
            input_row,
            text="  送信  ",
            command=self._on_send,
            bg=THEME["send_bg"],
            fg="#ffffff",
            font=(_FONT, 14, "bold"),
            activebackground=THEME["send_hover"],
            activeforeground="#ffffff",
            relief="flat",
            cursor="hand2",
            padx=28,
            pady=14,
        )
        self.send_btn.pack(side="right", fill="y")

        # ── サブボタン行 ──
        sub_frame = tk.Frame(input_bg, bg=THEME["bg2"])
        sub_frame.pack(fill="x", padx=16, pady=(6, 0))

        for text, cmd in [("プロフィール", self._edit_profile), ("クリア", self._clear_input)]:
            tk.Button(
                sub_frame, text=text, command=cmd,
                bg=THEME["bg2"], fg=THEME["text_sub"],
                font=(_FONT, 9),
                activebackground=THEME["bg3"],
                activeforeground=THEME["text"],
                relief="flat", cursor="hand2",
                padx=8, pady=1,
            ).pack(side="left", padx=(0, 4))

        # マイク
        self.mic_btn = tk.Button(
            sub_frame, text="マイク",
            command=self._on_mic,
            bg=THEME["bg2"], fg=THEME["accent"],
            font=(_FONT, 9),
            activebackground=THEME["bg3"],
            activeforeground=THEME["accent_hover"],
            relief="flat", cursor="hand2",
            padx=8, pady=1,
        )
        self.mic_btn.pack(side="right")

        tk.Label(
            sub_frame,
            text="Enter 送信 ・ Shift+Enter 改行",
            bg=THEME["bg2"], fg="#c0c0c8",
            font=(_FONT, 8),
        ).pack(side="right", padx=12)

        # 確実にフォーカスを設定
        self.root.after(200, self._force_focus_input)

    def _force_focus_input(self):
        """入力欄にフォーカスを強制設定"""
        self.input_text.focus_force()
        self.input_text.mark_set("insert", "1.0")

    def _setup_statusbar(self):
        status_frame = tk.Frame(self.root, bg=THEME["bg3"], pady=3)
        status_frame.pack(fill="x", side="bottom")

        self.status_var = tk.StringVar(value="準備中...")
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg=THEME["bg3"],
            fg=THEME["text_sub"],
            font=(_FONT, 8),
            anchor="w",
        ).pack(side="left", padx=16)

        self.model_label_var = tk.StringVar(value="")
        tk.Label(
            status_frame,
            textvariable=self.model_label_var,
            bg=THEME["bg3"],
            fg=THEME["accent"],
            font=(_FONT, 8, "bold"),
            anchor="e",
        ).pack(side="right", padx=16)

    # ── 初期化 ────────────────────────────────────────────
    def _init_client(self):
        self.client = create_openai_client()
        if self.client:
            self.status_var.set("OpenAI 接続OK")
        else:
            key = os.environ.get("OPENAI_API_KEY", "").strip()
            if not key:
                self.status_var.set("OPENAI_API_KEY が未設定です。.env ファイルを確認してください。")
            else:
                self.status_var.set("OpenAI クライアントの初期化に失敗しました。")

    def _check_voicevox(self):
        """バックグラウンドで VOICEVOX 接続確認"""
        def _worker():
            ok, msg = vvox.check_connection(self.settings.voicevox_url)
            self.msg_queue.put({"type": "voicevox_status", "ok": ok, "msg": msg})

            if ok:
                # きりたんの speaker_id を自動検索
                sid = vvox.find_kiritan_speaker_id(self.settings.voicevox_url)
                if sid is not None and sid != self.settings.voicevox_speaker_id:
                    self.msg_queue.put({
                        "type": "voicevox_speaker_found",
                        "speaker_id": sid,
                    })

        threading.Thread(target=_worker, daemon=True).start()

    def _check_voicevox_manual(self):
        self._set_status("VOICEVOX 接続確認中...")
        self._check_voicevox()

    # ── キュー処理 ────────────────────────────────────────
    def _start_queue_poll(self):
        self.root.after(100, self._poll_queue)

    def _poll_queue(self):
        try:
            while True:
                item = self.msg_queue.get_nowait()
                self._handle_queue_item(item)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._poll_queue)

    def _handle_queue_item(self, item: Dict[str, Any]):
        itype = item.get("type")

        if itype == "voicevox_status":
            ok = item["ok"]
            msg = item["msg"]
            color = THEME["status_ok"] if ok else THEME["status_ng"]
            self.voicevox_status_label.configure(fg=color)
            self._set_status(msg)

        elif itype == "voicevox_speaker_found":
            sid = item["speaker_id"]
            self.settings.voicevox_speaker_id = sid
            save_settings(self.settings)
            self._add_system_message(f"東北きりたんの Speaker ID を {sid} に設定しました。")

        elif itype == "reply":
            reply_text = item["text"]
            self._add_kiritan_message(reply_text)
            self.chat_history.append({"role": "assistant", "content": reply_text})
            self._set_busy(False)
            model_name = self.model_profile.get("model", "")
            self._set_status(f"返答完了  [model: {model_name}]")
            self.model_label_var.set(f"model: {model_name}")

            # 音声再生
            if self.voice_var.get():
                self._speak_async(reply_text)

        elif itype == "error":
            err_text = item["text"]
            self._add_system_message(f"エラー: {err_text}")
            self._set_busy(False)
            self._set_status(f"エラーが発生しました: {err_text}")

        elif itype == "mic_result":
            text = item.get("text", "").strip()
            if text:
                self.input_text.insert("end", text)
                self._set_status(f"音声認識: {text}")
            else:
                self._set_status("音声が認識できませんでした。")
            self.mic_btn.configure(state="normal")

        elif itype == "status":
            self._set_status(item["text"])

    # ── チャット表示 ──────────────────────────────────────
    def _add_message(self, speaker_tag: str, name_tag: str, speaker: str, text: str):
        self.chat_text.configure(state="normal")
        ts = datetime.now().strftime("%H:%M")

        self.chat_text.insert("end", f"\n[{ts}] ", "timestamp")
        self.chat_text.insert("end", f"{speaker}\n", name_tag)
        self.chat_text.insert("end", f"{text}\n", speaker_tag)

        self.chat_text.configure(state="disabled")
        self.chat_text.see("end")

    def _add_kiritan_message(self, text: str):
        self._add_message("kiritan_text", "kiritan_name", "きりたん", text)

    def _add_user_message(self, text: str):
        self._add_message("user_text", "user_name",
                          self.settings.user_profile.display_label(), text)

    def _add_system_message(self, text: str):
        self.chat_text.configure(state="normal")
        self.chat_text.insert("end", f"\n{text}\n", "system_text")
        self.chat_text.configure(state="disabled")
        self.chat_text.see("end")

    # ── 送信処理 ──────────────────────────────────────────
    def _on_return_key(self, event):
        # Shift+Enter は改行、Enter のみで送信
        if event.state & 0x1:  # Shiftキーが押されている
            return None
        self._on_send()
        return "break"

    def _on_send(self):
        text = self.input_text.get("1.0", "end").strip()
        if not text:
            return
        if self.is_generating:
            self._set_status("返答生成中です。しばらくお待ちください。")
            return

        if self.client is None:
            messagebox.showwarning(
                "OpenAI 未接続",
                "OPENAI_API_KEY が設定されていません。\n"
                ".env ファイルに OPENAI_API_KEY を設定してください。",
            )
            return

        # 入力をクリアして表示
        self.input_text.delete("1.0", "end")
        self._add_user_message(text)
        self.chat_history.append({"role": "user", "content": text})

        # 送信開始
        self._set_busy(True)
        self._set_status("返答を生成しています...")

        threading.Thread(target=self._generate_reply, args=(text,), daemon=True).start()

    def _generate_reply(self, user_text: str):
        try:
            reply = chat_with_history(
                client=self.client,
                user_text=user_text,
                history=self.chat_history[:-1],  # 最後に追加した user メッセージを除く（重複回避）
                system_prompt=self.system_prompt,
                preferred_model=self.model_profile.get("model", "gpt-4o-mini"),
                history_max=self.settings.chat_history_max,
            )
            self.msg_queue.put({"type": "reply", "text": reply})
        except Exception as e:
            self.msg_queue.put({"type": "error", "text": str(e)})

    def _speak_async(self, text: str):
        if not text:
            return

        def _worker():
            self.is_speaking = True
            self.msg_queue.put({"type": "status", "text": "音声合成・再生中..."})
            vvox.speak(
                text=vvox.sanitize_for_voice(text),
                speaker_id=self.settings.voicevox_speaker_id,
                speed_scale=self.settings.speed_scale,
                pitch_scale=self.settings.pitch_scale,
                intonation_scale=self.settings.intonation_scale,
                volume_scale=self.settings.volume_scale,
                base_url=self.settings.voicevox_url,
            )
            self.is_speaking = False
            self.msg_queue.put({"type": "status", "text": "再生完了"})

        threading.Thread(target=_worker, daemon=True).start()

    # ── マイク入力 ────────────────────────────────────────
    def _on_mic(self):
        self.mic_btn.configure(state="disabled")
        self._set_status("音声認識中...")
        threading.Thread(target=self._listen_mic, daemon=True).start()

    def _listen_mic(self):
        """
        SpeechRecognition を使ってマイク入力をテキスト化する。
        パッケージが未インストールの場合はメッセージを表示。
        """
        try:
            import speech_recognition as sr
        except ImportError:
            self.msg_queue.put({
                "type": "status",
                "text": "speech_recognition が未インストールです。pip install SpeechRecognition でインストールしてください。",
            })
            self.msg_queue.put({"type": "mic_result", "text": ""})
            return

        r = sr.Recognizer()
        try:
            with sr.Microphone() as mic:
                self.msg_queue.put({"type": "status", "text": "発話してください..."})
                r.adjust_for_ambient_noise(mic, duration=0.5)
                audio = r.listen(mic, phrase_time_limit=10)
            text = r.recognize_google(audio, language="ja-JP")
            self.msg_queue.put({"type": "mic_result", "text": text})
        except Exception as e:
            self.msg_queue.put({"type": "mic_result", "text": ""})

    # ── UI ヘルパー ───────────────────────────────────────
    def _set_busy(self, busy: bool):
        self.is_generating = busy
        state = "disabled" if busy else "normal"
        self.send_btn.configure(state=state)

    def _set_status(self, text: str):
        self.status_var.set(text)

    def _clear_input(self):
        self.input_text.delete("1.0", "end")

    # ── モデル関連 ────────────────────────────────────────
    def _get_model_profile(self, key: str) -> Dict[str, str]:
        return get_model_profile(key)

    def _model_profile_label(self, profile: Dict[str, str]) -> str:
        return f"{profile['key']}: {profile['label']}"

    def _on_style_change(self, _event=None):
        selected = self.style_var.get()
        key = selected.split(":")[0].strip()
        self.model_profile = self._get_model_profile(key)
        self.settings.model_key = self.model_profile["key"]
        self.system_prompt = compose_system_prompt(self.settings.user_profile, model_profile=self.model_profile)
        save_settings(self.settings)
        self._set_status(f"会話スタイルを変更: {self.model_profile['label']}")
        self.model_label_var.set(f"model: {self.model_profile.get('model', '')}")

    def _on_voice_toggle(self):
        self.settings.voice_enabled = self.voice_var.get()
        save_settings(self.settings)
        state = "ON" if self.settings.voice_enabled else "OFF"
        self._set_status(f"音声再生: {state}")

    # ── プロフィール編集 ──────────────────────────────────
    def _edit_profile(self):
        dlg = ProfileDialog(self.root, self.settings.user_profile)
        if dlg.result_profile is not None:
            self.settings.user_profile = dlg.result_profile
            self.system_prompt = compose_system_prompt(self.settings.user_profile, model_profile=self.model_profile)
            save_settings(self.settings)
            name = self.settings.user_profile.display_label()
            self._add_system_message(f"プロフィールを更新しました。呼び名: {name}")
            self._set_status(f"プロフィール更新完了: {name}")

    # ── 詳細設定 ──────────────────────────────────────────
    def _edit_settings(self):
        dlg = SettingsDialog(self.root, self.settings)
        if dlg.result_settings is not None:
            # ユーザープロフィールとmodel_keyは保持
            dlg.result_settings.user_profile = self.settings.user_profile
            dlg.result_settings.model_key = self.settings.model_key
            dlg.result_settings.voice_enabled = self.voice_var.get()
            self.settings = dlg.result_settings
            save_settings(self.settings)
            self._set_status("設定を保存しました。")
            # VOICEVOX 再接続確認
            self._check_voicevox()

    # ── 履歴 ─────────────────────────────────────────────
    def _save_history_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON ファイル", "*.json"), ("テキストファイル", "*.txt")],
            title="会話履歴を保存",
        )
        if not path:
            return
        if path.endswith(".txt"):
            try:
                with open(path, "w", encoding="utf-8") as f:
                    for msg in self.chat_history:
                        role = "きりたん" if msg["role"] == "assistant" else self.settings.user_profile.display_label()
                        f.write(f"{role}: {msg['content']}\n\n")
                self._set_status(f"履歴を保存しました: {path}")
            except Exception as e:
                messagebox.showerror("保存エラー", str(e))
        else:
            if save_chat_history(self.chat_history, Path(path)):
                self._set_status(f"履歴を保存しました: {path}")
            else:
                messagebox.showerror("保存エラー", "保存に失敗しました。")

    def _load_history_from(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON ファイル", "*.json")],
            title="会話履歴を読み込み",
        )
        if not path:
            return
        history = load_chat_history(Path(path))
        if not history:
            messagebox.showinfo("読み込み", "会話履歴がありませんでした。")
            return
        self.chat_history = history
        # チャット表示を更新
        self.chat_text.configure(state="normal")
        self.chat_text.delete("1.0", "end")
        self.chat_text.configure(state="disabled")
        for msg in history:
            if msg["role"] == "user":
                self._add_user_message(msg["content"])
            elif msg["role"] == "assistant":
                self._add_kiritan_message(msg["content"])
        self._set_status(f"履歴を読み込みました: {len(history)} 件")

    def _clear_history(self):
        if messagebox.askyesno("確認", "会話履歴をクリアしますか？"):
            self.chat_history.clear()
            self.chat_text.configure(state="normal")
            self.chat_text.delete("1.0", "end")
            self.chat_text.configure(state="disabled")
            self._add_system_message("会話履歴をクリアしました。")
            self._set_status("履歴をクリアしました。")

    # ── ウィンドウクローズ ────────────────────────────────
    def _on_close(self):
        save_settings(self.settings)
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()


# ── エントリポイント ──────────────────────────────────────
def main():
    # UTF-8 対応
    os.environ.setdefault("PYTHONUTF8", "1")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")

    root = tk.Tk()
    app = KiritanChatGUINew(root)
    app.run()


if __name__ == "__main__":
    main()
