"""
きりたん Chat Autoplay の GUI 版。
CLI で実装済みの会話ロジックを tkinter から呼び出し、
OpenAI → SeikaSay2 → VOICEROID 制御までをアプリ内で完結させる。
"""

import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog
from typing import Optional, Dict, Any

from kiritan_chat_cli import (
    create_client,
    MODEL_PROFILES,
    compose_system_prompt,
    UserProfile,
    DEFAULT_USER_PROFILE,
    start_phrase_tab_sentry,
    stop_phrase_tab_sentry,
    speak,
    chat_once,
    append_profile_history,
    profile_summary,
)


class ProfileDialog(simpledialog.Dialog):
    """ユーザープロフィールの簡易入力ダイアログ。"""

    def __init__(self, parent, profile: UserProfile):
        self.profile = profile
        self.result_data: Optional[Dict[str, str]] = None
        super().__init__(parent, title="プロフィールを編集")

    def body(self, master):
        tk.Label(master, text="名前").grid(row=0, column=0, sticky="w")
        self.name_var = tk.StringVar(value=self.profile.name)
        tk.Entry(master, textvariable=self.name_var).grid(row=0, column=1, sticky="ew")

        tk.Label(master, text="性別・ジェンダー").grid(row=1, column=0, sticky="w")
        self.gender_var = tk.StringVar(value=self.profile.gender)
        tk.Entry(master, textvariable=self.gender_var).grid(row=1, column=1, sticky="ew")

        tk.Label(master, text="年代感").grid(row=2, column=0, sticky="w")
        self.age_var = tk.StringVar(value=self.profile.age)
        tk.Entry(master, textvariable=self.age_var).grid(row=2, column=1, sticky="ew")

        master.grid_columnconfigure(1, weight=1)
        return master

    def apply(self):
        self.result_data = {
            "name": self.name_var.get().strip(),
            "gender": self.gender_var.get().strip(),
            "age": self.age_var.get().strip(),
        }


class KiritanChatGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("きりたん Chat Autoplay - GUI")

        self.chat_log = scrolledtext.ScrolledText(root, state="disabled", width=80, height=24)
        self.chat_log.pack(fill="both", expand=True, padx=8, pady=8)

        control_frame = tk.Frame(root)
        control_frame.pack(fill="x", padx=8)

        tk.Label(control_frame, text="スタイル:").pack(side="left")
        self.style_var = tk.StringVar(value=MODEL_PROFILES[0]["key"])
        self.style_combo = ttk.Combobox(
            control_frame,
            textvariable=self.style_var,
            values=[f"{p['key']}: {p['label']}" for p in MODEL_PROFILES],
            state="readonly",
            width=32,
        )
        self.style_combo.pack(side="left", padx=(4, 8))
        self.style_combo.bind("<<ComboboxSelected>>", self.on_style_change)

        self.speak_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(control_frame, text="音声で読み上げ", variable=self.speak_var).pack(side="left")

        ttk.Button(control_frame, text="プロフィール編集", command=self.edit_profile).pack(side="right", padx=(8, 0))

        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=8, pady=(0, 8))

        self.entry_var = tk.StringVar()
        self.entry = tk.Entry(input_frame, textvariable=self.entry_var)
        self.entry.pack(side="left", fill="x", expand=True)
        self.entry.bind("<Return>", self.on_send)

        self.send_button = ttk.Button(input_frame, text="送信", command=self.on_send)
        self.send_button.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="準備中...")
        ttk.Label(root, textvariable=self.status_var, anchor="w").pack(fill="x", padx=8, pady=(0, 8))

        # プロフィールとモデルの初期化
        self.user_profile = UserProfile(
            name=DEFAULT_USER_PROFILE.name,
            gender=DEFAULT_USER_PROFILE.gender,
            age=DEFAULT_USER_PROFILE.age,
        )
        self.conversation_profile = MODEL_PROFILES[0]
        self.system_prompt = compose_system_prompt(self.conversation_profile, self.user_profile)

        self.client = None
        self.message_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()
        self.root.after(100, self.process_queue)

        self.sentry_stop = None
        self.sentry_thread = None
        self.start_tab_guard()

        self.log("system", "きりたんとの会話を開始できます。スタイルやプロフィールを設定してから話しかけてみてください。")

    def start_tab_guard(self):
        try:
            self.sentry_stop, self.sentry_thread = start_phrase_tab_sentry()
            self.status_var.set("タブ監視を開始しました。")
        except Exception:
            self.status_var.set("タブ監視の開始に失敗しました。")

    def stop_tab_guard(self):
        stop_phrase_tab_sentry(self.sentry_stop, self.sentry_thread)

    def on_style_change(self, _event=None):
        key = self.style_var.get().split(":", 1)[0].strip()
        for profile in MODEL_PROFILES:
            if profile["key"] == key or profile["alias"] == key or profile["label"].startswith(key):
                self.conversation_profile = profile
                break
        self.system_prompt = compose_system_prompt(self.conversation_profile, self.user_profile)
        self.status_var.set(f"会話スタイル: {self.conversation_profile['label']}")

    def edit_profile(self):
        dlg = ProfileDialog(self.root, self.user_profile)
        if dlg.result_data:
            self.user_profile.name = dlg.result_data["name"]
            self.user_profile.gender = dlg.result_data["gender"]
            self.user_profile.age = dlg.result_data["age"]
            append_profile_history(self.user_profile)
            self.system_prompt = compose_system_prompt(self.conversation_profile, self.user_profile)
            self.log("system", f"プロフィール更新: {profile_summary(self.user_profile)}")

    def ensure_client(self) -> bool:
        if self.client:
            return True
        try:
            self.client = create_client()
            self.status_var.set("OpenAI クライアントを初期化しました。")
            return True
        except Exception as e:
            messagebox.showerror("OpenAI エラー", f"クライアントの初期化に失敗しました: {e}")
            self.status_var.set("OpenAI クライアント未初期化")
            return False

    def on_send(self, _event=None):
        text = self.entry_var.get().strip()
        if not text:
            return
        if not self.ensure_client():
            return
        self.entry_var.set("")
        self.log("you", text)
        self.send_button.config(state="disabled")
        self.status_var.set("返信を生成しています...")
        threading.Thread(target=self._generate_reply, args=(text,), daemon=True).start()

    def _generate_reply(self, user_text: str):
        try:
            reply = chat_once(
                self.client,
                user_text=user_text,
                preferred_model=self.conversation_profile["model"],
                system_prompt=self.system_prompt,
            )
            self.message_queue.put({"type": "reply", "text": reply})
            if self.speak_var.get():
                speak(reply)
        except Exception as e:
            self.message_queue.put({"type": "error", "text": str(e)})

    def process_queue(self):
        try:
            while True:
                item = self.message_queue.get_nowait()
                if item["type"] == "reply":
                    self.log("kiritan", item["text"])
                    self.status_var.set(f"モデル: {self.conversation_profile['model']}")
                elif item["type"] == "error":
                    messagebox.showerror("エラー", item["text"])
                    self.status_var.set("エラーが発生しました。")
        except queue.Empty:
            pass
        finally:
            self.send_button.config(state="normal")
            self.root.after(100, self.process_queue)

    def log(self, speaker: str, text: str):
        self.chat_log.config(state="normal")
        self.chat_log.insert(tk.END, f"{speaker}> {text}\n")
        self.chat_log.see(tk.END)
        self.chat_log.config(state="disabled")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.mainloop()

    def on_close(self):
        self.stop_tab_guard()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = KiritanChatGUI(root)
    app.run()


if __name__ == "__main__":
    main()
