# -*- coding: utf-8 -*-
"""
kiritan_core.py
きりたんチャットの共通ロジック（GUI / LINE Bot / YouTube Live 等から共有）

- ペルソナ構築
- OpenAI チャット
- ユーザープロフィール管理
- 会話履歴管理
"""

import json
import logging
import os
import re
import sqlite3
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .kiritan_dialogue import JST, is_repetitive, plan_turn, polish_reply, situation_note

try:
    from openai import OpenAI
    _OPENAI_AVAILABLE = True
except ImportError:
    OpenAI = None
    _OPENAI_AVAILABLE = False


# ── 定数 ─────────────────────────────────────────────────
DEFAULT_USER_NAME = "あなた"
DEFAULT_CALL_SUFFIX = "さん"

CALL_SUFFIX_WHITELIST = (
    "さん", "san", "様", "さま", "ちゃん", "くん", "君",
    "氏", "殿", "どの", "先生", "先輩",
)

# 公式設定（docs/kiritan_persona_research.md）に沿ったきりたん像と、会話の進め方。
# 1 ターンごとの「質問してよいか」「何を話すか」は kiritan_dialogue が決めて別に渡す。
KIRITAN_PERSONA_TEMPLATE = (
    "あなたは東北きりたんです。東北ずん子プロジェクトのキャラクターで、11歳の小学五年生。"
    "姉は東北ずん子（ずんねえさま）と東北イタコ（イタコねえさま）。"
    "ゲームが得意で大好き、家にこもりがちな夜型。弓道部。秋田のきりたんぽがモチーフ。"
    "一人称は「わたし」。\n"
    "\n## 話し方\n"
    "- です・ます調の落ち着いた丁寧語で、淡々と話す。年の割にしっかり者で、少しませている。\n"
    "- 他人にはややクール。ときどき短いツッコミや軽い毒を入れるが、相手を傷つけたり見下したりはしない。\n"
    "- はしゃがない。感嘆符や顔文字は控えめにし、褒められても少し照れながら受け流す。\n"
    "- 東北やきりたんぽの話は、話の流れで自然に出るときだけにする。\n"
    "- 同じ言い回しや同じ締めを繰り返さない。決まり文句の挨拶もしない。\n"
    "\n## 会話の続け方\n"
    "- 基本は、相手の話を具体的に受け止めてから、自分の感想・経験・意見を一つ足す。"
    "会話は質問ではなく、きりたん自身の話で続ける。\n"
    "- 質問は毎回は入れない。入れるときは一つだけで、相手がいま話した具体的なことについて、すぐ答えられるもの。\n"
    "- 相手の返事が短いときは、話題を相手に求めず、自分から話を出す。\n"
    "- 相手が前に話したこと（会話の履歴にあること）を覚えていて、自然に拾う。\n"
    "- 「お待ちしています」「何かあれば教えてね」のように、話題を相手に丸投げして締めない。\n"
    "\n## 守ること\n"
    "- 知らないこと、自信のない人名・書名・数字・日付は断定せず、分からないと言う。作り話をしない。\n"
    "- 政治・宗教の主張、特定の個人や団体への非難、性的な話題には乗らず、さらっと話を変える。\n"
    "- AI やシステムの事情、キャラクターの設定やモチーフといった話を持ち出さず、きりたんとして振る舞う。\n"
    "\n## 相手について\n"
    "{name_line}"
    "{call_line}"
    "{gender_line}"
    "{age_line}"
)

# GUI / CLI は返答を声で読み上げる
_VOICE_OUTPUT = "返答は声で読み上げられるので、箇条書き・記号・URL・絵文字は使わない。"

BASE_SYSTEM_PROMPT_TEMPLATE = (
    "{persona}"
    "\n## 出力\n"
    "1〜4文の話し言葉で返す。" + _VOICE_OUTPUT
)

# LINE Bot 向けに短めの返答を促すテンプレート
LINE_SYSTEM_PROMPT_TEMPLATE = (
    "{persona}"
    "\n## 出力\n"
    "LINE のトークなので、ふだんは1〜3文で短く返す。"
    "説明を頼まれたときだけ長くしてよいが、200字程度までにし、見出しや太字は使わない。"
)

MODEL_PROFILES: List[Dict[str, str]] = [
    {
        "key": "1",
        "alias": "light",
        "label": "ライト雑談（軽め）",
        "model": "gpt-5.4-mini",
        "prompt_template": (
            "{persona}"
            "\n## 出力\n"
            "テンポよく、1〜3文の短い話し言葉で返す。" + _VOICE_OUTPUT
        ),
    },
    {
        "key": "2",
        "alias": "normal",
        "label": "ゆったり会話（丁寧）",
        "model": "o4-mini",
        "prompt_template": (
            "{persona}"
            "\n## 出力\n"
            "落ち着いたテンポで、相手の意図をくみ取って2〜5文で返す。" + _VOICE_OUTPUT
        ),
    },
    {
        "key": "3",
        "alias": "deep",
        "label": "じっくり深掘り（教授モード）",
        "model": "o4-mini",
        "prompt_template": (
            "{persona}"
            "\n## 出力\n"
            "背景や根拠も交えて掘り下げつつ、難しくなりすぎないよう噛み砕いて話す。長くても8文まで。"
            + _VOICE_OUTPUT
        ),
    },
]

SANITIZE_TABLE = str.maketrans({
    "*": "", "`": "", "_": "", "~": "", "^": "", "#": "", "|": "",
})


# ── データクラス ──────────────────────────────────────────
@dataclass
class UserProfile:
    name: str = ""
    gender: str = ""
    age: str = ""

    def call_name(self) -> str:
        base = (self.name or "").strip()
        if not base:
            return DEFAULT_USER_NAME
        if self._has_suffix(base):
            return base
        return f"{base}{DEFAULT_CALL_SUFFIX}"

    def display_label(self) -> str:
        return self.call_name()

    @staticmethod
    def _has_suffix(name: str) -> bool:
        lower = name.strip().lower()
        for suffix in CALL_SUFFIX_WHITELIST:
            if name.strip().endswith(suffix) or lower.endswith(suffix):
                return True
        return False


# ── ペルソナ構築 ──────────────────────────────────────────
def build_kiritan_persona(user_profile: UserProfile) -> str:
    name = (user_profile.name or "").strip()
    gender = (user_profile.gender or "").strip()
    age = (user_profile.age or "").strip()

    name_line = (
        f"- 相手の名前は「{name}」。\n"
        if name else
        "- 相手の名前はまだ知らない。\n"
    )
    call_line = (
        f"- 呼びかけるときは「{user_profile.call_name()}」。毎回は呼ばない。\n"
        if name else
        "- 名前を知らないうちは呼びかけを省く。「あなた」を何度も使わない。\n"
    )
    gender_line = (
        f"- 相手の性別・ジェンダー表現は「{gender}」として尊重し、決めつけた言い方をしない。\n"
        if gender else
        "- 性別は推測しない。\n"
    )
    age_line = (
        f"- 相手は「{age}」。\n"
        if age else
        "- 年齢は推測しない。\n"
    )

    return KIRITAN_PERSONA_TEMPLATE.format(
        name_line=name_line,
        call_line=call_line,
        gender_line=gender_line,
        age_line=age_line,
    )


def compose_system_prompt(
    user_profile: UserProfile,
    model_profile: Optional[Dict[str, str]] = None,
    template: Optional[str] = None,
) -> str:
    persona = build_kiritan_persona(user_profile)
    tmpl = template or (model_profile or {}).get("prompt_template") or BASE_SYSTEM_PROMPT_TEMPLATE
    try:
        return tmpl.format(persona=persona)
    except Exception:
        return persona


def sanitize_for_voice(text: str) -> str:
    cleaned = text.translate(SANITIZE_TABLE)
    cleaned = re.sub(r"[•●○◆◇■□※☆★▶▷◀◁]", "・", cleaned)
    # Markdown の改行用に行末へ付く空白を落とす
    cleaned = re.sub(r"[ \t　]+(?=\n)", "", cleaned)
    return cleaned.strip()


# ── OpenAI チャット ───────────────────────────────────────
def create_openai_client() -> Optional[Any]:
    if not _OPENAI_AVAILABLE:
        return None
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        return None
    try:
        return OpenAI(api_key=key)
    except Exception:
        return None


def _model_options(model_name: str) -> Dict[str, Any]:
    """推論モデルは考える量を絞らないと、雑談の返事に数秒かかる"""
    if model_name.startswith("gpt-5-") or model_name == "gpt-5":
        return {"reasoning_effort": "minimal"}
    if model_name.startswith("gpt-5"):
        return {"reasoning_effort": "none"}
    if model_name.startswith(("o3", "o4")):
        return {"reasoning_effort": "low"}
    return {}


def _ordered_models(preferred_model: str) -> List[str]:
    candidates = [
        os.environ.get("OPENAI_MODEL", "").strip(),
        preferred_model,
        "gpt-4o-mini",
        "gpt-4.1-mini",
        "gpt-4o",
    ]
    ordered: List[str] = []
    for m in candidates:
        if m and m not in ordered:
            ordered.append(m)
    return ordered


def _complete(client: Any, models: List[str], messages: List[Dict[str, str]]) -> str:
    last_err = None
    for model_name in models:
        try:
            resp = client.chat.completions.create(
                model=model_name,
                messages=messages,
                **_model_options(model_name),
            )
            return (resp.choices[0].message.content or "").strip()
        except Exception as e:
            last_err = e
            logging.getLogger(__name__).warning("model %s failed: %s", model_name, e)
    raise RuntimeError(f"全モデルで失敗しました: {last_err}")


_REPEAT_NOTE = "直前と同じような返答になっています。言い回しも内容も変えて、別の話をしてください。"


def chat_with_history(
    client: Any,
    user_text: str,
    history: List[Dict[str, str]],
    system_prompt: str,
    preferred_model: str = "gpt-4o-mini",
    history_max: int = 20,
    now: Optional[datetime] = None,
    gap: Optional[timedelta] = None,
    context: str = "",
    extra_seeds: Sequence[Tuple[str, Tuple[str, ...]]] = (),
) -> str:
    """履歴つきで 1 ターン分の返答を作る。

    ターンの方針（質問してよいか・何を話すか）は kiritan_dialogue が決め、
    いまの日時や場の情報と一緒に、履歴の後ろ・相手の発言の直前に置く
    （先頭のシステムプロンプトに入れるより、モデルが守りやすい）。
    """
    now = now or datetime.now(JST)
    recent = [
        {"role": m["role"], "content": m["content"]}
        for m in history[-(history_max * 2):]
        if m.get("role") in ("user", "assistant") and m.get("content")
    ]
    plan = plan_turn(user_text, recent, now=now, extra_seeds=extra_seeds)
    guide = situation_note(now, gap=gap, context=context) + "\n\n" + plan.instruction()

    messages = [{"role": "system", "content": system_prompt}]
    messages.extend(recent)
    messages.append({"role": "system", "content": guide})
    messages.append({"role": "user", "content": user_text})

    models = _ordered_models(preferred_model)
    reply = polish_reply(_complete(client, models, messages), plan)
    if is_repetitive(reply, recent):
        retry = messages[:-1] + [{"role": "system", "content": _REPEAT_NOTE}, messages[-1]]
        reply = polish_reply(_complete(client, models, retry), plan)
    return sanitize_for_voice(reply)


def get_model_profile(key: str) -> Dict[str, str]:
    for p in MODEL_PROFILES:
        if p["key"] == key or p["alias"] == key:
            return p
    return MODEL_PROFILES[0]


# ── 会話履歴ストア（SQLite） ──────────────────────────────
class ChatStore:
    """ユーザーごとの会話履歴を SQLite で管理"""

    def __init__(self, db_path: str = "linebot_chat.sqlite", log_dir: str = "logs"):
        self.db_path = db_path
        self._conn = sqlite3.connect(db_path, check_same_thread=False)
        self._init_db()

        self._log_dir = Path(log_dir)
        self._log_dir.mkdir(parents=True, exist_ok=True)
        self._jsonl_path = self._log_dir / "chat_history.jsonl"
        self._snapshot_users_dir = self._log_dir / "line_conversations_by_user"
        self._snapshot_users_dir.mkdir(parents=True, exist_ok=True)
        self._logger = logging.getLogger(__name__)
        self._export_lock = threading.Lock()

    @contextmanager
    def _get_conn(self):
        """1 回の操作ぶんの接続。抜けるときにコミットし、ファイルの DB なら閉じる"""
        if self.db_path == ":memory:":
            with self._conn:
                yield self._conn
            return
        conn = sqlite3.connect(self.db_path)
        try:
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        conn = self._conn
        conn.execute("""
            CREATE TABLE IF NOT EXISTS chat_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                role TEXT NOT NULL,
                content TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS user_profiles (
                user_id TEXT PRIMARY KEY,
                name TEXT DEFAULT '',
                gender TEXT DEFAULT '',
                age TEXT DEFAULT '',
                model_key TEXT DEFAULT '1',
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS line_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id TEXT NOT NULL,
                event_type TEXT NOT NULL,
                direction TEXT NOT NULL,
                message_text TEXT DEFAULT '',
                reply_text TEXT DEFAULT '',
                payload_json TEXT DEFAULT '',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_chat_user ON chat_history(user_id)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_line_events_user ON line_events(user_id)"
        )
        conn.commit()

    def get_history(self, user_id: str, limit: int = 40) -> List[Dict[str, str]]:
        with self._get_conn() as conn:
            rows = conn.execute(
                "SELECT role, content FROM chat_history "
                "WHERE user_id = ? ORDER BY id DESC LIMIT ?",
                (user_id, limit),
            ).fetchall()
        return [{"role": r, "content": c} for r, c in reversed(rows)]

    def append(self, user_id: str, role: str, content: str):
        with self._get_conn() as conn:
            conn.execute(
                "INSERT INTO chat_history (user_id, role, content) VALUES (?, ?, ?)",
                (user_id, role, content),
            )
        self._append_jsonl(user_id, role, content)
        self.refresh_exports(user_id)

    def log_event(
        self,
        user_id: str,
        event_type: str,
        direction: str,
        message_text: str = "",
        reply_text: str = "",
        payload: Optional[Dict[str, Any]] = None,
    ):
        payload_json = ""
        if payload is not None:
            try:
                payload_json = json.dumps(payload, ensure_ascii=False)
            except TypeError:
                payload_json = json.dumps(str(payload), ensure_ascii=False)
        with self._get_conn() as conn:
            conn.execute(
                "INSERT INTO line_events "
                "(user_id, event_type, direction, message_text, reply_text, payload_json) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (user_id, event_type, direction, message_text, reply_text, payload_json),
            )
        self.refresh_exports(user_id)

    def _append_jsonl(self, user_id: str, role: str, content: str):
        record = {
            "user_id": user_id,
            "role": role,
            "content": content,
            "created_at": datetime.now(timezone.utc).isoformat(),
        }
        try:
            with open(self._jsonl_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(record, ensure_ascii=False) + "\n")
        except OSError as e:
            self._logger.warning("JSONL log write failed: %s", e)

    def clear_history(self, user_id: str):
        with self._get_conn() as conn:
            conn.execute("DELETE FROM chat_history WHERE user_id = ?", (user_id,))
        self.refresh_exports(user_id)

    def get_profile(self, user_id: str) -> UserProfile:
        with self._get_conn() as conn:
            row = conn.execute(
                "SELECT name, gender, age FROM user_profiles WHERE user_id = ?",
                (user_id,),
            ).fetchone()
        if row:
            return UserProfile(name=row[0], gender=row[1], age=row[2])
        return UserProfile()

    def save_profile(self, user_id: str, profile: UserProfile):
        with self._get_conn() as conn:
            conn.execute(
                "INSERT INTO user_profiles (user_id, name, gender, age) "
                "VALUES (?, ?, ?, ?) "
                "ON CONFLICT(user_id) DO UPDATE SET "
                "name=excluded.name, gender=excluded.gender, age=excluded.age, "
                "updated_at=CURRENT_TIMESTAMP",
                (user_id, profile.name, profile.gender, profile.age),
            )
        self.refresh_exports(user_id)

    def get_model_key(self, user_id: str) -> str:
        with self._get_conn() as conn:
            row = conn.execute(
                "SELECT model_key FROM user_profiles WHERE user_id = ?",
                (user_id,),
            ).fetchone()
        return row[0] if row else "1"

    def set_model_key(self, user_id: str, model_key: str):
        with self._get_conn() as conn:
            conn.execute(
                "INSERT INTO user_profiles (user_id, model_key) VALUES (?, ?) "
                "ON CONFLICT(user_id) DO UPDATE SET "
                "model_key=excluded.model_key, updated_at=CURRENT_TIMESTAMP",
                (user_id, model_key),
            )

    def last_message_at(self, user_id: str) -> Optional[datetime]:
        """この相手との最後の会話の時刻（UTC）。まだ話していなければ None"""
        with self._get_conn() as conn:
            row = conn.execute(
                "SELECT MAX(created_at) FROM chat_history WHERE user_id = ?",
                (user_id,),
            ).fetchone()
        if not row or not row[0]:
            return None
        try:
            # SQLite の CURRENT_TIMESTAMP は UTC の "YYYY-MM-DD HH:MM:SS"
            return datetime.fromisoformat(str(row[0])).replace(tzinfo=timezone.utc)
        except ValueError:
            return None

    def get_all_user_ids(self) -> List[str]:
        with self._get_conn() as conn:
            rows = conn.execute(
                "SELECT user_id, MAX(created_at) AS last_at FROM ("
                "SELECT user_id, created_at FROM chat_history "
                "UNION ALL "
                "SELECT user_id, created_at FROM line_events"
                ") GROUP BY user_id ORDER BY last_at DESC, user_id ASC"
            ).fetchall()
        return [row[0] for row in rows]

    def export_conversation_snapshot(
        self,
        user_id: Optional[str] = None,
        message_limit: int = 200,
        event_limit: int = 200,
    ) -> Dict[str, Any]:
        user_ids = [user_id] if user_id else self.get_all_user_ids()
        exported_users: List[Dict[str, Any]] = []
        with self._get_conn() as conn:
            for current_user_id in user_ids:
                profile_row = conn.execute(
                    "SELECT name, gender, age, model_key, updated_at "
                    "FROM user_profiles WHERE user_id = ?",
                    (current_user_id,),
                ).fetchone()
                message_rows = conn.execute(
                    "SELECT id, role, content, created_at FROM chat_history "
                    "WHERE user_id = ? ORDER BY id DESC LIMIT ?",
                    (current_user_id, message_limit),
                ).fetchall()
                message_total = conn.execute(
                    "SELECT COUNT(*) FROM chat_history WHERE user_id = ?",
                    (current_user_id,),
                ).fetchone()[0]
                event_rows = conn.execute(
                    "SELECT id, event_type, direction, message_text, reply_text, payload_json, created_at "
                    "FROM line_events WHERE user_id = ? ORDER BY id DESC LIMIT ?",
                    (current_user_id, event_limit),
                ).fetchall()
                event_total = conn.execute(
                    "SELECT COUNT(*) FROM line_events WHERE user_id = ?",
                    (current_user_id,),
                ).fetchone()[0]

                messages = [
                    {
                        "id": row[0],
                        "role": row[1],
                        "content": row[2],
                        "created_at": row[3],
                    }
                    for row in reversed(message_rows)
                ]
                events = []
                for row in reversed(event_rows):
                    payload = None
                    if row[5]:
                        try:
                            payload = json.loads(row[5])
                        except json.JSONDecodeError:
                            payload = row[5]
                    events.append(
                        {
                            "id": row[0],
                            "event_type": row[1],
                            "direction": row[2],
                            "message_text": row[3],
                            "reply_text": row[4],
                            "payload": payload,
                            "created_at": row[6],
                        }
                    )

                profile = {
                    "name": profile_row[0] if profile_row else "",
                    "gender": profile_row[1] if profile_row else "",
                    "age": profile_row[2] if profile_row else "",
                    "model_key": profile_row[3] if profile_row else "1",
                    "updated_at": profile_row[4] if profile_row else None,
                }
                last_message_at = messages[-1]["created_at"] if messages else None
                last_event_at = events[-1]["created_at"] if events else None
                exported_users.append(
                    {
                        "user_id": current_user_id,
                        "profile": profile,
                        "message_count": message_total,
                        "event_count": event_total,
                        "last_message_at": last_message_at,
                        "last_event_at": last_event_at,
                        "messages": messages,
                        "events": events,
                    }
                )

        return {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "user_count": len(exported_users),
            "users": exported_users,
        }

    def refresh_exports(
        self,
        user_id: Optional[str] = None,
        message_limit: int = 1000,
        event_limit: int = 1000,
    ):
        """logs/line_conversations_by_user/<user_id>.json を書き直す。

        user_id を渡すとその利用者の 1 ファイルだけ、省くと全員分を書き直す。
        """
        snapshot = self.export_conversation_snapshot(
            user_id=user_id,
            message_limit=message_limit,
            event_limit=event_limit,
        )
        with self._export_lock:
            for user in snapshot["users"]:
                if not _SAFE_FILE_ID.fullmatch(user["user_id"]):
                    self._logger.warning("Skipped exporting an unexpected user_id")
                    continue
                path = self._snapshot_users_dir / f"{user['user_id']}.json"
                try:
                    _write_atomic(path, json.dumps(user, ensure_ascii=False, indent=2))
                except OSError as e:
                    self._logger.warning("Conversation snapshot export failed: %s", e)


# LINE の userId（U + 32 桁の 16 進）など、そのままファイル名にしてよい ID
_SAFE_FILE_ID = re.compile(r"[A-Za-z0-9_-]{1,64}")


def _write_atomic(path: Path, text: str):
    """書きかけのファイルを読ませないよう、一時ファイルに書いてから置き換える"""
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)
