# -*- coding: utf-8 -*-
"""
linebot_server.py
きりたんチャット LINE Bot サーバー

起動方法:
  1. .env に以下を設定:
     OPENAI_API_KEY=sk-...
     LINE_CHANNEL_ACCESS_TOKEN=...
     LINE_CHANNEL_SECRET=...
  2. pip install flask line-bot-sdk openai python-dotenv
  3. python linebot_server.py
  4. ngrok http 5000 （開発時）
  5. LINE Developers で Webhook URL を設定
"""

import os
import sys
import logging

# .env 読み込み
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from flask import Flask, request, abort

from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    ApiClient,
    Configuration,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage,
)
from linebot.v3.webhooks import (
    FollowEvent,
    MessageEvent,
    TextMessageContent,
)
from linebot.v3.exceptions import InvalidSignatureError

from core import kiritan_core as core

# ── ログ設定 ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# ── 環境変数 ──────────────────────────────────────────────
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")

if not CHANNEL_ACCESS_TOKEN or not CHANNEL_SECRET:
    logger.warning(
        "LINE_CHANNEL_ACCESS_TOKEN / LINE_CHANNEL_SECRET が未設定です。\n"
        ".env ファイルを確認してください。"
    )

# ── Flask / LINE SDK 初期化 ───────────────────────────────
app = Flask(__name__)

configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# ── きりたんチャット初期化 ────────────────────────────────
openai_client = core.create_openai_client()
# DB パスは環境変数 LINEBOT_DB_PATH で上書き可能（Render ディスクマウント等）
_db_path = os.environ.get("LINEBOT_DB_PATH", "linebot_chat.sqlite")
chat_store = core.ChatStore(db_path=_db_path)

if openai_client:
    logger.info("OpenAI クライアント初期化OK")
else:
    logger.error("OpenAI クライアント初期化失敗。OPENAI_API_KEY を確認してください。")


# ── コマンド処理 ──────────────────────────────────────────
COMMANDS = {
    "/reset": "会話履歴をリセットします。",
    "/profile": "プロフィール設定。例: /profile 名前=たくま 性別=男性 年代=20代",
    "/style": "会話スタイル変更。1=ライト雑談 2=ゆったり会話 3=じっくり深掘り",
    "/help": "コマンド一覧を表示します。",
}


def handle_command(user_id: str, text: str) -> str:
    parts = text.strip().split(maxsplit=1)
    cmd = parts[0].lower()
    arg = parts[1] if len(parts) > 1 else ""

    if cmd == "/help":
        lines = ["きりたんチャット コマンド一覧:"]
        for c, desc in COMMANDS.items():
            lines.append(f"  {c} - {desc}")
        lines.append("\n何も付けずにメッセージを送ると、きりたんと会話できます。")
        return "\n".join(lines)

    if cmd == "/reset":
        chat_store.clear_history(user_id)
        return "会話履歴をリセットしました。新しい気持ちでお話ししましょう。"

    if cmd == "/profile":
        return _handle_profile(user_id, arg)

    if cmd == "/style":
        return _handle_style(user_id, arg)

    return f"不明なコマンドです: {cmd}\n/help でコマンド一覧を確認できます。"


def _handle_profile(user_id: str, arg: str) -> str:
    if not arg:
        profile = chat_store.get_profile(user_id)
        return (
            f"現在のプロフィール:\n"
            f"  名前: {profile.name or '（未設定）'}\n"
            f"  性別: {profile.gender or '（未設定）'}\n"
            f"  年代: {profile.age or '（未設定）'}\n\n"
            f"変更例: /profile 名前=たくま 性別=男性 年代=20代"
        )
    profile = chat_store.get_profile(user_id)
    for pair in arg.split():
        if "=" in pair:
            k, v = pair.split("=", 1)
            k = k.strip()
            if k in ("名前", "name"):
                profile.name = v.strip()
            elif k in ("性別", "gender"):
                profile.gender = v.strip()
            elif k in ("年代", "年齢", "age"):
                profile.age = v.strip()
    chat_store.save_profile(user_id, profile)
    return (
        f"プロフィールを更新しました。\n"
        f"  名前: {profile.name or '（未設定）'}\n"
        f"  性別: {profile.gender or '（未設定）'}\n"
        f"  年代: {profile.age or '（未設定）'}\n\n"
        f"きりたん: {profile.call_name()}、よろしくね。"
    )


def _handle_style(user_id: str, arg: str) -> str:
    if not arg:
        lines = ["会話スタイル一覧:"]
        for p in core.MODEL_PROFILES:
            lines.append(f"  {p['key']}: {p['label']} (model: {p['model']})")
        lines.append("\n例: /style 1")
        return "\n".join(lines)

    profile = core.get_model_profile(arg.strip())
    return f"会話スタイルを「{profile['label']}」に変更しました。"


# ── チャット応答 ──────────────────────────────────────────
def generate_reply(user_id: str, user_text: str) -> str:
    if not openai_client:
        return "ごめんなさい、今ちょっと調子が悪いみたい…。しばらくしてからまた話しかけてね。"

    user_profile = chat_store.get_profile(user_id)
    model_key = chat_store.get_model_key(user_id)
    model_profile = core.get_model_profile(model_key)

    system_prompt = core.compose_system_prompt(
        user_profile=user_profile,
        template=core.LINE_SYSTEM_PROMPT_TEMPLATE,
    )

    history = chat_store.get_history(user_id)

    try:
        reply = core.chat_with_history(
            client=openai_client,
            user_text=user_text,
            history=history,
            system_prompt=system_prompt,
            preferred_model=model_profile["model"],
        )
    except Exception as e:
        logger.error(f"Chat error for user {user_id}: {e}")
        return "うーん、ちょっとうまく言葉が出てこないみたい…。もう一度話しかけてもらえる？"

    chat_store.append(user_id, "user", user_text)
    chat_store.append(user_id, "assistant", reply)

    return reply


# ── Webhook ───────────────────────────────────────────────
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    logger.info(f"Webhook received: {body[:200]}")

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        logger.warning("Invalid signature")
        abort(400)
    except Exception as e:
        logger.error(f"Webhook handler error: {e}")
        # SDK パースエラーでも 200 を返す（LINE側のリトライを防ぐ）
        # 実際のメッセージ処理はフォールバックで行う
        try:
            import json as _json
            data = _json.loads(body)
            for event in data.get("events", []):
                if event.get("type") == "message" and event.get("message", {}).get("type") == "text":
                    _handle_message_fallback(event)
                elif event.get("type") == "follow":
                    _handle_follow_fallback(event)
        except Exception as e2:
            logger.error(f"Fallback handler error: {e2}")

    return "OK"


def _handle_message_fallback(event: dict):
    """SDK パースに失敗した場合のフォールバック処理"""
    user_id = event.get("source", {}).get("userId", "")
    text = event.get("message", {}).get("text", "").strip()
    reply_token = event.get("replyToken", "")
    if not user_id or not text or not reply_token:
        return
    logger.info(f"Fallback message from {user_id}: {text[:100]}")

    if text.startswith("/"):
        reply_text = handle_command(user_id, text)
    else:
        reply_text = generate_reply(user_id, text)

    with ApiClient(configuration) as api_client:
        from linebot.v3.messaging import PushMessageRequest
        api = MessagingApi(api_client)
        api.push_message(
            PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=reply_text)],
            )
        )
    logger.info(f"Fallback push sent to {user_id}")


def _handle_follow_fallback(event: dict):
    """フォローイベントのフォールバック"""
    reply_token = event.get("replyToken", "")
    if not reply_token:
        return
    welcome = (
        "はじめまして、東北きりたんです。\n"
        "わたしとおしゃべりしましょう！\n\n"
        "メッセージを送ってくれたら、お返事するね。\n"
        "/help でコマンド一覧も見れるよ。"
    )
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(
            ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=welcome)],
            )
        )


@handler.add(FollowEvent)
def handle_follow(event: FollowEvent):
    user_id = event.source.user_id
    logger.info(f"New follower: {user_id}")

    welcome = (
        "はじめまして、東北きりたんです。\n"
        "わたしとおしゃべりしましょう！\n\n"
        "メッセージを送ってくれたら、お返事するね。\n"
        "/help でコマンド一覧も見れるよ。"
    )

    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=welcome)],
            )
        )


@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event: MessageEvent):
    user_id = event.source.user_id
    text = event.message.text.strip()
    logger.info(f"Message from {user_id}: {text[:100]}")

    if text.startswith("/"):
        reply_text = handle_command(user_id, text)
    else:
        reply_text = generate_reply(user_id, text)

    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=reply_text)],
            )
        )
    logger.info(f"Reply sent to {user_id}")


# ── ヘルスチェック ────────────────────────────────────────
@app.route("/", methods=["GET"])
def health():
    return "きりたんチャット LINE Bot is running."


# ── エントリポイント ──────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.info(f"Starting LINE Bot server on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)
