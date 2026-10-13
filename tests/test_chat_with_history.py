# -*- coding: utf-8 -*-
from datetime import datetime, timedelta, timezone
from types import SimpleNamespace

import pytest

from core import kiritan_core as core
from core.kiritan_dialogue import JST

NOW = datetime(2026, 10, 12, 21, 5, tzinfo=JST)


class FakeClient:
    """chat.completions.create の呼び出しを記録し、決めた返答を順に返す"""

    def __init__(self, replies, fail_models=()):
        self.replies = list(replies)
        self.fail_models = set(fail_models)
        self.calls = []
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, model, messages, **kwargs):
        self.calls.append({"model": model, "messages": messages, "kwargs": kwargs})
        if model in self.fail_models:
            raise RuntimeError(f"{model} is unavailable")
        content = self.replies.pop(0)
        return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))])


@pytest.fixture(autouse=True)
def _no_model_override(monkeypatch):
    monkeypatch.delenv("OPENAI_MODEL", raising=False)


def _chat(client, user_text, history=(), **kwargs):
    return core.chat_with_history(
        client, user_text, list(history), "SYSTEM", preferred_model="gpt-4o-mini", now=NOW, **kwargs
    )


def test_guide_is_placed_right_before_user_message():
    client = FakeClient(["こんばんは。今夜も夜ふかしの予定です。"])
    history = [{"role": "user", "content": "やあ"}, {"role": "assistant", "content": "どうも。"}]
    _chat(client, "こんばんは", history, context="ここは LINE。")

    messages = client.calls[0]["messages"]
    assert messages[0] == {"role": "system", "content": "SYSTEM"}
    assert messages[1:3] == history
    assert messages[3]["role"] == "system"
    assert "【今回の返答の方針】" in messages[3]["content"]
    assert "21時05分" in messages[3]["content"]
    assert "ここは LINE。" in messages[3]["content"]
    assert messages[4] == {"role": "user", "content": "こんばんは"}


def test_reply_is_polished():
    client = FakeClient(["はい、きりたんです。何か話したいことがあれば教えてくださいね。"])
    assert _chat(client, "きりたん") == "はい、きりたんです。"


def test_questions_dropped_after_asking():
    client = FakeClient(["パズルいいですね。どのパズルですか？わたしは落ち物が得意です。"])
    history = [
        {"role": "user", "content": "ゲーム好き"},
        {"role": "assistant", "content": "いいですね。どのゲームが好きですか？"},
    ]
    assert _chat(client, "パズル系", history) == "パズルいいですね。わたしは落ち物が得意です。"


def test_repetitive_reply_is_regenerated_once():
    same = "こんばんは。今日は何をしていましたか？"
    client = FakeClient([same, "こんばんは。わたしはさっきまでゲームをしていました。"])
    history = [{"role": "user", "content": "こんばんは"}, {"role": "assistant", "content": same}]
    reply = _chat(client, "こんばんは", history)
    assert reply == "こんばんは。わたしはさっきまでゲームをしていました。"
    assert len(client.calls) == 2
    assert any("同じような返答" in m["content"] for m in client.calls[1]["messages"] if m["role"] == "system")


def test_overused_phrase_is_regenerated_once():
    history = [
        {"role": "user", "content": "聞こえる？"},
        {"role": "assistant", "content": "聞こえています。夜は静かで落ち着きます。"},
        {"role": "user", "content": "なにしてるの"},
        {"role": "assistant", "content": "ゲームをしています。静かな部屋だと落ち着きますね。"},
    ]
    client = FakeClient(["机を片づけると落ち着きます。", "机を片づけていました。ついでに攻略メモも整理しました。"])
    reply = _chat(client, "そっか", history)
    assert reply == "机を片づけていました。ついでに攻略メモも整理しました。"
    assert len(client.calls) == 2
    assert any("何度も使っています" in m["content"] for m in client.calls[1]["messages"] if m["role"] == "system")


def test_no_regeneration_without_overused_phrase():
    client = FakeClient(["机を片づけていました。"])
    _chat(client, "そっか", [{"role": "user", "content": "やあ"}, {"role": "assistant", "content": "どうも。"}])
    assert len(client.calls) == 1


def test_falls_back_to_next_model():
    client = FakeClient(["ゲームの話ならいくらでもできます。"], fail_models={"o4-mini"})
    core.chat_with_history(client, "ゲーム好き？", [], "SYSTEM", preferred_model="o4-mini", now=NOW)
    assert [c["model"] for c in client.calls] == ["o4-mini", "gpt-4o-mini"]


def test_all_models_fail():
    client = FakeClient([], fail_models={"gpt-4o-mini", "gpt-4.1-mini", "gpt-4o"})
    with pytest.raises(RuntimeError):
        _chat(client, "やあ")


@pytest.mark.parametrize(
    "model, options",
    [
        ("gpt-4o-mini", {}),
        ("gpt-5-mini", {"reasoning_effort": "minimal"}),
        ("gpt-5.4-mini", {"reasoning_effort": "none"}),
        ("o4-mini", {"reasoning_effort": "low"}),
    ],
)
def test_model_options(model, options):
    assert core._model_options(model) == options


def test_history_ignores_extra_fields_and_roles():
    client = FakeClient(["そうですね。"])
    history = [
        {"role": "system", "content": "old"},
        {"role": "user", "content": "a", "created_at": "x"},
        {"role": "assistant", "content": ""},
    ]
    _chat(client, "うん", history)
    sent = client.calls[0]["messages"]
    assert {"role": "user", "content": "a"} in sent
    assert all(m.get("content") != "old" for m in sent)


def test_markdown_line_break_spaces_are_removed():
    client = FakeClient(["はい。  \nゲームをしていました。"])
    assert _chat(client, "なにしてるの") == "はい。\nゲームをしていました。"


def test_persona_has_no_fixed_question_rule():
    prompt = core.compose_system_prompt(core.UserProfile(), template=core.LINE_SYSTEM_PROMPT_TEMPLATE)
    assert "11歳" in prompt
    assert "質問を一つ添えて" not in prompt
    assert "呼びかけを省く" in prompt


def test_chat_store_gap_and_style(tmp_path):
    store = core.ChatStore(db_path=str(tmp_path / "chat.sqlite"), log_dir=str(tmp_path / "logs"))
    assert store.last_message_at("u1") is None
    store.append("u1", "user", "やあ")
    last = store.last_message_at("u1")
    assert last is not None and last.tzinfo == timezone.utc
    assert datetime.now(timezone.utc) - last < timedelta(minutes=5)

    assert store.get_model_key("u1") == "1"
    store.set_model_key("u1", "2")
    assert store.get_model_key("u1") == "2"
    store.save_profile("u1", core.UserProfile(name="テスト"))
    assert store.get_model_key("u1") == "2"
