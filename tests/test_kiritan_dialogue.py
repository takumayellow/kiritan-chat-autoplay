# -*- coding: utf-8 -*-
from datetime import datetime, timedelta

import pytest

from core.kiritan_dialogue import (
    JST,
    MOVE_ACK,
    MOVE_ANSWER,
    MOVE_FAREWELL,
    MOVE_GREET,
    MOVE_OFFER_TOPIC,
    MOVE_REACT,
    TurnPlan,
    classify_input,
    has_question,
    is_filler,
    is_repetitive,
    overused_phrases,
    pick_topic_seed,
    plan_turn,
    polish_reply,
    situation_note,
    split_sentences,
)

NOW = datetime(2026, 10, 12, 21, 5, tzinfo=JST)


@pytest.mark.parametrize(
    "text, move",
    [
        ("おはよう", MOVE_GREET),
        ("おーい", MOVE_GREET),
        ("やっほー！", MOVE_GREET),
        ("こんばんは〜", MOVE_GREET),
        ("あ", MOVE_OFFER_TOPIC),
        ("はい", MOVE_ACK),
        ("なるほど", MOVE_ACK),
        ("わかる", MOVE_ACK),
        ("わかる？", MOVE_ANSWER),
        ("w", MOVE_ACK),
        ("きりたん", MOVE_OFFER_TOPIC),
        ("", MOVE_OFFER_TOPIC),
        ("おやすみ", MOVE_FAREWELL),
        ("またね！", MOVE_FAREWELL),
        ("そろそろ寝る", MOVE_FAREWELL),
        ("寝る前にゲームしてた", MOVE_REACT),
        ("きりたんの好きな食べ物は", MOVE_ANSWER),
        ("理科大について教えて", MOVE_ANSWER),
        ("今日は理大祭の日？", MOVE_ANSWER),
        ("どう？", MOVE_ANSWER),
        ("マックでバイトしてるの", MOVE_REACT),
        ("昨日は友達とカラオケに行ってきた", MOVE_REACT),
    ],
)
def test_classify_input(text, move):
    assert classify_input(text) == move


def test_split_sentences_keeps_closing_brackets():
    assert split_sentences("「そうなんですか？」と聞きました。次です") == [
        "「そうなんですか？」",
        "と聞きました。",
        "次です",
    ]


def test_has_question():
    assert has_question("いいですね。どれが好きですか？")
    assert has_question("いいですね。どれが好きですか。")
    assert has_question("それ、おいしいのかな。")
    assert has_question("アラームはもうセットしましたか。")
    assert not has_question("いいですね。わたしも好きです。")
    assert not has_question("どれにするか迷っています。")


@pytest.mark.parametrize(
    "sentence",
    [
        "お話を楽しみにしています。",
        "何か話したいことがあれば、気軽に教えてくださいね。",
        "いつでもお待ちしています！",
        "何でも大丈夫ですよ。",
    ],
)
def test_is_filler(sentence):
    assert is_filler(sentence)


def test_is_filler_ignores_content():
    assert not is_filler("きりたんぽ鍋がおいしい季節になりました。")


def test_polish_removes_trailing_filler():
    plan = TurnPlan(move=MOVE_GREET, allow_question=True)
    assert polish_reply("はい、きりたんです！何か話したいことがあれば教えてくださいね。", plan) == "はい、きりたんです！"


def test_polish_removes_questions_when_not_allowed():
    plan = TurnPlan(move=MOVE_REACT, allow_question=False)
    out = polish_reply("カラオケいいですね。何を歌ったんですか？わたしはゲームの曲ばかり歌います。", plan)
    assert out == "カラオケいいですね。わたしはゲームの曲ばかり歌います。"


def test_polish_keeps_only_last_question():
    plan = TurnPlan(move=MOVE_REACT, allow_question=True)
    out = polish_reply("バイトお疲れさまです。何時までですか？忙しいですか？わたしは夜型なので平気です。", plan)
    assert out == "バイトお疲れさまです。忙しいですか？わたしは夜型なので平気です。"


def test_polish_keeps_short_reply_intact():
    plan = TurnPlan(move=MOVE_REACT, allow_question=False)
    # 質問を削ると相づちしか残らないときは削らない
    assert polish_reply("へえ。どこで？", plan) == "へえ。どこで？"


def test_polish_keeps_newlines():
    plan = TurnPlan(move=MOVE_ANSWER, allow_question=True)
    text = "ゲームなら、メニューから遊べます。\nわたしのおすすめはリズムゲームです。"
    assert polish_reply(text, plan) == text


def test_plan_no_question_right_after_question():
    history = [
        {"role": "user", "content": "ゲーム好き"},
        {"role": "assistant", "content": "いいですね。どのゲームが好きですか？"},
    ]
    plan = plan_turn("パズル系", history, now=NOW)
    assert plan.move == MOVE_REACT
    assert plan.allow_question is False
    assert "質問しない" in plan.instruction()


def test_plan_allows_question_after_statement():
    history = [
        {"role": "user", "content": "ゲーム好き"},
        {"role": "assistant", "content": "わたしもです。最近は夜ふかししてばかりです。"},
    ]
    plan = plan_turn("昨日は徹夜した", history, now=NOW)
    assert plan.allow_question is True


def test_plan_suggests_question_after_two_statements():
    history = [
        {"role": "user", "content": "こんばんは"},
        {"role": "assistant", "content": "こんばんは。夜は落ち着きます。"},
        {"role": "user", "content": "暇"},
        {"role": "assistant", "content": "わたしはゲームをしていました。"},
    ]
    plan = plan_turn("今日はバイトだった", history, now=NOW)
    assert plan.suggest_question is True
    assert "具体的な質問を一つだけ入れる" in plan.instruction()


def test_plan_no_suggestion_when_asked_recently():
    history = [
        {"role": "user", "content": "ゲーム好き"},
        {"role": "assistant", "content": "どのゲームですか？"},
        {"role": "user", "content": "パズル"},
        {"role": "assistant", "content": "わたしも好きです。"},
    ]
    plan = plan_turn("最近は毎日やってる", history, now=NOW)
    assert plan.allow_question is True
    assert plan.suggest_question is False


def test_plan_no_suggestion_for_short_input():
    plan = plan_turn("あ", [], now=NOW)
    assert plan.suggest_question is False


def test_plan_farewell_never_asks():
    plan = plan_turn("おやすみ", [], now=NOW)
    assert plan.move == MOVE_FAREWELL
    assert plan.allow_question is False


def test_plan_offer_topic_has_seed():
    plan = plan_turn("あ", [], now=NOW)
    assert plan.move == MOVE_OFFER_TOPIC
    assert plan.topic_seed
    assert plan.topic_seed in plan.instruction()


def test_pick_topic_seed_skips_recent_topics():
    history = [{"role": "assistant", "content": "紅葉がきれいで、鍋がおいしいですね。"}]
    for turn in range(6):
        h = history + [{"role": "user", "content": "x"}] * turn
        assert pick_topic_seed(h, NOW) != "紅葉と、きりたんぽ鍋がおいしくなる季節"


def test_pick_topic_seed_uses_extra_seeds():
    seeds = (("理大祭の展示", ("理大祭",)),)
    picked = {pick_topic_seed([{"role": "user", "content": "x"}] * t, NOW, seeds) for t in range(10)}
    assert "理大祭の展示" in picked


def test_situation_note():
    note = situation_note(NOW, gap=timedelta(days=3), context="ここは LINE。")
    assert "2026年10月12日（月）21時05分、夜" in note
    assert "3日ぶり" in note
    assert note.endswith("ここは LINE。")


def test_situation_note_no_gap_for_short_breaks():
    assert "ぶり" not in situation_note(NOW, gap=timedelta(hours=2))


def test_is_repetitive():
    history = [{"role": "assistant", "content": "はい、きりたんです。何かお話ししたいことはありますか？"}]
    assert is_repetitive("はい、きりたんです。何かお話ししたいことはありますか。", history)
    assert not is_repetitive("きりたんぽ鍋の季節ですね。", history)


def test_plan_ack_deepens_current_topic():
    history = [
        {"role": "user", "content": "鍋食べたい"},
        {"role": "assistant", "content": "寒い日の鍋はいいですね。"},
    ]
    plan = plan_turn("なるほど", history, now=NOW)
    assert plan.move == MOVE_ACK
    assert plan.topic_seed
    assert "一歩だけ深める" in plan.instruction()


def test_plan_ack_without_history_offers_topic():
    plan = plan_turn("うん", [], now=NOW)
    assert plan.move == MOVE_OFFER_TOPIC


def _turns(*pairs):
    history = []
    for user, reply in pairs:
        history += [{"role": "user", "content": user}, {"role": "assistant", "content": reply}]
    return history


def test_overused_phrases_finds_repeated_expression():
    history = _turns(
        ("聞こえる？", "聞こえています。夜は静かで落ち着きます。"),
        ("なにしてるの", "ゲームをしています。静かな部屋だと落ち着きますね。"),
    )
    found = overused_phrases(history, "そっか")
    assert any(p.startswith("落ち着") for p in found)


def test_overused_phrases_ignores_topic_words():
    history = _turns(
        ("秋田行ったことある？", "秋田の冬は長いです。"),
        ("寒そう", "秋田の雪はかなり積もります。"),
    )
    assert overused_phrases(history, "なるほど") == ()


def test_overused_phrases_ignores_context_words():
    history = _turns(
        ("こんばんは", "夜なので眠いです。"),
        ("うん", "夜なのでゲームをしています。"),
    )
    assert overused_phrases(history, "そっか", context="いまは夜。") == ()


def test_overused_phrases_keeps_other_words_when_context_has_date():
    history = _turns(
        ("こんばんは", "今日は静かで落ち着きます。"),
        ("うん", "ゲームの後は静かで落ち着きます。"),
    )
    found = overused_phrases(history, "そっか", context="いまは 2026年10月13日（火）21時00分、夜。")
    assert any(p.startswith("静かで") for p in found)


def test_overused_phrases_prefers_stem_used_more_often():
    history = _turns(
        ("a", "部屋だと落ち着きます。"),
        ("b", "夜は落ち着きます。"),
        ("c", "ゲームをすると落ち着いてきます。"),
        ("d", "雨の音で落ち着いています。"),
    )
    assert "落ち着" in overused_phrases(history, "うん")


def test_overused_phrases_needs_two_replies():
    history = _turns(("やっほー", "静かで落ち着きます。静かで落ち着きます。"))
    assert overused_phrases(history, "うん") == ()


def test_plan_tells_model_to_avoid_overused_phrases():
    history = _turns(
        ("聞こえる？", "聞こえています。夜は静かで落ち着きます。"),
        ("なにしてるの", "ゲームをしています。静かな部屋だと落ち着きますね。"),
    )
    plan = plan_turn("そっか", history, now=NOW)
    assert plan.avoid_phrases
    assert "続けて使っている" in plan.instruction()
    assert "続けて使っている" not in plan_turn("そっか", [], now=NOW).instruction()
