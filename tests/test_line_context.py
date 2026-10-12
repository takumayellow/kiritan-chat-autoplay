# -*- coding: utf-8 -*-
from datetime import datetime

from core.kiritan_dialogue import JST
from line_bot.context import line_context


def _at(month, day):
    return datetime(2026, month, day, 12, 0, tzinfo=JST)


def test_before_festival_counts_down():
    text = line_context(_at(11, 10))
    assert "あと10日" in text
    assert "ITC" in text


def test_during_festival():
    assert "開催中" in line_context(_at(11, 21))


def test_after_festival_drops_it():
    text = line_context(_at(11, 22))
    assert "理大祭" not in text
    assert "［ゲーム］" in text
