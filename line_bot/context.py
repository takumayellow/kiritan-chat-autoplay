# -*- coding: utf-8 -*-
"""
LINE のきりたんが知っておくべき場の情報。

モデルは「自分がどこで何のために話しているか」を知らないので、聞かれても答えられず、
作り話をしてしまう。日付で変わる情報（理大祭）は、終わったら出さない。
"""

from datetime import date, datetime
from typing import Tuple

RIDAISAI_DAYS = (date(2026, 11, 20), date(2026, 11, 21))

GAMES = (
    "ITC オープンキャンパス（ITC の活動を講座のように受けられるサイト）",
    "葛飾キャンパスデイズ（葛飾キャンパスを 3D で歩ける探索ゲーム）",
    "おじちゃっと（おじさん構文のキャラとチャットできる）",
    "SynCadence Rhythm（ピアノの名曲で遊ぶリズムゲーム）",
    "LANTERNA パチンコ（お金を賭けないオリジナルのパチンコ）",
    "Cube Trainer（ルービックキューブの次の一手を教えるアプリ）",
)

# きりたんが自分から話す話題に、この場ならではのものを足す
SEEDS: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
    ("メニューの［ゲーム］にある作品のどれか一つ（遊んでみた感想として）", ("ゲーム", "パチンコ", "キャンパス", "リズム", "キューブ")),
)


def line_context(now: datetime) -> str:
    lines = [
        "ここは東京理科大学のサークル ITC の Prog 班が作った LINE のきりたんチャット。",
        "きりたんは小学生なので ITC の部員ではなく、展示や作品づくりの準備もしていない。このチャットで話し相手をしている。",
        "トーク画面下のメニュー［ゲーム］から、Prog 班の作品を遊べる: " + "、".join(GAMES) + "。",
        "作品の中身や遊び方を聞かれたら、この一覧に書いてあることだけを話し、知らない細部は作らない。",
    ]
    today = now.date()
    first, last = RIDAISAI_DAYS
    if today <= last:
        if today < first:
            days = (first - today).days
            lines.append(f"理大祭（東京理科大学の学園祭）は 11月20日（金）・21日（土）で、あと{days}日。ITC も展示をする。")
        else:
            lines.append("今日は理大祭（11月20日・21日）の開催中。ITC の展示に来た人が話しかけていることが多い。")
    return "\n".join(lines)
