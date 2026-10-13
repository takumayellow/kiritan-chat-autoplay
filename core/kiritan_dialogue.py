# -*- coding: utf-8 -*-
"""
kiritan_dialogue.py
1 ターンごとの会話の進め方を決め、生成後の返答を整える。

プロンプトに「質問しすぎないで」と書くだけでは、モデルは毎回の返答を質問で締めてしまう。
そこで、直前までの会話から「今回は何をするターンか」と「質問してよいか」を
コード側で決め、その方針だけをモデルに渡す。生成後も、話題を相手に丸投げする
締めの定型句と、方針に反した質問を取り除く。

docs/conversation_design.md に設計の理由とログの分析がある。
"""

import difflib
import re
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Sequence, Tuple

JST = timezone(timedelta(hours=9))

# ── 入力の分類 ────────────────────────────────────────────
MOVE_ANSWER = "answer"          # 相手の質問に答える
MOVE_REACT = "react"            # 相手の話を受け止め、自分の話を足す
MOVE_ACK = "ack"                # 相づちを受けて、直前の話題を一歩深める
MOVE_GREET = "greet"            # 挨拶を返す
MOVE_OFFER_TOPIC = "offer_topic"  # 相手の発言が短いので、きりたんから話題を出す
MOVE_FAREWELL = "farewell"      # 見送る

_TRAILING_NOISE = re.compile(r"[\s。、．，！!？?〜~ーｰw笑♪☆★…・]+$")

_GREETING = re.compile(
    r"^(おはよ[うー]?(ございます)?|こんにち[はわ]|こんばん[はわ]|やっほ[ーう]?|"
    r"はじめまして|ハロー|はろー|hello|hi|hey|おーい|おい|よっ?|やあ|ども|どうも|"
    r"ただいま|お疲れ(様|さま)?(です)?|おつかれ(さま)?|おひさ(しぶり)?|久しぶり)$",
    re.IGNORECASE,
)
_FAREWELL = re.compile(
    r"(おやすみ|またね|ばいばい|バイバイ|じゃあね|じゃね|さようなら|さよなら|"
    r"寝るね|寝ます|寝る[わよ]?$|落ちるね|またあとで|また明日)"
)
# 相づち。直前の話を聞いていることは伝わるが、話を進める材料は無い
_ACK = {
    "はい", "うん", "ううん", "ええ", "そう", "そうだね", "そうなんだ", "そっか", "へー", "へえ",
    "ふーん", "ふむ", "なるほど", "たしかに", "確かに", "わかった", "了解", "りょ", "ok", "おk",
    "うんうん", "いいね", "草", "笑", "w", "ほう", "まあ", "わかる", "分かる", "わかります",
    "分かります", "それな", "すごい", "すごいね", "えー", "まじか", "マジか",
}
# 呼びかけや一文字だけの発言。話題が無い
_LOW_CONTENT = {"あ", "え", "お", "ん", "きりたん", "きりたんちゃん", "きりちゃん", "なに", "なんでもない"}
_QUESTION_TAIL = re.compile(
    r"([？?]|教えて(ください|ほしい|欲しい)?|知ってる|知ってます|わかる|分かる|"
    r"(です|ます|でし|まし)か|かな|かね|って|なに|何|どう|どれ|どこ|いつ|誰|だれ|なんで|なぜ)$"
)
_QUESTION_WORD = re.compile(r"(何|なに|なん|どう|どこ|どれ|どんな|いつ|誰|だれ|なぜ|なんで|どうして|教えて)")


def normalize_input(text: str) -> str:
    return _TRAILING_NOISE.sub("", (text or "").strip())


def classify_input(text: str) -> str:
    """相手の発言から、今回のターンの基本の動きを決める"""
    raw = (text or "").strip()
    norm = normalize_input(raw)
    lowered = norm.lower()
    if not norm:
        # 「w」「笑」だけの返事は笑っているという相づち
        return MOVE_ACK if re.fullmatch(r"[wWｗ笑草]+", raw) else MOVE_OFFER_TOPIC
    if _FAREWELL.search(norm) and len(norm) <= 20:
        return MOVE_FAREWELL
    if _GREETING.match(lowered):
        return MOVE_GREET
    if lowered in _ACK and not re.search(r"[？?]\s*$", raw):
        return MOVE_ACK
    if lowered in _LOW_CONTENT or len(norm) <= 1:
        return MOVE_OFFER_TOPIC
    if re.search(r"[？?]\s*$", raw) or _QUESTION_TAIL.search(norm):
        return MOVE_ANSWER
    # 「きりたんの好きな食べ物は」のように「は」で止める短い問いかけ
    if len(norm) <= 25 and (norm.endswith("は") or _QUESTION_WORD.search(norm)):
        return MOVE_ANSWER
    return MOVE_REACT


# ── 返答の文単位の扱い ─────────────────────────────────────
# 文の区切り。区切りの後ろの空白や改行も、その文に含めて持つ（削った後に元の改行を保つため）
_SENTENCE = re.compile(r"[^。！？!?\n]*[。！？!?]+[」』）)]*[ \t　]*\n*|[^。！？!?\n]+\n*|\n+")


def _spans(text: str) -> List[str]:
    return [m.group(0) for m in _SENTENCE.finditer(text or "") if m.group(0)]


def split_sentences(text: str) -> List[str]:
    return [s.strip() for s in _spans(text) if s.strip()]


# 「？」で終わる文に加え、「〜ですか。」「〜かな。」のように句点で終わる問いかけも質問とみなす
_QUESTION_END = re.compile(
    r"([？?]|(です|ます|でした|ました|ません|でしょう|のです|んです|ある|いる|する|した)か[。．]?|かな[。．]?)[」』）)]*$"
)


def is_question(sentence: str) -> bool:
    return bool(_QUESTION_END.search(sentence.strip()))


def has_question(text: str) -> bool:
    return any(is_question(s) for s in split_sentences(text))


# 相手に話題を投げ返すだけで中身の無い締め（ログで 49 返答中 25 回出ていた）
_FILLER = [
    re.compile(r"(お待ち|待って|楽しみに)(して)?(い)?ます"),
    re.compile(r"(話|聞|お話|教え).{0,12}(楽しみ|嬉しい|うれしい)です"),
    re.compile(r"(あれば|あったら|なら|ときは|時は).{0,16}(教えて|話して|お話し|聞かせて|どうぞ|声をかけて)"),
    re.compile(r"(何でも|なんでも|どんなことでも).{0,10}(大丈夫|いい|どうぞ)"),
    re.compile(r"いつでも.{0,10}(聞いて|待って|お話し|どうぞ)"),
    re.compile(r"^(気軽に|お気軽に|遠慮なく).{0,20}$"),
]


def is_filler(sentence: str) -> bool:
    return any(p.search(sentence) for p in _FILLER)


# ── 話題の種 ──────────────────────────────────────────────
# (話題, 最近話したかを調べる語)。きりたんが自分から話すときに使う。
TOPIC_SEEDS: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
    ("最近遊んでいるゲームのこと", ("ゲーム", "プレイ", "ランク")),
    ("ずんねえさまやイタコねえさまとのちょっとした出来事", ("ねえさま", "ずん子", "イタコ")),
    ("弓道部の練習", ("弓道", "部活", "練習")),
    ("夜ふかしや休みの日の過ごし方", ("夜ふかし", "夜更かし", "休み")),
    ("きりたんぽや秋田の食べ物", ("きりたんぽ", "秋田", "比内")),
)

_SEASON_SEEDS: Dict[int, Tuple[str, Tuple[str, ...]]] = {
    1: ("寒さとこたつ", ("こたつ", "寒")),
    2: ("雪の日の過ごし方", ("雪",)),
    3: ("春が近づいてきたこと", ("春",)),
    4: ("桜や新学期", ("桜", "新学期")),
    5: ("過ごしやすい季節の外出", ("散歩", "お出かけ")),
    6: ("梅雨の雨の日の過ごし方", ("雨", "梅雨")),
    7: ("夏の暑さ", ("暑", "夏")),
    8: ("夏休みやお祭り", ("夏休み", "祭")),
    9: ("秋の夜長", ("夜長", "秋")),
    10: ("紅葉と、きりたんぽ鍋がおいしくなる季節", ("紅葉", "鍋")),
    11: ("朝晩の冷え込みと温かい食べ物", ("冷え", "温か", "鍋")),
    12: ("冬支度や年末", ("年末", "冬")),
}


def pick_topic_seed(
    history: Sequence[Dict[str, str]],
    now: datetime,
    extra_seeds: Sequence[Tuple[str, Tuple[str, ...]]] = (),
    lookback: int = 6,
) -> Optional[str]:
    """最近の自分の返答で触れていない話題を、ターンごとに順番をずらして選ぶ"""
    seeds = list(extra_seeds) + [_SEASON_SEEDS[now.month]] + list(TOPIC_SEEDS)
    recent = " ".join(
        m.get("content", "") for m in history[-lookback * 2:] if m.get("role") == "assistant"
    )
    fresh = [label for label, words in seeds if not any(w in recent for w in words)]
    if not fresh:
        return None
    turn = sum(1 for m in history if m.get("role") == "user")
    return fresh[turn % len(fresh)]


# ── 時間と場の情報 ─────────────────────────────────────────
def time_of_day(now: datetime) -> str:
    h = now.hour
    if 4 <= h < 10:
        return "朝"
    if 10 <= h < 16:
        return "昼"
    if 16 <= h < 19:
        return "夕方"
    if 19 <= h < 24:
        return "夜"
    return "深夜"


_WEEKDAYS = "月火水木金土日"


def situation_note(now: datetime, gap: Optional[timedelta] = None, context: str = "") -> str:
    """いまの日時・前回からの間隔・場の情報。モデルが自分では知り得ないもの"""
    lines = [
        f"いまは {now.year}年{now.month}月{now.day}日（{_WEEKDAYS[now.weekday()]}）"
        f"{now.hour}時{now.minute:02d}分、{time_of_day(now)}。"
    ]
    if gap is not None and gap >= timedelta(hours=12):
        days = gap.days
        span = f"{days}日" if days >= 1 else "半日"
        lines.append(f"この相手と前に話してから{span}ぶり。")
    if context.strip():
        lines.append(context.strip())
    return "\n".join(lines)


# ── ターンの方針 ──────────────────────────────────────────
@dataclass(frozen=True)
class TurnPlan:
    move: str
    allow_question: bool
    topic_seed: Optional[str] = None
    suggest_question: bool = False  # 相手の話を受けるターンで、しばらく質問していないとき
    avoid_phrases: Tuple[str, ...] = ()  # 最近の返答で続けて使っている言い回し（口癖）

    def instruction(self) -> str:
        lines = ["【今回の返答の方針】", _MOVE_TEXT[self.move].format(seed=self.topic_seed or "自分の最近のこと")]
        if self.move == MOVE_FAREWELL:
            lines.append("引き止めず、質問もしない。")
        elif self.suggest_question:
            lines.append(
                "しばらく聞き返していないので、自分の話を足したうえで、最後に相手がいま話したことについての"
                "具体的な質問を一つだけ入れる（一言で答えられるもの）。自然に聞けることが無いときだけ省く。" + _VAGUE
            )
        elif self.allow_question:
            lines.append(
                "質問は入れなくてよい。入れるなら最後に一つだけ、相手がいま話した具体的なことについて、"
                "一言で答えられるもの（はい/いいえや二択）にする。" + _VAGUE
            )
        else:
            lines.append("今回は質問しない。疑問文を使わず、感想や自分の話で終える。")
        if self.avoid_phrases:
            quoted = "".join(f"「{p}」" for p in self.avoid_phrases)
            lines.append(f"最近の返答で{quoted}を続けて使っている。今回は使わず、別の言い方や別の話にする。")
        lines.append(
            "「お待ちしています」「何かあれば教えてね」「お話を楽しみにしています」のように"
            "話題を相手に丸投げする締めは使わない。"
        )
        return "\n".join(lines)


_VAGUE = "「今日はどんな一日？」「何か話したいことは？」のような大きくて漠然とした質問はしない。"

_MOVE_TEXT = {
    MOVE_ANSWER: "相手の質問にまず答える。知らないことや自信のない固有名詞・数字は断定せず、分からないと言う。",
    MOVE_REACT: "相手の話を具体的に受け止め、それに対するきりたん自身の感想・経験・意見を一つ足す。",
    MOVE_ACK: (
        "相手は相づちを返している。直前の話題を、きりたん自身の具体的な体験や意見で一歩だけ深める。"
        "前の返答の言い直しはしない。続ける材料が無ければ、前の話からつなげて「{seed}」の話に移る（唐突に変えない）。"
    ),
    MOVE_GREET: "挨拶を短く返し、いまの時間帯や季節に合った自分の近況か小ネタを一つ話す。",
    MOVE_OFFER_TOPIC: (
        "相手の発言は短く、話題が無い。話題を求めて相手に投げ返さず、"
        "きりたんの方から「{seed}」の話を、自分の体験や意見として短く一つ話す。"
    ),
    MOVE_FAREWELL: "短く見送る。",
}


_KANJI = "一-鿿々"
_KANA_RUN = re.compile(f"[{_KANJI}ぁ-ゟ゠-ヿー]+")
# 漢字とかなが混ざっていても、言い回しというより普通の言葉なので口癖として扱わない
_COMMON_WORDS = ("思い", "思う", "言い", "言う", "感じ", "気が", "少し", "自分", "好き", "分か", "話し", "一番", "今日", "本当")


def _phrase_grams(text: str, shortest: int = 3, longest: int = 6) -> set:
    """漢字で始まり、かなを含む 3〜6 字の断片（「落ち着き」「静かな」「悪くない」の形）"""
    grams = set()
    for run in _KANA_RUN.findall(text):
        for n in range(shortest, longest + 1):
            for i in range(len(run) - n + 1):
                g = run[i:i + n]
                if re.match(f"[{_KANJI}]", g) and re.search("[ぁ-ゟ]", g):
                    grams.add(g)
    return grams


def overused_phrases(
    history: Sequence[Dict[str, str]],
    user_text: str = "",
    context: str = "",
    lookback: int = 4,
    limit: int = 3,
) -> Tuple[str, ...]:
    """直近の自分の返答のうち 2 回以上で使った言い回し。

    先頭の漢字の語が相手の発言や場の情報に出てくるものは、口癖ではなく今の話題の言葉なので除く。
    """
    recent = [m.get("content", "") for m in history if m.get("role") == "assistant"][-lookback:]
    topic = user_text + context + "".join(
        m.get("content", "") for m in list(history)[-lookback * 2:] if m.get("role") == "user"
    )
    counts: Dict[str, int] = {}
    for reply in recent:
        for g in _phrase_grams(reply):
            counts[g] = counts.get(g, 0) + 1
    found = [
        g for g, c in counts.items()
        if c >= 2
        and not any(w in g for w in _COMMON_WORDS)
        and re.match(f"[{_KANJI}]+", g).group() not in topic
    ]
    # 同じ回数なら長いほうを残す。短いほうが多く使われていれば、活用の違いも含められる短いほうを残す
    found = [g for g in found if not any(g != o and g in o and counts[o] >= counts[g] for o in found)]
    found = [g for g in found if not any(g != o and o in g for o in found)]
    found.sort(key=lambda g: (-counts[g], -len(g), g))
    return tuple(found[:limit])


def plan_turn(
    user_text: str,
    history: Sequence[Dict[str, str]],
    now: Optional[datetime] = None,
    extra_seeds: Sequence[Tuple[str, Tuple[str, ...]]] = (),
    context: str = "",
) -> TurnPlan:
    now = now or datetime.now(JST)
    move = classify_input(user_text)
    replies = [m.get("content", "") for m in history if m.get("role") == "assistant"]
    if move == MOVE_ACK and not replies:
        move = MOVE_OFFER_TOPIC  # 深める話題がまだ無い

    # 質問した次のターンは質問しない。続けて聞かれると、答えるだけの会話になる
    allow_question = move != MOVE_FAREWELL and not (replies and has_question(replies[-1]))
    # 相手が自分のことを話してくれたときは、2 回続けて質問していなければ聞き返してよいと伝える。
    # 「入れなくてよい」だけだと、モデルによっては一度も聞き返さなくなる
    suggest_question = (
        allow_question and move == MOVE_REACT and not any(has_question(r) for r in replies[-2:])
    )

    seed = None
    if move in (MOVE_OFFER_TOPIC, MOVE_ACK, MOVE_GREET):
        seed = pick_topic_seed(history, now, extra_seeds)
    return TurnPlan(
        move=move,
        allow_question=allow_question,
        topic_seed=seed,
        suggest_question=suggest_question,
        avoid_phrases=overused_phrases(history, user_text, context),
    )


# ── 生成後の整え ──────────────────────────────────────────
_MIN_KEEP = 12  # 質問を削った後にこれより短くなるなら削らない（返答が相づちだけになる）
_MIN_REST = 4   # 締めの定型句は、前に「はい、きりたんです！」程度でも残れば削る


def polish_reply(text: str, plan: TurnPlan) -> str:
    """丸投げの締めと、方針に反した質問を取り除く。中身が残らないときは元のまま返す"""
    original = (text or "").strip()
    spans = [sp for sp in _spans(original) if sp.strip()]
    if not spans:
        return original

    def length(items: List[str]) -> int:
        return len("".join(sp.strip() for sp in items))

    kept = list(spans)
    while len(kept) > 1 and is_filler(kept[-1].strip()) and length(kept[:-1]) >= _MIN_REST:
        kept.pop()

    questions = [i for i, sp in enumerate(kept) if is_question(sp)]
    # 質問しないターンでは全部、するターンでも最後の一つ以外は落とす（質問攻めにしない）
    drop = set(questions if not plan.allow_question else questions[:-1])
    if drop:
        trimmed = [sp for i, sp in enumerate(kept) if i not in drop]
        if length(trimmed) >= _MIN_KEEP:
            kept = trimmed

    return "".join(kept).strip()


def is_repetitive(reply: str, history: Sequence[Dict[str, str]], lookback: int = 5, threshold: float = 0.72) -> bool:
    """最近の自分の返答とほぼ同じ文面か"""
    recent = [m.get("content", "") for m in history if m.get("role") == "assistant"][-lookback:]
    return any(difflib.SequenceMatcher(None, reply, r).ratio() >= threshold for r in recent)
