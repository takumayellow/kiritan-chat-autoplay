# -*- coding: utf-8 -*-
"""
kiritan_voicevox.py
VOICEVOX HTTP API 連携モジュール

VOICEVOX エンジン（http://localhost:50021）と HTTP で通信して
東北きりたんの音声を合成・再生する。

依存: requests, sounddevice, soundfile (or simpleaudio)
  pip install requests sounddevice soundfile
"""

import io
import os
import threading
import time
from typing import Optional, Tuple

import requests

# ── VOICEVOX 設定 ─────────────────────────────────────────
VOICEVOX_BASE_URL: str = os.environ.get("VOICEVOX_URL", "http://localhost:50021")

# 東北きりたん の Speaker ID（VOICEVOX デフォルト）
# VOICEVOX で /speakers を叩いて確認できる。
# 東北きりたん: style_id=58 が一般的だが環境により異なるため環境変数で上書き可能。
KIRITAN_SPEAKER_ID: int = int(os.environ.get("VOICEVOX_SPEAKER_ID", "58"))

# 接続タイムアウト（秒）
REQUEST_TIMEOUT: float = float(os.environ.get("VOICEVOX_TIMEOUT", "10.0"))


# ── 接続チェック ──────────────────────────────────────────
def check_connection(base_url: str = VOICEVOX_BASE_URL) -> Tuple[bool, str]:
    """
    VOICEVOX エンジンへの接続を確認する。
    Returns: (接続OK: bool, メッセージ: str)
    """
    try:
        resp = requests.get(f"{base_url}/version", timeout=3.0)
        resp.raise_for_status()
        version = resp.text.strip().strip('"')
        return True, f"VOICEVOX v{version} 接続OK"
    except requests.exceptions.ConnectionError:
        return False, f"VOICEVOX エンジンに接続できません ({base_url})\nエンジンが起動しているか確認してください。"
    except requests.exceptions.Timeout:
        return False, f"VOICEVOX エンジンへの接続がタイムアウトしました ({base_url})"
    except Exception as e:
        return False, f"VOICEVOX 接続エラー: {e}"


def get_speakers(base_url: str = VOICEVOX_BASE_URL) -> list:
    """
    利用可能なスピーカー一覧を取得する。
    Returns: スピーカー情報のリスト（dict）
    """
    try:
        resp = requests.get(f"{base_url}/speakers", timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp.json()
    except Exception:
        return []


def find_kiritan_speaker_id(base_url: str = VOICEVOX_BASE_URL) -> Optional[int]:
    """
    VOICEVOX スピーカー一覧から「東北きりたん」の speaker_id を自動検索。
    見つからなければ None を返す。
    """
    speakers = get_speakers(base_url)
    for speaker in speakers:
        name = speaker.get("name", "")
        if "きりたん" in name or "kiritan" in name.lower():
            styles = speaker.get("styles", [])
            if styles:
                return styles[0].get("id")
    return None


# ── 音声合成 ──────────────────────────────────────────────
def synthesize(
    text: str,
    speaker_id: int = KIRITAN_SPEAKER_ID,
    speed_scale: float = 1.0,
    pitch_scale: float = 0.0,
    intonation_scale: float = 1.0,
    volume_scale: float = 1.0,
    base_url: str = VOICEVOX_BASE_URL,
) -> Optional[bytes]:
    """
    テキストを音声合成して WAV バイト列を返す。
    失敗した場合は None を返す。

    Args:
        text: 読み上げるテキスト
        speaker_id: VOICEVOX のスピーカーID
        speed_scale: 話速（0.5〜2.0、デフォルト 1.0）
        pitch_scale: 音高（-0.15〜0.15、デフォルト 0.0）
        intonation_scale: 抑揚（0.0〜2.0、デフォルト 1.0）
        volume_scale: 音量（0.0〜2.0、デフォルト 1.0）
        base_url: VOICEVOX エンジンの URL
    """
    if not text or not text.strip():
        return None

    try:
        # Step 1: audio_query を作成
        query_resp = requests.post(
            f"{base_url}/audio_query",
            params={"text": text, "speaker": speaker_id},
            timeout=REQUEST_TIMEOUT,
        )
        query_resp.raise_for_status()
        query = query_resp.json()

        # パラメータを上書き
        query["speedScale"] = max(0.5, min(2.0, speed_scale))
        query["pitchScale"] = max(-0.15, min(0.15, pitch_scale))
        query["intonationScale"] = max(0.0, min(2.0, intonation_scale))
        query["volumeScale"] = max(0.0, min(2.0, volume_scale))

        # Step 2: synthesis
        synth_resp = requests.post(
            f"{base_url}/synthesis",
            params={"speaker": speaker_id},
            json=query,
            timeout=REQUEST_TIMEOUT * 3,
        )
        synth_resp.raise_for_status()
        return synth_resp.content

    except requests.exceptions.ConnectionError:
        return None
    except Exception:
        return None


# ── 音声再生 ──────────────────────────────────────────────
def play_wav_bytes(wav_bytes: bytes) -> bool:
    """
    WAV バイト列を再生する。sounddevice + soundfile を優先し、
    失敗した場合は simpleaudio にフォールバック。
    Returns: 再生開始できた場合 True
    """
    if not wav_bytes:
        return False

    # sounddevice + soundfile を試みる
    try:
        import sounddevice as sd
        import soundfile as sf

        buf = io.BytesIO(wav_bytes)
        data, samplerate = sf.read(buf, dtype="float32")
        sd.play(data, samplerate=samplerate)
        sd.wait()
        return True
    except ImportError:
        pass
    except Exception:
        pass

    # simpleaudio にフォールバック
    try:
        import simpleaudio as sa  # type: ignore

        wave_obj = sa.WaveObject.from_wave_file(io.BytesIO(wav_bytes))
        play_obj = wave_obj.play()
        play_obj.wait_done()
        return True
    except ImportError:
        pass
    except Exception:
        pass

    # pygame にフォールバック
    try:
        import pygame  # type: ignore

        pygame.mixer.init()
        pygame.mixer.music.load(io.BytesIO(wav_bytes))
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            time.sleep(0.1)
        return True
    except ImportError:
        pass
    except Exception:
        pass

    return False


def speak(
    text: str,
    speaker_id: int = KIRITAN_SPEAKER_ID,
    speed_scale: float = 1.0,
    pitch_scale: float = 0.0,
    intonation_scale: float = 1.0,
    volume_scale: float = 1.0,
    base_url: str = VOICEVOX_BASE_URL,
    on_complete: Optional[callable] = None,
) -> bool:
    """
    テキストを音声合成して再生する（ブロッキング）。

    Args:
        text: 読み上げるテキスト
        speaker_id: VOICEVOX のスピーカーID
        speed_scale: 話速（0.5〜2.0）
        pitch_scale: 音高（-0.15〜0.15）
        intonation_scale: 抑揚（0.0〜2.0）
        volume_scale: 音量（0.0〜2.0）
        base_url: VOICEVOX エンジンの URL
        on_complete: 再生完了後に呼ばれるコールバック（引数なし）
    Returns:
        成功した場合 True
    """
    wav_bytes = synthesize(
        text,
        speaker_id=speaker_id,
        speed_scale=speed_scale,
        pitch_scale=pitch_scale,
        intonation_scale=intonation_scale,
        volume_scale=volume_scale,
        base_url=base_url,
    )
    if wav_bytes is None:
        return False

    result = play_wav_bytes(wav_bytes)
    if on_complete:
        try:
            on_complete()
        except Exception:
            pass
    return result


def speak_async(
    text: str,
    speaker_id: int = KIRITAN_SPEAKER_ID,
    speed_scale: float = 1.0,
    pitch_scale: float = 0.0,
    intonation_scale: float = 1.0,
    volume_scale: float = 1.0,
    base_url: str = VOICEVOX_BASE_URL,
    on_complete: Optional[callable] = None,
) -> threading.Thread:
    """
    テキストを非同期で音声合成・再生する。
    Returns: 再生スレッド（daemon=True）
    """
    def _worker():
        speak(
            text,
            speaker_id=speaker_id,
            speed_scale=speed_scale,
            pitch_scale=pitch_scale,
            intonation_scale=intonation_scale,
            volume_scale=volume_scale,
            base_url=base_url,
            on_complete=on_complete,
        )

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
    return thread


# ── テキスト前処理 ────────────────────────────────────────
import re as _re

_SANITIZE_TABLE = str.maketrans({
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
    kiritan_chat_cli.py の同名関数と同等の処理。
    """
    cleaned = text.translate(_SANITIZE_TABLE)
    cleaned = _re.sub(r"[•●○◆◇■□※☆★▶▷◀◁]", "・", cleaned)
    return cleaned


# ── 簡易テスト ────────────────────────────────────────────
if __name__ == "__main__":
    print("VOICEVOX 接続テスト...")
    ok, msg = check_connection()
    print(msg)

    if ok:
        # きりたんの speaker_id を自動検索
        sid = find_kiritan_speaker_id()
        if sid is not None:
            print(f"東北きりたん の speaker_id: {sid}")
        else:
            sid = KIRITAN_SPEAKER_ID
            print(f"東北きりたん が見つかりませんでした。デフォルト speaker_id={sid} を使用します。")

        print("テスト音声を再生します...")
        success = speak("こんにちは、東北きりたんです。テスト再生を行います。", speaker_id=sid)
        if success:
            print("再生完了！")
        else:
            print("再生に失敗しました。sounddevice または simpleaudio をインストールしてください。")
