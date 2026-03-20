# きりたんチャット GUI版（VOICEVOX対応）

東北きりたんと会話できる tkinter ベースの GUI チャットアプリです。
OpenAI API で返答を生成し、VOICEVOX HTTP API できりたんの声で読み上げます。

## ファイル構成

| ファイル | 説明 |
|---|---|
| `kiritan_chat_gui_new.py` | メイン GUI アプリ（本ファイル） |
| `kiritan_voicevox.py` | VOICEVOX HTTP API 連携モジュール |

## 必要環境

- Python 3.9 以上
- [VOICEVOX](https://voicevox.hiroshiba.jp/) エンジン（ローカル起動）
- OpenAI API キー

## インストール

```bash
pip install -r requirements.txt
```

音声再生には以下のいずれかが必要です（優先順）：

```bash
pip install sounddevice soundfile   # 推奨
# または
pip install simpleaudio             # 代替
```

マイク入力を使う場合：

```bash
pip install SpeechRecognition
```

## 起動方法

1. VOICEVOX エンジンを起動する（http://localhost:50021）
2. `.env` ファイルに API キーを設定する：

```
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxx
```

3. アプリを起動する：

```bash
python kiritan_chat_gui_new.py
```

## 画面構成

```
┌─────────────────────────────────────────────────────┐
│  きりたんチャット [VOICEVOX版]   スタイル: [  ▼] 音声□ ● VOICEVOX │
├─────────────────────────────────────────────────────┤
│ ┌─────────────────────────────────────────────────┐ │
│ │                                                 │ │
│ │ [システム] きりたんとの会話を開始しましょう！   │ │
│ │                                                 │ │
│ │ [HH:MM] あなた                                  │ │
│ │  テストです                                     │ │
│ │                                                 │ │
│ │ [HH:MM] きりたん                               │ │
│ │  はい、テストですね！何か試してみますか？       │ │
│ │                                                 │ │
│ └─────────────────────────────────────────────────┘ │
├─────────────────────────────────────────────────────┤
│ ┌─────────────────────────────────────────────────┐ │
│ │ テキスト入力欄（複数行）                        │ │
│ └─────────────────────────────────────────────────┘ │
│  [クリア] [プロフィール]          [マイク 🎤] [送信 ↵] │
├─────────────────────────────────────────────────────┤
│ ステータスバー                       model: gpt-4o-mini │
└─────────────────────────────────────────────────────┘
```

## 機能

### チャット機能
- テキスト入力で送信（Enter キーまたは「送信」ボタン）
- Shift+Enter で改行
- 会話履歴を保持して文脈のある返答を生成
- チャット履歴のスクロール表示（タイムスタンプ付き）

### 音声機能
- VOICEVOX HTTP API で東北きりたんの声で自動読み上げ
- 音声ON/OFF トグル
- マイク入力で音声認識（SpeechRecognition が必要）

### 会話スタイル
| スタイル | モデル | 特徴 |
|---|---|---|
| ライト雑談 | gpt-4o-mini | 軽快なテンポ |
| ゆったり会話 | o4-mini | 丁寧・落ち着き |
| じっくり深掘り | o4-mini-high | 深い解説 |

### プロフィール設定
- 名前・性別・年代感を設定するとペルソナが変化
- きりたんがその人に合った話し方で接してくれる

### 設定
- VOICEVOX URL（デフォルト: http://localhost:50021）
- Speaker ID（自動検索、手動変更も可）
- 話速 / 音高 / 抑揚 / 音量の調整

### 履歴
- 会話履歴を JSON または TXT で保存
- 過去の会話を JSON から読み込み

## 環境変数

| 変数名 | デフォルト | 説明 |
|---|---|---|
| `OPENAI_API_KEY` | （必須） | OpenAI API キー |
| `OPENAI_MODEL` | （自動選択） | 使用するモデルを固定する場合 |
| `VOICEVOX_URL` | `http://localhost:50021` | VOICEVOX エンジンの URL |
| `VOICEVOX_SPEAKER_ID` | `58` | きりたんの Speaker ID（自動検索で上書き） |
| `VOICEVOX_TIMEOUT` | `10.0` | HTTP リクエストタイムアウト（秒） |

## VOICEVOX Speaker ID の確認

VOICEVOX エンジン起動後、以下で利用可能なスピーカーを確認できます：

```bash
curl http://localhost:50021/speakers | python -m json.tool
```

東北きりたんの ID が `58` でない場合、設定画面（メニュー > 設定 > VOICEVOX・音声設定）で変更してください。

## ペルソナについて

`kiritan_chat_cli.py` と同じ「東北きりたんEX」ペルソナを使用しています：

- 14歳、東北ずん子と東北イタコの妹
- 落ち着いた声色、親しみやすい口調
- きりたんぽや東北の話題を好む
- 一人称は「きりたん」または「わたし」

## 既存ファイルとの関係

本ファイルは既存ファイルを変更せず、新規追加のみです：

| 既存ファイル | 関係 |
|---|---|
| `kiritan_chat_cli.py` | ペルソナ・プロンプト設定を参考に実装 |
| `kiritan_chat_gui_app.py` | tkinter GUI の設計を参考に実装 |

音声合成エンジンの違い：
- 既存: VOICEROID＋ 東北きりたん EX（SeikaSay2.exe / pywinauto）
- 本ファイル: VOICEVOX HTTP API（ローカルサーバー）
