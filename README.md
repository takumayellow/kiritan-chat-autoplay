# きりたん Chat Autoplay (CLI版 / GUI版)

VOICEROID＋ 東北きりたん EX と OpenAI API を組み合わせて、テキスト／音声の往復会話を自動化するためのツールセットです。SeikaSay2.exe をバックエンドに据え、PowerShell（CLI版）から自然に会話を続けられるようにしています。

## 概要
- **会話生成**: OpenAI API（`OPENAI_MODEL` が未設定なら `gpt-4o-mini` などにフォールバック）で返答を取得。
- **読み上げ**: AssistantSeika 同梱の `SeikaSay2.exe` CLI で VOICEROID を起動せずに再生。速度 (`speed`)、話者 (`cid`) を直接指定できます。
- **UI 制御**: pywinauto（UIA）で VOICEROID の「フレーズ編集」タブを監視し、音声効果タブへ飛んでも自動で戻す watchdog を常駐。
- **入力モード**: `mode text / mic / loop / dual` を切り替え可能。`time N` で録音秒数、`speed X` で再生速度を調整。
- **プロフィール記録**: 名前／ジェンダー／年代をヒアリングし、`logs/profile_history.jsonl` に JSONL で保存。 honorific は名前の末尾に含める方式に一本化。
- **フォーカス保護**: 既定では PowerShell のフォーカスを奪わず、VOICEROID 画面を勝手に中央へ移動させません（`KIRITAN_ALLOW_WINDOW_FOCUS=1` で従来挙動に戻せます）。

## 最近のアップデート
- **ペルソナ再構築**: 8 月時点の「親しみやすい口調＋質問で締める」スタイルをベースに、落ち着いたトーンでも自然に質問を添えるプロンプトへ更新。毎返答の末尾に 1 つだけ質問を足すよう System Prompt を強化。
- **呼称管理の自動化**: honorific 入力欄を廃止し、名前に接尾辞が含まれていない場合は自動で「さん」を付ける仕様に変更。履歴にも自動整形後の呼び名を保持。
- **タブガードの堅牢化**: UIA パターン取得を安全化（`iface_invoke` 取得失敗でも落ちない）、`SeikaSay2` 完了後に追加で `ensure_phrase_tab_with_retry` を走らせるなど、音声効果タブからの復帰精度を向上。
- **ウィンドウフォーカス制御**: `KIRITAN_ALLOW_WINDOW_FOCUS` で VOICEROID の前面化を明示的に許可するまで一切フォーカスを奪わない設計に変更。誤って PowerShell が背面に飛ぶ問題を解消。

## 既知の課題
- **音声効果タブへ自動遷移する仕様**: SeikaSay2.exe は毎回「音声効果」パラメータを更新・リセットするため、VOICEROID 本体が再生完了直後に音声効果タブを前面に戻してしまいます。根本的に止める術はなく、監視スレッド `_guard_phrase_tab` で検知→即フレーズ編集へ復帰するアプローチを採用しています。
- **GUI 直接操作との二者択一**: CLI での安定再生を優先した結果、VOICEROID の画面に触っている最中はタブが揺れ動きます。画面自体を見せたくない場合は GUI を閉じた状態で CLI のみを表示する運用を推奨します。

## 使い方

### 必要環境
- Windows 10/11（VOICEROID＋ 東北きりたん EX がインストール済み）
- Python 3.11 以上（`pywinauto`, `pywin32`, `openai`, `speechrecognition`, `sounddevice` などを `pip install -r requirements.txt` で導入）
- OpenAI API Key（`OPENAI_API_KEY`）
- AssistantSeika / SeikaSay2.exe（`SEIKA_EXE` でパスを上書き可能）

### 実行手順
1. 仮想環境を有効化し依存パッケージをインストールします。
   `pip install -r requirements.txt`
2. 必要な環境変数を設定します。
   - `OPENAI_API_KEY`（必須）
   - `OPENAI_MODEL`（任意。未設定なら `gpt-4o-mini`→`o4-mini`…の順にフォールバック）
   - `SEIKA_EXE`（任意。SeikaSay2.exe のフルパスを上書きしたい場合）
   - `KIRITAN_ALLOW_WINDOW_FOCUS`（任意。`1` で VOICEROID を前面化・クリック可能）
   - `KIRITAN_TAB_DEBUG`（任意。`1` でタブ監視ログを PowerShell に出力）
3. CLI を起動します。
   `python kiritan_chat_cli.py`
4. 画面の指示に従ってプロフィールを入力し、`mode dual`（既定）/`mode mic`/`mode loop` などを選びながら会話します。

### 主なコマンド
| コマンド | 説明 |
| --- | --- |
| `mode text/mic/loop/dual` | 入力モード切り替え（`voice`/`system` などの別名あり） |
| `time N` | 録音秒数（mic / loop 共通） |
| `speed X` | 読み上げ速度。`0.5～4.0` の範囲で制限 |
| `style` | 会話スタイル（light/normal/deep）を再選択 |
| `profile` / `name` | 名前・ジェンダー・年代の再入力 |
| `exit` | 終了 |

### ログと履歴
- `logs/profile_history.jsonl` にプロフィール入力の履歴を JSONL 形式で追記します。
- `KIRITAN_TAB_DEBUG=1` を指定すると、VOICEROID が音声効果タブへ遷移したタイミングが `[tab-debug …]` ログに記録されます。

---

## voiceinputting との連携

[voiceinputting](https://github.com/takumayellow/voiceinputting) は OpenAI 音声認識 API を使って音声をテキスト化するブリッジツールです。
きりたん Chat Autoplay の `mode mic` と組み合わせることで、マイク入力から自動的にテキスト変換→きりたん読み上げという完全ハンズフリーフローを構築できます。

### 連携方法

```
[マイク入力]
    |
    v
voiceinputting (gpt-4o-mini-transcribe でテキスト化)
    |
    v  テキスト文字列
kiritan_chat_cli.py の標準入力 / mode text
    |
    v
OpenAI API (返答生成)
    |
    v
SeikaSay2.exe → VOICEROID 読み上げ
```

#### 手順

1. **voiceinputting をセットアップ**（別ウィンドウで起動）:
   ```powershell
   cd ..\voiceinputting
   pip install -r requirements.txt
   python -m src.voice_to_codex --auto-send
   ```

2. **このリポジトリを `mode text` で起動**:
   ```powershell
   python kiritan_chat_cli.py
   # 起動後: mode text
   ```

3. voiceinputting が文字起こしした内容を、きりたん Chat の入力にコピー＆ペーストして送信します。

> **将来的な統合**: 両ツールを 1 プロセスに統合するパイプライン実装も検討中です。`OPENAI_API_KEY` は共通の環境変数を使用します。

---

## トラブルシューティング
- **VOICEROID が見つからない**: VOICEROID＋ 東北きりたん EX を起動し、タイトルに「VOICEROID」「きりたん」が含まれていることを確認してください。
- **SeikaSay2.exe が見つからない**: `SEIKA_EXE` にフルパスを設定するか、デフォルト配置（`AssistantSeika/.../SeikaSay2.exe`）に置きます。
- **UIA の NULL COM pointer エラー**: `_activate_control` で握り潰すように対処済みですが、再発する場合は VOICEROID ウィンドウを一度最小化→復帰してください。
- **音声効果タブに固定したい**: 逆に音声効果タブを維持したい場合は `ensure_phrase_tab()` 呼び出し箇所をコメントアウトするか、`start_phrase_tab_sentry()` を起動しないカスタムモードを作成してください。なお公式仕様上、録音後に音声効果側へ戻る挙動は避けられません。

## ライセンス / 貢献
- 本リポジトリのコードは作者の個人検証用途として公開しており、ライセンスは同梱ファイルに準じます。
- 変更や改善の提案は Pull Request / Issue で受け付けています。README の内容を更新する場合は、最近の仕様変更（ペルソナ、タブ制御、フォーカス動作など）との整合性にご注意ください。
