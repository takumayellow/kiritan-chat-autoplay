# きりたん Chat Autoplay（GUI優先・CLI併用）

## 1) “完成版”の全体像
- **会話生成**: OpenAI API（`OPENAI_MODEL` 未指定時は `gpt-4o-mini → o4-mini-high → o3-mini → gpt-4o` 順で自動フォールバック）
- **読み上げ**: `SeikaSay2.exe` の **CLI** で再生
- **音声入力**: `mode mic` で自動録音（Enter で途中停止 / `time N` で録音秒数を変更）
- **UI 安定化**: VOICEROID のタブが勝手に動く問題に対し、**UIA** で「**フレーズ編集**」に自動復帰  
  watchdog が `TabItem/ボタン/ラジオ` を監視し、必要なら `Ctrl+1` 擬似入力まで使って復帰
- **フォーカス**: 毎回の再生後に PowerShell を前面復帰
- **速度**: 入力値をそのまま Seika に渡し、`0.5–4.0x` にクランプ（2倍速化バグを解消）
- **ペルソナ**: 東北きりたんの年齢・口調・好物（きりたんぽ）・秋田弁の言い回しをプロンプトに明記し、AIらしい口調に戻らないよう調整
- **呼称ヒアリング**: 起動直後に名前・呼称・ジェンダー・年代を番号メニューで確認し、回答に沿ってペルソナへ反映（`profile` / `name` コマンドで再設定可）
- **プロフィール履歴**: 入力されたプロフィールは `logs/profile_history.jsonl` に JSONL 形式で追記され、後から分析できる
- **プロフィール履歴**: 入力されたプロフィールは `logs/profile_history.jsonl` に JSONL 形式で追記され、後から分析できる
- 実装方針は **「CLIで読み上げ」＋「GUIでタブ復帰」** の二本立て（GUI操作は必要最小限）

## 2) GUI（pywinauto）を使う背景と現実解
- 直接 VOICEROID を操作して再生ボタンを押す設計を検討
- しかし実運用では以下で不安定:
  - 32/64bit 非一致の警告（VOICEROID は 32bit）
  - Win32 backend は要素探索が不安定・遅い
  - **ウィンドウタイトルの「全角＋（VOICEROID＋）」** を見落とすと検出失敗
  - 読み上げ後にタブが「音声効果」等へ飛ぶ
- 結論: **再生は CLI、GUI はタブ復帰のみ**に限定するのが堅牢

## 3) なぜ CLI（SeikaSay2.exe）でやり切るのが正解か
- クリックやウィンドウ状態に依存せず **確実に再生**
- `-cid` や `-speed` を引数で明示制御
- 既知問題と対処
  - `invalid option: -play` → VOICEROID 本体に投げていた → **SeikaSay2.exe** に向ける
  - `Process "SeikaSay2.exe" not found` → パス誤り → 既定値/環境変数でカバー
  - **2倍速化** → 内部の二重適用をやめ、入力値をそのまま渡す（`0.5–4.0` クランプ）

## 4) “タブが動く”問題の最終対処
- 再生後にタブが他へ移動する事象あり
- 監視スレッド `_guard_phrase_tab` を調整し、再生停止直後も約 1.4 秒 linger しながら監視を継続  
  → 遅れて「音声効果」へ飛ぶケースも再度「フレーズ編集」へ戻せるようになった
- 再生開始前に `ensure_phrase_tab_with_retry()` を複数回実行し、実際に話し出す時点でタブが「フレーズ編集」にあることを確認してから `SeikaSay2.exe -play` を発火
- `force_phrase_tab()` で VOICEROID ウィンドウへフォーカスを戻しつつ高速にリトライし、再生前後で確実にタブを引き戻す
- `KIRITAN_TAB_DEBUG=1` を設定するとタブ監視の詳細ログが PowerShell に出力され、原因調査時に役立つ
- VOICEROID 本体が起動していない場合は数秒間隔で注意喚起を出し、ログが連打されないよう抑制
- **UIA** で `TabItem` だけでなく **ボタン/ラジオも列挙**し「フレーズ編集」へ復帰
  - `select()` → `invoke()` → `click_input()` → `Ctrl+1` 擬似入力まで段階的に試行
- これは `tab_switch_test.py` で検証済み → 本体 `ensure_phrase_tab()` に統合
- 手動検証メモ（Windows 実機で実施推奨）
  - 連続再生を数回繰り返し、停止後もタブが「フレーズ編集」を維持しているか確認
  - 再生停止直後に VOICEROID 側で速度・プリセットを変更し、タブがずれても 1 秒以内に戻るか確認
  - PowerShell ログに `⚠️/✖️` が頻出しないかを観察（連続失敗時は watchdog 間隔を調整）

## 5) 依存・環境・実行
- 必須: Windows（VOICEROID＋ 東北きりたん EX が稼働）、Python 3.11+
- ライブラリ:  
  `pip install openai pywinauto pywin32 speechrecognition sounddevice`
  （mic/loop を使わないなら `speechrecognition` と `sounddevice` は不要）
- 環境変数  
  - `OPENAI_API_KEY`（必須）  
  - `OPENAI_MODEL`（任意）  
  - `SEIKA_EXE`（任意: `SeikaSay2.exe` の絶対パスで既定値上書き）
- 実行:  
  `python kiritan-chat-autoplay.py`
- プロンプト:  
  `mode dual|text|mic|loop | time N | speed X | exit`  
  例）`speed 1.2`, `mode mic`, `time 6`
- いつでも `Ctrl+C` で安全に終了できます。`mode loop`/`mode system` は Stereo Mix などのシステム音入力デバイスが必要で、取得できなければ自動で text モードへ戻ります。
- 動作フロー:  
  起動直後に **フレーズ編集へ復帰** → 返答生成 → `SeikaSay2 -play` で再生  
  再生中は watchdog がタブを監視し続け、終了後も再度タブ復帰＋前面復帰を保证

## 6) /debug にある検証資材
- `tab_switch_test.py`: タブ列挙と復帰の確定版
- `debug_step*_*.py`: ウィンドウ列挙/接続/PID 解決など  
  ※タイトルは **VOICEROID＋ 東北きりたん EX**（全角＋）

## 7) よくあるエラーと即時対処
- `NameError: find_voiceroid_handle ...` → 検証関数の移植漏れ。現行コードは修正済み
- `invalid option: -play` → 送信先が VOICEROID 本体。**SeikaSay2.exe** に向ける
- `UIA: NULL COM pointer access` → `select → invoke → click_input` フォールバックを実装
- `Window not found` → タイトルの **全角＋** を確認

## 8) 本体のキーノート（後続開発の入口）
- `ensure_phrase_tab()`：タブ復帰の中核（起動時＆再生後に必ず呼ぶ）
- `speak()`：SeikaSay2 CLI 実行。終了後にタブ復帰＋前面復帰
- `chat_once()`：モデルのフォールバック実装
- 速度：`0.5–4.0` にクランプ、二重掛け禁止

## 9) 拡張の出発点
- 簡易トースト/ログ化、速度・抑揚プリセット、録音系の安定化（必要時のみ）


## モード構成

- **GUI 基本版（推奨）**: kiritan_chat_gui.py  
  AssistantSeika への依存なし。VOICEROID＋東北きりたん EX を起動してから実行。  
  返答は VOICEROID で再生され、ターミナルにも表示されます。

- **CLI 版**: kiritan_chat_cli.py  
  SeikaSay2.exe 経由で再生（HTTP/WCF 不要の環境ならこちらでも可）。
