# きりたん Chat Autoplay（GUI優先・CLI併用）

## 1) “完成版”の全体像
- **会話生成**: OpenAI API（`OPENAI_MODEL` 未指定時は `gpt-4o-mini → o4-mini-high → o3-mini → gpt-4o` 順で自動フォールバック）
- **読み上げ**: `SeikaSay2.exe` の **CLI** で再生
- **UI 安定化**: VOICEROID のタブが勝手に動く問題に対し、**UIA** で「**フレーズ編集**」に自動復帰  
  起動時と毎回の再生後に `select() → invoke() → click_input()` の順でフォールバック
- **フォーカス**: 毎回の再生後に PowerShell を前面復帰
- **速度**: 入力値をそのまま Seika に渡し、`0.5–4.0x` にクランプ（2倍速化バグを解消）
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
- **UIA** で `TabItem` を列挙し「**フレーズ編集**」を
  - `select()` → 失敗なら `invoke()` → さらに失敗なら `click_input()` で確実に復帰
- これは `tab_switch_test.py` で検証済み → 本体 `ensure_phrase_tab()` に統合

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
- 動作フロー:  
  起動直後に **フレーズ編集へ復帰** → 返答生成 → `SeikaSay2 -play` で再生 → 再生後にタブ復帰＋前面復帰

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

