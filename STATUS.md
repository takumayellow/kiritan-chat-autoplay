# きりたんチャット プロジェクト現況

**ブランチ**: `feature/line-bot`
**最終更新**: 2026-03-21

---

## 完了済み

### プロジェクト構造整理
- [x] モードごとにフォルダ分け（`core/` `cli/` `gui/` `line_bot/` `legacy/`）
- [x] 旧バージョンを `legacy/` に移動
- [x] 設定ファイルを `config/` に整理
- [x] エントリポイント（`run_cli.py` `run_gui.py` `run_linebot.py`）作成
- [x] 起動スクリプト（`start_linebot.bat` `start_gui.bat` `stop_all.bat`）作成

### 共通モジュール (`core/`)
- [x] `kiritan_core.py` - ペルソナ構築、OpenAIチャット、ユーザープロフィール、SQLite会話履歴
- [x] `kiritan_voicevox.py` - VOICEVOX HTTP API連携（音声合成・再生）

### GUI版 (`gui/`)
- [x] `kiritan_chat_gui_new.py` - VOICEVOX対応tkinter GUI
- [x] 白ベース・モダンデザイン（Segoe UIフォント）
- [x] きりたん立ち絵（右パネル、自動リサイズ）
- [x] きりたんアイコン（ヘッダー）
- [x] 入力欄のレイアウト修正（pack順序問題を解消）
- [x] VOICEVOX自動接続・Speaker ID自動検出
- [x] 会話スタイル切替（ライト/ゆったり/じっくり）

### LINE Bot (`line_bot/`)
- [x] `server.py` - Flask Webhookサーバー
- [x] テキスト会話（OpenAI API経由）
- [x] コマンド（/help /reset /profile /style）
- [x] ウェルカムメッセージ（友だち追加時）
- [x] SDKパースエラーのフォールバック処理
- [x] LINE公式アカウント プロフィール画像・背景画像（`assets/line_icon.png` `line_background.png`）
- [x] セットアップガイド（`docs/linebot_setup.md`）
- [x] ngrok接続確認（検証ボタンで成功確認済み）

### GitHub Issues
- [x] #4 LINE Bot対応
- [x] #5 YouTube Live常時配信で視聴者と会話
- [x] #6 LINE を軸にした集客導線の構築

---

## 未解決・残タスク

### LINE Bot - メッセージ返信が動かない（最優先）
- ngrok → サーバーの接続は確認済み（検証ボタン成功）
- LINEからメッセージ送信時にサーバーログに `Webhook received` が出ない
- **原因調査が必要**:
  - [ ] LINE Official Account Managerの応答設定の再確認
  - [ ] Webhookイベントの送信設定（LINE Developers側）
  - [ ] サーバーのログレベルをDEBUGに上げて詳細確認
  - [ ] ngrokのinspect画面（http://localhost:4040）でリクエスト内容確認
  - [ ] Channel Secretの再発行（公開してしまったため）

### LINE Bot - 追加機能
- [ ] 音声メッセージ返信（VOICEVOX → WAV → m4a → LINE AudioMessage）
- [ ] リッチメニュー画像作成・設定
- [ ] 本番デプロイ（Railway / Render）

### GUI版
- [ ] テキスト入力が一部環境で効かない問題の根本原因調査
- [ ] `kiritan_chat_gui_app.py`（旧GUI）の `core/` 移行対応

### 全体
- [ ] `kiritan_chat_autoplay_overview.pdf` を `docs/` に移動（ルートに残っている）
- [ ] `README.md` をフォルダ構成変更に合わせて更新
- [ ] Channel Secretの再発行と `.env` 更新
- [ ] mainブランチへのマージ

---

## ファイル構成

```
kiritan-chat-autoplay/
├── run_cli.py               # CLI起動
├── run_gui.py               # GUI起動
├── run_linebot.py           # LINE Bot起動
├── start_linebot.bat        # LINE Bot一括起動（サーバー+ngrok）
├── start_gui.bat            # GUI一括起動（VOICEVOX+アプリ）
├── stop_all.bat             # 全停止
├── core/                    # 共通ロジック
│   ├── kiritan_core.py      #   ペルソナ/OpenAI/DB
│   └── kiritan_voicevox.py  #   VOICEVOX連携
├── cli/                     # CLI版（SeikaSay2+VOICEROID）
│   └── kiritan_chat_cli.py
├── gui/                     # GUI版（VOICEVOX+tkinter）
│   ├── kiritan_chat_gui_new.py
│   └── kiritan_chat_gui_app.py
├── line_bot/                # LINE Bot
│   ├── server.py
│   └── requirements.txt
├── legacy/                  # 旧バージョン
├── config/                  # 設定ファイル
├── assets/                  # 画像素材
├── docs/                    # ドキュメント・PDF
└── debug/                   # デバッグスクリプト
```

## 関連リンク
- GitHub: https://github.com/takumayellow/kiritan-chat-autoplay
- Issues: #4(LINE Bot) #5(YouTube Live) #6(集客導線)
