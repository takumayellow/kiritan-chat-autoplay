# きりたんトーク LINE Bot セットアップガイド

## 1. LINE公式アカウント設定

### プロフィール
| 項目 | 内容 |
|---|---|
| アカウント名 | きりたんトーク |
| アイコン | `assets/line_icon.png` (640x640) |
| 背景画像 | `assets/line_background.png` (1080x878) |
| ステータスメッセージ | 東北きりたんとおしゃべりできるAIチャット |

### あいさつメッセージ
友だち追加時に自動送信される。LINE Official Account Manager → ホーム → あいさつメッセージ で設定:

```
友だち追加ありがとう！東北きりたんです。

わたしとおしゃべりしてみませんか？
メッセージを送ってくれたら、お返事するね。

便利なコマンド:
📝 /profile 名前=○○  → 名前を教えてくれたら覚えるよ
🎨 /style 1~3  → 会話スタイルを変えられるよ
🔄 /reset  → 会話をリセット
❓ /help  → コマンド一覧

まずは気軽に話しかけてね！
```

### 応答設定
LINE Official Account Manager → 設定 → 応答設定:

| 項目 | 設定値 |
|---|---|
| 応答モード | Bot |
| あいさつメッセージ | ON |
| 応答メッセージ | OFF |
| Webhook | ON |

### LINEコール
OFF（不要）

---

## 2. Messaging API 接続

### LINE Developers コンソール側
https://developers.line.biz/ でチャネルを開く:

| 項目 | 設定値 |
|---|---|
| Webhook URL | `https://あなたのサーバー/callback` |
| Webhookの利用 | ON |
| 応答メッセージ | OFF（二重返信防止） |

### 必要なキー（チャネル設定画面から取得）
- **Channel Secret**: 「チャネル基本設定」タブ
- **Channel Access Token**: 「Messaging API設定」タブ → 発行ボタン

---

## 3. サーバー側の設定

### .env ファイル
プロジェクトルートの `.env` に以下を追加:

```
OPENAI_API_KEY=sk-xxxxxxxx
LINE_CHANNEL_ACCESS_TOKEN=xxxxxxxx
LINE_CHANNEL_SECRET=xxxxxxxx
```

### 依存パッケージ
```bash
pip install -r linebot/requirements.txt
```

### 起動
```bash
python run_linebot.py
```
デフォルトで `http://0.0.0.0:5000` で起動。

---

## 4. 開発環境での公開（ngrok）

LINEのWebhookはHTTPS公開URLが必要。開発時は ngrok を使用:

```bash
# インストール（未インストールの場合）
# https://ngrok.com/ からダウンロード

# ローカルサーバーを公開
ngrok http 5000
```

表示された `https://xxxx.ngrok-free.app` をコピーし、
LINE Developers → Webhook URL に `https://xxxx.ngrok-free.app/callback` を設定。

### 検証ボタン
Webhook URL を設定後、「検証」ボタンを押して 200 OK が返れば接続成功。

---

## 5. 本番デプロイ

### Railway（推奨・無料枠あり）
```bash
# Railway CLI インストール
npm install -g @railway/cli

# デプロイ
railway login
railway init
railway up
```

`Procfile` をプロジェクトルートに作成:
```
web: python run_linebot.py
```

環境変数は Railway のダッシュボードで設定。

### Render（代替）
- Web Service として新規作成
- Build Command: `pip install -r linebot/requirements.txt`
- Start Command: `python run_linebot.py`
- 環境変数をダッシュボードで設定

---

## 6. リッチメニュー

LINE Official Account Manager → ホーム → リッチメニュー で作成。

### 推奨レイアウト（3分割）
```
┌──────────┬──────────┬──────────┐
│  💬       │  🎨       │  📦       │
│ きりたんと  │ スタイル    │ プロダクト  │
│ 話す      │ 変更       │ 一覧      │
└──────────┴──────────┴──────────┘
```

| ボタン | タイプ | アクション |
|---|---|---|
| きりたんと話す | テキスト | `こんにちは！` |
| スタイル変更 | テキスト | `/style` |
| プロダクト一覧 | URL | ポートフォリオURL |

### 画像サイズ
- 大: 2500x1686px
- 小: 2500x843px

---

## 7. QRコード（学祭向け）

LINE Official Account Manager → 友だちを増やす → 友だち追加ガイド → QRコード

ダウンロードして印刷し、ブースに掲示。

---

## 8. Bot コマンド一覧

| コマンド | 説明 |
|---|---|
| (通常メッセージ) | きりたんと会話 |
| `/help` | コマンド一覧 |
| `/reset` | 会話履歴リセット |
| `/profile` | プロフィール確認 |
| `/profile 名前=○○ 性別=○○ 年代=○○` | プロフィール設定 |
| `/style` | スタイル一覧 |
| `/style 1` | ライト雑談 (gpt-4o-mini) |
| `/style 2` | ゆったり会話 (o4-mini) |
| `/style 3` | じっくり深掘り (o4-mini-high) |

---

## 画像アセット

| ファイル | 用途 | サイズ |
|---|---|---|
| `assets/line_icon.png` | プロフィールアイコン | 640x640 |
| `assets/line_background.png` | プロフィール背景 | 1080x878 |
| `assets/kiritan_small.png` | アイコン原画 | - |
| `assets/kiritan_stand.png` | 立ち絵原画 | - |
