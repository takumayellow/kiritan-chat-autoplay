# LINE Bot 本番デプロイガイド

このドキュメントでは、きりたんチャット LINE Bot を Railway または Render にデプロイする手順を説明します。

---

## 前提条件

- LINE Developers でチャンネル作成済み
- OpenAI API キー取得済み
- GitHub リポジトリへのプッシュ権限

---

## 1. Render へのデプロイ

### 1-1. Render アカウント作成
1. [render.com](https://render.com) でアカウント登録（GitHub 連携）
2. 無料プランで Web Service が 1 つ利用可能（月 750 時間）

### 1-2. サービス作成
1. ダッシュボード → **New +** → **Web Service**
2. リポジトリを選択し、以下を設定:

| 項目 | 値 |
|------|-----|
| Runtime | Python 3 |
| Build Command | `pip install -r line_bot/requirements.txt` |
| Start Command | `gunicorn --bind 0.0.0.0:$PORT --workers 1 --timeout 120 run_linebot:app` |

または `render.yaml` を使えばこれらの設定が自動適用されます。

### 1-3. 環境変数の設定
Render ダッシュボードの **Environment** タブで以下を設定:

```
OPENAI_API_KEY=sk-...
LINE_CHANNEL_ACCESS_TOKEN=...
LINE_CHANNEL_SECRET=...
```

### 1-4. SQLite の永続化
`render.yaml` に Disk（1GB）の設定が含まれています。  
サーバー側では `db_path` を `/app/data/linebot_chat.sqlite` に変更することで永続化できます。

> **注意**: Render の無料プランではディスクが使えないため、永続化が必要な場合は有料プランへのアップグレードか、PostgreSQL への移行を検討してください。

### 1-5. LINE Webhook URL の設定
デプロイ後に表示される URL（例: `https://kiritan-chat-linebot.onrender.com`）を  
LINE Developers の **Webhook URL** に設定:

```
https://<your-service>.onrender.com/callback
```

---

## 2. Railway へのデプロイ

### 2-1. Railway アカウント作成
1. [railway.app](https://railway.app) でアカウント登録
2. **Hobby プラン**（$5/月）で月 500 時間利用可能

### 2-2. デプロイ手順
1. ダッシュボード → **New Project** → **Deploy from GitHub repo**
2. リポジトリを選択（`Dockerfile` が自動検出されます）
3. **Variables** タブで環境変数を設定（Render と同じ 3 つ）
4. デプロイ完了後、**Settings** → **Networking** でドメインを生成
5. LINE Developers の Webhook URL を更新

---

## 3. Docker でのローカル動作確認

```bash
# イメージビルド
docker build -t kiritan-chat-linebot .

# 起動（環境変数ファイルを使用）
docker run -p 5000:5000 \
  --env-file .env \
  kiritan-chat-linebot
```

---

## 4. GitHub Actions CI/CD

`.github/workflows/linebot_ci.yml` により以下が自動実行されます:

- **PR 時**: インポートチェック・ヘルスチェックテスト・Docker ビルド
- **main/master push 時**: 同上

Render への自動デプロイは Render の **Auto-Deploy** 機能を有効化することで実現できます（GitHub 連携済みの場合、main ブランチへの push で自動デプロイ）。

### GitHub Secrets の設定
将来的に CI からデプロイする場合は、リポジトリの **Settings → Secrets** に追加:

| Secret 名 | 用途 |
|-----------|------|
| `RENDER_API_KEY` | Render CLI でのデプロイ（オプション） |
| `RAILWAY_TOKEN` | Railway CLI でのデプロイ（オプション） |

---

## 5. 死活監視

### 無料の監視サービス
| サービス | 無料枠 | 備考 |
|---------|--------|------|
| [UptimeRobot](https://uptimerobot.com) | 50 モニター、5 分間隔 | 最も広く使われている |
| [Pulsetic](https://pulsetic.com) | 5 モニター、3 分間隔 | シンプルな UI |
| [Freshping](https://freshping.io) | 50 チェック | Freshworks 系 |

> **注意（pulsetic.com について）**: Pulsetic はサービスとして実在し、シンプルな死活監視として利用できます。  
> ただし、GitHub Issues に「このサイトが使いやすい」といった内容で投稿されるコメントは、  
> AI 生成・宣伝目的のスパムコメントである可能性があります。サービス自体の信頼性と  
> コメントの信頼性は別物として判断することを推奨します。  
> 死活監視には実績ある **UptimeRobot** を第一候補とするのが安心です。

### UptimeRobot 設定例
1. [uptimerobot.com](https://uptimerobot.com) でアカウント登録
2. **Add New Monitor**:
   - Monitor Type: HTTP(s)
   - URL: `https://<your-service>.onrender.com/`
   - Monitoring Interval: 5 minutes
3. Alert Contacts に LINE Notify や Discord webhook を設定

---

## 6. 無料枠まとめと選定

| サービス | 無料枠 | スリープ | SQLite永続化 | 推奨度 |
|---------|--------|---------|-------------|--------|
| Render | 750h/月 | 15分で停止 | 有料プランのみ | △ |
| Railway | 500h/月 ($5~) | なし | ボリューム可 | ○ |
| Fly.io | 3 shared VM | なし | ボリューム可 | ○ |

### 推奨: Railway
- スリープなし（LINE Bot に最適）
- ボリュームマウントで SQLite 永続化可能
- 月 $5 で安定稼働

### Render の無料プランの注意点
- 15 分間リクエストがないとスリープ（コールドスタートで ~30 秒かかる）
- LINE Bot は応答が遅れることがある → UptimeRobot で 5 分ごとにヘルスチェックすることで回避可能
