# きりたんチャット LINE Bot - Docker イメージ
# Railway / Render / Fly.io 等のコンテナ環境向け

FROM python:3.11-slim

WORKDIR /app

# 依存パッケージのインストール
COPY line_bot/requirements.txt ./line_bot/requirements.txt
RUN pip install --no-cache-dir -r line_bot/requirements.txt

# アプリケーションコードのコピー
COPY core/ ./core/
COPY line_bot/ ./line_bot/
COPY run_linebot.py ./

# ポート公開（Railway / Render は PORT 環境変数で上書きする）
EXPOSE 5000

# 本番起動（gunicorn）
# PORT 環境変数は Railway/Render が自動設定する。デフォルトは 5000
CMD gunicorn --bind 0.0.0.0:${PORT:-5000} --workers 1 --timeout 120 run_linebot:app
