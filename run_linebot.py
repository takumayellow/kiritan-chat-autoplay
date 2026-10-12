# -*- coding: utf-8 -*-
"""LINE Botサーバーを起動

開発時:
  python run_linebot.py

本番 (gunicorn):
  gunicorn run_linebot:app
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from line_bot.server import app
import os

if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")  # ngrok で公開するので LAN には開けない。本番は gunicorn
    port = int(os.environ.get("PORT", 5000))
    app.run(host=host, port=port, debug=False)
