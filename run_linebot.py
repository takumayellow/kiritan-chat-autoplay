# -*- coding: utf-8 -*-
"""LINE Botサーバーを起動"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from line_bot.server import app
import os

port = int(os.environ.get("PORT", 5000))
app.run(host="0.0.0.0", port=port, debug=False)
