# -*- coding: utf-8 -*-
"""CLI版きりたんチャットを起動"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from cli.kiritan_chat_cli import main
main()
