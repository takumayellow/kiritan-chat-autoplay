# -*- coding: utf-8 -*-
"""GUI版きりたんチャット（VOICEVOX）を起動"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from gui.kiritan_chat_gui_new import main
main()
