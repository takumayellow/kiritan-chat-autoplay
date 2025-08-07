# debug_step1_enum_windows.py
from pywinauto import findwindows

# 正規表現ではなく、とりあえず全ウィンドウを取得
windows = findwindows.find_elements()

for w in windows:
    print(f"handle={w.handle!r}  title={w.name!r}")
