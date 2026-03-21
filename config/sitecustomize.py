# Auto JP font + safe defaults for matplotlib (project-wide)
import matplotlib as mpl
from matplotlib import font_manager
candidates = ["IPAexGothic","IPAPGothic","Noto Sans CJK JP","TakaoPGothic","Yu Gothic","MS Gothic","DejaVu Sans"]
for name in candidates:
    try:
        font_manager.findfont(name, fallback_to_default=False)
        mpl.rcParams["font.family"] = name
        break
    except Exception:
        continue
mpl.rcParams["axes.unicode_minus"] = False  # マイナス記号の豆腐回避
mpl.rcParams["text.usetex"] = False        # 外部TeX不要（matplotlib内の数式だけにする）
