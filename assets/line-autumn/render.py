"""LINE 公式アカウント用の秋の画像を HTML から書き出す.

    python assets/line-autumn/render.py

background.png: プロフィール背景 (1080x878)
rich_menu.jpg:  リッチメニュー (2500x843. LINE の上限 1MB に収めるため JPEG)

必要なもの: playwright (Python)・Pillow と Microsoft Edge か Chromium. フォントは Google Fonts から読むのでネットにつなぐ.
"""
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
EDGE = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
TARGETS = [
    ("background.html", {"width": 1080, "height": 878}, "background.png"),
    ("rich_menu.html", {"width": 2500, "height": 843}, "rich_menu.png"),
]


def main() -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=str(EDGE) if EDGE.exists() else None)
        for html, size, out in TARGETS:
            page = browser.new_page(viewport=size)
            page.goto((HERE / html).as_uri(), wait_until="networkidle")
            page.evaluate("document.fonts.ready")
            page.screenshot(path=str(HERE / out))
            page.close()
            print("wrote", HERE / out)
        browser.close()
    rich = HERE / "rich_menu.png"
    Image.open(rich).convert("RGB").save(HERE / "rich_menu.jpg", quality=90)
    rich.unlink()
    print("wrote", HERE / "rich_menu.jpg")


if __name__ == "__main__":
    main()
