"""flyer.html から印刷用の PNG (A4 縦, 約 288dpi) と PDF (A4 1 ページ) を作る.

    python assets/flyer-ridaisai/render.py

必要なもの: playwright (Python)・Pillow と Microsoft Edge か Chromium. フォントは Google Fonts から読むのでネットにつなぐ.
"""
from pathlib import Path

from PIL import Image
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
EDGE = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
A4 = {"width": 794, "height": 1122}


def main() -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=str(EDGE) if EDGE.exists() else None)
        page = browser.new_page(viewport=A4, device_scale_factor=3)
        page.goto((HERE / "flyer.html").as_uri(), wait_until="networkidle")
        page.evaluate("document.fonts.ready")
        page.screenshot(path=str(HERE / "flyer.png"))
        browser.close()
    # ブラウザの印刷で PDF にするとレイアウトが画面とずれて 2 ページにあふれるので,
    # 確かめた PNG をそのまま A4 1 ページに貼る (幅 210mm で約 288dpi)
    png = Image.open(HERE / "flyer.png").convert("RGB")
    png.save(HERE / "flyer.pdf", resolution=png.width / (210 / 25.4))
    print("wrote", HERE / "flyer.png", "and", HERE / "flyer.pdf")


if __name__ == "__main__":
    main()
