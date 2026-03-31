# -*- coding: utf-8 -*-
"""
setup_rich_menu.py
LINE Bot リッチメニューのセットアップスクリプト

使い方:
  1. .env に LINE_CHANNEL_ACCESS_TOKEN を設定
  2. assets/rich_menu.png を用意（2500x843 推奨）
  3. python line_bot/setup_rich_menu.py

リッチメニュー画像がない場合はプレースホルダーを自動生成します。
"""

import json
import os
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import requests

# ── 設定 ─────────────────────────────────────────────────
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
API_BASE = "https://api.line.me/v2/bot"

# プロダクトリンク集サイト（GitHub Pages）
PRODUCTS_URL = "https://takumayellow.github.io/kiritan-chat-autoplay/products/"

# YouTube チャンネル（仮URL - 実際のチャンネルURLに差し替えてください）
YOUTUBE_URL = os.environ.get(
    "KIRITAN_YOUTUBE_URL",
    "https://www.youtube.com/@takumayellow",
)

# リッチメニュー画像パス
ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
RICH_MENU_IMAGE = ASSETS_DIR / "rich_menu.png"

# ── リッチメニュー定義 ──────────────────────────────────────
# 画像サイズ: 2500 x 843（half サイズ、3列 x 2行 = 6エリア）
RICH_MENU_BODY = {
    "size": {"width": 2500, "height": 843},
    "selected": True,
    "name": "きりたんチャット メニュー",
    "chatBarText": "メニューを開く",
    "areas": [
        {
            "bounds": {"x": 0, "y": 0, "width": 833, "height": 421},
            "action": {
                "type": "uri",
                "label": "YouTube",
                "uri": YOUTUBE_URL,
            },
        },
        {
            "bounds": {"x": 833, "y": 0, "width": 834, "height": 421},
            "action": {
                "type": "uri",
                "label": "ゲーム",
                "uri": PRODUCTS_URL,
            },
        },
        {
            "bounds": {"x": 1667, "y": 0, "width": 833, "height": 421},
            "action": {
                "type": "message",
                "label": "ヘルプ",
                "text": "/help",
            },
        },
        {
            "bounds": {"x": 0, "y": 421, "width": 833, "height": 422},
            "action": {
                "type": "message",
                "label": "プロフィール設定",
                "text": "/profile",
            },
        },
        {
            "bounds": {"x": 833, "y": 421, "width": 834, "height": 422},
            "action": {
                "type": "message",
                "label": "会話スタイル",
                "text": "/style",
            },
        },
        {
            "bounds": {"x": 1667, "y": 421, "width": 833, "height": 422},
            "action": {
                "type": "message",
                "label": "リセット",
                "text": "/reset",
            },
        },
    ],
}


def _headers(content_type: str = "application/json"):
    return {
        "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
        "Content-Type": content_type,
    }


def create_rich_menu() -> str:
    """リッチメニューを作成し、リッチメニューIDを返す"""
    resp = requests.post(
        f"{API_BASE}/richmenu",
        headers=_headers(),
        json=RICH_MENU_BODY,
    )
    resp.raise_for_status()
    rich_menu_id = resp.json()["richMenuId"]
    print(f"[OK] リッチメニュー作成: {rich_menu_id}")
    return rich_menu_id


def upload_rich_menu_image(rich_menu_id: str, image_path: Path):
    """リッチメニュー画像をアップロード"""
    if not image_path.exists():
        print(f"[WARN] 画像が見つかりません: {image_path}")
        print("       プレースホルダー画像を生成します...")
        _generate_placeholder_image(image_path)

    with open(image_path, "rb") as f:
        resp = requests.post(
            f"https://api-data.line.me/v2/bot/richmenu/{rich_menu_id}/content",
            headers={
                "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
                "Content-Type": "image/png",
            },
            data=f.read(),
        )
    resp.raise_for_status()
    print(f"[OK] 画像アップロード完了")


def set_default_rich_menu(rich_menu_id: str):
    """デフォルトリッチメニューに設定"""
    resp = requests.post(
        f"{API_BASE}/user/all/richmenu/{rich_menu_id}",
        headers=_headers(),
    )
    resp.raise_for_status()
    print(f"[OK] デフォルトリッチメニューに設定完了")


def delete_all_rich_menus():
    """既存のリッチメニューを全削除"""
    resp = requests.get(
        f"{API_BASE}/richmenu/list",
        headers=_headers(),
    )
    resp.raise_for_status()
    menus = resp.json().get("richmenus", [])
    for menu in menus:
        mid = menu["richMenuId"]
        del_resp = requests.delete(
            f"{API_BASE}/richmenu/{mid}",
            headers=_headers(),
        )
        del_resp.raise_for_status()
        print(f"[OK] 削除: {mid}")
    if not menus:
        print("[INFO] 既存のリッチメニューはありません")


def list_rich_menus():
    """リッチメニュー一覧を表示"""
    resp = requests.get(
        f"{API_BASE}/richmenu/list",
        headers=_headers(),
    )
    resp.raise_for_status()
    menus = resp.json().get("richmenus", [])
    if not menus:
        print("[INFO] リッチメニューはありません")
        return
    for menu in menus:
        print(f"  ID: {menu['richMenuId']}")
        print(f"  名前: {menu['name']}")
        print(f"  エリア数: {len(menu.get('areas', []))}")
        print()


def _generate_placeholder_image(output_path: Path):
    """プレースホルダーのリッチメニュー画像を生成（PIL使用）"""
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("[ERROR] Pillow が必要です: pip install Pillow")
        sys.exit(1)

    width, height = 2500, 843
    img = Image.new("RGB", (width, height), "#2C2F33")
    draw = ImageDraw.Draw(img)

    labels = [
        ("YouTube", "#FF0000"),
        ("ゲーム", "#7289DA"),
        ("ヘルプ", "#43B581"),
        ("プロフィール", "#FAA61A"),
        ("スタイル", "#F47FFF"),
        ("リセット", "#99AAB5"),
    ]

    col_width = width // 3
    row_height = height // 2

    for i, (label, color) in enumerate(labels):
        col = i % 3
        row = i // 3
        x0 = col * col_width + 4
        y0 = row * row_height + 4
        x1 = (col + 1) * col_width - 4
        y1 = (row + 1) * row_height - 4

        draw.rounded_rectangle([x0, y0, x1, y1], radius=20, fill=color)

        try:
            font = ImageFont.truetype("arial.ttf", 48)
        except (OSError, IOError):
            font = ImageFont.load_default()

        bbox = draw.textbbox((0, 0), label, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = x0 + (x1 - x0 - tw) // 2
        ty = y0 + (y1 - y0 - th) // 2
        draw.text((tx, ty), label, fill="white", font=font)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(output_path)
    print(f"[OK] プレースホルダー画像を生成: {output_path}")


def main():
    if not CHANNEL_ACCESS_TOKEN:
        print("[ERROR] LINE_CHANNEL_ACCESS_TOKEN が設定されていません")
        print("       .env ファイルに設定してください")
        sys.exit(1)

    import argparse
    parser = argparse.ArgumentParser(description="きりたんチャット リッチメニューセットアップ")
    parser.add_argument("--delete-all", action="store_true", help="既存リッチメニューを全削除")
    parser.add_argument("--list", action="store_true", help="リッチメニュー一覧を表示")
    parser.add_argument("--image", type=str, default=None, help="リッチメニュー画像パス")
    parser.add_argument("--youtube-url", type=str, default=None, help="YouTubeチャンネルURL")
    args = parser.parse_args()

    if args.list:
        list_rich_menus()
        return

    if args.delete_all:
        delete_all_rich_menus()
        return

    if args.youtube_url:
        RICH_MENU_BODY["areas"][0]["action"]["uri"] = args.youtube_url

    image_path = Path(args.image) if args.image else RICH_MENU_IMAGE

    print("=== きりたんチャット リッチメニューセットアップ ===\n")

    print("1. 既存リッチメニューを削除...")
    delete_all_rich_menus()

    print("\n2. 新しいリッチメニューを作成...")
    rich_menu_id = create_rich_menu()

    print("\n3. リッチメニュー画像をアップロード...")
    upload_rich_menu_image(rich_menu_id, image_path)

    print("\n4. デフォルトリッチメニューに設定...")
    set_default_rich_menu(rich_menu_id)

    print(f"\n=== 完了 ===")
    print(f"リッチメニューID: {rich_menu_id}")
    print(f"画像: {image_path}")
    print(f"\nボタン配置:")
    for area in RICH_MENU_BODY["areas"]:
        action = area["action"]
        action_detail = action.get("uri") or action.get("text", "")
        print(f"  [{area['action']['label']}] → {action['type']}: {action_detail}")


if __name__ == "__main__":
    main()
