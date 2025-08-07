"""
Tab Switch Test Suite for VOICEROID+ "Phrase Edit" Tab
Usage: python tab_switch_test.py

This script contains the following steps:
 1) Enumerate all top‑level windows via win32gui
 2) Find the VOICEROID window handle and PID
 3) Connect via UIA backend by PID and HWND (better for TabItem enumeration)
 4) Locate all TabItem elements and print their names
 5) Select the "フレーズ編集" tab by name
"""
import time
import win32gui
import win32process
from pywinauto import Application, timings


def enum_windows():
    print("=== Step 1: Enumerate Window Titles ===")
    def callback(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if title:
            print(f"hwnd={hwnd}, title='{title}'")
    win32gui.EnumWindows(callback, None)


def find_voiceroid_handle():
    print("=== Step 2: Find VOICEROID window and PID ===")
    target = 'VOICEROID＋ 東北きりたん EX'
    hwnd = win32gui.FindWindow(None, target)
    if not hwnd:
        print(f"✖️ Window '{target}' not found")
        return None, None
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    print(f"◎ Found hwnd={hwnd}, PID={pid}")
    return hwnd, pid


def connect_by_pid_hwnd(pid, hwnd):
    print("=== Step 3: Connect via UIA backend by PID and HWND ===")
    try:
        app = timings.wait_until_passes(
            5, 0.5,
            lambda: Application(backend='uia').connect(process=pid, visible_only=False)
        )
        win = app.window(handle=hwnd)
        print("◎ Connected (UIA) to VOICEROID process with HWND")
        return win
    except Exception as e:
        print(f"✖️ Connect failed: {e}")
        return None


def list_tab_items(win):
    print("=== Step 4: Locate all TabItem elements ===")
    try:
        items = win.descendants(control_type='TabItem')
        print(f"◎ Found {len(items)} TabItem(s):")
        for tab in items:
            info = tab.element_info
            print(f"  - '{info.name}' (auto_id='{info.automation_id}')")
        return items
    except Exception as e:
        print(f"✖️ Failed to list TabItems: {e}")
        return []


def switch_tab(win, items):
    """Step 5: Select the 'フレーズ編集' tab with multiple fallbacks"""
    print("=== Step 5: Select the 'フレーズ編集' tab ===")
    # Ensure window has focus
    try:
        win.set_focus()
    except Exception:
        pass
    for tab in items:
        if tab.element_info.name == 'フレーズ編集':
            # try select pattern
            try:
                tab.select()
                print("◎ 'フレーズ編集' tab selected via select()")
                return
            except Exception:
                pass
            # try invoke pattern
            try:
                tab.invoke()
                print("◎ 'フレーズ編集' tab selected via invoke()")
                return
            except Exception:
                pass
            # fallback to mouse click
            try:
                tab.click_input()
                print("◎ 'フレーズ編集' tab clicked successfully")
                return
            except Exception as e:
                print(f"✖️ クリックでの選択に失敗: {e}")
                return
    print("✖️ 'フレーズ編集' tab not found among TabItems")

if __name__ == '__main__':
    enum_windows()
    hwnd, pid = find_voiceroid_handle()
    if pid:
        win = connect_by_pid_hwnd(pid, hwnd)
        if win:
            items = list_tab_items(win)
            switch_tab(win, items)
