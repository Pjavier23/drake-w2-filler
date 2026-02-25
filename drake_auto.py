#!/usr/bin/env python3
"""
Drake Full-Auto Engine
======================
Handles launching Drake, opening a client return by SSN,
navigating to the W-2 screen, and filling all fields.
"""

import os
import re
import time
import subprocess
import pyautogui
import pygetwindow as gw
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
DRAKE_EXE_PATHS = [
    r"D:\DRAKE25\Drake32.exe",
    r"D:\DRAKE25\Drake.exe",
    r"D:\DRAKE25\DrakeTax.exe",
    r"C:\Drake25\Drake32.exe",
]
DRAKE_WINDOW_KEYWORDS = ['Drake 2025', 'Drake Tax', 'Data Entry', 'Drake 20']
LAUNCH_WAIT    = 8     # seconds to wait for Drake to fully load after launch
NAV_DELAY      = 0.4   # seconds between navigation steps
TYPING_DELAY   = 0.15  # seconds between keystrokes when filling fields

pyautogui.FAILSAFE = True
pyautogui.PAUSE    = 0.05

# ── Drake Window ──────────────────────────────────────────────────────────────

def find_drake_window(keyword_filter=None):
    """Find Drake window. Optional keyword_filter to match specific screen."""
    for title in gw.getAllTitles():
        for kw in DRAKE_WINDOW_KEYWORDS:
            if kw.lower() in title.lower():
                if keyword_filter and keyword_filter.lower() not in title.lower():
                    continue
                wins = gw.getWindowsWithTitle(title)
                if wins:
                    return wins[0]
    return None

def focus_window(win):
    """Bring window to front."""
    try:
        if win.isMinimized:
            win.restore()
            time.sleep(0.5)
        win.activate()
        time.sleep(0.8)
        return True
    except Exception as e:
        print(f"  focus error: {e}")
        return False

def wait_for_window(keyword, timeout=15):
    """Wait until a Drake window containing keyword appears."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        win = find_drake_window(keyword)
        if win:
            return win
        time.sleep(0.5)
    return None

# ── Drake Launch ──────────────────────────────────────────────────────────────

def find_drake_exe():
    """Find the Drake executable."""
    for path in DRAKE_EXE_PATHS:
        if Path(path).exists():
            return path
    # Search D:\DRAKE25\ for any .exe
    base = Path(r"D:\DRAKE25")
    if base.exists():
        exes = list(base.glob("*.exe"))
        if exes:
            # Prefer Drake32 or Drake in name
            for exe in exes:
                if 'drake' in exe.name.lower():
                    return str(exe)
            return str(exes[0])
    return None

def launch_drake(log_fn=print):
    """Launch Drake if not already running. Returns True if ready."""
    # Already open?
    if find_drake_window():
        log_fn("  ✅ Drake already open")
        return True

    exe = find_drake_exe()
    if not exe:
        raise RuntimeError(
            "Cannot find Drake executable.\n"
            "Expected at D:\\DRAKE25\\Drake32.exe\n"
            "Please set the correct path in drake_auto.py → DRAKE_EXE_PATHS"
        )

    log_fn(f"  🚀 Launching Drake: {exe}")
    subprocess.Popen([exe])

    log_fn(f"  ⏳ Waiting for Drake to load ({LAUNCH_WAIT}s)...")
    win = wait_for_window('Drake', timeout=LAUNCH_WAIT + 10)
    if not win:
        raise RuntimeError("Drake launched but window did not appear. Try again.")

    time.sleep(2)  # Extra settle time
    log_fn("  ✅ Drake loaded")
    return True

# ── Open Client Return ────────────────────────────────────────────────────────

def open_return_by_ssn(ssn: str, log_fn=print):
    """
    Open a client return in Drake by SSN.
    Uses File → Open Return menu, types SSN, presses Enter.
    
    SSN format: 123-45-6789 or 123456789
    """
    if not ssn:
        raise ValueError("No SSN provided — cannot open client return automatically.")

    # Normalize SSN — remove dashes, handle masked (XXX-XX-1234)
    ssn_clean = re.sub(r'\D', '', ssn)
    if len(ssn_clean) != 9:
        # Masked SSN — can't auto-open, wait for user to do it manually
        log_fn(f"⚠️  SSN is masked ({ssn}) — cannot auto-open return.")
        log_fn("   Please open the client return in Drake manually.")
        log_fn("   Waiting 15 seconds...")
        time.sleep(15)
        return

    win = find_drake_window()
    if not win:
        raise RuntimeError("Drake window not found")

    focus_window(win)
    log_fn(f"  📂 Opening return for SSN: {ssn_clean[:3]}-{ssn_clean[3:5]}-{ssn_clean[5:]}")

    # File → Open Return (standard Drake shortcut)
    pyautogui.hotkey('alt', 'f')   # Open File menu
    time.sleep(NAV_DELAY)
    pyautogui.press('o')           # Open Return
    time.sleep(NAV_DELAY * 2)

    # Type SSN in the open dialog
    pyautogui.typewrite(ssn_clean, interval=0.05)
    time.sleep(NAV_DELAY)
    pyautogui.press('enter')
    time.sleep(2)  # Wait for return to load

    log_fn("  ✅ Return opened")

def navigate_to_w2_screen(log_fn=print):
    """
    Navigate to the W-2 data entry screen from an open return.
    Drake path: typically via keyboard shortcut or menu
    W-2 screen shortcut in Drake: Ctrl+W or via Screen Menu
    """
    win = find_drake_window()
    if not win:
        raise RuntimeError("Drake window not found")

    focus_window(win)
    log_fn("  🗺️  Navigating to W-2 screen...")

    # Try Drake's built-in W-2 shortcut (screen number 2 in income section)
    # In Drake, you can type a screen code to jump directly
    # W-2 is typically accessible by pressing W2 or via Screen Menu
    
    # Method 1: Type screen code directly (Drake accepts screen codes)
    pyautogui.hotkey('ctrl', 'home')  # Go to main return screen
    time.sleep(NAV_DELAY)
    
    # In Drake data entry, type "W2" to jump to W-2 screen
    pyautogui.typewrite('W2', interval=0.1)
    time.sleep(NAV_DELAY)
    pyautogui.press('enter')
    time.sleep(1)

    # Verify we're on the W-2 screen
    w2_win = wait_for_window('W-2', timeout=5)
    if w2_win:
        log_fn("  ✅ W-2 screen loaded")
        return True

    # Method 2: Try Screen Menu navigation
    log_fn("  ℹ️  Trying Screen Menu navigation...")
    pyautogui.hotkey('alt', 's')   # Screen Menu (if available)
    time.sleep(NAV_DELAY)

    log_fn("  ⚠️  W-2 screen navigation may need calibration")
    return True  # Continue anyway — user may already be on W-2 screen

# ── Full Auto Pipeline ────────────────────────────────────────────────────────

def run_full_auto(data: dict, fill_fn, log_fn=print):
    """
    Full auto pipeline:
    1. Launch Drake (if needed)
    2. Open client return by SSN
    3. Navigate to W-2 screen
    4. Fill W-2 data
    """
    ssn = data.get('employee_ssn', '')

    # Step 1: Launch Drake
    log_fn("🚀 Step 1: Launching Drake...")
    launch_drake(log_fn)

    # Step 2: Open client return
    if ssn:
        log_fn("📂 Step 2: Opening client return...")
        open_return_by_ssn(ssn, log_fn)
    else:
        log_fn("⚠️  Step 2: No SSN found — skipping auto-open")
        log_fn("   Please open the client return manually in Drake")
        log_fn("   Waiting 10 seconds for you to do that...")
        time.sleep(10)

    # Step 3: Navigate to W-2 screen
    log_fn("🗺️  Step 3: Navigating to W-2 screen...")
    navigate_to_w2_screen(log_fn)

    # Step 4: Fill W-2 data
    log_fn("⌨️  Step 4: Filling W-2 data...")
    fill_fn(data)

    log_fn("✅ Full auto complete!")
