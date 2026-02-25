#!/usr/bin/env python3
"""
Drake W-2 Auto-Filler  v2.0  — Full Auto Mode
==============================================
Drop W-2 PDFs into the inbox folder.
Script extracts data, launches Drake, opens the client return,
navigates to the W-2 screen, and fills everything automatically.

Modes:
  FULL AUTO  — launches Drake, opens return, navigates, fills
  MANUAL     — Drake already open on W-2 screen, just fills

Requirements:
  pip install pyautogui pyperclip pdfplumber watchdog pytesseract pillow psutil pygetwindow
"""

import os
import re
import time
import json
import shutil
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
from pathlib import Path

import pyautogui
import pyperclip
import pdfplumber
import psutil
import pygetwindow as gw
from drake_auto import run_full_auto, launch_drake, open_return_by_ssn, navigate_to_w2_screen

# ── Config ────────────────────────────────────────────────────────────────────
WATCH_FOLDER = r"C:\W2_Inbox"
DONE_FOLDER  = r"C:\W2_Done"
ERROR_FOLDER = r"C:\W2_Errors"
KEYSTROKE_DELAY = 0.25   # seconds between Tab presses (increase if Drake is slow)
FILL_DELAY      = 0.15   # seconds after typing a value

pyautogui.FAILSAFE = True   # Emergency stop: slam mouse to top-left corner
pyautogui.PAUSE    = 0.05

# ── W-2 Field Extraction ──────────────────────────────────────────────────────

def clean_amount(val: str) -> str:
    """Strip $ signs, commas, spaces from a dollar amount."""
    return re.sub(r'[$,\s]', '', val or '').strip()

def extract_w2_from_pdf(pdf_path: str) -> dict:
    """
    Extract W-2 fields using three independent strategies:
    1. pdfplumber table extraction
    2. pdfplumber word-level extraction
    3. Multi-pattern regex on raw text
    Takes the best result from each strategy.
    Never crashes — always returns whatever was found.
    """
    data = {
        'employee_ssn': '', 'ein': '',
        'employer_name': '', 'employer_street': '',
        'employer_city': '', 'employer_state': '', 'employer_zip': '',
        'box1_wages': '', 'box2_fed_tax': '', 'box3_ss_wages': '',
        'box4_ss_tax': '', 'box5_med_wages': '', 'box6_med_tax': '',
        'box7_ss_tips': '', 'box8_alloc': '', 'box10_dep_care': '',
        'box11_nonqual': '', 'box12': [], 'box13_statutory': False,
        'box13_retirement': False, 'box13_sick_pay': False, 'box14': [],
        'box15_state': '', 'box15_state_id': '',
        'box16_wages': '', 'box17_tax': ''
    }

    text = ""
    words_by_page = []

    # ── Extract text and words ────────────────────────────────────────────────
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
                try:
                    words = page.extract_words(keep_blank_chars=False, x_tolerance=3, y_tolerance=3)
                    words_by_page.append({'words': words, 'width': page.width, 'height': page.height})
                except Exception:
                    pass
    except Exception as e:
        print(f"  pdfplumber error: {e}")

    # OCR fallback for scanned PDFs
    if len(text.strip()) < 50:
        try:
            import pytesseract
            from pdf2image import convert_from_path
            print("  Falling back to OCR...")
            images = convert_from_path(pdf_path, dpi=300)
            for img in images:
                text += pytesseract.image_to_string(img) + "\n"
        except Exception as e:
            print(f"  OCR fallback error: {e}")

    if not text.strip():
        print("  ⚠️  Could not extract any text from PDF")
        return data

    # Log raw text for debugging
    print(f"\n{'='*60}\nRAW TEXT (first 3000 chars):\n{'='*60}")
    print(text[:3000])
    print('='*60 + "\n")

    # ── Strategy 1: Word-position based extraction ────────────────────────────
    # W-2 box amounts are in standardized positions on the page
    # Right half of page = wage/tax amounts; left half = labels/identifiers
    def extract_by_position(words_data):
        if not words_data:
            return {}
        result = {}
        for page_data in words_data:
            words = page_data['words']
            w     = page_data['width']
            h     = page_data['height']
            # Build a lookup: text → (x0, top, x1, bottom)
            word_map = [(wd['text'], wd['x0'], wd['top'], wd['x1'], wd['bottom']) for wd in words]

            def near_label(label_text, dx=300, dy=20):
                """Find dollar amount within dx/dy pixels of a label."""
                for i, (txt, x0, top, x1, bot) in enumerate(word_map):
                    if label_text.lower() in txt.lower():
                        # Search nearby words for a dollar amount
                        for j, (t2, x2, t2_top, x2_1, t2_bot) in enumerate(word_map):
                            if abs(t2_top - top) < dy and x2 > x0 and x2 < x0 + dx:
                                cleaned = clean_amount(t2)
                                if re.match(r'^\d[\d,]*(?:\.\d{2})?$', cleaned):
                                    return cleaned
                        # Also check next few words
                        for j in range(i+1, min(i+5, len(word_map))):
                            t2 = word_map[j][0]
                            cleaned = clean_amount(t2)
                            if re.match(r'^\d[\d,]*(?:\.\d{2})?$', cleaned) and float(cleaned.replace(',','')) > 0:
                                return cleaned
                return ''

            result['box1_wages']    = result.get('box1_wages')    or near_label('Wages, tips')
            result['box2_fed_tax']  = result.get('box2_fed_tax')  or near_label('Federal income tax')
            result['box3_ss_wages'] = result.get('box3_ss_wages') or near_label('Social security wages')
            result['box4_ss_tax']   = result.get('box4_ss_tax')   or near_label('Social security tax with')
            result['box5_med_wages']= result.get('box5_med_wages')or near_label('Medicare wages')
            result['box6_med_tax']  = result.get('box6_med_tax')  or near_label('Medicare tax with')
        return {k: v for k, v in result.items() if v}

    pos_data = extract_by_position(words_by_page)

    # ── Strategy 2: Regex on raw text ────────────────────────────────────────
    def amt(pattern):
        """Find first dollar amount matching pattern."""
        m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if m:
            return clean_amount(m.group(1))
        return ''

    def amt_multi(*patterns):
        for p in patterns:
            v = amt(p)
            if v:
                return v
        return ''

    # SSN — full or partial/masked
    ssn_m = re.search(r'\b(\d{3}-\d{2}-\d{4})\b', text)
    if not ssn_m:
        ssn_m = re.search(r'\b([Xx*\d]{3}-[Xx*\d]{2}-\d{4})\b', text)
    data['employee_ssn'] = ssn_m.group(1) if ssn_m else ''

    # EIN — always XX-XXXXXXX format
    ein_m = re.search(r'\b(\d{2}-\d{7})\b', text)
    data['ein'] = ein_m.group(1) if ein_m else ''

    # Employer name — first all-caps line with 3+ chars after EIN or "employer name"
    name_m = re.search(
        r'(?:employer.{0,20}name|EIN[\s\S]{0,100}?\n)[\s\S]{0,5}\n([A-Z][A-Z0-9 &.,\'-]{3,})\n',
        text, re.IGNORECASE
    )
    if not name_m:
        # Try: largest all-caps word cluster on page
        name_m = re.search(r'\n([A-Z]{2}[A-Z0-9 &.,\'-]{3,}(?:LLC|INC|CORP|CO\.?|LTD)?)\n', text)
    data['employer_name'] = name_m.group(1).strip() if name_m else ''

    # Address
    street_m = re.search(r'\b(\d{2,5}\s+[A-Z][A-Za-z0-9 ]+(?:ST|AVE|BLVD|DR|RD|LN|WAY|PKWY|CT|CIR|HWY|SUITE?)[A-Za-z0-9 .]*)\b', text, re.IGNORECASE)
    data['employer_street'] = street_m.group(1).strip() if street_m else ''

    city_m = re.search(r'([A-Z][A-Za-z ]{2,}),?\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)', text)
    if city_m:
        data['employer_city']  = city_m.group(1).strip()
        data['employer_state'] = city_m.group(2).strip()
        data['employer_zip']   = city_m.group(3).strip()

    # Wages — try many patterns
    data['box1_wages'] = amt_multi(
        r'(?:^|\s)1\s+Wages[,\s]+tips[^\d\n]{0,80}?([\d,]+\.\d{2})',
        r'Wages,\s*tips[^\d\n]{0,100}([\d,]+\.\d{2})',
        r'Wages[^\d\n]{0,60}([\d,]+\.\d{2})',
    )
    data['box2_fed_tax'] = amt_multi(
        r'(?:^|\s)2\s+Federal[^\d\n]{0,80}?([\d,]+\.\d{2})',
        r'Federal\s+income\s+tax\s+withheld[^\d\n]{0,60}([\d,]+\.\d{2})',
        r'Federal\s+income\s+tax[^\d\n]{0,60}([\d,]+\.\d{2})',
    )
    data['box3_ss_wages'] = amt_multi(
        r'(?:^|\s)3\s+Social\s+security\s+wages[^\d\n]{0,60}([\d,]+\.\d{2})',
        r'Social\s+security\s+wages[^\d\n]{0,60}([\d,]+\.\d{2})',
    )
    data['box4_ss_tax'] = amt_multi(
        r'(?:^|\s)4\s+Social\s+security\s+tax[^\d\n]{0,60}([\d,]+\.\d{2})',
        r'Social\s+security\s+tax\s+withheld[^\d\n]{0,60}([\d,]+\.\d{2})',
    )
    data['box5_med_wages'] = amt_multi(
        r'(?:^|\s)5\s+Medicare\s+wages[^\d\n]{0,60}([\d,]+\.\d{2})',
        r'Medicare\s+wages[^\d\n]{0,60}([\d,]+\.\d{2})',
    )
    data['box6_med_tax'] = amt_multi(
        r'(?:^|\s)6\s+Medicare\s+tax[^\d\n]{0,60}([\d,]+\.\d{2})',
        r'Medicare\s+tax\s+withheld[^\d\n]{0,60}([\d,]+\.\d{2})',
    )
    data['box7_ss_tips']   = amt_multi(r'Social\s+security\s+tips[^\d\n]{0,60}([\d,]+\.\d{2})')
    data['box10_dep_care'] = amt_multi(r'Dependent\s+care[^\d\n]{0,60}([\d,]+\.\d{2})')

    # Box 12 codes — e.g. "12a Code D  1234.56"
    data['box12'] = []
    for m in re.finditer(r'\b(?:12[a-d]?\s+)?([A-Z]{1,2})\s+([\d,]+\.\d{2})\b', text):
        code = m.group(1).upper()
        # Only valid W-2 box 12 codes (A-HH range, common ones)
        valid_codes = {'A','B','C','D','E','F','G','H','J','K','L','M','N','P','Q','R','S','T','V','W','Y','Z','AA','BB','DD','EE','FF','GG','HH'}
        if code in valid_codes:
            amt_val = clean_amount(m.group(2))
            if float(amt_val.replace(',', '')) > 0:
                data['box12'].append({'code': code, 'amount': amt_val})

    # Box 14 — real codes like "SDDI 123.45" or "CA SDI 234.56"
    data['box14'] = []
    bad_words = {'STREET','AVE','BLVD','RD','DR','LN','SUITE','PAGE','FORM','COPY','DEPT','BOX'}
    for m in re.finditer(r'^([A-Z][A-Z0-9 ]{1,15})\s+([\d,]+\.\d{2})$', text, re.MULTILINE):
        label = m.group(1).strip()
        if not any(w in label.upper() for w in bad_words) and len(label) <= 15:
            data['box14'].append({'label': label, 'amount': clean_amount(m.group(2))})

    # State (15-20)
    state_sec = re.search(
        r'15[^\n]{0,5}([A-Z]{2})[^\n]{0,40}?([\w-]{3,})[^\n]{0,60}?([\d,]+\.\d{2})[^\n]{0,60}?([\d,]+\.\d{2})',
        text, re.IGNORECASE
    )
    if state_sec:
        data['box15_state']    = state_sec.group(1)
        data['box15_state_id'] = state_sec.group(2)
        data['box16_wages']    = clean_amount(state_sec.group(3))
        data['box17_tax']      = clean_amount(state_sec.group(4))
    else:
        data['box15_state']    = data.get('employer_state', '')
        data['box16_wages']    = data.get('box1_wages', '')

    # ── Merge position-based results (fill gaps) ──────────────────────────────
    for k, v in pos_data.items():
        if not data.get(k) and v:
            data[k] = v
            print(f"  📍 Position extraction filled: {k} = {v}")

    # ── Strategy 3: Last resort — find all dollar amounts by order ────────────
    # If wages still missing, grab largest dollar amount in doc as Box 1
    if not data['box1_wages']:
        all_amounts = re.findall(r'\b(\d{1,3}(?:,\d{3})+\.\d{2}|\d{4,6}\.\d{2})\b', text)
        if all_amounts:
            # Sort descending — largest is likely gross wages
            sorted_amts = sorted(all_amounts, key=lambda x: float(x.replace(',','')), reverse=True)
            # Filter to reasonable wage range ($1,000 - $500,000)
            wages = [a for a in sorted_amts if 1000 <= float(a.replace(',','')) <= 500000]
            if wages:
                data['box1_wages'] = wages[0]
                print(f"  🔍 Last-resort: Box 1 wages = {data['box1_wages']}")
            # Box 3 (SS wages) usually = Box 1 if not found separately
            if not data['box3_ss_wages'] and data['box1_wages']:
                data['box3_ss_wages'] = data['box1_wages']
            if not data['box5_med_wages'] and data['box1_wages']:
                data['box5_med_wages'] = data['box1_wages']

    # ── Log results ───────────────────────────────────────────────────────────
    print("\n── Extraction Results ──────────────────────────")
    key_fields = ['employee_ssn','ein','employer_name','box1_wages','box2_fed_tax',
                  'box3_ss_wages','box4_ss_tax','box5_med_wages','box6_med_tax',
                  'box15_state','box16_wages','box17_tax']
    for k in key_fields:
        v = data.get(k, '')
        icon = '✅' if v else '❌'
        print(f"  {icon} {k}: {v or '—'}")
    print("─────────────────────────────────────────────\n")

    return data


# ── Drake Detection & Focus ───────────────────────────────────────────────────

DRAKE_PROCESS_NAMES  = ['Drake32.exe', 'Drake.exe', 'DrakeTax.exe', 'drake32.exe']
DRAKE_WINDOW_KEYWORDS = ['Drake 2025', 'Drake Tax', 'Data Entry', 'Drake 20']
CALIBRATION_FILE     = Path(__file__).parent / 'calibration.json'

def is_drake_running() -> bool:
    """Check if any Drake process is running."""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] in DRAKE_PROCESS_NAMES:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False

def find_drake_window():
    """Find the Drake Tax window. Returns the window or None."""
    all_windows = gw.getAllTitles()
    for title in all_windows:
        for keyword in DRAKE_WINDOW_KEYWORDS:
            if keyword.lower() in title.lower():
                windows = gw.getWindowsWithTitle(title)
                if windows:
                    return windows[0]
    return None

def focus_drake_window() -> bool:
    """
    Find Drake window, bring it to front, maximize if needed.
    Returns True if successful.
    """
    win = find_drake_window()
    if not win:
        return False
    try:
        if win.isMinimized:
            win.restore()
            time.sleep(0.5)
        win.activate()
        time.sleep(0.8)   # give Windows time to actually bring it forward
        return True
    except Exception as e:
        print(f"  Window focus error: {e}")
        return False

def save_calibration(x: int, y: int):
    """
    Save EIN field position as offset relative to Drake window top-left.
    This means calibration stays valid even if Drake moves to a different
    position or monitor.
    """
    win = find_drake_window()
    if win:
        offset_x = x - win.left
        offset_y = y - win.top
        with open(CALIBRATION_FILE, 'w') as f:
            json.dump({'offset_x': offset_x, 'offset_y': offset_y, 'abs_x': x, 'abs_y': y}, f)
        print(f"  💾 Calibration saved: EIN offset from Drake window = ({offset_x}, {offset_y})")
        print(f"       Works even if Drake moves to a different position or monitor.")
    else:
        # Fallback: save absolute if Drake window not found
        with open(CALIBRATION_FILE, 'w') as f:
            json.dump({'offset_x': None, 'offset_y': None, 'abs_x': x, 'abs_y': y}, f)
        print(f"  💾 Calibration saved (absolute): EIN at ({x}, {y})")

def load_calibration() -> tuple:
    """
    Load EIN field position.
    Returns absolute (x, y) by adding saved offset to current Drake window position.
    Falls back to saved absolute coords if Drake window not found.
    """
    if not CALIBRATION_FILE.exists():
        return None, None
    try:
        with open(CALIBRATION_FILE) as f:
            d = json.load(f)
        offset_x = d.get('offset_x')
        offset_y = d.get('offset_y')
        if offset_x is not None and offset_y is not None:
            win = find_drake_window()
            if win:
                return win.left + offset_x, win.top + offset_y
        # Fallback to absolute
        return d.get('abs_x'), d.get('abs_y')
    except Exception:
        pass
    return None, None

def run_calibration(log_fn=print):
    """
    Interactive calibration: user hovers over EIN field, presses Enter.
    Saves coordinates for all future runs.
    """
    log_fn("📍 CALIBRATION MODE")
    log_fn("   1. Switch to Drake and click the EIN field")
    log_fn("   2. Leave your mouse ON the EIN field")
    log_fn("   3. Switch back here and click 'Save Position'")
    log_fn("   (You have 5 seconds after clicking Save Position)")

def click_ein_field() -> bool:
    """
    Click the EIN field using the best available method, in priority order:
      1. Saved calibration coordinates (fastest, most reliable)
      2. Image recognition using ein_label.png reference
      3. Estimated position relative to Drake window (last resort)
    Returns True if successfully clicked.
    """
    # Priority 1: Saved calibration
    x, y = load_calibration()
    if x and y:
        pyautogui.click(x, y)
        time.sleep(0.3)
        print(f"  ✅ EIN field clicked from calibration ({x}, {y})")
        return True

    # Priority 2: Image recognition
    try:
        ein_loc = pyautogui.locateOnScreen(
            str(Path(__file__).parent / 'assets' / 'ein_label.png'),
            confidence=0.7
        )
        if ein_loc:
            # Click to the right of the label where the input box is
            pyautogui.click(ein_loc.left + ein_loc.width + 120, ein_loc.top + ein_loc.height // 2)
            time.sleep(0.3)
            print("  ✅ EIN field found via image recognition")
            return True
    except Exception:
        pass

    # Priority 3: Estimated position relative to Drake window
    win = find_drake_window()
    if win:
        x = win.left + int(win.width * 0.35)
        y = win.top  + int(win.height * 0.30)
        pyautogui.click(x, y)
        time.sleep(0.3)
        print(f"  ⚠️  Used estimated EIN position ({x}, {y}) — run Calibrate for accuracy")
        return True

    return False


# ── Drake Screen Filler ───────────────────────────────────────────────────────

def tab(n=1):
    """Press Tab n times."""
    for _ in range(n):
        pyautogui.press('tab')
        time.sleep(KEYSTROKE_DELAY)

def type_field(value: str, then_tab: bool = True):
    """Select all, paste value, then Tab."""
    if value:
        pyperclip.copy(str(value))
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(FILL_DELAY)
    if then_tab:
        tab()

def skip(n=1):
    """Tab past n fields without changing them."""
    tab(n)

def fill_drake_w2_screen(data: dict):
    """
    Fill the Drake W-2 entry screen.

    Tab order on the Drake W-2 screen (confirmed from screenshot):
      [TS dropdown] [F checkbox] [Special tax treatment]
      EIN → Name → Name cont → Street → City → State → ZIP
      Box1 → Box2 → Box3 → Box4 → Box5 → Box6 → Box7 → Box8 → Box9 → Box10
      Employee First → Employee Last → Employee Street → Employee City → Employee State → Employee ZIP
      Box11 → Box12(row1: Code, Amount, Year) × 3 rows
      Box13 checkboxes (Statutory, Retirement, Sick pay)
      Box14(row1: Label, Amount) × 2 rows
      Box15(ST, StateID) → Box16 → Box17 → Box18 → Box19 → Box20 × 2 rows
    """

    # ── Step 1: Check Drake window is visible (more reliable than process name check)
    if not find_drake_window():
        raise RuntimeError(
            "Could not find the Drake Tax window.\n"
            "Make sure Drake is open with a return loaded and the W-2 screen visible."
        )

    # Countdown is handled by the caller (process_pdf) before this function runs
    print("  ⌨️  Starting field fill...")

    # Navigate to EIN field precisely regardless of where cursor is
    # Ctrl+Home goes to first field (TS), then Tab x3 lands on EIN
    pyautogui.hotkey('ctrl', 'home')
    time.sleep(0.3)
    tab(3)  # Skip: TS dropdown → F checkbox → Special tax treatment → lands on EIN

    # ── EIN
    type_field(data.get('ein', ''))

    # ── Employer Name
    type_field(data.get('employer_name', ''))

    # ── Name cont (skip)
    skip(1)

    # ── Street
    type_field(data.get('employer_street', ''))

    # ── City
    type_field(data.get('employer_city', ''))

    # ── State (2-letter dropdown — just type it)
    type_field(data.get('employer_state', ''))

    # ── ZIP
    type_field(data.get('employer_zip', ''))

    # ── Box 1 — Wages
    type_field(data.get('box1_wages', ''))

    # ── Box 2 — Federal tax w/h
    type_field(data.get('box2_fed_tax', ''))

    # ── Box 3 — SS wages
    type_field(data.get('box3_ss_wages', ''))

    # ── Box 4 — SS w/h
    type_field(data.get('box4_ss_tax', ''))

    # ── Box 5 — Medicare wages
    type_field(data.get('box5_med_wages', ''))

    # ── Box 6 — Medicare tax w/h
    type_field(data.get('box6_med_tax', ''))

    # ── Box 7 — SS tips
    type_field(data.get('box7_ss_tips', ''))

    # ── Box 8 — Allocated tips
    type_field(data.get('box8_alloc', ''))

    # ── Box 9 (skip — reserved)
    skip(1)

    # ── Box 10 — Dep care benefit
    type_field(data.get('box10_dep_care', ''))

    # ── Employee name/address (skip — same as screen 1 usually)
    skip(6)

    # ── Box 11 — Nonqualified plan
    type_field(data.get('box11_nonqual', ''))

    # ── Box 12 rows (3 rows × Code + Amount + Year)
    box12 = data.get('box12', [])
    for i in range(3):
        if i < len(box12):
            type_field(box12[i].get('code', ''))
            type_field(box12[i].get('amount', ''))
            skip(1)  # Year
        else:
            skip(3)

    # ── Box 13 checkboxes — Statutory, Retirement, Sick pay
    # Checkboxes: press Space to check, Tab to move on
    def checkbox(checked: bool):
        if checked:
            pyautogui.press('space')
            time.sleep(FILL_DELAY)
        tab(1)

    checkbox(data.get('box13_statutory', False))
    checkbox(data.get('box13_retirement', False))
    checkbox(data.get('box13_sick_pay', False))

    # ── Box 14 — Other (2 rows: label + amount)
    box14 = data.get('box14', [])
    for i in range(2):
        if i < len(box14):
            type_field(box14[i].get('label', ''))
            type_field(box14[i].get('amount', ''))
        else:
            skip(2)

    # ── Box 15-20 — State row 1
    type_field(data.get('box15_state', ''))
    type_field(data.get('box15_state_id', ''))
    type_field(data.get('box16_wages', ''))
    type_field(data.get('box17_tax', ''))
    skip(2)  # Box 18, 19, 20 (local — skip unless provided)

    print("✅ Fill complete! Review fields in Drake before saving.")


# ── GUI ───────────────────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("Drake W-2 Auto-Filler")
        root.geometry("600x520")
        root.configure(bg="#1a1a1a")
        root.resizable(False, False)

        # Header
        tk.Label(root, text="🦅 Drake W-2 Auto-Filler",
                 font=("Segoe UI", 18, "bold"),
                 fg="#00d9ff", bg="#1a1a1a").pack(pady=(20, 4))

        # Mode toggle
        self.full_auto = tk.BooleanVar(value=True)
        mode_frame = tk.Frame(root, bg="#1a1a1a")
        mode_frame.pack(pady=(4, 0))
        tk.Label(mode_frame, text="Mode:", font=("Segoe UI", 10),
                 fg="#888", bg="#1a1a1a").pack(side="left", padx=(0, 6))
        tk.Radiobutton(mode_frame, text="🤖 Full Auto  (opens Drake + return automatically)",
                       variable=self.full_auto, value=True,
                       font=("Segoe UI", 9), fg="#00ff88", bg="#1a1a1a",
                       selectcolor="#0f0f0f", activebackground="#1a1a1a").pack(side="left")
        tk.Radiobutton(mode_frame, text="✋ Manual  (Drake already open on W-2 screen)",
                       variable=self.full_auto, value=False,
                       font=("Segoe UI", 9), fg="#ffcc00", bg="#1a1a1a",
                       selectcolor="#0f0f0f", activebackground="#1a1a1a").pack(side="left", padx=(12, 0))

        # Log area
        self.log = scrolledtext.ScrolledText(
            root, height=16, font=("Consolas", 9),
            bg="#0f0f0f", fg="#ccc", insertbackground="white",
            relief="flat", bd=0, padx=8, pady=8
        )
        self.log.pack(fill="both", padx=16, pady=12, expand=True)

        # Buttons
        btn_frame = tk.Frame(root, bg="#1a1a1a")
        btn_frame.pack(pady=8)

        self.start_btn = tk.Button(
            btn_frame, text="▶  Start Watching",
            font=("Segoe UI", 10, "bold"),
            bg="#00d9ff", fg="#000", relief="flat",
            padx=20, pady=8, cursor="hand2",
            command=self.start_watching
        )
        self.start_btn.pack(side="left", padx=8)

        tk.Button(
            btn_frame, text="📂  Open Inbox",
            font=("Segoe UI", 10),
            bg="#2a2a2a", fg="#ccc", relief="flat",
            padx=20, pady=8, cursor="hand2",
            command=lambda: os.startfile(WATCH_FOLDER)
        ).pack(side="left", padx=8)

        tk.Button(
            btn_frame, text="🧪  Test Extract",
            font=("Segoe UI", 10),
            bg="#2a2a2a", fg="#ccc", relief="flat",
            padx=20, pady=8, cursor="hand2",
            command=self.test_extract
        ).pack(side="left", padx=8)

        tk.Button(
            btn_frame, text="📍  Calibrate",
            font=("Segoe UI", 10),
            bg="#2a2a2a", fg="#ffcc00", relief="flat",
            padx=20, pady=8, cursor="hand2",
            command=self.calibrate
        ).pack(side="left", padx=8)

        tk.Label(root, text="⚠  Emergency stop: slam mouse to top-left corner",
                 font=("Segoe UI", 8), fg="#ff6b6b", bg="#1a1a1a").pack(pady=(0, 10))

        self.observer = None
        self.log_msg("Ready. Click 'Start Watching' to begin.\n")

    def log_msg(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def start_watching(self):
        if self.observer and self.observer.is_alive():
            self.log_msg("Already watching.")
            return

        for folder in [WATCH_FOLDER, DONE_FOLDER, ERROR_FOLDER]:
            Path(folder).mkdir(parents=True, exist_ok=True)

        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler

        app = self

        class Handler(FileSystemEventHandler):
            def on_created(self, event):
                if event.is_directory:
                    return
                if event.src_path.lower().endswith('.pdf'):
                    time.sleep(1)
                    threading.Thread(target=app.process_pdf, args=(event.src_path,), daemon=True).start()

        self.observer = Observer()
        self.observer.schedule(Handler(), WATCH_FOLDER, recursive=False)
        self.observer.start()

        self.start_btn.config(text="✅  Watching...", state="disabled", bg="#005566")
        self.log_msg(f"👀 Watching: {WATCH_FOLDER}")
        self.log_msg(f"   Inbox  → {WATCH_FOLDER}")
        self.log_msg(f"   Done   → {DONE_FOLDER}")
        self.log_msg(f"   Errors → {ERROR_FOLDER}\n")
        self.log_msg("Drop a W-2 PDF into the inbox folder to start.\n")

    def process_pdf(self, pdf_path: str):
        try:
            self.log_msg(f"📄 PDF detected: {Path(pdf_path).name}")
            data = extract_w2_from_pdf(pdf_path)
            self.log_msg(f"📊 Extracted data:\n{json.dumps(data, indent=2)}\n")

            # Show confirmation dialog on main thread
            confirmed = [False]

            def ask():
                summary = (
                    f"Ready to fill Drake W-2\n\n"
                    + "\n".join([
                        f"{'✅' if data.get(k) else '❌'} {lbl}: {data.get(k) or '—'}"
                        for k, lbl in [
                            ('employee_ssn',  'SSN         '),
                            ('ein',           'EIN         '),
                            ('employer_name', 'Employer    '),
                            ('box1_wages',    'Box 1 Wages '),
                            ('box2_fed_tax',  'Box 2 FedTax'),
                            ('box3_ss_wages', 'Box 3 SS Wg '),
                            ('box4_ss_tax',   'Box 4 SS Tax'),
                            ('box5_med_wages','Box 5 Med Wg'),
                            ('box6_med_tax',  'Box 6 Med Tx'),
                            ('box15_state',   'State       '),
                        ]
                    ]) + "\n\n"
                    f"─────────────────────────────\n" +
                    (
                        f"FULL AUTO: Will open Drake, find client\nby SSN, navigate to W-2, and fill.\nJust click YES and stay out of the way."
                        if self.full_auto.get() else
                        f"MANUAL: Click YES, then you have\n5 seconds to click the EIN field in Drake."
                    ) +
                    f"\n─────────────────────────────\n\n"
                    f"Click YES to fill, NO to skip."
                )
                confirmed[0] = messagebox.askyesno("Confirm Fill", summary)
                dialog_done.set()

            dialog_done = threading.Event()
            self.root.after(0, ask)
            dialog_done.wait(timeout=120)  # wait up to 2 min for user to click YES/NO

            if confirmed[0]:
                if self.full_auto.get():
                    # ── FULL AUTO MODE ──────────────────────────────────────
                    self.log_msg("🤖 Full Auto mode — taking control...")
                    run_full_auto(data, fill_drake_w2_screen, log_fn=self.log_msg)
                else:
                    # ── MANUAL MODE ─────────────────────────────────────────
                    self.log_msg("⏳ 5 seconds — click ANYWHERE on the Drake W-2 screen!")
                    for i in range(5, 0, -1):
                        self.log_msg(f"   {i}...")
                        time.sleep(1)
                    self.log_msg("⌨️  Typing...")
                    fill_drake_w2_screen(data)
                self.log_msg("✅ Drake fill complete!\n")

                done_path = Path(DONE_FOLDER) / Path(pdf_path).name
                shutil.move(pdf_path, done_path)
                self.log_msg(f"📁 Moved to Done: {done_path.name}\n")
            else:
                self.log_msg("⏭  Skipped (user cancelled)\n")

        except RuntimeError as e:
            self.log_msg(f"⚠️  {e}\n")
            # Don't move to errors — let user fix and re-drop
        except Exception as e:
            self.log_msg(f"❌ Error: {e}\n")
            try:
                err_path = Path(ERROR_FOLDER) / Path(pdf_path).name
                shutil.move(pdf_path, err_path)
            except Exception:
                pass

    def calibrate(self):
        """
        Calibration: user moves mouse to EIN field in Drake,
        then we capture the position.
        """
        msg = (
            "CALIBRATION — Do this once:\n\n"
            "1. Click OK\n"
            "2. Switch to Drake (you have 5 seconds)\n"
            "3. Hover your mouse over the EIN input field\n"
            "4. HOLD STILL — script captures position automatically\n\n"
            "Click OK to start the 5-second countdown."
        )
        if not messagebox.askokcancel("Calibrate EIN Field", msg):
            return

        self.log_msg("📍 Calibrating... switch to Drake and hover over EIN field")

        def do_capture():
            for i in range(5, 0, -1):
                self.log_msg(f"   Capturing in {i}...")
                time.sleep(1)
            x, y = pyautogui.position()
            save_calibration(x, y)
            self.log_msg(f"✅ Calibration saved! EIN field = ({x}, {y})")
            self.log_msg("   All future fills will click this exact position.\n")
            messagebox.showinfo("Calibration Complete", f"EIN field saved at ({x}, {y})\nYou're all set!")

        threading.Thread(target=do_capture, daemon=True).start()

    def test_extract(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Pick a W-2 PDF to test",
            filetypes=[("PDF files", "*.pdf")]
        )
        if path:
            threading.Thread(target=self._run_test, args=(path,), daemon=True).start()

    def _run_test(self, path):
        try:
            self.log_msg(f"🧪 Testing extraction: {Path(path).name}")
            data = extract_w2_from_pdf(path)
            self.log_msg(f"Result:\n{json.dumps(data, indent=2)}\n")
        except Exception as e:
            self.log_msg(f"❌ Error: {e}\n")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
