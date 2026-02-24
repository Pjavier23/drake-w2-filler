#!/usr/bin/env python3
"""
Drake W-2 Auto-Filler
=====================
Watches a folder for W-2 PDFs, extracts field data via OCR,
then fills Drake Tax's W-2 entry screen using keyboard automation.

How it works:
  1. Run this script (keep it running in background)
  2. Open Drake, open the client return, navigate to the W-2 entry screen
  3. Drop a W-2 PDF into C:\W2_Inbox\
  4. Script detects the PDF, reads the data, checks Drake is open
  5. Confirmation popup shows extracted data — click YES
  6. Script auto-focuses the Drake window, finds the EIN field, fills everything

Requirements:
  pip install pyautogui pyperclip pdfplumber watchdog pytesseract pillow psutil pygetwindow
  Also install Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
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
    Extract W-2 fields from a PDF.
    Tries pdfplumber text extraction first (works on digital PDFs).
    Falls back to pytesseract OCR for scanned images.
    """
    text = ""

    # Attempt 1: digital text extraction
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
    except Exception as e:
        print(f"  pdfplumber error: {e}")

    # Attempt 2: OCR fallback
    if len(text.strip()) < 50:
        try:
            import pytesseract
            from pdf2image import convert_from_path
            print("  Falling back to OCR...")
            images = convert_from_path(pdf_path, dpi=300)
            for img in images:
                text += pytesseract.image_to_string(img) + "\n"
        except Exception as e:
            print(f"  OCR error: {e}")

    if not text.strip():
        raise ValueError("Could not extract text from PDF")

    # Log full raw text so we can debug extraction patterns
    print(f"\n{'='*60}")
    print("RAW EXTRACTED TEXT:")
    print('='*60)
    print(text[:3000])
    print('='*60 + "\n")

    data = {}

    def find(pattern, flags=re.IGNORECASE):
        m = re.search(pattern, text, flags)
        return m.group(1).strip() if m else ""

    # Employer EIN
    data['ein'] = find(r'(?:b\.?\s+)?Employer(?:\'s)?\s+(?:identification|ID|id)\s+(?:number|no\.?)[:\s]+([\d]{2}-[\d]{7})')
    if not data['ein']:
        data['ein'] = find(r'EIN[:\s]+([\d]{2}-[\d]{7})')

    # Employer Name
    data['employer_name'] = find(r'(?:c\.?\s+)?Employer(?:\'s)?\s+name[,\s]+address.*?[\r\n]+([^\r\n]+)')

    # Employer Street/City/State/ZIP — usually line after name
    addr_match = re.search(
        r'(?:c\.?\s+)?Employer(?:\'s)?\s+name.*?[\r\n]+[^\r\n]+[\r\n]+([^\r\n]+)[\r\n]+([^\r\n,]+)[,\s]+([A-Z]{2})\s+([\d]{5}(?:-[\d]{4})?)',
        text, re.IGNORECASE | re.DOTALL
    )
    if addr_match:
        data['employer_street'] = addr_match.group(1).strip()
        data['employer_city']   = addr_match.group(2).strip()
        data['employer_state']  = addr_match.group(3).strip()
        data['employer_zip']    = addr_match.group(4).strip()
    else:
        data['employer_street'] = ''
        data['employer_city']   = ''
        data['employer_state']  = ''
        data['employer_zip']    = ''

    # Wage boxes
    def find_box(box_num, label_pattern):
        patterns = [
            rf'(?:^|\s){box_num}\s+{label_pattern}[:\s]+\$?([\d,]+(?:\.\d{{2}})?)',
            rf'{label_pattern}[:\s\n]+\$?([\d,]+(?:\.\d{{2}})?)',
        ]
        for p in patterns:
            val = find(p)
            if val:
                return clean_amount(val)
        return ''

    data['box1_wages']         = find_box('1', r'Wages,?\s+tips')
    data['box2_fed_tax']       = find_box('2', r'Federal\s+income\s+tax\s+withheld')
    data['box3_ss_wages']      = find_box('3', r'Social\s+security\s+wages')
    data['box4_ss_tax']        = find_box('4', r'Social\s+security\s+tax\s+withheld')
    data['box5_med_wages']     = find_box('5', r'Medicare\s+wages')
    data['box6_med_tax']       = find_box('6', r'Medicare\s+tax\s+withheld')
    data['box7_ss_tips']       = find_box('7', r'Social\s+security\s+tips')
    data['box8_alloc_tips']    = find_box('8', r'Allocated\s+tips')
    data['box10_dep_care']     = find_box('10', r'Dependent\s+care')
    data['box11_nonqual']      = find_box('11', r'Nonqualified\s+plans?')

    # Box 12 codes (up to 4 entries like "12a D 5000.00")
    data['box12'] = []
    for m in re.finditer(r'12[a-d]?\s+([A-Z]{1,2})\s+\$?([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE):
        data['box12'].append({'code': m.group(1).upper(), 'amount': clean_amount(m.group(2))})

    # Box 13 checkboxes
    data['box13_statutory']   = bool(re.search(r'statutory\s+employee.*?[Xx✓]|[Xx✓].*?statutory\s+employee', text, re.IGNORECASE))
    data['box13_retirement']  = bool(re.search(r'retirement\s+plan.*?[Xx✓]|[Xx✓].*?retirement\s+plan', text, re.IGNORECASE))
    data['box13_sick_pay']    = bool(re.search(r'third.?party\s+sick.*?[Xx✓]|[Xx✓].*?sick\s+pay', text, re.IGNORECASE))

    # Box 14 other
    data['box14'] = []
    for m in re.finditer(r'14\s+Other[:\s]+([^\n]+)\s+([\d,]+(?:\.\d{2})?)', text, re.IGNORECASE):
        data['box14'].append({'label': m.group(1).strip(), 'amount': clean_amount(m.group(2))})

    # State info (box 15-20)
    state_m = re.search(
        r'15\s+([A-Z]{2})\s+([\w-]+)\s+16\s+([\d,]+(?:\.\d{2})?)\s+17\s+([\d,]+(?:\.\d{2})?)',
        text, re.IGNORECASE
    )
    if state_m:
        data['box15_state']    = state_m.group(1)
        data['box15_state_id'] = state_m.group(2)
        data['box16_wages']    = clean_amount(state_m.group(3))
        data['box17_tax']      = clean_amount(state_m.group(4))
    else:
        data['box15_state']    = data.get('employer_state', '')
        data['box15_state_id'] = ''
        data['box16_wages']    = data.get('box1_wages', '')
        data['box17_tax']      = ''

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
    type_field(data.get('box8_alloc_tips', ''))

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
        tk.Label(root, text="Drop a W-2 PDF → auto-fills Drake",
                 font=("Segoe UI", 10), fg="#888", bg="#1a1a1a").pack()

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
                    f"EIN:          {data.get('ein') or '—'}\n"
                    f"Employer:     {data.get('employer_name') or '—'}\n"
                    f"Box 1 Wages:  {data.get('box1_wages') or '—'}\n"
                    f"Box 2 Fed Tax:{data.get('box2_fed_tax') or '—'}\n"
                    f"Box 3 SS Wg:  {data.get('box3_ss_wages') or '—'}\n"
                    f"Box 4 SS Tax: {data.get('box4_ss_tax') or '—'}\n"
                    f"Box 5 Med Wg: {data.get('box5_med_wages') or '—'}\n"
                    f"Box 6 Med Tx: {data.get('box6_med_tax') or '—'}\n"
                    f"State:        {data.get('box15_state') or '—'}\n\n"
                    f"─────────────────────────────\n"
                    f"Click YES, then you have 5 seconds to:\n"
                    f"  1. Switch to Drake\n"
                    f"  2. Click the EIN input field\n"
                    f"Then hold still — script fills everything.\n"
                    f"─────────────────────────────\n\n"
                    f"Click YES to fill, NO to skip."
                )
                confirmed[0] = messagebox.askyesno("Confirm Fill", summary)
                dialog_done.set()

            dialog_done = threading.Event()
            self.root.after(0, ask)
            dialog_done.wait(timeout=120)  # wait up to 2 min for user to click YES/NO

            if confirmed[0]:
                self.log_msg("⏳ 5 seconds — click the EIN field in Drake NOW!")
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
