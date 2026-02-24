# Drake W-2 Auto-Filler

Automates W-2 data entry in Drake Tax 2025.  
Drop a W-2 PDF → it reads the fields → fills Drake's screen automatically.

---

## Quick Start

### 1. Install (one time only)

Double-click `install.bat`

If you don't have Python:
- Download from https://www.python.org/downloads/
- ✅ Check **"Add Python to PATH"** during install

### 2. Run

Double-click `run.bat`

### 3. Use it

1. Open Drake Tax
2. Open the client's return
3. Navigate to the **W-2 entry screen** (where you see the EIN field)
4. In the W-2 Filler app, click **▶ Start Watching**
5. Drop a W-2 PDF into `C:\W2_Inbox\`
6. A popup shows the extracted data — review it
7. Click **YES**, then **immediately click the EIN field in Drake**
8. The script fills everything in ~10 seconds

Done PDFs go to `C:\W2_Done\`  
Failed PDFs go to `C:\W2_Errors\`

---

## What Gets Filled

| Drake Field | Source |
|-------------|--------|
| EIN | Box b |
| Employer Name | Box c |
| Employer Street/City/State/ZIP | Box c |
| Box 1 — Wages | Box 1 |
| Box 2 — Federal Tax W/H | Box 2 |
| Box 3 — SS Wages | Box 3 |
| Box 4 — SS Tax W/H | Box 4 |
| Box 5 — Medicare Wages | Box 5 |
| Box 6 — Medicare Tax W/H | Box 6 |
| Box 12 codes/amounts | Box 12a-d |
| Box 13 checkboxes | Box 13 |
| Box 14 other | Box 14 |
| Box 15-17 State/Tax | Box 15-17 |

---

## Tips

- **Slow computer?** Open `drake_w2_filler.py` and increase `KEYSTROKE_DELAY` from 0.25 to 0.4
- **Emergency stop:** Slam mouse to the **top-left corner** of screen — PyAutoGUI will abort
- **Test without filling:** Click **🧪 Test Extract** to see what the script reads from any PDF without touching Drake
- **Multiple W-2s:** Drop them one at a time. Wait for the popup before dropping the next

---

## Troubleshooting

**"Could not extract text from PDF"**
- Install Tesseract OCR for scanned PDFs: https://github.com/UB-Mannheim/tesseract/wiki
- Install to default path: `C:\Program Files\Tesseract-OCR\`

**Fields filled in wrong boxes**
- The Tab order assumes you're starting with cursor in the EIN field
- Click once in the EIN field right after clicking YES in the popup

**Script types too fast / Drake misses fields**
- Increase `KEYSTROKE_DELAY` in the script (line ~50)

**Numbers are missing**
- Use "🧪 Test Extract" first to verify the PDF data is readable
- Some W-2 PDFs are scanned images — need Tesseract installed

---

## Files

```
drake-w2-filler/
├── drake_w2_filler.py   ← main script
├── install.bat          ← run once to install dependencies
├── run.bat              ← double-click to launch
├── requirements.txt     ← Python packages
└── README.md            ← this file
```
