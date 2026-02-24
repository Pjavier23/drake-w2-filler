@echo off
title Drake W-2 Filler - Installer
color 0B
echo.
echo  ============================================
echo   Drake W-2 Auto-Filler - Setup
echo  ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found!
    echo  Download from: https://www.python.org/downloads/
    echo  Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  [1/3] Python found. Installing packages...
python -m pip install --upgrade pip
python -m pip install pyautogui pyperclip pdfplumber watchdog pytesseract pdf2image Pillow psutil pygetwindow

echo.
echo  [2/3] Checking Tesseract OCR...
tesseract --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [WARNING] Tesseract not installed.
    echo  Only needed for scanned/image PDFs.
    echo  Download: https://github.com/UB-Mannheim/tesseract/wiki
    echo  Install to default location: C:\Program Files\Tesseract-OCR\
    echo.
) else (
    echo  Tesseract found.
)

echo  [3/3] Creating folders...
mkdir "C:\W2_Inbox" 2>nul
mkdir "C:\W2_Done"  2>nul
mkdir "C:\W2_Errors" 2>nul
echo  Folders created: C:\W2_Inbox, C:\W2_Done, C:\W2_Errors

echo.
echo  ============================================
echo   Installation complete!
echo   Run: python drake_w2_filler.py
echo   Or double-click: run.bat
echo  ============================================
echo.
pause
