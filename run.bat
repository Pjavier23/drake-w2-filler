@echo off
title Drake W-2 Auto-Filler
cd /d "%~dp0"
python -m pip install pyautogui pyperclip pdfplumber watchdog Pillow psutil pygetwindow -q
python drake_w2_filler.py
pause
