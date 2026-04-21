@echo off
echo CZN Fragment Overlay v4
echo ========================
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator rights...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
echo Admin: OK
echo.

if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo ERROR: Tesseract not found!
    echo.
    echo Please download and install Tesseract:
    echo https://github.com/UB-Mannheim/tesseract/wiki
    echo.
    echo Direct download link:
    echo https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.5.0.20241111.exe
    echo.
    pause
    exit /b 1
)
echo Tesseract: OK
echo.

pip install pyautogui Pillow pytesseract -q
echo.
echo Starting overlay...
echo F9 = scan    Right-click = quit
echo.
python czn_overlay.py
echo.
echo --- ended ---
pause
