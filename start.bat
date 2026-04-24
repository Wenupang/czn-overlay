@echo off
cd /d "%~dp0"
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs -WorkingDirectory '%~dp0'"
    exit /b
)
echo CZN Fragment Rater v5
echo =====================
pip install pyautogui -q
echo Starting...
python "%~dp0czn_overlay.py"
pause
