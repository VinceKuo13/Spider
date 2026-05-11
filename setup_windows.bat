@echo off
echo [1/2] Installing Python dependencies...
pip install -r requirements.txt
echo [2/2] Installing Playwright Chromium for xinjiang.py...
python -m playwright install chromium
echo Done. You can run: python tianjin.py
pause
