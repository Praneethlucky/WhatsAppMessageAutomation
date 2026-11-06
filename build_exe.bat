@echo off
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name WhatsAppBulkSender main.py
echo.
echo EXE created at dist\WhatsAppBulkSender.exe
pause
