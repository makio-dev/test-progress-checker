@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo Building exe...
pyinstaller --onefile --windowed --name "ITB_Checker" checker.py

echo Done! Output: dist\ITB_Checker.exe
pause
