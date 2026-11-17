@echo off
REM Rebuilds ResumeTailorAI.exe (needs: pip install -r app\requirements.txt)
cd /d "%~dp0app"
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name ResumeTailorAI --icon "%~dp0app\icon.ico" ^
  --add-data "%~dp0app\ui;ui" ^
  --collect-submodules webview ^
  --exclude-module tkinter --exclude-module matplotlib --exclude-module numpy ^
  --distpath "%~dp0." --workpath "%~dp0build" --specpath "%~dp0build" ^
  main.py
