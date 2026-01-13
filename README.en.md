# PyCombiner

[??????? ??????](README.md)

PyCombiner is a lightweight Windows GUI to launch and monitor your projects/scripts (.py/.ps1/.bat/.exe) from a single window. It focuses on day?to?day convenience: a project list with enable toggles, batch start of enabled items, log tabs, and proper process termination.

The UI is built with PySide6 and supports light/dark themes, Mica on Windows 11, and color?coded status labels. Logs are shown in separate tabs with a line limit to avoid unbounded growth.

![App screenshot](https://github.com/user-attachments/assets/9548f2d3-89b2-4b41-bfdc-ab20d5d6d390)

### Features
- Start/stop projects and batch?start enabled ones.
- Color?coded statuses (running/starting/stopped/crashed).
- Log tabs with ?Clear log? button.
- UI language switch (RU/EN).
- App autostart on Windows login (HKCU\Run).
- Proper process?tree termination.

### Requirements
- Windows 10/11
- Python 3.10+ (for running from source)

### Install from GitHub
```bash
git clone https://github.com/Omnividente/PyCombiner.git
cd PyCombiner
python -m venv .venv
.venv\Scriptsctivate
pip install PySide6
```

### Installation (if you already have the sources)
```bash
python -m venv .venv
.venv\Scriptsctivate
pip install PySide6
```

### Configuration
No repo?local config is required. The app stores its settings here:
```
%APPDATA%\PyCombiner\config.json
```
Logs are stored in:
```
%APPDATA%\PyCombiner\logs
```

### Run
```bash
python pycombiner.py
```

### Build EXE
```bash
pyinstaller --noconfirm --clean --onefile --windowed --name PyCombiner ^
  --collect-all PySide6 --hidden-import shiboken6 ^
  --icon pycombiner.ico pycombiner.py
```

### Common issues
- **PyInstaller warns about missing SQL driver DLLs**  
  This is expected if you don?t use Qt SQL. The warnings can be ignored.
- **No icon after build**  
  Ensure `pycombiner.ico` is in the project root. Windows may cache icons; restarting Explorer can help.

### Security
Do not commit secrets or personal settings. Local artifacts (build/dist, .env, caches, etc.) are excluded via `.gitignore`.

### Project structure
- `pycombiner.py` ? main app file (UI + process logic).
- `pycombiner.ico` ? app/EXE icon.
- `PyCombiner.spec` ? optional PyInstaller spec.

### License and authorship
This project is released under the MIT License ? you may use it for personal and commercial purposes as long as the copyright notice and license text are kept. See the LICENSE file.  
Developer: Omnividente (Telegram: @Omnividente).
