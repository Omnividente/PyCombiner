# PyCombiner

## Русский

PyCombiner — это компактный Windows‑GUI для запуска и мониторинга ваших проектов/скриптов (.py/.ps1/.bat/.exe) из одного окна. Приложение ориентировано на быстрый повседневный запуск: список проектов, переключатель «вкл.», автозапуск выбранных, вкладки логов и корректное завершение процессов.

Интерфейс построен на PySide6 и поддерживает светлую/тёмную тему, Mica на Windows 11 и цветовую индикацию статуса. Логи отображаются в отдельных вкладках, а объём логов ограничен, чтобы UI не разрастался со временем.

![Скриншот приложения](assets/screenshot.png)

### Возможности
- Запуск/остановка проектов и «пакетный» старт включённых.
- Цветовая индикация статусов (running/starting/stopped/crashed).
- Вкладки логов с кнопкой «Очистить лог».
- Переключение языка интерфейса (RU/EN).
- Автозапуск приложения при входе в Windows (HKCU\Run).
- Корректное завершение дерева процессов.

### Требования
- Windows 10/11
- Python 3.10+ (для запуска из исходников)

### Установка из GitHub
```bash
git clone https://github.com/Omnividente/PyCombiner.git
cd PyCombiner
python -m venv .venv
.venv\Scripts\activate
pip install PySide6
```

### Установка (если исходники уже скачаны)
```bash
python -m venv .venv
.venv\Scripts\activate
pip install PySide6
```

### Настройка
Отдельные конфиги в репозитории не требуются. Приложение само создаёт и хранит настройки в:
```
%APPDATA%\PyCombiner\config.json
```
Логи сохраняются в:
```
%APPDATA%\PyCombiner\logs
```

### Запуск
```bash
python pycombiner.py
```

### Сборка EXE
```bash
pyinstaller --noconfirm --clean --onefile --windowed --name PyCombiner ^
  --collect-all PySide6 --hidden-import shiboken6 ^
  --icon pycombiner.ico pycombiner.py
```

### Типовые ошибки и решения
- **PyInstaller ругается на отсутствующие DLL для SQL‑драйверов**  
  Это нормально, если вы не используете Qt SQL. Предупреждения можно игнорировать.
- **После сборки не видно иконку**  
  Проверьте, что `pycombiner.ico` лежит в корне проекта. Windows может кэшировать иконки — иногда помогает перезапуск проводника.

### Безопасность
Не коммитьте личные настройки и секреты. Локальные файлы (build/dist, .env, кэш и т.д.) исключены через `.gitignore`.

### Структура проекта
- `pycombiner.py` — основной файл приложения (GUI, логика запуска).
- `pycombiner.ico` — иконка приложения/EXE.
- `PyCombiner.spec` — опциональный spec для PyInstaller.

### Лицензия и авторство
Проект распространяется под лицензией MIT — можно свободно использовать в личных и коммерческих целях при сохранении уведомления об авторских правах и текста лицензии. См. файл LICENSE.  
Разработчик: Omnividente (Telegram: @Omnividente).

---

## English

PyCombiner is a lightweight Windows GUI to launch and monitor your projects/scripts (.py/.ps1/.bat/.exe) from a single window. It focuses on day‑to‑day convenience: a project list with enable toggles, batch start of enabled items, log tabs, and proper process termination.

The UI is built with PySide6 and supports light/dark themes, Mica on Windows 11, and color‑coded status labels. Logs are shown in separate tabs with a line limit to avoid unbounded growth.

![App screenshot](assets/screenshot.png)

### Features
- Start/stop projects and batch‑start enabled ones.
- Color‑coded statuses (running/starting/stopped/crashed).
- Log tabs with “Clear log” button.
- UI language switch (RU/EN).
- App autostart on Windows login (HKCU\Run).
- Proper process‑tree termination.

### Requirements
- Windows 10/11
- Python 3.10+ (for running from source)

### Install from GitHub
```bash
git clone https://github.com/Omnividente/PyCombiner.git
cd PyCombiner
python -m venv .venv
.venv\Scripts\activate
pip install PySide6
```

### Installation (if you already have the sources)
```bash
python -m venv .venv
.venv\Scripts\activate
pip install PySide6
```

### Configuration
No repo‑local config is required. The app stores its settings here:
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
  This is expected if you don’t use Qt SQL. The warnings can be ignored.
- **No icon after build**  
  Ensure `pycombiner.ico` is in the project root. Windows may cache icons; restarting Explorer can help.

### Security
Do not commit secrets or personal settings. Local artifacts (build/dist, .env, caches, etc.) are excluded via `.gitignore`.

### Project structure
- `pycombiner.py` — main app file (UI + process logic).
- `pycombiner.ico` — app/EXE icon.
- `PyCombiner.spec` — optional PyInstaller spec.

### License and authorship
This project is released under the MIT License — you may use it for personal and commercial purposes as long as the copyright notice and license text are kept. See the LICENSE file.  
Developer: Omnividente (Telegram: @Omnividente).
