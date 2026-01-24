# PyCombiner

[English version](README.en.md)

PyCombiner — лёгкое Windows‑приложение для запуска и мониторинга ваших проектов/скриптов (.py/.ps1/.bat/.exe) из одного окна. Оно ориентировано на повседневное удобство: список проектов с переключателем, запуск включённых одной кнопкой, вкладки логов и корректное завершение процессов.

Интерфейс сделан на PySide6 и поддерживает светлую/тёмную тему, Mica на Windows 11 и цветовые статусы. Логи отображаются во вкладках и ограничиваются по длине, чтобы не разрастались бесконтрольно.

![Скриншот приложения](https://github.com/user-attachments/assets/a5c797f9-1254-4e4d-bbcc-b02b24fb2622)

### Возможности
- Старт/стоп проектов и запуск включённых одной кнопкой.
- Цветовые статусы (running/starting/stopped/crashed).
- Вкладки логов с кнопкой «Очистить лог».
- Переключение языка интерфейса (RU/EN).
- Автозапуск при входе в Windows (HKCU\Run).
- Автозапуск при старте Windows без входа (Планировщик задач).
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

### Установка (если исходники уже есть)
```bash
python -m venv .venv
.venv\Scripts\activate
pip install PySide6
```

### Настройка
Локальный конфиг в репозитории не требуется. Настройки сохраняются здесь:
```
%APPDATA%\PyCombiner\config.json
```
Логи сохраняются здесь:
```
%APPDATA%\PyCombiner\logs
```

### Запуск
```bash
python pycombiner.py
```

### Автозапуск без входа (Планировщик задач)
1) В меню **Настройки** включите опцию **«Автозапуск PyCombiner при старте Windows (без входа)»**.  
2) Приложение попросит пароль Windows (он **не сохраняется**).  
3) Будет создана задача, которая запускает PyCombiner в headless‑режиме (`--headless --autostart`).

### Как работает GUI, когда запущен headless
- Headless управляет процессами и пишет состояние в:
  - `%APPDATA%\PyCombiner\state.json`
  - `%APPDATA%\PyCombiner\logs`
- При запуске GUI он подключается как клиент: читает `state.json`, отображает логи и отправляет команды через `%APPDATA%\PyCombiner\commands`.
- Это значит, что вы можете войти в систему, открыть GUI и сразу видеть всё, что происходило до входа, без остановки ботов.
- Если headless не запущен, GUI работает автономно (как обычный локальный режим).

### Сборка EXE
```bash
pyinstaller --noconfirm --clean --onefile --windowed --name PyCombiner ^
  --collect-all PySide6 --hidden-import shiboken6 ^
  --icon pycombiner.ico pycombiner.py
```

### Типовые ошибки и решения
- **PyInstaller ругается на отсутствующие DLL для SQL**  
  Это ожидаемо, если вы не используете Qt SQL. Предупреждения можно игнорировать.
- **У EXE нет иконки**  
  Убедитесь, что `pycombiner.ico` лежит в корне проекта. Windows может кэшировать иконки — перезапуск Explorer помогает.

### Безопасность
Не коммитьте секреты и личные настройки. Локальные артефакты (build/dist, .env, кэши и т.п.) исключены через `.gitignore`.

### Структура проекта
- `pycombiner.py` — основной файл приложения (UI + логика процессов).
- `pycombiner.ico` — иконка приложения/EXE.
- `PyCombiner.spec` — опциональный spec для PyInstaller.

### Лицензия и авторство
Проект распространяется по лицензии MIT — можно использовать в личных и коммерческих целях при сохранении текста лицензии и копирайта. См. файл LICENSE.  
Разработчик: Omnividente (Telegram: @Omnividente).
