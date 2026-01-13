# PyCombiner

[English version](README.en.md)

PyCombiner - легкое Windows GUI приложение для запуска и мониторинга ваших проектов/скриптов (.py/.ps1/.bat/.exe) из одного окна. Он ориентирован на повседневное удобство: список проектов с переключателем включения, запуск включенных одной кнопкой, вкладки логов и корректное завершение процессов.

Интерфейс сделан на PySide6 и поддерживает светлую/темную тему, Mica на Windows 11 и цветовые статусы. Логи отображаются в отдельных вкладках, а их длина ограничена, чтобы не разрастались бесконтрольно.

![Скриншот приложения](https://github.com/user-attachments/assets/a5c797f9-1254-4e4d-bbcc-b02b24fb2622)

### Возможности
- Старт/стоп проектов и запуск включенных одной кнопкой.
- Цветовые статусы (running/starting/stopped/crashed).
- Вкладки логов с кнопкой "Очистить лог".
- Переключение языка интерфейса (RU/EN).
- Автозапуск при входе в Windows (HKCU\Run).
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
  Убедитесь, что `pycombiner.ico` лежит в корне проекта. Windows может кэшировать иконки - перезапуск Explorer помогает.

### Безопасность
Не коммитьте секреты и личные настройки. Локальные артефакты (build/dist, .env, кэши и т.п.) исключены через `.gitignore`.

### Структура проекта
- `pycombiner.py` - основной файл приложения (UI + логика процессов).
- `pycombiner.ico` - иконка приложения/EXE.
- `PyCombiner.spec` - опциональный spec для PyInstaller.

### Лицензия и авторство
Проект распространяется по лицензии MIT - можно использовать в личных и коммерческих целях при сохранении текста лицензии и копирайта. См. файл LICENSE.  
Разработчик: Omnividente (Telegram: @Omnividente).
