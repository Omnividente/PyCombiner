# -*- coding: utf-8 -*-
"""
PyCombiner — GUI-комбайн для запуска ваших проектов (Windows 10/11, PySide6)

Фичи:
- Светлая/тёмная темы, Mica/тёмный титлбар (Win11), мягкий Win-стиль
- Список проектов с тумблером "Вкл." (persist в config.json)
- Последовательный запуск "включённых"
- Корректный стоп: taskkill /T /F (убивает дерево)
- Санитарная очистка зомби перед стартом проекта
- Автозапуск PyCombiner при входе в Windows (HKCU\Run)
- Вкладки логов для каждого проекта
"""

from __future__ import annotations
from PySide6.QtCore import Qt
from PySide6 import QtCore, QtGui, QtWidgets

import argparse
import faulthandler
import json
import locale
import os
import platform
import re
import shlex
import subprocess
import sys
import time
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Dict, Tuple
ANSI_RE = re.compile(r"\x1B\[[0-?]*[ -/]*[@-~]")


def _strip_ansi(text: str) -> str:
    try:
        return ANSI_RE.sub('', text)
    except Exception:
        return text


# ------------------------ Константы/пути -----------------------------------
APP_NAME = "PyCombiner"
APP_VERSION = "1.1.3"
ORG_NAME = "PyCombiner"
AUTOSTART_TASK_NAME = "PyCombiner Startup"

APPDATA_DIR = Path(os.environ.get("APPDATA", str(
    Path.home() / "AppData" / "Roaming"))) / "PyCombiner"
CONFIG_PATH = APPDATA_DIR / "config.json"
LOGS_DIR = APPDATA_DIR / "logs"
STATE_PATH = APPDATA_DIR / "state.json"
COMMANDS_DIR = APPDATA_DIR / "commands"
DAEMON_PID_PATH = APPDATA_DIR / "daemon.pid"
APP_LOG_PATH = APPDATA_DIR / "app.log"
APPDATA_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)
COMMANDS_DIR.mkdir(parents=True, exist_ok=True)
LOG_MAX_LINES = 300
LOG_ROTATE_MAX_BYTES = 5 * 1024 * 1024
LOG_ROTATE_COUNT = 3
PID_CACHE_TTL_SEC = 2.0
_PID_CACHE: Dict[int, Tuple[float, bool]] = {}
PENDING_STATUS_GRACE_SEC = 3.0
NETWORK_CHECK_INTERVAL_SEC = 10
NETWORK_CHECK_TIMEOUT_SEC = 3


# ------------------------ Утилиты ------------------------------------------

def ensure_dirs():
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    COMMANDS_DIR.mkdir(parents=True, exist_ok=True)


def log_app(message: str) -> None:
    try:
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        APP_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with APP_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"[{ts}] {message}\n")
    except Exception:
        pass


def install_debug_handlers() -> None:
    try:
        fh = APPDATA_DIR / "app_crash.log"
        fh.parent.mkdir(parents=True, exist_ok=True)
        with fh.open("a", encoding="utf-8") as f:
            faulthandler.enable(file=f, all_threads=True)
    except Exception:
        pass

    def _excepthook(exc_type, exc, tb):
        try:
            log_app("Unhandled exception:\n" +
                    "".join(traceback.format_exception(exc_type, exc, tb)))
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = _excepthook

    def _qt_msg_handler(mode, context, message):
        try:
            log_app(f"Qt[{mode}] {message}")
        except Exception:
            pass

    try:
        QtCore.qInstallMessageHandler(_qt_msg_handler)
    except Exception:
        pass


def set_data_dir(path: str) -> None:
    global APPDATA_DIR, CONFIG_PATH, LOGS_DIR, STATE_PATH, COMMANDS_DIR, DAEMON_PID_PATH, APP_LOG_PATH
    if not path:
        return
    p = Path(path).expanduser()
    APPDATA_DIR = p
    CONFIG_PATH = APPDATA_DIR / "config.json"
    LOGS_DIR = APPDATA_DIR / "logs"
    STATE_PATH = APPDATA_DIR / "state.json"
    COMMANDS_DIR = APPDATA_DIR / "commands"
    DAEMON_PID_PATH = APPDATA_DIR / "daemon.pid"
    APP_LOG_PATH = APPDATA_DIR / "app.log"
    ensure_dirs()


def _env_lookup(env: Optional[Dict[str, str]], key: str, default: str = "") -> str:
    if env is None:
        return os.environ.get(key, default)
    if key in env and env[key] is not None:
        return str(env[key])
    key_lower = key.lower()
    for k, v in env.items():
        if k.lower() == key_lower and v is not None:
            return str(v)
    return default


def _normalize_env_snapshot(raw: object) -> Dict[str, str]:
    if not isinstance(raw, dict):
        return {}
    out: Dict[str, str] = {}
    for k, v in raw.items():
        if not isinstance(k, str):
            continue
        if v is None:
            continue
        out[k] = str(v)
    return out


def capture_env_snapshot() -> Dict[str, str]:
    return {str(k): str(v) for k, v in os.environ.items()}


def maybe_update_env_snapshot(cfg: "Config") -> None:
    current = _normalize_env_snapshot(cfg.data.get("env_snapshot"))
    snapshot = capture_env_snapshot()
    if snapshot != current:
        cfg.data["env_snapshot"] = snapshot
        cfg.save()


def build_process_environment(snapshot: Optional[Dict[str, str]]) -> QtCore.QProcessEnvironment:
    if snapshot:
        env = QtCore.QProcessEnvironment()
        for k, v in snapshot.items():
            env.insert(k, v)
        return env
    return QtCore.QProcessEnvironment.systemEnvironment()


def get_win_build() -> int:
    try:
        return sys.getwindowsversion().build  # type: ignore[attr-defined]
    except Exception:
        return 19041


def get_system_is_light() -> bool:
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        ) as k:
            v, _ = winreg.QueryValueEx(k, "AppsUseLightTheme")
            return bool(v)
    except Exception:
        return True


def get_accent_color_argb(default: int = 0xFF2196F3) -> int:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\DWM") as k:
            v, _ = winreg.QueryValueEx(k, "ColorizationColor")
            return int(v)
    except Exception:
        return default


def argb_to_qcolor(argb: int) -> QtGui.QColor:
    a = (argb >> 24) & 0xFF
    r = (argb >> 16) & 0xFF
    g = (argb >> 8) & 0xFF
    b = (argb >> 0) & 0xFF
    c = QtGui.QColor(r, g, b)
    c.setAlpha(a)
    return c


def decode_bytes(data: bytes) -> str:
    if not data:
        return ""
    try:
        return data.decode("utf-8", errors="replace")
    except Exception:
        try:
            return data.decode(locale.getpreferredencoding(False) or "cp1251", errors="replace")
        except Exception:
            return data.decode("cp1251", errors="replace")


def is_network_ready() -> bool:
    if os.name != "nt":
        return True
    try:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        ps = (
            "Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue "
            "| Where-Object { $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '127.0.0.1' } "
            "| Select-Object -First 1 -ExpandProperty IPAddress"
        )
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps],
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=flags,
            timeout=NETWORK_CHECK_TIMEOUT_SEC,
        )
        return bool(out.strip())
    except Exception:
        # Fail-open to avoid blocking launches on unexpected errors.
        return True


# ------------------------ DWM / заголовок / Mica ---------------------------

def _dwm_set_attribute(hwnd: int, attr: int, value: int) -> None:
    try:
        import ctypes
        dwmapi = ctypes.windll.dwmapi
        val = ctypes.c_int(value)
        dwmapi.DwmSetWindowAttribute(ctypes.c_void_p(hwnd), ctypes.c_int(attr),
                                     ctypes.byref(val), ctypes.sizeof(val))
    except Exception:
        pass


def _set_immersive_dark(hwnd: int, enabled: bool) -> None:
    # 20 (Win11), 19 (Win10)
    for a in (20, 19):
        _dwm_set_attribute(hwnd, a, 1 if enabled else 0)


def enable_mica_and_titlebar(win: QtWidgets.QWidget, *, mica_light: bool, dark_title: bool) -> None:
    hwnd = int(win.winId())
    build = get_win_build()
    _set_immersive_dark(hwnd, dark_title)
    if build >= 22000:
        # 38 = DWMWA_SYSTEMBACKDROP_TYPE, 2 = Mica
        _dwm_set_attribute(hwnd, 38, 2 if mica_light else 0)


# ------------------------ Иконка -------------------------------------------

def build_fallback_icon(size: int = 256, accent: Optional[QtGui.QColor] = None) -> QtGui.QIcon:
    if accent is None:
        accent = argb_to_qcolor(get_accent_color_argb())
    pm = QtGui.QPixmap(size, size)
    pm.fill(QtCore.Qt.transparent)
    p = QtGui.QPainter(pm)
    p.setRenderHint(QtGui.QPainter.Antialiasing, True)
    rect = QtCore.QRectF(0, 0, size, size)
    path = QtGui.QPainterPath()
    r = size * 0.22
    path.addRoundedRect(rect.adjusted(size * 0.08, size *
                        0.08, -size * 0.08, -size * 0.08), r, r)
    p.setPen(QtCore.Qt.NoPen)
    p.setBrush(accent)
    p.drawPath(path)
    p.setPen(QtGui.QPen(QtGui.QColor("white")))
    font = QtGui.QFont("Segoe UI Variable", int(
        size * 0.38), QtGui.QFont.DemiBold)
    p.setFont(font)
    p.drawText(pm.rect(), Qt.AlignCenter, "Py")
    p.end()
    ic = QtGui.QIcon()
    ic.addPixmap(pm)
    return ic


def load_app_icon() -> QtGui.QIcon:
    for c in (
        Path(sys.argv[0]).with_name("pycombiner.ico"),
        Path.cwd() / "pycombiner.ico",
        APPDATA_DIR / "pycombiner.ico",
        Path(sys.argv[0]).with_name("app.ico"),
        Path.cwd() / "app.ico",
        APPDATA_DIR / "app.ico",
    ):
        if c.exists():
            try:
                return QtGui.QIcon(str(c))
            except Exception:
                pass
    return build_fallback_icon()


# ------------------------ Тема / QSS ---------------------------------------

class Theme:
    System = "system"
    Light = "light"
    Dark = "dark"


class Lang:
    System = "system"
    RU = "ru"
    EN = "en"


def get_system_lang() -> str:
    try:
        loc = locale.getdefaultlocale()[0] or ""
    except Exception:
        loc = ""
    return Lang.RU if loc.lower().startswith("ru") else Lang.EN


def resolve_lang(lang: str) -> str:
    if not lang or lang == Lang.System:
        return get_system_lang()
    return lang if lang in (Lang.RU, Lang.EN) else Lang.EN


STRINGS = {
    Lang.RU: {
        "menu_file": "Файл",
        "menu_settings": "Настройки",
        "menu_appearance": "Оформление",
        "menu_language": "Язык",
        "menu_help": "Справка",
        "act_exit": "Выход",
        "act_autostart": "Автозапуск PyCombiner при входе в Windows",
        "act_autostart_task": "Автозапуск PyCombiner при старте Windows (без входа)",
        "act_theme_system": "Системная тема",
        "act_theme_light": "Светлая",
        "act_theme_dark": "Тёмная",
        "act_use_mica": "Фон Mica (Windows 11)",
        "act_about": "О программе",
        "lang_system": "Системный",
        "lang_ru": "Русский",
        "lang_en": "English",
        "btn_add": "Добавить",
        "btn_edit": "Изменить",
        "btn_del": "Удалить",
        "btn_start": "Старт",
        "btn_stop": "Стоп",
        "btn_restart": "Перезапуск",
        "btn_start_enabled": "Старт (включённые)",
        "btn_stop_all": "Стоп все",
        "header_enabled": "Вкл.",
        "header_name": "Имя",
        "header_status": "Статус",
        "header_cmd": "Команда",
        "header_cwd": "Рабочая папка",
        "tab_clear": "Очистить лог",
        "tab_clear_tip": "Очистить лог текущей вкладки",
        "msg_stop_running": "Сначала остановите запущенный проект.",
        "msg_no_enabled": "Нет включённых проектов для запуска.",
        "msg_no_cmd": "Не указана команда запуска.",
        "log_start": "Запущен: {cmd}",
        "log_stop": "Остановка…",
        "log_finish": "Завершён (code={code}, status={status}).",
        "log_adopted": "Уже запущен (pid={pid}), подхвачен.",
        "log_start_error": "[!] Ошибка запуска: {err}",
        "log_proc_error": "[!] Ошибка процесса: {err}",
        "log_wait_net": "Ожидаю сеть...",
        "about_text": "{app}\nКомбайн процессов с Fluent-оформлением.\nКонфиг: {config}\nЛоги: {logs}",
        "tray_show": "Показать окно",
        "tray_start_enabled": "Старт включённых",
        "tray_exit": "Выход",
        "tray_minimized": "Свернуто в трей",
        "dlg_title": "Проект",
        "dlg_label_cmd": "Команда (.py|.ps1|.bat|.exe):",
        "dlg_label_cwd": "Рабочая папка:",
        "dlg_label_args": "Параметры запуска:",
        "dlg_label_name": "Имя:",
        "dlg_placeholder_args": "например: --env prod --threads 4",
        "dlg_browse": "Обзор",
        "dlg_pick_cmd": "Выберите файл",
        "dlg_pick_cwd": "Выберите папку",
        "dlg_filter_cmd": "Скрипты/исполняемые (*.ps1 *.bat *.cmd *.exe *.py);;Все файлы (*.*)",
        "dlg_chk_enabled": "Включать при старте",
        "dlg_chk_autorst": "Авто‑перезапуск при падении",
        "dlg_chk_clear_log": "Очищать лог при старте",
        "dlg_app_autostart_group": "Автозапуск приложения",
        "dlg_app_autostart_run": "Запускать PyCombiner при входе в Windows",
        "dlg_app_autostart_task": "Запускать PyCombiner при старте Windows (без входа)",
        "dlg_app_autostart_task_tip": "Потребуется пароль Windows. Пароль в приложении не хранится.",
        "dlg_default_name": "Проект",
        "msg_task_user_missing": "Не удалось определить пользователя Windows.",
        "msg_task_password_title": "Пароль Windows",
        "msg_task_password_label": "Введите пароль для пользователя {user}. Он нужен, чтобы запускать PyCombiner при старте системы без входа.",
        "msg_task_password_empty": "Пароль не введён — автозапуск при старте Windows не включён.",
        "msg_task_enable_failed": "Не удалось создать задачу автозапуска.\n{err}",
        "msg_task_disable_failed": "Не удалось удалить задачу автозапуска.\n{err}",
        "headless_connected": "Headless: подключен (pid={pid})",
        "headless_disconnected": "Headless: не подключен",
    },
    Lang.EN: {
        "menu_file": "File",
        "menu_settings": "Settings",
        "menu_appearance": "Appearance",
        "menu_language": "Language",
        "menu_help": "Help",
        "act_exit": "Exit",
        "act_autostart": "Start PyCombiner on Windows login",
        "act_autostart_task": "Start PyCombiner at Windows startup (no login)",
        "act_theme_system": "System theme",
        "act_theme_light": "Light",
        "act_theme_dark": "Dark",
        "act_use_mica": "Mica background (Windows 11)",
        "act_about": "About",
        "lang_system": "System",
        "lang_ru": "Russian",
        "lang_en": "English",
        "btn_add": "Add",
        "btn_edit": "Edit",
        "btn_del": "Delete",
        "btn_start": "Start",
        "btn_stop": "Stop",
        "btn_restart": "Restart",
        "btn_start_enabled": "Start (enabled)",
        "btn_stop_all": "Stop all",
        "header_enabled": "On",
        "header_name": "Name",
        "header_status": "Status",
        "header_cmd": "Command",
        "header_cwd": "Working dir",
        "tab_clear": "Clear log",
        "tab_clear_tip": "Clear current tab log",
        "msg_stop_running": "Stop the running project first.",
        "msg_no_enabled": "No enabled projects to start.",
        "msg_no_cmd": "Launch command is empty.",
        "log_start": "Started: {cmd}",
        "log_stop": "Stopping…",
        "log_finish": "Finished (code={code}, status={status}).",
        "log_adopted": "Already running (pid={pid}), adopted.",
        "log_start_error": "[!] Start error: {err}",
        "log_proc_error": "[!] Process error: {err}",
        "log_wait_net": "Waiting for network...",
        "about_text": "{app}\nProcess combiner with Fluent styling.\nConfig: {config}\nLogs: {logs}",
        "tray_show": "Show window",
        "tray_start_enabled": "Start enabled",
        "tray_exit": "Exit",
        "tray_minimized": "Minimized to tray",
        "dlg_title": "Project",
        "dlg_label_cmd": "Command (.py|.ps1|.bat|.exe):",
        "dlg_label_cwd": "Working directory:",
        "dlg_label_args": "Launch arguments:",
        "dlg_label_name": "Name:",
        "dlg_placeholder_args": "e.g. --env prod --threads 4",
        "dlg_browse": "Browse",
        "dlg_pick_cmd": "Select file",
        "dlg_pick_cwd": "Select folder",
        "dlg_filter_cmd": "Scripts/Executables (*.ps1 *.bat *.cmd *.exe *.py);;All files (*.*)",
        "dlg_chk_enabled": "Enable on startup",
        "dlg_chk_autorst": "Auto‑restart on crash",
        "dlg_chk_clear_log": "Clear log on start",
        "dlg_app_autostart_group": "App autostart",
        "dlg_app_autostart_run": "Start PyCombiner on Windows login",
        "dlg_app_autostart_task": "Start PyCombiner at Windows startup (no login)",
        "dlg_app_autostart_task_tip": "Windows password is required. The password is not stored in the app.",
        "dlg_default_name": "Project",
        "msg_task_user_missing": "Unable to determine the Windows user.",
        "msg_task_password_title": "Windows password",
        "msg_task_password_label": "Enter the password for user {user}. It is required to run PyCombiner at system startup without login.",
        "msg_task_password_empty": "Password not provided — startup task was not enabled.",
        "msg_task_enable_failed": "Failed to create the startup task.\n{err}",
        "msg_task_disable_failed": "Failed to remove the startup task.\n{err}",
        "headless_connected": "Headless: connected (pid={pid})",
        "headless_disconnected": "Headless: not connected",
    },
}

STATUS_LABELS = {
    Lang.RU: {
        "running": "работает",
        "starting": "запуск",
        "waiting": "ожидание сети",
        "stopping": "остановка",
        "stopped": "остановлен",
        "crashed": "ошибка",
    },
    Lang.EN: {
        "running": "running",
        "starting": "starting",
        "waiting": "waiting for network",
        "stopping": "stopping",
        "stopped": "stopped",
        "crashed": "crashed",
    },
}


def tr(lang: str, key: str, **kwargs) -> str:
    lang = resolve_lang(lang)
    text = STRINGS.get(lang, {}).get(key, STRINGS[Lang.EN].get(key, key))
    return text.format(**kwargs) if kwargs else text


def status_label(lang: str, status: str) -> str:
    lang = resolve_lang(lang)
    status = (status or "").strip().lower()
    return STATUS_LABELS.get(lang, STATUS_LABELS[Lang.EN]).get(status, status)


def theme_colors_hex(theme: str) -> dict:
    is_light = (theme == Theme.Light) or (
        theme == Theme.System and get_system_is_light())
    return {
        "bg":     "#f5f6f7" if is_light else "#202428",
        "card":   "#ffffff" if is_light else "#2a2e34",
        "text":   "#121212" if is_light else "#e6e6e6",
        "sub":    "#666a70" if is_light else "#b0b4bb",
        "border": "#e3e5e8" if is_light else "#3a3f45",
    }


def build_qss(theme: str, accent: QtGui.QColor) -> str:
    cols = theme_colors_hex(theme)
    acc = accent.name()
    return f"""
    QWidget {{ color: {cols['text']}; font-size: 13px; font-family: "Segoe UI Variable","Segoe UI"; }}
    QMainWindow {{ background: {cols['bg']}; }}

    QMenuBar {{ background: {cols['card']}; color: {cols['text']}; border-bottom: 1px solid {cols['border']}; }}
    QMenuBar::item {{ padding: 6px 10px; border-radius: 6px; }}
    QMenuBar::item:selected {{ background: {acc}33; }}

    QFrame#Card {{ background: {cols['card']}; border: 1px solid {cols['border']}; border-radius: 14px; }}

    QPushButton {{ background: transparent; border: 1px solid {cols['border']}; padding: 6px 12px; border-radius: 10px; }}
    QPushButton:hover {{ border-color: {acc}; }}
    QPushButton:pressed {{ background: {acc}2a; }}
    QToolButton {{ background: transparent; border: 1px solid {cols['border']}; padding: 4px 10px; border-radius: 8px; }}
    QToolButton:hover {{ border-color: {acc}; }}
    QToolButton:pressed {{ background: {acc}2a; }}

    QMenu {{ background: {cols['card']}; color: {cols['text']}; border: 1px solid {cols['border']}; border-radius: 10px; padding: 6px; }}
    QMenu::item {{ padding: 6px 10px; border-radius: 6px; }}
    QMenu::item:selected {{ background: {acc}33; }}
    QMenu::separator {{ height: 1px; background: {cols['border']}; margin: 6px; }}

    QDialog, QMessageBox, QFileDialog {{ background: {cols['card']}; color: {cols['text']}; }}
    QLabel, QCheckBox, QRadioButton {{ color: {cols['text']}; }}
    QCheckBox::indicator {{ width: 18px; height: 18px; border-radius: 5px; border: 1px solid {cols['border']}; background: {cols['card']}; }}
    QCheckBox::indicator:checked {{ background: {acc}; border-color: {acc}; }}

    QTreeWidget, QTreeView, QTableView {{ background: transparent; color: {cols['text']}; }}
    QHeaderView::section {{ background: {cols['card']}; padding: 6px; border: none; color: {cols['sub']}; border-bottom: 1px solid {cols['border']}; font-weight: 600; }}
    QTreeWidget::item {{ padding: 4px 6px; }}
    QTreeWidget::item:selected {{ background: {acc}22; }}
    QTreeView::indicator {{ width: 0px; height: 0px; }}
    QTreeWidget::branch {{ background: transparent; }}

    QLineEdit {{ border: 1px solid {cols['border']}; border-radius: 8px; padding: 6px 8px; background: {cols['card']}; color: {cols['text']}; }}
    QComboBox, QTextEdit, QPlainTextEdit {{ border: 1px solid {cols['border']}; border-radius: 8px; background: {cols['card']}; color: {cols['text']}; }}
    QTextEdit::viewport, QPlainTextEdit::viewport {{ background: {cols['card']}; color: {cols['text']}; }}
    QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {{ border: 1px solid {acc}; }}

    QTabWidget::pane {{
        border: 1px solid {cols['border']};
        border-radius: 12px;
        top: -1px;
        background: {cols['card']};
    }}
    QTabBar::tab {{
        background: transparent;
        color: {cols['sub']};
        padding: 6px 12px;
        margin: 0 6px 6px 0;
        border: 1px solid transparent;
        border-radius: 8px;
    }}
    QTabBar::tab:hover {{ background: {acc}18; }}
    QTabBar::tab:selected {{
        background: {acc}12;
        color: {cols['text']};
        border: 1px solid {acc};
        font-weight: 600;
    }}
    """


# ------------------------ Жёсткое убийство/очистка -------------------------

def _win_taskkill_tree(pid: int):
    """Жёстко прибить процесс и всех его детей (Windows)."""
    if not pid or pid <= 0:
        return
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        subprocess.run(
            ["taskkill", "/PID", str(int(pid)), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
            creationflags=flags,
        )
    except Exception:
        pass


def _win_find_project_pids(command: str, work_dir: str, exclude_pids: Optional[set[int]] = None) -> List[int]:
    """
    Ищем процессы проекта по подстрокам командной строки/рабочей папки.
    Возвращаем список PID (Windows, через PowerShell/CIM).
    """
    if platform.system() != "Windows":
        return []

    needles = []
    if command:
        needles.append(command)
        needles.append(os.path.basename(command))
    if work_dir:
        needles.append(work_dir)
    needles = [n for n in needles if n]

    if not needles:
        return []

    # В PowerShell оборачиваем шаблон в ОДИНАРНЫЕ кавычки: '*needle*'
    # Если внутри есть одинарная кавычка — удваиваем её.
    cond_parts = []
    for n in needles:
        esc = n.replace("'", "''")
        cond_parts.append(f"($_.CommandLine -like '*{esc}*')")
    cond = " -and ".join(cond_parts)
    # avoid matching the powershell process running this query
    cond = f"($_.ProcessId -ne $PID) -and ({cond})"

    ps = (
        f"Get-CimInstance Win32_Process | Where-Object {{ {cond} }} "
        f"| Select-Object -ExpandProperty ProcessId"
    )

    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps],
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=flags,
        )
    except Exception:
        out = ""

    pids: List[int] = []
    for line in (out or "").splitlines():
        line = line.strip()
        if not line or not line.isdigit():
            continue
        try:
            pid = int(line)
        except Exception:
            continue
        if exclude_pids and pid in exclude_pids:
            continue
        pids.append(pid)
    return pids


def _win_kill_project_zombies(command: str, work_dir: str, exclude_pids: Optional[set[int]] = None):
    """
    Перед стартом пробуем найти и погасить зависшие процессы проекта
    по подстрокам из командной строки/рабочей папки.
    Делается через PowerShell и запрос CIM Win32_Process.
    """
    for pid in _win_find_project_pids(command, work_dir, exclude_pids):
        try:
            _win_taskkill_tree(pid)
        except Exception:
            pass


# ------------------------ Конфиг/модель ------------------------------------

@dataclass
class Project:
    pid: str
    name: str
    cmd: str
    cwd: str
    args: str = ""
    enabled: bool = False
    autorestart: bool = True
    clear_log_on_start: bool = False
    status: str = "stopped"

    # runtime (не сериализуем):
    process: Optional[QtCore.QProcess] = field(
        default=None, repr=False, compare=False, init=False)
    item: Optional[QtWidgets.QTreeWidgetItem] = field(
        default=None, repr=False, compare=False, init=False)
    switch: Optional["Switch"] = field(
        default=None, repr=False, compare=False, init=False)
    log: Optional["LogView"] = field(
        default=None, repr=False, compare=False, init=False)
    stopping: bool = field(default=False, repr=False,
                           compare=False, init=False)
    waiting_network: bool = field(
        default=False, repr=False, compare=False, init=False)
    external_pid: Optional[int] = field(
        default=None, repr=False, compare=False, init=False)
    restart_pending: bool = field(
        default=False, repr=False, compare=False, init=False)

    def to_dict(self) -> dict:
        return {
            "pid": self.pid,
            "name": self.name,
            "cmd": self.cmd,
            "cwd": self.cwd, "args": self.args,
            "enabled": self.enabled,
            "autorestart": self.autorestart,
            "clear_log_on_start": self.clear_log_on_start,
            "status": self.status,
        }

    @staticmethod
    def from_dict(d: dict) -> "Project":
        return Project(
            pid=d.get("pid") or uuid.uuid4().hex,
            name=d.get("name", "Project"),
            cmd=d.get("cmd", ""),
            cwd=d.get("cwd", ""),
            args=str(d.get("args", "")),
            enabled=bool(d.get("enabled", False)),
            autorestart=bool(d.get("autorestart", True)),
            clear_log_on_start=bool(d.get("clear_log_on_start", False)),
            status=d.get("status", "stopped"),
        )


class Config:
    def __init__(self, path: Path):
        self.path = path
        self.data: Dict[str, object] = {}
        self.load()

    def load(self):
        ensure_dirs()
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text("utf-8"))
            except Exception:
                self.data = {}
        self.data.setdefault("projects", [])
        self.data.setdefault("theme", Theme.System)
        self.data.setdefault("use_mica", True)
        self.data.setdefault("autostart_run", False)
        self.data.setdefault("autostart_task", False)
        self.data.setdefault("language", Lang.System)
        self.data.setdefault("env_snapshot", {})

    def save(self):
        ensure_dirs()
        self.path.write_text(json.dumps(
            self.data, ensure_ascii=False, indent=2), "utf-8")

    def get_projects(self) -> List[Project]:
        # type: ignore[list-item]
        projects = [Project.from_dict(p)
                    for p in self.data.get("projects", [])]
        for p in projects:
            p.status = "stopped"
        return projects

    def set_projects(self, projects: List[Project]) -> None:
        self.data["projects"] = [p.to_dict() for p in projects]
        self.save()


def _safe_filename(text: str) -> str:
    if not text:
        return "log"
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", text).strip("_") or "log"


def log_path_for_project(p: "Project") -> Path:
    return LOGS_DIR / f"{_safe_filename(p.pid)}.log"


def _rotate_log_file(path: Path) -> None:
    if LOG_ROTATE_COUNT <= 1:
        try:
            path.write_text("", "utf-8")
        except Exception:
            pass
        return
    backups = max(LOG_ROTATE_COUNT - 1, 0)
    try:
        oldest = path.with_suffix(path.suffix + f".{backups}")
        if oldest.exists():
            oldest.unlink()
        for i in range(backups - 1, 0, -1):
            src = path.with_suffix(path.suffix + f".{i}")
            dst = path.with_suffix(path.suffix + f".{i + 1}")
            if src.exists():
                src.replace(dst)
        if path.exists():
            path.replace(path.with_suffix(path.suffix + ".1"))
    except Exception:
        pass


def _clear_log_backups(path: Path) -> None:
    backups = max(LOG_ROTATE_COUNT - 1, 0)
    for i in range(1, backups + 1):
        p = path.with_suffix(path.suffix + f".{i}")
        try:
            if p.exists():
                p.unlink()
        except Exception:
            pass


def append_project_log(p: "Project", text: str) -> None:
    if not text:
        return
    text = _strip_ansi(text)
    try:
        path = log_path_for_project(p)
        data = text.encode("utf-8", errors="replace")
        if LOG_ROTATE_MAX_BYTES > 0:
            try:
                size = path.stat().st_size if path.exists() else 0
                if size + len(data) > LOG_ROTATE_MAX_BYTES:
                    _rotate_log_file(path)
                    if len(data) > LOG_ROTATE_MAX_BYTES:
                        data = data[-LOG_ROTATE_MAX_BYTES:]
            except Exception:
                pass
        with path.open("ab") as f:
            f.write(data)
    except Exception:
        pass


def read_log_tail(path: Path, max_lines: int = LOG_MAX_LINES) -> str:
    if not path.exists():
        return ""
    try:
        size = path.stat().st_size
        chunk = 256 * 1024
        start = max(0, size - chunk)
        with path.open("rb") as f:
            f.seek(start)
            data = f.read()
        text = decode_bytes(data)
        lines = text.splitlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
        return ("\n".join(lines) + ("\n" if text.endswith("\n") else ""))
    except Exception:
        return ""


def read_log_from_offset(path: Path, offset: int) -> Tuple[str, int]:
    if not path.exists():
        return "", 0
    try:
        size = path.stat().st_size
        if size < offset:
            offset = 0
        with path.open("rb") as f:
            f.seek(offset)
            data = f.read()
        return decode_bytes(data), size
    except Exception:
        return "", offset


def write_state(projects: List[Project]) -> None:
    try:
        data = {
            "updated_at": time.time(),
            "projects": [
                {
                    "pid": p.pid,
                    "name": p.name,
                    "cmd": p.cmd,
                    "cwd": p.cwd,
                    "args": p.args,
                    "enabled": p.enabled,
                    "autorestart": p.autorestart,
                    "status": p.status,
                    "os_pid": (
                        int(p.process.processId())
                        if p.process and p.process.state() == QtCore.QProcess.Running
                        else (int(p.external_pid) if p.external_pid else None)
                    ),
                }
                for p in projects
            ],
        }
        tmp = STATE_PATH.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
        tmp.replace(STATE_PATH)
    except Exception:
        pass


def read_state() -> dict:
    if not STATE_PATH.exists():
        return {}
    try:
        return json.loads(STATE_PATH.read_text("utf-8"))
    except Exception:
        return {}


def is_pid_running(pid: int) -> bool:
    if not pid or pid <= 0:
        return False
    now = time.monotonic()
    cached = _PID_CACHE.get(pid)
    if cached and (now - cached[0]) < PID_CACHE_TTL_SEC:
        return cached[1]
    alive = False
    if os.name == "nt":
        try:
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            out = subprocess.check_output(
                ["tasklist", "/FO", "CSV", "/NH", "/FI", f"PID eq {pid}"],
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=flags,
            )
            if "No tasks are running" in out:
                alive = False
            else:
                # CSV line format: "Image Name","PID",...
                if re.search(rf'","{pid}"', out):
                    alive = True
                elif re.search(rf"\\b{pid}\\b", out):
                    alive = True
        except Exception:
            pass
    if not alive:
        try:
            os.kill(pid, 0)
            alive = True
        except PermissionError:
            # Process exists but we may not have rights to query it.
            alive = True
        except OSError as e:
            # On Windows, access denied may come as generic OSError (winerror=5).
            if getattr(e, "winerror", None) == 5 or getattr(e, "errno", None) in (5, 13):
                alive = True
            else:
                alive = False
        except Exception:
            alive = False
    _PID_CACHE[pid] = (now, alive)
    if len(_PID_CACHE) > 256:
        # drop oldest entries
        for k, _v in sorted(_PID_CACHE.items(), key=lambda it: it[1][0])[:64]:
            _PID_CACHE.pop(k, None)
    return alive


def read_daemon_pid() -> Optional[int]:
    if not DAEMON_PID_PATH.exists():
        return None
    try:
        pid = int(DAEMON_PID_PATH.read_text("utf-8").strip())
    except Exception:
        try:
            DAEMON_PID_PATH.unlink()
        except Exception:
            pass
        return None
    if not is_pid_running(pid):
        try:
            DAEMON_PID_PATH.unlink()
        except Exception:
            pass
        return None
    return pid


def is_daemon_running() -> bool:
    return read_daemon_pid() is not None


def is_state_fresh(max_age_sec: float = 5.0) -> bool:
    try:
        data = read_state()
        ts = float(data.get("updated_at") or 0.0)
        if ts <= 0:
            return False
        return (time.time() - ts) <= max_age_sec
    except Exception:
        return False


# ------------------------ Пользовательские виджеты -------------------------

class Switch(QtWidgets.QCheckBox):
    """Win11-подобный тумблер (виджет, а не делегат)."""

    def __init__(self, accent: QtGui.QColor, parent=None):
        super().__init__(parent)
        self.accent = accent
        self.setCursor(Qt.PointingHandCursor)
        self.setText("")
        self._h = 22
        self._w = int(self._h * 1.9)
        self.setFixedSize(self._w + 8, self._h)

    def sizeHint(self) -> QtCore.QSize:
        return QtCore.QSize(self._w + 8, self._h)

    def paintEvent(self, e: QtGui.QPaintEvent) -> None:
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        r = self.rect().adjusted(4, 0, -4, 0)
        track = QtCore.QRectF(r.x(), r.center().y() -
                              self._h / 2, self._w, self._h)
        p.setPen(QtCore.Qt.NoPen)
        bg = QtGui.QColor(self.accent) if self.isChecked(
        ) else self.palette().mid().color()
        if self.isChecked():
            bg.setAlpha(220)
        p.setBrush(bg)
        p.drawRoundedRect(track, self._h / 2, self._h / 2)
        knob = self._h - 4
        kx = track.left() + 2 if not self.isChecked() else track.right() - knob - 2
        kr = QtCore.QRectF(kx, track.top() + 2, knob, knob)
        p.setBrush(QtGui.QColor("white"))
        p.setPen(QtGui.QPen(QtGui.QColor(0, 0, 0, 30)))
        p.drawEllipse(kr)
        p.end()


class LogView(QtWidgets.QPlainTextEdit):
    def __init__(self, *, theme: str, accent: QtGui.QColor, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setLineWrapMode(QtWidgets.QPlainTextEdit.NoWrap)
        self.setFont(QtGui.QFont(
            "Cascadia Mono, Consolas, JetBrains Mono, Courier New", 10))
        # Ограничиваем размер лога, чтобы не раздувать память.
        self.document().setMaximumBlockCount(LOG_MAX_LINES)
        self.apply_palette(theme, accent)

    def apply_palette(self, theme: str, accent: QtGui.QColor):
        cols = theme_colors_hex(theme)
        pal = self.palette()
        pal.setColor(QtGui.QPalette.Base, QtGui.QColor(cols["card"]))
        pal.setColor(QtGui.QPalette.Text, QtGui.QColor(cols["text"]))
        pal.setColor(QtGui.QPalette.Highlight, accent)
        pal.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(
            "#000000" if cols["card"] == "#ffffff" else "#ffffff"))
        self.setPalette(pal)

    def append_text(self, text: str):
        text = _strip_ansi(text)
        self.moveCursor(QtGui.QTextCursor.End)
        self.insertPlainText(text)
        self.moveCursor(QtGui.QTextCursor.End)


class StatusDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, color_for_status, parent=None):
        super().__init__(parent)
        self._color_for_status = color_for_status

    def paint(self, painter, option, index):
        opt = QtWidgets.QStyleOptionViewItem(option)
        status_code = index.data(Qt.ItemDataRole.UserRole) or index.data()
        color = self._color_for_status(str(status_code or ""))
        opt.palette.setColor(QtGui.QPalette.Text, color)
        opt.palette.setColor(QtGui.QPalette.HighlightedText, color)
        opt.font.setBold(True)
        super().paint(painter, opt, index)


# ------------------------ Автозапуск Windows Run ---------------------------

def set_windows_run_autostart(enable: bool) -> None:
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE
        ) as k:
            name = "PyCombiner"
            if enable:
                cmd = get_self_executable_for_run()
                winreg.SetValueEx(k, name, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(k, name)
                except FileNotFoundError:
                    pass
    except Exception:
        traceback.print_exc()


def get_windows_username() -> str:
    user = os.environ.get("USERNAME", "") or ""
    domain = os.environ.get("USERDOMAIN", "") or os.environ.get("COMPUTERNAME", "") or ""
    if domain and user and "\\" not in user:
        return f"{domain}\\{user}"
    return user


def set_windows_task_autostart(enable: bool, username: Optional[str] = None, password: Optional[str] = None) -> Tuple[bool, str]:
    try:
        if enable:
            if not username or password is None:
                return False, "Missing credentials"
            command, args_list = get_self_run_parts(
                headless=True, data_dir=APPDATA_DIR)
            args_str = subprocess.list2cmdline(args_list) if os.name == "nt" else " ".join(
                shlex.quote(a) for a in args_list)
            xml = build_task_xml(command, args_str, username)
            xml_path = APPDATA_DIR / "autostart_task.xml"
            xml_path.write_text(xml, encoding="utf-16")
            args = [
                "schtasks",
                "/Create",
                "/F",
                "/TN",
                AUTOSTART_TASK_NAME,
                "/XML",
                str(xml_path),
                "/RU",
                username,
                "/RP",
                password,
            ]
        else:
            check = subprocess.run(
                ["schtasks", "/Query", "/TN", AUTOSTART_TASK_NAME],
                capture_output=True,
                text=True,
            )
            if check.returncode != 0:
                return True, ""
            args = ["schtasks", "/Delete", "/TN", AUTOSTART_TASK_NAME, "/F"]
        res = subprocess.run(args, capture_output=True, text=True)
        ok = (res.returncode == 0)
        msg = (res.stdout or "") + ("\n" + res.stderr if res.stderr else "")
        if enable and ok:
            try:
                if 'xml_path' in locals() and xml_path.exists():
                    xml_path.unlink()
            except Exception:
                pass
            verify = subprocess.run(
                ["schtasks", "/Query", "/TN", AUTOSTART_TASK_NAME],
                capture_output=True,
                text=True,
            )
            if verify.returncode != 0:
                return False, (msg.strip() + "\nTask verification failed.").strip()
        return ok, msg.strip()
    except Exception as e:
        return False, str(e)


def get_self_executable_for_run(*, headless: bool = False, data_dir: Optional[Path] = None) -> str:
    """
    Возвращает команду для автозапуска PyCombiner с флагом --autostart,
    корректно экранированную под текущую ОС.
    """
    if getattr(sys, "frozen", False):
        # собранный exe
        args = [str(Path(sys.executable)), "--autostart"]
    else:
        # запускаем интерпретатор + текущий скрипт
        args = [str(Path(sys.executable)), str(
            Path(sys.argv[0]).resolve()), "--autostart"]
    if headless:
        args.append("--headless")
    if data_dir:
        args += ["--data-dir", str(data_dir)]

    if os.name == "nt":
        # правильное quoting для командной строки Windows
        return subprocess.list2cmdline(args)
    else:
        # безопасное quoting для POSIX
        return " ".join(shlex.quote(a) for a in args)


def get_self_run_parts(*, headless: bool = False, data_dir: Optional[Path] = None, autostart: bool = True) -> Tuple[str, List[str]]:
    if getattr(sys, "frozen", False):
        args = [str(Path(sys.executable))]
    else:
        args = [str(Path(sys.executable)), str(Path(sys.argv[0]).resolve())]
    if headless:
        args.append("--headless")
    if autostart:
        args.append("--autostart")
    if data_dir:
        args += ["--data-dir", str(data_dir)]
    return args[0], args[1:]


def build_task_xml(command: str, arguments: str, username: str) -> str:
    ts = time.strftime("%Y-%m-%dT%H:%M:%S")
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>{ts}</Date>
    <Author>PyCombiner</Author>
  </RegistrationInfo>
  <Triggers>
    <BootTrigger>
      <Enabled>true</Enabled>
    </BootTrigger>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{username}</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{username}</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{command}</Command>
      <Arguments>{arguments}</Arguments>
    </Exec>
  </Actions>
</Task>
"""


def shutil_which(name: str, env: Optional[Dict[str, str]] = None) -> Optional[str]:
    path_value = _env_lookup(env, "PATH", "")
    for p in (path_value or "").split(os.pathsep):
        candidate = Path(p) / name
        if candidate.exists():
            return str(candidate)
    return None


def find_python_executable(env: Optional[Dict[str, str]] = None) -> Optional[str]:
    candidates: List[Path] = []
    env_value = _env_lookup(env, "PYCOMBINER_PYTHON")
    if not env_value:
        env_value = _env_lookup(env, "PYTHON_EXE")
    if env_value:
        candidates.append(Path(env_value))
    # Try PATH
    for exe in ("python.exe", "python3.exe", "py.exe"):
        p = shutil_which(exe, env=env)
        if p:
            candidates.append(Path(p))
    # Try Windows py launcher directly
    sysroot = _env_lookup(env, "SystemRoot", r"C:\Windows")
    candidates.append(Path(sysroot) / "py.exe")

    # Try registry (user + machine)
    if os.name == "nt":
        try:
            import winreg

            def _iter_install_paths(root, base):
                try:
                    with winreg.OpenKey(root, base) as key:
                        i = 0
                        while True:
                            try:
                                ver = winreg.EnumKey(key, i)
                            except OSError:
                                break
                            i += 1
                            try:
                                with winreg.OpenKey(key, ver + "\\InstallPath") as k2:
                                    try:
                                        ip, _ = winreg.QueryValueEx(k2, "")
                                        if ip:
                                            yield Path(ip) / "python.exe"
                                    except Exception:
                                        pass
                                    try:
                                        ep, _ = winreg.QueryValueEx(
                                            k2, "ExecutablePath")
                                        if ep:
                                            yield Path(ep)
                                    except Exception:
                                        pass
                            except Exception:
                                continue
                except Exception:
                    return

            reg_paths = [
                (winreg.HKEY_CURRENT_USER, r"Software\Python\PythonCore"),
                (winreg.HKEY_LOCAL_MACHINE, r"Software\Python\PythonCore"),
                (winreg.HKEY_LOCAL_MACHINE,
                 r"Software\Wow6432Node\Python\PythonCore"),
            ]
            for root, base in reg_paths:
                for p in _iter_install_paths(root, base):
                    candidates.append(p)
        except Exception:
            pass

    for c in candidates:
        try:
            if c.exists():
                return str(c)
        except Exception:
            continue
    return None


def program_and_args_for_cmd(cmd: str, env: Optional[Dict[str, str]] = None) -> Tuple[str, List[str]]:
    path = Path(cmd.strip().strip('"'))
    suf = path.suffix.lower()
    if suf == ".ps1":
        pwsh = shutil_which("pwsh.exe", env=env) or shutil_which("pwsh", env=env)
        if pwsh:
            prog = pwsh
        else:
            prog = shutil_which("powershell.exe", env=env) or shutil_which("powershell", env=env) or "powershell.exe"
        return prog, ["-NoLogo", "-ExecutionPolicy", "Bypass", "-File", str(path)]
    if suf in (".bat", ".cmd"):
        env_cmd = _env_lookup(env, "ComSpec")
        if not env_cmd:
            env_cmd = _env_lookup(env, "COMSPEC")
        cmd_exe = env_cmd or shutil_which("cmd.exe", env=env) or "cmd.exe"
        return cmd_exe, ["/c", str(path)]
    if suf == ".exe":
        return str(path), []
    if suf == ".py":
        # запуск .py: в сборке sys.executable указывает на PyCombiner.exe,
        # поэтому ищем интерпретатор рядом со скриптом или используем системный.
        venv_py = None
        try:
            cand = path.parent / ".venv"
            if os.name == "nt":
                vp = cand / "Scripts" / "python.exe"
            else:
                vp = cand / "bin" / "python3"
            if vp.exists():
                venv_py = str(vp)
        except Exception:
            pass
        if venv_py:
            return venv_py, ["-u", str(path)]
        if not getattr(sys, "frozen", False):
            return sys.executable, ["-u", str(path)]
        py_exec = find_python_executable(env=env)
        if py_exec:
            if os.path.basename(py_exec).lower().startswith("py"):
                return py_exec, ["-3", "-u", str(path)]
            return py_exec, ["-u", str(path)]
        if os.name == "nt":
            return "py", ["-3", "-u", str(path)]
        return "python3", ["-u", str(path)]
    parts = shlex.split(cmd, posix=False)
    return (parts[0], parts[1:]) if parts else (cmd, [])


class HeadlessController(QtCore.QObject):
    def __init__(self, cfg: Config, *, autostart: bool):
        super().__init__()
        self.cfg = cfg
        self._language = cfg.data.get("language", Lang.System)
        self._env_snapshot = _normalize_env_snapshot(cfg.data.get("env_snapshot"))
        if not self._env_snapshot:
            log_app("Headless: env snapshot missing, using system environment")
        self._autostart = bool(autostart)
        self._started_at = time.time()
        self.projects: List[Project] = cfg.get_projects()
        enabled_count = sum(1 for p in self.projects if p.enabled)
        log_app(
            f"Headless: config={CONFIG_PATH} projects={len(self.projects)} enabled={enabled_count} autostart={self._autostart}"
        )
        self._cmd_timer = QtCore.QTimer(self)
        self._cmd_timer.timeout.connect(self._process_commands)
        self._cmd_timer.start(1000)
        self._state_timer = QtCore.QTimer(self)
        self._state_timer.timeout.connect(self._write_state)
        self._state_timer.start(1000)

        self._write_pid()
        self._write_state()
        self._cleanup_stale_commands()

        app = QtCore.QCoreApplication.instance()
        if app:
            app.aboutToQuit.connect(self._cleanup)

        if autostart:
            QtCore.QTimer.singleShot(300, self._autostart_when_network_ready)

    def _autostart_when_network_ready(self) -> None:
        if is_network_ready():
            for p in self.projects:
                if p.waiting_network:
                    p.waiting_network = False
                    if p.status == "waiting":
                        p.status = "stopped"
            self._write_state()
            self.start_enabled()
            return
        any_waiting = False
        for p in self.projects:
            if p.enabled:
                p.waiting_network = True
                p.status = "waiting"
                any_waiting = True
        if any_waiting:
            log_app("Headless: waiting for network")
        self._write_state()
        QtCore.QTimer.singleShot(
            NETWORK_CHECK_INTERVAL_SEC * 1000, self._autostart_when_network_ready
        )

    def _write_pid(self):
        try:
            DAEMON_PID_PATH.write_text(str(os.getpid()), "utf-8")
        except Exception:
            pass

    def _cleanup(self):
        try:
            self.stop_all(reason="headless_exit")
        except Exception:
            pass
        try:
            if DAEMON_PID_PATH.exists():
                DAEMON_PID_PATH.unlink()
        except Exception:
            pass

    def _write_state(self):
        write_state(self.projects)

    def _is_stale_command(self, data: dict) -> bool:
        try:
            ts = float(data.get("ts") or 0)
        except Exception:
            ts = 0.0
        return ts and ts < (self._started_at - 2)

    def _cleanup_stale_commands(self) -> None:
        ensure_dirs()
        for path in COMMANDS_DIR.glob("cmd-*.json"):
            try:
                data = json.loads(path.read_text("utf-8"))
            except Exception:
                data = {}
            if self._is_stale_command(data):
                try:
                    log_app(f"Headless: drop stale cmd {path.name}")
                    path.unlink()
                except Exception:
                    pass

    def _process_commands(self):
        ensure_dirs()
        for path in sorted(COMMANDS_DIR.glob("cmd-*.json")):
            try:
                data = json.loads(path.read_text("utf-8"))
            except Exception:
                try:
                    path.unlink()
                except Exception:
                    pass
                continue
            if self._is_stale_command(data):
                try:
                    log_app(f"Headless: ignore stale cmd {data.get('action')} pid={data.get('pid')}")
                finally:
                    try:
                        path.unlink()
                    except Exception:
                        pass
                continue
            action = (data.get("action") or "").lower()
            pid = data.get("pid")
            try:
                log_app(f"Headless: cmd={action} pid={pid}")
                if action == "start" and pid:
                    self.start_project_by_pid(pid)
                elif action == "stop" and pid:
                    self.stop_project_by_pid(pid, reason="cmd_stop")
                elif action == "restart" and pid:
                    self.restart_project_by_pid(pid)
                elif action == "stop_all":
                    self.stop_all(reason="cmd_stop_all")
                elif action == "start_enabled":
                    self.start_enabled()
                elif action == "reload":
                    self.reload_config()
            finally:
                try:
                    path.unlink()
                except Exception:
                    pass

    def reload_config(self):
        self.cfg.load()
        self._env_snapshot = _normalize_env_snapshot(self.cfg.data.get("env_snapshot"))
        new_projects = {p.pid: p for p in self.cfg.get_projects()}
        current = {p.pid: p for p in self.projects}
        # stop projects removed from config
        for pid, p in list(current.items()):
            if pid not in new_projects:
                self.stop_project(p, reason="removed_from_config")
                current.pop(pid, None)
        # update existing and add new
        updated: List[Project] = []
        for pid, np in new_projects.items():
            if pid in current:
                cp = current[pid]
                cp.name = np.name
                cp.cmd = np.cmd
                cp.cwd = np.cwd
                cp.args = np.args
                cp.enabled = np.enabled
                cp.autorestart = np.autorestart
                updated.append(cp)
            else:
                updated.append(np)
        self.projects = updated
        self._write_state()

    def start_enabled(self):
        enabled = [p for p in self.projects if p.enabled]
        log_app(f"Headless: start_enabled -> {len(enabled)} project(s)")
        for p in self.projects:
            if p.enabled and not (p.process and p.process.state() == QtCore.QProcess.Running):
                self.start_project(p)

    def start_project_by_pid(self, pid: str):
        for p in self.projects:
            if p.pid == pid:
                self.start_project(p)
                return

    def restart_project_by_pid(self, pid: str):
        for p in self.projects:
            if p.pid == pid:
                self.restart_project(p)
                return

    def stop_project_by_pid(self, pid: str, reason: str = ""):
        for p in self.projects:
            if p.pid == pid:
                self.stop_project(p, reason=reason or "cmd_stop")
                return

    def restart_project(self, p: Project) -> None:
        if p.process and p.process.state() == QtCore.QProcess.Running:
            p.restart_pending = True
            self.stop_project(p, reason="cmd_restart")
            return
        # нет активного процесса — останавливаем зомби и сразу запускаем
        p.restart_pending = False
        p.stopping = False
        self.stop_project(p, reason="cmd_restart")
        self.start_project(p)

    def stop_all(self, reason: str = ""):
        for p in self.projects:
            self.stop_project(p, reason=reason or "stop_all")

    def start_project(self, p: Project):
        if p.process and p.process.state() == QtCore.QProcess.Running:
            return
        if not p.cmd:
            log_app(f"Headless: skip start {p.name} (empty cmd)")
            return
        if p.stopping:
            return

        exclude = set()
        for q in self.projects:
            if q.process and q.process.state() == QtCore.QProcess.Running:
                try:
                    exclude.add(int(q.process.processId()))
                except Exception:
                    pass
        conflict_running = any((q is not p) and q.process and q.process.state() == QtCore.QProcess.Running and (
            q.cmd == p.cmd or (p.cwd and q.cwd and q.cwd == p.cwd)) for q in self.projects)
        existing_pids: List[int] = []
        if not conflict_running:
            existing_pids = _win_find_project_pids(p.cmd, p.cwd, exclude)
            if not existing_pids:
                _win_kill_project_zombies(p.cmd, p.cwd, exclude)
        if existing_pids:
            p.external_pid = existing_pids[0]
            p.status = "running"
            self._write_state()
            append_project_log(
                p,
                self._format_log(self._tr("log_adopted", pid=p.external_pid)),
            )
            log_app(f"Headless: adopt {p.name} pid={p.external_pid}")
            return
        p.waiting_network = False
        if p.clear_log_on_start:
            try:
                path = log_path_for_project(p)
                path.write_text("", "utf-8")
                _clear_log_backups(path)
            except Exception:
                pass
        env_snapshot = self._env_snapshot if self._env_snapshot else None
        program, args = program_and_args_for_cmd(p.cmd, env=env_snapshot)
        extra = (p.args or '').strip()
        if extra:
            try:
                args += shlex.split(extra, posix=False)
            except Exception:
                args += extra.split()

        proc = QtCore.QProcess(self)
        env = build_process_environment(self._env_snapshot)
        env.insert("PYTHONIOENCODING", "utf-8")
        env.insert("COMBINER", "1")
        proc.setProcessEnvironment(env)
        proc.setProgram(program)
        proc.setArguments(args)
        if p.cwd:
            proc.setWorkingDirectory(p.cwd)
        proc.setProcessChannelMode(QtCore.QProcess.MergedChannels)
        log_app(f"Headless: start {p.name} -> {program} {args} cwd={p.cwd}")

        proc.readyReadStandardOutput.connect(
            lambda p_=p, pr=proc: self._on_proc_output(p_, pr))
        proc.readyReadStandardError.connect(
            lambda p_=p, pr=proc: self._on_proc_output(p_, pr))
        proc.finished.connect(lambda code, status,
                              p_=p, pr=proc: self._on_proc_finished(p_, pr, code, status))
        proc.errorOccurred.connect(
            lambda err, p_=p, pr=proc: self._on_proc_error(p_, pr, err))
        proc.started.connect(lambda p_=p: self._on_proc_started(p_))

        p.stopping = False
        p.external_pid = None
        p.status = "starting"
        self._write_state()
        try:
            proc.start()
        except Exception as e:
            log_app(f"Headless: start failed {p.name}: {e}")
            append_project_log(p, self._format_log(
                self._tr("log_start_error", err=e)))
            p.status = "stopped"
            self._write_state()
            return

        p.process = proc
        append_project_log(
            p,
            self._format_log(self._tr("log_start", cmd=p.cmd)),
        )

    def stop_project(self, p: Project, reason: str = ""):
        if reason:
            log_app(f"Headless: stop {p.name} pid={p.pid} reason={reason}")
        p.waiting_network = False
        if p.process and p.process.state() == QtCore.QProcess.Running:
            p.stopping = True
            p.status = "stopping"
            self._write_state()
            pid = int(p.process.processId() or 0)
            try:
                p.process.terminate()
                if not p.process.waitForFinished(1500):
                    _win_taskkill_tree(pid)
            except Exception:
                _win_taskkill_tree(pid)
        else:
            if p.external_pid and is_pid_running(int(p.external_pid)):
                try:
                    _win_taskkill_tree(int(p.external_pid))
                except Exception:
                    pass
            else:
                _win_kill_project_zombies(p.cmd, p.cwd)
            p.status = "stopped"
            p.external_pid = None
            self._write_state()
        append_project_log(p, self._format_log(self._tr("log_stop")))

    def _on_proc_output(self, p: Project, pr: QtCore.QProcess):
        data = pr.readAll().data()
        text = decode_bytes(data)
        if text:
            append_project_log(p, text)

    def _on_proc_started(self, p: Project):
        p.status = "running"
        p.external_pid = None
        p.waiting_network = False
        self._write_state()
        log_app(f"Headless: running {p.name}")

    def _on_proc_finished(self, p: Project, pr: QtCore.QProcess, code: int, status: QtCore.QProcess.ExitStatus):
        append_project_log(
            p,
            self._format_log(self._tr(
                "log_finish",
                code=code,
                status=('CrashExit' if status == QtCore.QProcess.CrashExit else 'NormalExit')
            )),
        )
        log_app(f"Headless: finished {p.name} code={code} status={status}")
        if p.process is not pr:
            return
        was_stopping = p.stopping
        p.status = "stopped" if (was_stopping or (status == QtCore.QProcess.NormalExit and code == 0)) else "crashed"
        self._write_state()

        pr_autorestart = (
            p.autorestart and not was_stopping and status == QtCore.QProcess.CrashExit)
        p.process = None
        p.external_pid = None
        if p.stopping:
            p.stopping = False
        if p.restart_pending:
            p.restart_pending = False
            QtCore.QTimer.singleShot(300, lambda: self.start_project(p))
            return
        if pr_autorestart:
            QtCore.QTimer.singleShot(2000, lambda: self.start_project(p))

    def _on_proc_error(self, p: Project, pr: QtCore.QProcess, err: QtCore.QProcess.ProcessError):
        if p.process is not pr:
            return
        append_project_log(p, self._format_log(
            self._tr("log_proc_error", err=err)))
        self._write_state()

    def _format_log(self, text: str) -> str:
        return f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] {text}\n"

    def _tr(self, key: str, **kwargs) -> str:
        return tr(self._language, key, **kwargs)


# ------------------------ Главное окно -------------------------------------

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, cfg: Config):
        super().__init__()
        self.cfg = cfg
        self.projects: List[Project] = cfg.get_projects()
        self._start_enabled_queue: List[Project] = []
        self._theme = self.cfg.data.get("theme", Theme.System)
        self._accent = argb_to_qcolor(get_accent_color_argb())
        self._language = self.cfg.data.get("language", Lang.System)
        self._syncing_selection = False
        self._syncing_autostart_task = False
        self._client_mode = self._should_use_client_mode()
        self._log_offsets: Dict[str, int] = {}
        self._pending_actions: Dict[str, Tuple[str, float]] = {}
        self._last_headless_spawn = 0.0

        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setWindowIcon(load_app_icon())

        self._really_quit = False
        self._build_tray()
        self._build_ui()
        self._populate_projects()
        self.apply_language()
        self.apply_theme()
        QtCore.QTimer.singleShot(0, self.apply_theme)
        self._update_daemon_indicator()
        if self._client_mode:
            self._state_timer = QtCore.QTimer(self)
            self._state_timer.timeout.connect(self._refresh_state_from_file)
            self._state_timer.start(1000)
            self._log_timer = QtCore.QTimer(self)
            self._log_timer.timeout.connect(self._poll_log_updates)
            self._log_timer.start(700)
            QtCore.QTimer.singleShot(0, self._refresh_state_from_file)
            self._daemon_watch_timer = QtCore.QTimer(self)
            self._daemon_watch_timer.timeout.connect(self._monitor_daemon)
            self._daemon_watch_timer.start(1500)
        else:
            QtCore.QTimer.singleShot(300, self.on_start_enabled)

    def _should_use_client_mode(self) -> bool:
        if is_daemon_running():
            return True
        if is_state_fresh(5.0):
            return True
        return bool(self.cfg.data.get("autostart_task", False))

    def _check_daemon_ready(self) -> None:
        if is_daemon_running():
            if getattr(self, "_wait_daemon_timer", None):
                self._wait_daemon_timer.stop()
            self._refresh_state_from_file()
        self._update_daemon_indicator()

    def _spawn_headless(self) -> bool:
        try:
            cmd, args = get_self_run_parts(
                headless=True, data_dir=APPDATA_DIR, autostart=True)
            env = dict(os.environ)
            # Drop PyInstaller/Qt injection from GUI so headless starts clean.
            for k in list(env.keys()):
                if k.startswith("_PYI_"):
                    env.pop(k, None)
            for k in ("QT_PLUGIN_PATH", "QML2_IMPORT_PATH", "PYTHONPATH", "PYTHONHOME"):
                env.pop(k, None)
            flags = 0
            if os.name == "nt":
                flags |= getattr(subprocess, "CREATE_NO_WINDOW", 0)
                flags |= getattr(subprocess, "DETACHED_PROCESS", 0)
            subprocess.Popen(
                [cmd] + args,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=flags,
                env=env,
            )
            log_app("GUI: spawned headless controller")
            return True
        except Exception as e:
            log_app(f"GUI: failed to spawn headless: {e}")
            return False

    def _monitor_daemon(self) -> None:
        self._update_daemon_indicator()
        if is_daemon_running():
            return
        if not bool(self.cfg.data.get("autostart_task", False)):
            return
        now = time.time()
        if now - self._last_headless_spawn < 10:
            return
        if self._spawn_headless():
            self._last_headless_spawn = now

    def _update_daemon_indicator(self) -> None:
        if not getattr(self, "lbl_daemon", None):
            return
        pid = read_daemon_pid()
        if pid:
            text = self._tr("headless_connected", pid=pid)
            color = self._status_color("running").name()
        else:
            text = self._tr("headless_disconnected")
            color = self._status_color("stopped").name()
        self.lbl_daemon.setText(text)
        self.lbl_daemon.setStyleSheet(f"color: {color}; font-weight: 600;")
    # ---------- UI ----------

    def _build_ui(self):
        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)
        v = QtWidgets.QVBoxLayout(cw)
        v.setContentsMargins(12, 8, 12, 12)
        v.setSpacing(10)

        self._build_menu()

        card = QtWidgets.QFrame(objectName="Card")
        v.addWidget(card, 1)
        main = QtWidgets.QVBoxLayout(card)
        main.setContentsMargins(10, 10, 10, 10)
        main.setSpacing(10)

        # кнопки
        btns = QtWidgets.QHBoxLayout()
        main.addLayout(btns)
        self.btn_add = QtWidgets.QPushButton("Добавить")
        self.btn_edit = QtWidgets.QPushButton("Изменить")
        self.btn_del = QtWidgets.QPushButton("Удалить")
        self.btn_start = QtWidgets.QPushButton("Старт")
        self.btn_stop = QtWidgets.QPushButton("Стоп")
        self.btn_restart = QtWidgets.QPushButton("Перезапуск")
        self.btn_start_enabled = QtWidgets.QPushButton("Старт (включённые)")
        self.btn_stop_all = QtWidgets.QPushButton("Стоп все")
        for b in (self.btn_add, self.btn_edit, self.btn_del, self.btn_start, self.btn_stop, self.btn_restart, self.btn_start_enabled, self.btn_stop_all):
            btns.addWidget(b)
        btns.addStretch(1)
        self.lbl_daemon = QtWidgets.QLabel()
        self.lbl_daemon.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        btns.addWidget(self.lbl_daemon)

        # список проектов
        self.tree = QtWidgets.QTreeWidget()
        self.tree.setHeaderLabels(
            ["Вкл.", "Имя", "Статус", "Команда", "Рабочая папка"])
        self.tree.setRootIsDecorated(False)
        self.tree.setIndentation(0)
        self.tree.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.tree.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tree.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._status_delegate = StatusDelegate(self._status_color, self.tree)
        self.tree.setItemDelegateForColumn(2, self._status_delegate)
        main.addWidget(self.tree, 1)

        hdr = self.tree.header()
        hdr.setMinimumSectionSize(80)
        hdr.setStretchLastSection(False)
        from PySide6.QtWidgets import QHeaderView
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        self.tree.setColumnWidth(0, 76)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.Stretch)
        hdr.setSectionResizeMode(4, QHeaderView.Stretch)

        # логи
        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setDocumentMode(True)
        self.btn_clear_log = QtWidgets.QToolButton()
        self.btn_clear_log.setText("Очистить лог")
        self.btn_clear_log.setCursor(Qt.PointingHandCursor)
        self.btn_clear_log.setToolTip("Очистить лог текущей вкладки")
        self.btn_clear_log.clicked.connect(self.on_clear_log)
        self.tabs.setCornerWidget(self.btn_clear_log, Qt.TopRightCorner)
        main.addWidget(self.tabs, 2)

        # сигналы
        self.tree.itemDoubleClicked.connect(lambda *_: self.on_edit())
        QtGui.QShortcut(QtGui.QKeySequence("F2"), self, activated=self.on_edit)
        QtGui.QShortcut(QtGui.QKeySequence("Delete"),
                        self, activated=self.on_delete)
        QtGui.QShortcut(QtGui.QKeySequence("Ctrl+L"),
                        self, activated=self.on_clear_log)

        self.btn_add.clicked.connect(self.on_add)
        self.btn_edit.clicked.connect(self.on_edit)
        self.btn_del.clicked.connect(self.on_delete)
        self.btn_start.clicked.connect(self.on_start_selected)
        self.btn_stop.clicked.connect(self.on_stop_selected)
        self.btn_restart.clicked.connect(self.on_restart_selected)
        self.btn_start_enabled.clicked.connect(self.on_start_enabled)
        self.btn_stop_all.clicked.connect(self.on_stop_all)

        self.tree.itemSelectionChanged.connect(self._on_selection_changed)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self._refresh_action_buttons()

    def _build_menu(self):
        mb = self.menuBar()
        self.menu_file = mb.addMenu("")
        self.act_exit = self.menu_file.addAction("")
        self.act_exit.triggered.connect(self._quit_from_tray)

        self.menu_settings = mb.addMenu("")
        self.act_autostart = self.menu_settings.addAction("", lambda: None)
        self.act_autostart.setCheckable(True)
        self.act_autostart.setChecked(
            bool(self.cfg.data.get("autostart_run", False)))
        self.act_autostart.toggled.connect(self.on_toggle_autostart)
        self.act_autostart_task = self.menu_settings.addAction("", lambda: None)
        self.act_autostart_task.setCheckable(True)
        self.act_autostart_task.setChecked(
            bool(self.cfg.data.get("autostart_task", False)))
        self.act_autostart_task.toggled.connect(self.on_toggle_autostart_task)

        self.menu_appearance = mb.addMenu("")
        self.act_theme_system = self.menu_appearance.addAction("")
        self.act_theme_light = self.menu_appearance.addAction("")
        self.act_theme_dark = self.menu_appearance.addAction("")
        for a in (self.act_theme_system, self.act_theme_light, self.act_theme_dark):
            a.setCheckable(True)
        self.act_theme_system.triggered.connect(
            lambda: self.set_theme(Theme.System))
        self.act_theme_light .triggered.connect(
            lambda: self.set_theme(Theme.Light))
        self.act_theme_dark  .triggered.connect(
            lambda: self.set_theme(Theme.Dark))

        self.menu_appearance.addSeparator()
        self.act_use_mica = self.menu_appearance.addAction("")
        self.act_use_mica.setCheckable(True)
        self.act_use_mica.setChecked(bool(self.cfg.data.get("use_mica", True)))
        self.act_use_mica.triggered.connect(self.apply_theme)

        self.menu_language = mb.addMenu("")
        self.act_lang_system = self.menu_language.addAction("")
        self.act_lang_ru = self.menu_language.addAction("")
        self.act_lang_en = self.menu_language.addAction("")
        for a in (self.act_lang_system, self.act_lang_ru, self.act_lang_en):
            a.setCheckable(True)
        lang_group = QtGui.QActionGroup(self.menu_language)
        lang_group.setExclusive(True)
        for a in (self.act_lang_system, self.act_lang_ru, self.act_lang_en):
            lang_group.addAction(a)
        self.act_lang_system.triggered.connect(
            lambda: self.set_language(Lang.System))
        self.act_lang_ru.triggered.connect(lambda: self.set_language(Lang.RU))
        self.act_lang_en.triggered.connect(lambda: self.set_language(Lang.EN))

        self.menu_help = mb.addMenu("")
        self.act_about = self.menu_help.addAction("", self.on_about)

    def set_theme(self, theme: str):
        self.cfg.data["theme"] = theme
        self.cfg.save()
        self.apply_theme()

    def set_language(self, lang: str):
        self._language = lang
        self.cfg.data["language"] = lang
        self.cfg.save()
        self.apply_language()

    def _lang(self) -> str:
        return resolve_lang(self._language)

    def _tr(self, key: str, **kwargs) -> str:
        return tr(self._language, key, **kwargs)

    def apply_theme(self):
        # type: ignore[assignment]
        theme = self.cfg.data.get("theme", Theme.System)
        is_light = (theme == Theme.Light) or (
            theme == Theme.System and get_system_is_light())
        accent = argb_to_qcolor(get_accent_color_argb())
        self._theme = theme
        self._accent = accent

        self.setStyleSheet(build_qss(theme, accent))

        # Заголовок/фон окна
        if self.cfg.data.get("use_mica", True) and is_light and get_win_build() >= 22000:
            enable_mica_and_titlebar(self, mica_light=True, dark_title=False)
        else:
            enable_mica_and_titlebar(
                self, mica_light=False, dark_title=not is_light)

        # Логи перекрасить (важно для старта в тёмной теме)
        for log in self.findChildren(LogView):
            log.apply_palette(theme, accent)

        # обновим цвет акцента у тумблеров
        for p in self.projects:
            if p.switch:
                p.switch.accent = accent
                p.switch.update()
            self._apply_status_style(p)
            self._apply_tab_status(p)
        self._update_daemon_indicator()

        # обновить состояние меню
        self.act_theme_system.setChecked(theme == Theme.System)
        self.act_theme_light .setChecked(theme == Theme.Light)
        self.act_theme_dark  .setChecked(theme == Theme.Dark)
        self.act_use_mica.setChecked(bool(self.cfg.data.get("use_mica", True)))

    def apply_language(self):
        # Кнопки
        self.btn_add.setText(self._tr("btn_add"))
        self.btn_edit.setText(self._tr("btn_edit"))
        self.btn_del.setText(self._tr("btn_del"))
        self.btn_start.setText(self._tr("btn_start"))
        self.btn_stop.setText(self._tr("btn_stop"))
        self.btn_restart.setText(self._tr("btn_restart"))
        self.btn_start_enabled.setText(self._tr("btn_start_enabled"))
        self.btn_stop_all.setText(self._tr("btn_stop_all"))

        # Заголовки таблицы
        self.tree.setHeaderLabels([
            self._tr("header_enabled"),
            self._tr("header_name"),
            self._tr("header_status"),
            self._tr("header_cmd"),
            self._tr("header_cwd"),
        ])

        # Вкладки логов
        self.btn_clear_log.setText(self._tr("tab_clear"))
        self.btn_clear_log.setToolTip(self._tr("tab_clear_tip"))

        self._update_daemon_indicator()

        # Меню
        self.menu_file.setTitle(self._tr("menu_file"))
        self.menu_settings.setTitle(self._tr("menu_settings"))
        self.menu_appearance.setTitle(self._tr("menu_appearance"))
        self.menu_language.setTitle(self._tr("menu_language"))
        self.menu_help.setTitle(self._tr("menu_help"))

        self.act_exit.setText(self._tr("act_exit"))
        self.act_autostart.setText(self._tr("act_autostart"))
        if getattr(self, "act_autostart_task", None):
            self.act_autostart_task.setText(self._tr("act_autostart_task"))
        self.act_theme_system.setText(self._tr("act_theme_system"))
        self.act_theme_light.setText(self._tr("act_theme_light"))
        self.act_theme_dark.setText(self._tr("act_theme_dark"))
        self.act_use_mica.setText(self._tr("act_use_mica"))
        self.act_about.setText(self._tr("act_about"))
        self.act_lang_system.setText(self._tr("lang_system"))
        self.act_lang_ru.setText(self._tr("lang_ru"))
        self.act_lang_en.setText(self._tr("lang_en"))

        self.act_lang_system.setChecked(self._language == Lang.System)
        self.act_lang_ru.setChecked(self._language == Lang.RU)
        self.act_lang_en.setChecked(self._language == Lang.EN)

        # Трей
        if getattr(self, "tray", None):
            if getattr(self, "act_tray_show", None):
                self.act_tray_show.setText(self._tr("tray_show"))
            if getattr(self, "act_tray_start_enabled", None):
                self.act_tray_start_enabled.setText(
                    self._tr("tray_start_enabled"))
            if getattr(self, "act_tray_exit", None):
                self.act_tray_exit.setText(self._tr("tray_exit"))

        # Обновим подписи статусов
        for p in self.projects:
            if p.item:
                p.item.setText(2, self._status_label(p.status))
                p.item.setData(2, Qt.ItemDataRole.UserRole, p.status)
        self.tree.viewport().update()

    def _status_label(self, status: str) -> str:
        return status_label(self._language, status)

    def _status_color(self, status: str) -> QtGui.QColor:
        status = (status or "").strip().lower()
        theme = self._theme
        is_light = (theme == Theme.Light) or (
            theme == Theme.System and get_system_is_light())
        palette = {
            "running": "#16a34a" if is_light else "#22c55e",
            "starting": "#d97706" if is_light else "#f59e0b",
            "waiting": "#d97706" if is_light else "#f59e0b",
            "stopping": "#d97706" if is_light else "#f59e0b",
            "stopped": "#dc2626" if is_light else "#f87171",
            "crashed": "#b91c1c" if is_light else "#ef4444",
        }
        return QtGui.QColor(palette.get(status, self._accent.name()))

    def _apply_status_style(self, p: Project) -> None:
        if not p.item:
            return
        p.item.setForeground(2, QtGui.QBrush(self._status_color(p.status)))
        font = p.item.font(2)
        font.setBold(True)
        p.item.setFont(2, font)

    def _apply_tab_status(self, p: Project) -> None:
        if not p.log:
            return
        idx = self.tabs.indexOf(p.log)
        if idx < 0:
            return
        if (p.status or "").strip().lower() == "crashed":
            self.tabs.tabBar().setTabTextColor(
                idx, self._status_color(p.status))
        else:
            self.tabs.tabBar().setTabTextColor(idx, QtGui.QColor())

    def _project_by_tab_index(self, index: int) -> Optional[Project]:
        if index < 0 or index >= len(self.projects):
            return None
        return self.projects[index]

    def _load_log_tail(self, p: Project) -> None:
        if not p.log:
            return
        path = log_path_for_project(p)
        text = read_log_tail(path, LOG_MAX_LINES)
        p.log.setPlainText(text)
        try:
            p.log.moveCursor(QtGui.QTextCursor.End)
            sb = p.log.verticalScrollBar()
            sb.setValue(sb.maximum())
        except Exception:
            pass
        try:
            self._log_offsets[p.pid] = path.stat().st_size
        except Exception:
            self._log_offsets[p.pid] = 0

    def _clear_log_for_project(self, p: Project) -> None:
        if p.log:
            p.log.clear()
        path = log_path_for_project(p)
        try:
            path.write_text("", "utf-8")
        except Exception:
            pass
        _clear_log_backups(path)
        self._log_offsets[p.pid] = 0

    def _poll_log_updates(self) -> None:
        idx = self.tabs.currentIndex()
        p = self._project_by_tab_index(idx)
        if not p or not p.log:
            return
        path = log_path_for_project(p)
        offset = self._log_offsets.get(p.pid, 0)
        text, new_offset = read_log_from_offset(path, offset)
        if text:
            try:
                sb = p.log.verticalScrollBar()
                at_bottom = sb.value() >= (sb.maximum() - 2)
                if at_bottom:
                    p.log.append_text(text)
                    sb.setValue(sb.maximum())
                else:
                    val = sb.value()
                    p.log.append_text(text)
                    sb.setValue(val)
            except Exception:
                p.log.append_text(text)
        self._log_offsets[p.pid] = new_offset

    def _refresh_state_from_file(self) -> None:
        state = read_state()
        items = {p.get("pid"): p for p in state.get("projects", []) if p.get("pid")}
        for p in self.projects:
            sp = items.get(p.pid)
            if not sp:
                continue
            new_status = sp.get("status", p.status)
            if self._should_ignore_status_update(p, new_status):
                continue
            if new_status != p.status:
                p.status = new_status
                self._update_row_status(p)
        self._update_daemon_indicator()

    def _send_command(self, action: str, pid: Optional[str] = None) -> None:
        ensure_dirs()
        payload = {
            "id": uuid.uuid4().hex,
            "action": action,
            "pid": pid,
            "ts": time.time(),
        }
        path = COMMANDS_DIR / f"cmd-{payload['id']}.json"
        tmp = COMMANDS_DIR / f"cmd-{payload['id']}.tmp"
        try:
            tmp.write_text(json.dumps(payload, ensure_ascii=False), "utf-8")
            tmp.replace(path)
        except Exception:
            pass

    def _apply_app_autostart_settings(self, data: dict) -> None:
        run_enabled = data.get("app_autostart_run")
        task_enabled = data.get("app_autostart_task")
        if run_enabled is not None and getattr(self, "act_autostart", None):
            if bool(self.cfg.data.get("autostart_run", False)) != bool(run_enabled):
                self.act_autostart.setChecked(bool(run_enabled))
        if task_enabled is not None and getattr(self, "act_autostart_task", None):
            if bool(self.cfg.data.get("autostart_task", False)) != bool(task_enabled):
                self.act_autostart_task.setChecked(bool(task_enabled))

    # ---------- заполнение ----------
    def _populate_projects(self):
        self.tree.clear()
        self.tabs.clear()
        # type: ignore[assignment]
        theme = self.cfg.data.get("theme", Theme.System)
        accent = argb_to_qcolor(get_accent_color_argb())

        for p in self.projects:
            item = QtWidgets.QTreeWidgetItem()
            item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            item.setText(1, p.name)
            item.setText(2, self._status_label(p.status))
            item.setData(2, Qt.ItemDataRole.UserRole, p.status)
            item.setText(3, p.cmd)
            item.setText(4, p.cwd)
            item.setData(0, Qt.ItemDataRole.UserRole, p.pid)
            self.tree.addTopLevelItem(item)
            p.item = item
            self._apply_status_style(p)

            sw = Switch(accent, self.tree)
            sw.setChecked(p.enabled)
            sw.toggled.connect(
                lambda checked, proj=p: self._on_switch_toggled(proj, checked))
            self.tree.setItemWidget(item, 0, sw)
            p.switch = sw

            te = LogView(theme=theme, accent=accent, parent=self.tabs)
            self.tabs.addTab(te, p.name)
            p.log = te
            self._apply_tab_status(p)
            self._log_offsets[p.pid] = 0
            if self._client_mode:
                self._load_log_tail(p)

        if self.tree.topLevelItemCount() > 0:
            self.tree.setCurrentItem(self.tree.topLevelItem(0))
            self.tabs.setCurrentIndex(0)
        self._refresh_action_buttons()

    def _on_switch_toggled(self, p: Project, checked: bool) -> None:
        p.enabled = checked
        self.cfg.set_projects(self.projects)
        if self._client_mode:
            self._send_command("reload")

    def _selected_project(self) -> Optional[Project]:
        it = self.tree.currentItem()
        if not it:
            return None
        pid = it.data(0, Qt.ItemDataRole.UserRole)
        for p in self.projects:
            if p.pid == pid:
                return p
        return None

    def _refresh_action_buttons(self) -> None:
        p = self._selected_project()
        if not p:
            for btn in (self.btn_start, self.btn_stop, self.btn_restart, self.btn_edit, self.btn_del):
                btn.setEnabled(False)
            self.btn_start_enabled.setEnabled(False)
            self.btn_stop_all.setEnabled(False)
            return

        status = (p.status or "").strip().lower()
        can_start = status in ("stopped", "crashed")
        can_stop = status in ("running", "starting", "stopping", "waiting")
        can_restart = status in ("running", "starting", "stopped", "crashed")

        self.btn_start.setEnabled(can_start)
        self.btn_stop.setEnabled(can_stop)
        self.btn_restart.setEnabled(can_restart and status != "stopping")
        self.btn_edit.setEnabled(True)
        self.btn_del.setEnabled(status in ("stopped", "crashed") or (not self._client_mode))

        can_start_enabled = any((pr.status or "").strip().lower() in ("stopped", "crashed") and pr.enabled for pr in self.projects)
        can_stop_all = any((pr.status or "").strip().lower() in ("running", "starting", "stopping", "waiting") for pr in self.projects)
        self.btn_start_enabled.setEnabled(can_start_enabled)
        self.btn_stop_all.setEnabled(can_stop_all)


    def _update_row_status(self, p: Project):
        if p.item:
            p.item.setText(2, self._status_label(p.status))
            p.item.setData(2, Qt.ItemDataRole.UserRole, p.status)
            self._apply_status_style(p)
            self.tree.viewport().update()
        if p.log:
            idx = self.tabs.indexOf(p.log)
            if idx >= 0:
                self.tabs.setTabText(idx, p.name)
        self._apply_tab_status(p)
        if not self._client_mode:
            write_state(self.projects)
        self._refresh_action_buttons()

    def _set_pending_action(self, p: Project, action: str) -> None:
        self._pending_actions[p.pid] = (action, time.monotonic())

    def _should_ignore_status_update(self, p: Project, new_status: str) -> bool:
        pending = self._pending_actions.get(p.pid)
        if not pending:
            return False
        action, ts = pending
        if (time.monotonic() - ts) > PENDING_STATUS_GRACE_SEC:
            self._pending_actions.pop(p.pid, None)
            return False
        status = (new_status or "").strip().lower()
        if action == "start":
            if status in ("stopped", "crashed", "stopping"):
                return True
            if status == "running":
                self._pending_actions.pop(p.pid, None)
        elif action == "stop":
            if status in ("running", "starting", "stopping", "waiting"):
                return True
            if status in ("stopped", "crashed"):
                self._pending_actions.pop(p.pid, None)
        elif action == "restart":
            if status == "running":
                return True
            # увидели переход в остановку/запуск — ждём финального running
            self._pending_actions[p.pid] = ("restart_wait", ts)
        elif action == "restart_wait":
            if status == "running":
                self._pending_actions.pop(p.pid, None)
        return False

    # ---------- кнопки ----------
    def on_add(self):
        dlg = ProjectDialog(
            self,
            lang=self._language,
            autostart_run=bool(self.cfg.data.get("autostart_run", False)),
            autostart_task=bool(self.cfg.data.get("autostart_task", False)),
        )
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            d = dlg.get_data()
            self._apply_app_autostart_settings(d)
            d.pop("app_autostart_run", None)
            d.pop("app_autostart_task", None)
            p = Project(pid=uuid.uuid4().hex, **d)
            self.projects.append(p)
            self.cfg.set_projects(self.projects)
            self._populate_projects()
            if self._client_mode:
                self._send_command("reload")

    def on_edit(self):
        p = self._selected_project()
        if not p:
            return
        init = p.to_dict()
        dlg = ProjectDialog(
            self,
            init=init,
            lang=self._language,
            autostart_run=bool(self.cfg.data.get("autostart_run", False)),
            autostart_task=bool(self.cfg.data.get("autostart_task", False)),
        )
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            d = dlg.get_data()
            self._apply_app_autostart_settings(d)
            d.pop("app_autostart_run", None)
            d.pop("app_autostart_task", None)
            p.name = d["name"]
            p.cmd = d["cmd"]
            p.cwd = d["cwd"]
            p.args = d.get("args", "")
            p.autorestart = d["autorestart"]
            p.enabled = d["enabled"]
            p.clear_log_on_start = bool(d.get("clear_log_on_start", False))
            if p.item is not None:
                p.item.setText(1, p.name)
                p.item.setText(3, p.cmd)
                p.item.setText(4, p.cwd)
            if p.log is not None:
                idx = self.tabs.indexOf(p.log)
                if idx >= 0:
                    self.tabs.setTabText(idx, p.name)
            if p.switch is not None:
                p.switch.blockSignals(True)
                p.switch.setChecked(p.enabled)
                p.switch.blockSignals(False)
            self.cfg.set_projects(self.projects)
            if self._client_mode:
                self._send_command("reload")

    def on_delete(self):
        p = self._selected_project()
        if not p:
            return
        if (p.process and p.process.state() == QtCore.QProcess.Running) or (
            self._client_mode and (p.status or "").strip().lower() in ("running", "starting")
        ):
            QtWidgets.QMessageBox.warning(
                self, APP_NAME, self._tr("msg_stop_running"))
            return
        self.projects = [x for x in self.projects if x.pid != p.pid]
        self.cfg.set_projects(self.projects)
        self._populate_projects()
        if self._client_mode:
            self._send_command("reload")

    def on_start_selected(self):
        p = self._selected_project()
        if not p:
            return
        if self._client_mode:
            if p.clear_log_on_start:
                self._clear_log_for_project(p)
            p.status = "starting"
            self._update_row_status(p)
            self._set_pending_action(p, "start")
            self._send_command("start", p.pid)
            return
        self.start_project(p)

    def on_stop_selected(self):
        p = self._selected_project()
        if not p:
            return
        if self._client_mode:
            p.status = "stopping"
            self._update_row_status(p)
            self._set_pending_action(p, "stop")
            self._send_command("stop", p.pid)
            return
        self.stop_project(p)

    def on_restart_selected(self):
        p = self._selected_project()
        if not p:
            return
        if self._client_mode:
            if p.clear_log_on_start:
                self._clear_log_for_project(p)
            cur_status = (p.status or "").strip().lower()
            if cur_status in ("running", "starting", "stopping", "waiting"):
                p.status = "stopping"
            else:
                p.status = "starting"
            self._update_row_status(p)
            self._set_pending_action(p, "restart")
            self._send_command("restart", p.pid)
            return
        self.stop_project(p)
        self.start_project(p)

    def on_start_enabled(self):
        if self._client_mode:
            for p in self.projects:
                if p.enabled and (p.status or "").strip().lower() in ("stopped", "crashed"):
                    if p.clear_log_on_start:
                        self._clear_log_for_project(p)
                    p.status = "starting"
                    self._update_row_status(p)
                    self._set_pending_action(p, "start")
            self._send_command("start_enabled")
            return
        targets = [p for p in self.projects if p.enabled and not (
            p.process and p.process.state() == QtCore.QProcess.Running)]
        if not targets:
            QtWidgets.QMessageBox.information(
                self, APP_NAME, self._tr("msg_no_enabled"))
            return
        self._start_enabled_queue = targets[:]
        self._start_enabled_next()

    def _start_enabled_next(self):
        if not self._start_enabled_queue:
            return
        p = self._start_enabled_queue.pop(0)
        self.start_project(p)
        QtCore.QTimer.singleShot(250, self._start_enabled_next)

    def on_stop_all(self):
        if self._client_mode:
            for p in self.projects:
                if (p.status or "").strip().lower() in ("running", "starting", "stopping"):
                    p.status = "stopping"
                    self._update_row_status(p)
                    self._set_pending_action(p, "stop")
            self._send_command("stop_all")
            return
        for p in self.projects:
            self.stop_project(p)

    def on_clear_log(self):
        idx = self.tabs.currentIndex()
        if idx < 0:
            return
        p = self._project_by_tab_index(idx)
        if p:
            self._clear_log_for_project(p)

    # ---------- автозапуск ----------
    def on_toggle_autostart(self, enabled: bool):
        self.cfg.data["autostart_run"] = enabled
        self.cfg.save()
        set_windows_run_autostart(enabled)

    def _set_autostart_task_checked(self, value: bool):
        if getattr(self, "act_autostart_task", None) is None:
            return
        self._syncing_autostart_task = True
        try:
            self.act_autostart_task.setChecked(value)
        finally:
            self._syncing_autostart_task = False

    def on_toggle_autostart_task(self, enabled: bool):
        if self._syncing_autostart_task:
            return
        if enabled:
            user = get_windows_username()
            if not user:
                QtWidgets.QMessageBox.warning(
                    self, APP_NAME, self._tr("msg_task_user_missing"))
                self._set_autostart_task_checked(False)
                return
            pwd, ok = QtWidgets.QInputDialog.getText(
                self,
                self._tr("msg_task_password_title"),
                self._tr("msg_task_password_label", user=user),
                QtWidgets.QLineEdit.Password,
            )
            if not ok or not pwd:
                QtWidgets.QMessageBox.information(
                    self, APP_NAME, self._tr("msg_task_password_empty"))
                self._set_autostart_task_checked(False)
                return
            ok_task, err = set_windows_task_autostart(
                True, username=user, password=pwd)
            if not ok_task:
                QtWidgets.QMessageBox.warning(
                    self, APP_NAME, self._tr("msg_task_enable_failed", err=err))
                self._set_autostart_task_checked(False)
                return
            self.cfg.data["autostart_task"] = True
            self.cfg.save()
        else:
            ok_task, err = set_windows_task_autostart(False)
            if not ok_task:
                QtWidgets.QMessageBox.warning(
                    self, APP_NAME, self._tr("msg_task_disable_failed", err=err))
                self._set_autostart_task_checked(True)
                return
            self.cfg.data["autostart_task"] = False
            self.cfg.save()

    # ---------- about ----------
    def on_about(self):
        text = self._tr(
            "about_text",
            app=f"{APP_NAME} v{APP_VERSION}",
            config=CONFIG_PATH,
            logs=LOGS_DIR,
        )
        QtWidgets.QMessageBox.information(self, APP_NAME, text)

    # ---------- запуск/стоп процессов ----------

    def start_project(self, p: Project):
        if p.process and p.process.state() == QtCore.QProcess.Running:
            return
        if not p.cmd:
            QtWidgets.QMessageBox.warning(
                self, APP_NAME, self._tr("msg_no_cmd"))
            return

        # санитарная очистка зомби перед запуском (бережно)
        # соберём PID-ы текущих запущенных проектов, чтобы их не трогать
        exclude = set()
        for q in self.projects:
            if q.process and q.process.state() == QtCore.QProcess.Running:
                try:
                    exclude.add(int(q.process.processId()))
                except Exception:
                    pass
        # если уже бежит проект с той же командой или той же рабочей папкой — зачистку пропускаем
        conflict_running = any((q is not p) and q.process and q.process.state() == QtCore.QProcess.Running and (
            q.cmd == p.cmd or (p.cwd and q.cwd and q.cwd == p.cwd)) for q in self.projects)
        if not conflict_running:
            _win_kill_project_zombies(p.cmd, p.cwd, exclude)
        if p.clear_log_on_start:
            self._clear_log_for_project(p)
        p.waiting_network = False
        program, args = program_and_args_for_cmd(p.cmd)
        # Добавим пользовательские параметры запуска
        extra = (p.args or '').strip()
        if extra:
            try:
                args += shlex.split(extra, posix=False)
            except Exception:
                args += extra.split()
        proc = QtCore.QProcess(self)
        env = QtCore.QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONIOENCODING", "utf-8")
        env.insert("COMBINER", "1")
        proc.setProcessEnvironment(env)
        proc.setProgram(program)
        proc.setArguments(args)
        if p.cwd:
            proc.setWorkingDirectory(p.cwd)
        proc.setProcessChannelMode(QtCore.QProcess.MergedChannels)

        proc.readyReadStandardOutput.connect(
            lambda p_=p, pr=proc: self._on_proc_output(p_, pr))
        proc.readyReadStandardError.connect(
            lambda p_=p, pr=proc: self._on_proc_output(p_, pr))
        proc.finished.connect(lambda code, status,
                              p_=p, pr=proc: self._on_proc_finished(p_, pr, code, status))
        proc.errorOccurred.connect(
            lambda err, p_=p, pr=proc: self._on_proc_error(p_, pr, err))
        proc.started.connect(lambda p_=p: self._on_proc_started(p_))

        p.stopping = False
        p.status = "starting"
        self._update_row_status(p)
        try:
            proc.start()
        except Exception as e:
            if p.log:
                p.log.append_text(self._tr("log_start_error", err=e) + "\n")
            append_project_log(p, self._tr("log_start_error", err=e) + "\n")
            p.status = "stopped"
            self._update_row_status(p)
            return

        p.process = proc
        if p.log:
            p.log.append_text(
                f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
                f"{self._tr('log_start', cmd=p.cmd)}\n")
        append_project_log(
            p,
            f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
            f"{self._tr('log_start', cmd=p.cmd)}\n",
        )

    def _on_proc_started(self, p: Project):
        p.status = "running"
        self._update_row_status(p)

    def stop_project(self, p: Project):
        # мягко → жёстко → зачистка зомби
        p.waiting_network = False
        if p.process and p.process.state() == QtCore.QProcess.Running:
            p.stopping = True
            p.status = "stopping"
            self._update_row_status(p)
            pid = int(p.process.processId() or 0)
            try:
                p.process.terminate()
                if not p.process.waitForFinished(1500):
                    _win_taskkill_tree(pid)
            except Exception:
                _win_taskkill_tree(pid)
        else:
            _win_kill_project_zombies(p.cmd, p.cwd)
            p.status = "stopped"
            self._update_row_status(p)
        if p.log:
            p.log.append_text(
                f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
                f"{self._tr('log_stop')}\n"
            )
        append_project_log(
            p,
            f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
            f"{self._tr('log_stop')}\n",
        )

    def _on_proc_output(self, p: Project, pr: QtCore.QProcess):
        data = pr.readAll().data()
        text = decode_bytes(data)
        if p.log and text:
            p.log.append_text(text)
        if text:
            append_project_log(p, text)

    def _on_proc_finished(self, p: Project, pr: QtCore.QProcess, code: int, status: QtCore.QProcess.ExitStatus):
        if p.log:
            p.log.append_text(
                f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
                f"{self._tr('log_finish', code=code, status=('CrashExit' if status==QtCore.QProcess.CrashExit else 'NormalExit'))}\n"
            )
        append_project_log(
            p,
            f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
            f"{self._tr('log_finish', code=code, status=('CrashExit' if status==QtCore.QProcess.CrashExit else 'NormalExit'))}\n",
        )
        if p.process is not pr:
            return
        was_stopping = p.stopping
        p.status = "stopped" if (was_stopping or (status == QtCore.QProcess.NormalExit and code == 0)) else "crashed"
        self._update_row_status(p)

        pr_autorestart = (
            p.autorestart and not was_stopping and status == QtCore.QProcess.CrashExit)
        p.process = None
        if p.stopping:
            p.stopping = False

        if pr_autorestart:
            QtCore.QTimer.singleShot(2000, lambda: self.start_project(p))

    def _on_proc_error(self, p: Project, pr: QtCore.QProcess, err: QtCore.QProcess.ProcessError):
        if p.process is not pr:
            return
        if p.log:
            p.log.append_text(self._tr("log_proc_error", err=err) + "\n")
        append_project_log(p, self._tr("log_proc_error", err=err) + "\n")

    def _on_selection_changed(self):
        if self._syncing_selection:
            return
        it = self.tree.currentItem()
        if not it:
            return
        pid = it.data(0, Qt.ItemDataRole.UserRole)
        # активировать вкладку по имени
        for idx, prj in enumerate(self.projects):
            if prj.pid == pid:
                self._syncing_selection = True
                try:
                    self.tabs.setCurrentIndex(idx)
                finally:
                    self._syncing_selection = False
                break
        self._refresh_action_buttons()

    def _on_tab_changed(self, index: int):
        if self._syncing_selection:
            return
        if index < 0 or index >= len(self.projects):
            return
        prj = self.projects[index]
        if not prj.item:
            return
        self._syncing_selection = True
        try:
            self.tree.setCurrentItem(prj.item)
        finally:
            self._syncing_selection = False
        if self._client_mode:
            self._load_log_tail(prj)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        # Реальный выход по крестику
        try:
            self.cfg.set_projects(self.projects)
        except Exception:
            traceback.print_exc()
        if not self._client_mode:
            # Остановка всех процессов
            for p in self.projects:
                try:
                    self.stop_project(p)
                except Exception:
                    pass
            QtWidgets.QApplication.processEvents(QtCore.QEventLoop.AllEvents, 50)
            for p in self.projects:
                try:
                    if p.process and p.process.state() == QtCore.QProcess.Running:
                        pid = int(p.process.processId() or 0)
                        _win_taskkill_tree(pid)
                        p.process.waitForFinished(800)
                    p.log = None
                except Exception:
                    pass
        super().closeEvent(event)

    def showEvent(self, e: QtGui.QShowEvent) -> None:
        super().showEvent(e)
        if not getattr(self, "_centered_once", False):
            self._centered_once = True
            screen = self.screen() or QtWidgets.QApplication.primaryScreen()
            geo = self.frameGeometry()
            geo.moveCenter(screen.availableGeometry().center())
            self.move(geo.topLeft())

    def _build_tray(self):
        self.tray = QtWidgets.QSystemTrayIcon(self)
        icon = self.windowIcon() or load_app_icon()
        self.tray.setIcon(icon)
        menu = QtWidgets.QMenu()
        self.act_tray_show = menu.addAction("")
        self.act_tray_show.triggered.connect(
            lambda: (self.showNormal(), self.raise_(), self.activateWindow()))
        menu.addSeparator()
        self.act_tray_start_enabled = menu.addAction("")
        self.act_tray_start_enabled.triggered.connect(self.on_start_enabled)
        menu.addSeparator()
        self.act_tray_exit = menu.addAction("")
        self.act_tray_exit.triggered.connect(self._quit_from_tray)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self._on_tray_activated)
        self.tray.show()

    def _on_tray_activated(self, reason):
        if reason == QtWidgets.QSystemTrayIcon.Trigger:
            self.showNormal()
            self.raise_()
            self.activateWindow()

    def _quit_from_tray(self):
        self._really_quit = True
        self.close()

    def changeEvent(self, event: QtCore.QEvent) -> None:
        # Сворачивать в трей только при нажатии кнопки «Свернуть»
        if event.type() == QtCore.QEvent.WindowStateChange:
            if self.isMinimized() and QtWidgets.QSystemTrayIcon.isSystemTrayAvailable():
                QtCore.QTimer.singleShot(0, self.hide)
                try:
                    self.tray.showMessage(
                        APP_NAME, self._tr("tray_minimized"), QtWidgets.QSystemTrayIcon.Information, 1200)
                except Exception:
                    pass
        super().changeEvent(event)

# ------------------------ Диалог проекта -----------------------------------


class ProjectDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent=None,
        init: Optional[dict] = None,
        lang: str = Lang.System,
        autostart_run: Optional[bool] = None,
        autostart_task: Optional[bool] = None,
    ):
        super().__init__(parent)
        self._lang = resolve_lang(lang)
        self._tr = lambda key, **kwargs: tr(self._lang, key, **kwargs)
        self.setWindowTitle(self._tr("dlg_title"))
        self.setModal(True)
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.setSpacing(8)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        lay.addLayout(form)

        self.ed_name = QtWidgets.QLineEdit()
        self.ed_cmd = QtWidgets.QLineEdit()
        self.ed_cwd = QtWidgets.QLineEdit()
        self.ed_args = QtWidgets.QLineEdit()
        self.ed_args.setPlaceholderText(self._tr("dlg_placeholder_args"))

        # авто-подстановка CWD/Имени из выбранного файла
        self.ed_cmd.editingFinished.connect(self._autofill_from_cmd)

        btn_cmd = QtWidgets.QPushButton(self._tr("dlg_browse"))
        btn_cwd = QtWidgets.QPushButton(self._tr("dlg_browse"))

        row = QtWidgets.QHBoxLayout()
        row.addWidget(self.ed_cmd, 1)
        row.addWidget(btn_cmd)
        form.addRow(self._tr("dlg_label_cmd"), row)

        row2 = QtWidgets.QHBoxLayout()
        row2.addWidget(self.ed_cwd, 1)
        row2.addWidget(btn_cwd)
        form.addRow(self._tr("dlg_label_cwd"), row2)
        form.addRow(self._tr("dlg_label_args"), self.ed_args)

        form.addRow(self._tr("dlg_label_name"), self.ed_name)

        self.chk_enabled = QtWidgets.QCheckBox(self._tr("dlg_chk_enabled"))
        self.chk_autorst = QtWidgets.QCheckBox(self._tr("dlg_chk_autorst"))
        self.chk_autorst.setChecked(True)
        lay.addWidget(self.chk_enabled)
        lay.addWidget(self.chk_autorst)
        self.chk_clear_log = QtWidgets.QCheckBox(self._tr("dlg_chk_clear_log"))
        lay.addWidget(self.chk_clear_log)

        self.grp_autostart = QtWidgets.QGroupBox(
            self._tr("dlg_app_autostart_group"))
        autol = QtWidgets.QVBoxLayout(self.grp_autostart)
        autol.setContentsMargins(10, 6, 10, 6)
        autol.setSpacing(4)
        self.chk_app_autostart_run = QtWidgets.QCheckBox(
            self._tr("dlg_app_autostart_run"))
        self.chk_app_autostart_task = QtWidgets.QCheckBox(
            self._tr("dlg_app_autostart_task"))
        self.chk_app_autostart_task.setToolTip(
            self._tr("dlg_app_autostart_task_tip"))
        autol.addWidget(self.chk_app_autostart_run)
        autol.addWidget(self.chk_app_autostart_task)
        lay.addWidget(self.grp_autostart)

        bb = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        lay.addWidget(bb)

        btn_cmd.clicked.connect(self._pick_cmd)
        btn_cwd.clicked.connect(self._pick_cwd)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)

        if init:
            self.ed_name.setText(init.get("name", ""))
            self.ed_cmd.setText(init.get("cmd", ""))
            self.ed_cwd.setText(init.get("cwd", ""))
            self.ed_args.setText(init.get("args", ""))
            self.chk_enabled.setChecked(bool(init.get("enabled", False)))
            self.chk_autorst.setChecked(bool(init.get("autorestart", True)))
            self.chk_clear_log.setChecked(bool(init.get("clear_log_on_start", False)))
        if autostart_run is not None:
            self.chk_app_autostart_run.setChecked(bool(autostart_run))
        if autostart_task is not None:
            self.chk_app_autostart_task.setChecked(bool(autostart_task))

    def _autofill_from_cmd(self):
        path_str = self.ed_cmd.text().strip()
        if not path_str:
            return
        try:
            p = Path(path_str)
            # Рабочая папка = папка файла, если поле пустое
            if not self.ed_cwd.text().strip():
                self.ed_cwd.setText(str(p.parent))
            # Имя проекта = имя родительской папки (если пусто), иначе имя файла без расширения
            if not self.ed_name.text().strip():
                self.ed_name.setText(p.parent.name or p.stem)
        except Exception:
            pass

    def _pick_cmd(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, self._tr("dlg_pick_cmd"), "",
            self._tr("dlg_filter_cmd")
        )
        if path:
            self.ed_cmd.setText(path)
            self._autofill_from_cmd()

    def _pick_cwd(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(
            self, self._tr("dlg_pick_cwd"), "")
        if d:
            self.ed_cwd.setText(d)

    def get_data(self) -> dict:
        return {
            "name": self.ed_name.text().strip() or self._tr("dlg_default_name"),
            "cmd": self.ed_cmd.text().strip(),
            "cwd": self.ed_cwd.text().strip(),
            "args": self.ed_args.text().strip(),
            "enabled": self.chk_enabled.isChecked(),
            "autorestart": self.chk_autorst.isChecked(),
            "clear_log_on_start": self.chk_clear_log.isChecked(),
            "app_autostart_run": self.chk_app_autostart_run.isChecked(),
            "app_autostart_task": self.chk_app_autostart_task.isChecked(),
        }


# ------------------------ main ---------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--autostart", action="store_true",
                        help="Запустить включённые проекты автоматически.")
    parser.add_argument("--headless", action="store_true",
                        help="Запуск без интерфейса (фоновый контроллер).")
    parser.add_argument("--data-dir", default="",
                        help="Переопределить папку данных PyCombiner.")
    args = parser.parse_args()

    if args.data_dir:
        set_data_dir(args.data_dir)
    ensure_dirs()
    install_debug_handlers()

    cfg = Config(CONFIG_PATH)
    if not args.headless:
        maybe_update_env_snapshot(cfg)
    if bool(cfg.data.get("autostart_run", False)):
        set_windows_run_autostart(True)
    if args.headless:
        if is_daemon_running():
            return
        app = QtCore.QCoreApplication(sys.argv)
        log_app(
            f"Headless controller start autostart={args.autostart} data_dir={APPDATA_DIR}"
        )
        controller = HeadlessController(cfg, autostart=args.autostart)
        # keep reference, otherwise GC may stop timers in headless mode
        app._headless_controller = controller  # type: ignore[attr-defined]
        sys.exit(app.exec())

    app = QtWidgets.QApplication(sys.argv)
    app.setOrganizationName(ORG_NAME)
    app.setApplicationName(APP_NAME)
    try:
        ic = load_app_icon()
        if ic:
            app.setWindowIcon(ic)
    except Exception:
        pass

    win = MainWindow(cfg)
    win.resize(1200, 800)
    win.show()
    log_app("GUI started")

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
