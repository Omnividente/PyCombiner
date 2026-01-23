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
import json
import locale
import os
import platform
import re
import shlex
import subprocess
import sys
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
APP_VERSION = "1.1"
ORG_NAME = "PyCombiner"
AUTOSTART_TASK_NAME = "PyCombiner Startup"

APPDATA_DIR = Path(os.environ.get("APPDATA", str(
    Path.home() / "AppData" / "Roaming"))) / "PyCombiner"
CONFIG_PATH = APPDATA_DIR / "config.json"
LOGS_DIR = APPDATA_DIR / "logs"
APPDATA_DIR.mkdir(parents=True, exist_ok=True)
LOGS_DIR.mkdir(parents=True, exist_ok=True)
LOG_MAX_LINES = 300


# ------------------------ Утилиты ------------------------------------------

def ensure_dirs():
    APPDATA_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)


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
        "log_start_error": "[!] Ошибка запуска: {err}",
        "log_proc_error": "[!] Ошибка процесса: {err}",
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
        "log_start_error": "[!] Start error: {err}",
        "log_proc_error": "[!] Process error: {err}",
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
    },
}

STATUS_LABELS = {
    Lang.RU: {
        "running": "работает",
        "starting": "запуск",
        "stopped": "остановлен",
        "crashed": "ошибка",
    },
    Lang.EN: {
        "running": "running",
        "starting": "starting",
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


def _win_kill_project_zombies(command: str, work_dir: str, exclude_pids: Optional[set[int]] = None):
    """
    Перед стартом пробуем найти и погасить зависшие процессы проекта
    по подстрокам из командной строки/рабочей папки.
    Делается через PowerShell и запрос CIM Win32_Process.
    """
    if platform.system() != "Windows":
        return

    needles = []
    if command:
        needles.append(command)
        needles.append(os.path.basename(command))
    if work_dir:
        needles.append(work_dir)
    needles = [n for n in needles if n]

    if not needles:
        return

    # В PowerShell оборачиваем шаблон в ОДИНАРНЫЕ кавычки: '*needle*'
    # Если внутри есть одинарная кавычка — удваиваем её.
    cond_parts = []
    for n in needles:
        esc = n.replace("'", "''")
        cond_parts.append(f"($_.CommandLine -like '*{esc}*')")
    cond = " -and ".join(cond_parts)

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

    for m in re.findall(r"\d+", out or ""):
        try:
            pid = int(m)
        except Exception:
            continue
        if exclude_pids and pid in exclude_pids:
            continue
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

    def to_dict(self) -> dict:
        return {
            "pid": self.pid,
            "name": self.name,
            "cmd": self.cmd,
            "cwd": self.cwd, "args": self.args,
            "enabled": self.enabled,
            "autorestart": self.autorestart,
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
            cmd = get_self_executable_for_run()
            args = [
                "schtasks",
                "/Create",
                "/F",
                "/TN",
                AUTOSTART_TASK_NAME,
                "/SC",
                "ONSTART",
                "/RL",
                "HIGHEST",
                "/RU",
                username,
                "/RP",
                password,
                "/TR",
                cmd,
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
        return ok, msg.strip()
    except Exception as e:
        return False, str(e)


def get_self_executable_for_run() -> str:
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

    if os.name == "nt":
        # правильное quoting для командной строки Windows
        return subprocess.list2cmdline(args)
    else:
        # безопасное quoting для POSIX
        return " ".join(shlex.quote(a) for a in args)


def shutil_which(name: str) -> Optional[str]:
    for p in os.environ.get("PATH", "").split(os.pathsep):
        candidate = Path(p) / name
        if candidate.exists():
            return str(candidate)
    return None


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

        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setWindowIcon(load_app_icon())

        self._really_quit = False
        self._build_tray()
        self._build_ui()
        self._populate_projects()
        self.apply_language()
        self.apply_theme()
        QtCore.QTimer.singleShot(0, self.apply_theme)

        QtCore.QTimer.singleShot(300, self.on_start_enabled)
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

        if self.tree.topLevelItemCount() > 0:
            self.tree.setCurrentItem(self.tree.topLevelItem(0))
            self.tabs.setCurrentIndex(0)

    def _on_switch_toggled(self, p: Project, checked: bool) -> None:
        p.enabled = checked
        self.cfg.set_projects(self.projects)

    def _selected_project(self) -> Optional[Project]:
        it = self.tree.currentItem()
        if not it:
            return None
        pid = it.data(0, Qt.ItemDataRole.UserRole)
        for p in self.projects:
            if p.pid == pid:
                return p
        return None

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

    def on_delete(self):
        p = self._selected_project()
        if not p:
            return
        if p.process and p.process.state() == QtCore.QProcess.Running:
            QtWidgets.QMessageBox.warning(
                self, APP_NAME, self._tr("msg_stop_running"))
            return
        self.projects = [x for x in self.projects if x.pid != p.pid]
        self.cfg.set_projects(self.projects)
        self._populate_projects()

    def on_start_selected(self):
        p = self._selected_project()
        if p:
            self.start_project(p)

    def on_stop_selected(self):
        p = self._selected_project()
        if p:
            self.stop_project(p)

    def on_restart_selected(self):
        p = self._selected_project()
        if not p:
            return
        self.stop_project(p)
        self.start_project(p)

    def on_start_enabled(self):
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
        for p in self.projects:
            self.stop_project(p)

    def on_clear_log(self):
        idx = self.tabs.currentIndex()
        if idx < 0:
            return
        widget = self.tabs.widget(idx)
        if isinstance(widget, LogView):
            widget.clear()

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
    def _program_and_args_for_cmd(self, cmd: str) -> Tuple[str, List[str]]:
        path = Path(cmd.strip().strip('"'))
        suf = path.suffix.lower()
        if suf == ".ps1":
            prog = "pwsh.exe" if shutil_which(
                "pwsh.exe") or shutil_which("pwsh") else "powershell.exe"
            return prog, ["-NoLogo", "-ExecutionPolicy", "Bypass", "-File", str(path)]
        if suf in (".bat", ".cmd"):
            return "cmd.exe", ["/c", str(path)]
        if suf == ".exe":
            return str(path), []
        if suf == ".py":
            # запуск .py: в сборке sys.executable указывает на PyCombiner.exe,
            # поэтому ищем интерпретатор рядом со скриптом или используем системный.
            # 1) локальный venv рядом со скриптом
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
            # 2) если не frozen — можно использовать текущий интерпретатор
            if not getattr(sys, "frozen", False):
                return sys.executable, ["-u", str(path)]
            # 3) запасной вариант: системный python
            if os.name == "nt":
                return "py", ["-3", "-u", str(path)]
            return "python3", ["-u", str(path)]
        parts = shlex.split(cmd, posix=False)
        return (parts[0], parts[1:]) if parts else (cmd, [])

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
        program, args = self._program_and_args_for_cmd(p.cmd)
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
        env = QtCore.QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONIOENCODING", "utf-8")
        env.insert("COMBINER", "1")
        env.insert("PYTHONIOENCODING", "utf-8")
        proc.setProcessEnvironment(env)

        proc.readyReadStandardOutput.connect(
            lambda p_=p, pr=proc: self._on_proc_output(p_, pr))
        proc.readyReadStandardError.connect(
            lambda p_=p, pr=proc: self._on_proc_output(p_, pr))
        proc.finished.connect(lambda code, status,
                              p_=p: self._on_proc_finished(p_, code, status))
        proc.errorOccurred.connect(
            lambda err, p_=p: self._on_proc_error(p_, err))
        proc.started.connect(lambda p_=p: self._on_proc_started(p_))

        p.stopping = False
        p.status = "starting"
        self._update_row_status(p)
        try:
            proc.start()
        except Exception as e:
            if p.log:
                p.log.append_text(self._tr("log_start_error", err=e) + "\n")
            p.status = "stopped"
            self._update_row_status(p)
            return

        p.process = proc
        if p.log:
            p.log.append_text(
                f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
                f"{self._tr('log_start', cmd=p.cmd)}\n")

    def _on_proc_started(self, p: Project):
        p.status = "running"
        self._update_row_status(p)

    def stop_project(self, p: Project):
        # мягко → жёстко → зачистка зомби
        if p.process and p.process.state() == QtCore.QProcess.Running:
            p.stopping = True
            pid = int(p.process.processId() or 0)
            try:
                p.process.terminate()
                if not p.process.waitForFinished(1500):
                    _win_taskkill_tree(pid)
            except Exception:
                _win_taskkill_tree(pid)
            finally:
                p.process = None
        else:
            _win_kill_project_zombies(p.cmd, p.cwd)

        p.status = "stopped"
        self._update_row_status(p)
        if p.log:
            p.log.append_text(
                f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
                f"{self._tr('log_stop')}\n"
            )

    def _on_proc_output(self, p: Project, pr: QtCore.QProcess):
        data = pr.readAll().data()
        text = decode_bytes(data)
        if p.log and text:
            p.log.append_text(text)

    def _on_proc_finished(self, p: Project, code: int, status: QtCore.QProcess.ExitStatus):
        if p.log:
            p.log.append_text(
                f"[{QtCore.QDateTime.currentDateTime().toString('yyyy-MM-dd hh:mm:ss')}] "
                f"{self._tr('log_finish', code=code, status=('CrashExit' if status==QtCore.QProcess.CrashExit else 'NormalExit'))}\n"
            )
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

    def _on_proc_error(self, p: Project, err: QtCore.QProcess.ProcessError):
        if p.log:
            p.log.append_text(self._tr("log_proc_error", err=err) + "\n")

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

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        # Реальный выход по крестику
        try:
            self.cfg.set_projects(self.projects)
        except Exception:
            traceback.print_exc()
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
            "app_autostart_run": self.chk_app_autostart_run.isChecked(),
            "app_autostart_task": self.chk_app_autostart_task.isChecked(),
        }


# ------------------------ main ---------------------------------------------

def main():
    ensure_dirs()
    parser = argparse.ArgumentParser()
    parser.add_argument("--autostart", action="store_true",
                        help="Запустить включённые проекты автоматически.")
    args = parser.parse_args()

    app = QtWidgets.QApplication(sys.argv)
    app.setOrganizationName(ORG_NAME)
    app.setApplicationName(APP_NAME)
    try:
        ic = load_app_icon()
        if ic:
            app.setWindowIcon(ic)
    except Exception:
        pass

    cfg = Config(CONFIG_PATH)
    win = MainWindow(cfg)
    win.resize(1200, 800)
    win.show()

    if args.autostart:
        QtCore.QTimer.singleShot(500, win.on_start_enabled)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
