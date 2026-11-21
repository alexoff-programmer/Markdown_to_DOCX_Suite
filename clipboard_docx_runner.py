#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown to DOCX Suite — Windows 11 Edition
ClipboardDocxRunner & Batch Converter
Полное соответствие Fluent Design System 2.0 (Mica/Acrylic, Segoe UI Variable, Dark Theme):
- Главное окно со вкладками: Пакетный конвертер, Буфер обмена, Библиотека шаблонов, Параметры
- Hero Drag & Drop зона для файлов (.md, .txt, .html) и целых папок с рекурсивным поиском
- Очередь задач со статусами-пиллами (Готово, В обработке, Ожидание, Ошибка)
- Неблокирующий многопоточный воркер QThread и анимированный прогресс-бар
- Модальное плавающее окно конвертации буфера с карточками шаблонов (1, 2, 3)
- Windows 11 Toast-уведомления с кнопками «Открыть файл» и «Папка»
- Системный трей и глобальные хоткеи (Ctrl+Alt+W, Ctrl+Alt+B)
- Полная интеграция с ядром ГОСТ (gost_docx_parser и gost_docx_formatter)
"""

from __future__ import annotations

import os
import gc
import copy
import sys
import json
import time
import re
import shutil
import subprocess
import argparse
from pathlib import Path
from typing import Tuple, Optional, Dict, List, Any

# Импорт ядра ГОСТ
# Поддержка работы как из исходников, так и внутри скомпилированного PyInstaller EXE
meipass = getattr(sys, "_MEIPASS", None)
if meipass:
    sys.path.insert(0, str(Path(meipass) / "gost_engine"))
    sys.path.insert(0, str(meipass))

ENGINE_DIR = Path(__file__).resolve().parent.parent / "gost_engine"
if not ENGINE_DIR.exists():
    ENGINE_DIR = Path(__file__).resolve().parent / "gost_engine"
if not ENGINE_DIR.exists():
    ENGINE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ENGINE_DIR))

try:
    from gost_docx_parser import MarkdownASTParser, tokenize_inlines, typography_cleanup, NBSP
    from gost_docx_formatter import (
        GostDocxFormatter, latex_to_omml,
        append_markdown_to_docx, merge_docx_files, update_docx_styles
    )
    HAS_ENGINE = True
except ImportError:
    HAS_ENGINE = False
    MarkdownASTParser = None
    GostDocxFormatter = None

try:
    import pyperclip
except ImportError:
    pyperclip = None

try:
    import keyboard
except ImportError:
    keyboard = None

if sys.platform == "win32":
    import winreg
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
else:
    winreg = None
    user32 = None
    kernel32 = None

# PySide6 GUI
try:
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QSystemTrayIcon, QMenu, QDialog, QWidget,
        QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QCheckBox, QRadioButton,
        QPushButton, QButtonGroup, QFrame, QGraphicsDropShadowEffect, QGraphicsOpacityEffect,
        QFileDialog, QScrollArea, QComboBox, QTableWidget, QTableWidgetItem, QHeaderView,
        QProgressBar, QStackedWidget, QAbstractItemView, QTextEdit, QPlainTextEdit, QSplitter,
        QListWidget, QSizePolicy, QSpinBox, QDoubleSpinBox, QColorDialog, QSizeGrip
    )
    from template_editor_widget import TemplateEditorWidget
    from PySide6.QtGui import (
        QIcon, QPixmap, QPainter, QColor, QFont, QCursor, QAction,
        QDragEnterEvent, QDropEvent
    )
    try:
        from PySide6.QtSvg import QSvgRenderer
        HAS_QTSVG = True
    except ImportError:
        QSvgRenderer = None
        HAS_QTSVG = False
    from PySide6.QtCore import (
        Qt, QObject, QAbstractNativeEventFilter, QPoint, QRect, Signal, Slot,
        QThread, QTimer, QSize, QPropertyAnimation, QEasingCurve, QFileSystemWatcher
    )
    HAS_PYSIDE6 = True
except ImportError:
    HAS_PYSIDE6 = False

# ==============================================================================
# 1. КОНСТАНТЫ И ПУТИ КОНФИГУРАЦИИ
# ==============================================================================

APP_NAME = "ClipboardDocxRunner"
APP_DATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / APP_NAME
APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = APP_DATA_DIR / "config.json"
TEMPLATES_REGISTRY_FILE = APP_DATA_DIR / "custom_templates_registry.json"
LOG_FILE = APP_DATA_DIR / "docx_runner.log"

DEFAULT_CONFIG = {
    "hotkey": "ctrl+alt+w",
    "batch_hotkey": "ctrl+alt+b",
    "default_template": "dstu_2026_gost_russian_names",
    "open_after_create": True,
    "autostart": False,
    "two_columns": False,
    "strip_body_emphasis": False,
    "merge_apply_template": False,
    "save_mode": "same_dir",
    "output_dir": str((Path.home() / "Desktop" / "Word_Documents").resolve())
}

AUTOSTART_REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTOSTART_VALUE_NAME = "MarkdownToDocxSuite"

def get_windows_autostart_command() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --minimized'
    py_exe = Path(sys.executable)
    pyw = py_exe.with_name("pythonw.exe")
    runner_py = Path(__file__).resolve()
    exe_str = str(pyw if pyw.exists() else py_exe)
    return f'"{exe_str}" "{runner_py}" --minimized'

def is_windows_autostart_enabled() -> bool:
    if sys.platform != "win32" or winreg is None:
        return bool(load_config().get("autostart", False))
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0, winreg.KEY_READ) as key:
            val, _ = winreg.QueryValueEx(key, AUTOSTART_VALUE_NAME)
            return bool(val)
    except FileNotFoundError:
        return False
    except Exception:
        return bool(load_config().get("autostart", False))

def set_windows_autostart(enabled: bool) -> Tuple[bool, str]:
    if sys.platform != "win32" or winreg is None:
        return True, ""
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, AUTOSTART_REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, AUTOSTART_VALUE_NAME, 0, winreg.REG_SZ, get_windows_autostart_command())
            else:
                try:
                    winreg.DeleteValue(key, AUTOSTART_VALUE_NAME)
                except FileNotFoundError:
                    pass
        return True, ""
    except Exception as e:
        log(f"Ошибка настройки автозагрузки Windows: {e}")
        return False, str(e)

def log(msg: str):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] {msg}\n"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(entry)
    except Exception:
        pass
    if sys.stdout and not getattr(sys, "frozen", False):
        try:
            print(entry.strip())
        except Exception:
            pass

def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return {**DEFAULT_CONFIG, **json.load(f)}
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

def get_templates_dirs() -> List[Path]:
    """Возвращает список директорий поиска шаблонов в порядке приоритета:
    1) Пользовательские папки (рядом с EXE/скриптом, рабочая папка, APPDATA) — для мгновенного подхвата новых шаблонов;
    2) Внутренняя папка _MEIPASS (встроенные шаблоны при запуске из одиночного EXE).
    """
    dirs = [
        get_base_dir() / "templates",
        Path.cwd() / "templates",
        get_base_dir().parent / "templates",
        APP_DATA_DIR / "templates"
    ]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        dirs.append(Path(meipass) / "templates")

    # Гарантируем наличие локальной папки templates
    primary_dir = get_base_dir() / "templates"
    try:
        primary_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    unique = []
    seen = set()
    for d in dirs:
        try:
            r = d.resolve()
            if r not in seen and r.exists() and r.is_dir():
                seen.add(r)
                unique.append(r)
        except Exception:
            pass
    return unique

def get_primary_templates_dir() -> Path:
    """Основная папка, куда импортируются или сохраняются пользовательские шаблоны."""
    d = get_base_dir() / "templates"
    try:
        d.mkdir(parents=True, exist_ok=True)
        return d.resolve()
    except Exception:
        fallback = APP_DATA_DIR / "templates"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback.resolve()

def register_custom_template_path(file_path: Path):
    """Регистрирует путь к импортированному или новому шаблону в реестре и в .templates_index.json."""
    try:
        resolved = str(file_path.resolve())
    except Exception:
        resolved = str(file_path)

    # 1. Сохраняем в APPDATA реестр
    reg_list = []
    if TEMPLATES_REGISTRY_FILE.exists():
        try:
            with open(TEMPLATES_REGISTRY_FILE, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
                if isinstance(data, list):
                    reg_list = data
        except Exception:
            pass
    if resolved not in reg_list:
        reg_list.append(resolved)
        try:
            with open(TEMPLATES_REGISTRY_FILE, "w", encoding="utf-8") as f:
                json.dump(reg_list, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # 2. Сохраняем в локальный индекс папки шаблона
    try:
        idx_file = file_path.parent / ".templates_index.json"
        idx_list = []
        if idx_file.exists():
            with open(idx_file, "r", encoding="utf-8-sig") as f:
                d = json.load(f)
                if isinstance(d, list):
                    idx_list = d
        if file_path.name not in idx_list:
            idx_list.append(file_path.name)
            with open(idx_file, "w", encoding="utf-8") as f:
                json.dump(idx_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def unregister_template_path(file_path: Path):
    """Удаляет путь шаблона из реестра и .templates_index.json при удалении шаблона."""
    try:
        resolved = str(file_path.resolve())
    except Exception:
        resolved = str(file_path)
    if TEMPLATES_REGISTRY_FILE.exists():
        try:
            with open(TEMPLATES_REGISTRY_FILE, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if isinstance(data, list) and resolved in data:
                data = [x for x in data if x != resolved]
                with open(TEMPLATES_REGISTRY_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    try:
        idx_file = file_path.parent / ".templates_index.json"
        if idx_file.exists():
            with open(idx_file, "r", encoding="utf-8-sig") as f:
                d = json.load(f)
            if isinstance(d, list) and file_path.name in d:
                d = [x for x in d if x != file_path.name]
                with open(idx_file, "w", encoding="utf-8") as f:
                    json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def get_registered_template_paths() -> List[Path]:
    paths = []
    if TEMPLATES_REGISTRY_FILE.exists():
        try:
            with open(TEMPLATES_REGISTRY_FILE, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
                if isinstance(data, list):
                    for p_str in data:
                        p = Path(p_str)
                        if p.exists() and p.is_file():
                            paths.append(p)
        except Exception:
            pass
    return paths

_TEMPLATES_LIST_CACHE: Dict[str, Any] = {"sig": None, "items": []}

def list_available_templates() -> List[dict]:
    """Сканирует папки шаблонов с мгновенным кэшированием по сигнатуре файлов (путь, mtime, размер)."""
    cur_sig = get_templates_signature()
    if _TEMPLATES_LIST_CACHE["sig"] == cur_sig and _TEMPLATES_LIST_CACHE["items"]:
        return list(_TEMPLATES_LIST_CACHE["items"])
    res = {}
    known_template_filenames = [
        "dstu_2026_gost_russian_names.json",
        "dstu_2026_gost_v3.json",
        "vak_rinc_article.json",
        "gost_34_tz.json",
        "corporate.json",
        "default.json",
        "ieee_vak_two_column.json"
    ]
    for d in get_templates_dirs():
        cand_files = []
        # 1. Полное динамическое чтение директории через os.listdir
        try:
            for entry in sorted(os.listdir(d)):
                if entry.lower().endswith(".json"):
                    fp = d / entry
                    if fp.is_file() and fp not in cand_files:
                        cand_files.append(fp)
        except Exception:
            pass

        # 2. Страховочная проверка известных имен
        for kn in known_template_filenames:
            kp = d / kn
            if kp.exists() and kp.is_file() and kp not in cand_files:
                cand_files.append(kp)

        # 3. Дополнительный обход через glob
        try:
            for gf in sorted(d.glob("*.json")):
                if gf.name != ".templates_index.json" and gf not in cand_files:
                    cand_files.append(gf)
        except Exception:
            pass

        # 4. Чтение локального индекса .templates_index.json
        idx_file = d / ".templates_index.json"
        if idx_file.exists():
            try:
                with open(idx_file, "r", encoding="utf-8-sig") as f:
                    idx_names = json.load(f)
                if isinstance(idx_names, list):
                    for fn in idx_names:
                        ip = d / fn
                        if ip.exists() and ip.is_file() and ip not in cand_files:
                            cand_files.append(ip)
            except Exception:
                pass

        # 5. Зарегистрированные пользовательские пути
        for rp in get_registered_template_paths():
            if rp not in cand_files:
                cand_files.append(rp)

        for f in cand_files:
            try:
                # utf-8-sig безопасно читает файлы, сохраненные в Блокноте Windows с BOM
                with open(f, "r", encoding="utf-8-sig") as jf:
                    data = json.load(jf)
                if isinstance(data, dict) and "styles" in data and isinstance(data["styles"], dict):
                    t_name = str(data.get("name") or f.stem).strip()
                    data["name"] = t_name
                    if not data.get("display_title"):
                        data["display_title"] = t_name
                    if not data.get("description"):
                        data["description"] = f"Пользовательский шаблон ({f.name})"
                    data["_file_path"] = str(f.resolve())
                    data["_file_stem"] = f.stem
                    if t_name not in res:
                        res[t_name] = data
            except Exception as ex:
                log(f"Ошибка чтения шаблона {f}: {ex}")

    p_order = {
        "dstu_2026_gost_russian_names": 0,
        "dstu_2026_gost_v3": 1,
        "vak_rinc_article": 2,
        "gost_34_tz": 3,
        "corporate": 4,
        "default": 5
    }
    sorted_items = sorted(
        res.values(),
        key=lambda t: (p_order.get(t.get("name", ""), 50), t.get("display_title", t.get("name", "")).lower())
    )
    _TEMPLATES_LIST_CACHE["sig"] = cur_sig
    _TEMPLATES_LIST_CACHE["items"] = sorted_items
    return sorted_items

def get_templates_signature() -> tuple:
    """Возвращает сигнатуру (путь, mtime_ns, size) всех .json файлов шаблонов для надёжного отслеживания изменений."""
    sig = []
    seen = set()
    for d in get_templates_dirs():
        try:
            for fp in sorted(d.glob("*.json")):
                if fp.name == ".templates_index.json" or not fp.is_file():
                    continue
                rp = str(fp.resolve())
                if rp in seen:
                    continue
                seen.add(rp)
                st = fp.stat()
                sig.append((rp, st.st_mtime_ns, st.st_size))
        except Exception:
            pass
    for rp in get_registered_template_paths():
        try:
            rps = str(rp.resolve())
            if rps not in seen and rp.exists() and rp.is_file():
                seen.add(rps)
                st = rp.stat()
                sig.append((rps, st.st_mtime_ns, st.st_size))
        except Exception:
            pass
    return tuple(sorted(sig))

def load_template(name_or_path: str) -> dict:
    """Универсальная загрузка шаблона по имени (name), имени файла (stem) или полному пути."""
    raw = (name_or_path or "dstu_2026_gost_russian_names").strip()
    clean = raw.lower()
    if clean.endswith(".json"):
        clean = clean[:-5]

    p = Path(raw)
    if p.exists() and p.is_file():
        try:
            with open(p, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if isinstance(data, dict) and "styles" in data:
                return data
        except Exception:
            pass

    avail = list_available_templates()
    for t in avail:
        t_name = str(t.get("name", "")).strip()
        t_stem = str(t.get("_file_stem", "")).strip()
        t_title = str(t.get("display_title", "")).strip()
        if raw in (t_name, t_stem, t_title) or clean in (t_name.lower(), t_stem.lower(), t_title.lower()):
            return t

    for d in get_templates_dirs():
        for cand in (raw, f"{raw}.json", clean, f"{clean}.json"):
            cp = d / cand
            if cp.exists() and cp.is_file():
                try:
                    with open(cp, "r", encoding="utf-8-sig") as f:
                        return json.load(f)
                except Exception:
                    pass

    if avail:
        return avail[0]
    raise FileNotFoundError(f"Шаблон не найден: {name_or_path}")

def analyze_markdown_text(text: str) -> dict:
    if not text:
        return {"lines": 0, "chars": 0, "tables": 0, "formulas": 0, "headings": 0}
    lines = [l for l in text.splitlines() if l.strip()]
    chars = len(text)
    sep_tables = len(re.findall(r'^\s*\|(?:\s*:?-{2,}:?\s*\|)+\s*$', text, flags=re.MULTILINE))
    if sep_tables == 0:
        pipe_lines = len(re.findall(r'^\s*\|.*?\|\s*$', text, flags=re.MULTILINE))
        sep_tables = max(pipe_lines // 3, 1) if pipe_lines >= 3 else 0
    formulas = len(re.findall(r'\$\$[\s\S]+?\$\$|\$(?:[^\$\\]|\\.)+\$|\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]', text))
    headings = len(re.findall(r'^\s*#{1,6}\s+', text, flags=re.MULTILINE))
    return {
        "lines": len(lines),
        "chars": chars,
        "tables": sep_tables,
        "formulas": formulas,
        "headings": headings
    }

def extract_semantic_title(markdown_text: str) -> str:
    lines = [l.strip() for l in markdown_text.splitlines() if l.strip()]
    if not lines:
        return f"Document_{time.strftime('%Y%m%d_%H%M%S')}"
    for l in lines[:5]:
        m = re.match(r'^#+\s+(.+)$', l)
        if m:
            clean = re.sub(r'[^\w\s-]', '', m.group(1)).strip()
            clean = re.sub(r'[-\s]+', '_', clean)
            if clean:
                return clean[:40]
    first_line = re.sub(r'[^\w\s-]', '', lines[0]).strip()
    first_line = re.sub(r'[-\s]+', '_', first_line)
    return first_line[:30] or f"Document_{time.strftime('%Y%m%d_%H%M%S')}"

# ==============================================================================
# 2. СТИЛИЗАЦИЯ И ИНТЕРФЕЙС (1-В-1 ПО ДИЗАЙНУ MARKDOWN_TO_DOCX_SUITE.HTML)
# ==============================================================================


UI_CYRILLIC_FONTS = [
    ("По умолчанию (Inter / Segoe UI)", "Inter"),
    ("Cambria (с засечками)", "Cambria"),
    ("Georgia (с засечками)", "Georgia"),
    ("Times New Roman (с засечками)", "Times New Roman"),
    ("PT Serif (кириллица)", "PT Serif"),
    ("Charter (с засечками)", "Charter"),
    ("Liberation Serif (кириллица)", "Liberation Serif"),
    ("Segoe UI (Windows 11)", "Segoe UI"),
    ("PT Sans (кириллица)", "PT Sans"),
    ("Roboto (без засечек)", "Roboto"),
]

THEMES = {
    "light": {
        "bg_canvas": "#EEF2F6",
        "bg_window": "#FFFFFF",
        "bg_subtle": "#F8FAFC",
        "bg_muted": "#F1F5F9",
        "bg_hover": "#E2E8F0",
        "border": "#E2E8F0",
        "border_strong": "#CBD5E1",
        "text": "#0F172A",
        "text_2": "#334155",
        "text_mute": "#64748B",
        "accent": "#2563EB",
        "accent_hover": "#1D4ED8",
        "accent_soft": "#EFF6FF",
        "accent_text": "#1E40AF",
        "teal": "#0D9488",
        "teal_soft": "#F0FDFA",
        "teal_text": "#0F766E",
    },
    "dark": {
        "bg_canvas": "#090D16",
        "bg_window": "#111827",
        "bg_subtle": "#182032",
        "bg_muted": "#1E293B",
        "bg_hover": "#334155",
        "border": "#263244",
        "border_strong": "#37475E",
        "text": "#F8FAFC",
        "text_2": "#CBD5E1",
        "text_mute": "#94A3B8",
        "accent": "#3B82F6",
        "accent_hover": "#2563EB",
        "accent_soft": "#1A2B4B",
        "accent_text": "#93C5FD",
        "teal": "#14B8A6",
        "teal_soft": "#112E31",
        "teal_text": "#5EEAD4",
    }
}


def _ensure_ui_svg(filename: str, svg_content: str) -> str:
    svg_p = APP_DATA_DIR / filename
    if not svg_p.exists():
        try:
            svg_p.write_text(svg_content, encoding="utf-8")
        except Exception:
            return ""
    return svg_p.resolve().as_posix()

def _get_check_svg_url() -> str:
    svg_p = APP_DATA_DIR / "check_white.svg"
    if not svg_p.exists():
        try:
            svg_p.write_text(
                "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'>"
                "<path fill='none' stroke='#FFFFFF' stroke-width='2.2' "
                "stroke-linecap='round' stroke-linejoin='round' d='M3.5 8.5l3 3 6-6.5'/>"
                "</svg>",
                encoding="utf-8",
            )
        except Exception:
            return ""
    return svg_p.resolve().as_posix()

def build_stylesheet(theme_name: str = "light", ui_font: str = "Inter") -> str:
    chk_url = _get_check_svg_url()
    chk_img_rule = f'image: url("{chk_url}");' if chk_url else ""
    t_cur_spin = THEMES.get(theme_name, THEMES["light"])
    stroke_col = t_cur_spin["text_2"]
    up_url = _ensure_ui_svg(
        f"spin_up_{theme_name}.svg",
        f"<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 12'><path fill='none' stroke='{stroke_col}' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round' d='M3 7.5L6 4.5l3 3'/></svg>"
    )
    dn_url = _ensure_ui_svg(
        f"spin_dn_{theme_name}.svg",
        f"<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 12'><path fill='none' stroke='{stroke_col}' stroke-width='1.7' stroke-linecap='round' stroke-linejoin='round' d='M3 4.5L6 7.5l3-3'/></svg>"
    )
    up_img_rule = f'image: url("{up_url}");' if up_url else ""
    dn_img_rule = f'image: url("{dn_url}");' if dn_url else ""

    t = THEMES.get(theme_name, THEMES["light"])
    return f"""
    QWidget {{
        font-family: '{ui_font}', 'Segoe UI Variable Text', 'Segoe UI', 'Segoe UI Emoji', sans-serif;
    }}
    QMainWindow {{
        background: transparent;
    }}
    QWidget#MainRoot {{
        background: transparent;
    }}
    QDialog {{
        background-color: {t['bg_window']};
        color: {t['text']};
        font-size: 13.5px;
    }}
    QFrame#WindowFrame {{
        background-color: {t['bg_window']};
        color: {t['text']};
        border: 1px solid {t['border_strong']};
        border-radius: 14px;
    }}
    QLabel {{
        color: {t['text']};
        background: transparent;
    }}
    QWidget#WorkspaceBox {{
        background-color: {t['bg_window']};
    }}
    QFrame#AppHeader {{
        background-color: {t['bg_subtle']};
        border-bottom: 1px solid {t['border']};
        border-top-left-radius: 13px;
        border-top-right-radius: 13px;
    }}
    QFrame#AppFooter {{
        background-color: {t['bg_subtle']};
        border-top: 1px solid {t['border']};
        border-bottom-left-radius: 13px;
        border-bottom-right-radius: 13px;
    }}
    QPushButton.WinBtn {{
        background: transparent;
        border: none;
        border-radius: 7px;
        color: {t['text_mute']};
        font-size: 13px;
        font-weight: 600;
        padding: 0px;
    }}
    QPushButton.WinBtn:hover {{
        background-color: {t['bg_muted']};
        color: {t['text']};
    }}
    QPushButton#WinCloseBtn:hover {{
        background-color: #EF4444;
        color: #FFFFFF;
    }}
    QFrame#SegmentedBox {{
        background-color: {t['bg_muted']};
        border: 1px solid {t['border']};
        border-radius: 9px;
    }}
    QPushButton.SegBtn {{
        background: transparent;
        border: none;
        color: {t['text_mute']};
        font-size: 13px;
        font-weight: 500;
        min-height: 30px;
        padding: 6px 16px;
        border-radius: 7px;
    }}
    QPushButton.SegBtn:hover {{
        color: {t['text']};
    }}
    QPushButton.SegBtn[active="true"] {{
        background-color: {t['bg_window']};
        color: {t['text']};
        font-weight: 600;
        border: 1px solid {t['border']};
    }}
    QPushButton.NavBtn {{
        background: transparent;
        border: none;
        color: {t['text_2']};
        font-size: 13px;
        font-weight: 500;
        text-align: left;
        min-height: 32px;
        padding: 6px 12px;
        border-radius: 7px;
    }}
    QPushButton.NavBtn:hover {{
        background-color: {t['bg_muted']};
        color: {t['text']};
    }}
    QPushButton.NavBtn[active="true"] {{
        background-color: {t['accent_soft']};
        color: {t['accent_text']};
        font-weight: 600;
    }}
    QPushButton#TableTplBtn {{
        background-color: {t['bg_subtle']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        padding: 4px 10px;
        color: {t['accent_text']};
        font-size: 12px;
        font-weight: 600;
        text-align: left;
    }}
    QPushButton#TableTplBtn:hover {{
        background-color: {t['accent_soft']};
        border-color: {t['accent']};
    }}
    QPushButton#TplPill {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 8px;
        padding: 6px 12px;
        color: {t['accent_text']};
        font-size: 13px;
        font-weight: 600;
        text-align: left;
    }}
    QPushButton#TplPill:hover {{
        border-color: {t['accent']};
        background-color: {t['accent_soft']};
    }}
    QPushButton#HotkeyChip {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 6px 10px;
        color: {t['teal_text']};
        font-family: 'JetBrains Mono', 'Consolas', monospace;
        font-size: 12px;
        font-weight: 600;
    }}
    QPushButton#HotkeyChip:hover {{
        background-color: {t['teal_soft']};
        border-color: {t['teal']};
    }}
    QPushButton.IconBtn {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        color: {t['text_2']};
        font-size: 14px;
        padding: 0px;
    }}
    QPushButton.IconBtn:hover {{
        background-color: {t['bg_muted']};
        border-color: {t['border_strong']};
        color: {t['text']};
    }}
    QFrame.SurfaceCard, QFrame.PaneCard, QFrame.ContextBar {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 10px;
    }}
    QFrame.ContextBar {{
        background-color: {t['bg_subtle']};
    }}
    QFrame.PaneHeader {{
        background-color: {t['bg_subtle']};
        border-bottom: 1px solid {t['border']};
        border-top-left-radius: 10px;
        border-top-right-radius: 10px;
    }}
    QFrame.PaneFooter {{
        background-color: {t['bg_subtle']};
        border-top: 1px solid {t['border']};
        border-bottom-left-radius: 10px;
        border-bottom-right-radius: 10px;
    }}
    QFrame.DropHero {{
        background-color: {t['bg_subtle']};
        border: 2px dashed {t['border_strong']};
        border-radius: 10px;
    }}
    QFrame.DropHero[dragover="true"] {{
        background-color: {t['accent_soft']};
        border: 2px dashed {t['accent']};
    }}
    QFrame#ResultBanner {{
        background-color: {t['teal_soft']};
        border: 1px solid {t['teal']};
        border-radius: 8px;
    }}
    QPushButton.PrimaryBtn {{
        background-color: {t['accent']};
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        min-height: 36px;
        padding: 7px 20px;
        font-size: 13.5px;
        font-weight: 600;
    }}
    QPushButton.PrimaryBtn:hover {{
        background-color: {t['accent_hover']};
    }}
    QPushButton.PrimaryBtn:disabled {{
        background-color: {t['border_strong']};
        color: {t['text_mute']};
    }}
    QPushButton.SecondaryBtn {{
        background-color: {t['bg_window']};
        color: {t['text_2']};
        border: 1px solid {t['border_strong']};
        border-radius: 7px;
        min-height: 32px;
        padding: 6px 16px;
        font-size: 12.5px;
        font-weight: 500;
    }}
    QPushButton.SecondaryBtn:hover {{
        background-color: {t['bg_muted']};
        color: {t['text']};
        border-color: {t['text_mute']};
    }}
    QPushButton.SecondaryBtn:disabled {{
        color: {t['text_mute']};
        border-color: {t['border']};
    }}
    QPushButton#PathPill {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 8px;
        padding: 7px 12px;
        color: {t['text_2']};
        font-size: 12.5px;
        text-align: left;
    }}
    QPushButton#PathPill:hover {{
        border-color: {t['accent']};
    }}
    QLineEdit, QComboBox {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 7px;
        padding: 6px 10px;
        color: {t['text']};
        font-size: 13px;
    }}
    QLineEdit:focus, QComboBox:focus {{
        border-color: {t['accent']};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 24px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {t['bg_window']};
        color: {t['text']};
        selection-background-color: {t['accent_soft']};
        selection-color: {t['accent_text']};
        border: 1px solid {t['border_strong']};
        border-radius: 6px;
        padding: 4px;
    }}
    QTextEdit#MdEditor {{
        background-color: {t['bg_window']};
        color: {t['text']};
        border: none;
        padding: 12px;
        font-family: 'JetBrains Mono', 'Consolas', monospace;
        font-size: 13px;
    }}
    QScrollArea#PreviewScroll {{
        background-color: {t['bg_subtle']};
        border: none;
        border-bottom-left-radius: 9px;
        border-bottom-right-radius: 9px;
    }}
    QScrollArea#PopScroll {{
        background-color: transparent;
        border: none;
    }}
    QWidget#PreviewBody {{
        background-color: {t['bg_subtle']};
    }}
    QLabel.StatBadge {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        padding: 3px 8px;
        font-family: 'JetBrains Mono', 'Consolas', monospace;
        font-size: 11.5px;
        color: {t['text_2']};
    }}
    QFrame.OutlineItem {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 7px;
    }}
    QLabel.OutlineLvl {{
        background-color: {t['accent_soft']};
        color: {t['accent_text']};
        border-radius: 4px;
        padding: 2px 6px;
        font-family: 'JetBrains Mono', 'Consolas', monospace;
        font-size: 11px;
        font-weight: 600;
    }}
    QFrame#TableWrap {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 10px;
    }}
    QTableWidget {{
        background-color: transparent;
        border: none;
        gridline-color: transparent;
        color: {t['text_2']};
    }}
    QTableCornerButton::section {{
        background-color: {t['bg_subtle']};
        border: none;
    }}
    QHeaderView {{
        background-color: {t['bg_subtle']};
        border: none;
    }}
    QTableWidget::item {{
        padding: 8px 10px;
        border-bottom: 1px solid {t['border']};
    }}
    QTableWidget::item:selected {{
        background-color: {t['accent_soft']};
        color: {t['text']};
    }}
    QHeaderView::section {{
        background-color: {t['bg_subtle']};
        color: {t['text_mute']};
        padding: 8px 10px;
        font-size: 12px;
        font-weight: 600;
        border: none;
        border-bottom: 1px solid {t['border']};
    }}
    QListWidget {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 4px;
        color: {t['text']};
    }}
    QListWidget::item {{
        background-color: {t['bg_subtle']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        padding: 7px 10px;
        margin-bottom: 4px;
    }}
    QListWidget::item:selected {{
        background-color: {t['accent_soft']};
        border-color: {t['accent']};
        color: {t['text']};
    }}
    QProgressBar {{
        background-color: {t['bg_muted']};
        border: 1px solid {t['border']};
        border-radius: 6px;
        min-height: 22px;
        max-height: 24px;
        text-align: center;
        color: {t['text']};
        font-size: 11.5px;
        font-weight: 600;
    }}
    QProgressBar::chunk {{
        background-color: {t['accent']};
        border-radius: 5px;
    }}
    QSplitter::handle {{
        background-color: transparent;
        width: 12px;
    }}
    QSplitter::handle:hover {{
        background-color: {t['accent_soft']};
        border-radius: 4px;
    }}
    QCheckBox, QRadioButton {{
        color: {t['text_2']};
        spacing: 7px;
        font-size: 13px;
    }}
    QCheckBox::indicator, QRadioButton::indicator {{
        width: 15px;
        height: 15px;
        border: 1px solid {t['border_strong']};
        background-color: {t['bg_window']};
        border-radius: 4px;
    }}
    QRadioButton::indicator {{
        border-radius: 8px;
    }}
    QCheckBox::indicator:checked {{
        background-color: {t['accent']};
        border-color: {t['accent']};
        {chk_img_rule}
    }}
    QRadioButton::indicator:checked {{
        background-color: {t['accent']};
        border-color: {t['accent']};
    }}
    QFrame#OverlayPanel {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 12px;
    }}
    QFrame#DrawerPanel {{
        background-color: {t['bg_window']};
        border-left: 1px solid {t['border_strong']};
        border-top-right-radius: 13px;
        border-bottom-right-radius: 13px;
    }}
    QFrame.TplCard {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 8px;
    }}
    QFrame.TplCard:hover {{
        background-color: {t['bg_subtle']};
        border-color: {t['border_strong']};
    }}
    QFrame.TplCard[selected="true"] {{
        background-color: {t['accent_soft']};
        border: 1px solid {t['accent']};
    }}
    QPushButton.TplCard {{
        text-align: left;
        background-color: {t['bg_window']};
        border: 1px solid {t['border']};
        border-radius: 8px;
        padding: 10px 12px;
        color: {t['text']};
    }}
    QPushButton.TplCard:hover {{
        background-color: {t['bg_subtle']};
        border-color: {t['border_strong']};
    }}
    QPushButton.TplCard[selected="true"] {{
        background-color: {t['accent_soft']};
        border: 1px solid {t['accent']};
    }}
    QScrollBar:vertical {{
        background: transparent;
        width: 8px;
    }}
    QScrollBar::handle:vertical {{
        background: {t['border_strong']};
        border-radius: 4px;
        min-height: 24px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {t['text_mute']};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}
    QScrollBar:horizontal {{
        background: transparent;
        height: 8px;
    }}
    QScrollBar::handle:horizontal {{
        background: {t['border_strong']};
        border-radius: 4px;
        min-width: 24px;
    }}
    QScrollBar::handle:horizontal:hover {{
        background: {t['text_mute']};
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0;
    }}
    QMenu {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 10px;
        padding: 6px;
        color: {t['text']};
    }}
    QMenu::item {{
        padding: 7px 14px;
        border-radius: 6px;
        font-size: 13px;
        color: {t['text']};
    }}
    QMenu::item:selected {{
        background-color: {t['accent_soft']};
        color: {t['accent_text']};
    }}
    QMenu::item:disabled {{
        color: {t['text_mute']};
    }}
    QMenu::separator {{
        height: 1px;
        background: {t['border']};
        margin: 5px 8px;
    }}
    QMenu::indicator {{
        width: 15px;
        height: 15px;
        border-radius: 4px;
        border: 1px solid {t['border_strong']};
        background-color: {t['bg_window']};
        margin-left: 4px;
    }}
    QMenu::indicator:checked {{
        background-color: {t['accent']};
        border-color: {t['accent']};
        {chk_img_rule}
    }}
    QSpinBox, QDoubleSpinBox {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 7px;
        padding: 5px 24px 5px 8px;
        color: {t['text']};
        font-size: 13px;
        min-height: 22px;
    }}
    QSpinBox::up-button, QDoubleSpinBox::up-button {{
        subcontrol-origin: border;
        subcontrol-position: top right;
        width: 20px;
        border-left: 1px solid {t['border']};
        border-top-right-radius: 6px;
        background-color: {t['bg_subtle']};
    }}
    QSpinBox::down-button, QDoubleSpinBox::down-button {{
        subcontrol-origin: border;
        subcontrol-position: bottom right;
        width: 20px;
        border-left: 1px solid {t['border']};
        border-top: 1px solid {t['border']};
        border-bottom-right-radius: 6px;
        background-color: {t['bg_subtle']};
    }}
    QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover,
    QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {{
        background-color: {t['bg_muted']};
    }}
    QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{
        width: 10px;
        height: 10px;
        {up_img_rule}
    }}
    QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{
        width: 10px;
        height: 10px;
        {dn_img_rule}
    }}
    QPlainTextEdit {{
        background-color: {t['bg_window']};
        border: 1px solid {t['border_strong']};
        border-radius: 7px;
        padding: 5px 8px;
        color: {t['text']};
        font-size: 13px;
    }}
    QSpinBox:focus, QDoubleSpinBox:focus, QPlainTextEdit:focus {{
        border-color: {t['accent']};
    }}
    QToolTip {{
        background-color: {t['text']};
        color: {t['bg_window']};
        border: none;
        border-radius: 6px;
        padding: 5px 9px;
        font-size: 12px;
    }}
    QPushButton.PrimaryBtn:pressed {{
        background-color: {t['accent_text']};
    }}
    QPushButton.SecondaryBtn:pressed {{
        background-color: {t['bg_hover']};
    }}
    QPushButton::menu-indicator {{
        image: none;
        width: 0px;
    }}
    """


if HAS_PYSIDE6:

    def create_app_icon() -> QIcon:
        pix = QPixmap(64, 64)
        pix.fill(Qt.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor("#2563EB"))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(4, 4, 56, 56, 14, 14)
        p.setPen(QColor("#FFFFFF"))
        p.setFont(QFont("Segoe UI", 28, QFont.Bold))
        p.drawText(QRect(4, 4, 56, 56), Qt.AlignCenter, "W")
        p.end()
        return QIcon(pix)

    SVG_ICONS = {
        "settings": (
            "<svg viewBox='0 0 24 24' fill='none' stroke='{color}' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>"
            "<circle cx='12' cy='12' r='3'/>"
            "<path d='M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z'/>"
            "</svg>"
        ),
        "edit": (
            "<svg viewBox='0 0 24 24' fill='none' stroke='{color}' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>"
            "<path d='M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7'/>"
            "<path d='M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z'/>"
            "</svg>"
        ),
        "import": (
            "<svg viewBox='0 0 24 24' fill='none' stroke='{color}' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>"
            "<path d='M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4'/>"
            "<polyline points='7 10 12 15 17 10'/>"
            "<line x1='12' y1='15' x2='12' y2='3'/>"
            "</svg>"
        ),
        "refresh": (
            "<svg viewBox='0 0 24 24' fill='none' stroke='{color}' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>"
            "<polyline points='23 4 23 10 17 10'/>"
            "<polyline points='1 20 1 14 7 14'/>"
            "<path d='M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15'/>"
            "</svg>"
        ),
    }

    _SVG_ICON_CACHE: Dict[Tuple[str, str, int], QIcon] = {}

    def make_svg_icon(name: str, color: str = "#334155", size: int = 16) -> QIcon:
        cache_key = (name, color, size)
        if cache_key in _SVG_ICON_CACHE:
            return _SVG_ICON_CACHE[cache_key]
        svg_tmpl = SVG_ICONS.get(name, "")
        if not svg_tmpl:
            return QIcon()
        svg_bytes = svg_tmpl.format(color=color).encode("utf-8")
        scale = 4
        px_sz = size * scale
        pix = QPixmap(px_sz, px_sz)
        pix.fill(Qt.transparent)
        if HAS_QTSVG and QSvgRenderer is not None:
            renderer = QSvgRenderer(svg_bytes)
            painter = QPainter(pix)
            painter.setRenderHint(QPainter.Antialiasing, True)
            painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
            renderer.render(painter)
            painter.end()
        else:
            pix.loadFromData(svg_bytes, "SVG")
        pix.setDevicePixelRatio(float(scale))
        icon = QIcon()
        icon.addPixmap(pix)
        _SVG_ICON_CACHE[cache_key] = icon
        return icon

    def apply_menu_theme(menu, theme_name: str = "light"):
        if menu is None:
            return
        menu.setWindowFlags(menu.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        menu.setAttribute(Qt.WA_TranslucentBackground, True)
        t = theme_name if isinstance(theme_name, dict) else THEMES.get(theme_name, THEMES["light"])
        menu.setStyleSheet(
            f"QMenu {{ background-color: {t['bg_window']}; border: 1px solid {t['border_strong']}; "
            f"border-radius: 10px; padding: 6px; color: {t['text']}; }}"
            f"QMenu::item {{ padding: 7px 14px; border-radius: 6px; font-size: 13px; color: {t['text']}; }}"
            f"QMenu::item:selected {{ background-color: {t['accent_soft']}; color: {t['accent_text']}; }}"
            f"QMenu::item:disabled {{ color: {t['text_mute']}; }}"
            f"QMenu::separator {{ height: 1px; background: {t['border']}; margin: 4px 8px; }}"
        )

    def install_styled_context_menu(widget, theme_getter=None):
        """Заменяет стандартное системное контекстное меню PySide на органичное меню в стиле приложения."""
        if widget is None:
            return
        widget.setContextMenuPolicy(Qt.CustomContextMenu)

        def _show_menu(pos):
            menu = QMenu(widget)
            t_name = theme_getter() if callable(theme_getter) else "light"
            if isinstance(t_name, dict):
                t = t_name
            else:
                t = THEMES.get(t_name, THEMES["light"])
            menu.setStyleSheet(
                f"QMenu {{ background-color: {t['bg_window']}; border: 1px solid {t['border_strong']}; "
                f"border-radius: 10px; padding: 6px; color: {t['text']}; }}"
                f"QMenu::item {{ padding: 7px 14px; border-radius: 6px; font-size: 13px; }}"
                f"QMenu::item:selected {{ background-color: {t['accent_soft']}; color: {t['accent_text']}; }}"
                f"QMenu::item:disabled {{ color: {t['text_mute']}; }}"
                f"QMenu::separator {{ height: 1px; background: {t['border']}; margin: 4px 8px; }}"
            )
            ro = getattr(widget, "isReadOnly", lambda: False)()
            if hasattr(widget, "undo") and not ro:
                act_u = menu.addAction("Отменить")
                act_u.triggered.connect(widget.undo)
            if hasattr(widget, "redo") and not ro:
                act_r = menu.addAction("Повторить")
                act_r.triggered.connect(widget.redo)
            if menu.actions():
                menu.addSeparator()
            if hasattr(widget, "cut") and not ro:
                act_cut = menu.addAction("Вырезать")
                act_cut.triggered.connect(widget.cut)
            if hasattr(widget, "copy"):
                act_cp = menu.addAction("Копировать")
                act_cp.triggered.connect(widget.copy)
            if hasattr(widget, "paste") and not ro:
                act_ps = menu.addAction("Вставить")
                act_ps.triggered.connect(widget.paste)
            menu.addSeparator()
            if hasattr(widget, "selectAll"):
                act_sa = menu.addAction("Выделить всё")
                act_sa.triggered.connect(widget.selectAll)
            if hasattr(widget, "clear") and not ro:
                act_cl = menu.addAction("Очистить")
                act_cl.triggered.connect(widget.clear)
            menu.exec(widget.mapToGlobal(pos))
            menu.deleteLater()

        widget.customContextMenuRequested.connect(_show_menu)

    def show_themed_dialog(parent, title: str, message: str, theme_name: str = "light", is_error: bool = False):
        t = THEMES.get(theme_name, THEMES["light"])
        dlg = QDialog(parent)
        dlg.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        dlg.setAttribute(Qt.WA_TranslucentBackground, True)
        dlg.setMinimumWidth(380)
        outer = QVBoxLayout(dlg)
        outer.setContentsMargins(12, 12, 12, 12)
        box = QFrame()
        box.setAttribute(Qt.WA_StyledBackground, True)
        border_col = "#EF4444" if is_error else t["border_strong"]
        box.setStyleSheet(
            f"QFrame {{ background-color: {t['bg_window']}; border: 1px solid {border_col}; border-radius: 14px; }}"
        )

        outer.addWidget(box)

        lay = QVBoxLayout(box)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(10)
        lbl_t = QLabel(title)
        lbl_t.setStyleSheet(f"color: {t['text']}; font-size: 14.5px; font-weight: 700; border: none; background: transparent;")
        lay.addWidget(lbl_t)
        lbl_m = QLabel(message)
        lbl_m.setWordWrap(True)
        lbl_m.setStyleSheet(f"color: {t['text_2']}; font-size: 13px; border: none; background: transparent;")
        lay.addWidget(lbl_m)
        row = QHBoxLayout()
        row.addStretch()
        btn_ok = QPushButton("Понятно")
        btn_ok.setProperty("class", "PrimaryBtn")
        btn_ok.setCursor(QCursor(Qt.PointingHandCursor))
        btn_ok.clicked.connect(dlg.accept)
        row.addWidget(btn_ok)
        lay.addLayout(row)
        dlg.exec()
        dlg.deleteLater()

    class BatchWorker(QThread):
        progress_updated = Signal(int, int, str)
        item_finished = Signal(str, str, str, bool)
        all_completed = Signal(int, int)

        def __init__(self, task_items: List[Dict[str, Any]], templates_cache: Dict[str, dict], parent=None, two_columns: bool = False, strip_body_emphasis: Optional[bool] = None):
            super().__init__(parent)
            self.task_items = task_items
            self.templates_cache = templates_cache
            self.two_columns = two_columns
            self.strip_body_emphasis = strip_body_emphasis
            self._is_cancelled = False

        def cancel(self):
            self._is_cancelled = True

        def run(self):
            total = len(self.task_items)
            success_count = 0
            fail_count = 0
            for idx, item in enumerate(self.task_items, 1):
                if self._is_cancelled:
                    break
                in_p = Path(item["input_path"])
                out_p = Path(item["output_path"])
                tpl_name = item.get("template_name", "dstu_2026_gost_russian_names")
                self.progress_updated.emit(idx, total, in_p.name)
                try:
                    raw_tpl = load_template(tpl_name) or self.templates_cache.get(tpl_name)
                    tpl_cfg = json.loads(json.dumps(raw_tpl))
                    if self.two_columns:
                        tpl_cfg.setdefault("page", {})["columns"] = 2
                        tpl_cfg["page"].setdefault("column_spacing_mm", 5.0)
                    content = in_p.read_text(encoding="utf-8-sig")
                    if HAS_ENGINE:
                        parser = MarkdownASTParser(tpl_cfg.get("structural_titles", []))
                        ast_nodes = parser.parse(content)
                        strip_em = self.strip_body_emphasis if self.strip_body_emphasis is not None else load_config().get("strip_body_emphasis", False)
                        formatter = GostDocxFormatter(tpl_cfg, strip_body_emphasis=True if strip_em else None)
                        formatter.build_document(ast_nodes, str(out_p), base_dir=str(in_p.resolve().parent))
                    else:
                        raise RuntimeError("Ядро gost_engine не найдено.")
                    success_count += 1
                    self.item_finished.emit(str(in_p), str(out_p), "Готово", True)
                except Exception as e:
                    fail_count += 1
                    self.item_finished.emit(str(in_p), str(out_p), f"Ошибка: {e}", False)
            self.all_completed.emit(success_count, fail_count)


    class DropHeroWidget(QFrame):
        files_dropped = Signal(list)

        def __init__(self, parent=None):
            super().__init__(parent)
            self.setProperty("class", "DropHero")
            self.setAcceptDrops(True)
            lay = QHBoxLayout(self)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(14)

            copy_col = QVBoxLayout()
            copy_col.setSpacing(3)
            lbl_h = QLabel("Перетащите файлы (.md, .txt, .html) или целую папку")
            lbl_h.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            copy_col.addWidget(lbl_h)
            lbl_s = QLabel("Файлы будут обработаны выбранным шаблоном в фоновом потоке без зависания окна.")
            lbl_s.setStyleSheet("font-size: 12.5px; opacity: 0.72;")
            copy_col.addWidget(lbl_s)
            lay.addLayout(copy_col, 1)

            self.btn_box = QHBoxLayout()
            self.btn_box.setSpacing(8)
            lay.addLayout(self.btn_box)

        def dragEnterEvent(self, event: QDragEnterEvent):
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                self.setProperty("dragover", "true")
                self.style().unpolish(self)
                self.style().polish(self)

        def dragLeaveEvent(self, event):
            self.setProperty("dragover", "false")
            self.style().unpolish(self)
            self.style().polish(self)

        def dropEvent(self, event: QDropEvent):
            self.setProperty("dragover", "false")
            self.style().unpolish(self)
            self.style().polish(self)
            paths = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile()]
            if paths:
                self.files_dropped.emit(paths)


    class ToastOverlay(QWidget):
        """Плавающее toast-уведомление с плавной анимацией windowOpacity."""
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
            self.setAttribute(Qt.WA_TranslucentBackground)
            self.setAttribute(Qt.WA_ShowWithoutActivating, True)

            lay = QVBoxLayout(self)
            lay.setContentsMargins(6, 6, 6, 6)
            self.pill = QFrame()
            self.pill.setStyleSheet("background-color: #0F172A; color: #FFFFFF; border-radius: 8px;")
            p_lay = QHBoxLayout(self.pill)
            p_lay.setContentsMargins(16, 10, 16, 10)
            self.lbl = QLabel("")
            self.lbl.setStyleSheet("color: #FFFFFF; font-size: 13px; font-weight: 500; background: transparent;")
            p_lay.addWidget(self.lbl)
            lay.addWidget(self.pill)

            self._anim = QPropertyAnimation(self, b"windowOpacity", self)
            self._anim.setDuration(220)
            self._anim.setEasingCurve(QEasingCurve.OutCubic)
            self._anim.finished.connect(self._on_faded)

            self._hide_timer = QTimer(self)
            self._hide_timer.setSingleShot(True)
            self._hide_timer.timeout.connect(self._fade_out)

        def show_toast(self, message: str, theme_name: str = "light", target_path: Path = None):
            t = THEMES.get(theme_name, THEMES["light"])
            self.pill.setStyleSheet(
                f"background-color: {t['bg_subtle']}; color: {t['text']}; "
                f"border: 1px solid {t['border_strong']}; border-radius: 10px;"
            )
            self.lbl.setStyleSheet(f"color: {t['text']}; font-size: 13px; font-weight: 500; background: transparent; border: none;")
            self.lbl.setText(message)
            self.adjustSize()
            scr_obj = QApplication.primaryScreen()
            if scr_obj is not None:
                scr = scr_obj.availableGeometry()
                self.move(scr.right() - self.width() - 22, scr.bottom() - self.height() - 22)
            self._anim.stop()
            self.setWindowOpacity(0.0)
            self.show()
            self._anim.setStartValue(0.0)
            self._anim.setEndValue(1.0)
            self._anim.start()
            self._hide_timer.start(2600)

        def _fade_out(self):
            self._anim.stop()
            self._anim.setStartValue(self.windowOpacity())
            self._anim.setEndValue(0.0)
            self._anim.start()

        def _on_faded(self):
            if self.windowOpacity() <= 0.05:
                self.hide()


    class TemplateSelectorDialog(QDialog):
        """Безрамочное модальное окно быстрой сборки из буфера (Ctrl+Alt+W)."""
        def __init__(self, detected_title: str, stats: dict, theme_name: str = "light", parent=None):
            super().__init__(parent)
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground, True)
            self.detected_title = detected_title
            self.stats = stats
            self.theme_name = theme_name
            self.avail_templates = list_available_templates()
            cfg = load_config()
            self.ui_font = cfg.get("ui_font", "Inter")
            self.selected_template = cfg.get("default_template") or (self.avail_templates[0]["name"] if self.avail_templates else "dstu_2026_gost_russian_names")
            self.append_mode = False
            self.append_target = cfg.get("last_append_target", "")
            self.two_columns = bool(cfg.get("two_columns", False))
            self.strip_body_emphasis = bool(cfg.get("strip_body_emphasis", False))
            self._drag_pos = None
            self._fade_anim = QPropertyAnimation(self, b"windowOpacity", self)
            self._fade_anim.setDuration(160)
            self._fade_anim.setEasingCurve(QEasingCurve.OutCubic)

            self.setFixedWidth(520)
            self.setStyleSheet(build_stylesheet(self.theme_name, self.ui_font))
            self._build_ui()

        def showEvent(self, event):
            super().showEvent(event)
            self.setWindowOpacity(0.0)
            self._fade_anim.stop()
            self._fade_anim.setStartValue(0.0)
            self._fade_anim.setEndValue(1.0)
            self._fade_anim.start()
            self.raise_()
            self.activateWindow()
            if sys.platform == "win32" and user32:
                try:
                    hwnd = int(self.winId())
                    user32.SetForegroundWindow(hwnd)
                    user32.BringWindowToTop(hwnd)
                except Exception:
                    pass

        def mousePressEvent(self, event):
            if event.button() == Qt.LeftButton:
                self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()

        def mouseMoveEvent(self, event):
            if self._drag_pos is not None and event.buttons() & Qt.LeftButton:
                self.move(event.globalPosition().toPoint() - self._drag_pos)
                event.accept()

        def mouseReleaseEvent(self, event):
            self._drag_pos = None

        def keyPressEvent(self, event):
            if event.key() == Qt.Key_Escape:
                self.reject()
                return
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self._on_accept()
                return
            key_text = event.text()
            if key_text.isdigit() and 1 <= int(key_text) <= 9:
                idx = int(key_text) - 1
                if idx < len(self.card_widgets):
                    _, t_name = self.card_widgets[idx]
                    self._select_card(t_name)
                    self._on_accept()
                    return
            super().keyPressEvent(event)

        def _build_ui(self):
            t = THEMES.get(self.theme_name, THEMES["light"])

            outer = QVBoxLayout(self)
            outer.setContentsMargins(12, 12, 12, 12)

            container = QFrame()
            container.setObjectName("DlgContainer")
            container.setAttribute(Qt.WA_StyledBackground, True)
            container.setStyleSheet(
                f"QFrame#DlgContainer {{ background-color: {t['bg_window']}; "
                f"border: 1px solid {t['border_strong']}; border-radius: 14px; }}"
            )

            outer.addWidget(container)

            lay = QVBoxLayout(container)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            hdr = QHBoxLayout()
            lbl_h = QLabel("Быстрая сборка из буфера (Ctrl+Alt+W)")
            lbl_h.setStyleSheet(f"color: {t['text']}; font-size: 14px; font-weight: 700; background: transparent;")
            hdr.addWidget(lbl_h)
            hdr.addStretch()
            btn_esc = QPushButton("Esc")
            btn_esc.setProperty("class", "SecondaryBtn")
            btn_esc.clicked.connect(self.reject)
            hdr.addWidget(btn_esc)
            lay.addLayout(hdr)

            info_box = QFrame()
            info_box.setAttribute(Qt.WA_StyledBackground, True)
            info_box.setStyleSheet(f"background-color: {t['bg_subtle']}; border: 1px solid {t['border']}; border-radius: 8px;")
            ib_lay = QHBoxLayout(info_box)
            ib_lay.setContentsMargins(12, 8, 12, 8)
            lbl_st = QLabel(
                f"В буфере: <b>{self.stats.get('lines', 0)} строк</b> · "
                f"<b>{self.stats.get('formulas', 0)} формул LaTeX</b> · "
                f"<b>{self.stats.get('tables', 0)} таблиц</b>"
            )
            lbl_st.setStyleSheet(f"color: {t['text_2']}; font-size: 12.5px; background: transparent;")
            ib_lay.addWidget(lbl_st)
            lay.addWidget(info_box)

            # Однородный прозрачный список карточек с горизонтальной и вертикальной прокруткой в едином стиле
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setMinimumHeight(210)
            scroll.setMaximumHeight(280)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
            scroll.viewport().setStyleSheet("background: transparent;")

            cont = QWidget()
            cont.setStyleSheet("background: transparent;")
            c_lay = QVBoxLayout(cont)
            c_lay.setContentsMargins(0, 2, 4, 4)
            c_lay.setSpacing(6)

            self.card_widgets = []
            for idx, tpl in enumerate(self.avail_templates):
                t_name = tpl.get("name", f"tpl_{idx}")
                t_title = tpl.get("display_title", t_name)
                t_desc = tpl.get("specs_summary") or tpl.get("description", "")
                is_sel = (t_name == self.selected_template)

                card = QFrame()
                card.setProperty("class", "TplCard")
                card.setProperty("selected", "true" if is_sel else "false")
                card.setAttribute(Qt.WA_StyledBackground, True)
                card.setCursor(QCursor(Qt.PointingHandCursor))
                card.mousePressEvent = lambda e, n=t_name: self._select_card(n)

                cl = QVBoxLayout(card)
                cl.setContentsMargins(12, 9, 12, 9)
                cl.setSpacing(2)

                top_r = QHBoxLayout()
                lbl_t = QLabel(f"{idx+1}. {t_title}")
                lbl_t.setStyleSheet(f"color: {t['text']}; font-size: 13px; font-weight: 600; background: transparent;")
                top_r.addWidget(lbl_t, 1)
                if is_sel:
                    lbl_a = QLabel("Активен")
                    lbl_a.setStyleSheet(f"color: {t['accent_text']}; font-size: 11.5px; font-weight: 600; background: transparent;")
                    top_r.addWidget(lbl_a)
                cl.addLayout(top_r)

                if t_desc:
                    lbl_d = QLabel(t_desc)
                    lbl_d.setStyleSheet(f"color: {t['text_mute']}; font-size: 12px; background: transparent;")
                    cl.addWidget(lbl_d)

                c_lay.addWidget(card)
                self.card_widgets.append((card, t_name))

            c_lay.addStretch()
            scroll.setWidget(cont)
            lay.addWidget(scroll)

            # Опция Append без рамки (только галочка и выбор файла)
            app_box = QWidget()
            app_box.setStyleSheet("background: transparent;")
            ab_lay = QVBoxLayout(app_box)
            ab_lay.setContentsMargins(2, 4, 2, 2)
            ab_lay.setSpacing(6)
            self.chk_append = QCheckBox("Дописать в конец существующего DOCX (Append)")
            self.chk_append.toggled.connect(self._on_append_toggled)
            ab_lay.addWidget(self.chk_append)
            row_a = QHBoxLayout()
            self.ed_target = QLineEdit(self.append_target)
            self.ed_target.setPlaceholderText("Выберите целевой .docx...")
            self.ed_target.setEnabled(False)
            install_styled_context_menu(self.ed_target, lambda: self.theme_name)
            row_a.addWidget(self.ed_target, 1)
            self.btn_browse = QPushButton("Выбрать .docx")
            self.btn_browse.setProperty("class", "SecondaryBtn")
            self.btn_browse.setEnabled(False)
            self.btn_browse.clicked.connect(self._browse_target)
            row_a.addWidget(self.btn_browse)
            ab_lay.addLayout(row_a)
            lay.addWidget(app_box)

            foot = QHBoxLayout()
            btn_dlg_opts = QPushButton("⚙️ Опции ▾")
            btn_dlg_opts.setProperty("class", "SecondaryBtn")
            btn_dlg_opts.setCursor(QCursor(Qt.PointingHandCursor))
            dlg_menu = QMenu(btn_dlg_opts)
            self.act_dlg_2col = dlg_menu.addAction("2 колонки (IEEE / ВАК)")
            self.act_dlg_2col.setCheckable(True)
            self.act_dlg_2col.setChecked(self.two_columns)
            self.act_dlg_2col.toggled.connect(lambda c: setattr(self, "two_columns", bool(c)))
            self.act_dlg_strip = dlg_menu.addAction("Убрать жирный/курсив в тексте")
            self.act_dlg_strip.setCheckable(True)
            self.act_dlg_strip.setChecked(self.strip_body_emphasis)
            self.act_dlg_strip.toggled.connect(lambda c: setattr(self, "strip_body_emphasis", bool(c)))
            btn_dlg_opts.setMenu(dlg_menu)
            foot.addWidget(btn_dlg_opts)
            foot.addStretch()
            btn_ok = QPushButton("Сформировать DOCX →")
            btn_ok.setProperty("class", "PrimaryBtn")
            btn_ok.setMinimumHeight(38)
            btn_ok.clicked.connect(self._on_accept)
            foot.addWidget(btn_ok)
            lay.addLayout(foot)

        def _select_card(self, t_name: str):
            self.selected_template = t_name
            for card, name in self.card_widgets:
                card.setProperty("selected", "true" if name == t_name else "false")
                card.style().unpolish(card)
                card.style().polish(card)

        def _on_append_toggled(self, checked: bool):
            self.append_mode = checked
            self.ed_target.setEnabled(checked)
            self.btn_browse.setEnabled(checked)

        def _browse_target(self):
            f, _ = QFileDialog.getOpenFileName(self, "Выберите целевой DOCX", "", "Word Documents (*.docx)")
            if f:
                self.ed_target.setText(f)
                self.append_target = f

        def _on_accept(self):
            self.append_mode = self.chk_append.isChecked()
            self.append_target = self.ed_target.text().strip()
            if self.append_mode:
                if not self.append_target or not Path(self.append_target).exists():
                    show_themed_dialog(self, "Файл не выбран", "Укажите существующий файл .docx для дозаписи.", self.theme_name)
                    return
                cfg = load_config()
                cfg["last_append_target"] = self.append_target
                save_config(cfg)
            self.accept()


    class HotkeyEmitter(QObject):
        trigger_clipboard = Signal()
        trigger_main_window = Signal()


    class UnifiedMainWindow(QMainWindow):
        """Главное окно приложения — 1-в-1 по дизайн-системе markdown_to_docx_suite.html."""
        def __init__(self, app_controller=None):
            super().__init__()
            self.app_controller = app_controller
            self.config = load_config()
            self.theme_name = self.config.get("theme", "light")
            self.mode = "text"
            self.templates_cache = {t["name"]: t for t in list_available_templates()}
            self.active_tpl = self.config.get("default_template", "dstu_2026_gost_russian_names")
            if self.active_tpl not in self.templates_cache and self.templates_cache:
                self.active_tpl = next(iter(self.templates_cache))

            self.queue_items: List[Dict[str, Any]] = []
            self.merge_files: List[str] = []
            self._qt_base_dir: Optional[str] = None
            self._last_output: Optional[Path] = None

            self.setWindowTitle("Markdown to DOCX Suite")
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Window)
            self.setAttribute(Qt.WA_TranslucentBackground, True)
            self._win_drag_pos = None
            self._force_closing = False
            self._win_fade = QPropertyAnimation(self, b"windowOpacity", self)
            self._win_fade.setEasingCurve(QEasingCurve.OutCubic)
            self._win_fade.finished.connect(self._on_win_fade_finished)
            self.setMinimumSize(1060, 690)
            self.resize(1200, 740)
            self.setAcceptDrops(True)

            self._build_ui()
            self._install_edge_resize()
            self._setup_template_watcher()
            self._apply_theme()
            self._sync_template_ui()
            self._parse_preview()
            self._update_footer_button()

        def _on_win_fade_finished(self):
            if self.windowOpacity() <= 0.05:
                self.hide()
                if hasattr(self, "_tpl_poll_timer"):
                    self._tpl_poll_timer.stop()
                gc.collect()

        def showEvent(self, event):
            super().showEvent(event)
            if hasattr(self, "_tpl_poll_timer") and not self._tpl_poll_timer.isActive():
                self._check_templates_changed()
                self._tpl_poll_timer.start()
            self._win_fade.stop()
            self._win_fade.setDuration(190)
            self.setWindowOpacity(0.0)
            self._win_fade.setStartValue(0.0)
            self._win_fade.setEndValue(1.0)
            self._win_fade.start()

        def closeEvent(self, event):
            if self._force_closing:
                event.accept()
                return
            event.ignore()
            self._win_fade.stop()
            self._win_fade.setDuration(160)
            self._win_fade.setStartValue(self.windowOpacity())
            self._win_fade.setEndValue(0.0)
            self._win_fade.start()
            if self.app_controller and hasattr(self.app_controller, "toast"):
                self.app_controller.toast.show_toast("Свёрнуто в трей. Хоткеи Ctrl+Alt+W / Ctrl+Alt+B активны", self.theme_name)

        def cleanup(self):
            self._force_closing = True
            for t_attr in ("_tpl_poll_timer", "_tpl_reload_timer", "_preview_timer"):
                t_obj = getattr(self, t_attr, None)
                if t_obj is not None:
                    t_obj.stop()
            if hasattr(self, "_tpl_watcher"):
                try:
                    dirs = self._tpl_watcher.directories()
                    files = self._tpl_watcher.files()
                    if dirs:
                        self._tpl_watcher.removePaths(dirs)
                    if files:
                        self._tpl_watcher.removePaths(files)
                except Exception:
                    pass
            w = getattr(self, "worker", None)
            if w is not None:
                try:
                    if w.isRunning():
                        w.cancel()
                        w.wait(2000)
                except Exception:
                    pass
            self.hide()

        def dragEnterEvent(self, event: QDragEnterEvent):
            if event.mimeData().hasUrls():
                event.acceptProposedAction()

        def dropEvent(self, event: QDropEvent):
            paths = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile()]
            if not paths:
                return
            docx_paths = [p for p in paths if Path(p).is_file() and Path(p).suffix.lower() == ".docx"]
            if docx_paths:
                if self.mode == "text" and self.rb_append.isChecked() and len(docx_paths) == 1:
                    self.ed_append_target.setText(docx_paths[0])
                    self.config["last_append_target"] = docx_paths[0]
                    save_config(self.config)
                    self._toast(f"Выбран целевой DOCX: {Path(docx_paths[0]).name}")
                else:
                    self._switch_mode("merge")
                    added_m = 0
                    for dp in docx_paths:
                        if dp not in self.merge_files:
                            self.merge_files.append(dp)
                            added_m += 1
                    if added_m:
                        self._sync_merge_list_ui()
                        self._toast(f"Добавлено DOCX для склейки: {added_m}")
                return
            if len(paths) == 1 and Path(paths[0]).is_file() and Path(paths[0]).suffix.lower() in (".md", ".markdown", ".txt"):
                fp = Path(paths[0])
                self._qt_base_dir = str(fp.resolve().parent)
                self.ed_input.setPlainText(fp.read_text(encoding="utf-8-sig"))
                self._switch_mode("text")
                self._toast(f"Загружен файл: {fp.name}")
            else:
                self._switch_mode("files")
                self._on_files_dropped(paths)

        def resizeEvent(self, event):
            super().resizeEvent(event)
            self._layout_overlays()

        def _on_ui_font_changed(self):
            ui_f = self.combo_ui_font.currentData() or "Inter"
            self.config["ui_font"] = ui_f
            save_config(self.config)
            self._apply_theme()

        def _apply_theme(self):
            ui_f = self.config.get("ui_font", "Inter")
            if QApplication.instance() is not None:
                QApplication.instance().setFont(QFont(ui_f, 10))
            self.setStyleSheet(build_stylesheet(self.theme_name, ui_f))
            if hasattr(self, "btn_theme_light") and hasattr(self, "btn_theme_dark"):
                self.btn_theme_light.setProperty("active", "true" if self.theme_name == "light" else "false")
                self.btn_theme_dark.setProperty("active", "true" if self.theme_name == "dark" else "false")
                for b_th in (self.btn_theme_light, self.btn_theme_dark):
                    b_th.style().unpolish(b_th)
                    b_th.style().polish(b_th)
            if hasattr(self, "tpl_popover"):
                self._render_tpl_cards()
            t_cur = THEMES.get(self.theme_name, THEMES["light"])
            if hasattr(self, "btn_settings"):
                self.btn_settings.setIcon(make_svg_icon("settings", t_cur["text_2"], 17))
            if hasattr(self, "btn_pop_edit"):
                self.btn_pop_edit.setIcon(make_svg_icon("edit", t_cur["text_2"], 15))
            if hasattr(self, "btn_pop_imp"):
                self.btn_pop_imp.setIcon(make_svg_icon("import", t_cur["text_2"], 15))
            if hasattr(self, "btn_pop_ref"):
                self.btn_pop_ref.setIcon(make_svg_icon("refresh", t_cur["text_2"], 15))
            if hasattr(self, "lbl_hdr_tpl_text"):
                self.lbl_hdr_tpl_text.setStyleSheet(f"color: {t_cur['accent_text']}; background: transparent; border: none; font-size: 13px; font-weight: 600;")
            for m_obj in self.findChildren(QMenu):
                apply_menu_theme(m_obj, self.theme_name)
            if self.app_controller and getattr(self.app_controller, "_tray_menu", None) is not None:
                apply_menu_theme(self.app_controller._tray_menu, self.theme_name)
            if getattr(self, "template_editor", None) is not None:
                self.template_editor.apply_theme()
            if hasattr(self, "table_queue"):
                self._refresh_queue_table()
            if hasattr(self, "ed_input"):
                self._parse_preview()

        def _set_theme(self, name: str):
            if name not in ("light", "dark") or name == self.theme_name:
                return
            self.theme_name = name
            self.config["theme"] = self.theme_name
            save_config(self.config)
            self._apply_theme()
            self._toast("Включена тёмная тема" if name == "dark" else "Включена светлая тема")

        def _toggle_theme(self):
            self._set_theme("dark" if self.theme_name == "light" else "light")

        def _toggle_maximize(self):
            if self.isMaximized():
                self.showNormal()
                if hasattr(self, "_outer_lay"):
                    self._outer_lay.setContentsMargins(2, 2, 2, 2)
            else:
                if hasattr(self, "_outer_lay"):
                    self._outer_lay.setContentsMargins(0, 0, 0, 0)
                self.showMaximized()

        def _install_edge_resize(self):
            self._resize_margin = 7
            self._resize_edges = Qt.Edges()
            self._resize_active = False
            self._resize_start_pos = None
            self._resize_start_geom = None
            if QApplication.instance() is not None:
                QApplication.instance().installEventFilter(self)

        def _detect_edges(self, global_pos: QPoint) -> Qt.Edges:
            if self.isMaximized():
                return Qt.Edges()
            local = self.mapFromGlobal(global_pos)
            w, h = self.width(), self.height()
            m = self._resize_margin
            if not (0 <= local.x() <= w and 0 <= local.y() <= h):
                return Qt.Edges()
            edges = Qt.Edges()
            if local.x() <= m:
                edges |= Qt.LeftEdge
            elif local.x() >= w - m:
                edges |= Qt.RightEdge
            if local.y() <= m:
                edges |= Qt.TopEdge
            elif local.y() >= h - m:
                edges |= Qt.BottomEdge
            return edges

        def _cursor_for_edges(self, edges: Qt.Edges):
            if edges in (Qt.LeftEdge | Qt.TopEdge, Qt.RightEdge | Qt.BottomEdge):
                return Qt.SizeFDiagCursor
            if edges in (Qt.RightEdge | Qt.TopEdge, Qt.LeftEdge | Qt.BottomEdge):
                return Qt.SizeBDiagCursor
            if edges & (Qt.LeftEdge | Qt.RightEdge):
                return Qt.SizeHorCursor
            if edges & (Qt.TopEdge | Qt.BottomEdge):
                return Qt.SizeVerCursor
            return None

        def eventFilter(self, obj, event):
            try:
                if self.isVisible() and hasattr(event, "globalPosition") and isinstance(obj, QWidget) and obj.window() is self:
                    etype = event.type()
                    if etype == event.Type.MouseMove and not self._resize_active:
                        edges = self._detect_edges(event.globalPosition().toPoint())
                        cur = self._cursor_for_edges(edges)
                        if cur is not None:
                            self.setCursor(cur)
                        elif self.cursor().shape() in (Qt.SizeHorCursor, Qt.SizeVerCursor, Qt.SizeFDiagCursor, Qt.SizeBDiagCursor):
                            self.unsetCursor()
                    elif etype == event.Type.MouseButtonPress and event.button() == Qt.LeftButton:
                        edges = self._detect_edges(event.globalPosition().toPoint())
                        if edges:
                            wh = self.windowHandle()
                            if wh is not None and wh.startSystemResize(edges):
                                return True
                            self._resize_active = True
                            self._resize_edges = edges
                            self._resize_start_pos = event.globalPosition().toPoint()
                            self._resize_start_geom = self.geometry()
                            return True
                    elif etype == event.Type.MouseMove and self._resize_active:
                        delta = event.globalPosition().toPoint() - self._resize_start_pos
                        g = QRect(self._resize_start_geom)
                        min_w, min_h = self.minimumWidth(), self.minimumHeight()
                        if self._resize_edges & Qt.LeftEdge:
                            nw = max(min_w, g.width() - delta.x())
                            g.setLeft(g.right() - nw + 1)
                        if self._resize_edges & Qt.RightEdge:
                            g.setWidth(max(min_w, g.width() + delta.x()))
                        if self._resize_edges & Qt.TopEdge:
                            nh = max(min_h, g.height() - delta.y())
                            g.setTop(g.bottom() - nh + 1)
                        if self._resize_edges & Qt.BottomEdge:
                            g.setHeight(max(min_h, g.height() + delta.y()))
                        self.setGeometry(g)
                        return True
                    elif etype == event.Type.MouseButtonRelease and self._resize_active:
                        self._resize_active = False
                        self.unsetCursor()
                        return True
            except Exception:
                pass
            return super().eventFilter(obj, event)

        def _hdr_mouse_press(self, event):
            if event.button() == Qt.LeftButton:
                self._win_drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()

        def _hdr_mouse_move(self, event):
            if self._win_drag_pos is not None and event.buttons() & Qt.LeftButton and not self.isMaximized():
                self.move(event.globalPosition().toPoint() - self._win_drag_pos)
                event.accept()

        def _hdr_mouse_release(self, event):
            self._win_drag_pos = None

        def _build_ui(self):
            outer_root = QWidget()
            outer_root.setObjectName("MainRoot")
            outer_root.setMouseTracking(True)
            self.setCentralWidget(outer_root)
            self._outer_lay = QVBoxLayout(outer_root)
            self._outer_lay.setContentsMargins(2, 2, 2, 2)
            self._outer_lay.setSpacing(0)

            root = QFrame()
            root.setObjectName("WindowFrame")
            root.setAttribute(Qt.WA_StyledBackground, True)

            self.window_frame = root
            self._outer_lay.addWidget(root)

            main_lay = QVBoxLayout(root)
            main_lay.setContentsMargins(1, 1, 1, 1)
            main_lay.setSpacing(0)

            # --- 1. HEADER ---
            hdr = QFrame()
            hdr.setObjectName("AppHeader")
            hdr.setAttribute(Qt.WA_StyledBackground, True)
            hdr.mousePressEvent = self._hdr_mouse_press
            hdr.mouseMoveEvent = self._hdr_mouse_move
            hdr.mouseReleaseEvent = self._hdr_mouse_release
            hdr.mouseDoubleClickEvent = lambda e: self._toggle_maximize()
            h_lay = QHBoxLayout(hdr)
            h_lay.setContentsMargins(18, 12, 18, 12)
            h_lay.setSpacing(12)

            # Brand
            brand_box = QHBoxLayout()
            brand_box.setSpacing(10)
            lbl_icon = QLabel("W")
            lbl_icon.setFixedSize(30, 30)
            lbl_icon.setAlignment(Qt.AlignCenter)
            lbl_icon.setStyleSheet("background-color: #2563EB; color: #FFFFFF; border-radius: 8px; font-weight: 700; font-size: 13px;")
            brand_box.addWidget(lbl_icon)
            lbl_brand = QLabel("Markdown to DOCX")
            lbl_brand.setStyleSheet("font-weight: 700; font-size: 14px;")
            brand_box.addWidget(lbl_brand)
            h_lay.addLayout(brand_box)
            h_lay.addStretch()

            # Segmented Mode Switcher
            seg = QFrame()
            seg.setObjectName("SegmentedBox")
            seg_lay = QHBoxLayout(seg)
            seg_lay.setContentsMargins(3, 3, 3, 3)
            seg_lay.setSpacing(2)
            self.seg_buttons = {}
            for m_id, m_lbl in [("text", "Текст / MD"), ("files", "Файлы"), ("merge", "Сборка DOCX"), ("templates", "Шаблоны")]:
                b = QPushButton(m_lbl)
                b.setProperty("class", "SegBtn")
                b.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
                b.setProperty("active", "true" if m_id == "text" else "false")
                b.setCursor(QCursor(Qt.PointingHandCursor))
                b.clicked.connect(lambda _, mid=m_id: self._switch_mode(mid))
                seg_lay.addWidget(b)
                self.seg_buttons[m_id] = b
            h_lay.addWidget(seg)
            h_lay.addStretch()

            # Right controls
            self.btn_tpl_open = QPushButton("")
            self.btn_tpl_open.setFixedSize(250, 34)
            self.btn_tpl_open.setObjectName("TplPill")
            self.btn_tpl_open.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_tpl_open.clicked.connect(self._toggle_tpl_popover)
            tpl_btn_lay = QHBoxLayout(self.btn_tpl_open)
            tpl_btn_lay.setContentsMargins(11, 0, 10, 0)
            tpl_btn_lay.setSpacing(6)
            self.lbl_hdr_tpl_text = QLabel("Шаблон: ...")
            self.lbl_hdr_tpl_text.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.lbl_hdr_tpl_text.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            self.lbl_hdr_tpl_text.setStyleSheet("background: transparent; border: none; font-size: 13px; font-weight: 600;")
            lbl_hdr_arrow = QLabel("▾")
            lbl_hdr_arrow.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            lbl_hdr_arrow.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            lbl_hdr_arrow.setStyleSheet("background: transparent; border: none; font-size: 12px;")
            tpl_btn_lay.addWidget(self.lbl_hdr_tpl_text, 1)
            tpl_btn_lay.addWidget(lbl_hdr_arrow, 0)
            h_lay.addWidget(self.btn_tpl_open)

            self.btn_sim_hotkey = QPushButton("Ctrl+Alt+W")
            self.btn_sim_hotkey.setMinimumHeight(34)
            self.btn_sim_hotkey.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            self.btn_sim_hotkey.setObjectName("HotkeyChip")
            self.btn_sim_hotkey.setToolTip("Быстрая сборка из буфера обмена (Ctrl+Alt+W)")
            self.btn_sim_hotkey.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_sim_hotkey.clicked.connect(self._trigger_hotkey_dialog)
            h_lay.addWidget(self.btn_sim_hotkey)

            self.btn_settings = QPushButton("")
            self.btn_settings.setProperty("class", "IconBtn")
            self.btn_settings.setFixedSize(34, 34)
            self.btn_settings.setIconSize(QSize(17, 17))
            self.btn_settings.setToolTip("Параметры приложения")
            self.btn_settings.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_settings.clicked.connect(self._open_settings_drawer)
            h_lay.addWidget(self.btn_settings)

            btn_w_hide = QPushButton("✕")
            btn_w_hide.setObjectName("WinCloseBtn")
            btn_w_hide.setProperty("class", "WinBtn")
            btn_w_hide.setFixedSize(34, 34)
            btn_w_hide.setToolTip("Скрыть в системный трей")
            btn_w_hide.setCursor(QCursor(Qt.PointingHandCursor))
            btn_w_hide.clicked.connect(self.close)
            h_lay.addWidget(btn_w_hide)

            main_lay.addWidget(hdr)

            # --- 2. WORKSPACE (STACKED MODES) ---
            ws = QWidget()
            ws.setObjectName("WorkspaceBox")
            ws.setAttribute(Qt.WA_StyledBackground, True)
            ws_lay = QVBoxLayout(ws)
            ws_lay.setContentsMargins(18, 16, 18, 16)
            ws_lay.setSpacing(14)

            self.stack = QStackedWidget()
            self.stack.addWidget(self._build_text_panel())
            self.stack.addWidget(self._build_files_panel())
            self.stack.addWidget(self._build_merge_panel())
            self.stack.addWidget(self._build_templates_panel())
            ws_lay.addWidget(self.stack, 1)
            main_lay.addWidget(ws, 1)

            # --- 3. FOOTER ---
            ftr = QFrame()
            ftr.setObjectName("AppFooter")
            f_lay = QHBoxLayout(ftr)
            f_lay.setContentsMargins(18, 12, 18, 12)
            f_lay.setSpacing(12)

            self.btn_save_mode = QPushButton("Сохранять: Рядом с исходным файлом  ✏️")
            self.btn_save_mode.setMinimumHeight(34)
            self.btn_save_mode.setObjectName("PathPill")
            self.btn_save_mode.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_save_mode.clicked.connect(self._toggle_save_mode)
            f_lay.addWidget(self.btn_save_mode)
            self._sync_save_mode_label()

            f_lay.addStretch()

            self.btn_main_run = QPushButton("Сформировать DOCX →")
            self.btn_main_run.setMinimumHeight(38)
            self.btn_main_run.setProperty("class", "PrimaryBtn")
            self.btn_main_run.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_main_run.clicked.connect(self._on_main_run)
            f_lay.addWidget(self.btn_main_run)
            size_grip = QSizeGrip(self)
            size_grip.setFixedSize(14, 14)
            size_grip.setStyleSheet("background: transparent;")
            f_lay.addWidget(size_grip, 0, Qt.AlignBottom | Qt.AlignRight)
            main_lay.addWidget(ftr)

            # --- 4. OVERLAYS (BACKDROP, TPL POPOVER, SETTINGS DRAWER) ---
            self._build_overlays(root)

        def _build_text_panel(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(14)

            # Context Bar (New / Append + 2-col)
            ctx = QFrame()
            ctx.setProperty("class", "ContextBar")
            ctx_lay = QVBoxLayout(ctx)
            ctx_lay.setContentsMargins(14, 10, 14, 10)
            ctx_lay.setSpacing(8)

            r1 = QHBoxLayout()
            r1.setSpacing(16)
            self.rb_new = QRadioButton("Создать новый DOCX")
            self.rb_append = QRadioButton("Дописать в существующий (Append)")
            self.rb_new.setChecked(True)
            self.rb_new.toggled.connect(self._on_append_radio_changed)
            r1.addWidget(self.rb_new)
            r1.addWidget(self.rb_append)
            r1.addStretch()

            self.btn_text_opts = QPushButton("⚙️ Опции вёрстки ▾")
            self.btn_text_opts.setProperty("class", "SecondaryBtn")
            self.btn_text_opts.setCursor(QCursor(Qt.PointingHandCursor))
            self.text_opts_menu = QMenu(self.btn_text_opts)
            self.act_two_col = self.text_opts_menu.addAction("2 колонки (IEEE / ВАК)")
            self.act_two_col.setCheckable(True)
            self.act_two_col.setChecked(bool(self.config.get("two_columns", False)))
            self.act_two_col.toggled.connect(self._on_two_col_toggled)

            self.act_strip_emphasis = self.text_opts_menu.addAction("Убрать жирный/курсив в тексте")
            self.act_strip_emphasis.setCheckable(True)
            self.act_strip_emphasis.setChecked(bool(self.config.get("strip_body_emphasis", False)))
            self.act_strip_emphasis.toggled.connect(self._on_strip_emphasis_toggled)

            self.text_opts_menu.addSeparator()
            self.act_keep_styles = self.text_opts_menu.addAction("Сохранить стили исходного файла (Append)")
            self.act_keep_styles.setCheckable(True)
            self.act_keep_styles.setChecked(True)
            self.act_keep_styles.toggled.connect(lambda c: self._toast("Стили исходного файла сохраняются" if c else "Стили исходного файла будут обновлены"))

            self.act_page_break = self.text_opts_menu.addAction("Начинать с новой страницы (Append)")
            self.act_page_break.setCheckable(True)
            self.act_page_break.setChecked(False)
            self.act_page_break.toggled.connect(lambda c: self._toast("Вставка с новой страницы" if c else "Вставка без разрыва страницы"))

            self.btn_text_opts.setMenu(self.text_opts_menu)
            r1.addWidget(self.btn_text_opts)
            ctx_lay.addLayout(r1)

            # Progressive disclosure row for Append
            self.append_opts_widget = QWidget()
            ao_lay = QHBoxLayout(self.append_opts_widget)
            ao_lay.setContentsMargins(0, 4, 0, 0)
            ao_lay.setSpacing(8)
            self.ed_append_target = QLineEdit(self.config.get("last_append_target", ""))
            self.ed_append_target.setPlaceholderText("Выберите целевой .docx файл...")
            install_styled_context_menu(self.ed_append_target, lambda: self.theme_name)
            ao_lay.addWidget(self.ed_append_target, 1)
            btn_pick = QPushButton("Выбрать .docx")
            btn_pick.setProperty("class", "SecondaryBtn")
            btn_pick.clicked.connect(self._pick_append_target)
            ao_lay.addWidget(btn_pick)
            self.append_opts_widget.setVisible(False)
            ctx_lay.addWidget(self.append_opts_widget)
            lay.addWidget(ctx)

            # Split Editor + Preview
            splitter = QSplitter(Qt.Horizontal)
            splitter.setHandleWidth(12)

            # Left Pane
            left_pane = QFrame()
            left_pane.setProperty("class", "PaneCard")
            lp_lay = QVBoxLayout(left_pane)
            lp_lay.setContentsMargins(0, 0, 0, 0)
            lp_lay.setSpacing(0)

            lp_hdr = QFrame()
            lp_hdr.setProperty("class", "PaneHeader")
            lph_lay = QHBoxLayout(lp_hdr)
            lph_lay.setContentsMargins(14, 9, 14, 9)
            lph_lay.setSpacing(6)
            lbl_l = QLabel("Исходный текст (Markdown + LaTeX)")
            lbl_l.setStyleSheet("font-size: 13px; font-weight: 600;")
            lph_lay.addWidget(lbl_l)
            lph_lay.addStretch()
            b_paste = QPushButton("Вставить")
            b_paste.setProperty("class", "SecondaryBtn")
            b_paste.setCursor(QCursor(Qt.PointingHandCursor))
            b_paste.clicked.connect(self._paste_clip)
            lph_lay.addWidget(b_paste)

            b_more = QPushButton("Действия ▾")
            b_more.setProperty("class", "SecondaryBtn")
            b_more.setCursor(QCursor(Qt.PointingHandCursor))
            more_menu = QMenu(b_more)
            more_menu.addAction("Открыть .md файл...").triggered.connect(self._open_md_file)
            more_menu.addAction("Загрузить пример ДГТУ").triggered.connect(self._load_sample)
            more_menu.addSeparator()
            more_menu.addAction("Очистить редактор").triggered.connect(self._clear_editor)
            b_more.setMenu(more_menu)
            lph_lay.addWidget(b_more)
            lp_lay.addWidget(lp_hdr)

            self.ed_input = QTextEdit()
            self.ed_input.setObjectName("MdEditor")
            self.ed_input.setAcceptDrops(False)
            install_styled_context_menu(self.ed_input, lambda: self.theme_name)
            self.ed_input.setPlaceholderText("Вставьте или введите Markdown, формулы LaTeX ($$...$$), таблицы или блоки mermaid/plantuml/dot...")
            self._preview_timer = QTimer(self)
            self._preview_timer.setSingleShot(True)
            self._preview_timer.setInterval(180)
            self._preview_timer.timeout.connect(self._parse_preview)
            self.ed_input.textChanged.connect(lambda: self._preview_timer.start())
            lp_lay.addWidget(self.ed_input, 1)

            lp_ftr = QFrame()
            lp_ftr.setProperty("class", "PaneFooter")
            lpf_lay = QHBoxLayout(lp_ftr)
            lpf_lay.setContentsMargins(14, 7, 14, 7)
            lbl_hint = QLabel("Совет: перетащите .md прямо в это окно")
            lbl_hint.setStyleSheet("font-size: 12px; opacity: 0.7;")
            lpf_lay.addWidget(lbl_hint)
            lpf_lay.addStretch()
            self.lbl_char_count = QLabel("0 симв.")
            self.lbl_char_count.setStyleSheet("font-size: 12px; opacity: 0.7;")
            lpf_lay.addWidget(self.lbl_char_count)
            lp_lay.addWidget(lp_ftr)
            splitter.addWidget(left_pane)

            # Right Pane
            right_pane = QFrame()
            right_pane.setProperty("class", "PaneCard")
            rp_lay = QVBoxLayout(right_pane)
            rp_lay.setContentsMargins(0, 0, 0, 0)
            rp_lay.setSpacing(0)

            rp_hdr = QFrame()
            rp_hdr.setProperty("class", "PaneHeader")
            rph_lay = QHBoxLayout(rp_hdr)
            rph_lay.setContentsMargins(14, 9, 14, 9)
            rph_lay.setSpacing(6)
            lbl_r = QLabel("Структура будущего DOCX")
            lbl_r.setStyleSheet("font-size: 13px; font-weight: 600;")
            rph_lay.addWidget(lbl_r)
            rph_lay.addStretch()
            self.lbl_stat_sec = QLabel("0 разд.")
            self.lbl_stat_sec.setProperty("class", "StatBadge")
            self.lbl_stat_frm = QLabel("0 ф-л")
            self.lbl_stat_frm.setProperty("class", "StatBadge")
            self.lbl_stat_tbl = QLabel("0 табл.")
            self.lbl_stat_tbl.setProperty("class", "StatBadge")
            rph_lay.addWidget(self.lbl_stat_sec)
            rph_lay.addWidget(self.lbl_stat_frm)
            rph_lay.addWidget(self.lbl_stat_tbl)
            rp_lay.addWidget(rp_hdr)

            self.preview_scroll = QScrollArea()
            self.preview_scroll.setObjectName("PreviewScroll")
            self.preview_scroll.setWidgetResizable(True)
            self.preview_body = QWidget()
            self.preview_body.setObjectName("PreviewBody")
            self.outline_lay = QVBoxLayout(self.preview_body)
            self.outline_lay.setContentsMargins(14, 14, 14, 14)
            self.outline_lay.setSpacing(6)
            self.outline_lay.addStretch()
            self.preview_scroll.setWidget(self.preview_body)
            rp_lay.addWidget(self.preview_scroll, 1)

            # Result Banner (Hidden until generated)
            self.result_banner = QFrame()
            self.result_banner.setObjectName("ResultBanner")
            rb_wrap = QVBoxLayout()
            rb_wrap.setContentsMargins(12, 0, 12, 12)
            rb_lay = QHBoxLayout(self.result_banner)
            rb_lay.setContentsMargins(12, 9, 12, 9)
            rb_lay.setSpacing(8)
            self.lbl_result_name = QLabel("✓ Готово: document.docx")
            self.lbl_result_name.setStyleSheet("font-size: 13px; font-weight: 600;")
            rb_lay.addWidget(self.lbl_result_name, 1)
            btn_ow = QPushButton("Открыть в Word")
            btn_ow.setProperty("class", "SecondaryBtn")
            btn_ow.clicked.connect(self._open_last_file)
            rb_lay.addWidget(btn_ow)
            btn_of = QPushButton("Папка")
            btn_of.setProperty("class", "SecondaryBtn")
            btn_of.clicked.connect(self._open_last_folder)
            rb_lay.addWidget(btn_of)
            self.result_banner.setVisible(False)
            rb_wrap.addWidget(self.result_banner)
            rp_lay.addLayout(rb_wrap)

            splitter.addWidget(right_pane)
            splitter.setSizes([580, 440])
            lay.addWidget(splitter, 1)
            return w

        def _build_files_panel(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(12)

            self.drop_hero = DropHeroWidget()
            self.drop_hero.files_dropped.connect(self._on_files_dropped)

            b_add_f = QPushButton("+ Файлы")
            b_add_f.setProperty("class", "SecondaryBtn")
            b_add_f.setCursor(QCursor(Qt.PointingHandCursor))
            b_add_f.clicked.connect(self._browse_files)
            self.drop_hero.btn_box.addWidget(b_add_f)

            b_queue_menu = QPushButton("Очередь ▾")
            b_queue_menu.setProperty("class", "SecondaryBtn")
            b_queue_menu.setCursor(QCursor(Qt.PointingHandCursor))
            q_menu = QMenu(b_queue_menu)
            q_menu.addAction("+ Добавить папку...").triggered.connect(self._browse_folder)
            q_menu.addSeparator()
            q_menu.addAction("Выбрать все файлы").triggered.connect(lambda: self._set_all_queue_selected(True))
            q_menu.addAction("Снять выделение").triggered.connect(lambda: self._set_all_queue_selected(False))
            q_menu.addAction("Удалить невыбранные").triggered.connect(self._remove_unselected_queue)
            q_menu.addSeparator()
            q_menu.addAction("Очистить очередь").triggered.connect(self._clear_queue)
            b_queue_menu.setMenu(q_menu)
            self.drop_hero.btn_box.addWidget(b_queue_menu)

            self.btn_batch_opts = QPushButton("⚙️ Опции пакета ▾")
            self.btn_batch_opts.setProperty("class", "SecondaryBtn")
            self.btn_batch_opts.setCursor(QCursor(Qt.PointingHandCursor))
            batch_menu = QMenu(self.btn_batch_opts)
            self.act_batch_2col = batch_menu.addAction("2 колонки (IEEE / ВАК)")
            self.act_batch_2col.setCheckable(True)
            self.act_batch_2col.setChecked(bool(self.config.get("two_columns", False)))
            self.act_batch_2col.toggled.connect(self._on_two_col_toggled)

            self.act_batch_strip = batch_menu.addAction("Убрать жирный/курсив в тексте")
            self.act_batch_strip.setCheckable(True)
            self.act_batch_strip.setChecked(bool(self.config.get("strip_body_emphasis", False)))
            self.act_batch_strip.toggled.connect(self._on_strip_emphasis_toggled)

            batch_menu.addSeparator()
            batch_menu.addAction("Назначить активный шаблон всем выбранным").triggered.connect(self._apply_active_tpl_to_selected)
            self.btn_batch_opts.setMenu(batch_menu)
            self.drop_hero.btn_box.addWidget(self.btn_batch_opts)

            lay.addWidget(self.drop_hero)

            self.progress_bar = QProgressBar()
            self.progress_bar.setMinimumHeight(22)
            self.progress_bar.setTextVisible(True)
            self.progress_bar.setFormat("%p% (%v/%m)")
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(False)
            self._progress_anim = QPropertyAnimation(self.progress_bar, b"value", self)
            self._progress_anim.setDuration(180)
            self._progress_anim.setEasingCurve(QEasingCurve.OutCubic)
            lay.addWidget(self.progress_bar)

            # Баннер результата пакетной обработки (Открыть файл / Папка)
            self.batch_result_banner = QFrame()
            self.batch_result_banner.setObjectName("ResultBanner")
            self.batch_result_banner.setAttribute(Qt.WA_StyledBackground, True)
            br_lay = QHBoxLayout(self.batch_result_banner)
            br_lay.setContentsMargins(12, 9, 12, 9)
            br_lay.setSpacing(8)
            self.lbl_batch_result = QLabel("✓ Конвертация завершена")
            self.lbl_batch_result.setStyleSheet("font-size: 13px; font-weight: 600;")
            br_lay.addWidget(self.lbl_batch_result, 1)
            btn_bo = QPushButton("Открыть в Word")
            btn_bo.setProperty("class", "SecondaryBtn")
            btn_bo.clicked.connect(self._open_last_file)
            br_lay.addWidget(btn_bo)
            btn_bf = QPushButton("Папка")
            btn_bf.setProperty("class", "SecondaryBtn")
            btn_bf.clicked.connect(self._open_last_folder)
            br_lay.addWidget(btn_bf)
            self.batch_result_banner.setVisible(False)
            lay.addWidget(self.batch_result_banner)

            tbl_wrap = QFrame()
            tbl_wrap.setObjectName("TableWrap")
            tbl_wrap.setAttribute(Qt.WA_StyledBackground, True)
            tw_lay = QVBoxLayout(tbl_wrap)
            tw_lay.setContentsMargins(2, 2, 2, 2)
            tw_lay.setSpacing(0)

            self.table_queue = QTableWidget(0, 7)
            self.table_queue.setFrameShape(QFrame.NoFrame)
            self.table_queue.setShowGrid(False)
            self.table_queue.verticalHeader().setVisible(False)
            self.table_queue.verticalHeader().setDefaultSectionSize(54)
            self.table_queue.setHorizontalHeaderLabels(["✓", "Файл", "Папка сохранения", "Размер", "Шаблон", "Статус", "Действия"])
            self.table_queue.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
            self.table_queue.setColumnWidth(0, 38)
            self.table_queue.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
            self.table_queue.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
            self.table_queue.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
            self.table_queue.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
            self.table_queue.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
            self.table_queue.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
            self.table_queue.setColumnWidth(6, 108)
            self.table_queue.setFocusPolicy(Qt.NoFocus)
            self.table_queue.setSelectionMode(QAbstractItemView.SingleSelection)
            self.table_queue.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.table_queue.setContextMenuPolicy(Qt.CustomContextMenu)
            self.table_queue.customContextMenuRequested.connect(self._on_queue_context_menu)
            tw_lay.addWidget(self.table_queue)
            lay.addWidget(tbl_wrap, 1)
            return w

        def _build_merge_panel(self) -> QWidget:
            w = QWidget()
            lay = QHBoxLayout(w)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(14)

            # Left: Merge DOCX
            c1 = QFrame()
            c1.setProperty("class", "SurfaceCard")
            c1.setMinimumWidth(0)
            l1 = QVBoxLayout(c1)
            l1.setContentsMargins(16, 16, 16, 16)
            l1.setSpacing(10)
            lbl1 = QLabel("Склейка нескольких DOCX в один")
            lbl1.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            l1.addWidget(lbl1)
            lbl1_s = QLabel("Перетаскивайте файлы мышью для изменения порядка (ПКМ — удалить).")
            lbl1_s.setWordWrap(True)
            lbl1_s.setStyleSheet("font-size: 12.5px; opacity: 0.7;")
            l1.addWidget(lbl1_s)

            self.list_merge = QListWidget()
            self.list_merge.setMinimumWidth(0)
            self.list_merge.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
            self.list_merge.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.list_merge.setTextElideMode(Qt.ElideMiddle)
            self.list_merge.setDragDropMode(QAbstractItemView.InternalMove)
            self.list_merge.setDefaultDropAction(Qt.MoveAction)
            self.list_merge.model().rowsMoved.connect(self._on_merge_rows_moved)
            self.list_merge.setContextMenuPolicy(Qt.CustomContextMenu)
            self.list_merge.customContextMenuRequested.connect(self._on_merge_context_menu)
            l1.addWidget(self.list_merge, 1)

            m_bar = QHBoxLayout()
            m_bar.setSpacing(6)
            b_add_m = QPushButton("+ Добавить .docx")
            b_add_m.setProperty("class", "SecondaryBtn")
            b_add_m.setCursor(QCursor(Qt.PointingHandCursor))
            b_add_m.clicked.connect(self._merge_add)
            m_bar.addWidget(b_add_m)

            m_bar.addStretch()

            b_m_opts = QPushButton("⚙️ Опции склейки ▾")
            b_m_opts.setProperty("class", "SecondaryBtn")
            b_m_opts.setCursor(QCursor(Qt.PointingHandCursor))
            m_opts_menu = QMenu(b_m_opts)
            self.act_merge_pb = m_opts_menu.addAction("Разрыв страницы между файлами")
            self.act_merge_pb.setCheckable(True)
            self.act_merge_pb.setChecked(True)
            self.act_merge_apply_tpl = m_opts_menu.addAction("Применить активный шаблон после склейки")
            self.act_merge_apply_tpl.setCheckable(True)
            self.act_merge_apply_tpl.setChecked(bool(self.config.get("merge_apply_template", False)))
            self.act_merge_apply_tpl.toggled.connect(self._on_merge_apply_toggled)
            m_opts_menu.addSeparator()
            self.act_merge_2col = m_opts_menu.addAction("2 колонки (IEEE / ВАК)")
            self.act_merge_2col.setCheckable(True)
            self.act_merge_2col.setChecked(bool(self.config.get("two_columns", False)))
            self.act_merge_2col.toggled.connect(self._on_two_col_toggled)
            self.act_merge_strip = m_opts_menu.addAction("Убрать жирный/курсив в тексте")
            self.act_merge_strip.setCheckable(True)
            self.act_merge_strip.setChecked(bool(self.config.get("strip_body_emphasis", False)))
            self.act_merge_strip.toggled.connect(self._on_strip_emphasis_toggled)
            b_m_opts.setMenu(m_opts_menu)
            m_bar.addWidget(b_m_opts)

            l1.addLayout(m_bar)
            lay.addWidget(c1, 1)

            # Right: Restyle DOCX
            c2 = QFrame()
            c2.setProperty("class", "SurfaceCard")
            l2 = QVBoxLayout(c2)
            l2.setContentsMargins(16, 16, 16, 16)
            l2.setSpacing(12)
            lbl2 = QLabel("Применить шаблон к готовому DOCX")
            lbl2.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            l2.addWidget(lbl2)
            lbl2_s = QLabel("Обновляет шрифты, интервалы, колонки и стандартные стили (включая «Обычный») без изменения текста, с учётом активных опций вёрстки.")
            lbl2_s.setWordWrap(True)
            lbl2_s.setStyleSheet("font-size: 12.5px; opacity: 0.7;")
            l2.addWidget(lbl2_s)

            r_row = QHBoxLayout()
            self.ed_restyle_target = QLineEdit()
            self.ed_restyle_target.setPlaceholderText("Выберите .docx файл...")
            install_styled_context_menu(self.ed_restyle_target, lambda: self.theme_name)
            r_row.addWidget(self.ed_restyle_target, 1)
            btn_rb = QPushButton("Обзор...")
            btn_rb.setProperty("class", "SecondaryBtn")
            btn_rb.clicked.connect(self._pick_restyle_target)
            r_row.addWidget(btn_rb)
            l2.addLayout(r_row)
            l2.addStretch()

            r_foot = QHBoxLayout()
            r_foot.addStretch()
            btn_apply = QPushButton("Обновить стили документа")
            btn_apply.setProperty("class", "SecondaryBtn")
            btn_apply.setCursor(QCursor(Qt.PointingHandCursor))
            btn_apply.clicked.connect(self._run_restyle)
            r_foot.addWidget(btn_apply)
            l2.addLayout(r_foot)
            lay.addWidget(c2, 1)
            return w

        def _build_templates_panel(self) -> QWidget:
            self._tpl_editor_host = QWidget()
            self._tpl_editor_host_lay = QVBoxLayout(self._tpl_editor_host)
            self._tpl_editor_host_lay.setContentsMargins(0, 0, 0, 0)
            self.template_editor = None
            return self._tpl_editor_host

        def _ensure_template_editor(self):
            if self.template_editor is None:
                self.template_editor = TemplateEditorWidget(
                    templates_provider=lambda: self.templates_cache,
                    primary_dir_provider=get_primary_templates_dir,
                    register_path_fn=register_custom_template_path,
                    unregister_path_fn=unregister_template_path,
                    context_menu_installer=lambda w: install_styled_context_menu(w, lambda: self.theme_name),
                    theme_getter=lambda: THEMES.get(self.theme_name, THEMES["light"]),
                    parent=self._tpl_editor_host,
                )
                self.template_editor.template_saved.connect(self._on_editor_template_saved)
                self.template_editor.template_deleted.connect(lambda _: self.reload_templates(silent=True))
                self.template_editor.refresh_requested.connect(lambda: self.reload_templates(silent=False))
                self.template_editor.toast_requested.connect(self._toast)
                self._tpl_editor_host_lay.addWidget(self.template_editor)
                self.template_editor.apply_theme()
            return self.template_editor

        def _build_overlays(self, parent_widget: QWidget):
            self.backdrop = QFrame(parent_widget)
            self.backdrop.setStyleSheet("background-color: rgba(15, 23, 42, 0.32); border-radius: 13px;")
            self.backdrop.setVisible(False)
            self.backdrop.mousePressEvent = lambda e: self._close_overlays()

            # Template Popover
            self.tpl_popover = QFrame(parent_widget)
            self.tpl_popover.setObjectName("OverlayPanel")
            self.tpl_popover.setAttribute(Qt.WA_StyledBackground, True)
            self.tpl_popover.setVisible(False)


            p_lay = QVBoxLayout(self.tpl_popover)
            p_lay.setContentsMargins(2, 2, 2, 4)
            p_lay.setSpacing(0)

            p_hdr = QFrame()
            p_hdr.setProperty("class", "PaneHeader")
            ph_lay = QHBoxLayout(p_hdr)
            ph_lay.setContentsMargins(14, 10, 14, 10)
            lbl_ph = QLabel("Библиотека шаблонов ГОСТ")
            lbl_ph.setStyleSheet("font-size: 13.5px; font-weight: 600;")
            ph_lay.addWidget(lbl_ph)
            ph_lay.addStretch()
            self.btn_pop_edit = QPushButton("")
            self.btn_pop_edit.setProperty("class", "IconBtn")
            self.btn_pop_edit.setFixedSize(30, 30)
            self.btn_pop_edit.setIconSize(QSize(15, 15))
            self.btn_pop_edit.setToolTip("Изменить в редакторе шаблонов")
            self.btn_pop_edit.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_pop_edit.clicked.connect(self._open_template_editor_from_popover)
            ph_lay.addWidget(self.btn_pop_edit)

            self.btn_pop_imp = QPushButton("")
            self.btn_pop_imp.setProperty("class", "IconBtn")
            self.btn_pop_imp.setFixedSize(30, 30)
            self.btn_pop_imp.setIconSize(QSize(15, 15))
            self.btn_pop_imp.setToolTip("Импортировать шаблон (.json)")
            self.btn_pop_imp.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_pop_imp.clicked.connect(self._import_template)
            ph_lay.addWidget(self.btn_pop_imp)

            self.btn_pop_ref = QPushButton("")
            self.btn_pop_ref.setProperty("class", "IconBtn")
            self.btn_pop_ref.setFixedSize(30, 30)
            self.btn_pop_ref.setIconSize(QSize(15, 15))
            self.btn_pop_ref.setToolTip("Обновить список шаблонов")
            self.btn_pop_ref.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_pop_ref.clicked.connect(lambda: self.reload_templates(silent=False))
            ph_lay.addWidget(self.btn_pop_ref)
            btn_pc = QPushButton("✕")
            btn_pc.setProperty("class", "WinBtn")
            btn_pc.setFixedSize(30, 30)
            btn_pc.setCursor(QCursor(Qt.PointingHandCursor))
            btn_pc.clicked.connect(self._close_overlays)
            ph_lay.addWidget(btn_pc)
            p_lay.addWidget(p_hdr)

            self.tpl_scroll = QScrollArea()
            self.tpl_scroll.setObjectName("PopScroll")
            self.tpl_scroll.setWidgetResizable(True)
            self.tpl_list_widget = QWidget()
            self.tpl_list_lay = QVBoxLayout(self.tpl_list_widget)
            self.tpl_list_lay.setContentsMargins(10, 10, 10, 10)
            self.tpl_list_lay.setSpacing(6)
            self.tpl_scroll.setWidget(self.tpl_list_widget)
            p_lay.addWidget(self.tpl_scroll, 1)

            # Settings Slide-over Drawer
            self.settings_drawer = QFrame(parent_widget)
            self.settings_drawer.setObjectName("DrawerPanel")
            self.settings_drawer.setAttribute(Qt.WA_StyledBackground, True)
            self.settings_drawer.setVisible(False)


            d_lay = QVBoxLayout(self.settings_drawer)
            d_lay.setContentsMargins(0, 0, 0, 0)
            d_lay.setSpacing(0)

            d_hdr = QFrame()
            d_hdr.setProperty("class", "PaneHeader")
            dh_lay = QHBoxLayout(d_hdr)
            dh_lay.setContentsMargins(16, 12, 16, 12)
            lbl_dh = QLabel("Параметры приложения")
            lbl_dh.setStyleSheet("font-size: 13.5px; font-weight: 600;")
            dh_lay.addWidget(lbl_dh)
            dh_lay.addStretch()
            btn_dc = QPushButton("✕")
            btn_dc.setProperty("class", "WinBtn")
            btn_dc.setFixedSize(30, 30)
            btn_dc.setCursor(QCursor(Qt.PointingHandCursor))
            btn_dc.clicked.connect(self._close_overlays)
            dh_lay.addWidget(btn_dc)
            d_lay.addWidget(d_hdr)

            d_body = QVBoxLayout()
            d_body.setContentsMargins(16, 16, 16, 16)
            d_body.setSpacing(14)

            lbl_g1 = QLabel("Поведение после сборки")
            lbl_g1.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.65;")
            d_body.addWidget(lbl_g1)
            self.chk_auto_open = QCheckBox("Автоматически открывать документ в Word")
            self.chk_auto_open.setChecked(self.config.get("open_after_create", True))
            self.chk_auto_open.toggled.connect(self._save_drawer_settings)
            d_body.addWidget(self.chk_auto_open)

            self.chk_autostart = QCheckBox("Запускать при старте Windows (в трей)")
            self.chk_autostart.setChecked(is_windows_autostart_enabled())
            self.chk_autostart.toggled.connect(self._on_autostart_toggled)
            d_body.addWidget(self.chk_autostart)

            lbl_gt = QLabel("Тема оформления")
            lbl_gt.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.65; margin-top: 8px;")
            d_body.addWidget(lbl_gt)

            theme_seg = QFrame()
            theme_seg.setObjectName("SegmentedBox")
            ts_lay = QHBoxLayout(theme_seg)
            ts_lay.setContentsMargins(3, 3, 3, 3)
            ts_lay.setSpacing(2)
            self.btn_theme_light = QPushButton("☀️ Светлая")
            self.btn_theme_light.setProperty("class", "SegBtn")
            self.btn_theme_light.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_theme_light.clicked.connect(lambda: self._set_theme("light"))
            self.btn_theme_dark = QPushButton("🌙 Тёмная")
            self.btn_theme_dark.setProperty("class", "SegBtn")
            self.btn_theme_dark.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_theme_dark.clicked.connect(lambda: self._set_theme("dark"))
            ts_lay.addWidget(self.btn_theme_light, 1)
            ts_lay.addWidget(self.btn_theme_dark, 1)
            d_body.addWidget(theme_seg)

            lbl_gf = QLabel("Шрифт интерфейса приложения")
            lbl_gf.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.65; margin-top: 8px;")
            d_body.addWidget(lbl_gf)
            self.combo_ui_font = QComboBox()
            saved_ui_f = self.config.get("ui_font", "Inter")
            sel_ui_idx = 0
            for idx_f, (f_lbl, f_val) in enumerate(UI_CYRILLIC_FONTS):
                self.combo_ui_font.addItem(f_lbl, f_val)
                if f_val == saved_ui_f:
                    sel_ui_idx = idx_f
            self.combo_ui_font.setCurrentIndex(sel_ui_idx)
            self.combo_ui_font.currentIndexChanged.connect(self._on_ui_font_changed)
            d_body.addWidget(self.combo_ui_font)

            lbl_g2 = QLabel("Глобальные горячие клавиши")
            lbl_g2.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.65; margin-top: 8px;")
            d_body.addWidget(lbl_g2)
            lbl_hk1 = QLabel("Ctrl+Alt+W — мгновенная сборка из буфера")
            lbl_hk2 = QLabel("Ctrl+Alt+B — развернуть главное окно")
            d_body.addWidget(lbl_hk1)
            d_body.addWidget(lbl_hk2)

            lbl_g3 = QLabel("Папка сохранения (при ручном выборе)")
            lbl_g3.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.65; margin-top: 8px;")
            d_body.addWidget(lbl_g3)
            self.ed_out_dir = QLineEdit(self.config.get("output_dir", ""))
            install_styled_context_menu(self.ed_out_dir, lambda: self.theme_name)
            d_body.addWidget(self.ed_out_dir)
            btn_od = QPushButton("Выбрать папку...")
            btn_od.setProperty("class", "SecondaryBtn")
            btn_od.clicked.connect(self._pick_custom_out_dir)
            d_body.addWidget(btn_od)

            d_body.addStretch()
            lbl_tray = QLabel("Закрытие окна сворачивает программу в системный трей, сохраняя работу горячих клавиш.")
            lbl_tray.setWordWrap(True)
            lbl_tray.setStyleSheet("font-size: 12px; opacity: 0.65;")
            d_body.addWidget(lbl_tray)
            d_lay.addLayout(d_body, 1)

            self._drawer_anim = QPropertyAnimation(self.settings_drawer, b"geometry", self)
            self._drawer_anim.setDuration(220)
            self._drawer_anim.setEasingCurve(QEasingCurve.OutCubic)

            self._pop_anim = QPropertyAnimation(self.tpl_popover, b"geometry", self)
            self._pop_anim.setDuration(180)
            self._pop_anim.setEasingCurve(QEasingCurve.OutCubic)

        def _layout_overlays(self):
            if not hasattr(self, "backdrop"):
                return
            rw = self.window_frame.width()
            rh = self.window_frame.height()
            self.backdrop.setGeometry(0, 0, rw, rh)
            pw = min(440, max(320, rw - 36))
            ph = min(460, max(280, rh - 100))
            self.tpl_popover.setGeometry(rw - pw - 18, 60, pw, ph)
            if self.settings_drawer.isVisible():
                dw = min(360, rw)
                self.settings_drawer.setGeometry(rw - dw, 0, dw, rh)

        def _close_overlays(self):
            self.backdrop.setVisible(False)
            if self.tpl_popover.isVisible():
                self.tpl_popover.setVisible(False)
            if self.settings_drawer.isVisible():
                rw = self.window_frame.width()
                rh = self.window_frame.height()
                dw = self.settings_drawer.width()
                self._drawer_anim.stop()
                self._drawer_anim.setStartValue(self.settings_drawer.geometry())
                self._drawer_anim.setEndValue(QRect(rw, 0, dw, rh))
                def _hide_dr():
                    self.settings_drawer.setVisible(False)
                    try:
                        self._drawer_anim.finished.disconnect(_hide_dr)
                    except Exception:
                        pass
                self._drawer_anim.finished.connect(_hide_dr)
                self._drawer_anim.start()

        def _toggle_tpl_popover(self):
            if self.tpl_popover.isVisible():
                self._close_overlays()
                return
            self._close_overlays()
            self.reload_templates(silent=True)
            self._layout_overlays()
            rw = self.window_frame.width()
            rh = self.window_frame.height()
            pw = min(440, max(340, rw - 36))
            ph = min(460, max(300, rh - 100))
            px = rw - pw - 18
            py = 60
            self.backdrop.setVisible(True)
            self.backdrop.raise_()
            self.tpl_popover.setGeometry(px, py - 12, pw, 0)
            self.tpl_popover.setVisible(True)
            self.tpl_popover.raise_()
            self._pop_anim.stop()
            self._pop_anim.setStartValue(QRect(px, py - 10, pw, 40))
            self._pop_anim.setEndValue(QRect(px, py, pw, ph))
            self._pop_anim.start()

        def _open_settings_drawer(self):
            self._close_overlays()
            self._layout_overlays()
            rw = self.window_frame.width()
            rh = self.window_frame.height()
            dw = min(360, rw)
            self.backdrop.setVisible(True)
            self.backdrop.raise_()
            self.settings_drawer.setGeometry(rw, 0, dw, rh)
            self.settings_drawer.setVisible(True)
            self.settings_drawer.raise_()
            self._drawer_anim.stop()
            self._drawer_anim.setStartValue(QRect(rw, 0, dw, rh))
            self._drawer_anim.setEndValue(QRect(rw - dw, 0, dw, rh))
            self._drawer_anim.start()

        def _render_tpl_cards(self):
            self.tpl_scroll.setUpdatesEnabled(False)
            try:
                self._render_tpl_cards_inner()
            finally:
                self.tpl_scroll.setUpdatesEnabled(True)

        def _render_tpl_cards_inner(self):
            while self.tpl_list_lay.count():
                it = self.tpl_list_lay.takeAt(0)
                w_old = it.widget()
                if w_old is not None:
                    w_old.setParent(None)
                    w_old.deleteLater()
            t = THEMES.get(self.theme_name, THEMES["light"])
            self.tpl_list_widget.setStyleSheet("background: transparent;")
            for k, v in self.templates_cache.items():
                title = v.get("display_title", k)
                desc = v.get("specs_summary") or v.get("description", "")
                is_sel = (k == self.active_tpl)

                card = QFrame()
                card.setProperty("class", "TplCard")
                card.setProperty("selected", "true" if is_sel else "false")
                card.setAttribute(Qt.WA_StyledBackground, True)
                card.setCursor(QCursor(Qt.PointingHandCursor))
                card.mousePressEvent = lambda e, tid=k: self._select_template(tid)

                cl = QVBoxLayout(card)
                cl.setContentsMargins(12, 10, 12, 10)
                cl.setSpacing(3)

                top_r = QHBoxLayout()
                lbl_t = QLabel(title)
                lbl_t.setStyleSheet(f"color: {t['text']}; font-size: 13px; font-weight: 600; background: transparent;")
                top_r.addWidget(lbl_t, 1)
                if is_sel:
                    lbl_a = QLabel("Активен")
                    lbl_a.setStyleSheet(f"color: {t['accent_text']}; font-size: 11.5px; font-weight: 600; background: transparent;")
                    top_r.addWidget(lbl_a)
                cl.addLayout(top_r)

                if desc:
                    lbl_d = QLabel(desc)
                    lbl_d.setWordWrap(True)
                    lbl_d.setStyleSheet(f"color: {t['text_mute']}; font-size: 12px; background: transparent;")
                    cl.addWidget(lbl_d)

                self.tpl_list_lay.addWidget(card)
            self.tpl_list_lay.addStretch()

        def _select_template(self, tpl_id: str):
            self.active_tpl = tpl_id
            self.config["default_template"] = tpl_id
            save_config(self.config)
            self._sync_template_ui()
            self._close_overlays()
            self._toast(f"Выбран шаблон: {self.templates_cache[tpl_id].get('display_title', tpl_id)}")

        def _open_template_editor_from_popover(self):
            self._close_overlays()
            self._switch_mode("templates")
            if getattr(self, "template_editor", None) is not None and not getattr(self.template_editor, "_dirty", False):
                self.template_editor.load_templates(self.active_tpl)

        def _on_editor_template_saved(self, tpl_id: str):
            self.active_tpl = tpl_id
            self.config["default_template"] = tpl_id
            save_config(self.config)
            self.reload_templates(silent=True)

        def _sync_template_ui(self):
            tpl = self.templates_cache.get(self.active_tpl, {})
            t_title = tpl.get("display_title", self.active_tpl)
            if hasattr(self, "lbl_hdr_tpl_text"):
                elided_t = self.lbl_hdr_tpl_text.fontMetrics().elidedText(f"Шаблон: {t_title}", Qt.ElideRight, 206)
                self.lbl_hdr_tpl_text.setText(elided_t)
            self.btn_tpl_open.setToolTip(f"Активный шаблон: {t_title}")
            cols = int(tpl.get("page", {}).get("columns", 1) or 1)
            if cols > 1 and not self.config.get("two_columns", False):
                self._sync_two_col_actions(True)
            self._render_tpl_cards()
            self._refresh_queue_table()

        def _setup_template_watcher(self):
            self._last_tpl_sig = get_templates_signature()
            try:
                self._tpl_watcher = QFileSystemWatcher(self)
                self._refresh_watcher_paths()
                self._tpl_reload_timer = QTimer(self)
                self._tpl_reload_timer.setSingleShot(True)
                self._tpl_reload_timer.setInterval(200)
                self._tpl_reload_timer.timeout.connect(self._check_templates_changed)
                self._tpl_watcher.directoryChanged.connect(lambda _: self._tpl_reload_timer.start())
                self._tpl_watcher.fileChanged.connect(lambda _: self._tpl_reload_timer.start())
            except Exception as ex:
                log(f"Ошибка QFileSystemWatcher: {ex}")

            # Надёжный лёгкий опрос (1.5 сек) на случай облачной синхронизации (Google Drive) или переименований
            self._tpl_poll_timer = QTimer(self)
            self._tpl_poll_timer.setInterval(1500)
            self._tpl_poll_timer.timeout.connect(self._check_templates_changed)
            self._tpl_poll_timer.start()

        def _refresh_watcher_paths(self):
            if not hasattr(self, "_tpl_watcher"):
                return
            try:
                dirs = [str(d) for d in get_templates_dirs() if d.exists()]
                files = [fp for fp, _, _ in get_templates_signature()]
                cur_dirs = set(self._tpl_watcher.directories())
                cur_files = set(self._tpl_watcher.files())
                add_d = [d for d in dirs if d not in cur_dirs]
                add_f = [f for f in files if f not in cur_files]
                if add_d:
                    self._tpl_watcher.addPaths(add_d)
                if add_f:
                    self._tpl_watcher.addPaths(add_f)
            except Exception:
                pass

        def _check_templates_changed(self):
            new_sig = get_templates_signature()
            if new_sig != getattr(self, "_last_tpl_sig", None):
                self._last_tpl_sig = new_sig
                self._refresh_watcher_paths()
                self.reload_templates(silent=True)

        def reload_templates(self, silent: bool = False):
            self._last_tpl_sig = get_templates_signature()
            self._refresh_watcher_paths()
            fresh = list_available_templates()
            self.templates_cache = {t["name"]: t for t in fresh}
            if self.active_tpl not in self.templates_cache and self.templates_cache:
                self.active_tpl = next(iter(self.templates_cache))
            self._sync_template_ui()
            if getattr(self, "template_editor", None) is not None:
                self.template_editor.load_templates(self.active_tpl)
            if not silent:
                self._toast(f"Найдено шаблонов: {len(self.templates_cache)}")

        def _import_template(self):
            files, _ = QFileDialog.getOpenFileNames(self, "Импорт JSON-шаблона", "", "JSON (*.json)")
            if not files:
                return
            dest_dir = get_primary_templates_dir()
            cnt = 0
            for f_str in files:
                src = Path(f_str)
                try:
                    data = json.loads(src.read_text(encoding="utf-8-sig"))
                    if isinstance(data, dict) and "styles" in data:
                        dst = dest_dir / src.name
                        shutil.copy2(src, dst)
                        register_custom_template_path(dst)
                        cnt += 1
                except Exception:
                    pass
            self.reload_templates(silent=True)
            self._toast(f"Импортировано шаблонов: {cnt}" if cnt else "Ошибка: некорректный JSON шаблона")

        def _switch_mode(self, mode_id: str):
            self.mode = mode_id
            idx_map = {"text": 0, "files": 1, "merge": 2, "templates": 3}
            self.stack.setCurrentIndex(idx_map.get(mode_id, 0))
            for k, btn in self.seg_buttons.items():
                btn.setProperty("active", "true" if k == mode_id else "false")
                btn.style().unpolish(btn)
                btn.style().polish(btn)
            if hasattr(self, "btn_tpl_open"):
                self.btn_tpl_open.setVisible(mode_id != "templates")
            if mode_id in ("files", "templates"):
                self.reload_templates(silent=True)
            if mode_id == "templates":
                self._ensure_template_editor().load_templates(self.active_tpl)
            self._update_footer_button()

        def _on_append_radio_changed(self):
            is_app = self.rb_append.isChecked()
            self.append_opts_widget.setVisible(is_app)
            self._update_footer_button()

        def _update_footer_button(self):
            if self.mode == "text":
                self.btn_main_run.setText("Дописать в DOCX →" if self.rb_append.isChecked() else "Сформировать DOCX →")
            elif self.mode == "files":
                sel_cnt = sum(1 for q in self.queue_items if q.get("selected", True))
                self.btn_main_run.setText(f"Конвертировать выбранные ({sel_cnt}) →" if sel_cnt else "Конвертировать выбранные →")
            elif self.mode == "merge":
                self.btn_main_run.setText("Склеить DOCX →")
            else:
                self.btn_main_run.setText("Сохранить шаблон →")

        def _toggle_save_mode(self):
            cur = self.config.get("save_mode", "same_dir")
            if cur == "same_dir":
                d = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения DOCX", self.config.get("output_dir", ""))
                if d:
                    self.config["save_mode"] = "custom_dir"
                    self.config["output_dir"] = d
                    save_config(self.config)
            else:
                self.config["save_mode"] = "same_dir"
                save_config(self.config)
            self._sync_save_mode_label()
            self._refresh_queue_table()

        def _sync_save_mode_label(self):
            if self.config.get("save_mode", "same_dir") == "same_dir":
                self.btn_save_mode.setText("Сохранять: Рядом с исходным файлом  ✏️")
            else:
                out_d = Path(self.config.get("output_dir", "")).name or self.config.get("output_dir", "")
                self.btn_save_mode.setText(f"Сохранять: {out_d}  ✏️")

        def _save_drawer_settings(self):
            self.config["open_after_create"] = self.chk_auto_open.isChecked()
            save_config(self.config)
            self._toast("Настройки сохранены")

        def _on_autostart_toggled(self, checked: bool):
            ok, err = set_windows_autostart(bool(checked))
            if ok:
                self.config["autostart"] = bool(checked)
                save_config(self.config)
                self._toast("Автозагрузка Windows включена" if checked else "Автозагрузка Windows отключена")
            else:
                self.chk_autostart.blockSignals(True)
                self.chk_autostart.setChecked(not checked)
                self.chk_autostart.blockSignals(False)
                self._toast(f"Не удалось изменить автозагрузку: {err[:45]}")

        def _pick_custom_out_dir(self):
            d = QFileDialog.getExistingDirectory(self, "Выберите папку сохранения", self.ed_out_dir.text())
            if d:
                self.ed_out_dir.setText(d)
                self.config["output_dir"] = d
                self.config["save_mode"] = "custom_dir"
                save_config(self.config)
                self._sync_save_mode_label()

        def _toast(self, msg: str):
            if self.app_controller and hasattr(self.app_controller, "toast"):
                self.app_controller.toast.show_toast(msg, self.theme_name)

        # --- Text / MD Panel actions ---
        def _load_sample(self):
            self._qt_base_dir = None
            sample_md = (
                "# 1 Постановка задачи\n\n"
                "Проектирование автоматизированного конвертера документации по ГОСТ 7.32-2017.\n\n"
                "## 1.1 Расчётная часть\n\n"
                "$$E = m c^2 + \\int_0^T P(t)\\,dt \\tag{1}$$\n\n"
                "| Параметр | Обозначение | Значение |\n"
                "| :--- | :---: | ---: |\n"
                "| Частота опроса | $f_s$ | 1000 Гц |\n"
                "| Погрешность | $\\Delta$ | 0,05 % |\n"
            )
            self.ed_input.setPlainText(sample_md)
            self._toast("Загружен образец ДГТУ")

        def _paste_clip(self):
            txt = ""
            if pyperclip:
                try:
                    txt = pyperclip.paste()
                except Exception:
                    pass
            if not txt and QApplication.clipboard():
                txt = QApplication.clipboard().text()
            if txt:
                self._qt_base_dir = None
                self.ed_input.setPlainText(txt)
                self._toast("Текст вставлен из буфера")

        def _open_md_file(self):
            f, _ = QFileDialog.getOpenFileName(self, "Открыть Markdown", "", "Markdown (*.md *.markdown *.txt)")
            if f:
                fp = Path(f)
                self._qt_base_dir = str(fp.resolve().parent)
                self.ed_input.setPlainText(fp.read_text(encoding="utf-8-sig"))

        def _clear_editor(self):
            self.ed_input.clear()
            self.ed_input.document().clearUndoRedoStacks()
            self.result_banner.setVisible(False)
            self._parse_preview()
            gc.collect()

        def _pick_append_target(self):
            f, _ = QFileDialog.getOpenFileName(self, "Выберите целевой DOCX", "", "Word Documents (*.docx)")
            if f:
                self.ed_append_target.setText(f)
                self.config["last_append_target"] = f
                save_config(self.config)

        def _parse_preview(self):
            raw = self.ed_input.toPlainText()
            cur_key = (raw, self.theme_name, self.active_tpl)
            if getattr(self, "_last_preview_key", None) == cur_key:
                return
            self._last_preview_key = cur_key
            self.preview_scroll.setUpdatesEnabled(False)
            try:
                self._parse_preview_inner(raw)
            finally:
                self.preview_scroll.setUpdatesEnabled(True)

        def _parse_preview_inner(self, raw: str):
            self.lbl_char_count.setText(f"{len(raw)} симв.")
            st = analyze_markdown_text(raw)
            self.lbl_stat_sec.setText(f"{st['headings']} разд.")
            self.lbl_stat_frm.setText(f"{st['formulas']} ф-л")
            self.lbl_stat_tbl.setText(f"{st['tables']} табл.")

            while self.outline_lay.count():
                it = self.outline_lay.takeAt(0)
                w_old = it.widget()
                if w_old is not None:
                    w_old.setParent(None)
                    w_old.deleteLater()

            if not raw.strip():
                lbl_e = QLabel("Введите или вставьте текст слева, чтобы увидеть структуру документа.")
                lbl_e.setWordWrap(True)
                lbl_e.setStyleSheet("opacity: 0.65; padding: 10px;")
                self.outline_lay.addWidget(lbl_e)
                self.outline_lay.addStretch()
                return

            if HAS_ENGINE:
                try:
                    tpl_cfg = self.templates_cache.get(self.active_tpl, {})
                    nodes = MarkdownASTParser(tpl_cfg.get("structural_titles", [])).parse(raw)
                    for n in nodes[:40]:
                        nt = n["type"]
                        if nt == "heading":
                            tag = f"H{n.get('outline_level', 0) + 1}"
                            title = n.get("text", "")
                        elif nt == "table":
                            tag = "Табл"
                            title = f"Таблица ({len(n.get('header', []))} стлб. × {len(n.get('rows', []))} строк)"
                        elif nt == "formula_block":
                            tag = "Ф-ла"
                            title = f"{n.get('number') or ''} {n.get('latex', '')[:45]}"
                        elif nt == "list_item":
                            tag = n.get("raw_marker", "•")
                            title = n.get("text", "")[:60]
                        elif nt in ("figure_image", "rendered_diagram"):
                            tag = "Рис"
                            title = n.get("path") or n.get("lang", "Диаграмма")
                        else:
                            continue
                        t = THEMES.get(self.theme_name, THEMES["light"])
                        row = QFrame()
                        row.setProperty("class", "OutlineItem")
                        row.setAttribute(Qt.WA_StyledBackground, True)
                        row.setMinimumHeight(34)
                        row.setStyleSheet(f"background-color: {t['bg_window']}; border: 1px solid {t['border']}; border-radius: 7px;")
                        rl = QHBoxLayout(row)
                        rl.setContentsMargins(10, 6, 10, 6)
                        rl.setSpacing(8)
                        l_tag = QLabel(tag)
                        l_tag.setStyleSheet(f"background-color: {t['accent_soft']}; color: {t['accent_text']}; border-radius: 4px; padding: 2px 6px; font-family: 'Consolas', monospace; font-size: 11px; font-weight: 600;")
                        rl.addWidget(l_tag)
                        l_txt = QLabel(title)
                        l_txt.setStyleSheet(f"color: {t['text']}; font-size: 13px; background: transparent; border: none;")
                        rl.addWidget(l_txt, 1)
                        self.outline_lay.addWidget(row)
                except Exception:
                    pass
            self.outline_lay.addStretch()

        def _sync_two_col_actions(self, checked: bool):
            for act_name in ("act_two_col", "act_batch_2col", "act_merge_2col"):
                act = getattr(self, act_name, None)
                if act is not None and act.isChecked() != checked:
                    act.blockSignals(True)
                    act.setChecked(checked)
                    act.blockSignals(False)
            self._update_opts_badges()

        def _sync_strip_actions(self, checked: bool):
            for act_name in ("act_strip_emphasis", "act_batch_strip", "act_merge_strip"):
                act = getattr(self, act_name, None)
                if act is not None and act.isChecked() != checked:
                    act.blockSignals(True)
                    act.setChecked(checked)
                    act.blockSignals(False)
            self._update_opts_badges()

        def _update_opts_badges(self):
            parts = []
            if self.config.get("two_columns", False):
                parts.append("2 кол.")
            if self.config.get("strip_body_emphasis", False):
                parts.append("Без жирного")
            suffix = f" ({ ' · '.join(parts) })" if parts else ""
            if hasattr(self, "btn_text_opts"):
                self.btn_text_opts.setText(f"⚙️ Опции вёрстки{suffix} ▾")
            if hasattr(self, "btn_batch_opts"):
                self.btn_batch_opts.setText(f"⚙️ Опции пакета{suffix} ▾")

        def _on_two_col_toggled(self, checked: bool):
            self.config["two_columns"] = bool(checked)
            save_config(self.config)
            self._sync_two_col_actions(bool(checked))
            self._toast("Включены 2 колонки (IEEE / ВАК)" if checked else "Включена 1 колонка")

        def _on_strip_emphasis_toggled(self, checked: bool):
            self.config["strip_body_emphasis"] = bool(checked)
            save_config(self.config)
            self._sync_strip_actions(bool(checked))
            self._toast("Жирный/курсив в тексте будет очищаться" if checked else "Форматирование текста сохраняется")

        def _on_merge_apply_toggled(self, checked: bool):
            self.config["merge_apply_template"] = bool(checked)
            save_config(self.config)
            self._toast("Шаблон будет применён к результату склейки" if checked else "Стили исходных файлов сохранятся")

        def _get_effective_template_cfg(self, tpl_id: Optional[str] = None) -> dict:
            cfg = json.loads(json.dumps(load_template(tpl_id or self.active_tpl)))
            if self.config.get("two_columns", False):
                cfg.setdefault("page", {})["columns"] = 2
                cfg["page"].setdefault("column_spacing_mm", 5.0)
            if self.config.get("strip_body_emphasis", False):
                cfg["strip_body_emphasis"] = True
            return cfg

        def _on_main_run(self):
            if self.mode == "text":
                text = self.ed_input.toPlainText()
                if not text.strip():
                    self._toast("Введите или вставьте Markdown текст")
                    return
                tpl_cfg = self._get_effective_template_cfg()
                try:
                    if self.rb_append.isChecked():
                        target = self.ed_append_target.text().strip()
                        if not target or not Path(target).exists():
                            show_themed_dialog(self, "Файл не выбран", "Укажите существующий файл .docx для режима Append.", self.theme_name)
                            return
                        out_p = Path(target).resolve()
                        shutil.copy2(out_p, out_p.with_suffix(".docx.bak"))
                        append_markdown_to_docx(
                            str(out_p), text, tpl_cfg,
                            output_path=str(out_p),
                            update_existing_styles=not self.act_keep_styles.isChecked(),
                            add_page_break=self.act_page_break.isChecked(),
                            base_dir=self._qt_base_dir,
                        )
                        self.lbl_result_name.setText(f"✓ Дописано: {out_p.name}")
                    else:
                        title = extract_semantic_title(text)
                        out_dir = Path(self.config.get("output_dir", str((Path.home() / "Desktop" / "Word_Documents").resolve())))
                        out_dir.mkdir(parents=True, exist_ok=True)
                        out_p = out_dir / f"{title}.docx"
                        ast_nodes = MarkdownASTParser(tpl_cfg.get("structural_titles", [])).parse(text)
                        GostDocxFormatter(tpl_cfg).build_document(ast_nodes, str(out_p), base_dir=self._qt_base_dir)
                        self.lbl_result_name.setText(f"✓ Готово: {out_p.name}")
                    self._last_output = out_p
                    self.result_banner.setVisible(True)
                    self._toast(f"Собрано: {out_p.name}")
                    if self.config.get("open_after_create", True):
                        self._open_last_file()
                except Exception as ex:
                    show_themed_dialog(self, "Ошибка сборки", str(ex), self.theme_name, is_error=True)
            elif self.mode == "files":
                self._start_batch_conversion()
            elif self.mode == "merge":
                self._run_merge()
            elif self.mode == "templates" and getattr(self, "template_editor", None) is not None:
                self.template_editor.save_current_template()

        def _open_last_file(self):
            if self._last_output and self._last_output.exists():
                if sys.platform == "win32":
                    os.startfile(str(self._last_output))
                else:
                    subprocess.Popen(["xdg-open", str(self._last_output)])

        def _open_last_folder(self):
            if self._last_output:
                d = self._last_output.parent
                if sys.platform == "win32":
                    os.startfile(str(d))
                else:
                    subprocess.Popen(["xdg-open", str(d)])

        # --- Files / Batch actions ---
        def _on_files_dropped(self, paths: List[str]):
            exts = {".md", ".markdown", ".txt", ".html"}
            added = 0
            for p_str in paths:
                p = Path(p_str)
                if p.is_file() and p.suffix.lower() in exts:
                    added += self._add_queue_item(p)
                elif p.is_dir():
                    for sub in p.rglob("*"):
                        if sub.is_file() and sub.suffix.lower() in exts:
                            added += self._add_queue_item(sub)
            self._refresh_queue_table()
            if added:
                self._toast(f"Добавлено файлов: {added}")

        def _add_queue_item(self, p: Path) -> int:
            if any(x["input_path"] == str(p) for x in self.queue_items):
                return 0
            if self.config.get("save_mode", "same_dir") == "same_dir":
                out_p = p.with_suffix(".docx")
            else:
                out_p = Path(self.config.get("output_dir", "")) / f"{p.stem}.docx"
            sz_kb = max(0.1, p.stat().st_size / 1024.0)
            self.queue_items.append({
                "input_path": str(p),
                "output_path": str(out_p),
                "size": f"{sz_kb:.1f} КБ",
                "template_name": self.active_tpl,
                "selected": True,
                "custom_tpl": False,
                "status": "Ожидание"
            })
            return 1

        def _browse_files(self):
            files, _ = QFileDialog.getOpenFileNames(self, "Выберите файлы", "", "Markdown & Text (*.md *.markdown *.txt *.html)")
            if files:
                self._on_files_dropped(files)

        def _browse_folder(self):
            d = QFileDialog.getExistingDirectory(self, "Выберите папку")
            if d:
                self._on_files_dropped([d])

        def _open_path_in_os(self, target_p: Path):
            if not target_p.exists():
                return
            if sys.platform == "win32":
                os.startfile(str(target_p))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target_p)])
            else:
                subprocess.Popen(["xdg-open", str(target_p)])

        def _clear_queue(self):
            self.queue_items.clear()
            self.progress_bar.setVisible(False)
            self.batch_result_banner.setVisible(False)
            self._refresh_queue_table()

        def _set_all_queue_selected(self, state: bool):
            for it in self.queue_items:
                it["selected"] = bool(state)
            self._refresh_queue_table()
            self._toast("Выбраны все файлы" if state else "Выделение снято")

        def _remove_unselected_queue(self):
            before = len(self.queue_items)
            self.queue_items = [it for it in self.queue_items if it.get("selected", True)]
            self._refresh_queue_table()
            removed = before - len(self.queue_items)
            self._toast(f"Удалено невыбранных: {removed}" if removed else "Нет невыбранных файлов")

        def _apply_active_tpl_to_selected(self):
            cnt = 0
            for it in self.queue_items:
                if it.get("selected", True):
                    it["template_name"] = self.active_tpl
                    it["custom_tpl"] = False
                    cnt += 1
            self._refresh_queue_table()
            self._toast(f"Шаблон назначен для файлов: {cnt}")

        def _on_row_check_toggled(self, idx: int, checked: bool):
            if 0 <= idx < len(self.queue_items):
                self.queue_items[idx]["selected"] = bool(checked)
                self._update_footer_button()

        def _on_row_tpl_changed(self, idx: int, tpl_id: str):
            if 0 <= idx < len(self.queue_items) and tpl_id:
                self.queue_items[idx]["template_name"] = tpl_id
                self.queue_items[idx]["custom_tpl"] = True
                t_name = self.templates_cache.get(tpl_id, {}).get("display_title", tpl_id)
                cw = self.table_queue.cellWidget(idx, 4)
                btn = cw.findChild(QPushButton, "TableTplBtn") if cw is not None else None
                if btn is not None:
                    btn.setToolTip(t_name)
                    btn.setText(btn.fontMetrics().elidedText(f"{t_name} ▾", Qt.ElideRight, 196))
                    m = btn.menu()
                    if m is not None:
                        for act in m.actions():
                            act.setChecked(act.text() == t_name)
                else:
                    self._refresh_queue_table()
                self._toast(f"{Path(self.queue_items[idx]['input_path']).name}: {t_name}")

        def _on_queue_context_menu(self, pos):
            row = self.table_queue.rowAt(pos.y())
            if row < 0 or row >= len(self.queue_items):
                return
            item = self.queue_items[row]
            menu = QMenu(self.table_queue)
            apply_menu_theme(menu, self.theme_name)
            is_sel = item.get("selected", True)
            menu.addAction("Снять выделение" if is_sel else "Выбрать файл").triggered.connect(
                lambda: self._toggle_row_selection(row)
            )
            sub_tpl = menu.addMenu("Назначить шаблон")
            apply_menu_theme(sub_tpl, self.theme_name)
            for tid, cfg in self.templates_cache.items():
                act_s = sub_tpl.addAction(cfg.get("display_title", tid))
                act_s.setCheckable(True)
                act_s.setChecked(tid == item.get("template_name"))
                act_s.triggered.connect(lambda _, i=row, t_id=tid: self._on_row_tpl_changed(i, t_id))
            menu.addSeparator()
            out_p = Path(item["output_path"])
            menu.addAction("Открыть папку").triggered.connect(lambda: self._open_path_in_os(out_p.parent))
            if item.get("status") == "Готово" and out_p.exists():
                menu.addAction("Открыть документ").triggered.connect(lambda: self._open_path_in_os(out_p))
            menu.addSeparator()
            menu.addAction("Удалить из очереди").triggered.connect(lambda: self._remove_queue_row(row))
            menu.exec(self.table_queue.viewport().mapToGlobal(pos))
            menu.deleteLater()

        def _toggle_row_selection(self, row: int):
            if 0 <= row < len(self.queue_items):
                self.queue_items[row]["selected"] = not self.queue_items[row].get("selected", True)
                self._refresh_queue_table()

        def _remove_queue_row(self, row: int):
            if 0 <= row < len(self.queue_items):
                name = Path(self.queue_items[row]["input_path"]).name
                self.queue_items.pop(row)
                self._refresh_queue_table()
                self._toast(f"Удалён из очереди: {name}")

        def _refresh_queue_table(self):
            if not hasattr(self, "table_queue"):
                return
            self.table_queue.setUpdatesEnabled(False)
            try:
                self._refresh_queue_table_inner()
            finally:
                self.table_queue.setUpdatesEnabled(True)

        def _refresh_queue_table_inner(self):
            for old_r in range(self.table_queue.rowCount()):
                for old_c in (0, 4, 6):
                    cw = self.table_queue.cellWidget(old_r, old_c)
                    if cw is not None:
                        cw.setParent(None)
                        cw.deleteLater()
            self.table_queue.setRowCount(0)
            tpl_items = [(tid, cfg.get("display_title", tid)) for tid, cfg in self.templates_cache.items()]
            for idx, item in enumerate(self.queue_items):
                if item["status"] == "Ожидание" and not item.get("custom_tpl", False):
                    item["template_name"] = self.active_tpl
                r = self.table_queue.rowCount()
                self.table_queue.insertRow(r)
                in_p = Path(item["input_path"])
                out_dir = in_p.parent if self.config.get("save_mode", "same_dir") == "same_dir" else Path(self.config.get("output_dir", ""))
                out_p = out_dir / f"{in_p.stem}.docx"
                item["output_path"] = str(out_p)

                chk_w = QWidget()
                chk_w.setStyleSheet("background: transparent;")
                cl_lay = QHBoxLayout(chk_w)
                cl_lay.setContentsMargins(0, 0, 0, 0)
                cl_lay.setAlignment(Qt.AlignCenter)
                cb = QCheckBox()
                cb.setChecked(item.get("selected", True))
                cb.toggled.connect(lambda c, i=idx: self._on_row_check_toggled(i, c))
                cl_lay.addWidget(cb)
                self.table_queue.setCellWidget(r, 0, chk_w)

                self.table_queue.setItem(r, 1, QTableWidgetItem(in_p.name))
                self.table_queue.setItem(r, 2, QTableWidgetItem(str(out_dir)))
                self.table_queue.setItem(r, 3, QTableWidgetItem(item["size"]))

                cur_tid = item["template_name"]
                cur_lbl = self.templates_cache.get(cur_tid, {}).get("display_title", cur_tid)
                tpl_w = QWidget()
                tpl_w.setStyleSheet("background: transparent;")
                tl = QHBoxLayout(tpl_w)
                tl.setContentsMargins(4, 4, 4, 4)
                btn_row_tpl = QPushButton()
                btn_row_tpl.setObjectName("TableTplBtn")
                btn_row_tpl.setAttribute(Qt.WA_StyledBackground, True)
                btn_row_tpl.setFixedHeight(32)
                btn_row_tpl.setMaximumWidth(220)
                btn_row_tpl.setToolTip(cur_lbl)
                btn_row_tpl.setText(btn_row_tpl.fontMetrics().elidedText(f"{cur_lbl} ▾", Qt.ElideRight, 196))
                btn_row_tpl.setCursor(QCursor(Qt.PointingHandCursor))
                rm = QMenu(btn_row_tpl)
                apply_menu_theme(rm, self.theme_name)
                for tid, t_lbl in tpl_items:
                    act_t = rm.addAction(t_lbl)
                    act_t.setCheckable(True)
                    act_t.setChecked(tid == cur_tid)
                    act_t.triggered.connect(lambda _, i=idx, t_id=tid: self._on_row_tpl_changed(i, t_id))
                btn_row_tpl.setMenu(rm)
                tl.addWidget(btn_row_tpl)
                self.table_queue.setCellWidget(r, 4, tpl_w)

                st_item = QTableWidgetItem(item["status"])
                if item["status"] == "Готово":
                    st_item.setForeground(QColor("#0D9488"))
                elif item["status"].startswith("Ошибка"):
                    st_item.setForeground(QColor("#EF4444"))
                self.table_queue.setItem(r, 5, st_item)

                act_w = QWidget()
                act_w.setStyleSheet("background: transparent;")
                al = QHBoxLayout(act_w)
                al.setContentsMargins(6, 4, 6, 4)
                al.setSpacing(8)
                al.setAlignment(Qt.AlignCenter)
                btn_fld = QPushButton("📁")
                btn_fld.setProperty("class", "IconBtn")
                btn_fld.setFixedSize(30, 30)
                btn_fld.setToolTip(f"Открыть папку: {out_dir}")
                btn_fld.clicked.connect(lambda _, d=out_dir: self._open_path_in_os(d))
                al.addWidget(btn_fld)
                if item["status"] == "Готово" and out_p.exists():
                    btn_op = QPushButton("📄")
                    btn_op.setProperty("class", "IconBtn")
                    btn_op.setFixedSize(30, 30)
                    btn_op.setToolTip(f"Открыть документ: {out_p.name}")
                    btn_op.clicked.connect(lambda _, fp=out_p: self._open_path_in_os(fp))
                    al.addWidget(btn_op)
                self.table_queue.setRowHeight(r, 54)
                self.table_queue.setCellWidget(r, 6, act_w)
            self._update_footer_button()

        def _animate_progress(self, val: int):
            if hasattr(self, "_progress_anim"):
                self._progress_anim.stop()
                self._progress_anim.setStartValue(self.progress_bar.value())
                self._progress_anim.setEndValue(val)
                self._progress_anim.start()
            else:
                self.progress_bar.setValue(val)

        def _start_batch_conversion(self):
            selected_items = [it for it in self.queue_items if it.get("selected", True)]
            if not selected_items:
                self._toast("Отметьте галочкой хотя бы один файл для конвертации")
                return
            self.batch_result_banner.setVisible(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setMaximum(len(selected_items))
            self.progress_bar.setValue(0)
            self.btn_main_run.setEnabled(False)
            self.worker = BatchWorker(
                selected_items,
                self.templates_cache,
                self,
                two_columns=bool(self.config.get("two_columns", False)),
                strip_body_emphasis=bool(self.config.get("strip_body_emphasis", False)),
            )
            self.worker.progress_updated.connect(lambda c, t, fn: self._animate_progress(c))
            self.worker.item_finished.connect(self._on_batch_item_done)
            self.worker.all_completed.connect(self._on_batch_done)
            self.worker.finished.connect(self.worker.deleteLater)
            self.worker.start()

        def _on_batch_item_done(self, in_str: str, out_str: str, st: str, ok: bool):
            for idx, it in enumerate(self.queue_items):
                if it["input_path"] == in_str:
                    it["status"] = st
                    st_item = QTableWidgetItem(st)
                    if st == "Готово":
                        st_item.setForeground(QColor("#0D9488"))
                    elif st.startswith("Ошибка"):
                        st_item.setForeground(QColor("#EF4444"))
                    self.table_queue.setItem(idx, 5, st_item)
                    if ok:
                        self._last_output = Path(out_str)
                    break

        def _on_batch_done(self, ok_cnt: int, fail_cnt: int):
            self.btn_main_run.setEnabled(True)
            if HAS_ENGINE:
                try:
                    latex_to_omml.cache_clear()
                except Exception:
                    pass
            gc.collect()
            self._refresh_queue_table()
            if ok_cnt > 0 and self._last_output:
                self.lbl_batch_result.setText(f"✓ Готово файлов: {ok_cnt} (ошибок: {fail_cnt}) — {self._last_output.name}")
                self.batch_result_banner.setVisible(True)
                if self.config.get("open_after_create", True):
                    self._open_last_file()
            self._toast(f"Пакетная конвертация завершена ({ok_cnt} готово, {fail_cnt} ошибок)")

        # --- Merge / Restyle actions ---
        def _sync_merge_list_ui(self):
            self.list_merge.clear()
            for i, fp in enumerate(self.merge_files, 1):
                it = QListWidgetItem(f"{i}. {Path(fp).name}")
                it.setData(Qt.UserRole, fp)
                it.setToolTip(fp)
                self.list_merge.addItem(it)

        def _on_merge_rows_moved(self, *args):
            reordered = []
            for r in range(self.list_merge.count()):
                fp = self.list_merge.item(r).data(Qt.UserRole)
                if fp:
                    reordered.append(fp)
            if reordered:
                self.merge_files = reordered
                QTimer.singleShot(0, self._sync_merge_list_ui)
                self._toast("Порядок документов обновлён")

        def _on_merge_context_menu(self, pos):
            menu = QMenu(self.list_merge)
            apply_menu_theme(menu, self.theme_name)
            row = self.list_merge.row(self.list_merge.itemAt(pos))
            if row >= 0:
                self.list_merge.setCurrentRow(row)
                menu.addAction("Удалить выбранный").triggered.connect(self._merge_remove)
                menu.addSeparator()
            if self.merge_files:
                menu.addAction("Очистить список").triggered.connect(self._merge_clear)
            else:
                menu.addAction("+ Добавить .docx").triggered.connect(self._merge_add)
            menu.exec(self.list_merge.viewport().mapToGlobal(pos))
            menu.deleteLater()

        def _merge_add(self):
            files, _ = QFileDialog.getOpenFileNames(self, "Выберите DOCX файлы", "", "Word Documents (*.docx)")
            for f in files:
                self.merge_files.append(f)
            if files:
                self._sync_merge_list_ui()

        def _merge_move(self, delta: int):
            r = self.list_merge.currentRow()
            if r < 0:
                return
            nr = r + delta
            if 0 <= nr < len(self.merge_files):
                self.merge_files[r], self.merge_files[nr] = self.merge_files[nr], self.merge_files[r]
                self.list_merge.clear()
                for i, fp in enumerate(self.merge_files, 1):
                    self.list_merge.addItem(f"{i}. {Path(fp).name}")
                self.list_merge.setCurrentRow(nr)

        def _merge_remove(self):
            r = self.list_merge.currentRow()
            if r >= 0:
                self.merge_files.pop(r)
                self.list_merge.clear()
                for i, fp in enumerate(self.merge_files, 1):
                    self.list_merge.addItem(f"{i}. {Path(fp).name}")

        def _merge_clear(self):
            self.merge_files.clear()
            self.list_merge.clear()
            self._toast("Список склейки очищен")

        def _run_merge(self):
            if len(self.merge_files) < 2:
                self._toast("Добавьте минимум 2 файла .docx для склейки")
                return
            def_out = str(Path(self.merge_files[0]).with_name(Path(self.merge_files[0]).stem + "_merged.docx"))
            out_f, _ = QFileDialog.getSaveFileName(self, "Сохранить объединенный DOCX", def_out, "Word Documents (*.docx)")
            if not out_f:
                return
            try:
                res = merge_docx_files(self.merge_files[0], self.merge_files[1:], output_path=out_f, add_page_break=self.act_merge_pb.isChecked())
                if self.act_merge_apply_tpl.isChecked() or self.config.get("two_columns", False) or self.config.get("strip_body_emphasis", False):
                    eff_cfg = self._get_effective_template_cfg()
                    if not self.act_merge_apply_tpl.isChecked():
                        minimal_cfg = {}
                        if self.config.get("two_columns", False):
                            minimal_cfg["page"] = {"columns": 2, "column_spacing_mm": 5.0}
                        if self.config.get("strip_body_emphasis", False):
                            minimal_cfg["strip_body_emphasis"] = True
                        update_docx_styles(res, minimal_cfg, output_path=res)
                    else:
                        update_docx_styles(res, eff_cfg, output_path=res)
                self._last_output = Path(res)
                self._toast(f"Склеено: {Path(res).name}")
                if self.config.get("open_after_create", True):
                    self._open_last_file()
            except Exception as ex:
                show_themed_dialog(self, "Ошибка склейки", str(ex), self.theme_name, is_error=True)

        def _pick_restyle_target(self):
            f, _ = QFileDialog.getOpenFileName(self, "Выберите DOCX", "", "Word Documents (*.docx)")
            if f:
                self.ed_restyle_target.setText(f)

        def _run_restyle(self):
            target = self.ed_restyle_target.text().strip()
            if not target or not Path(target).exists():
                show_themed_dialog(self, "Файл не выбран", "Укажите существующий файл .docx.", self.theme_name)
                return
            try:
                tp = Path(target).resolve()
                shutil.copy2(tp, tp.with_suffix(".docx.bak"))
                update_docx_styles(str(tp), self._get_effective_template_cfg(), output_path=str(tp))
                self._toast(f"Стили обновлены: {tp.name}")
            except Exception as ex:
                show_themed_dialog(self, "Ошибка", str(ex), self.theme_name, is_error=True)

        def _trigger_hotkey_dialog(self):
            if self.app_controller:
                self.app_controller.trigger_clipboard_convert()


    class ClipboardDocxApp:
        """Контроллер приложения: системный трей, нативные Win32-хоткеи и главное окно."""
        def __init__(self, qapp: QApplication):
            self.qapp = qapp
            self.config = load_config()
            self.toast = ToastOverlay()
            self.main_window = UnifiedMainWindow(app_controller=self)

            self.hotkey_emitter = HotkeyEmitter()
            self.hotkey_emitter.trigger_clipboard.connect(self.trigger_clipboard_convert)
            self.hotkey_emitter.trigger_main_window.connect(self.show_main_window)

            self._exiting = False
            self._setup_tray()
            self._setup_hotkeys()
            self.qapp.aboutToQuit.connect(self.cleanup_on_exit)

        def _setup_tray(self):
            self.tray_icon = QSystemTrayIcon(self.qapp)
            app_icon = create_app_icon()
            self.qapp.setWindowIcon(app_icon)
            self.main_window.setWindowIcon(app_icon)
            self.tray_icon.setIcon(app_icon)
            self.tray_icon.setToolTip("Markdown to DOCX Suite (Ctrl+Alt+W / Ctrl+Alt+B)")

            menu = QMenu()
            apply_menu_theme(menu, self.main_window.theme_name)
            self._tray_menu = menu
            act_w = menu.addAction("Открыть окно")
            act_w.triggered.connect(self.show_main_window)
            act_c = menu.addAction("Собрать из буфера (Ctrl+Alt+W)")
            act_c.triggered.connect(self.trigger_clipboard_convert)
            menu.addSeparator()
            act_q = menu.addAction("Выход")
            act_q.triggered.connect(self.exit_app)
            self.tray_icon.setContextMenu(menu)
            self.tray_icon.activated.connect(lambda r: self.show_main_window() if r in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick) else None)
            self.tray_icon.show()

        def show_main_window(self):
            self.main_window.show()
            self.main_window.setWindowState(self.main_window.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            self.main_window.raise_()
            self.main_window.activateWindow()

        def _setup_hotkeys(self):
            self._win_hotkey_filter = None
            win_ok = False
            if sys.platform == "win32" and user32 is not None:
                try:
                    WM_HOTKEY = 0x0312
                    MOD_ALT = 0x0001
                    MOD_CONTROL = 0x0002
                    MOD_NOREPEAT = 0x4000
                    HOTKEY_CLIP_ID = 101
                    HOTKEY_MAIN_ID = 102

                    ok1 = bool(user32.RegisterHotKey(None, HOTKEY_CLIP_ID, MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, ord('W')))
                    ok2 = bool(user32.RegisterHotKey(None, HOTKEY_MAIN_ID, MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, ord('B')))

                    class _WinHotkeyFilter(QAbstractNativeEventFilter):
                        def __init__(self, emitter):
                            super().__init__()
                            self.emitter = emitter

                        def nativeEventFilter(self, eventType, message):
                            if eventType in (b"windows_generic_MSG", b"windows_dispatcher_MSG"):
                                msg = wintypes.MSG.from_address(int(message))
                                if msg.message == WM_HOTKEY:
                                    if msg.wParam == HOTKEY_CLIP_ID:
                                        self.emitter.trigger_clipboard.emit()
                                        return True, 0
                                    elif msg.wParam == HOTKEY_MAIN_ID:
                                        self.emitter.trigger_main_window.emit()
                                        return True, 0
                            return False, 0

                    if ok1 or ok2:
                        self._win_hotkey_filter = _WinHotkeyFilter(self.hotkey_emitter)
                        self.qapp.installNativeEventFilter(self._win_hotkey_filter)
                        win_ok = (ok1 and ok2)
                        log(f"Win32 RegisterHotKey: Ctrl+Alt+W={ok1}, Ctrl+Alt+B={ok2}")
                except Exception as ex:
                    log(f"Win32 RegisterHotKey error: {ex}")

            if not win_ok and keyboard:
                try:
                    keyboard.clear_all_hotkeys()
                    keyboard.add_hotkey("ctrl+alt+w", self.hotkey_emitter.trigger_clipboard.emit)
                    keyboard.add_hotkey("ctrl+alt+b", self.hotkey_emitter.trigger_main_window.emit)
                except Exception as ex:
                    log(f"Warning setting up hotkeys: {ex}")

        def trigger_clipboard_convert(self):
            text = ""
            if pyperclip:
                try:
                    text = pyperclip.paste()
                except Exception:
                    pass
            if not text and QApplication.clipboard():
                text = QApplication.clipboard().text()

            if not text or not text.strip():
                self.toast.show_toast("Буфер обмена пуст! Скопируйте Markdown или формулу.", self.main_window.theme_name)
                return

            st = analyze_markdown_text(text)
            title = extract_semantic_title(text)
            dlg = TemplateSelectorDialog(title, st, theme_name=self.main_window.theme_name)
            res_code = dlg.exec()
            if res_code != QDialog.Accepted:
                dlg.deleteLater()
                return

            sel_tpl = dlg.selected_template
            is_app_mode = getattr(dlg, "append_mode", False)
            app_target = getattr(dlg, "append_target", "")
            dlg_2col = getattr(dlg, "two_columns", False)
            dlg_strip = getattr(dlg, "strip_body_emphasis", False)
            dlg.deleteLater()
            cfg = load_config()
            try:
                tpl_cfg = json.loads(json.dumps(load_template(sel_tpl)))
                if dlg_2col:
                    tpl_cfg.setdefault("page", {})["columns"] = 2
                    tpl_cfg["page"].setdefault("column_spacing_mm", 5.0)
                if dlg_strip or cfg.get("strip_body_emphasis", False):
                    tpl_cfg["strip_body_emphasis"] = True
                if is_app_mode and app_target:
                    out_p = Path(app_target).resolve()
                    if out_p.exists():
                        shutil.copy2(out_p, out_p.with_suffix(".docx.bak"))
                    append_markdown_to_docx(str(out_p), text, tpl_cfg, output_path=str(out_p))
                else:
                    out_dir = Path(cfg.get("output_dir", str((Path.home() / "Desktop" / "Word_Documents").resolve())))
                    out_dir.mkdir(parents=True, exist_ok=True)
                    out_p = out_dir / f"{title}.docx"
                    ast_nodes = MarkdownASTParser(tpl_cfg.get("structural_titles", [])).parse(text)
                    GostDocxFormatter(tpl_cfg).build_document(ast_nodes, str(out_p))

                self.toast.show_toast(f"Готово: {out_p.name}", self.main_window.theme_name, target_path=out_p)
                if cfg.get("open_after_create", True):
                    if sys.platform == "win32":
                        os.startfile(str(out_p))
                    else:
                        subprocess.Popen(["xdg-open", str(out_p)])
            except Exception as e:
                self.toast.show_toast(f"Ошибка сборки: {str(e)[:60]}", self.main_window.theme_name)

        def cleanup_on_exit(self):
            if getattr(self, "_exiting", False):
                return
            self._exiting = True
            if sys.platform == "win32" and user32 is not None:
                try:
                    user32.UnregisterHotKey(None, 101)
                    user32.UnregisterHotKey(None, 102)
                except Exception:
                    pass
            if getattr(self, "_win_hotkey_filter", None) is not None:
                try:
                    self.qapp.removeNativeEventFilter(self._win_hotkey_filter)
                except Exception:
                    pass
                self._win_hotkey_filter = None
            if keyboard is not None:
                try:
                    keyboard.unhook_all()
                except Exception:
                    pass
            try:
                self.toast._hide_timer.stop()
                self.toast.hide()
                self.toast.deleteLater()
            except Exception:
                pass
            try:
                try:
                    self.qapp.removeEventFilter(self.main_window)
                except Exception:
                    pass
                self.main_window.cleanup()
                self.main_window.deleteLater()
            except Exception:
                pass
            try:
                self.tray_icon.hide()
                self.tray_icon.deleteLater()
            except Exception:
                pass
            _SVG_ICON_CACHE.clear()
            gc.collect()

        def exit_app(self):
            self.cleanup_on_exit()
            self.qapp.quit()

else:
    UnifiedMainWindow = None
    ClipboardDocxApp = None


# ==============================================================================
# 3. CLI ТОЧКА ВХОДА
# ==============================================================================

def main():
    import argparse
    if len(sys.argv) >= 2 and sys.argv[1] in ("--append", "--append-clipboard", "--merge", "--apply-styles"):
        ap = argparse.ArgumentParser(description="Markdown to DOCX Suite — CLI")
        ap.add_argument("--append", nargs=2, metavar=("TARGET_DOCX", "INPUT_MD"), help="Добавить Markdown в конец существующего DOCX")
        ap.add_argument("--append-clipboard", metavar="TARGET_DOCX", help="Добавить текст из буфера обмена в конец существующего DOCX")
        ap.add_argument("--merge", nargs="+", metavar="DOCX", help="Склеить DOCX файлы (первый — базовый, остальные добавляются в конец)")
        ap.add_argument("--apply-styles", metavar="TARGET_DOCX", help="Обновить стандартные стили существующего DOCX по шаблону")
        ap.add_argument("--template", default="dstu_2026_gost_russian_names", help="Имя шаблона")
        ap.add_argument("-o", "--output", default=None, help="Выходной файл (по умолчанию перезаписывает TARGET_DOCX)")
        ap.add_argument("--update-styles", action="store_true", help="Обновить существующие стили документа по шаблону")
        ap.add_argument("--page-break", action="store_true", help="Вставить разрыв страницы перед добавляемым содержимым")
        args = ap.parse_args()

        if args.append:
            target_docx, in_md = args.append
            tpl_cfg = load_template(args.template)
            with open(in_md, "r", encoding="utf-8") as f:
                content = f.read()
            res = append_markdown_to_docx(
                target_docx, content, tpl_cfg,
                output_path=args.output,
                update_existing_styles=args.update_styles,
                add_page_break=args.page_break,
                base_dir=str(Path(in_md).resolve().parent),
            )
            print(f"ГОТОВО (append): {res}")
            return

        if args.append_clipboard:
            if not pyperclip:
                print("ОШИБКА: pyperclip не установлен.")
                sys.exit(1)
            content = pyperclip.paste()
            if not content or not content.strip():
                print("ОШИБКА: буфер обмена пуст.")
                sys.exit(1)
            tpl_cfg = load_template(args.template)
            res = append_markdown_to_docx(
                args.append_clipboard, content, tpl_cfg,
                output_path=args.output,
                update_existing_styles=args.update_styles,
                add_page_break=args.page_break,
            )
            print(f"ГОТОВО (append-clipboard): {res}")
            return

        if args.merge:
            if len(args.merge) < 2:
                print("ОШИБКА: для --merge нужно минимум 2 файла DOCX.")
                sys.exit(1)
            res = merge_docx_files(
                args.merge[0], args.merge[1:],
                output_path=args.output,
                add_page_break=args.page_break,
            )
            print(f"ГОТОВО (merge): {res}")
            return

        if args.apply_styles:
            tpl_cfg = load_template(args.template)
            res = update_docx_styles(args.apply_styles, tpl_cfg, output_path=args.output)
            print(f"ГОТОВО (apply-styles): {res}")
            return

    if len(sys.argv) >= 3 and not sys.argv[1].startswith("-"):
        in_md = sys.argv[1]
        out_docx = sys.argv[2]
        tmpl = sys.argv[3] if len(sys.argv) > 3 else "dstu_2026_gost_russian_names"
        with open(in_md, "r", encoding="utf-8") as f:
            content = f.read()
        tpl_cfg = load_template(tmpl)
        parser = MarkdownASTParser(tpl_cfg.get("structural_titles", []))
        ast_nodes = parser.parse(content)
        formatter = GostDocxFormatter(tpl_cfg)
        formatter.build_document(ast_nodes, out_docx, base_dir=str(Path(in_md).resolve().parent))
        print(f"ГОТОВО: {out_docx}")
        return

    if not HAS_PYSIDE6:
        print("ВНИМАНИЕ: PySide6 не установлен. Используйте CLI-режим.")
        sys.exit(1)

    try:
        try:
            os.chdir(str(get_base_dir()))
        except Exception:
            pass
        start_minimized = any(a in ("--minimized", "--tray") for a in sys.argv[1:])
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        init_font = load_config().get("ui_font", "Inter")
        app.setFont(QFont(init_font, 10))
        app.setQuitOnLastWindowClosed(False)
        runner = ClipboardDocxApp(app)
        if not start_minimized:
            runner.show_main_window()
        else:
            runner.toast.show_toast("Запущено в трее (Ctrl+Alt+W)", runner.main_window.theme_name)
        sys.exit(app.exec())
    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        log(f"FATAL STARTUP ERROR:\n{tb}")
        if sys.platform == "win32" and user32 is not None:
            try:
                user32.MessageBoxW(None, f"Ошибка запуска:\n{exc}\n\nПодробности в {LOG_FILE}", "ClipboardDocxRunner — Ошибка", 0x10)
            except Exception:
                pass
        raise

if __name__ == "__main__":
    main()
