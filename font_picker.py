#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Системный выбор шрифтов с автоматической группировкой в подпапки по типам:
- С засечками
- Без засечек
- Моноширинные
- Рукописные и декоративные
- Прочие системные
"""

from __future__ import annotations

from typing import Dict, List, Tuple, Optional, Callable

try:
    from PySide6.QtWidgets import (
        QDialog, QVBoxLayout, QHBoxLayout, QFrame, QLabel, QLineEdit,
        QPushButton, QTreeWidget, QTreeWidgetItem, QWidget, QSizePolicy
    )
    from PySide6.QtGui import QFont, QFontDatabase, QCursor
    from PySide6.QtCore import Qt, Signal, QPropertyAnimation, QEasingCurve
    HAS_PYSIDE6 = True
except ImportError:
    HAS_PYSIDE6 = False


_SERIF_HINTS = (
    "times", "serif", "georgia", "cambria", "garamond", "palatino", "charter",
    "book", "century", "minion", "baskerville", "didot", "bodoni", "constantia",
    "sylfaen", "antiqua", "schoolbook", "rockwell", "slab", "playfair",
    "merriweather", "lora", "crimson", "literata", "eb garamond", "dejavu serif",
    "liberation serif", "pt serif", "noto serif", "ibm plex serif",
)

_SANS_HINTS = (
    "sans", "segoe", "inter", "roboto", "arial", "calibri", "helvetica",
    "tahoma", "verdana", "trebuchet", "candara", "corbel", "montserrat",
    "ubuntu", "lato", "fira", "bahnschrift", "aptos", "franklin", "gill",
    "futura", "avenir", "rubik", "nunito", "manrope", "golos", "yandex",
)

_SCRIPT_HINTS = (
    "script", "cursive", "brush", "comic", "impact", "pacifico", "lobster",
    "hand", "calligraph", "corsiva", "segoe print", "gabriola", "mistral",
    "papyrus", "chiller", "jokerman", "vivaldi", "kunstler", "freestyle",
)

_SKIP_FONTS = (
    "wingdings", "webdings", "symbol", "marlett", "mt extra", "hololens",
    "segoe mdl2", "segoe fluent",
)

FONT_CATEGORIES = [
    ("serif", "С засечками"),
    ("sans", "Без засечек"),
    ("mono", "Моноширинные"),
    ("display", "Рукописные и декоративные"),
    ("other", "Прочие системные"),
]

CATEGORY_LABELS = dict(FONT_CATEGORIES)

_SYSTEM_FONTS_CACHE: Optional[List[Tuple[str, str]]] = None


def classify_font_family(family: str, is_fixed_pitch: bool = False) -> str:
    low = (family or "").strip().lower()
    if is_fixed_pitch or any(k in low for k in ("mono", "courier", "consolas", "fixed", "terminal", "menlo", "monaco", "hack", "cascadia")):
        return "mono"
    if any(k in low for k in _SCRIPT_HINTS):
        return "display"
    if any(k in low for k in _SERIF_HINTS):
        return "serif"
    if any(k in low for k in _SANS_HINTS):
        return "sans"
    return "other"


def get_system_fonts_grouped() -> List[Tuple[str, str]]:
    """Возвращает кэшированный список (family_name, category_id) всех системных шрифтов."""
    global _SYSTEM_FONTS_CACHE
    if _SYSTEM_FONTS_CACHE is not None:
        return _SYSTEM_FONTS_CACHE

    result: List[Tuple[str, str]] = []
    if HAS_PYSIDE6:
        try:
            families = sorted(set(QFontDatabase.families()), key=lambda s: s.lower())
            for fam in families:
                low = fam.lower()
                if any(sk in low for sk in _SKIP_FONTS):
                    continue
                is_fixed = QFontDatabase.isFixedPitch(fam)
                cat = classify_font_family(fam, is_fixed)
                result.append((fam, cat))
        except Exception:
            pass

    if not result:
        for fam, cat in [
            ("Times New Roman", "serif"), ("Cambria", "serif"), ("Georgia", "serif"),
            ("PT Serif", "serif"), ("Liberation Serif", "serif"),
            ("Segoe UI", "sans"), ("Inter", "sans"), ("Arial", "sans"),
            ("Calibri", "sans"), ("PT Sans", "sans"), ("Roboto", "sans"),
            ("Courier New", "mono"), ("Consolas", "mono"), ("JetBrains Mono", "mono"),
        ]:
            result.append((fam, cat))

    _SYSTEM_FONTS_CACHE = result
    return result


if HAS_PYSIDE6:

    class FontPickerDialog(QDialog):
        """Окно выбора системных шрифтов с группировкой в подпапки по типам и поиском."""
        def __init__(
            self,
            current_font: str = "Times New Roman",
            default_category: str = "all",
            theme_tokens: Optional[dict] = None,
            context_menu_installer: Optional[Callable[[QWidget], None]] = None,
            parent=None,
        ):
            super().__init__(parent)
            self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
            self.setAttribute(Qt.WA_TranslucentBackground, True)
            self.setMinimumSize(640, 540)
            self.resize(680, 580)

            self._t = theme_tokens or {
                "bg_window": "#FFFFFF", "bg_subtle": "#F8FAFC", "bg_muted": "#F1F5F9",
                "border": "#E2E8F0", "border_strong": "#CBD5E1",
                "text": "#0F172A", "text_2": "#334155", "text_mute": "#64748B",
                "accent": "#2563EB", "accent_soft": "#EFF6FF", "accent_text": "#1E40AF",
            }
            self._install_ctx = context_menu_installer
            self.selected_font = current_font
            self._default_cat = default_category
            self._fonts = get_system_fonts_grouped()

            self._build_ui()
            self._populate_tree()
            self._fade = QPropertyAnimation(self, b"windowOpacity", self)
            self._fade.setDuration(150)
            self._fade.setEasingCurve(QEasingCurve.OutCubic)

        def showEvent(self, event):
            super().showEvent(event)
            self.setWindowOpacity(0.0)
            self._fade.stop()
            self._fade.setStartValue(0.0)
            self._fade.setEndValue(1.0)
            self._fade.start()

        def _build_ui(self):
            t = self._t
            outer = QVBoxLayout(self)
            outer.setContentsMargins(6, 6, 6, 6)

            box = QFrame()
            box.setObjectName("FontPickerBox")
            box.setAttribute(Qt.WA_StyledBackground, True)
            box.setStyleSheet(
                f"QFrame#FontPickerBox {{ background-color: {t['bg_window']}; border: 1px solid {t['border_strong']}; border-radius: 14px; }}"
            )
            outer.addWidget(box)

            lay = QVBoxLayout(box)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            hdr = QHBoxLayout()
            lbl = QLabel("Выбор системного шрифта")
            lbl.setStyleSheet(f"color: {t['text']}; font-size: 15px; font-weight: 700; border: none; background: transparent;")
            hdr.addWidget(lbl, 1)
            btn_x = QPushButton("✕")
            btn_x.setProperty("class", "WinBtn")
            btn_x.setFixedSize(30, 30)
            btn_x.setCursor(QCursor(Qt.PointingHandCursor))
            btn_x.clicked.connect(self.reject)
            hdr.addWidget(btn_x)
            lay.addLayout(hdr)

            self.ed_search = QLineEdit()
            self.ed_search.setPlaceholderText("Поиск шрифта по названию (например: Times, Cambria, Segoe, Mono)...")
            if self._install_ctx:
                self._install_ctx(self.ed_search)
            self.ed_search.textChanged.connect(self._populate_tree)
            lay.addWidget(self.ed_search)

            self.tree_fonts = QTreeWidget()
            self.tree_fonts.setHeaderHidden(True)
            self.tree_fonts.setIndentation(18)
            self.tree_fonts.setAnimated(True)
            self.tree_fonts.setFocusPolicy(Qt.NoFocus)
            self.tree_fonts.setStyleSheet(
                f"QTreeWidget {{ background-color: {t['bg_subtle']}; border: 1px solid {t['border']}; "
                f"border-radius: 9px; padding: 6px; color: {t['text']}; }}"
                f"QTreeWidget::item {{ padding: 5px 8px; border-radius: 6px; }}"
                f"QTreeWidget::item:hover {{ background-color: {t['bg_muted']}; }}"
                f"QTreeWidget::item:selected {{ background-color: {t['accent_soft']}; color: {t['accent_text']}; }}"
            )
            self.tree_fonts.currentItemChanged.connect(self._on_current_changed)
            self.tree_fonts.itemDoubleClicked.connect(self._on_item_double_clicked)
            lay.addWidget(self.tree_fonts, 1)

            self.lbl_preview = QLabel("АаБбВвГг — Пример текста 12345")
            self.lbl_preview.setWordWrap(True)
            self.lbl_preview.setStyleSheet(
                f"background-color: {t['bg_subtle']}; color: {t['text']}; border: 1px solid {t['border']}; "
                f"border-radius: 8px; padding: 10px 12px; font-size: 14px;"
            )
            lay.addWidget(self.lbl_preview)

            foot = QHBoxLayout()
            foot.setSpacing(10)
            self.lbl_count = QLabel("")
            self.lbl_count.setStyleSheet(f"color: {t['text_mute']}; font-size: 12.5px; border: none; background: transparent;")
            foot.addWidget(self.lbl_count)
            foot.addStretch()
            btn_cancel = QPushButton("Отмена")
            btn_cancel.setProperty("class", "SecondaryBtn")
            btn_cancel.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            btn_cancel.setMinimumHeight(34)
            btn_cancel.setCursor(QCursor(Qt.PointingHandCursor))
            btn_cancel.clicked.connect(self.reject)
            foot.addWidget(btn_cancel)
            btn_ok = QPushButton("Выбрать шрифт")
            btn_ok.setProperty("class", "PrimaryBtn")
            btn_ok.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            btn_ok.setMinimumHeight(36)
            btn_ok.setCursor(QCursor(Qt.PointingHandCursor))
            btn_ok.clicked.connect(self._accept_selected)
            foot.addWidget(btn_ok)
            lay.addLayout(foot)

        def _populate_tree(self):
            q = self.ed_search.text().strip().lower()
            self.tree_fonts.clear()

            by_cat: Dict[str, List[str]] = {cid: [] for cid, _ in FONT_CATEGORIES}
            for fam, cat in self._fonts:
                if q and q not in fam.lower():
                    continue
                by_cat.setdefault(cat, []).append(fam)

            shown = 0
            selected_item: Optional[QTreeWidgetItem] = None
            t = self._t
            folder_font = QFont()
            folder_font.setBold(True)
            folder_font.setPointSize(10)

            for cid, clbl in FONT_CATEGORIES:
                fams = by_cat.get(cid, [])
                if not fams:
                    continue
                cat_item = QTreeWidgetItem(self.tree_fonts, [f"{clbl} ({len(fams)})"])
                cat_item.setFont(0, folder_font)
                cat_item.setFlags(cat_item.flags() & ~Qt.ItemIsSelectable)
                for fam in fams:
                    child = QTreeWidgetItem(cat_item, [fam])
                    child.setData(0, Qt.UserRole, fam)
                    child.setFont(0, QFont(fam, 11))
                    if fam.lower() == (self.selected_font or "").lower():
                        selected_item = child
                    shown += 1
                # При поиске или если это выбранная категория — раскрываем подпапку
                if q:
                    cat_item.setExpanded(True)

            self.lbl_count.setText(f"Найдено шрифтов: {shown}")
            if selected_item is not None:
                if q:
                    self.tree_fonts.setCurrentItem(selected_item)
                    self.tree_fonts.scrollToItem(selected_item)

        def _on_current_changed(self, cur: QTreeWidgetItem, prev: QTreeWidgetItem):
            if cur is None:
                return
            fam = cur.data(0, Qt.UserRole)
            if fam:
                self.selected_font = fam
                t = self._t
                self.lbl_preview.setText(f"{fam}: Автоматизированная система ГОСТ 7.32 — 12345")
                self.lbl_preview.setStyleSheet(
                    f"font-family: '{fam}'; background-color: {t['bg_subtle']}; color: {t['text']}; "
                    f"border: 1px solid {t['border']}; border-radius: 8px; padding: 10px 12px; font-size: 14px;"
                )

        def _on_item_double_clicked(self, item: QTreeWidgetItem, col: int):
            fam = item.data(0, Qt.UserRole)
            if fam:
                self.selected_font = fam
                self.accept()

        def _accept_selected(self):
            cur = self.tree_fonts.currentItem()
            if cur is not None and cur.data(0, Qt.UserRole):
                self.selected_font = cur.data(0, Qt.UserRole)
            self.accept()


    class FontPickerButton(QPushButton):
        """Кнопка выбора системного шрифта с фиксированным левым отступом 12px и открытием FontPickerDialog."""
        font_changed = Signal(str)

        def __init__(
            self,
            initial_font: str = "Times New Roman",
            default_category: str = "all",
            max_width: int = 190,
            theme_getter: Optional[Callable[[], dict]] = None,
            context_menu_installer: Optional[Callable[[QWidget], None]] = None,
            parent=None,
        ):
            super().__init__("", parent)
            self._font_name = initial_font
            self._default_cat = default_category
            self._max_w = max_width
            self._theme_getter = theme_getter
            self._install_ctx = context_menu_installer
            self.setObjectName("InputPill")
            self.setFixedHeight(34)
            self.setMinimumWidth(110)
            self.setMaximumWidth(max_width)
            self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            self.setCursor(QCursor(Qt.PointingHandCursor))
            lay = QHBoxLayout(self)
            lay.setContentsMargins(0, 0, 10, 0)
            lay.setSpacing(6)
            self._lbl_text = QLabel("")
            self._lbl_text.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self._lbl_text.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            self._lbl_text.setStyleSheet("background: transparent; border: none; font-size: 13px;")
            self._lbl_arr = QLabel("▾")
            self._lbl_arr.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self._lbl_arr.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            self._lbl_arr.setStyleSheet("background: transparent; border: none; font-size: 12px;")
            lay.addWidget(self._lbl_text, 1)
            lay.addWidget(self._lbl_arr, 0)
            self.clicked.connect(self._open_picker)
            self.set_font_name(initial_font)

        def font_name(self) -> str:
            return self._font_name

        def set_font_name(self, name: str):
            self._font_name = (name or "Times New Roman").strip()
            cat = "other"
            for fam, c in get_system_fonts_grouped():
                if fam.lower() == self._font_name.lower():
                    cat = c
                    break
            cat_lbl = CATEGORY_LABELS.get(cat, "")
            elided = self._lbl_text.fontMetrics().elidedText(self._font_name, Qt.ElideRight, max(70, self._max_w - 36))
            self._lbl_text.setText(elided)
            self.setToolTip(f"Системный шрифт: {self._font_name}" + (f" ({cat_lbl})" if cat_lbl else ""))

        def _open_picker(self):
            tokens = self._theme_getter() if self._theme_getter else None
            dlg = FontPickerDialog(
                current_font=self._font_name,
                default_category=self._default_cat,
                theme_tokens=tokens,
                context_menu_installer=self._install_ctx,
                parent=self.window(),
            )
            res = dlg.exec()
            chosen = dlg.selected_font
            dlg.deleteLater()
            if res == QDialog.Accepted and chosen and chosen != self._font_name:
                self.set_font_name(chosen)
                self.font_changed.emit(chosen)
