#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Полноценный органичный редактор JSON-шаблонов ГОСТ / Word для Markdown to DOCX Suite.
Поддерживает все параметры шаблонов:
- Страница, поля, колонки, strip_body_emphasis
- Шрифты (primary, monospace) и структурные заголовки
- Все стили документа (канонические русские имена, цвет, размер, начертание, выравнивание, интервалы, уровни структуры) + живой предпросмотр
- Таблицы (border_style, двойная граница шапки, заливка, отступы ячеек, repeat_header, cant_split)
- Списки (отступы уровней, выступы маркеров, символы, разделитель, интервалы)
"""

from __future__ import annotations

import copy
import json
import re
from pathlib import Path
from typing import Dict, Any, Callable, Optional, List

try:
    from PySide6.QtWidgets import (
        QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
        QCheckBox, QPushButton, QFrame, QScrollArea, QComboBox, QDoubleSpinBox,
        QSpinBox, QListWidget, QListWidgetItem, QStackedWidget, QPlainTextEdit,
        QColorDialog, QMenu, QSizePolicy, QDialog
    )
    from PySide6.QtGui import QColor, QFont, QCursor, QAction
    from PySide6.QtCore import Qt, Signal
    HAS_PYSIDE6 = True
except ImportError:
    HAS_PYSIDE6 = False


CANONICAL_STYLE_ORDER = [
    "Заголовок структурный",
    "Заголовок 1",
    "Заголовок 2",
    "Заголовок 3",
    "Основной текст",
    "Примечание",
    "Цитата",
    "Эпиграф",
    "Название таблицы",
    "Ячейка таблицы",
    "Рисунок",
    "Подпись рисунка",
    "Исходный код",
    "Название листинга",
]

DEFAULT_NEW_STYLE = {
    "font": "primary",
    "size_pt": 14.0,
    "bold": False,
    "italic": False,
    "all_caps": False,
    "color": "#000000",
    "alignment": "justify",
    "line_spacing": 1.5,
    "first_line_indent_mm": 12.5,
    "space_before_pt": 0.0,
    "space_after_pt": 0.0,
    "page_break_before": False,
    "keep_with_next": False,
}

DOC_FONT_CHOICES = [
    "Times New Roman",
    "Cambria",
    "Georgia",
    "Arial",
    "Calibri",
    "PT Serif",
    "PT Sans",
    "Liberation Serif",
    "Courier New",
    "Consolas",
    "JetBrains Mono",
]


def create_blank_template(name: str, display_title: str) -> dict:
    return {
        "name": name,
        "display_title": display_title,
        "description": "Пользовательский шаблон оформления",
        "specs_summary": "Times New Roman • 14 пт • Встроенные стили Word",
        "tags": ["Пользовательский"],
        "page": {
            "width_mm": 210.0,
            "height_mm": 297.0,
            "margin_left_mm": 30.0,
            "margin_right_mm": 20.0,
            "margin_top_mm": 20.0,
            "margin_bottom_mm": 20.0,
            "columns": 1,
            "column_spacing_mm": 5.0,
            "different_first_page": True,
            "page_numbers": False,
            "span_headings": True,
        },
        "fonts": {
            "primary": {
                "name": "Times New Roman",
                "ascii": "Times New Roman",
                "hAnsi": "Times New Roman",
                "cs": "Times New Roman",
                "eastAsia": "Times New Roman",
            },
            "monospace": {
                "name": "Courier New",
                "ascii": "Courier New",
                "hAnsi": "Courier New",
                "cs": "Courier New",
                "eastAsia": "Courier New",
            },
        },
        "structural_titles": [
            "аннотация",
            "содержание",
            "введение",
            "заключение",
            "список использованных источников",
            "приложение",
        ],
        "strip_body_emphasis": False,
        "styles": {
            "Заголовок структурный": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 16.0, "bold": True, "alignment": "center",
                "line_spacing": 1.0, "first_line_indent_mm": 0.0,
                "space_before_pt": 12.0, "space_after_pt": 12.0,
                "page_break_before": True, "keep_with_next": True, "outline_level": 0,
            },
            "Заголовок 1": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 16.0, "bold": True, "alignment": "left",
                "line_spacing": 1.0, "first_line_indent_mm": 12.5,
                "space_before_pt": 16.0, "space_after_pt": 12.0,
                "page_break_before": True, "keep_with_next": True, "outline_level": 0,
            },
            "Заголовок 2": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 14.0, "bold": True, "alignment": "left",
                "line_spacing": 1.0, "first_line_indent_mm": 12.5,
                "space_before_pt": 14.0, "space_after_pt": 10.0,
                "keep_with_next": True, "outline_level": 1,
            },
            "Заголовок 3": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 14.0, "bold": True, "alignment": "left",
                "line_spacing": 1.0, "first_line_indent_mm": 12.5,
                "space_before_pt": 12.0, "space_after_pt": 8.0,
                "keep_with_next": True, "outline_level": 2,
            },
            "Основной текст": {**DEFAULT_NEW_STYLE},
            "Примечание": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 12.0, "line_spacing": 1.15,
                "space_before_pt": 4.0, "space_after_pt": 4.0,
            },
            "Цитата": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 13.0, "italic": True, "line_spacing": 1.15,
                "first_line_indent_mm": 0.0, "space_before_pt": 6.0, "space_after_pt": 6.0,
            },
            "Название таблицы": {
                **DEFAULT_NEW_STYLE,
                "alignment": "left", "line_spacing": 1.0, "first_line_indent_mm": 0.0,
                "space_before_pt": 12.0, "space_after_pt": 6.0, "keep_with_next": True,
            },
            "Ячейка таблицы": {
                **DEFAULT_NEW_STYLE,
                "size_pt": 12.0, "alignment": "left", "line_spacing": 1.0,
                "first_line_indent_mm": 0.0, "space_before_pt": 2.0, "space_after_pt": 2.0,
            },
            "Рисунок": {
                **DEFAULT_NEW_STYLE,
                "alignment": "center", "line_spacing": 1.0, "first_line_indent_mm": 0.0,
                "space_before_pt": 12.0, "space_after_pt": 6.0, "keep_with_next": True,
            },
            "Подпись рисунка": {
                **DEFAULT_NEW_STYLE,
                "alignment": "center", "line_spacing": 1.0, "first_line_indent_mm": 0.0,
                "space_before_pt": 6.0, "space_after_pt": 12.0,
            },
            "Исходный код": {
                **DEFAULT_NEW_STYLE,
                "font": "monospace", "size_pt": 11.0, "alignment": "left",
                "line_spacing": 1.0, "first_line_indent_mm": 0.0,
            },
            "Название листинга": {
                **DEFAULT_NEW_STYLE,
                "alignment": "left", "line_spacing": 1.0, "first_line_indent_mm": 0.0,
                "space_before_pt": 12.0, "space_after_pt": 4.0, "keep_with_next": True,
            },
        },
        "tables": {
            "border_style": "full_grid",
            "header_double_border": True,
            "header_shading": None,
            "cell_padding_mm": {"top": 1.0, "bottom": 1.0, "left": 1.8, "right": 1.8},
            "repeat_header": True,
            "cant_split": True,
        },
        "lists": {
            "base_indent_mm": 12.5,
            "level_step_mm": 6.0,
            "hanging_indent_mm": {"dash": 4.5, "bullet": 4.5, "ordered": 6.5, "task": 5.5},
            "dash_symbol": "–",
            "bullet_symbols": ["•", "◦", "▪"],
            "marker_separator": "\u00a0",
            "space_before_pt": 0.0,
            "space_after_pt": 0.0,
        },
    }


def strip_internal_keys(cfg: dict) -> dict:
    clean = copy.deepcopy(cfg)
    for k in list(clean.keys()):
        if k.startswith("_"):
            del clean[k]
    return clean


if HAS_PYSIDE6:

    class TemplateEditorWidget(QWidget):
        """Полноценный редактор шаблонов, органично встроенный в UnifiedMainWindow."""
        template_saved = Signal(str)
        toast_requested = Signal(str)

        def __init__(
            self,
            templates_provider: Callable[[], Dict[str, dict]],
            primary_dir_provider: Callable[[], Path],
            register_path_fn: Callable[[Path], None],
            context_menu_installer: Optional[Callable[[QWidget], None]] = None,
            theme_getter: Optional[Callable[[], dict]] = None,
            parent=None,
        ):
            super().__init__(parent)
            self._templates_provider = templates_provider
            self._primary_dir_provider = primary_dir_provider
            self._register_path_fn = register_path_fn
            self._install_ctx = context_menu_installer
            self._theme_getter = theme_getter

            self._working_cfg: dict = {}
            self._current_tpl_id: str = ""
            self._current_style_key: str = ""
            self._loading: bool = False
            self._dirty: bool = False

            self._build_ui()

        def _t(self) -> dict:
            if self._theme_getter:
                return self._theme_getter()
            return {
                "bg_window": "#FFFFFF", "bg_subtle": "#F8FAFC", "bg_muted": "#F1F5F9",
                "border": "#E2E8F0", "border_strong": "#CBD5E1",
                "text": "#0F172A", "text_2": "#334155", "text_mute": "#64748B",
                "accent": "#2563EB", "accent_soft": "#EFF6FF", "accent_text": "#1E40AF",
                "teal": "#0D9488", "teal_soft": "#F0FDFA", "teal_text": "#0F766E",
            }

        def _attach_ctx(self, w: QWidget):
            if self._install_ctx:
                self._install_ctx(w)

        def _mark_dirty(self, dirty: bool = True):
            self._dirty = dirty
            if hasattr(self, "lbl_dirty_state"):
                t = self._t()
                if dirty:
                    self.lbl_dirty_state.setText("● Есть несохранённые изменения")
                    self.lbl_dirty_state.setStyleSheet(f"color: {t['accent_text']}; font-size: 12px; font-weight: 600;")
                else:
                    self.lbl_dirty_state.setText("✓ Все изменения сохранены")
                    self.lbl_dirty_state.setStyleSheet(f"color: {t['teal_text']}; font-size: 12px; font-weight: 600;")

        def _build_ui(self):
            root = QVBoxLayout(self)
            root.setContentsMargins(0, 0, 0, 0)
            root.setSpacing(12)

            # Top Action Bar
            bar = QFrame()
            bar.setProperty("class", "ContextBar")
            b_lay = QHBoxLayout(bar)
            b_lay.setContentsMargins(14, 10, 14, 10)
            b_lay.setSpacing(8)

            lbl_pick = QLabel("Шаблон:")
            lbl_pick.setStyleSheet("font-size: 13px; font-weight: 600;")
            b_lay.addWidget(lbl_pick)

            self.btn_tpl_picker = QPushButton("Выберите шаблон ▾")
            self.btn_tpl_picker.setObjectName("TplPill")
            self.btn_tpl_picker.setMinimumHeight(34)
            self.btn_tpl_picker.setMinimumWidth(240)
            self.btn_tpl_picker.setCursor(QCursor(Qt.PointingHandCursor))
            self.tpl_picker_menu = QMenu(self.btn_tpl_picker)
            self.btn_tpl_picker.setMenu(self.tpl_picker_menu)
            b_lay.addWidget(self.btn_tpl_picker)

            b_lay.addStretch()

            self.lbl_dirty_state = QLabel("✓ Сохранено")
            b_lay.addWidget(self.lbl_dirty_state)

            btn_new = QPushButton("+ Новый")
            btn_new.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            btn_new.setProperty("class", "SecondaryBtn")
            btn_new.setMinimumHeight(32)
            btn_new.setCursor(QCursor(Qt.PointingHandCursor))
            btn_new.clicked.connect(self._create_new_template)
            b_lay.addWidget(btn_new)

            btn_dup = QPushButton("Копия")
            btn_dup.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            btn_dup.setProperty("class", "SecondaryBtn")
            btn_dup.setMinimumHeight(32)
            btn_dup.setCursor(QCursor(Qt.PointingHandCursor))
            btn_dup.clicked.connect(self._duplicate_template)
            b_lay.addWidget(btn_dup)

            root.addWidget(bar)

            # Main Body: Left Navigation + Right Section Stack
            body = QHBoxLayout()
            body.setSpacing(12)

            # Left Sidebar
            nav_card = QFrame()
            nav_card.setProperty("class", "PaneCard")
            nav_card.setAttribute(Qt.WA_StyledBackground, True)
            nav_card.setFixedWidth(262)
            n_lay = QVBoxLayout(nav_card)
            n_lay.setContentsMargins(10, 12, 10, 12)
            n_lay.setSpacing(6)

            lbl_sec = QLabel("Разделы шаблона")
            lbl_sec.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.7; padding: 2px 4px;")
            n_lay.addWidget(lbl_sec)

            self.nav_buttons: List[QPushButton] = []
            sections = [
                ("📄  Страница и колонки", 0),
                ("🔤  Шрифты и заголовки", 1),
                ("🎨  Стили документа", 2),
                ("📊  Таблицы", 3),
                ("📝  Списки", 4),
            ]
            for s_lbl, idx in sections:
                sb = QPushButton(s_lbl)
                sb.setProperty("class", "NavBtn")
                sb.setProperty("active", "true" if idx == 2 else "false")
                sb.setMinimumHeight(34)
                sb.setCursor(QCursor(Qt.PointingHandCursor))
                sb.clicked.connect(lambda _, i=idx: self._switch_section(i))
                n_lay.addWidget(sb)
                self.nav_buttons.append(sb)

            n_lay.addSpacing(8)
            lbl_st_list = QLabel("Быстрый выбор стиля")
            lbl_st_list.setStyleSheet("font-size: 12px; font-weight: 600; opacity: 0.7; padding: 2px 4px;")
            n_lay.addWidget(lbl_st_list)

            self.list_styles = QListWidget()
            self.list_styles.currentItemChanged.connect(self._on_style_selected)
            n_lay.addWidget(self.list_styles, 1)

            st_btns = QHBoxLayout()
            st_btns.setSpacing(6)
            btn_add_st = QPushButton("+ Стиль")
            btn_add_st.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            btn_add_st.setProperty("class", "SecondaryBtn")
            btn_add_st.clicked.connect(self._add_custom_style)
            st_btns.addWidget(btn_add_st, 1)
            btn_del_st = QPushButton("Удалить")
            btn_del_st.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
            btn_del_st.setProperty("class", "SecondaryBtn")
            btn_del_st.clicked.connect(self._delete_selected_style)
            st_btns.addWidget(btn_del_st)
            n_lay.addLayout(st_btns)

            body.addWidget(nav_card)

            # Right Content Card
            right_card = QFrame()
            right_card.setProperty("class", "PaneCard")
            right_card.setAttribute(Qt.WA_StyledBackground, True)
            rc_lay = QVBoxLayout(right_card)
            rc_lay.setContentsMargins(0, 0, 0, 0)
            rc_lay.setSpacing(0)

            self.sec_stack = QStackedWidget()
            self.sec_stack.addWidget(self._wrap_scroll(self._build_page_section()))
            self.sec_stack.addWidget(self._wrap_scroll(self._build_fonts_section()))
            self.sec_stack.addWidget(self._wrap_scroll(self._build_style_section()))
            self.sec_stack.addWidget(self._wrap_scroll(self._build_tables_section()))
            self.sec_stack.addWidget(self._wrap_scroll(self._build_lists_section()))
            self.sec_stack.setCurrentIndex(2)
            rc_lay.addWidget(self.sec_stack, 1)

            body.addWidget(right_card, 1)
            root.addLayout(body, 1)

        def _wrap_scroll(self, inner: QWidget) -> QScrollArea:
            inner.setObjectName("EditorScrollInner")
            inner.setStyleSheet("QWidget#EditorScrollInner { background: transparent; }")
            sc = QScrollArea()
            sc.setWidgetResizable(True)
            sc.setFrameShape(QFrame.NoFrame)
            sc.setObjectName("PopScroll")
            sc.setStyleSheet("QScrollArea { background: transparent; border: none; }")
            sc.viewport().setStyleSheet("background: transparent;")
            sc.setWidget(inner)
            return sc

        def _switch_section(self, idx: int):
            self.sec_stack.setCurrentIndex(idx)
            for i, btn in enumerate(self.nav_buttons):
                btn.setProperty("active", "true" if i == idx else "false")
                btn.style().unpolish(btn)
                btn.style().polish(btn)

        # --- Section 0: Page & Columns ---
        def _build_page_section(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            lbl = QLabel("Параметры страницы, колонок и метаданные")
            lbl.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            lay.addWidget(lbl)

            g = QGridLayout()
            g.setSpacing(10)

            self.ed_tpl_title = QLineEdit()
            self._attach_ctx(self.ed_tpl_title)
            self.ed_tpl_title.textEdited.connect(self._on_meta_changed)
            g.addWidget(QLabel("Название в списке:"), 0, 0)
            g.addWidget(self.ed_tpl_title, 0, 1, 1, 3)

            self.ed_tpl_desc = QLineEdit()
            self._attach_ctx(self.ed_tpl_desc)
            self.ed_tpl_desc.textEdited.connect(self._on_meta_changed)
            g.addWidget(QLabel("Описание:"), 1, 0)
            g.addWidget(self.ed_tpl_desc, 1, 1, 1, 3)

            self.sp_pw = self._make_dspin(50, 500, 1.0, 210.0, " мм")
            self.sp_ph = self._make_dspin(50, 600, 1.0, 297.0, " мм")
            g.addWidget(QLabel("Ширина страницы:"), 2, 0)
            g.addWidget(self.sp_pw, 2, 1)
            g.addWidget(QLabel("Высота страницы:"), 2, 2)
            g.addWidget(self.sp_ph, 2, 3)

            self.sp_ml = self._make_dspin(0, 100, 1.0, 30.0, " мм")
            self.sp_mr = self._make_dspin(0, 100, 1.0, 20.0, " мм")
            self.sp_mt = self._make_dspin(0, 100, 1.0, 20.0, " мм")
            self.sp_mb = self._make_dspin(0, 100, 1.0, 20.0, " мм")
            g.addWidget(QLabel("Левое поле:"), 3, 0)
            g.addWidget(self.sp_ml, 3, 1)
            g.addWidget(QLabel("Правое поле:"), 3, 2)
            g.addWidget(self.sp_mr, 3, 3)
            g.addWidget(QLabel("Верхнее поле:"), 4, 0)
            g.addWidget(self.sp_mt, 4, 1)
            g.addWidget(QLabel("Нижнее поле:"), 4, 2)
            g.addWidget(self.sp_mb, 4, 3)

            self.sp_cols = QSpinBox()
            self.sp_cols.setRange(1, 3)
            self.sp_cols.valueChanged.connect(self._on_page_changed)
            self.sp_col_gap = self._make_dspin(1.0, 30.0, 0.5, 5.0, " мм")
            g.addWidget(QLabel("Количество колонок:"), 5, 0)
            g.addWidget(self.sp_cols, 5, 1)
            g.addWidget(QLabel("Интервал между колонками:"), 5, 2)
            g.addWidget(self.sp_col_gap, 5, 3)

            for sp in (self.sp_pw, self.sp_ph, self.sp_ml, self.sp_mr, self.sp_mt, self.sp_mb, self.sp_col_gap):
                sp.valueChanged.connect(self._on_page_changed)
            lay.addLayout(g)

            self.chk_diff_first = QCheckBox("Особый колонтитул на первой странице (different_first_page)")
            self.chk_page_nums = QCheckBox("Нумерация страниц в нижнем колонтитуле (page_numbers)")
            self.chk_span_h = QCheckBox("Заголовки 1 уровня на всю ширину при 2 колонках (span_headings)")
            self.chk_tpl_strip_em = QCheckBox("Убирать жирный и курсив в обычном тексте и списках (strip_body_emphasis)")
            for c in (self.chk_diff_first, self.chk_page_nums, self.chk_span_h, self.chk_tpl_strip_em):
                c.toggled.connect(self._on_page_changed)
                lay.addWidget(c)

            lay.addStretch()
            return w

        # --- Section 1: Fonts & Structural Titles ---
        def _build_fonts_section(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            lbl = QLabel("Шрифты документа и структурные заголовки")
            lbl.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            lay.addWidget(lbl)

            g = QGridLayout()
            g.setSpacing(10)
            self.combo_font_primary = QComboBox()
            self.combo_font_primary.setEditable(True)
            self.combo_font_primary.addItems(DOC_FONT_CHOICES)
            self.combo_font_primary.currentTextChanged.connect(self._on_fonts_changed)
            self._attach_ctx(self.combo_font_primary.lineEdit())

            self.combo_font_mono = QComboBox()
            self.combo_font_mono.setEditable(True)
            self.combo_font_mono.addItems(["Courier New", "Consolas", "JetBrains Mono", "DejaVu Sans Mono"])
            self.combo_font_mono.currentTextChanged.connect(self._on_fonts_changed)
            self._attach_ctx(self.combo_font_mono.lineEdit())

            g.addWidget(QLabel("Основной шрифт (primary):"), 0, 0)
            g.addWidget(self.combo_font_primary, 0, 1)
            g.addWidget(QLabel("Моноширинный шрифт (monospace):"), 1, 0)
            g.addWidget(self.combo_font_mono, 1, 1)
            lay.addLayout(g)

            lbl_st = QLabel("Структурные заголовки (по одному в строке, без номера):")
            lbl_st.setStyleSheet("font-size: 13px; font-weight: 600; margin-top: 6px;")
            lay.addWidget(lbl_st)

            self.ed_struct_titles = QPlainTextEdit()
            self.ed_struct_titles.setMinimumHeight(170)
            self._attach_ctx(self.ed_struct_titles)
            self.ed_struct_titles.textChanged.connect(self._on_struct_titles_changed)
            lay.addWidget(self.ed_struct_titles)
            lay.addStretch()
            return w

        # --- Section 2: Style Editor + Live Preview ---
        def _build_style_section(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            hdr = QHBoxLayout()
            self.lbl_style_title = QLabel("Стиль: Основной текст")
            self.lbl_style_title.setStyleSheet("font-size: 15px; font-weight: 700;")
            hdr.addWidget(self.lbl_style_title, 1)
            self.lbl_style_word_map = QLabel("")
            self.lbl_style_word_map.setProperty("class", "StatBadge")
            hdr.addWidget(self.lbl_style_word_map)
            lay.addLayout(hdr)

            # Live Preview Card
            self.preview_card = QFrame()
            self.preview_card.setProperty("class", "ContextBar")
            pc_lay = QVBoxLayout(self.preview_card)
            pc_lay.setContentsMargins(16, 12, 16, 12)
            pc_lay.setSpacing(4)
            lbl_pv_cap = QLabel("Предпросмотр стиля в документе:")
            lbl_pv_cap.setStyleSheet("font-size: 11px; opacity: 0.65; font-weight: 600;")
            pc_lay.addWidget(lbl_pv_cap)
            self.lbl_style_preview = QLabel("Пример абзаца: Автоматизированная система проектирования по ГОСТ 7.32-2017.")
            self.lbl_style_preview.setWordWrap(True)
            pc_lay.addWidget(self.lbl_style_preview)
            lay.addWidget(self.preview_card)

            g = QGridLayout()
            g.setSpacing(10)

            self.combo_st_font = QComboBox()
            self.combo_st_font.addItem("Основной (primary)", "primary")
            self.combo_st_font.addItem("Моноширинный (monospace)", "monospace")
            self.combo_st_font.currentIndexChanged.connect(self._on_style_prop_changed)

            self.sp_st_size = self._make_dspin(6.0, 48.0, 0.5, 14.0, " пт")
            self.combo_st_align = QComboBox()
            for val, txt in [("justify", "По ширине"), ("left", "По левому краю"), ("center", "По центру"), ("right", "По правому краю")]:
                self.combo_st_align.addItem(txt, val)
            self.combo_st_align.currentIndexChanged.connect(self._on_style_prop_changed)

            color_row = QHBoxLayout()
            color_row.setSpacing(6)
            self.ed_st_color = QLineEdit("#000000")
            self.ed_st_color.setMaxLength(7)
            self._attach_ctx(self.ed_st_color)
            self.ed_st_color.textEdited.connect(self._on_style_prop_changed)
            self.btn_st_color = QPushButton("")
            self.btn_st_color.setFixedSize(30, 30)
            self.btn_st_color.setCursor(QCursor(Qt.PointingHandCursor))
            self.btn_st_color.clicked.connect(self._pick_style_color)
            color_row.addWidget(self.ed_st_color, 1)
            color_row.addWidget(self.btn_st_color)

            g.addWidget(QLabel("Гарнитура:"), 0, 0)
            g.addWidget(self.combo_st_font, 0, 1)
            g.addWidget(QLabel("Кегль:"), 0, 2)
            g.addWidget(self.sp_st_size, 0, 3)

            g.addWidget(QLabel("Выравнивание:"), 1, 0)
            g.addWidget(self.combo_st_align, 1, 1)
            g.addWidget(QLabel("Цвет текста:"), 1, 2)
            g.addLayout(color_row, 1, 3)

            self.sp_st_line = self._make_dspin(0.8, 3.0, 0.05, 1.5, "×")
            self.sp_st_indent = self._make_dspin(0.0, 50.0, 0.5, 12.5, " мм")
            self.sp_st_before = self._make_dspin(0.0, 72.0, 1.0, 0.0, " пт")
            self.sp_st_after = self._make_dspin(0.0, 72.0, 1.0, 0.0, " пт")

            g.addWidget(QLabel("Межстрочный интервал:"), 2, 0)
            g.addWidget(self.sp_st_line, 2, 1)
            g.addWidget(QLabel("Отступ первой строки:"), 2, 2)
            g.addWidget(self.sp_st_indent, 2, 3)

            g.addWidget(QLabel("Интервал перед:"), 3, 0)
            g.addWidget(self.sp_st_before, 3, 1)
            g.addWidget(QLabel("Интервал после:"), 3, 2)
            g.addWidget(self.sp_st_after, 3, 3)

            self.combo_st_outline = QComboBox()
            self.combo_st_outline.addItem("Обычный абзац (нет)", None)
            self.combo_st_outline.addItem("Уровень 1 (Заголовок 1)", 0)
            self.combo_st_outline.addItem("Уровень 2 (Заголовок 2)", 1)
            self.combo_st_outline.addItem("Уровень 3 (Заголовок 3)", 2)
            self.combo_st_outline.currentIndexChanged.connect(self._on_style_prop_changed)
            g.addWidget(QLabel("Уровень структуры:"), 4, 0)
            g.addWidget(self.combo_st_outline, 4, 1, 1, 3)

            for sp in (self.sp_st_size, self.sp_st_line, self.sp_st_indent, self.sp_st_before, self.sp_st_after):
                sp.valueChanged.connect(self._on_style_prop_changed)
            lay.addLayout(g)

            flags_row = QHBoxLayout()
            flags_row.setSpacing(14)
            self.chk_st_bold = QCheckBox("Полужирный")
            self.chk_st_italic = QCheckBox("Курсив")
            self.chk_st_caps = QCheckBox("ПРОПИСНЫЕ")
            self.chk_st_pb = QCheckBox("С новой страницы")
            self.chk_st_kwn = QCheckBox("Не отрывать от следующего")
            for c in (self.chk_st_bold, self.chk_st_italic, self.chk_st_caps, self.chk_st_pb, self.chk_st_kwn):
                c.toggled.connect(self._on_style_prop_changed)
                flags_row.addWidget(c)
            flags_row.addStretch()
            lay.addLayout(flags_row)

            lay.addStretch()
            return w

        # --- Section 3: Tables ---
        def _build_tables_section(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            lbl = QLabel("Оформление таблиц")
            lbl.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            lay.addWidget(lbl)

            g = QGridLayout()
            g.setSpacing(10)
            self.combo_tbl_border = QComboBox()
            self.combo_tbl_border.addItem("Полная сетка (ГОСТ)", "full_grid")
            self.combo_tbl_border.addItem("Открытая книга (IEEE / ВАК)", "gost_open")
            self.combo_tbl_border.currentIndexChanged.connect(self._on_tables_changed)
            g.addWidget(QLabel("Стиль границ:"), 0, 0)
            g.addWidget(self.combo_tbl_border, 0, 1)

            shade_row = QHBoxLayout()
            self.chk_tbl_shade = QCheckBox("Заливка шапки:")
            self.chk_tbl_shade.toggled.connect(self._on_tables_changed)
            self.ed_tbl_shade = QLineEdit("#F1F5F9")
            self._attach_ctx(self.ed_tbl_shade)
            self.ed_tbl_shade.textEdited.connect(self._on_tables_changed)
            shade_row.addWidget(self.chk_tbl_shade)
            shade_row.addWidget(self.ed_tbl_shade, 1)
            g.addLayout(shade_row, 0, 2, 1, 2)

            self.sp_pad_t = self._make_dspin(0.0, 10.0, 0.2, 1.0, " мм")
            self.sp_pad_b = self._make_dspin(0.0, 10.0, 0.2, 1.0, " мм")
            self.sp_pad_l = self._make_dspin(0.0, 10.0, 0.2, 1.8, " мм")
            self.sp_pad_r = self._make_dspin(0.0, 10.0, 0.2, 1.8, " мм")
            g.addWidget(QLabel("Отступ сверху:"), 1, 0)
            g.addWidget(self.sp_pad_t, 1, 1)
            g.addWidget(QLabel("Отступ снизу:"), 1, 2)
            g.addWidget(self.sp_pad_b, 1, 3)
            g.addWidget(QLabel("Отступ слева:"), 2, 0)
            g.addWidget(self.sp_pad_l, 2, 1)
            g.addWidget(QLabel("Отступ справа:"), 2, 2)
            g.addWidget(self.sp_pad_r, 2, 3)
            for sp in (self.sp_pad_t, self.sp_pad_b, self.sp_pad_l, self.sp_pad_r):
                sp.valueChanged.connect(self._on_tables_changed)
            lay.addLayout(g)

            self.chk_tbl_double = QCheckBox("Двойная нижняя граница под шапкой таблицы")
            self.chk_tbl_repeat = QCheckBox("Повторять шапку при переносе на следующую страницу")
            self.chk_tbl_split = QCheckBox("Запретить разрыв строк таблицы (cant_split)")
            for c in (self.chk_tbl_double, self.chk_tbl_repeat, self.chk_tbl_split):
                c.toggled.connect(self._on_tables_changed)
                lay.addWidget(c)

            lay.addStretch()
            return w

        # --- Section 4: Lists ---
        def _build_lists_section(self) -> QWidget:
            w = QWidget()
            lay = QVBoxLayout(w)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(12)

            lbl = QLabel("Параметры списков Word")
            lbl.setStyleSheet("font-size: 14.5px; font-weight: 600;")
            lay.addWidget(lbl)

            g = QGridLayout()
            g.setSpacing(10)
            self.sp_lst_base = self._make_dspin(0.0, 40.0, 0.5, 12.5, " мм")
            self.sp_lst_step = self._make_dspin(1.0, 25.0, 0.5, 6.0, " мм")
            self.sp_lst_hdash = self._make_dspin(1.0, 20.0, 0.5, 4.5, " мм")
            self.sp_lst_hbul = self._make_dspin(1.0, 20.0, 0.5, 4.5, " мм")
            self.sp_lst_hord = self._make_dspin(1.0, 20.0, 0.5, 6.5, " мм")
            self.sp_lst_htask = self._make_dspin(1.0, 20.0, 0.5, 5.5, " мм")
            self.sp_lst_sb = self._make_dspin(0.0, 24.0, 0.5, 0.0, " пт")
            self.sp_lst_sa = self._make_dspin(0.0, 24.0, 0.5, 0.0, " пт")

            g.addWidget(QLabel("Базовый отступ:"), 0, 0)
            g.addWidget(self.sp_lst_base, 0, 1)
            g.addWidget(QLabel("Шаг вложенности:"), 0, 2)
            g.addWidget(self.sp_lst_step, 0, 3)

            g.addWidget(QLabel("Выступ тире:"), 1, 0)
            g.addWidget(self.sp_lst_hdash, 1, 1)
            g.addWidget(QLabel("Выступ маркера:"), 1, 2)
            g.addWidget(self.sp_lst_hbul, 1, 3)

            g.addWidget(QLabel("Выступ нумерации:"), 2, 0)
            g.addWidget(self.sp_lst_hord, 2, 1)
            g.addWidget(QLabel("Выступ чекбокса:"), 2, 2)
            g.addWidget(self.sp_lst_htask, 2, 3)

            g.addWidget(QLabel("Интервал перед:"), 3, 0)
            g.addWidget(self.sp_lst_sb, 3, 1)
            g.addWidget(QLabel("Интервал после:"), 3, 2)
            g.addWidget(self.sp_lst_sa, 3, 3)

            self.ed_lst_dash = QLineEdit("–")
            self._attach_ctx(self.ed_lst_dash)
            self.ed_lst_dash.textEdited.connect(self._on_lists_changed)
            self.ed_lst_bullets = QLineEdit("•, ◦, ▪")
            self._attach_ctx(self.ed_lst_bullets)
            self.ed_lst_bullets.textEdited.connect(self._on_lists_changed)
            self.combo_lst_sep = QComboBox()
            self.combo_lst_sep.addItem("Неразрывный пробел (ГОСТ)", "\u00a0")
            self.combo_lst_sep.addItem("Обычный пробел", " ")
            self.combo_lst_sep.addItem("Табуляция", "\t")
            self.combo_lst_sep.currentIndexChanged.connect(self._on_lists_changed)

            g.addWidget(QLabel("Символ тире:"), 4, 0)
            g.addWidget(self.ed_lst_dash, 4, 1)
            g.addWidget(QLabel("Маркеры уровней:"), 4, 2)
            g.addWidget(self.ed_lst_bullets, 4, 3)
            g.addWidget(QLabel("Разделитель маркера:"), 5, 0)
            g.addWidget(self.combo_lst_sep, 5, 1, 1, 3)

            for sp in (self.sp_lst_base, self.sp_lst_step, self.sp_lst_hdash, self.sp_lst_hbul, self.sp_lst_hord, self.sp_lst_htask, self.sp_lst_sb, self.sp_lst_sa):
                sp.valueChanged.connect(self._on_lists_changed)
            lay.addLayout(g)
            lay.addStretch()
            return w

        def _make_dspin(self, mn: float, mx: float, step: float, val: float, suffix: str = "") -> QDoubleSpinBox:
            sp = QDoubleSpinBox()
            sp.setRange(mn, mx)
            sp.setSingleStep(step)
            sp.setDecimals(2 if step < 0.5 else 1)
            sp.setValue(val)
            if suffix:
                sp.setSuffix(suffix)
            return sp

        # --- Data binding ---
        def apply_theme(self):
            t = self._t()
            if hasattr(self, "tpl_picker_menu"):
                self.tpl_picker_menu.setWindowFlags(self.tpl_picker_menu.windowFlags() | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
                self.tpl_picker_menu.setAttribute(Qt.WA_TranslucentBackground, True)
                self.tpl_picker_menu.setStyleSheet(
                    f"QMenu {{ background-color: {t['bg_window']}; border: 1px solid {t['border_strong']}; "
                    f"border-radius: 10px; padding: 6px; color: {t['text']}; }}"
                    f"QMenu::item {{ padding: 7px 14px; border-radius: 6px; font-size: 13px; color: {t['text']}; }}"
                    f"QMenu::item:selected {{ background-color: {t['accent_soft']}; color: {t['accent_text']}; }}"
                )
            self._mark_dirty(self._dirty)
            self._update_style_preview()
            if hasattr(self, "ed_st_color"):
                self._sync_color_swatch(self.ed_st_color.text().strip() or "#000000")

        def load_templates(self, active_tpl_id: str = ""):
            self._loading = True
            tpls = self._templates_provider()
            self.tpl_picker_menu.clear()
            target_tid = active_tpl_id if active_tpl_id in tpls else (next(iter(tpls)) if tpls else "")
            for tid, cfg in tpls.items():
                title = cfg.get("display_title", tid)
                act = self.tpl_picker_menu.addAction(title)
                act.setCheckable(True)
                act.setChecked(tid == target_tid)
                act.triggered.connect(lambda _, t_id=tid: self._on_tpl_picked(t_id))
            if target_tid and target_tid in tpls:
                self.btn_tpl_picker.setText(f"{tpls[target_tid].get('display_title', target_tid)} ▾")
                self._load_template_data(target_tid, tpls[target_tid])
            self._loading = False

        def _on_tpl_picked(self, tid: str):
            tpls = self._templates_provider()
            if tid and tid in tpls:
                self.btn_tpl_picker.setText(f"{tpls[tid].get('display_title', tid)} ▾")
                for act, (k, _) in zip(self.tpl_picker_menu.actions(), tpls.items()):
                    act.setChecked(k == tid)
                self._load_template_data(tid, tpls[tid])

        def _load_template_data(self, tid: str, cfg: dict):
            self._loading = True
            self._current_tpl_id = tid
            self._working_cfg = copy.deepcopy(cfg)

            # Meta & Page
            self.ed_tpl_title.setText(str(self._working_cfg.get("display_title", tid)))
            self.ed_tpl_desc.setText(str(self._working_cfg.get("description", "")))
            pg = self._working_cfg.setdefault("page", {})
            self.sp_pw.setValue(float(pg.get("width_mm", 210.0)))
            self.sp_ph.setValue(float(pg.get("height_mm", 297.0)))
            self.sp_ml.setValue(float(pg.get("margin_left_mm", 30.0)))
            self.sp_mr.setValue(float(pg.get("margin_right_mm", 20.0)))
            self.sp_mt.setValue(float(pg.get("margin_top_mm", 20.0)))
            self.sp_mb.setValue(float(pg.get("margin_bottom_mm", 20.0)))
            self.sp_cols.setValue(int(pg.get("columns", 1) or 1))
            self.sp_col_gap.setValue(float(pg.get("column_spacing_mm", 5.0)))
            self.chk_diff_first.setChecked(bool(pg.get("different_first_page", False)))
            self.chk_page_nums.setChecked(bool(pg.get("page_numbers", False)))
            self.chk_span_h.setChecked(bool(pg.get("span_headings", True)))
            self.chk_tpl_strip_em.setChecked(bool(self._working_cfg.get("strip_body_emphasis", False)))

            # Fonts & Structural titles
            fnt = self._working_cfg.setdefault("fonts", {})
            self.combo_font_primary.setCurrentText(str(fnt.get("primary", {}).get("name", "Times New Roman")))
            self.combo_font_mono.setCurrentText(str(fnt.get("monospace", {}).get("name", "Courier New")))
            self.ed_struct_titles.setPlainText("\n".join(self._working_cfg.get("structural_titles", [])))

            # Tables
            tbl = self._working_cfg.setdefault("tables", {})
            idx_tb = self.combo_tbl_border.findData(tbl.get("border_style", "full_grid"))
            self.combo_tbl_border.setCurrentIndex(max(0, idx_tb))
            shade = tbl.get("header_shading")
            self.chk_tbl_shade.setChecked(bool(shade))
            self.ed_tbl_shade.setText(str(shade or "#F1F5F9"))
            self.ed_tbl_shade.setEnabled(bool(shade))
            pad = tbl.get("cell_padding_mm", {})
            self.sp_pad_t.setValue(float(pad.get("top", 1.0)))
            self.sp_pad_b.setValue(float(pad.get("bottom", 1.0)))
            self.sp_pad_l.setValue(float(pad.get("left", 1.8)))
            self.sp_pad_r.setValue(float(pad.get("right", 1.8)))
            self.chk_tbl_double.setChecked(bool(tbl.get("header_double_border", True)))
            self.chk_tbl_repeat.setChecked(bool(tbl.get("repeat_header", True)))
            self.chk_tbl_split.setChecked(bool(tbl.get("cant_split", True)))

            # Lists
            lst = self._working_cfg.setdefault("lists", {})
            self.sp_lst_base.setValue(float(lst.get("base_indent_mm", 12.5)))
            self.sp_lst_step.setValue(float(lst.get("level_step_mm", 6.0)))
            hang = lst.get("hanging_indent_mm", {})
            if isinstance(hang, (int, float)):
                hang = {"dash": hang, "bullet": hang, "ordered": hang, "task": hang}
            self.sp_lst_hdash.setValue(float(hang.get("dash", 4.5)))
            self.sp_lst_hbul.setValue(float(hang.get("bullet", 4.5)))
            self.sp_lst_hord.setValue(float(hang.get("ordered", 6.5)))
            self.sp_lst_htask.setValue(float(hang.get("task", 5.5)))
            self.sp_lst_sb.setValue(float(lst.get("space_before_pt", 0.0)))
            self.sp_lst_sa.setValue(float(lst.get("space_after_pt", 0.0)))
            self.ed_lst_dash.setText(str(lst.get("dash_symbol", "–")))
            self.ed_lst_bullets.setText(", ".join(lst.get("bullet_symbols", ["•", "◦", "▪"])))
            sep_idx = self.combo_lst_sep.findData(lst.get("marker_separator", "\u00a0"))
            self.combo_lst_sep.setCurrentIndex(max(0, sep_idx))

            # Styles list
            self._populate_style_list()
            self._loading = False
            self._mark_dirty(False)

        def _populate_style_list(self, select_key: str = ""):
            self.list_styles.clear()
            styles = self._working_cfg.setdefault("styles", {})
            ordered = [k for k in CANONICAL_STYLE_ORDER if k in styles]
            for k in styles.keys():
                if k not in ordered:
                    ordered.append(k)
            target_row = 0
            for idx, sk in enumerate(ordered):
                self.list_styles.addItem(QListWidgetItem(sk))
                if sk == select_key:
                    target_row = idx
            if ordered:
                self.list_styles.setCurrentRow(target_row)
                self._load_style_into_form(ordered[target_row])

        def _on_style_selected(self, current: QListWidgetItem, previous: QListWidgetItem):
            if self._loading or current is None:
                return
            self._switch_section(2)
            self._load_style_into_form(current.text())

        def _load_style_into_form(self, style_key: str):
            self._loading = True
            self._current_style_key = style_key
            s_def = self._working_cfg.get("styles", {}).get(style_key, {})
            self.lbl_style_title.setText(f"Стиль: {style_key}")

            word_map_hint = {
                "Заголовок структурный": "Word: Title",
                "Заголовок 1": "Word: Heading 1",
                "Заголовок 2": "Word: Heading 2",
                "Заголовок 3": "Word: Heading 3",
                "Основной текст": "Word: Normal",
                "Обычный": "Word: Normal",
                "Подпись рисунка": "Word: Caption",
                "Подпись к рисунку": "Word: Caption",
                "Цитата": "Word: Quote",
            }.get(style_key, "Пользовательский стиль")
            self.lbl_style_word_map.setText(word_map_hint)

            f_idx = self.combo_st_font.findData(s_def.get("font", "primary"))
            self.combo_st_font.setCurrentIndex(max(0, f_idx))
            self.sp_st_size.setValue(float(s_def.get("size_pt", 14.0)))
            a_idx = self.combo_st_align.findData(s_def.get("alignment", "justify"))
            self.combo_st_align.setCurrentIndex(max(0, a_idx))

            col_val = str(s_def.get("color", "#000000") or "#000000")
            if not col_val.startswith("#"):
                col_val = "#" + col_val
            self.ed_st_color.setText(col_val.upper())
            self._sync_color_swatch(col_val)

            self.sp_st_line.setValue(float(s_def.get("line_spacing", 1.5)))
            self.sp_st_indent.setValue(float(s_def.get("first_line_indent_mm", 0.0)))
            self.sp_st_before.setValue(float(s_def.get("space_before_pt", 0.0)))
            self.sp_st_after.setValue(float(s_def.get("space_after_pt", 0.0)))

            ol = s_def.get("outline_level", None)
            ol_idx = self.combo_st_outline.findData(ol)
            self.combo_st_outline.setCurrentIndex(max(0, ol_idx))

            self.chk_st_bold.setChecked(bool(s_def.get("bold", False)))
            self.chk_st_italic.setChecked(bool(s_def.get("italic", False)))
            self.chk_st_caps.setChecked(bool(s_def.get("all_caps", False)))
            self.chk_st_pb.setChecked(bool(s_def.get("page_break_before", False)))
            self.chk_st_kwn.setChecked(bool(s_def.get("keep_with_next", False)))

            self._loading = False
            self._update_style_preview()

        def _sync_color_swatch(self, hex_col: str):
            c = QColor(hex_col)
            if not c.isValid():
                c = QColor("#000000")
            self.btn_st_color.setStyleSheet(
                f"background-color: {c.name()}; border: 1px solid #94A3B8; border-radius: 6px;"
            )

        def _pick_style_color(self):
            cur = QColor(self.ed_st_color.text().strip() or "#000000")
            chosen = QColorDialog.getColor(cur, self, "Выберите цвет текста")
            if chosen.isValid():
                self.ed_st_color.setText(chosen.name().upper())
                self._on_style_prop_changed()

        def _update_style_preview(self):
            if not hasattr(self, "lbl_style_preview"):
                return
            s_def = self._working_cfg.get("styles", {}).get(self._current_style_key, {})
            f_key = s_def.get("font", "primary")
            f_name = self._working_cfg.get("fonts", {}).get(f_key, {}).get("name", "Times New Roman")
            sz = max(9, min(22, int(round(float(s_def.get("size_pt", 14.0))))))
            col = str(s_def.get("color", "#000000") or "#000000")
            if not col.startswith("#"):
                col = "#" + col
            if not QColor(col).isValid():
                col = "#000000"
            weight = "700" if s_def.get("bold") else "400"
            fstyle = "italic" if s_def.get("italic") else "normal"
            align_map = {
                "left": Qt.AlignLeft,
                "center": Qt.AlignHCenter,
                "right": Qt.AlignRight,
                "justify": Qt.AlignJustify,
            }
            self.lbl_style_preview.setAlignment(align_map.get(s_def.get("alignment", "justify"), Qt.AlignLeft))
            txt = "Пример абзаца: Автоматизированная система проектирования по ГОСТ 7.32-2017."
            if s_def.get("all_caps"):
                txt = txt.upper()
            self.lbl_style_preview.setText(txt)
            t = self._t()
            # Лист документа в предпросмотре: светлая бумага с адаптивной рамкой, чтобы цвет текста из шаблона (#000000 и др.) всегда читался
            self.lbl_style_preview.setStyleSheet(
                f"font-family: '{f_name}', serif; font-size: {sz}px; font-weight: {weight}; "
                f"font-style: {fstyle}; color: {col}; background-color: #FFFFFF; "
                f"padding: 10px 12px; border-radius: 6px; border: 1px solid {t['border_strong']};"
            )

        # --- Handlers updating _working_cfg ---
        def _on_meta_changed(self):
            if self._loading:
                return
            self._working_cfg["display_title"] = self.ed_tpl_title.text().strip() or self._current_tpl_id
            self._working_cfg["description"] = self.ed_tpl_desc.text().strip()
            self._mark_dirty(True)

        def _on_page_changed(self):
            if self._loading:
                return
            pg = self._working_cfg.setdefault("page", {})
            pg["width_mm"] = round(self.sp_pw.value(), 1)
            pg["height_mm"] = round(self.sp_ph.value(), 1)
            pg["margin_left_mm"] = round(self.sp_ml.value(), 1)
            pg["margin_right_mm"] = round(self.sp_mr.value(), 1)
            pg["margin_top_mm"] = round(self.sp_mt.value(), 1)
            pg["margin_bottom_mm"] = round(self.sp_mb.value(), 1)
            pg["columns"] = int(self.sp_cols.value())
            pg["column_spacing_mm"] = round(self.sp_col_gap.value(), 1)
            pg["different_first_page"] = self.chk_diff_first.isChecked()
            pg["page_numbers"] = self.chk_page_nums.isChecked()
            pg["span_headings"] = self.chk_span_h.isChecked()
            self._working_cfg["strip_body_emphasis"] = self.chk_tpl_strip_em.isChecked()
            self._mark_dirty(True)

        def _on_fonts_changed(self):
            if self._loading:
                return
            p_name = self.combo_font_primary.currentText().strip() or "Times New Roman"
            m_name = self.combo_font_mono.currentText().strip() or "Courier New"
            self._working_cfg["fonts"] = {
                "primary": {"name": p_name, "ascii": p_name, "hAnsi": p_name, "cs": p_name, "eastAsia": p_name},
                "monospace": {"name": m_name, "ascii": m_name, "hAnsi": m_name, "cs": m_name, "eastAsia": m_name},
            }
            self._update_style_preview()
            self._mark_dirty(True)

        def _on_struct_titles_changed(self):
            if self._loading:
                return
            lines = [l.strip().lower() for l in self.ed_struct_titles.toPlainText().splitlines() if l.strip()]
            self._working_cfg["structural_titles"] = lines
            self._mark_dirty(True)

        def _on_style_prop_changed(self):
            if self._loading or not self._current_style_key:
                return
            s_def = self._working_cfg.setdefault("styles", {}).setdefault(self._current_style_key, {})
            s_def["font"] = self.combo_st_font.currentData() or "primary"
            s_def["size_pt"] = round(self.sp_st_size.value(), 2)
            s_def["alignment"] = self.combo_st_align.currentData() or "justify"
            col = self.ed_st_color.text().strip()
            if re.match(r'^#[0-9A-Fa-f]{6}$', col):
                s_def["color"] = col.upper()
                self._sync_color_swatch(col)
            s_def["line_spacing"] = round(self.sp_st_line.value(), 2)
            s_def["first_line_indent_mm"] = round(self.sp_st_indent.value(), 1)
            s_def["space_before_pt"] = round(self.sp_st_before.value(), 1)
            s_def["space_after_pt"] = round(self.sp_st_after.value(), 1)
            ol = self.combo_st_outline.currentData()
            if ol is not None:
                s_def["outline_level"] = int(ol)
            elif "outline_level" in s_def:
                del s_def["outline_level"]
            s_def["bold"] = self.chk_st_bold.isChecked()
            s_def["italic"] = self.chk_st_italic.isChecked()
            s_def["all_caps"] = self.chk_st_caps.isChecked()
            s_def["page_break_before"] = self.chk_st_pb.isChecked()
            s_def["keep_with_next"] = self.chk_st_kwn.isChecked()
            self._update_style_preview()
            self._mark_dirty(True)

        def _on_tables_changed(self):
            if self._loading:
                return
            tbl = self._working_cfg.setdefault("tables", {})
            tbl["border_style"] = self.combo_tbl_border.currentData() or "full_grid"
            use_shade = self.chk_tbl_shade.isChecked()
            self.ed_tbl_shade.setEnabled(use_shade)
            tbl["header_shading"] = (self.ed_tbl_shade.text().strip() or "#F1F5F9") if use_shade else None
            tbl["cell_padding_mm"] = {
                "top": round(self.sp_pad_t.value(), 2),
                "bottom": round(self.sp_pad_b.value(), 2),
                "left": round(self.sp_pad_l.value(), 2),
                "right": round(self.sp_pad_r.value(), 2),
            }
            tbl["header_double_border"] = self.chk_tbl_double.isChecked()
            tbl["repeat_header"] = self.chk_tbl_repeat.isChecked()
            tbl["cant_split"] = self.chk_tbl_split.isChecked()
            self._mark_dirty(True)

        def _on_lists_changed(self):
            if self._loading:
                return
            lst = self._working_cfg.setdefault("lists", {})
            lst["base_indent_mm"] = round(self.sp_lst_base.value(), 1)
            lst["level_step_mm"] = round(self.sp_lst_step.value(), 1)
            lst["hanging_indent_mm"] = {
                "dash": round(self.sp_lst_hdash.value(), 1),
                "bullet": round(self.sp_lst_hbul.value(), 1),
                "ordered": round(self.sp_lst_hord.value(), 1),
                "task": round(self.sp_lst_htask.value(), 1),
            }
            lst["space_before_pt"] = round(self.sp_lst_sb.value(), 1)
            lst["space_after_pt"] = round(self.sp_lst_sa.value(), 1)
            lst["dash_symbol"] = self.ed_lst_dash.text().strip() or "–"
            lst["bullet_symbols"] = [s.strip() for s in self.ed_lst_bullets.text().split(",") if s.strip()] or ["•"]
            lst["marker_separator"] = self.combo_lst_sep.currentData() or "\u00a0"
            self._mark_dirty(True)

        # --- Template & Style CRUD ---
        def _ask_name_dialog(self, title: str, label: str, default: str = "") -> Optional[str]:
            dlg = QDialog(self)
            dlg.setWindowTitle(title)
            dlg.setMinimumWidth(360)
            t = self._t()
            dlg.setStyleSheet(
                f"QDialog {{ background-color: {t['bg_window']}; color: {t['text']}; }}"
                f"QLabel {{ color: {t['text']}; }}"
            )
            lay = QVBoxLayout(dlg)
            lay.setContentsMargins(18, 16, 18, 16)
            lay.setSpacing(10)
            lay.addWidget(QLabel(label))
            ed = QLineEdit(default)
            self._attach_ctx(ed)
            lay.addWidget(ed)
            row = QHBoxLayout()
            row.addStretch()
            b_cancel = QPushButton("Отмена")
            b_cancel.setProperty("class", "SecondaryBtn")
            b_cancel.clicked.connect(dlg.reject)
            b_ok = QPushButton("Готово")
            b_ok.setProperty("class", "PrimaryBtn")
            b_ok.clicked.connect(dlg.accept)
            row.addWidget(b_cancel)
            row.addWidget(b_ok)
            lay.addLayout(row)
            if dlg.exec() == QDialog.Accepted:
                val = ed.text().strip()
                return val or None
            return None

        def _create_new_template(self):
            title = self._ask_name_dialog("Новый шаблон", "Название нового шаблона:", "Мой шаблон ГОСТ")
            if not title:
                return
            slug = re.sub(r'[^a-z0-9_]+', '_', title.lower())
            slug = slug.strip('_') or f"custom_tpl_{len(self._templates_provider()) + 1}"
            tpl_cfg = create_blank_template(slug, title)
            self._save_cfg_to_disk(tpl_cfg)
            self.template_saved.emit(slug)
            self.load_templates(slug)
            self.toast_requested.emit(f"Создан шаблон: {title}")

        def _duplicate_template(self):
            cur_title = self._working_cfg.get("display_title", self._current_tpl_id)
            title = self._ask_name_dialog("Копия шаблона", "Название копии:", f"{cur_title} (копия)")
            if not title:
                return
            slug = re.sub(r'[^a-z0-9_]+', '_', title.lower()).strip('_')
            if not slug or slug in self._templates_provider():
                slug = f"{self._current_tpl_id}_copy_{len(self._templates_provider()) + 1}"
            dup = strip_internal_keys(self._working_cfg)
            dup["name"] = slug
            dup["display_title"] = title
            self._save_cfg_to_disk(dup)
            self.template_saved.emit(slug)
            self.load_templates(slug)
            self.toast_requested.emit(f"Создана копия: {title}")

        def _add_custom_style(self):
            name = self._ask_name_dialog("Новый стиль", "Имя стиля (например: Листинг Примечание):")
            if not name:
                return
            styles = self._working_cfg.setdefault("styles", {})
            if name not in styles:
                styles[name] = copy.deepcopy(DEFAULT_NEW_STYLE)
                self._mark_dirty(True)
            self._populate_style_list(select_key=name)
            self._switch_section(2)
            self.toast_requested.emit(f"Добавлен стиль: {name}")

        def _delete_selected_style(self):
            if not self._current_style_key:
                return
            styles = self._working_cfg.get("styles", {})
            if len(styles) <= 1:
                self.toast_requested.emit("В шаблоне должен остаться хотя бы один стиль")
                return
            removed = self._current_style_key
            styles.pop(removed, None)
            self._mark_dirty(True)
            self._populate_style_list()
            self.toast_requested.emit(f"Удалён стиль: {removed}")

        def _save_cfg_to_disk(self, cfg: dict) -> Path:
            clean = strip_internal_keys(cfg)
            t_name = clean.get("name") or "custom_template"
            orig_path = cfg.get("_file_path")
            if orig_path and Path(orig_path).parent.exists():
                target = Path(orig_path)
            else:
                target = self._primary_dir_provider() / f"{t_name}.json"
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text(json.dumps(clean, ensure_ascii=False, indent=4), encoding="utf-8")
            self._register_path_fn(target)
            return target

        def save_current_template(self) -> Optional[str]:
            if not self._working_cfg:
                return None
            # Обновляем specs_summary автоматически для наглядности в карточках
            f_name = self._working_cfg.get("fonts", {}).get("primary", {}).get("name", "Times New Roman")
            body_sz = self._working_cfg.get("styles", {}).get("Основной текст", {}).get("size_pt", 14.0)
            cols = self._working_cfg.get("page", {}).get("columns", 1)
            col_txt = "2 колонки" if int(cols or 1) > 1 else "1 колонка"
            self._working_cfg["specs_summary"] = f"{f_name} • {body_sz:g} пт • {col_txt}"

            self._save_cfg_to_disk(self._working_cfg)
            self._mark_dirty(False)
            tid = self._working_cfg.get("name", self._current_tpl_id)
            self.template_saved.emit(tid)
            self.toast_requested.emit(f"Шаблон сохранён: {self._working_cfg.get('display_title', tid)}")
            return tid
