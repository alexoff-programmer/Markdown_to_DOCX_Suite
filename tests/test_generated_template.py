#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты для проверки генерации JSON-шаблонов по обновлённому навыку gost-md-bundle:
- Понятные русские названия стилей ('Заголовок 1', 'Основной текст', 'Обычный', 'Подпись рисунка' и т.д.)
- Прямое переопределение встроенных стилей Word (Normal, Heading1, Heading2, Heading3, Title, Caption, Quote)
- Отсутствие призрачных стилей (GOST_Heading_1..3, GOST_Body)
- Отсутствие коллизий w:aliases и утечек Word theme (asciiTheme, themeColor)
- Корректная сборка полного Markdown-документа по сгенерированному шаблону
"""

import json
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image, ImageDraw

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "gost_engine"))

from gost_docx_parser import MarkdownASTParser
from gost_docx_formatter import GostDocxFormatter

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}

THEME_ATTRS = (
    f"{{{W_NS}}}asciiTheme",
    f"{{{W_NS}}}hAnsiTheme",
    f"{{{W_NS}}}csTheme",
    f"{{{W_NS}}}eastAsiaTheme",
    f"{{{W_NS}}}themeColor",
    f"{{{W_NS}}}themeShade",
    f"{{{W_NS}}}themeTint",
)


def make_test_image(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (480, 280), "white")
    draw = ImageDraw.Draw(img)
    draw.rectangle([16, 16, 464, 264], outline="#1B365D", width=4)
    draw.line([40, 220, 180, 120, 320, 150, 440, 60], fill="#0D9488", width=4)
    img.save(path, dpi=(300, 300))


def build_generated_template(use_alt_aliases: bool = False) -> dict:
    body_key = "Обычный" if use_alt_aliases else "Основной текст"
    caption_key = "Подпись к рисунку" if use_alt_aliases else "Подпись рисунка"

    return {
        "name": "generated_cambria_russian_styles",
        "display_title": "Сгенерированный шаблон (Cambria / Русские стили)",
        "description": "Тестовый шаблон по навыку gost-md-bundle с переопределением встроенных стилей Word.",
        "page": {
            "width_mm": 210.0,
            "height_mm": 297.0,
            "margin_left_mm": 25.0,
            "margin_right_mm": 15.0,
            "margin_top_mm": 20.0,
            "margin_bottom_mm": 20.0,
            "different_first_page": True,
        },
        "fonts": {
            "primary": {
                "name": "Cambria",
                "ascii": "Cambria",
                "hAnsi": "Cambria",
                "cs": "Cambria",
                "eastAsia": "Cambria",
            },
            "monospace": {
                "name": "Consolas",
                "ascii": "Consolas",
                "hAnsi": "Consolas",
                "cs": "Consolas",
                "eastAsia": "Consolas",
            },
        },
        "structural_titles": [
            "аннотация",
            "введение",
            "заключение",
            "список использованных источников",
        ],
        "strip_body_emphasis": True,
        "styles": {
            "Заголовок структурный": {
                "font": "primary",
                "size_pt": 15.0,
                "bold": True,
                "color": "#1B365D",
                "alignment": "center",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 12.0,
                "space_after_pt": 10.0,
                "page_break_before": True,
                "keep_with_next": True,
                "outline_level": 0,
            },
            "Заголовок 1": {
                "font": "primary",
                "size_pt": 15.0,
                "bold": True,
                "color": "#1B365D",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 10.0,
                "space_before_pt": 14.0,
                "space_after_pt": 8.0,
                "page_break_before": True,
                "keep_with_next": True,
                "outline_level": 0,
            },
            "Заголовок 2": {
                "font": "primary",
                "size_pt": 13.5,
                "bold": True,
                "color": "#2C5282",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 10.0,
                "space_before_pt": 12.0,
                "space_after_pt": 6.0,
                "page_break_before": False,
                "keep_with_next": True,
                "outline_level": 1,
            },
            "Заголовок 3": {
                "font": "primary",
                "size_pt": 12.5,
                "bold": True,
                "color": "#000000",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 10.0,
                "space_before_pt": 10.0,
                "space_after_pt": 4.0,
                "page_break_before": False,
                "keep_with_next": True,
                "outline_level": 2,
            },
            body_key: {
                "font": "primary",
                "size_pt": 13.0,
                "bold": False,
                "color": "#000000",
                "alignment": "justify",
                "line_spacing": 1.4,
                "first_line_indent_mm": 10.0,
                "space_before_pt": 0.0,
                "space_after_pt": 0.0,
            },
            "Примечание": {
                "font": "primary",
                "size_pt": 11.5,
                "bold": False,
                "color": "#000000",
                "alignment": "justify",
                "line_spacing": 1.15,
                "first_line_indent_mm": 10.0,
                "space_before_pt": 4.0,
                "space_after_pt": 4.0,
            },
            "Цитата": {
                "font": "primary",
                "size_pt": 12.0,
                "bold": False,
                "italic": True,
                "color": "#333333",
                "alignment": "justify",
                "line_spacing": 1.15,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 6.0,
                "space_after_pt": 6.0,
            },
            "Эпиграф": {
                "font": "primary",
                "size_pt": 11.5,
                "bold": False,
                "italic": True,
                "color": "#000000",
                "alignment": "right",
                "line_spacing": 1.15,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 8.0,
                "space_after_pt": 8.0,
            },
            "Название таблицы": {
                "font": "primary",
                "size_pt": 13.0,
                "bold": False,
                "color": "#000000",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 10.0,
                "space_after_pt": 4.0,
                "keep_with_next": True,
            },
            "Ячейка таблицы": {
                "font": "primary",
                "size_pt": 11.0,
                "bold": False,
                "color": "#000000",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 2.0,
                "space_after_pt": 2.0,
            },
            "Рисунок": {
                "font": "primary",
                "size_pt": 13.0,
                "bold": False,
                "color": "#000000",
                "alignment": "center",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 10.0,
                "space_after_pt": 4.0,
                "keep_with_next": True,
            },
            caption_key: {
                "font": "primary",
                "size_pt": 12.0,
                "bold": False,
                "color": "#000000",
                "alignment": "center",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 4.0,
                "space_after_pt": 10.0,
            },
            "Исходный код": {
                "font": "monospace",
                "size_pt": 10.5,
                "bold": False,
                "color": "#000000",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 0.0,
                "space_after_pt": 0.0,
            },
            "Название листинга": {
                "font": "primary",
                "size_pt": 13.0,
                "bold": False,
                "color": "#000000",
                "alignment": "left",
                "line_spacing": 1.0,
                "first_line_indent_mm": 0.0,
                "space_before_pt": 10.0,
                "space_after_pt": 4.0,
                "keep_with_next": True,
            },
        },
        "tables": {
            "border_style": "full_grid",
            "header_double_border": True,
            "header_shading": "#E2E8F0",
            "cell_padding_mm": {
                "top": 1.2,
                "bottom": 1.2,
                "left": 2.0,
                "right": 2.0,
            },
            "repeat_header": True,
            "cant_split": True,
        },
        "lists": {
            "base_indent_mm": 10.0,
            "level_step_mm": 6.0,
            "hanging_indent_mm": {
                "dash": 4.5,
                "bullet": 4.5,
                "ordered": 6.5,
                "task": 5.5,
            },
            "dash_symbol": "–",
            "bullet_symbols": ["•", "◦", "▪"],
            "marker_separator": "\u00a0",
            "space_before_pt": 0.0,
            "space_after_pt": 0.0,
        },
    }


SAMPLE_MARKDOWN = """# Аннотация

В документе приведены результаты **экспериментального** исследования *автоматизированного* стенда.

# 1 Архитектура системы

## 1.1 Математическая модель

Энергия системы вычисляется по формуле:

$$E = mc^2 \\tag{1}$$

### 1.1.1 Основные требования

Система должна обеспечивать:
1) **быстрый** опрос датчиков;
   а) *непрерывную* фильтрацию сигналов;
      - частоту не ниже 200 Гц;
2) диагностику узлов.

> Цитата из технического регламента: все измерения выполняются в установившемся режиме.

Таблица 1 - Параметры стенда

| Параметр | Значение |
| :--- | :---: |
| Частота | 200 Гц |
| Погрешность | 0,5 % |

![График отклика](media/plot.png){width=120mm}

Рисунок 1 - Переходный процесс

Примечание - Данные получены при нормальных условиях.

Листинг 1 - Инициализация модуля
```python
def init_sensor(rate_hz: int) -> None:
    print(rate_hz)
```

# Список использованных источников

1. Смирнов А. М. Методы испытаний приводов. Ростов-на-Дону: ДГТУ, 2025. 180 с.
"""


class TestGeneratedTemplateOverrides(unittest.TestCase):
    def _compile_and_read_xml(self, cfg: dict):
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            make_test_image(tmp / "media" / "plot.png")
            out_docx = tmp / "generated_test.docx"

            ast_nodes = MarkdownASTParser(cfg.get("structural_titles", [])).parse(SAMPLE_MARKDOWN)
            GostDocxFormatter(cfg).build_document(ast_nodes, str(out_docx), base_dir=str(tmp))

            with zipfile.ZipFile(out_docx) as zf:
                styles_root = ET.fromstring(zf.read("word/styles.xml").decode("utf-8"))
                doc_root = ET.fromstring(zf.read("word/document.xml").decode("utf-8"))
                media_files = [n for n in zf.namelist() if n.startswith("word/media/")]
            return styles_root, doc_root, media_files

    def _assert_no_theme_leaks(self, element: ET.Element, context: str):
        for node in element.iter():
            for attr in THEME_ATTRS:
                self.assertNotIn(
                    attr,
                    node.attrib,
                    f"Найден theme-атрибут {attr} внутри {context}: {ET.tostring(node, encoding='unicode')}",
                )

    def _get_style_map(self, styles_root: ET.Element):
        return {
            st.get(f"{{{W_NS}}}styleId"): st
            for st in styles_root.findall("w:style", NS)
            if st.get(f"{{{W_NS}}}styleId")
        }

    def test_01_generated_template_overrides_builtin_word_styles_cleanly(self):
        cfg = build_generated_template(use_alt_aliases=False)
        styles_root, doc_root, media_files = self._compile_and_read_xml(cfg)

        self.assertEqual(len(media_files), 1)
        style_map = self._get_style_map(styles_root)

        # 1. Нет призрачных стилей
        for ghost_id in ("GOST_Heading_1", "GOST_Heading_2", "GOST_Heading_3", "GOST_Body"):
            self.assertNotIn(ghost_id, style_map, f"Обнаружен призрачный стиль {ghost_id}")

        # 2. Встроенные стили Word присутствуют и переопределены параметрами шаблона
        expected_builtin = {
            "Normal": ("Cambria", "26", "000000"),
            "Title": ("Cambria", "30", "1B365D"),
            "Heading1": ("Cambria", "30", "1B365D"),
            "Heading2": ("Cambria", "27", "2C5282"),
            "Heading3": ("Cambria", "25", "000000"),
            "Caption": ("Cambria", "24", "000000"),
            "Quote": ("Cambria", "24", "333333"),
        }
        for style_id, (font_name, sz_half_pt, color_hex) in expected_builtin.items():
            self.assertIn(style_id, style_map, f"Отсутствует встроенный стиль {style_id}")
            st = style_map[style_id]
            rfonts = st.find("./w:rPr/w:rFonts", NS)
            self.assertIsNotNone(rfonts, f"Нет w:rFonts у {style_id}")
            self.assertEqual(rfonts.get(f"{{{W_NS}}}ascii"), font_name)
            self.assertEqual(rfonts.get(f"{{{W_NS}}}hAnsi"), font_name)
            sz = st.find("./w:rPr/w:sz", NS)
            self.assertIsNotNone(sz, f"Нет w:sz у {style_id}")
            self.assertEqual(sz.get(f"{{{W_NS}}}val"), sz_half_pt)
            color = st.find("./w:rPr/w:color", NS)
            self.assertIsNotNone(color, f"Нет w:color у {style_id}")
            self.assertEqual(color.get(f"{{{W_NS}}}val").upper(), color_hex)
            self._assert_no_theme_leaks(st, f"style:{style_id}")

        # 3. Проверка docDefaults (без asciiTheme и с нашими значениями по умолчанию)
        doc_defaults = styles_root.find("./w:docDefaults", NS)
        self.assertIsNotNone(doc_defaults)
        self._assert_no_theme_leaks(doc_defaults, "docDefaults")
        def_rfonts = doc_defaults.find(".//w:rFonts", NS)
        self.assertEqual(def_rfonts.get(f"{{{W_NS}}}ascii"), "Cambria")

        # 4. Нет коллизий w:aliases (два разных styleId не должны иметь одинаковый alias)
        seen_aliases = {}
        for style_id, st in style_map.items():
            aliases_el = st.find("./w:aliases", NS)
            if aliases_el is not None:
                val = aliases_el.get(f"{{{W_NS}}}val", "")
                for alias in [a.strip().lower() for a in val.split(",") if a.strip()]:
                    self.assertNotIn(
                        alias,
                        seen_aliases,
                        f"Коллизия alias '{alias}' между {seen_aliases.get(alias)} и {style_id}",
                    )
                    seen_aliases[alias] = style_id

        # 5. В document.xml реально используются встроенные pStyle (Title, Heading1..3, Normal, Caption, Quote)
        used_pstyles = {
            p.get(f"{{{W_NS}}}val")
            for p in doc_root.findall(".//w:pPr/w:pStyle", NS)
        }
        for required_used in ("Title", "Heading1", "Heading2", "Heading3", "Normal", "Caption", "Quote"):
            self.assertIn(required_used, used_pstyles, f"Стиль {required_used} не применён в document.xml")

        # 6. strip_body_emphasis=True убрал жирный и курсив из обычного текста и списков
        for para in doc_root.findall(".//w:body/w:p", NS):
            pstyle_el = para.find("./w:pPr/w:pStyle", NS)
            pstyle = pstyle_el.get(f"{{{W_NS}}}val") if pstyle_el is not None else "Normal"
            if pstyle in ("Normal", "ListParagraph"):
                for b_el in para.findall("./w:r/w:rPr/w:b", NS):
                    self.assertIn(b_el.get(f"{{{W_NS}}}val"), ("0", "false"), f"Остался жирный в {pstyle}")
                for i_el in para.findall("./w:r/w:rPr/w:i", NS):
                    self.assertIn(i_el.get(f"{{{W_NS}}}val"), ("0", "false"), f"Остался курсив в {pstyle}")

    def test_02_generated_template_russian_aliases_no_collisions(self):
        cfg = build_generated_template(use_alt_aliases=True)
        styles_root, doc_root, _ = self._compile_and_read_xml(cfg)
        style_map = self._get_style_map(styles_root)

        self.assertIn("Normal", style_map)
        self.assertIn("Caption", style_map)
        self.assertEqual(style_map["Normal"].find("./w:rPr/w:sz", NS).get(f"{{{W_NS}}}val"), "26")
        self.assertEqual(style_map["Caption"].find("./w:rPr/w:sz", NS).get(f"{{{W_NS}}}val"), "24")
        self.assertIsNotNone(doc_root.find(".//m:oMath", NS))
        self.assertIsNotNone(doc_root.find(".//w:tbl", NS))


if __name__ == "__main__":
    unittest.main(verbosity=2)
