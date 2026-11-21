#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты функций:
- append_markdown_to_docx (добавление текста в существующий DOCX с сохранением исходного текста, форматирования и стилей документа)
- merge_docx_files (склейка DOCX-файлов с сохранением таблиц и переназначением rId изображений)
- update_docx_styles (изменение стандартных стилей DOCX, например 'Обычный' / 'Normal')
- Регулировка параметров списков (lists: base_indent_mm, level_step_mm, hanging_indent_mm, marker_separator, отступы/интервалы)
- CLI-режим clipboard_docx_runner.py (--append, --merge, --apply-styles)
"""

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from docx import Document
from docx.shared import Pt, Mm, RGBColor
from PIL import Image, ImageDraw

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "gost_engine"))

from gost_docx_parser import MarkdownASTParser, NBSP
from gost_docx_formatter import (
    GostDocxFormatter,
    append_markdown_to_docx,
    merge_docx_files,
    update_docx_styles,
)

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def load_template(name: str) -> dict:
    return json.loads((BASE_DIR / "templates" / f"{name}.json").read_text(encoding="utf-8"))


def make_test_png(path: Path, color: str):
    img = Image.new("RGB", (200, 100), color)
    ImageDraw.Draw(img).rectangle([10, 10, 190, 90], outline="black", width=3)
    img.save(path)


class TestAppendMergeAndLists(unittest.TestCase):
    def test_01_append_preserves_original_text_formatting_and_styles(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            base_docx = tmp / "base.docx"

            # Создаем исходный документ со своим стилем Normal (Arial 11 pt) и локальным форматированием
            doc = Document()
            normal = doc.styles["Normal"]
            normal.font.name = "Arial"
            normal.font.size = Pt(11)
            normal.font.color.rgb = RGBColor(0x11, 0x22, 0x33)

            doc.sections[0].header.paragraphs[0].text = "Исходный колонтитул"
            p1 = doc.add_paragraph("Исходный абзац: ")
            r_bold = p1.add_run("жирный фрагмент")
            r_bold.bold = True
            r_italic = p1.add_run(" и курсив.")
            r_italic.italic = True
            doc.save(base_docx)

            cfg = load_template("dstu_2026_gost_russian_names")
            out_docx = tmp / "appended.docx"
            append_markdown_to_docx(
                str(base_docx),
                "## 1.1 Добавленный подраздел\n\nНовый текст с формулой $E = mc^2$.\n",
                cfg,
                output_path=str(out_docx),
            )

            res = Document(out_docx)
            # Исходный колонтитул и исходный абзац сохранены
            self.assertEqual(res.sections[0].header.paragraphs[0].text, "Исходный колонтитул")
            self.assertEqual(res.paragraphs[0].text, "Исходный абзац: жирный фрагмент и курсив.")
            self.assertTrue(res.paragraphs[0].runs[1].bold)
            self.assertTrue(res.paragraphs[0].runs[2].italic)

            # Стиль Normal НЕ перезаписан шаблоном (остался Arial 11 pt)
            self.assertEqual(res.styles["Normal"].font.name, "Arial")
            self.assertEqual(res.styles["Normal"].font.size.pt, 11.0)

            # Добавленный текст в конце документа использует стиль Normal
            self.assertGreaterEqual(len(res.paragraphs), 3)
            self.assertIn("Добавленный подраздел", res.paragraphs[1].text)
            self.assertEqual(res.paragraphs[2].style.name, "Normal")

    def test_02_update_standard_style_normal_russian_name(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            target = tmp / "doc.docx"

            doc = Document()
            doc.styles["Normal"].font.name = "Calibri"
            doc.styles["Normal"].font.size = Pt(11)
            p = doc.add_paragraph("Сохраняемый текст документа.")
            p.runs[0].italic = True
            doc.save(target)

            update_docx_styles(
                str(target),
                {
                    "Обычный": {
                        "font": "Times New Roman",
                        "size_pt": 14.0,
                        "line_spacing": 1.5,
                        "first_line_indent_mm": 12.5,
                        "alignment": "justify",
                    }
                },
            )

            res = Document(target)
            self.assertEqual(res.paragraphs[0].text, "Сохраняемый текст документа.")
            self.assertTrue(res.paragraphs[0].runs[0].italic)
            st = res.styles["Normal"]
            self.assertEqual(st.font.name, "Times New Roman")
            self.assertEqual(st.font.size.pt, 14.0)
            self.assertAlmostEqual(st.paragraph_format.line_spacing, 1.5, places=2)
            self.assertAlmostEqual(st.paragraph_format.first_line_indent.mm, 12.5, places=1)

    def test_03_merge_docx_files_with_images_and_tables(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            img1 = tmp / "img1.png"
            img2 = tmp / "img2.png"
            make_test_png(img1, "red")
            make_test_png(img2, "blue")

            d1_path = tmp / "doc1.docx"
            d2_path = tmp / "doc2.docx"
            merged_path = tmp / "merged.docx"

            d1 = Document()
            d1.styles["Normal"].font.name = "Georgia"
            d1.add_paragraph("Первый документ")
            d1.add_picture(str(img1), width=Mm(40))
            d1.save(d1_path)

            d2 = Document()
            d2.add_paragraph("Второй документ")
            tbl = d2.add_table(rows=2, cols=2)
            tbl.cell(0, 0).text = "A"
            tbl.cell(1, 1).text = "B"
            d2.add_picture(str(img2), width=Mm(40))
            d2.save(d2_path)

            merge_docx_files(str(d1_path), [str(d2_path)], output_path=str(merged_path), add_page_break=True)

            res = Document(merged_path)
            all_text = "\n".join(p.text for p in res.paragraphs)
            self.assertIn("Первый документ", all_text)
            self.assertIn("Второй документ", all_text)
            self.assertEqual(len(res.tables), 1)
            self.assertEqual(res.tables[0].cell(1, 1).text, "B")
            self.assertEqual(res.styles["Normal"].font.name, "Georgia")

            with zipfile.ZipFile(merged_path) as z:
                media = [n for n in z.namelist() if n.startswith("word/media/")]
                doc_xml = z.read("word/document.xml").decode("utf-8")
            self.assertEqual(len(media), 2)
            self.assertEqual(doc_xml.count("<w:drawing>"), 2)

    def test_04_configurable_list_parameters(self):
        md = (
            "- первый пункт;\n"
            "  * вложенный маркер;\n"
            "1) нумерованный пункт.\n"
        )
        cfg = load_template("dstu_2026_gost_v3")
        cfg["lists"] = {
            "base_indent_mm": 10.0,
            "level_step_mm": 8.0,
            "hanging_indent_mm": {"dash": 5.0, "bullet": 5.0, "ordered": 7.0},
            "dash_symbol": "—",
            "bullet_symbols": ["▪", "◦"],
            "marker_separator": "\t",
            "space_before_pt": 2.0,
            "space_after_pt": 4.0,
        }
        ast = MarkdownASTParser().parse(md)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "custom_lists.docx"
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))

            doc = Document(out_docx)
            self.assertEqual(len(doc.paragraphs), 3)
            p0, p1, p2 = doc.paragraphs
            self.assertEqual(p0.text, "первый пункт;")
            self.assertEqual(p1.text, "вложенный маркер;")
            self.assertEqual(p2.text, "нумерованный пункт.")
            self.assertEqual(p0.paragraph_format.space_before.pt, 2.0)
            self.assertEqual(p0.paragraph_format.space_after.pt, 4.0)

            with zipfile.ZipFile(out_docx) as z:
                num_root = ET.fromstring(z.read("word/numbering.xml").decode("utf-8"))
            # Проверяем настроенные отступы и выступы в numbering.xml:
            # dash lvl 0: left = 15 мм (850 dxa), hanging = 5 мм (283 dxa)
            # bullet lvl 1: left = 23 мм (1304 dxa), hanging = 5 мм (283 dxa)
            # ordered lvl 0: left = 17 мм (964 dxa), hanging = 7 мм (397 dxa)
            inds = [(int(el.get(f"{{{NS['w']}}}left")), int(el.get(f"{{{NS['w']}}}hanging"))) for el in num_root.findall(".//w:lvl/w:pPr/w:ind", NS) if el.get(f"{{{NS['w']}}}left") and el.get(f"{{{NS['w']}}}hanging")]
            self.assertIn((850, 283), inds)
            self.assertIn((1304, 283), inds)
            self.assertIn((964, 397), inds)

    def test_05_cli_append_merge_apply_styles(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            d1 = tmp / "d1.docx"
            d2 = tmp / "d2.docx"
            md = tmp / "add.md"
            out_merged = tmp / "merged.docx"

            doc1 = Document()
            doc1.add_paragraph("Базовый текст")
            doc1.save(d1)

            doc2 = Document()
            doc2.add_paragraph("Дополнительный файл")
            doc2.save(d2)

            md.write_text("# 1 Раздел из CLI\n\nТекст из файла.\n", encoding="utf-8")

            runner = str(BASE_DIR / "clipboard_docx_runner.py")
            r1 = subprocess.run(
                [sys.executable, runner, "--append", str(d1), str(md)],
                capture_output=True, text=True, check=True,
            )
            self.assertIn("ГОТОВО (append)", r1.stdout)

            r2 = subprocess.run(
                [sys.executable, runner, "--merge", str(d1), str(d2), "-o", str(out_merged)],
                capture_output=True, text=True, check=True,
            )
            self.assertIn("ГОТОВО (merge)", r2.stdout)

            r3 = subprocess.run(
                [sys.executable, runner, "--apply-styles", str(out_merged), "--template", "dstu_2026_gost_v3"],
                capture_output=True, text=True, check=True,
            )
            self.assertIn("ГОТОВО (apply-styles)", r3.stdout)

            final_doc = Document(out_merged)
            txt = "\n".join(p.text for p in final_doc.paragraphs)
            self.assertIn("Базовый текст", txt)
            self.assertIn(f"Текст из{NBSP}файла.", txt)
            self.assertIn("Дополнительный файл", txt)
            self.assertEqual(final_doc.styles["Normal"].font.name, "Times New Roman")
            self.assertEqual(final_doc.styles["Normal"].font.size.pt, 14.0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
