#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Комплексные unit- и интеграционные тесты новых возможностей и исправлений:
1. Компиляция диаграмм mermaid, plantuml, dot внешним CLI-рендерером в стиль Рисунок
2. Сохранение исходного символа маркера списка (-, *, +, ., )) и уровней вложенности
3. Многоколоночная верстка (columns: 2, column_spacing_mm: 5.0) с переключением в 1 колонку
4. Поддержка первой особой страницы (different_first_page: true)
5. Декларативный блок tables (gost_open / full_grid, cell_padding_mm, header_shading, tblHeader, cantSplit)
6. Нумерованные выключные формулы с двумя позициями табуляции (центр + правый край)
7. Автоматический расчет ширин столбцов таблиц (min unbreakable + 75-й перцентиль)
8. Автоматическое связывание соседних блоков (keep_with_next + классификация подписей)
9. Контекстная типографика (Smart Typography) с исключением кода и формул
"""

import json
import os
import stat
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image, ImageDraw

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "gost_engine"))

from gost_docx_parser import MarkdownASTParser, typography_cleanup, tokenize_inlines, NBSP
from gost_docx_formatter import GostDocxFormatter, compute_smart_column_widths_mm

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}


def load_template(name: str) -> dict:
    p = BASE_DIR / "templates" / f"{name}.json"
    return json.loads(p.read_text(encoding="utf-8"))


class TestNewFeatures(unittest.TestCase):
    def test_01_diagram_cli_compilation_and_fallback(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            # Создаем mock CLI-рендереры, генерирующие валидные высокоDPI PNG
            png_script = tmp / "mock_renderer.py"
            png_script.write_text(
                "import sys\n"
                "from pathlib import Path\n"
                "from PIL import Image, ImageDraw\n"
                "args = sys.argv[1:]\n"
                "out_path = None\n"
                "if '-o' in args:\n"
                "    idx = args.index('-o')\n"
                "    cand = Path(args[idx + 1])\n"
                "    if cand.suffix.lower() == '.png':\n"
                "        out_path = cand\n"
                "    else:\n"
                "        out_path = cand / 'diagram.png'\n"
                "img = Image.new('RGB', (900, 450), 'white')\n"
                "d = ImageDraw.Draw(img)\n"
                "d.rectangle([20, 20, 880, 430], outline='navy', width=4)\n"
                "d.text((60, 200), 'Rendered Diagram 300 DPI', fill='black')\n"
                "img.save(out_path, dpi=(300, 300))\n",
                encoding="utf-8"
            )
            wrapper = tmp / "mock_cli.sh"
            wrapper.write_text(f"#!/bin/sh\nexec {sys.executable} {png_script} \"$@\"\n", encoding="utf-8")
            wrapper.chmod(wrapper.stat().st_mode | stat.S_IEXEC)

            md = (
                "# 1 Архитектура\n\n"
                "```mermaid\n"
                "graph TD\n"
                "  A[Start] --> B[End]\n"
                "```\n\n"
                "Рисунок 1 – Блок-схема алгоритма\n\n"
                "```plantuml\n"
                "@startuml\n"
                "Alice -> Bob: Hello\n"
                "@enduml\n"
                "```\n\n"
                "```dot\n"
                "digraph G { A -> B; }\n"
                "```\n"
            )
            parser = MarkdownASTParser()
            ast = parser.parse(md)
            diag_nodes = [n for n in ast if n['type'] == 'rendered_diagram']
            self.assertEqual([d['lang'] for d in diag_nodes], ['mermaid', 'plantuml', 'dot'])
            self.assertTrue(diag_nodes[0]['keep_with_next'])

            cfg = load_template("dstu_2026_gost_v3")
            cfg['diagrams'] = {
                'dpi': 300,
                'mermaid_cli': [sys.executable, str(png_script)],
                'plantuml_cli': [sys.executable, str(png_script)],
                'dot_cli': [sys.executable, str(png_script)],
            }
            out_docx = tmp / "diagrams.docx"
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))

            with zipfile.ZipFile(out_docx) as z:
                media_files = [n for n in z.namelist() if n.startswith("word/media/")]
                doc_xml = z.read("word/document.xml").decode("utf-8")
            self.assertGreaterEqual(len(media_files), 1)
            self.assertEqual(doc_xml.count("<w:drawing>"), 3)
            self.assertIn('w:val="GOST_Figure"', doc_xml)

            # Проверка фолбэка при отсутствии CLI
            cfg_no_cli = load_template("dstu_2026_gost_v3")
            cfg_no_cli['diagrams'] = {'mermaid_cli': '/nonexistent/mmdc', 'plantuml_cli': '/nonexistent/puml', 'dot_cli': '/nonexistent/dot'}
            out_fb = tmp / "diagrams_fallback.docx"
            GostDocxFormatter(cfg_no_cli).build_document(ast, str(out_fb))
            self.assertTrue(out_fb.exists())

    def test_02_list_markers_preserved_and_multilevel_indents(self):
        md = (
            "- пункт ГОСТ с тире;\n"
            "* маркированный пункт со звездочкой;\n"
            "+ маркированный пункт с плюсом;\n"
            "1. нумерованный с точкой;\n"
            "1) нумерованный со скобкой:\n"
            "   а) буквенный вложенный;\n"
            "      - вложенный третьего уровня.\n"
        )
        parser = MarkdownASTParser()
        ast = parser.parse(md)
        items = [n for n in ast if n['type'] == 'list_item']
        self.assertEqual(len(items), 7)
        self.assertEqual([it['raw_marker'] for it in items], ['-', '*', '+', '1.', '1)', 'а)', '-'])
        self.assertEqual([it['marker_char'] for it in items], ['-', '*', '+', '.', ')', ')', '-'])
        self.assertEqual([it['level'] for it in items], [0, 0, 0, 0, 0, 1, 2])
        self.assertEqual(items[0]['list_kind'], 'gost_dash')
        self.assertEqual(items[1]['list_kind'], 'bullet')
        self.assertEqual(items[3]['list_kind'], 'ordered_dot')
        self.assertEqual(items[4]['list_kind'], 'ordered_paren')
        self.assertEqual(items[5]['list_kind'], 'alpha_paren')

        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "lists.docx"
            cfg = load_template("dstu_2026_gost_russian_names")
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))
            with zipfile.ZipFile(out_docx) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8")
                num_xml = z.read("word/numbering.xml").decode("utf-8")
            root = ET.fromstring(doc_xml)
            num_root = ET.fromstring(num_xml)
            paras = root.findall(".//w:body/w:p", NS)
            self.assertEqual(len(paras), 7)
            # Проверяем нативный <w:numPr> и уровни вложенности <w:ilvl>
            ilvls = [int(p.find("w:pPr/w:numPr/w:ilvl", NS).get(f"{{{NS['w']}}}val")) for p in paras]
            self.assertEqual(ilvls, [0, 0, 0, 0, 0, 1, 2])
            # Проверяем, что в word/numbering.xml есть все форматы (–, •, %1., %1), %2), russianLower)
            lvl_texts = [el.get(f"{{{NS['w']}}}val") for el in num_root.findall(".//w:lvlText", NS)]
            num_fmts = [el.get(f"{{{NS['w']}}}val") for el in num_root.findall(".//w:numFmt", NS)]
            self.assertIn("–", lvl_texts)
            self.assertIn("•", lvl_texts)
            self.assertIn("%1.", lvl_texts)
            self.assertIn("%1)", lvl_texts)
            self.assertIn("%2)", lvl_texts)
            self.assertIn("russianLower", num_fmts)

    def test_03_multicolumn_and_different_first_page(self):
        md = (
            "# 1 Введение в двухколоночную верстку\n\n"
            "Первый абзац в двух колонках.\n\n"
            "## 1.1 Подраздел в колонке\n\n"
            "Текст подраздела.\n\n"
            "Таблица 1 Сводные параметры системы\n\n"
            "| Параметр | Значение | Единица | Примечание |\n"
            "| :--- | :---: | :---: | :--- |\n"
            "| Частота | 100 | МГц | Номинальная |\n\n"
            "Текст после широкой таблицы снова в двух колонках.\n"
        )
        cfg = load_template("ieee_vak_two_column")
        ast = MarkdownASTParser(cfg.get("structural_titles")).parse(md)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "two_col.docx"
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))
            with zipfile.ZipFile(out_docx) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8")
            root = ET.fromstring(doc_xml)
            cols_elems = root.findall(".//w:cols", NS)
            col_nums = [c.get(f"{{{NS['w']}}}num") for c in cols_elems]
            # Должны присутствовать секции как с 1 колонкой (для Заголовка 1 и широкой таблицы), так и с 2 колонками
            self.assertIn("1", col_nums)
            self.assertIn("2", col_nums)
            # Проверяем different_first_page -> <w:titlePg/>
            self.assertIsNotNone(root.find(".//w:titlePg", NS))
            # Проверяем continuous section breaks
            types = [t.get(f"{{{NS['w']}}}val") for t in root.findall(".//w:sectPr/w:type", NS)]
            self.assertIn("continuous", types)

    def test_04_declarative_tables_gost_open_and_full_grid(self):
        md = (
            "Таблица 1 Открытая таблица\n\n"
            "| Код | Наименование | Значение |\n"
            "| --- | --- | --- |\n"
            "| A1 | Первый параметр | 10,5 |\n"
        )
        ast = MarkdownASTParser().parse(md)
        cfg = load_template("ieee_vak_two_column")
        # Проверяем gost_open + header_shading + tblHeader + cantSplit + tblCellMar
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "tbl_open.docx"
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))
            with zipfile.ZipFile(out_docx) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8")
            root = ET.fromstring(doc_xml)
            tbl_borders = root.find(".//w:tblPr/w:tblBorders", NS)
            self.assertEqual(tbl_borders.find("w:left", NS).get(f"{{{NS['w']}}}val"), "none")
            self.assertEqual(tbl_borders.find("w:right", NS).get(f"{{{NS['w']}}}val"), "none")
            self.assertIsNotNone(root.find(".//w:trPr/w:tblHeader", NS))
            self.assertIsNotNone(root.find(".//w:trPr/w:cantSplit", NS))
            shd = root.find(".//w:tcPr/w:shd", NS)
            self.assertIsNotNone(shd)
            self.assertEqual(shd.get(f"{{{NS['w']}}}fill"), "F2F4F8")
            self.assertIsNotNone(root.find(".//w:tblPr/w:tblCellMar", NS))

    def test_05_numbered_formula_tab_stops(self):
        md = (
            "$$E = m c^2 \\tag{1}$$\n\n"
            "$$F = m a \\qquad (2)$$\n"
        )
        ast = MarkdownASTParser().parse(md)
        self.assertEqual(ast[0]['number'], "(1)")
        self.assertEqual(ast[1]['number'], "(2)")

        cfg = load_template("dstu_2026_gost_v3")
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "formulas.docx"
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))
            with zipfile.ZipFile(out_docx) as z:
                doc_xml = z.read("word/document.xml").decode("utf-8")
            root = ET.fromstring(doc_xml)
            # Не должно быть таблиц-контейнеров для формул
            self.assertEqual(len(root.findall(".//w:tbl", NS)), 0)
            tabs = root.findall(".//w:pPr/w:tabs/w:tab", NS)
            vals = [t.get(f"{{{NS['w']}}}val") for t in tabs]
            self.assertEqual(vals, ["center", "right", "center", "right"])
            self.assertEqual(len(root.findall(".//m:oMath", NS)), 2)

    def test_06_smart_table_column_widths(self):
        header = ["№", "Обозначение", "Подробное техническое описание алгоритма"]
        rows = [
            ["1", "$K_{p}$", "Пропорциональный коэффициент регулятора скорости вращения привода"],
            ["2", "$T_{i}$", "Постоянная времени интегрирования контура тока электродвигателя"],
        ]
        widths = compute_smart_column_widths_mm(header, rows, total_w_mm=160.0, cell_size_pt=10.0, h_padding_mm=3.0)
        self.assertAlmostEqual(sum(widths), 160.0, places=2)
        self.assertLess(widths[0], widths[1])
        self.assertLess(widths[1], widths[2])
        self.assertGreater(widths[2], 80.0)

    def test_07_adjacent_block_linking(self):
        md = (
            "Таблица 3 Характеристики привода\n\n"
            "| Параметр | Значение |\n"
            "| --- | --- |\n"
            "| Мощность | 5 кВт |\n\n"
            "![Схема](missing.png)\n\n"
            "Схема функциональная блока управления\n\n"
            "```python\n"
            "print('hello')\n"
            "```\n\n"
            "Листинг 1 Пример программы\n"
        )
        ast = MarkdownASTParser().parse(md)
        types = [n['type'] for n in ast]
        self.assertEqual(types, ['table_caption', 'table', 'figure_image', 'figure_caption', 'code_block', 'listing_caption'])
        self.assertTrue(ast[0]['keep_with_next'])
        self.assertTrue(ast[2]['keep_with_next'])
        self.assertTrue(ast[4]['keep_with_next'])

    def test_08_smart_typography_and_code_math_exclusion(self):
        s = 'А. М. Смирнов в статье описал рис. 1 и табл. 2 с кеглем 12 пт, полями 20 мм, диапазоном 10-20 и выводом - это важно. `12 пт и 10-20` и $10-20$ не трогать.'
        toks = tokenize_inlines(s)
        plain = "".join(t['val'] for t in toks if not t['code'] and not t['math'])
        code_tok = [t['val'] for t in toks if t['code']]
        math_tok = [t['val'] for t in toks if t['math']]

        self.assertIn(f"А.{NBSP}М.{NBSP}Смирнов", plain)
        self.assertIn(f"рис.{NBSP}1", plain)
        self.assertIn(f"табл.{NBSP}2", plain)
        self.assertIn(f"12{NBSP}пт", plain)
        self.assertIn(f"20{NBSP}мм", plain)
        self.assertIn(f"в{NBSP}статье", plain)
        self.assertIn(f"и{NBSP}табл.", plain)
        self.assertIn("10–20", plain)
        self.assertIn(f"выводом{NBSP}— это", plain)

        # Полное исключение инлайн-кода и формул
        self.assertEqual(code_tok, ["12 пт и 10-20"])
        self.assertEqual(math_tok, ["10-20"])

        # Блок кода не подвергается типографике
        md_code = "```python\nx = 10-20  # 12 пт в коде А. М. Смирнов\n```\n"
        ast_code = MarkdownASTParser().parse(md_code)
        self.assertEqual(ast_code[0]['lines'][0], "x = 10-20  # 12 пт в коде А. М. Смирнов")



    def test_09_strip_body_and_list_emphasis(self):
        md = (
            "# 1 Заголовок\n\n"
            "Обычный текст с **жирным** и *курсивом*.\n\n"
            "- **первый** пункт с *курсивом*;\n"
            "1) *второй* пункт с **выделением**.\n"
        )
        cfg = load_template("dstu_2026_gost_v3")
        cfg["strip_body_emphasis"] = True
        ast = MarkdownASTParser().parse(md)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "stripped.docx"
            GostDocxFormatter(cfg).build_document(ast, str(out_docx))
            from docx import Document
            doc = Document(out_docx)
            self.assertTrue(doc.paragraphs[0].runs[0].bold)
            for p_idx in (1, 2, 3):
                for r in doc.paragraphs[p_idx].runs:
                    self.assertFalse(bool(r.bold))
                    self.assertFalse(bool(r.italic))

if __name__ == "__main__":
    unittest.main(verbosity=2)
