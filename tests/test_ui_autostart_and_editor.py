#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты новых возможностей:
1. Автозагрузка Windows (реестр HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run, --minimized)
2. Редактор шаблонов (create_blank_template, strip_internal_keys, сохранение и компиляция DOCX)
3. Разные шаблоны для разных файлов в пакетной обработке + 2 колонки и strip_body_emphasis
4. Склейка DOCX с применением 2 колонок и очисткой жирного/курсива (update_docx_styles)
"""

import json
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest import mock

from docx import Document

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "gost_engine"))

import clipboard_docx_runner as cdr
from template_editor_widget import create_blank_template, strip_internal_keys
from gost_docx_parser import MarkdownASTParser
from gost_docx_formatter import GostDocxFormatter, merge_docx_files, update_docx_styles

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


class _FakeWinreg:
    HKEY_CURRENT_USER = object()
    KEY_READ = 1
    KEY_SET_VALUE = 2
    REG_SZ = 1

    def __init__(self):
        self.store = {}

    class _Ctx:
        def __init__(self, reg, path):
            self.reg = reg
            self.path = path

        def __enter__(self):
            return (self.reg, self.path)

        def __exit__(self, *args):
            return False

    def OpenKey(self, root, path, reserved, access):
        if path not in self.store:
            raise FileNotFoundError(path)
        return self._Ctx(self, path)

    def CreateKeyEx(self, root, path, reserved, access):
        self.store.setdefault(path, {})
        return self._Ctx(self, path)

    def SetValueEx(self, ctx, name, reserved, rtype, value):
        reg, path = ctx
        reg.store.setdefault(path, {})[name] = value

    def QueryValueEx(self, ctx, name):
        reg, path = ctx
        if name not in reg.store.get(path, {}):
            raise FileNotFoundError(name)
        return reg.store[path][name], self.REG_SZ

    def DeleteValue(self, ctx, name):
        reg, path = ctx
        if name not in reg.store.get(path, {}):
            raise FileNotFoundError(name)
        del reg.store[path][name]


class TestUIAutostartAndEditor(unittest.TestCase):
    def test_01_windows_autostart_registry_enable_and_disable(self):
        fake_reg = _FakeWinreg()
        with mock.patch.object(cdr.sys, "platform", "win32"), mock.patch.object(cdr, "winreg", fake_reg):
            self.assertFalse(cdr.is_windows_autostart_enabled())
            ok, err = cdr.set_windows_autostart(True)
            self.assertTrue(ok, err)
            self.assertTrue(cdr.is_windows_autostart_enabled())
            cmd = fake_reg.store[cdr.AUTOSTART_REG_PATH][cdr.AUTOSTART_VALUE_NAME]
            self.assertIn("--minimized", cmd)

            ok2, err2 = cdr.set_windows_autostart(False)
            self.assertTrue(ok2, err2)
            self.assertFalse(cdr.is_windows_autostart_enabled())

    def test_02_template_editor_create_modify_and_build_docx(self):
        tpl = create_blank_template("test_editor_tpl", "Мой редакторский шаблон")
        tpl["_file_path"] = "/tmp/should_be_stripped.json"
        tpl["page"]["columns"] = 2
        tpl["strip_body_emphasis"] = True
        tpl["styles"]["Основной текст"]["size_pt"] = 12.0
        clean = strip_internal_keys(tpl)
        self.assertNotIn("_file_path", clean)

        with tempfile.TemporaryDirectory() as td:
            tpl_path = Path(td) / "test_editor_tpl.json"
            tpl_path.write_text(json.dumps(clean, ensure_ascii=False, indent=2), encoding="utf-8")
            loaded = cdr.load_template(str(tpl_path))
            self.assertEqual(loaded["styles"]["Основной текст"]["size_pt"], 12.0)

            out_docx = Path(td) / "out.docx"
            ast = MarkdownASTParser(loaded.get("structural_titles", [])).parse(
                "# 1 Введение\n\nТекст с **жирным** выделением.\n"
            )
            GostDocxFormatter(loaded, strip_body_emphasis=True).build_document(ast, str(out_docx))
            with zipfile.ZipFile(out_docx) as zf:
                doc_root = ET.fromstring(zf.read("word/document.xml").decode("utf-8"))
            cols = [c.get(f"{{{W_NS}}}num") for c in doc_root.findall(".//w:cols", NS)]
            self.assertIn("2", cols)

    def test_03_merge_with_two_columns_and_strip_body_emphasis(self):
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            d1_p = tmp / "d1.docx"
            d2_p = tmp / "d2.docx"
            merged_p = tmp / "merged.docx"

            d1 = Document()
            p1 = d1.add_paragraph()
            r1 = p1.add_run("Жирный текст из первого файла")
            r1.bold = True
            d1.save(d1_p)

            d2 = Document()
            p2 = d2.add_paragraph()
            r2 = p2.add_run("Курсив из второго файла")
            r2.italic = True
            d2.save(d2_p)

            merge_docx_files(str(d1_p), [str(d2_p)], output_path=str(merged_p), add_page_break=True)
            update_docx_styles(
                str(merged_p),
                {"page": {"columns": 2, "column_spacing_mm": 5.0}, "strip_body_emphasis": True},
                output_path=str(merged_p),
            )

            res = Document(merged_p)
            for p in res.paragraphs:
                for r in p.runs:
                    self.assertFalse(bool(r.bold))
                    self.assertFalse(bool(r.italic))

            with zipfile.ZipFile(merged_p) as zf:
                doc_root = ET.fromstring(zf.read("word/document.xml").decode("utf-8"))
            cols = [c.get(f"{{{W_NS}}}num") for c in doc_root.findall(".//w:cols", NS)]
            self.assertIn("2", cols)

    def test_04_system_font_grouping_by_type(self):
        from font_picker import classify_font_family, get_system_fonts_grouped
        self.assertEqual(classify_font_family("Times New Roman", False), "serif")
        self.assertEqual(classify_font_family("Cambria", False), "serif")
        self.assertEqual(classify_font_family("Segoe UI", False), "sans")
        self.assertEqual(classify_font_family("Consolas", True), "mono")
        self.assertEqual(classify_font_family("Segoe Script", False), "display")
        grouped = get_system_fonts_grouped()
        self.assertGreater(len(grouped), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
