#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Комплексные тесты объединения разных стилей, append и использования исходных стилей DOCX:
1. append_markdown_to_docx с preserve_existing_styles=True: сохраняет и использует стили исходного DOCX
   (встроенные Normal/Heading1/Heading2/Quote и кастомный русский стиль 'Примечание'), без коллизий списков.
2. append_markdown_to_docx с update_existing_styles=True: переопределяет стили исходного DOCX новым шаблоном.
3. merge_docx_files: объединяет документы с разными стилями, переносит уникальные стили из второго DOCX,
   переназначает конфликтующие по имени styleId (без битых pStyle) и разводит numId в numbering.xml.
"""

import json
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Mm, RGBColor
from PIL import Image, ImageDraw

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "gost_engine"))

from gost_docx_formatter import append_markdown_to_docx, merge_docx_files

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def load_template(name: str) -> dict:
    return json.loads((BASE_DIR / "templates" / f"{name}.json").read_text(encoding="utf-8"))


def make_png(path: Path, color: str):
    img = Image.new("RGB", (180, 100), color)
    ImageDraw.Draw(img).rectangle([8, 8, 172, 92], outline="black", width=2)
    img.save(path)


class TestStylesAppendAndMerge(unittest.TestCase):
    def test_01_append_uses_original_docx_builtin_and_custom_russian_styles(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            base_docx = tmp / "original.docx"
            out_docx = tmp / "appended.docx"

            doc = Document()
            normal = doc.styles["Normal"]
            normal.font.name = "Georgia"
            normal.font.size = Pt(11)
            normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)

            h1 = doc.styles["Heading 1"]
            h1.font.name = "Georgia"
            h1.font.size = Pt(16)
            h1.font.bold = True
            h1.font.color.rgb = RGBColor(0x11, 0x44, 0x88)

            h2 = doc.styles["Heading 2"]
            h2.font.name = "Georgia"
            h2.font.size = Pt(13)
            h2.font.bold = True

            quote = doc.styles["Quote"]
            quote.font.name = "Georgia"
            quote.font.size = Pt(11)
            quote.font.italic = True
            quote.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

            note_st = doc.styles.add_style("Примечание", WD_STYLE_TYPE.PARAGRAPH)
            note_st.font.name = "Georgia"
            note_st.font.size = Pt(10)
            note_st.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

            doc.add_paragraph("Исходный заголовок", style=h1)
            doc.add_paragraph("Исходный абзац", style=normal)
            doc.add_paragraph("Исходный пункт списка", style="List Bullet")
            doc.save(base_docx)

            cfg = load_template("dstu_2026_gost_russian_names")
            md = (
                "# 2 Новый раздел\n\n"
                "## 2.1 Подраздел\n\n"
                "Добавленный текст в стиле исходного документа.\n\n"
                "> Добавленная цитата.\n\n"
                "Примечание - Добавленное примечание.\n\n"
                "- первый добавленный пункт;\n"
                "- второй добавленный пункт.\n"
            )
            append_markdown_to_docx(
                str(base_docx),
                md,
                cfg,
                output_path=str(out_docx),
                update_existing_styles=False,
            )

            res = Document(out_docx)
            # Исходные стили НЕ перезаписаны шаблоном (остались Georgia)
            self.assertEqual(res.styles["Normal"].font.name, "Georgia")
            self.assertEqual(res.styles["Normal"].font.size.pt, 11.0)
            self.assertEqual(res.styles["Heading 1"].font.name, "Georgia")
            self.assertEqual(res.styles["Heading 1"].font.size.pt, 16.0)
            self.assertEqual(res.styles["Heading 1"].font.color.rgb, RGBColor(0x11, 0x44, 0x88))
            self.assertEqual(res.styles["Примечание"].font.name, "Georgia")
            self.assertEqual(res.styles["Примечание"].font.size.pt, 10.0)

            # Добавленные параграфы используют именно стили исходного документа
            by_text = {p.text.replace("\u00a0", " "): p for p in res.paragraphs}
            self.assertEqual(by_text["2 Новый раздел"].style.name, "Heading 1")
            self.assertEqual(by_text["2.1 Подраздел"].style.name, "Heading 2")
            self.assertEqual(by_text["Добавленный текст в стиле исходного документа."].style.name, "Normal")
            self.assertEqual(by_text["Добавленная цитата."].style.name, "Quote")
            self.assertEqual(by_text["Примечание — Добавленное примечание."].style.name, "Примечание")

            # Списки исходного и добавленного текста имеют валидные numId в numbering.xml
            with zipfile.ZipFile(out_docx) as zf:
                num_root = ET.fromstring(zf.read("word/numbering.xml").decode("utf-8"))
                doc_root = ET.fromstring(zf.read("word/document.xml").decode("utf-8"))
            defined_nids = {el.get(f"{{{W_NS}}}numId") for el in num_root.findall("w:num", NS)}
            used_nids = [el.get(f"{{{W_NS}}}val") for el in doc_root.findall(".//w:numPr/w:numId", NS)]
            self.assertGreaterEqual(len(used_nids), 2)
            for nid in used_nids:
                self.assertIn(nid, defined_nids)

    def test_02_append_with_update_existing_styles_overrides_original(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            base_docx = tmp / "original.docx"
            out_docx = tmp / "updated.docx"

            doc = Document()
            doc.styles["Normal"].font.name = "Calibri"
            doc.styles["Normal"].font.size = Pt(11)
            doc.add_paragraph("Старый текст.")
            doc.save(base_docx)

            cfg = load_template("dstu_2026_gost_russian_names")
            append_markdown_to_docx(
                str(base_docx),
                "Новый текст после обновления стилей.\n",
                cfg,
                output_path=str(out_docx),
                update_existing_styles=True,
            )

            res = Document(out_docx)
            self.assertEqual(res.styles["Normal"].font.name, "Times New Roman")
            self.assertEqual(res.styles["Normal"].font.size.pt, 14.0)
            self.assertEqual(res.paragraphs[0].text, "Старый текст.")
            self.assertEqual(res.paragraphs[1].text, "Новый текст после обновления стилей.")
            self.assertEqual(res.paragraphs[1].style.name, "Normal")

    def test_03_merge_different_styles_remaps_name_collisions_and_lists(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            d1_path = tmp / "doc1.docx"
            d2_path = tmp / "doc2.docx"
            merged_path = tmp / "merged.docx"
            img2 = tmp / "pic2.png"
            make_png(img2, "teal")

            # Doc 1: Georgia, BoxDoc1, список 1
            d1 = Document()
            d1.styles["Normal"].font.name = "Georgia"
            d1.styles["Normal"].font.size = Pt(11)
            d1.styles["Heading 1"].font.name = "Georgia"
            d1.styles["Heading 1"].font.color.rgb = RGBColor(0x11, 0x44, 0x88)
            box1 = d1.styles.add_style("BoxDoc1", WD_STYLE_TYPE.PARAGRAPH)
            box1.font.name = "Verdana"
            box1.font.size = Pt(10)

            d1.add_paragraph("Заголовок из Doc1", style="Heading 1")
            d1.add_paragraph("Блок из Doc1", style=box1)
            from docx.oxml import parse_xml
            from docx.oxml.ns import nsdecls
            p_l1 = d1.add_paragraph("Пункт Doc1", style="List Bullet")
            d1_num = d1.part.numbering_part.numbering_definitions._numbering
            d1_num.append(parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>'))
            d1_num.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="1"><w:abstractNumId w:val="1"/></w:num>'))
            p_l1._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))
            d1.save(d1_path)

            # Doc 2: уникальный стиль CodeDoc2, конфликтующий по имени AltHeading1 ('heading 1'), таблица, картинка, свой список
            d2 = Document()
            code2 = d2.styles.add_style("CodeDoc2", WD_STYLE_TYPE.PARAGRAPH)
            code2.font.name = "Consolas"
            code2.font.size = Pt(10)
            code2.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

            # Создаем кастомный стиль с другим styleId, но именем как у встроенного Heading 1
            alt_h1 = d2.styles.add_style("AltHeading1Temp", WD_STYLE_TYPE.PARAGRAPH)
            alt_h1.element.find(f"{{{W_NS}}}name").set(f"{{{W_NS}}}val", "heading 1")

            d2.add_paragraph("Заголовок из Doc2", style=alt_h1)
            d2.add_paragraph("Код из Doc2", style=code2)
            p_l2 = d2.add_paragraph("Пункт Doc2", style="List Number")
            d2_num = d2.part.numbering_part.numbering_definitions._numbering
            d2_num.append(parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'))
            d2_num.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="1"><w:abstractNumId w:val="1"/></w:num>'))
            p_l2._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))
            tbl = d2.add_table(rows=2, cols=2)
            tbl.cell(0, 0).text = "Ключ"
            tbl.cell(1, 1).text = "Значение"
            d2.add_picture(str(img2), width=Mm(35))
            d2.save(d2_path)

            merge_docx_files(str(d1_path), [str(d2_path)], output_path=str(merged_path), add_page_break=True)

            res = Document(merged_path)
            # 1. Стили Doc1 сохранены
            self.assertEqual(res.styles["Normal"].font.name, "Georgia")
            self.assertEqual(res.styles["Heading 1"].font.color.rgb, RGBColor(0x11, 0x44, 0x88))
            self.assertIn("BoxDoc1", [s.name for s in res.styles])

            # 2. Уникальный стиль CodeDoc2 из Doc2 перенесён и применён
            self.assertIn("CodeDoc2", [s.name for s in res.styles])
            self.assertEqual(res.styles["CodeDoc2"].font.name, "Consolas")

            by_text = {p.text: p for p in res.paragraphs if p.text}
            self.assertEqual(by_text["Блок из Doc1"].style.name, "BoxDoc1")
            self.assertEqual(by_text["Код из Doc2"].style.name, "CodeDoc2")
            # 3. Конфликтующий по имени AltHeading1 переназначен на Heading 1 (без битого styleId)
            self.assertEqual(by_text["Заголовок из Doc2"].style.name, "Heading 1")

            # 4. Проверяем целостность pStyle и numbering.xml в XML
            with zipfile.ZipFile(merged_path) as zf:
                styles_root = ET.fromstring(zf.read("word/styles.xml").decode("utf-8"))
                doc_root = ET.fromstring(zf.read("word/document.xml").decode("utf-8"))
                num_root = ET.fromstring(zf.read("word/numbering.xml").decode("utf-8"))
                media = [n for n in zf.namelist() if n.startswith("word/media/")]

            all_style_ids = {s.get(f"{{{W_NS}}}styleId") for s in styles_root.findall("w:style", NS)}
            for pstyle in doc_root.findall(".//w:pPr/w:pStyle", NS):
                val = pstyle.get(f"{{{W_NS}}}val")
                self.assertIn(val, all_style_ids, f"Битая ссылка на стиль pStyle={val}")

            defined_nids = {el.get(f"{{{W_NS}}}numId") for el in num_root.findall("w:num", NS)}
            used_nids = [el.get(f"{{{W_NS}}}val") for el in doc_root.findall(".//w:numPr/w:numId", NS)]
            self.assertEqual(len(used_nids), 2)
            self.assertEqual(len(set(used_nids)), 2, "Списки из двух документов должны иметь разные numId")
            for nid in used_nids:
                self.assertIn(nid, defined_nids)

            self.assertEqual(len(res.tables), 1)
            self.assertEqual(len(media), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
