#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Стресс-тесты реалистичных Markdown-документов, сгенерированных нейросетью по навыку gost-md-bundle:
1. Научная статья ВАК / ДГТУ (структурные заголовки, формулы с \tag{1}, где-пояснения, таблицы с единицами, списки, сноски, изображения в media/ -> ZIP)
2. Двухколоночная статья IEEE (columns: 2, mermaid/dot диаграммы, широкая таблица, открытая сетка gost_open)
3. Техническое задание ГОСТ 34 (многоуровневые списки 1) -> а) -> -, листинги кода, примечания, callouts, без картинок -> только .md)
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

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}


def make_chart(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (600, 360), "white")
    d = ImageDraw.Draw(img)
    d.rectangle([20, 20, 580, 340], outline="#2563EB", width=4)
    d.line([60, 280, 200, 180, 360, 210, 520, 80], fill="#0D9488", width=5)
    img.save(path, dpi=(300, 300))


class TestLLMSkillPrompts(unittest.TestCase):
    def _validate_docx_openxml(self, docx_path: Path):
        with zipfile.ZipFile(docx_path) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
            styles_xml = z.read("word/styles.xml").decode("utf-8")
        doc_root = ET.fromstring(doc_xml)
        styles_root = ET.fromstring(styles_xml)
        self.assertIsNotNone(doc_root)
        self.assertIsNotNone(styles_root)
        return doc_root, styles_root

    def test_01_llm_vak_article_with_media_zip_bundle(self):
        md = (
            "# Аннотация\n\n"
            "В работе рассмотрен алгоритм векторного управления приводом А. М. Смирнова на частоте 20 кГц при температуре 20-25 °C.\n\n"
            "# 1 Постановка задачи\n\n"
            "Электромагнитный момент двигателя определяется выражением:\n\n"
            "$$M = \\frac{3}{2} p \\left(\\psi_d i_q - \\psi_q i_d\\right) \\tag{1}$$\n\n"
            "где $p$ - число пар полюсов;\n"
            "$i_d, i_q$ - проекции тока статора.\n\n"
            "Таблица 1 - Параметры экспериментального стенда (В миллиметрах)\n\n"
            "| Параметр | Обозначение | Значение |\n"
            "| :--- | :---: | ---: |\n"
            "| Длина плеча | $L_1$ | 450 |\n"
            "| Диаметр вала | $d$ | 24 |\n\n"
            "![График переходного процесса](media/step_response.png){width=130mm}\n\n"
            "1 - заданное значение; 2 - измеренный отклик\n\n"
            "Рисунок 1 - Переходный процесс контура скорости\n\n"
            "Примечание - Измерения выполнены с точностью 0,5 % (рис. 1, табл. 1)[^1].\n\n"
            "[^1]: Данные получены на стенде кафедры робототехники.\n\n"
            "# Список использованных источников\n\n"
            "1. Иванов И. И. Системы управления приводами. М.: Наука, 2024. 320 с.\n"
        )
        cfg = json.loads((BASE_DIR / "templates" / "vak_rinc_article.json").read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td)
            md_path = tmp / "article.md"
            img_path = tmp / "media" / "step_response.png"
            make_chart(img_path)
            md_path.write_text(md, encoding="utf-8")

            # Имитируем упаковку по навыку gost-md-bundle
            zip_path = tmp / "article.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(md_path, "article.md")
                zf.write(img_path, "media/step_response.png")

            # Распаковка и сборка скриптом
            unpack_dir = tmp / "unpacked"
            with zipfile.ZipFile(zip_path) as zf:
                zf.extractall(unpack_dir)

            ast_nodes = MarkdownASTParser(cfg.get("structural_titles", [])).parse((unpack_dir / "article.md").read_text(encoding="utf-8"))
            out_docx = tmp / "article.docx"
            GostDocxFormatter(cfg).build_document(ast_nodes, str(out_docx), base_dir=str(unpack_dir))

            doc_root, _ = self._validate_docx_openxml(out_docx)
            with zipfile.ZipFile(out_docx) as z:
                media = [n for n in z.namelist() if n.startswith("word/media/")]
            self.assertEqual(len(media), 1)
            self.assertEqual(len(doc_root.findall(".//m:oMath", NS)), 5)
            self.assertEqual(len(doc_root.findall(".//w:tbl", NS)), 1)

    def test_02_llm_ieee_two_column_article(self):
        md = (
            "# 1 Introduction\n\n"
            "This paper presents a two-column conference layout with automatic single-column wide tables.\n\n"
            "## 1.1 System Model\n\n"
            "$$y(t) = \\int_0^t u(\\tau)\\,d\\tau \\qquad (1)$$\n\n"
            "Таблица 1 - Comparative benchmark across controllers\n\n"
            "| Method | RMSE | Settling Time | Overshoot |\n"
            "| :--- | :---: | :---: | :---: |\n"
            "| PID | 0.142 | 1.8 s | 12 % |\n"
            "| MPC | 0.041 | 0.9 s | 2 % |\n\n"
            "```mermaid\n"
            "graph LR\n"
            "  Sensor --> Controller --> Actuator\n"
            "```\n\n"
            "Рисунок 1 - Control loop architecture\n"
        )
        cfg = json.loads((BASE_DIR / "templates" / "ieee_vak_two_column.json").read_text(encoding="utf-8"))
        ast_nodes = MarkdownASTParser(cfg.get("structural_titles", [])).parse(md)
        with tempfile.TemporaryDirectory() as td:
            out_docx = Path(td) / "ieee.docx"
            GostDocxFormatter(cfg).build_document(ast_nodes, str(out_docx))
            doc_root, _ = self._validate_docx_openxml(out_docx)
            col_nums = [c.get(f"{{{NS['w']}}}num") for c in doc_root.findall(".//w:cols", NS)]
            self.assertIn("1", col_nums)
            self.assertIn("2", col_nums)

    def test_03_llm_gost34_tz_nested_lists_and_code_no_images(self):
        md = (
            "# Введение\n\n"
            "Настоящее техническое задание распространяется на автоматизированную систему.\n\n"
            "# 1 Требования к системе\n\n"
            "Система должна обеспечивать выполнение следующих функций:\n"
            "1) сбор телеметрии:\n"
            "   а) опрос датчиков по шине CAN;\n"
            "   б) буферизация кадров:\n"
            "      - частота не менее 100 Гц;\n"
            "      - глубина очереди 1024 кадра;\n"
            "2) диагностика неисправностей.\n\n"
            "> [!IMPORTANT] Все интерфейсы должны быть гальванически развязаны.\n\n"
            "Листинг 1 - Инициализация фильтра CAN\n"
            "```c\n"
            "void can_init(uint32_t baud) {\n"
            "    CAN1->BTR = baud;\n"
            "}\n"
            "```\n"
        )
        cfg = json.loads((BASE_DIR / "templates" / "gost_34_tz.json").read_text(encoding="utf-8"))
        ast_nodes = MarkdownASTParser(cfg.get("structural_titles", [])).parse(md)
        with tempfile.TemporaryDirectory() as td:
            out_docx = Path(td) / "tz.docx"
            GostDocxFormatter(cfg).build_document(ast_nodes, str(out_docx))
            doc_root, _ = self._validate_docx_openxml(out_docx)
            with zipfile.ZipFile(out_docx) as z:
                num_xml = z.read("word/numbering.xml").decode("utf-8")
            num_root = ET.fromstring(num_xml)
            self.assertGreaterEqual(len(doc_root.findall(".//w:numPr", NS)), 6)
            self.assertGreaterEqual(len(num_root.findall(".//w:abstractNum", NS)), 3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
