#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты подтягивания изображений из Markdown по относительному, абсолютному,
URL-кодированному пути и file:// URI, а также через CLI clipboard_docx_runner.py.
"""

import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from urllib.parse import quote

from PIL import Image, ImageDraw

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR / "gost_engine"))

from gost_docx_parser import MarkdownASTParser
from gost_docx_formatter import GostDocxFormatter


def make_png(path: Path, color: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (240, 120), color)
    ImageDraw.Draw(img).rectangle([10, 10, 230, 110], outline="black", width=3)
    img.save(path)


class TestImagePaths(unittest.TestCase):
    def test_relative_absolute_urlencoded_and_file_uri(self):
        cfg = json.loads((BASE_DIR / "templates" / "dstu_2026_gost_v3.json").read_text(encoding="utf-8"))

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            md_dir = tmp / "docs"
            md_dir.mkdir()

            rel_img = md_dir / "media" / "rel_fig.png"
            parent_img = tmp / "shared" / "parent_fig.png"
            abs_img = tmp / "abs_dir" / "abs_fig.png"
            cyr_img = md_dir / "media" / "лабораторный стенд 2026.png"
            ang_img = md_dir / "media" / "fig with spaces.png"
            uri_img = tmp / "uri_dir" / "uri_fig.png"

            make_png(rel_img, "red")
            make_png(parent_img, "green")
            make_png(abs_img, "blue")
            make_png(cyr_img, "yellow")
            make_png(ang_img, "cyan")
            make_png(uri_img, "magenta")

            encoded_cyr = "media/" + quote("лабораторный стенд 2026.png")
            md_text = (
                "# 1 Иллюстрации\n\n"
                "![Относительный](media/rel_fig.png)\n\n"
                "![Родительский каталог](../shared/parent_fig.png)\n\n"
                f"![Абсолютный]({abs_img.resolve()})\n\n"
                f"![URL-кодированный кириллический]({encoded_cyr})\n\n"
                "![В угловых скобках с заголовком](<media/fig with spaces.png> \"Подпись\"){width=100mm}\n\n"
                f"![File URI]({uri_img.resolve().as_uri()})\n\n"
                "![Отсутствующий файл](media/not_found.png)\n"
            )

            md_file = md_dir / "paper.md"
            md_file.write_text(md_text, encoding="utf-8")
            out_docx = tmp / "out.docx"

            # Проверяем через CLI clipboard_docx_runner.py (как в реальном использовании)
            subprocess.run(
                [sys.executable, str(BASE_DIR / "clipboard_docx_runner.py"), str(md_file), str(out_docx), "dstu_2026_gost_v3"],
                capture_output=True, text=True, check=True,
            )

            with zipfile.ZipFile(out_docx) as z:
                media = [n for n in z.namelist() if n.startswith("word/media/")]
                doc_xml = z.read("word/document.xml").decode("utf-8")

            self.assertEqual(len(media), 6)
            self.assertEqual(doc_xml.count("<w:drawing>"), 6)
            self.assertIn("[Отсутствующий файл]", doc_xml)


if __name__ == "__main__":
    unittest.main(verbosity=2)
