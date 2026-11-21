#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — Синоним главной точки входа для запуска Markdown to DOCX Suite.
Перенаправляет запуск на clipboard_docx_runner.main().
"""

import sys
from clipboard_docx_runner import main

if __name__ == "__main__":
    main()
