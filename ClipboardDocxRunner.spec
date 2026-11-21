# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['clipboard_docx_runner.py'],
    pathex=['gost_engine'],
    binaries=[],
    datas=[('templates', 'templates'), ('gost_engine', 'gost_engine')],
    hiddenimports=['docx', 'docx.oxml', 'docx.oxml.ns', 'docx.enum.text', 'docx.enum.table', 'docx.enum.style', 'lxml', 'lxml.etree', 'PySide6', 'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets', 'pyperclip', 'keyboard', 'gost_docx_parser', 'gost_docx_formatter', 'template_editor_widget', 'font_picker'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ClipboardDocxRunner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_docx.ico'],
)
