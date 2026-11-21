#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py — Автономный кроссплатформенный сборщик ClipboardDocxRunner в EXE (PyInstaller).
Устойчив к русским путям с пробелами, кодировкам консоли CP1251/CP866 и блокировкам файлов в Windows.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# Защита от краха консоли Windows при выводе символов в кодировках CP1251 / CP866
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def main():
    print("=" * 60)
    print("   СБОРКА ClipboardDocxRunner В АВТОНОМНЫЙ EXE (PyInstaller)")
    print("=" * 60)
    print()

    base_dir = Path(__file__).resolve().parent
    os.chdir(base_dir)

    # 1. Завершение работающих процессов ClipboardDocxRunner.exe (снятие файловых блокировок)
    if sys.platform == "win32":
        try:
            subprocess.run(["taskkill", "/F", "/IM", "ClipboardDocxRunner.exe"], capture_output=True)
        except Exception:
            pass

    # 2. Проверка и установка зависимостей
    print("[1/5] Проверка зависимостей (PyInstaller, PySide6, python-docx, lxml, pyperclip, keyboard)...")
    req_packages = ["pyinstaller", "python-docx", "lxml", "pyperclip", "keyboard", "PySide6"]
    
    missing_pkgs = []
    for mod_name, pkg_name in [("PyInstaller", "pyinstaller"), ("docx", "python-docx"), ("lxml", "lxml"), ("pyperclip", "pyperclip"), ("keyboard", "keyboard"), ("PySide6", "PySide6")]:
        try:
            __import__(mod_name)
        except ImportError:
            missing_pkgs.append(pkg_name)
    if missing_pkgs:
        print(f"  [*] Установка недостающих пакетов: {', '.join(missing_pkgs)}...")
        cmd_install = [sys.executable, "-m", "pip", "install"] + missing_pkgs
        try:
            subprocess.check_call(cmd_install)
            print("  [OK] Зависимости успешно установлены.")
        except Exception as e:
            print(f"  [!] Ошибка при установке через pip: {e}")
            print("  Продолжаем сборку с уже установленными в системе пакетами...")

    # 3. Очистка старых сборок
    print("[2/5] Очистка директорий предыдущей сборки (build и dist)...")
    for d in ["build", "dist"]:
        dp = base_dir / d
        if dp.exists():
            try:
                shutil.rmtree(dp)
            except Exception as e:
                print(f"  [!] Не удалось удалить {d} (возможно, занято другим процессом): {e}")

    # 4. Поиск целевого скрипта
    target_script = "clipboard_docx_runner.py"
    if not (base_dir / target_script).exists():
        if (base_dir / "app" / "clipboard_docx_runner.py").exists():
            target_script = str(Path("app") / "clipboard_docx_runner.py")
        else:
            print(f"  [ОШИБКА] Не найден файл {target_script} в {base_dir}!")
            sys.exit(1)

    print(f"  [OK] Целевой файл: {target_script}")

    # 5. Проверка файла иконки
    icon_args = []
    for icon_name in ["app_docx.ico", "app/app_docx.ico"]:
        if (base_dir / icon_name).exists():
            icon_args = [f"--icon={icon_name}"]
            break

    # 6. Формирование параметров PyInstaller
    print("[3/5] Запуск PyInstaller...")

    # Разделитель путей для --add-data в Windows — точка с запятой (;)
    sep = ";" if sys.platform == "win32" else ":"

    # Используем относительные пути для --add-data и --paths:
    # это на 100% защищает от сбоев с двоеточием диска (C:), пробелами и кириллицей в путях Windows
    pyinstaller_args = [
        "--noconsole",
        "--onefile",
        "--clean",
        "--name=ClipboardDocxRunner",
        "--paths=gost_engine",
        "--hidden-import=docx",
        "--hidden-import=docx.oxml",
        "--hidden-import=docx.oxml.ns",
        "--hidden-import=docx.enum.text",
        "--hidden-import=docx.enum.table",
        "--hidden-import=docx.enum.style",
        "--hidden-import=lxml",
        "--hidden-import=lxml.etree",
        "--hidden-import=PySide6",
        "--hidden-import=PySide6.QtCore",
        "--hidden-import=PySide6.QtGui",
        "--hidden-import=PySide6.QtWidgets",
        "--hidden-import=pyperclip",
        "--hidden-import=keyboard",
        "--hidden-import=gost_docx_parser",
        "--hidden-import=gost_docx_formatter",
    ] + icon_args

    if (base_dir / "templates").exists():
        pyinstaller_args.append(f"--add-data=templates{sep}templates")
    if (base_dir / "gost_engine").exists():
        pyinstaller_args.append(f"--add-data=gost_engine{sep}gost_engine")

    pyinstaller_args.append(target_script)

    print("  Выполняется сборка EXE (это займет 20-40 секунд)...")
    
    # Запуск через прямой программный интерфейс PyInstaller (или через subprocess в случае сбоя)
    success = False
    try:
        import PyInstaller.__main__
        PyInstaller.__main__.run(pyinstaller_args)
        success = True
    except Exception as e:
        print(f"  [!] Запуск через PyInstaller API вызвал исключение: {e}")
        print("  Повторная попытка через python -m PyInstaller...")
        cmd_run = [sys.executable, "-m", "PyInstaller"] + pyinstaller_args
        res = subprocess.run(cmd_run)
        if res.returncode == 0:
            success = True

    if not success:
        print()
        print("=" * 60)
        print("  [ОШИБКА] Сборка завершилась с ошибкой!")
        print("  Убедитесь, что закрыты все экземпляры Word и ClipboardDocxRunner.")
        print("=" * 60)
        sys.exit(1)

    # 7. Финализация: копирование шаблонов и образцов документов рядом с EXE
    print("[4/5] Финализация: копирование шаблонов и образцов документов...")
    dist_templates = base_dir / "dist" / "templates"
    dist_templates.mkdir(parents=True, exist_ok=True)
    templates_dir = base_dir / "templates"
    if templates_dir.exists():
        copied_tpls = set()
        known_tpls = [
            "dstu_2026_gost_russian_names.json",
            "dstu_2026_gost_v3.json",
            "vak_rinc_article.json",
            "gost_34_tz.json",
            "corporate.json",
            "default.json",
            "ieee_vak_two_column.json",
            ".templates_index.json"
        ]
        for kn in known_tpls:
            kp = templates_dir / kn
            if kp.exists():
                shutil.copy2(kp, dist_templates / kp.name)
                copied_tpls.add(kp.name)
                print(f"  [OK] Скопирован шаблон: {kp.name}")
        for tpl in templates_dir.glob("*.json"):
            if tpl.name not in copied_tpls:
                shutil.copy2(tpl, dist_templates / tpl.name)
                print(f"  [OK] Скопирован шаблон: {tpl.name}")

    dist_samples = base_dir / "dist" / "samples"
    dist_samples.mkdir(parents=True, exist_ok=True)
    samples_dir = base_dir / "samples"
    if samples_dir.exists():
        for smp in samples_dir.glob("*.md"):
            shutil.copy2(smp, dist_samples / smp.name)
            print(f"  [OK] Скопирован образец: {smp.name}")

    exe_file = base_dir / "dist" / "ClipboardDocxRunner.exe"
    print()
    print("=" * 60)
    print("   СБОРКА УСПЕШНО ЗАВЕРШЕНА!")
    if exe_file.exists():
        print(f"   Файл: dist\\ClipboardDocxRunner.exe ({exe_file.stat().st_size // (1024*1024)} МБ)")
    else:
        print(f"   Файл собран в папке dist\\")
    print(f"   Шаблоны: dist\\templates\\")
    print(f"   Образцы: dist\\samples\\")
    print("=" * 60)
    print()

    # Открытие папки dist в проводнике Windows
    if sys.platform == "win32":
        try:
            os.startfile(str(base_dir / "dist"))
        except Exception:
            pass

if __name__ == "__main__":
    main()