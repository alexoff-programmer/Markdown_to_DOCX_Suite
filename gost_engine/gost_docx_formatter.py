#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gost_docx_formatter.py — Движок генерации и оформления документов DOCX по стандартам ГОСТ / ДГТУ / IEEE.
Отвечает за:
- Полноценную компиляцию LaTeX -> Office Math OpenXML (<m:oMath>, <m:oMathPara>)
- Настройку геометрии страниц, многоколоночной верстки (columns) и первой особой страницы (different_first_page)
- Регистрацию шрифтов Times New Roman и Courier New во всех XML-диапазонах
- Декларативную верстку таблиц (gost_open / full_grid, cell_padding_mm, header_shading, tblHeader, cantSplit)
- Автоматический расчет ширин столбцов таблиц (min unbreakable + 75-й перцентиль)
- Рендеринг нумерованных формул через две позиции табуляции (центр + правый край)
- Компиляцию диаграмм mermaid / plantuml / dot внешними CLI-рендерерами в стиль Рисунок
- Оформление списков с сохранением исходного маркера, цитат, выносок и связывание соседних блоков
"""

from __future__ import annotations

import copy
from functools import lru_cache
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Dict, Tuple, Any, Optional
from xml.sax.saxutils import escape

import docx
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

from gost_docx_parser import MarkdownASTParser, sanitize_and_repair_latex_escapes, tokenize_inlines, typography_cleanup, NBSP

DXA_PER_MM = 56.692913

GREEK_AND_SYMBOLS = {
    # Строчные греческие буквы
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ϵ',
    'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'vartheta': 'ϑ',
    'iota': 'ι', 'kappa': 'κ', 'varkappa': 'ϰ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν',
    'xi': 'ξ', 'pi': 'π', 'varpi': 'ϖ', 'rho': 'ρ', 'varrho': 'ϱ',
    'sigma': 'σ', 'varsigma': 'ς', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'ϕ',
    'varphi': 'φ', 'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',

    # Прописные греческие буквы
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ',
    'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ', 'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',

    # Операторы и отношения
    'cdot': '·', 'times': '×', 'div': '÷', 'pm': '±', 'mp': '∓',
    'ast': '∗', 'star': '⋆', 'circ': '°', 'bullet': '∙',
    'oplus': '⊕', 'ominus': '⊖', 'otimes': '⊗', 'odot': '⊙',
    'approx': '≈', 'sim': '∼', 'simeq': '≃', 'cong': '≅', 'equiv': '≡', 'propto': '∝',
    'neq': '≠', 'ne': '≠', 'leq': '≤', 'le': '≤', 'geq': '≥', 'ge': '≥',
    'll': '≪', 'gg': '≫', 'prec': '≺', 'succ': '≻', 'preceq': '⪯', 'succeq': '⪰',
    'infty': '∞', 'partial': '∂', 'nabla': '∇', 'hbar': 'ħ', 'ell': 'ℓ',
    'forall': '∀', 'exists': '∃', 'nexists': '∄', 'neg': '¬', 'lnot': '¬',
    'land': '∧', 'lor': '∨', 'wedge': '∧', 'vee': '∨',
    'in': '∈', 'notin': '∉', 'ni': '∋', 'subset': '⊂', 'supset': '⊃',
    'subseteq': '⊆', 'supseteq': '⊇', 'cup': '∪', 'cap': '∩', 'setminus': '∖',
    'emptyset': '∅', 'varnothing': '∅',
    'sum': '∑', 'prod': '∏', 'coprod': '∐', 'int': '∫', 'iint': '∬', 'iiint': '∭', 'oint': '∮',
    'longrightarrow': '⟶', 'longleftarrow': '⟵', 'longleftrightarrow': '⟷',
    'rightarrow': '→', 'to': '→', 'leftarrow': '←', 'gets': '←', 'leftrightarrow': '↔',
    'Longrightarrow': '⟹', 'Longleftarrow': '⟸', 'Longleftrightarrow': '⟺',
    'Rightarrow': '⇒', 'Leftarrow': '⇐', 'Leftrightarrow': '⇔',
    'implies': '⟹', 'iff': '⟺', 'mapsto': '↦', 'longmapsto': '⟼',
    'uparrow': '↑', 'downarrow': '↓', 'updownarrow': '↕',
    'nearrow': '↗', 'searrow': '↘', 'swarrow': '↙', 'nwarrow': '↖',
    'dots': '…', 'ldots': '…', 'cdots': '⋯', 'vdots': '⋮', 'ddots': '⋱',
    'quad': '\u2003', 'qquad': '\u2003\u2003', 'enspace': '\u2002', 'thinspace': '\u2009',
    'vert': '|', 'lvert': '|', 'rvert': '|', 'mid': '|',
    'Vert': '‖', 'lVert': '‖', 'rVert': '‖', 'parallel': '∥', 'perp': '⊥', 'angle': '∠',
    'langle': '⟨', 'rangle': '⟩', 'lceil': '⌈', 'rceil': '⌉', 'lfloor': '⌊', 'rfloor': '⌋',
    'degree': '°', 'prime': '′', 'triangle': '△', 'Box': '□', 'square': '□',
    'Re': 'ℜ', 'Im': 'ℑ', 'wp': '℘', 'aleph': 'ℵ', 'backslash': '\\',
}

UPRIGHT_FUNCTIONS = {
    'max', 'min', 'sup', 'inf', 'lim', 'limsup', 'liminf',
    'sin', 'cos', 'tan', 'tg', 'ctg', 'cot', 'sec', 'csc',
    'arcsin', 'arccos', 'arctan', 'arctg', 'arcctg',
    'sinh', 'cosh', 'tanh', 'sh', 'ch', 'th', 'cth',
    'ln', 'log', 'lg', 'exp', 'det', 'dim', 'ker', 'hom', 'arg', 'deg', 'gcd', 'Pr', 'mod',
    'sgn', 'sign', 'tr', 'diag', 'rank', 'const', 'var', 'cov', 'std'
}

ACCENT_COMMANDS = {
    'hat': '\u0302',
    'widehat': '\u0302',
    'tilde': '\u0303',
    'widetilde': '\u0303',
    'bar': '\u0304',
    'vec': '\u20D7',
    'overrightarrow': '\u20D7',
    'overleftarrow': '\u20D6',
    'dot': '\u0307',
    'ddot': '\u0308',
    'dddot': '\u20DB',
    'check': '\u030C',
    'breve': '\u0306',
    'acute': '\u0301',
    'grave': '\u0300',
    'mathring': '\u030A',
}

FONT_UPRIGHT_COMMANDS = {
    'mathrm', 'textrm', 'text', 'mbox', 'hbox', 'operatorname',
    'mathsf', 'mathtt', 'texttt', 'textnormal', 'rm'
}
FONT_BOLD_COMMANDS = {'mathbf', 'textbf', 'boldsymbol', 'bm', 'bf'}
FONT_ITALIC_COMMANDS = {'mathit', 'textit', 'it'}
FRAC_COMMANDS = {'frac', 'dfrac', 'tfrac', 'cfrac'}
NOOP_COMMANDS = {
    'displaystyle', 'textstyle', 'scriptstyle', 'scriptscriptstyle',
    'limits', 'nolimits', 'relax', 'allowbreak', 'noindent',
    'big', 'Big', 'bigg', 'Bigg', 'bigl', 'bigr', 'Bigl', 'Bigr',
    'biggl', 'biggr', 'Biggl', 'Biggr', 'bigm', 'Bigm', 'middle'
}

BINARY_RELATIONS = {
    '=', '≈', '∼', '≃', '≅', '≡', '∝', '≠', '≤', '≥', '≪', '≫',
    '≺', '≻', '⪯', '⪰', '∈', '∉', '∋', '⊂', '⊃', '⊆', '⊇',
    '→', '←', '↔', '⟶', '⟵', '⟷', '⇒', '⇐', '⇔', '⟹', '⟸', '⟺', '↦', '⟼',
    '<', '>'
}
BINARY_OPERATORS = {'+', '−', '±', '∓', '·', '×', '÷', '⊕', '⊖', '⊗', '∪', '∩', '∖', '∧', '∨'}
OPENING_TOKENS = {'(', '[', '{', '⟨', '⌈', '⌊', '|', '‖', ',', ';', ':'}


def _make_omml_run(text: str, style: str = 'math', size_pt: Optional[float] = None, bold: bool = False) -> str:
    if not text:
        return ""
    esc = escape(text)
    sz_xml = ""
    if size_pt:
        half_pts = int(round(size_pt * 2))
        sz_xml = f'<w:sz w:val="{half_pts}"/><w:szCs w:val="{half_pts}"/>'
    b_xml = '<w:b/><w:bCs/>' if (bold or style == 'bold') else ''

    if style == 'upright':
        return (
            '<m:r>'
            '<m:rPr><m:nor/></m:rPr>'
            f'<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>{b_xml}{sz_xml}</w:rPr>'
            f'<m:t xml:space="preserve">{esc}</m:t>'
            '</m:r>'
        )
    elif style == 'bold':
        return (
            '<m:r>'
            '<m:rPr><m:sty m:val="b"/></m:rPr>'
            f'<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math"/>{b_xml}{sz_xml}</w:rPr>'
            f'<m:t xml:space="preserve">{esc}</m:t>'
            '</m:r>'
        )
    else:
        m_sty = '<m:rPr><m:sty m:val="b"/></m:rPr>' if bold else ''
        return (
            '<m:r>'
            f'{m_sty}'
            f'<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math"/>{b_xml}{sz_xml}</w:rPr>'
            f'<m:t xml:space="preserve">{esc}</m:t>'
            '</m:r>'
        )


class LaTeXToOMMLCompiler:
    r"""
    Полноценный рекурсивный синтаксический транслятор LaTeX в нативный формат Word Office Math (OMML).
    Поддерживает произвольную глубину вложенности: дроби, корни, индексы, акценты (\widehat, \bar, \vec),
    скобки \left...\right, шрифтовые модификаторы (\mathrm, \text, \mathbf), матрицы и системы уравнений.
    """

    def __init__(self, latex: str, size_pt: Optional[float] = None, bold: bool = False):
        s = sanitize_and_repair_latex_escapes(latex or "").strip()
        if s.startswith('$$') and s.endswith('$$') and len(s) >= 4:
            s = s[2:-2].strip()
        elif s.startswith('$') and s.endswith('$') and len(s) >= 2:
            s = s[1:-1].strip()
        elif s.startswith(r'\[') and s.endswith(r'\]') and len(s) >= 4:
            s = s[2:-2].strip()
        elif s.startswith(r'\(') and s.endswith(r'\)') and len(s) >= 4:
            s = s[2:-2].strip()

        s = re.sub(r'\{\s*,\s*\}', ',', s)
        s = s.replace('...', '…')

        self.s = s
        self.n = len(s)
        self.pos = 0
        self.size_pt = size_pt
        self.bold = bold

    def _run(self, text: str, style: str = 'math') -> str:
        return _make_omml_run(text, style=style, size_pt=self.size_pt, bold=self.bold)

    def _ctrl_pr(self) -> str:
        sz_xml = ""
        if self.size_pt:
            half_pts = int(round(self.size_pt * 2))
            sz_xml = f'<w:sz w:val="{half_pts}"/><w:szCs w:val="{half_pts}"/>'
        return f'<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math"/>{sz_xml}</w:rPr></m:ctrlPr>'

    def compile(self) -> str:
        body = self.parse_sequence(stop_tokens=set(), style='math')
        if not body:
            body = self._run(" ")
        return (
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'{body}</m:oMath>'
        )

    def skip_spaces(self):
        while self.pos < self.n and self.s[self.pos] in (' ', '\t', '\r', '\n'):
            self.pos += 1

    def peek(self) -> str:
        return self.s[self.pos] if self.pos < self.n else ''

    def read_command(self) -> str:
        if self.pos >= self.n:
            return ''
        ch = self.s[self.pos]
        if ch.isalpha():
            start = self.pos
            while self.pos < self.n and self.s[self.pos].isalpha():
                self.pos += 1
            cmd = self.s[start:self.pos]
            if cmd in ('operatorname', 'tag') and self.pos < self.n and self.s[self.pos] == '*':
                self.pos += 1
            return cmd
        else:
            self.pos += 1
            return ch

    def read_balanced_group(self, open_ch: str = '{', close_ch: str = '}') -> str:
        self.skip_spaces()
        if self.pos >= self.n or self.s[self.pos] != open_ch:
            return ""
        self.pos += 1
        depth = 1
        out = []
        while self.pos < self.n and depth > 0:
            c = self.s[self.pos]
            if c == '\\' and self.pos + 1 < self.n:
                out.append(self.s[self.pos:self.pos + 2])
                self.pos += 2
                continue
            if c == open_ch:
                depth += 1
                out.append(c)
            elif c == close_ch:
                depth -= 1
                if depth > 0:
                    out.append(c)
            else:
                out.append(c)
            self.pos += 1
        return "".join(out)

    def parse_sub_expr(self, sub_latex: str, style: str = 'math') -> str:
        sub_compiler = LaTeXToOMMLCompiler(sub_latex, size_pt=self.size_pt, bold=self.bold)
        return sub_compiler.parse_sequence(stop_tokens=set(), style=style)

    def read_script_arg(self, style: str = 'math') -> str:
        self.skip_spaces()
        if self.pos >= self.n:
            return self._run("", style)
        if self.s[self.pos] == '{':
            grp = self.read_balanced_group('{', '}')
            return self.parse_sub_expr(grp, style=style)
        elif self.s[self.pos] == '\\':
            self.pos += 1
            cmd = self.read_command()
            if cmd in FONT_UPRIGHT_COMMANDS:
                grp = self.read_balanced_group('{', '}')
                return self.parse_sub_expr(grp, style='upright')
            elif cmd in FONT_BOLD_COMMANDS:
                grp = self.read_balanced_group('{', '}')
                return self.parse_sub_expr(grp, style='bold')
            elif cmd in GREEK_AND_SYMBOLS:
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return self._run(GREEK_AND_SYMBOLS[cmd], style)
            else:
                return self._run(cmd, style)
        else:
            ch = self.s[self.pos]
            self.pos += 1
            return self._run(ch, style)

    def read_delimiter_token(self) -> str:
        self.skip_spaces()
        if self.pos >= self.n:
            return ''
        ch = self.s[self.pos]
        if ch == '\\':
            self.pos += 1
            cmd = self.read_command()
            if cmd in ('{', '}'):
                return cmd
            if cmd in ('vert', 'lvert', 'rvert', 'mid', '|'):
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return '|'
            if cmd in ('Vert', 'lVert', 'rVert'):
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return '‖'
            if cmd in GREEK_AND_SYMBOLS:
                return GREEK_AND_SYMBOLS[cmd]
            return ''
        else:
            self.pos += 1
            if ch == '.':
                return ''
            return ch

    def parse_left_right(self, style: str = 'math') -> str:
        beg_chr = self.read_delimiter_token()
        depth = 1
        inner_chars = []
        end_chr = ')'
        while self.pos < self.n and depth > 0:
            if self.s.startswith(r'\left', self.pos) and (self.pos + 5 == self.n or not self.s[self.pos + 5].isalpha()):
                depth += 1
                inner_chars.append(r'\left')
                self.pos += 5
            elif self.s.startswith(r'\right', self.pos) and (self.pos + 6 == self.n or not self.s[self.pos + 6].isalpha()):
                depth -= 1
                self.pos += 6
                if depth == 0:
                    end_chr = self.read_delimiter_token()
                    break
                else:
                    inner_chars.append(r'\right')
            else:
                inner_chars.append(self.s[self.pos])
                self.pos += 1
        inner_xml = self.parse_sub_expr("".join(inner_chars), style=style)
        return (
            '<m:d>'
            f'<m:dPr><m:begChr m:val="{escape(beg_chr)}"/><m:endChr m:val="{escape(end_chr)}"/><m:grow m:val="1"/>{self._ctrl_pr()}</m:dPr>'
            f'<m:e>{inner_xml}</m:e>'
            '</m:d>'
        )

    def parse_environment(self, env_name: str, style: str = 'math') -> str:
        end_tag = f"\\end{{{env_name}}}"
        idx = self.s.find(end_tag, self.pos)
        if idx == -1:
            body_str = self.s[self.pos:]
            self.pos = self.n
        else:
            body_str = self.s[self.pos:idx]
            self.pos = idx + len(end_tag)

        rows = [r.strip() for r in re.split(r'\\\\', body_str) if r.strip()]
        if env_name == 'cases':
            eq_items = []
            for r in rows:
                parts = [p.strip() for p in r.split('&')]
                row_xml = (self._run("\u2003")).join(self.parse_sub_expr(p, style=style) for p in parts)
                eq_items.append(f'<m:e>{row_xml}</m:e>')
            return (
                f'<m:d><m:dPr><m:begChr m:val="{{"/><m:sepChr m:val=""/><m:endChr m:val=""/><m:grow m:val="1"/>{self._ctrl_pr()}</m:dPr>'
                f'<m:e><m:eqArr><m:eqArrPr/>{"".join(eq_items)}</m:eqArr></m:e></m:d>'
            )
        elif 'matrix' in env_name:
            m_rows = []
            for r in rows:
                cols = [c.strip() for c in r.split('&')]
                c_xml = "".join(f'<m:e>{self.parse_sub_expr(c, style=style)}</m:e>' for c in cols)
                m_rows.append(f'<m:mr>{c_xml}</m:mr>')
            mat_xml = f'<m:m><m:mPr/>{"".join(m_rows)}</m:m>'
            delims = {
                'pmatrix': ('(', ')'),
                'bmatrix': ('[', ']'),
                'Bmatrix': ('{', '}'),
                'vmatrix': ('|', '|'),
                'Vmatrix': ('‖', '‖'),
            }
            if env_name in delims:
                b_c, e_c = delims[env_name]
                return f'<m:d><m:dPr><m:begChr m:val="{b_c}"/><m:endChr m:val="{e_c}"/><m:grow m:val="1"/>{self._ctrl_pr()}</m:dPr><m:e>{mat_xml}</m:e></m:d>'
            return mat_xml
        else:
            eq_items = []
            for r in rows:
                clean_r = r.replace('&', ' ')
                eq_items.append(f'<m:e>{self.parse_sub_expr(clean_r, style=style)}</m:e>')
            return f'<m:eqArr><m:eqArrPr/>{"".join(eq_items)}</m:eqArr>'

    def parse_atom(self, style: str = 'math') -> Tuple[str, str]:
        if self.pos >= self.n:
            return ('empty', '')

        if style == 'upright':
            if self.s[self.pos] in (' ', '\t'):
                start = self.pos
                while self.pos < self.n and self.s[self.pos] in (' ', '\t'):
                    self.pos += 1
                return ('text', self.s[start:self.pos])
        else:
            self.skip_spaces()
            if self.pos >= self.n:
                return ('empty', '')

        ch = self.s[self.pos]

        if ch == '{':
            grp = self.read_balanced_group('{', '}')
            if not grp:
                return ('empty', '')
            return ('xml', self.parse_sub_expr(grp, style=style))

        if ch == '\\':
            self.pos += 1
            cmd = self.read_command()
            if not cmd:
                return ('empty', '')

            if cmd in NOOP_COMMANDS:
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return ('empty', '')

            if cmd == '{':
                depth_b = 1
                inner_b = []
                while self.pos < self.n and depth_b > 0:
                    if self.s.startswith(r'\{', self.pos):
                        depth_b += 1
                        inner_b.append(r'\{')
                        self.pos += 2
                    elif self.s.startswith(r'\}', self.pos):
                        depth_b -= 1
                        self.pos += 2
                        if depth_b == 0:
                            break
                        else:
                            inner_b.append(r'\}')
                    else:
                        inner_b.append(self.s[self.pos])
                        self.pos += 1
                inner_b_xml = self.parse_sub_expr("".join(inner_b), style=style)
                return (
                    'xml',
                    f'<m:d><m:dPr><m:begChr m:val="{{"/><m:endChr m:val="}}"/><m:grow m:val="1"/>{self._ctrl_pr()}</m:dPr>'
                    f'<m:e>{inner_b_xml}</m:e></m:d>'
                )
            if cmd == '}':
                return ('empty', '')
            if cmd in ('%', '&', '#', '$', '_'):
                return ('xml', self._run(cmd, style='upright'))
            if cmd in (';', ':'):
                return ('xml', self._run(' ', style='upright'))
            if cmd in (',', ' '):
                return ('xml', self._run(' ', style='upright'))
            if cmd in ('quad', 'qquad', 'enspace', 'thinspace'):
                sp_map = {'quad': '    ', 'qquad': '        ', 'enspace': '  ', 'thinspace': ' '}
                return ('xml', self._run(sp_map[cmd], style='upright'))
            if cmd in ('vert', 'lvert', 'rvert', 'mid', '|'):
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return ('xml', self._run('|', style='upright'))
            if cmd in ('Vert', 'lVert', 'rVert'):
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return ('xml', self._run('‖', style='upright'))
            if cmd == '!':
                return ('empty', '')

            if cmd in ('hspace', 'vspace', 'phantom', 'hphantom', 'vphantom'):
                self.read_balanced_group('{', '}')
                return ('text', ' ')

            if cmd == 'color':
                self.read_balanced_group('{', '}')
                return ('empty', '')
            if cmd == 'textcolor':
                self.read_balanced_group('{', '}')
                val_str = self.read_balanced_group('{', '}')
                return ('xml', self.parse_sub_expr(val_str, style=style))

            if cmd == 'left':
                return ('xml', self.parse_left_right(style=style))
            if cmd == 'right':
                self.read_delimiter_token()
                return ('empty', '')

            if cmd == 'begin':
                env_name = self.read_balanced_group('{', '}').strip()
                return ('xml', self.parse_environment(env_name, style=style))

            if cmd in FRAC_COMMANDS:
                num_str = self.read_balanced_group('{', '}')
                den_str = self.read_balanced_group('{', '}')
                num_xml = self.parse_sub_expr(num_str, style=style)
                den_xml = self.parse_sub_expr(den_str, style=style)
                return ('xml', f'<m:f><m:fPr>{self._ctrl_pr()}</m:fPr><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>')

            if cmd == 'binom':
                n_str = self.read_balanced_group('{', '}')
                k_str = self.read_balanced_group('{', '}')
                n_xml = self.parse_sub_expr(n_str, style=style)
                k_xml = self.parse_sub_expr(k_str, style=style)
                return (
                    'xml',
                    f'<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/><m:grow m:val="1"/>{self._ctrl_pr()}</m:dPr>'
                    f'<m:e><m:f><m:fPr><m:type m:val="noBar"/>{self._ctrl_pr()}</m:fPr><m:num>{n_xml}</m:num><m:den>{k_xml}</m:den></m:f></m:e></m:d>'
                )

            if cmd == 'sqrt':
                self.skip_spaces()
                deg_str = None
                if self.pos < self.n and self.s[self.pos] == '[':
                    deg_str = self.read_balanced_group('[', ']')
                val_str = self.read_balanced_group('{', '}')
                val_xml = self.parse_sub_expr(val_str, style=style)
                if deg_str:
                    deg_xml = self.parse_sub_expr(deg_str, style=style)
                    return ('xml', f'<m:rad><m:radPr>{self._ctrl_pr()}</m:radPr><m:deg>{deg_xml}</m:deg><m:e>{val_xml}</m:e></m:rad>')
                return ('xml', f'<m:rad><m:radPr><m:degHide m:val="1"/>{self._ctrl_pr()}</m:radPr><m:deg/><m:e>{val_xml}</m:e></m:rad>')

            if cmd in ACCENT_COMMANDS:
                acc_chr = ACCENT_COMMANDS[cmd]
                self.skip_spaces()
                if self.pos < self.n and self.s[self.pos] == '{':
                    arg_str = self.read_balanced_group('{', '}')
                elif self.pos < self.n:
                    arg_str = self.s[self.pos]
                    self.pos += 1
                else:
                    arg_str = ""
                arg_xml = self.parse_sub_expr(arg_str, style=style)
                return ('xml', f'<m:acc><m:accPr><m:chr m:val="{acc_chr}"/>{self._ctrl_pr()}</m:accPr><m:e>{arg_xml}</m:e></m:acc>')

            if cmd in ('overline', 'underline'):
                pos_val = 'top' if cmd == 'overline' else 'bot'
                arg_str = self.read_balanced_group('{', '}')
                arg_xml = self.parse_sub_expr(arg_str, style=style)
                return ('xml', f'<m:bar><m:barPr><m:pos m:val="{pos_val}"/>{self._ctrl_pr()}</m:barPr><m:e>{arg_xml}</m:e></m:bar>')

            if cmd in ('overbrace', 'underbrace'):
                pos_val = 'top' if cmd == 'overbrace' else 'bot'
                chr_val = '⏞' if cmd == 'overbrace' else '⏟'
                arg_str = self.read_balanced_group('{', '}')
                arg_xml = self.parse_sub_expr(arg_str, style=style)
                return ('xml', f'<m:groupChr><m:groupChrPr><m:chr m:val="{chr_val}"/><m:pos m:val="{pos_val}"/>{self._ctrl_pr()}</m:groupChrPr><m:e>{arg_xml}</m:e></m:groupChr>')

            if cmd in ('overset', 'stackrel'):
                top_str = self.read_balanced_group('{', '}')
                base_str = self.read_balanced_group('{', '}')
                return ('xml', f'<m:limUpp><m:limUppPr>{self._ctrl_pr()}</m:limUppPr><m:e>{self.parse_sub_expr(base_str, style=style)}</m:e><m:lim>{self.parse_sub_expr(top_str, style=style)}</m:lim></m:limUpp>')

            if cmd == 'underset':
                bot_str = self.read_balanced_group('{', '}')
                base_str = self.read_balanced_group('{', '}')
                return ('xml', f'<m:limLow><m:limLowPr>{self._ctrl_pr()}</m:limLowPr><m:e>{self.parse_sub_expr(base_str, style=style)}</m:e><m:lim>{self.parse_sub_expr(bot_str, style=style)}</m:lim></m:limLow>')

            if cmd == 'boxed':
                arg_str = self.read_balanced_group('{', '}')
                return ('xml', f'<m:borderBox><m:borderBoxPr>{self._ctrl_pr()}</m:borderBoxPr><m:e>{self.parse_sub_expr(arg_str, style=style)}</m:e></m:borderBox>')

            if cmd in FONT_UPRIGHT_COMMANDS:
                arg_str = self.read_balanced_group('{', '}')
                return ('xml', self.parse_sub_expr(arg_str, style='upright'))

            if cmd in FONT_BOLD_COMMANDS:
                arg_str = self.read_balanced_group('{', '}')
                return ('xml', self.parse_sub_expr(arg_str, style='bold'))

            if cmd in FONT_ITALIC_COMMANDS:
                arg_str = self.read_balanced_group('{', '}')
                return ('xml', self.parse_sub_expr(arg_str, style='math'))

            if cmd in UPRIGHT_FUNCTIONS:
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                save_pos = self.pos
                self.skip_spaces()
                next_ch = self.peek()
                self.pos = save_pos
                suffix = "\u2009" if (next_ch and next_ch not in ('(', '[', '_', '^', '{')) else ""
                return ('xml', self._run(cmd + suffix, style='upright'))

            if cmd in GREEK_AND_SYMBOLS:
                if self.pos + 1 < self.n and self.s[self.pos:self.pos + 2] == '{}':
                    self.pos += 2
                return ('text', GREEK_AND_SYMBOLS[cmd])

            if self.pos < self.n and self.s[self.pos] == '{':
                arg_str = self.read_balanced_group('{', '}')
                return ('xml', self.parse_sub_expr(arg_str, style=style))
            return ('text', cmd)

        if ch.isdigit():
            start = self.pos
            while self.pos < self.n and (
                self.s[self.pos].isdigit() or
                (self.s[self.pos] in (',', '.') and self.pos + 1 < self.n and self.s[self.pos + 1].isdigit())
            ):
                self.pos += 1
            return ('text', self.s[start:self.pos])

        if style == 'upright' and ch.isalpha():
            start = self.pos
            while self.pos < self.n and (self.s[self.pos].isalnum() or self.s[self.pos] in (' ', '-', '–', '.', ',')):
                if self.s[self.pos] in ('\\', '{', '}', '_', '^', '$'):
                    break
                self.pos += 1
            return ('text', self.s[start:self.pos])

        self.pos += 1
        if ch in ('<', '>', '|') and style != 'upright':
            sp_prefix = ' ' if (ch in ('<', '>') and self.pos < self.n) else ''
            return ('xml', self._run(f'{ch}{sp_prefix}', style='upright'))
        if ch in (',', ';') and style != 'upright':
            return ('xml', self._run(f'{ch} ', style='upright'))
        if ch == '-' and style != 'upright':
            ch = '−'
        elif ch == "'":
            ch = '′'
        return ('text', ch)

    def _atom_to_xml(self, kind: str, content: str, style: str) -> str:
        if kind == 'xml':
            return content
        return self._run(content, style=style)

    def parse_sequence(self, stop_tokens: set, style: str = 'math') -> str:
        raw_items: List[Tuple[str, str, str]] = []

        while self.pos < self.n:
            if style != 'upright':
                self.skip_spaces()
            if self.pos >= self.n or self.s[self.pos] in stop_tokens:
                break

            ch = self.s[self.pos]

            if ch in ('^', '_') and not raw_items:
                self.pos += 1
                save_pos = self.pos
                self.skip_spaces()
                if self.s.startswith(r'\circ', self.pos) and (self.pos + 5 == self.n or not self.s[self.pos + 5].isalpha()):
                    self.pos += 5
                    raw_items.append(('text', '°', 'upright'))
                    continue
                elif self.s.startswith(r'{\circ}', self.pos):
                    self.pos += 7
                    raw_items.append(('text', '°', 'upright'))
                    continue
                self.pos = save_pos
                script_xml = self.read_script_arg(style=style)
                empty_base = self._run('\u200B', style=style)
                if ch == '^':
                    raw_items.append(('xml', f'<m:sSup><m:sSupPr>{self._ctrl_pr()}</m:sSupPr><m:e>{empty_base}</m:e><m:sup>{script_xml}</m:sup></m:sSup>', style))
                else:
                    raw_items.append(('xml', f'<m:sSub><m:sSubPr>{self._ctrl_pr()}</m:sSubPr><m:e>{empty_base}</m:e><m:sub>{script_xml}</m:sub></m:sSub>', style))
                continue

            kind, content = self.parse_atom(style=style)
            if kind == 'empty' or not content:
                continue

            if style != 'upright':
                self.skip_spaces()
            while self.pos < self.n and self.s[self.pos] == "'":
                self.pos += 1
                if kind == 'text':
                    content += '′'
                else:
                    content = content + self._run('′', style=style)
                if style != 'upright':
                    self.skip_spaces()

            sub_xml = None
            sup_xml = None
            while self.pos < self.n and self.s[self.pos] in ('_', '^'):
                op = self.s[self.pos]
                self.pos += 1
                if op == '_':
                    sub_xml = self.read_script_arg(style=style)
                else:
                    save_p = self.pos
                    self.skip_spaces()
                    if self.s.startswith(r'\circ', self.pos) and (self.pos + 5 == self.n or not self.s[self.pos + 5].isalpha()):
                        self.pos += 5
                        if kind == 'text':
                            content += '°'
                        else:
                            sup_xml = self._run('°', style='upright')
                    elif self.s.startswith(r'{\circ}', self.pos):
                        self.pos += 7
                        if kind == 'text':
                            content += '°'
                        else:
                            sup_xml = self._run('°', style='upright')
                    else:
                        self.pos = save_p
                        sup_xml = self.read_script_arg(style=style)
                if style != 'upright':
                    self.skip_spaces()

            if sub_xml is not None or sup_xml is not None:
                base_xml = self._atom_to_xml(kind, content, style)
                if sub_xml is not None and sup_xml is not None:
                    wrapped = f'<m:sSubSup><m:sSubSupPr>{self._ctrl_pr()}</m:sSubSupPr><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub><m:sup>{sup_xml}</m:sup></m:sSubSup>'
                elif sub_xml is not None:
                    wrapped = f'<m:sSub><m:sSubPr>{self._ctrl_pr()}</m:sSubPr><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub></m:sSub>'
                else:
                    wrapped = f'<m:sSup><m:sSupPr>{self._ctrl_pr()}</m:sSupPr><m:e>{base_xml}</m:e><m:sup>{sup_xml}</m:sup></m:sSup>'
                raw_items.append(('xml', wrapped, style))
            else:
                raw_items.append((kind, content, style))

        if style != 'upright' and len(raw_items) > 1:
            spaced_items: List[Tuple[str, str, str]] = []
            for idx, (k, c, st) in enumerate(raw_items):
                prev_tok = raw_items[idx - 1][1] if idx > 0 else None
                next_tok = raw_items[idx + 1][1] if idx + 1 < len(raw_items) else None

                if k == 'text' and c in BINARY_RELATIONS:
                    if prev_tok is not None and prev_tok not in OPENING_TOKENS and not str(prev_tok).endswith(' '):
                        spaced_items.append(('text', ' ', st))
                    spaced_items.append((k, c, st))
                    if next_tok is not None and not str(next_tok).startswith(' '):
                        spaced_items.append(('text', ' ', st))
                elif k == 'text' and c in BINARY_OPERATORS:
                    is_unary = (
                        idx == 0 or
                        (raw_items[idx - 1][0] == 'text' and (
                            raw_items[idx - 1][1] in OPENING_TOKENS or
                            raw_items[idx - 1][1] in BINARY_RELATIONS or
                            raw_items[idx - 1][1] in BINARY_OPERATORS or
                            raw_items[idx - 1][1].strip() == ''
                        ))
                    )
                    if not is_unary:
                        if prev_tok is not None and not str(prev_tok).endswith(' '):
                            spaced_items.append(('text', ' ', st))
                        spaced_items.append((k, c, st))
                        if next_tok is not None and not str(next_tok).startswith(' '):
                            spaced_items.append(('text', ' ', st))
                    else:
                        spaced_items.append((k, c, st))
                elif k == 'text' and c in (',', ';'):
                    spaced_items.append((k, c, st))
                    if next_tok is not None and next_tok not in (' ', '\u2005', '\u2009', '\u2003', '\u2003\u2003'):
                        spaced_items.append(('text', ' ', st))
                else:
                    spaced_items.append((k, c, st))
            raw_items = spaced_items

        merged_xml = []
        buf_text = []
        buf_style = None

        for k, c, st in raw_items:
            if k == 'text':
                if buf_style == st:
                    buf_text.append(c)
                else:
                    if buf_text:
                        merged_xml.append(self._run("".join(buf_text), style=buf_style))
                    buf_text = [c]
                    buf_style = st
            else:
                if buf_text:
                    merged_xml.append(self._run("".join(buf_text), style=buf_style))
                    buf_text = []
                    buf_style = None
                merged_xml.append(c)

        if buf_text:
            merged_xml.append(self._run("".join(buf_text), style=buf_style))

        return "".join(merged_xml)


@lru_cache(maxsize=128)
def latex_to_omml(latex_str: str, size_pt: Optional[float] = None, bold: bool = False) -> str:
    return LaTeXToOMMLCompiler(latex_str, size_pt=size_pt, bold=bold).compile()


def render_inlines_to_paragraph(
    p,
    text: str,
    is_heading: bool = False,
    font_size_pt: Optional[float] = None,
    strip_emphasis: bool = False
):
    if font_size_pt is None and p.style and p.style.font and p.style.font.size:
        font_size_pt = p.style.font.size.pt

    tokens = tokenize_inlines(text)
    for tok in tokens:
        if tok.get('footnote'):
            fn_run = p.add_run(f"[{tok['footnote']}]")
            fn_run.font.superscript = True
            fn_run.font.size = Pt(10)
            continue

        if tok['math']:
            omml_str = latex_to_omml(tok['val'], size_pt=font_size_pt, bold=is_heading)
            p._p.append(parse_xml(omml_str))
        elif tok['code']:
            run = p.add_run(tok['val'])
            run.font.name = 'Courier New'
            run.font.size = Pt(font_size_pt if font_size_pt and font_size_pt <= 12 else 11)
            if is_heading:
                run.bold = True
        else:
            run = p.add_run(tok['val'])
            if strip_emphasis and not is_heading:
                run.bold = False
                run.italic = False
            else:
                run.bold = True if is_heading else tok['bold']
                if tok['italic']:
                    run.italic = True
            if tok['strike']:
                run.font.strike = True
            if tok['underline']:
                run.font.underline = True
            if tok['sub']:
                run.font.subscript = True
            if tok['sup']:
                run.font.superscript = True
            if tok['highlight']:
                rPr = run._r.get_or_add_rPr()
                rPr.append(parse_xml(f'<w:highlight {nsdecls("w")} w:val="yellow"/>'))
            if tok['kbd']:
                run.font.name = 'Courier New'
                run.font.size = Pt(10.5)


def _is_centered_table_cell(val: str) -> bool:
    s = (val or "").strip()
    if not s or s in ('–', '—', '-', '+', '±', '☐', '☑'):
        return True

    without_math = re.sub(r'\$\$[\s\S]+?\$\$|\$[^$]+\$', '0', s)
    without_emphasis = re.sub(r'[*_`]', '', without_math).strip()

    words = re.findall(r'[а-яА-ЯёЁa-zA-Z]{4,}', without_emphasis)
    if words:
        return False
    all_words = re.findall(r'[а-яА-ЯёЁa-zA-Z]{2,}', without_emphasis)
    if len(all_words) >= 2:
        return False
    return True


def _extract_visible_cell_text(raw_c: str) -> str:
    vis_c = re.sub(r'\\(?:mathrm|text|textrm|mathbf|mathit)\s*\{([^{}]*)\}', r'\1', raw_c or '')
    vis_c = re.sub(r'\\frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}', lambda m: m.group(1) if len(m.group(1)) >= len(m.group(2)) else m.group(2), vis_c)
    vis_c = re.sub(r'\\[a-zA-Z]+\{?\}?', 'X', vis_c)
    vis_c = re.sub(r'[${}_^\\*`]', '', vis_c)
    return vis_c.strip()


def _percentile_75(values: List[float]) -> float:
    if not values:
        return 0.0
    s = sorted(values)
    if len(s) == 1:
        return float(s[0])
    pos = 0.75 * (len(s) - 1)
    lo = int(pos)
    hi = min(lo + 1, len(s) - 1)
    frac = pos - lo
    return float(s[lo] + frac * (s[hi] - s[lo]))


def compute_smart_column_widths_mm(
    header: List[str],
    rows: List[List[str]],
    total_w_mm: float,
    cell_size_pt: float = 12.0,
    h_padding_mm: float = 3.6
) -> List[float]:
    """
    Вычисляет ширину столбцов таблицы на основе:
    - минимальной неразрывной ширины (по самому длинному слову или формуле);
    - оптимальной ширины (по 75-му перцентилю длины текста в ячейках с учетом кегля);
    - пропорционального распределения доступных миллиметров по информационной плотности.
    """
    cols = len(header)
    if cols == 0:
        return []

    char_w_mm = max(1.2, cell_size_pt * 0.352778 * 0.48)
    all_rows = [header] + rows

    min_widths = []
    opt_widths = []

    for c_i in range(cols):
        cell_lengths: List[float] = []
        max_unbreakable_chars = 3

        for r_i, r_cells in enumerate(all_rows):
            raw_c = r_cells[c_i] if c_i < len(r_cells) else ""
            vis_text = _extract_visible_cell_text(raw_c)
            eff_len = len(vis_text)
            cell_lengths.append(float(eff_len) * (1.05 if r_i == 0 else 1.0))

            # Самое длинное неразрывное слово или компактная формула
            tokens = re.findall(r'[^\s]+', vis_text)
            for tok in tokens:
                # Учитываем дефисные слова как потенциально переносимые только если очень длинные
                sub_parts = tok.split('-') if len(tok) > 16 else [tok]
                for sp in sub_parts:
                    if len(sp) > max_unbreakable_chars:
                        max_unbreakable_chars = len(sp)

        min_mm = max(10.0, max_unbreakable_chars * char_w_mm * 1.08 + h_padding_mm)
        p75_len = _percentile_75(cell_lengths)
        opt_mm = max(min_mm, p75_len * char_w_mm + h_padding_mm)

        min_widths.append(min_mm)
        opt_widths.append(opt_mm)

    sum_min = sum(min_widths)
    if sum_min >= total_w_mm:
        scale = total_w_mm / sum_min
        return [w * scale for w in min_widths]

    remaining_mm = total_w_mm - sum_min
    demands = [max(0.5, opt_widths[i] - min_widths[i]) for i in range(cols)]
    sum_demands = sum(demands) or 1.0

    # Если все столбцы уже помещаются в opt_widths, распределяем остаток пропорционально opt_widths
    if sum(opt_widths) <= total_w_mm:
        base_sum = sum(opt_widths) or 1.0
        extra_pool = total_w_mm - base_sum
        return [opt_widths[i] + extra_pool * (opt_widths[i] / base_sum) for i in range(cols)]

    return [min_widths[i] + remaining_mm * (demands[i] / sum_demands) for i in range(cols)]


STYLE_CANONICAL_MAP = {
    'заголовок 1': 'Heading 1',
    'заголовок1': 'Heading 1',
    'heading 1': 'Heading 1',
    'gost_heading_1': 'Heading 1',
    'h1': 'Heading 1',
    'раздел': 'Heading 1',

    'заголовок 2': 'Heading 2',
    'заголовок2': 'Heading 2',
    'heading 2': 'Heading 2',
    'gost_heading_2': 'Heading 2',
    'h2': 'Heading 2',
    'подраздел': 'Heading 2',

    'заголовок 3': 'Heading 3',
    'заголовок3': 'Heading 3',
    'heading 3': 'Heading 3',
    'gost_heading_3': 'Heading 3',
    'h3': 'Heading 3',
    'пункт': 'Heading 3',

    'заголовок структурный': 'Title',
    'структурный заголовок': 'Title',
    'заголовок': 'Title',
    'title': 'Title',
    'gost_heading_structural': 'Title',

    'обычный': 'Normal',
    'основной текст': 'Normal',
    'текст': 'Normal',
    'normal': 'Normal',
    'body': 'Normal',
    'gost_body': 'Normal',

    'примечание': 'Примечание',
    'сноска': 'Примечание',
    'gost_note': 'Примечание',

    'цитата': 'Quote',
    'выноска': 'Quote',
    'quote': 'Quote',
    'gost_quote': 'Quote',

    'эпиграф': 'Эпиграф',
    'gost_epigraph': 'Эпиграф',

    'подпись таблицы': 'Таблица Название',
    'название таблицы': 'Таблица Название',
    'gost_table_caption': 'Таблица Название',

    'ячейка таблицы': 'Таблица Ячейка',
    'текст таблицы': 'Таблица Ячейка',
    'gost_table_cell': 'Таблица Ячейка',

    'рисунок': 'GOST_Figure',
    'иллюстрация': 'GOST_Figure',
    'gost_figure': 'GOST_Figure',

    'подпись рисунка': 'Caption',
    'подпись к рисунку': 'Caption',
    'подпись': 'Caption',
    'caption': 'Caption',
    'gost_figure_caption': 'Caption',

    'листинг код': 'Листинг Код',
    'исходный код': 'Листинг Код',
    'код': 'Листинг Код',
    'gost_listing_code': 'Листинг Код',

    'листинг название': 'Листинг Название',
    'название листинга': 'Листинг Название',
    'подпись листинга': 'Листинг Название',
    'gost_listing_caption': 'Листинг Название',
}


WORD_BUILTIN_STYLES = {
    'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4',
    'Heading 5', 'Heading 6', 'Heading 7', 'Heading 8', 'Heading 9',
    'Title', 'Subtitle', 'Normal', 'Caption', 'Quote',
    'Header', 'Footer', 'List Paragraph', 'Body Text'
}

RUSSIAN_ALIASES = {
    'Heading 1': 'Заголовок 1,ГОСТ Заголовок 1,Раздел',
    'Heading 2': 'Заголовок 2,ГОСТ Заголовок 2,Подраздел',
    'Heading 3': 'Заголовок 3,ГОСТ Заголовок 3,Пункт',
    'Title': 'Заголовок структурный,Структурный заголовок',
    'Normal': 'Обычный,Основной текст,ГОСТ Основной текст',
    'Caption': 'Подпись,Название объекта,Рисунок',
    'Quote': 'Цитата,Блок цитирования',
}


class GostDocxFormatter:
    def __init__(
        self,
        template_cfg: dict,
        existing_docx_path: Optional[str] = None,
        preserve_existing_styles: bool = True,
        update_existing_styles: bool = False,
        strip_body_emphasis: Optional[bool] = None,
    ):
        self.cfg = template_cfg
        self.existing_docx_path = existing_docx_path
        self.is_append_mode = bool(existing_docx_path)
        self.preserve_existing_styles = bool(preserve_existing_styles and not update_existing_styles)
        self.strip_body_emphasis = bool(self.cfg.get('strip_body_emphasis', False) if strip_body_emphasis is None else strip_body_emphasis)

        if existing_docx_path:
            self.doc = Document(existing_docx_path)
            sec = self.doc.sections[-1] if self.doc.sections else None
            emu_per_mm = 36000.0
            self.page_cfg = dict(self.cfg.get('page', {}))
            if sec is not None:
                if sec.page_width:
                    self.page_cfg['width_mm'] = sec.page_width / emu_per_mm
                if sec.page_height:
                    self.page_cfg['height_mm'] = sec.page_height / emu_per_mm
                if sec.left_margin is not None:
                    self.page_cfg['margin_left_mm'] = sec.left_margin / emu_per_mm
                if sec.right_margin is not None:
                    self.page_cfg['margin_right_mm'] = sec.right_margin / emu_per_mm
                cols_el = sec._sectPr.find(qn('w:cols'))
                if cols_el is not None and cols_el.get(qn('w:num')):
                    self.num_columns = max(1, int(cols_el.get(qn('w:num'))))
                else:
                    self.num_columns = 1
            else:
                self.num_columns = 1
            self.column_spacing_mm = float(self.page_cfg.get('column_spacing_mm', 5.0))
            self.span_headings = bool(self.page_cfg.get('span_headings', True))
            self.different_first_page = False
            self.current_columns = self.num_columns
            self._has_content = True
            self.style_obj_map = {}
            self._index_existing_styles()
            self._register_styles(preserve_existing=self.preserve_existing_styles)
        else:
            self.doc = Document()
            self.page_cfg = self.cfg.get('page', {})
            self.num_columns = max(1, int(self.page_cfg.get('columns', 1) or 1))
            self.column_spacing_mm = float(self.page_cfg.get('column_spacing_mm', 5.0))
            self.span_headings = bool(self.page_cfg.get('span_headings', True))
            self.different_first_page = bool(self.page_cfg.get('different_first_page', False))
            self.current_columns = self.num_columns
            self._has_content = False
            self.style_obj_map = {}

            self._setup_page()
            self._register_styles(preserve_existing=False)

    def _apply_section_geometry(self, section, columns: int, is_first_section: bool = False):
        p = self.page_cfg
        section.page_width = Mm(p.get('width_mm', 210.0))
        section.page_height = Mm(p.get('height_mm', 297.0))
        section.left_margin = Mm(p.get('margin_left_mm', 30.0))
        section.right_margin = Mm(p.get('margin_right_mm', 20.0))
        section.top_margin = Mm(p.get('margin_top_mm', 20.0))
        section.bottom_margin = Mm(p.get('margin_bottom_mm', 20.0))

        if is_first_section and self.different_first_page:
            section.different_first_page_header_footer = True
            # Первая страница без колонтитулов
            for hdr_ftr in (section.first_page_header, section.first_page_footer):
                for par in hdr_ftr.paragraphs:
                    par.text = ""

        if is_first_section and p.get('page_numbers'):
            ftr = section.footer
            fp = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = fp.add_run()
            r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
            r._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'))
            r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
            r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))

        sectPr = section._sectPr
        for old_cols in sectPr.findall(qn('w:cols')):
            sectPr.remove(old_cols)
        space_dxa = int(round(self.column_spacing_mm * DXA_PER_MM))
        cols_el = parse_xml(f'<w:cols {nsdecls("w")} w:num="{columns}" w:space="{space_dxa}"/>')
        sectPr.append(cols_el)

    def _setup_page(self):
        s = self.doc.sections[0]
        self._apply_section_geometry(s, columns=self.num_columns, is_first_section=True)

    def _ensure_columns(self, target_cols: int):
        if self.num_columns <= 1:
            return
        if self.current_columns == target_cols:
            return
        if not self._has_content:
            self._apply_section_geometry(self.doc.sections[-1], columns=target_cols, is_first_section=(len(self.doc.sections) == 1))
            self.current_columns = target_cols
            return
        sec = self.doc.add_section(WD_SECTION.CONTINUOUS)
        sec.header.is_linked_to_previous = True
        sec.footer.is_linked_to_previous = True
        self._apply_section_geometry(sec, columns=target_cols, is_first_section=False)
        self.current_columns = target_cols

    def _get_full_text_width_mm(self) -> float:
        return (
            self.page_cfg.get('width_mm', 210.0)
            - self.page_cfg.get('margin_left_mm', 30.0)
            - self.page_cfg.get('margin_right_mm', 20.0)
        )

    def _get_current_text_width_mm(self) -> float:
        full_w = self._get_full_text_width_mm()
        if self.current_columns > 1:
            return (full_w - (self.current_columns - 1) * self.column_spacing_mm) / self.current_columns
        return full_w

    def _is_wide_table(self, node: dict) -> bool:
        if self.num_columns <= 1:
            return False
        header = node.get('header', [])
        rows = node.get('rows', [])
        cols = len(header)
        if cols >= 3:
            return True
        single_col_w = (self._get_full_text_width_mm() - (self.num_columns - 1) * self.column_spacing_mm) / self.num_columns
        cell_size_pt = self._get_style_cfg('GOST_Table_Cell').get('size_pt', 12.0)
        h_pad = self._get_table_h_padding_mm()
        widths = compute_smart_column_widths_mm(header, rows, single_col_w, cell_size_pt=cell_size_pt, h_padding_mm=h_pad)
        # Проверяем, не сжались ли столбцы слишком сильно в одной колонке
        char_w_mm = max(1.2, cell_size_pt * 0.352778 * 0.48)
        raw_need = 0.0
        for c_i in range(cols):
            max_w = max([len(_extract_visible_cell_text(r[c_i] if c_i < len(r) else "")) for r in [header] + rows] + [4])
            raw_need += min(55.0, max_w * char_w_mm + h_pad)
        return raw_need > single_col_w * 1.15 or min(widths or [999]) < 16.0

    def _resolve_style_name(self, raw_name: str) -> str:
        clean = raw_name.strip().lower()
        return STYLE_CANONICAL_MAP.get(clean, raw_name)

    def _get_style_cfg(self, name_or_key: str) -> dict:
        if not name_or_key:
            return {}
        styles_map = self.cfg.get('styles', {})
        if name_or_key in styles_map:
            return styles_map[name_or_key]
        clean = name_or_key.strip().lower()
        for k, v in styles_map.items():
            if k.lower() == clean:
                return v
        canonical = self._resolve_style_name(name_or_key)
        if canonical in styles_map:
            return styles_map[canonical]
        for k, v in styles_map.items():
            if self._resolve_style_name(k) == canonical:
                return v
            if v.get('ui_name') and v['ui_name'].lower() in (clean, canonical.lower()):
                return v
        return {}

    def _get_doc_style(self, name_or_key: str):
        if not name_or_key:
            return self.doc.styles['Normal']

        if hasattr(self, 'style_obj_map'):
            if name_or_key in self.style_obj_map:
                return self.style_obj_map[name_or_key]
            clean = name_or_key.strip().lower()
            if clean in self.style_obj_map:
                return self.style_obj_map[clean]

        canonical = self._resolve_style_name(name_or_key)
        if hasattr(self, 'style_obj_map'):
            if canonical in self.style_obj_map:
                return self.style_obj_map[canonical]
            if canonical.lower() in self.style_obj_map:
                return self.style_obj_map[canonical.lower()]

        for cand in [canonical, name_or_key]:
            try:
                return self.doc.styles[cand]
            except KeyError:
                pass

        return self.doc.styles['Normal']

    def _index_existing_styles(self):
        for s in self.doc.styles:
            if s.type != WD_STYLE_TYPE.PARAGRAPH:
                continue
            keys = set()
            if s.name:
                keys.add(s.name)
                keys.add(s.name.lower())
                canon = self._resolve_style_name(s.name)
                keys.add(canon)
                keys.add(canon.lower())
            if s.style_id:
                keys.add(s.style_id)
                keys.add(s.style_id.lower())
                canon_id = self._resolve_style_name(s.style_id)
                keys.add(canon_id)
                keys.add(canon_id.lower())
            aliases_el = s.element.find(qn('w:aliases'))
            if aliases_el is not None:
                aval = aliases_el.get(qn('w:val')) or ''
                for a in aval.split(','):
                    a_c = a.strip()
                    if a_c:
                        keys.add(a_c)
                        keys.add(a_c.lower())
            for k in keys:
                self.style_obj_map.setdefault(k, s)

    def _configure_doc_defaults(self):
        """Полная очистка и настройка <w:docDefaults> в word/styles.xml по параметрам шаблона."""
        styles_el = self.doc.styles.element
        doc_defaults = styles_el.find(qn('w:docDefaults'))
        if doc_defaults is None:
            doc_defaults = OxmlElement('w:docDefaults')
            styles_el.insert(0, doc_defaults)

        font_specs = self.cfg.get('fonts', {})
        primary_font = font_specs.get('primary', {}).get('name', 'Times New Roman')
        body_cfg = self._get_style_cfg('GOST_Body')
        body_sz_half_pt = int(round(body_cfg.get('size_pt', 14.0) * 2))
        body_line = int(round(body_cfg.get('line_spacing', 1.5) * 240))
        body_before = int(round(body_cfg.get('space_before_pt', 0.0) * 20))
        body_after = int(round(body_cfg.get('space_after_pt', 0.0) * 20))

        rPrDefault = doc_defaults.find(qn('w:rPrDefault'))
        if rPrDefault is None:
            rPrDefault = OxmlElement('w:rPrDefault')
            doc_defaults.append(rPrDefault)
        rPr = rPrDefault.find(qn('w:rPr'))
        if rPr is not None:
            rPrDefault.remove(rPr)
        rPr_xml = f"""
            <w:rPr {nsdecls("w")}>
                <w:rFonts w:ascii="{escape(primary_font)}" w:hAnsi="{escape(primary_font)}" w:eastAsia="{escape(primary_font)}" w:cs="{escape(primary_font)}"/>
                <w:sz w:val="{body_sz_half_pt}"/>
                <w:szCs w:val="{body_sz_half_pt}"/>
                <w:color w:val="000000"/>
                <w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>
            </w:rPr>
        """
        rPrDefault.append(parse_xml(rPr_xml))

        pPrDefault = doc_defaults.find(qn('w:pPrDefault'))
        if pPrDefault is None:
            pPrDefault = OxmlElement('w:pPrDefault')
            doc_defaults.append(pPrDefault)
        pPr = pPrDefault.find(qn('w:pPr'))
        if pPr is not None:
            pPrDefault.remove(pPr)
        pPr_xml = f"""
            <w:pPr {nsdecls("w")}>
                <w:spacing w:before="{body_before}" w:after="{body_after}" w:line="{body_line}" w:lineRule="auto"/>
            </w:pPr>
        """
        pPrDefault.append(parse_xml(pPr_xml))

    def _apply_clean_rpr_to_style_element(self, style_elem, f_data: dict, s_def: dict):
        """Очищает <w:rPr> стиля от заводских синих цветов Word (themeColor) и записывает параметры шаблона."""
        rPr = style_elem.get_or_add_rPr()
        tags_to_strip = [
            'w:rFonts', 'w:color', 'w:sz', 'w:szCs',
            'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u',
            'w:spacing', 'w:kern', 'w:caps', 'w:smallCaps'
        ]
        for t_name in tags_to_strip:
            for old_el in rPr.findall(qn(t_name)):
                rPr.remove(old_el)

        f_name = f_data.get('name', 'Times New Roman')
        rFonts = OxmlElement('w:rFonts')
        for attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
            rFonts.set(qn(f'w:{attr}'), f_data.get(attr, f_name))
        rPr.append(rFonts)

        if s_def.get('bold', False):
            rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
            rPr.append(parse_xml(f'<w:bCs {nsdecls("w")}/>'))
        else:
            rPr.append(parse_xml(f'<w:b {nsdecls("w")} w:val="0"/>'))
            rPr.append(parse_xml(f'<w:bCs {nsdecls("w")} w:val="0"/>'))

        if s_def.get('italic', False):
            rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
            rPr.append(parse_xml(f'<w:iCs {nsdecls("w")}/>'))
        else:
            rPr.append(parse_xml(f'<w:i {nsdecls("w")} w:val="0"/>'))
            rPr.append(parse_xml(f'<w:iCs {nsdecls("w")} w:val="0"/>'))

        if s_def.get('all_caps', False):
            rPr.append(parse_xml(f'<w:caps {nsdecls("w")}/>'))

        raw_color = str(s_def.get('color', '000000')).lstrip('#').upper()
        if len(raw_color) != 6:
            raw_color = '000000'
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{raw_color}"/>'))

        sz_half_pt = int(round(s_def.get('size_pt', 14.0) * 2))
        rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_half_pt}"/>'))
        rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz_half_pt}"/>'))

        for old_lang in rPr.findall(qn('w:lang')):
            rPr.remove(old_lang)
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>'))

    def _add_styled_paragraph(self, style_key: str):
        """Создает абзац с гарантированной записью <w:pStyle w:val='...'/> в XML документа."""
        style_obj = self._get_doc_style(style_key)
        p = self.doc.add_paragraph(style=style_obj)
        pPr = p._p.get_or_add_pPr()
        if pPr.find(qn('w:pStyle')) is None and style_obj and hasattr(style_obj, 'style_id'):
            pPr.insert(0, parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{style_obj.style_id}"/>'))
        return p

    def _register_styles(self, preserve_existing: bool = False):
        styles_map = self.cfg.get('styles', {})
        font_specs = self.cfg.get('fonts', {})

        align_lut = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }

        if not preserve_existing:
            self._configure_doc_defaults()
        if not hasattr(self, 'style_obj_map'):
            self.style_obj_map = {}

        for raw_s_name, s_def in styles_map.items():
            canonical_name = self._resolve_style_name(raw_s_name)

            # Если сохраняем существующие стили документа, ищем подходящий стиль в self.style_obj_map
            existing_style = None
            if preserve_existing:
                candidates = [raw_s_name, raw_s_name.lower(), canonical_name, canonical_name.lower()]
                if s_def.get('ui_name'):
                    candidates.extend([s_def['ui_name'], s_def['ui_name'].lower()])
                if s_def.get('aliases'):
                    for a in str(s_def['aliases']).split(','):
                        a_c = a.strip()
                        if a_c:
                            candidates.extend([a_c, a_c.lower()])
                for cand in candidates:
                    if cand in self.style_obj_map:
                        existing_style = self.style_obj_map[cand]
                        break
                if existing_style is not None:
                    for cand in candidates:
                        self.style_obj_map[cand] = existing_style
                    continue

            if canonical_name in WORD_BUILTIN_STYLES:
                target_name = canonical_name
                try:
                    style = self.doc.styles[target_name]
                except KeyError:
                    style = self.doc.styles.add_style(target_name, WD_STYLE_TYPE.PARAGRAPH)
            else:
                target_name = canonical_name
                try:
                    style = self.doc.styles[target_name]
                except KeyError:
                    try:
                        style = self.doc.styles.add_style(target_name, WD_STYLE_TYPE.PARAGRAPH)
                    except ValueError:
                        style = self.doc.styles[target_name]

            self.style_obj_map[raw_s_name] = style
            self.style_obj_map[raw_s_name.lower()] = style
            self.style_obj_map[canonical_name] = style
            self.style_obj_map[canonical_name.lower()] = style
            if s_def.get('ui_name'):
                self.style_obj_map[s_def['ui_name']] = style
                self.style_obj_map[s_def['ui_name'].lower()] = style

            f_key = s_def.get('font', 'primary')
            f_data = font_specs.get(f_key, font_specs.get('primary', {'name': 'Times New Roman'}))
            style.font.name = f_data['name']
            style.font.size = Pt(s_def.get('size_pt', 14.0))
            style.font.bold = s_def.get('bold', False)
            if s_def.get('italic'):
                style.font.italic = True

            style_elem = style.element
            if canonical_name in WORD_BUILTIN_STYLES:
                pPr_el = style_elem.find(qn('w:pPr'))
                if pPr_el is not None:
                    for tag_name in ('w:pBdr', 'w:shd'):
                        for old_el in pPr_el.findall(qn(tag_name)):
                            pPr_el.remove(old_el)
            self._apply_clean_rpr_to_style_element(style_elem, f_data, s_def)
            link_el = style_elem.find(qn('w:link'))
            if link_el is not None:
                linked_style_id = link_el.get(qn('w:val'))
                if linked_style_id:
                    for cand_s in self.doc.styles.element.findall(qn('w:style')):
                        if cand_s.get(qn('w:styleId')) == linked_style_id:
                            self._apply_clean_rpr_to_style_element(cand_s, f_data, s_def)
                            break
            alias_val = s_def.get('aliases') or RUSSIAN_ALIASES.get(canonical_name)
            if alias_val:
                alias_parts = [a.strip() for a in str(alias_val).split(',') if a.strip()]
                clean_aliases = []
                for a in alias_parts:
                    if a.lower() != canonical_name.lower() and a not in clean_aliases:
                        clean_aliases.append(a)
                    self.style_obj_map[a] = style
                    self.style_obj_map[a.lower()] = style

                if clean_aliases:
                    final_alias_str = ','.join(clean_aliases)
                    aliases_el = style_elem.find(qn('w:aliases'))
                    if aliases_el is None:
                        aliases_el = parse_xml(f'<w:aliases {nsdecls("w")} w:val="{escape(final_alias_str)}"/>')
                        style_elem.insert(1, aliases_el)
                    else:
                        aliases_el.set(qn('w:val'), final_alias_str)

            if style_elem.find(qn('w:qFormat')) is None:
                style_elem.append(parse_xml(f'<w:qFormat {nsdecls("w")}/>'))

            style.unhide_when_used = True
            style.hidden = False
            style.quick_style = True

            pPr = style.paragraph_format
            pPr.alignment = align_lut.get(s_def.get('alignment', 'justify'), WD_ALIGN_PARAGRAPH.JUSTIFY)
            pPr.line_spacing = s_def.get('line_spacing', 1.5)
            pPr.first_line_indent = Mm(s_def.get('first_line_indent_mm', 0.0))
            pPr.space_before = Pt(s_def.get('space_before_pt', 0.0))
            pPr.space_after = Pt(s_def.get('space_after_pt', 0.0))

            if s_def.get('keep_with_next'):
                pPr.keep_with_next = True
            if s_def.get('page_break_before'):
                pPr.page_break_before = True

            if 'outline_level' in s_def:
                pPr_el = style.element.get_or_add_pPr()
                for old_lvl in pPr_el.findall(qn('w:outlineLvl')):
                    pPr_el.remove(old_lvl)
                pPr_el.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{s_def["outline_level"]}"/>'))

    def build_document(self, ast_nodes: List[Dict[str, Any]], output_path: str, base_dir: Optional[str] = None):
        body_size_pt = self._get_style_cfg('GOST_Body').get('size_pt', 14.0)
        n_nodes = len(ast_nodes)

        for idx, node in enumerate(ast_nodes):
            ntype = node['type']
            if ntype != 'list_item':
                self._prev_node_was_list = False

            if ntype == 'heading':
                is_span_heading = (
                    self.num_columns > 1
                    and self.span_headings
                    and node.get('style') in ('GOST_Heading_1', 'GOST_Heading_Structural')
                )
                if is_span_heading:
                    self._ensure_columns(1)
                elif self.num_columns > 1:
                    self._ensure_columns(self.num_columns)

                h_size_pt = self._get_style_cfg(node['style']).get('size_pt', 16.0 if node['outline_level'] == 0 else 14.0)
                p = self.doc.add_paragraph(style=self._get_doc_style(node['style']))
                pPr = p._p.get_or_add_pPr()
                for old_lvl in pPr.findall(qn('w:outlineLvl')):
                    pPr.remove(old_lvl)
                pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{node["outline_level"]}"/>'))
                render_inlines_to_paragraph(p, node['text'], is_heading=True, font_size_pt=h_size_pt)
                self._has_content = True

                if is_span_heading:
                    # Если следующий узел — не другой span-заголовок и не широкая таблица, возвращаем колонки
                    nxt = ast_nodes[idx + 1] if idx + 1 < n_nodes else None
                    nxt_is_span = nxt and (
                        (nxt['type'] == 'heading' and nxt.get('style') in ('GOST_Heading_1', 'GOST_Heading_Structural'))
                        or (nxt['type'] == 'table' and self._is_wide_table(nxt))
                    )
                    if not nxt_is_span:
                        self._ensure_columns(self.num_columns)

            elif ntype == 'paragraph':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                txt = node['text']
                if re.match(r'^Примечани[ея]\b', txt):
                    p = self.doc.add_paragraph(style=self._get_doc_style('GOST_Note'))
                    render_inlines_to_paragraph(p, txt, is_heading=False, font_size_pt=12.0, strip_emphasis=self.strip_body_emphasis)
                else:
                    p = self._add_styled_paragraph('GOST_Body')
                    if re.match(r'^где\b', txt):
                        p.paragraph_format.first_line_indent = Mm(0)
                    render_inlines_to_paragraph(p, txt, is_heading=False, font_size_pt=body_size_pt, strip_emphasis=self.strip_body_emphasis)
                if node.get('keep_with_next'):
                    p.paragraph_format.keep_with_next = True
                self._has_content = True

            elif ntype == 'list_item':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                self._render_list_item(node, body_size_pt)
                self._has_content = True

            elif ntype == 'table_caption':
                nxt = ast_nodes[idx + 1] if idx + 1 < n_nodes else None
                if self.num_columns > 1 and nxt and nxt['type'] == 'table' and self._is_wide_table(nxt):
                    self._ensure_columns(1)
                elif self.num_columns > 1:
                    self._ensure_columns(self.num_columns)

                cap_size = self._get_style_cfg('GOST_Table_Caption').get('size_pt', 14.0)
                p = self.doc.add_paragraph(style=self._get_doc_style(self._resolve_style_name('GOST_Table_Caption')))
                p.paragraph_format.keep_with_next = True
                render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=cap_size)
                if node.get('unit'):
                    p_unit = self.doc.add_paragraph(style=self._get_doc_style('GOST_Table_Caption'))
                    p_unit.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    p_unit.paragraph_format.space_before = Pt(0)
                    p_unit.paragraph_format.keep_with_next = True
                    p_unit.add_run(node['unit'])
                self._has_content = True

            elif ntype == 'table':
                wide = self._is_wide_table(node)
                if wide:
                    self._ensure_columns(1)
                elif self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                self._render_table(node)
                self._has_content = True
                if wide:
                    self._ensure_columns(self.num_columns)

            elif ntype == 'figure_image':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                p = self.doc.add_paragraph(style=self._get_doc_style(self._resolve_style_name('GOST_Figure')))
                if node.get('keep_with_next'):
                    p.paragraph_format.keep_with_next = True
                img_path = self._resolve_image_path(node['path'], base_dir=base_dir)
                max_w_mm = self._get_current_text_width_mm() * 0.95
                if img_path and os.path.isfile(img_path):
                    req_w_mm = node['width_cm'] * 10.0 if node.get('width_cm') else min(140.0, max_w_mm)
                    p.add_run().add_picture(img_path, width=Mm(min(req_w_mm, max_w_mm)))
                else:
                    p.add_run(f"[{node.get('alt') or 'Иллюстрация'}]")
                self._has_content = True

            elif ntype == 'rendered_diagram':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                self._render_external_diagram(node)
                self._has_content = True

            elif ntype == 'figure_caption':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                fig_size = self._get_style_cfg('GOST_Figure_Caption').get('size_pt', 14.0)
                p = self.doc.add_paragraph(style=self._get_doc_style(self._resolve_style_name('GOST_Figure_Caption')))
                render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=fig_size)
                self._has_content = True

            elif ntype == 'formula_block':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                self._render_formula_block(node['latex'], node.get('number'), size_pt=body_size_pt)
                self._has_content = True

            elif ntype == 'diagram_block':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                self._render_diagram_block(node['lines'], keep_with_next=bool(node.get('keep_with_next', True)))
                self._has_content = True

            elif ntype == 'code_block':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                c_lines = node['lines']
                keep_next = bool(node.get('keep_with_next', False))
                for c_idx, l in enumerate(c_lines):
                    p = self.doc.add_paragraph(style=self._get_doc_style(self._resolve_style_name('GOST_Listing_Code')))
                    if keep_next or c_idx < len(c_lines) - 1:
                        p.paragraph_format.keep_with_next = True
                    p.add_run(l if l else " ")
                self._has_content = True

            elif ntype == 'listing_caption':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                cap_size = self._get_style_cfg('GOST_Listing_Caption').get('size_pt', 14.0)
                p = self.doc.add_paragraph(style=self._get_doc_style('GOST_Listing_Caption'))
                if node.get('keep_with_next') is False:
                    p.paragraph_format.keep_with_next = False
                render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=cap_size)
                self._has_content = True

            elif ntype == 'blockquote':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                p = self.doc.add_paragraph(style=self._get_doc_style('GOST_Quote'))
                render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=13.0)
                self._has_content = True

            elif ntype == 'callout':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                p = self._add_styled_paragraph('GOST_Body')
                pPr = p._p.get_or_add_pPr()
                pPr.append(parse_xml(f'''
                    <w:pBdr {nsdecls("w")}>
                        <w:left w:val="single" w:sz="24" w:space="8" w:color="0067C0"/>
                    </w:pBdr>
                '''))
                r_kind = p.add_run(f"[{node['kind']}] ")
                r_kind.bold = True
                r_kind.font.color.rgb = RGBColor(0, 103, 192)
                render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=body_size_pt, strip_emphasis=self.strip_body_emphasis)
                self._has_content = True

            elif ntype == 'footnote_definition':
                if self.num_columns > 1:
                    self._ensure_columns(self.num_columns)
                p = self.doc.add_paragraph(style=self._get_doc_style('GOST_Note'))
                p.paragraph_format.first_line_indent = Mm(0)
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                r_num = p.add_run(f"[{node['id']}] ")
                r_num.font.superscript = True
                render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=11.0)
                self._has_content = True

        resolved_out = Path(output_path).resolve()
        resolved_out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(resolved_out))

    def _resolve_image_path(self, raw_path: str, base_dir: Optional[str] = None) -> Optional[str]:
        if not raw_path:
            return None
        from urllib.parse import unquote, urlparse
        from urllib.request import url2pathname

        candidates_raw = []
        cleaned = raw_path.strip().strip('<>').strip('"').strip("'")
        if cleaned.lower().startswith('file://'):
            parsed = urlparse(cleaned)
            file_p = url2pathname(unquote(parsed.path))
            if parsed.netloc and not file_p.startswith('\\'):
                file_p = f'//{parsed.netloc}{file_p}'
            candidates_raw.append(file_p)
        candidates_raw.append(unquote(cleaned))
        if cleaned != unquote(cleaned):
            candidates_raw.append(cleaned)

        base_dirs = []
        if base_dir:
            base_dirs.append(Path(base_dir))
        if getattr(self, 'existing_docx_path', None):
            base_dirs.append(Path(self.existing_docx_path).resolve().parent)

        for cand in candidates_raw:
            if not cand:
                continue
            p_obj = Path(cand)
            if p_obj.is_file():
                return str(p_obj)
            if not p_obj.is_absolute():
                for bdir in base_dirs:
                    joined = (bdir / cand).resolve()
                    if joined.is_file():
                        return str(joined)
        return None

    def _ensure_list_style(self):
        """Настраивает встроенный стиль List Paragraph по параметрам GOST_Body шаблона."""
        if getattr(self, '_list_style_ready', False):
            return self._list_style_obj
        body_cfg = self._get_style_cfg('GOST_Body')
        font_specs = self.cfg.get('fonts', {})
        f_key = body_cfg.get('font', 'primary')
        f_data = font_specs.get(f_key, font_specs.get('primary', {'name': 'Times New Roman'}))

        try:
            lp_style = self.doc.styles['List Paragraph']
        except KeyError:
            lp_style = self.doc.styles.add_style('List Paragraph', WD_STYLE_TYPE.PARAGRAPH)

        if not self.preserve_existing_styles or not self.is_append_mode:
            self._apply_clean_rpr_to_style_element(lp_style.element, f_data, body_cfg)
            pPr = lp_style.paragraph_format
            pPr.line_spacing = body_cfg.get('line_spacing', 1.5)
            pPr.space_before = Pt(0.0)
            pPr.space_after = Pt(0.0)
            pPr.first_line_indent = Mm(0.0)

        self._list_style_obj = lp_style
        self._list_style_ready = True
        return lp_style

    def _get_or_create_abstract_num(self, list_kind: str, is_latin_alpha: bool = False) -> int:
        if not hasattr(self, '_abstract_num_cache'):
            self._abstract_num_cache = {}
        cache_key = (list_kind, is_latin_alpha)
        if cache_key in self._abstract_num_cache:
            return self._abstract_num_cache[cache_key]

        numbering_el = self.doc.part.numbering_part.numbering_definitions._numbering
        existing_ids = []
        for el in numbering_el.findall(qn('w:abstractNum')):
            aid = el.get(qn('w:abstractNumId'))
            if aid is not None and aid.isdigit():
                existing_ids.append(int(aid))
        new_aid = (max(existing_ids) + 1) if existing_ids else 10

        lists_cfg = self.cfg.get('lists', {})
        base_indent = float(lists_cfg.get('base_indent_mm', self._get_style_cfg('GOST_Body').get('first_line_indent_mm', 12.5)))
        level_step = float(lists_cfg.get('level_step_mm', 6.0))
        hang_cfg = lists_cfg.get('hanging_indent_mm', {})

        def _get_hang(k_name: str, default_v: float) -> float:
            if isinstance(hang_cfg, (int, float)):
                return float(hang_cfg)
            if isinstance(hang_cfg, dict) and k_name in hang_cfg:
                return float(hang_cfg[k_name])
            return default_v

        body_cfg = self._get_style_cfg('GOST_Body')
        font_specs = self.cfg.get('fonts', {})
        f_data = font_specs.get(body_cfg.get('font', 'primary'), font_specs.get('primary', {'name': 'Times New Roman'}))
        f_name = escape(f_data.get('name', 'Times New Roman'))

        dash_sym = str(lists_cfg.get('dash_symbol', '–'))
        b_syms = lists_cfg.get('bullet_symbols', ['•', '◦', '▪']) or ['•']

        lvls_xml = []
        for lvl in range(6):
            if list_kind == 'gost_dash':
                num_fmt = 'bullet'
                lvl_text = escape(dash_sym)
                hang = _get_hang('dash', 4.5)
            elif list_kind == 'bullet':
                num_fmt = 'bullet'
                lvl_text = escape(b_syms[lvl % len(b_syms)])
                hang = _get_hang('bullet', 4.5)
            elif list_kind == 'task_checked':
                num_fmt = 'bullet'
                lvl_text = '☑'
                hang = _get_hang('task', 5.5)
            elif list_kind == 'task_unchecked':
                num_fmt = 'bullet'
                lvl_text = '☐'
                hang = _get_hang('task', 5.5)
            elif list_kind == 'ordered_dot':
                num_fmt = 'decimal'
                lvl_text = f'%{lvl + 1}.'
                hang = _get_hang('ordered', 6.5)
            elif list_kind == 'ordered_paren':
                num_fmt = 'decimal'
                lvl_text = f'%{lvl + 1})'
                hang = _get_hang('ordered', 6.5)
            elif list_kind == 'alpha_paren':
                num_fmt = 'lowerLetter' if is_latin_alpha else 'russianLower'
                lvl_text = f'%{lvl + 1})'
                hang = _get_hang('ordered', 6.5)
            else:
                num_fmt = 'bullet'
                lvl_text = escape(dash_sym)
                hang = _get_hang('dash', 4.5)

            left_dxa = int(round((base_indent + lvl * level_step + hang) * DXA_PER_MM))
            hang_dxa = int(round(hang * DXA_PER_MM))
            lvls_xml.append(
                f'<w:lvl w:ilvl="{lvl}">'
                f'<w:start w:val="1"/>'
                f'<w:numFmt w:val="{num_fmt}"/>'
                f'<w:lvlText w:val="{lvl_text}"/>'
                f'<w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{left_dxa}" w:hanging="{hang_dxa}"/></w:pPr>'
                f'<w:rPr><w:rFonts w:ascii="{f_name}" w:hAnsi="{f_name}" w:cs="{f_name}" w:hint="default"/></w:rPr>'
                f'</w:lvl>'
            )

        abs_xml = (
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{new_aid}">'
            f'<w:multiLevelType w:val="multilevel"/>'
            + ''.join(lvls_xml)
            + '</w:abstractNum>'
        )
        abs_el = parse_xml(abs_xml)
        first_num = numbering_el.find(qn('w:num'))
        if first_num is not None:
            first_num.addprevious(abs_el)
        else:
            numbering_el.append(abs_el)

        self._abstract_num_cache[cache_key] = new_aid
        return new_aid

    def _create_num_instance(self, abstract_num_id: int, level: int = 0, start_val: int = 1) -> int:
        numbering_el = self.doc.part.numbering_part.numbering_definitions._numbering
        existing_nids = []
        for el in numbering_el.findall(qn('w:num')):
            nid = el.get(qn('w:numId'))
            if nid is not None and nid.isdigit():
                existing_nids.append(int(nid))
        new_nid = (max(existing_nids) + 1) if existing_nids else 10

        override_xml = f'<w:lvlOverride w:ilvl="{level}"><w:startOverride w:val="{start_val}"/></w:lvlOverride>'
        num_xml = (
            f'<w:num {nsdecls("w")} w:numId="{new_nid}">'
            f'<w:abstractNumId w:val="{abstract_num_id}"/>'
            f'{override_xml}'
            f'</w:num>'
        )
        numbering_el.append(parse_xml(num_xml))
        return new_nid

    def _render_list_item(self, node: dict, body_size_pt: float):
        lp_style = self._ensure_list_style()
        p = self.doc.add_paragraph(style=lp_style)
        pPr = p._p.get_or_add_pPr()
        if pPr.find(qn('w:pStyle')) is None and hasattr(lp_style, 'style_id'):
            pPr.insert(0, parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{lp_style.style_id}"/>'))

        list_kind = node.get('list_kind', 'gost_dash')
        raw_marker = node.get('raw_marker', node.get('marker', '–'))
        level = min(5, max(0, int(node.get('level', 0) or 0)))

        is_latin_alpha = False
        effective_kind = list_kind
        start_val = 1
        if list_kind == 'task':
            effective_kind = 'task_checked' if node.get('marker') == '☑' else 'task_unchecked'
        elif list_kind in ('ordered_dot', 'ordered_paren'):
            m_d = re.match(r'^(\d+)', str(raw_marker))
            if m_d:
                start_val = int(m_d.group(1))
        elif list_kind == 'alpha_paren':
            ch = str(raw_marker)[0].lower() if raw_marker else 'а'
            is_latin_alpha = ('a' <= ch <= 'z')

        if not hasattr(self, '_active_list_nums'):
            self._active_list_nums = {}

        prev_was_list = getattr(self, '_prev_node_was_list', False)
        list_key = (effective_kind, is_latin_alpha, level)
        if (not prev_was_list) or (list_key not in self._active_list_nums) or (list_kind in ('ordered_dot', 'ordered_paren') and start_val == 1 and getattr(self, '_last_list_key', None) != list_key):
            for k_old in list(self._active_list_nums.keys()):
                if k_old[2] >= level:
                    del self._active_list_nums[k_old]
            abs_id = self._get_or_create_abstract_num(effective_kind, is_latin_alpha=is_latin_alpha)
            num_id = self._create_num_instance(abs_id, level=level, start_val=start_val)
            self._active_list_nums[list_key] = num_id
        else:
            num_id = self._active_list_nums[list_key]

        self._last_list_key = list_key
        self._prev_node_was_list = True

        for old_np in pPr.findall(qn('w:numPr')):
            pPr.remove(old_np)
        pPr.append(parse_xml(
            f'<w:numPr {nsdecls("w")}>'
            f'<w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{num_id}"/>'
            f'</w:numPr>'
        ))

        lists_cfg = self.cfg.get('lists', {})
        if lists_cfg.get('space_before_pt') is not None:
            p.paragraph_format.space_before = Pt(float(lists_cfg['space_before_pt']))
        if lists_cfg.get('space_after_pt') is not None:
            p.paragraph_format.space_after = Pt(float(lists_cfg['space_after_pt']))
        if lists_cfg.get('line_spacing') is not None:
            p.paragraph_format.line_spacing = float(lists_cfg['line_spacing'])
        if lists_cfg.get('alignment'):
            align_lut = {
                'left': WD_ALIGN_PARAGRAPH.LEFT,
                'center': WD_ALIGN_PARAGRAPH.CENTER,
                'right': WD_ALIGN_PARAGRAPH.RIGHT,
                'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            if lists_cfg['alignment'] in align_lut:
                p.alignment = align_lut[lists_cfg['alignment']]
        if lists_cfg.get('keep_together'):
            p.paragraph_format.keep_together = True

        render_inlines_to_paragraph(p, node['text'], is_heading=False, font_size_pt=body_size_pt, strip_emphasis=self.strip_body_emphasis)

    def _compile_diagram_to_image(self, lang: str, code: str, work_dir: Path) -> Optional[Path]:
        diag_cfg = self.cfg.get('diagrams', {})
        dpi = int(diag_cfg.get('dpi', 300))

        def _cmd_prefix(val):
            return list(val) if isinstance(val, (list, tuple)) else [val]

        if lang == 'mermaid':
            cli = diag_cfg.get('mermaid_cli') or shutil.which('mmdc')
            if not cli:
                return None
            in_file = work_dir / 'diagram.mmd'
            out_file = work_dir / 'diagram.png'
            in_file.write_text(code, encoding='utf-8')
            scale = max(1, round(dpi / 96))
            cmd = _cmd_prefix(cli) + ['-i', str(in_file), '-o', str(out_file), '-s', str(scale), '-b', 'white']
        elif lang == 'plantuml':
            cli = diag_cfg.get('plantuml_cli') or shutil.which('plantuml')
            jar_path = diag_cfg.get('plantuml_jar') or os.environ.get('PLANTUML_JAR')
            in_file = work_dir / 'diagram.puml'
            out_file = work_dir / 'diagram.png'
            in_file.write_text(code, encoding='utf-8')
            if cli:
                cmd = _cmd_prefix(cli) + ['-tpng', '-o', str(work_dir), str(in_file)]
            elif jar_path and os.path.exists(jar_path):
                cmd = ['java', '-jar', jar_path, '-tpng', '-o', str(work_dir), str(in_file)]
            else:
                return None
        elif lang == 'dot':
            cli = diag_cfg.get('dot_cli') or shutil.which('dot')
            if not cli:
                return None
            in_file = work_dir / 'diagram.dot'
            out_file = work_dir / 'diagram.png'
            in_file.write_text(code, encoding='utf-8')
            cmd = _cmd_prefix(cli) + ['-Tpng', f'-Gdpi={dpi}', str(in_file), '-o', str(out_file)]
        else:
            return None

        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if res.returncode == 0 and out_file.exists() and out_file.stat().st_size > 0:
                return out_file
            # Также проверяем возможный EMF-выход, если кастомный CLI сгенерировал .emf
            emf_file = work_dir / 'diagram.emf'
            if res.returncode == 0 and emf_file.exists() and emf_file.stat().st_size > 0:
                return emf_file
        except Exception:
            return None
        return None

    def _render_external_diagram(self, node: dict):
        lang = node.get('lang', 'mermaid')
        code = node.get('code', '')
        keep_next = bool(node.get('keep_with_next', False))

        with tempfile.TemporaryDirectory(prefix='gost_diag_') as tmpdir:
            img_path = self._compile_diagram_to_image(lang, code, Path(tmpdir))
            if img_path is not None:
                p = self.doc.add_paragraph(style=self._get_doc_style(self._resolve_style_name('GOST_Figure')))
                if keep_next:
                    p.paragraph_format.keep_with_next = True
                max_w_mm = self._get_current_text_width_mm() * 0.95
                p.add_run().add_picture(str(img_path), width=Mm(min(140.0, max_w_mm)))
                return

        # Фолбэк, если внешний CLI-рендерер не установлен в системе
        self._render_diagram_block(node.get('lines', []), keep_with_next=keep_next)

    def _render_formula_block(self, latex_str: str, eq_num: Optional[str], size_pt: float = 14.0):
        if eq_num:
            col_w_mm = self._get_current_text_width_mm()
            p = self._add_styled_paragraph('GOST_Body')
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Mm(0)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.tab_stops.add_tab_stop(Mm(col_w_mm / 2.0), WD_TAB_ALIGNMENT.CENTER)
            p.paragraph_format.tab_stops.add_tab_stop(Mm(col_w_mm), WD_TAB_ALIGNMENT.RIGHT)

            p.add_run('\t')
            p._p.append(parse_xml(latex_to_omml(latex_str, size_pt=size_pt)))
            p.add_run('\t')
            p.add_run(eq_num)
        else:
            p = self._add_styled_paragraph('GOST_Body')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Mm(0)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            m_para = parse_xml(f'''
                <m:oMathPara {nsdecls("m")}>
                    {latex_to_omml(latex_str, size_pt=size_pt)}
                </m:oMathPara>
            ''')
            p._p.append(m_para)

    def _render_diagram_block(self, lines: List[str], keep_with_next: bool = True):
        if not lines:
            return
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblBorders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                <w:insideH w:val="none"/><w:insideV w:val="none"/>
            </w:tblBorders>
        ''')
        t._tbl.tblPr.append(tblBorders)

        trPr = t.rows[0]._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        cell = t.rows[0].cells[0]
        p = cell.paragraphs[0]
        p.style = self._get_doc_style('GOST_Listing_Code')
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Mm(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        if keep_with_next:
            p.paragraph_format.keep_with_next = True

        for idx, l in enumerate(lines):
            if idx > 0:
                p.add_run().add_break()
            run = p.add_run(l if l else " ")
            run.font.name = 'Courier New'
            run.font.size = Pt(9.5)

    def _get_table_padding_mm(self) -> Dict[str, float]:
        tbl_cfg = self.cfg.get('tables', {})
        pad = tbl_cfg.get('cell_padding_mm', {})
        if isinstance(pad, (int, float)):
            v = float(pad)
            return {'top': v, 'bottom': v, 'left': v, 'right': v}
        return {
            'top': float(pad.get('top', 1.0)),
            'bottom': float(pad.get('bottom', 1.0)),
            'left': float(pad.get('left', 1.8)),
            'right': float(pad.get('right', 1.8)),
        }

    def _get_table_h_padding_mm(self) -> float:
        p = self._get_table_padding_mm()
        return p['left'] + p['right']

    def _render_table(self, node: dict):
        header = node['header']
        rows = node['rows']
        alignments = node.get('alignments', [])
        cols = len(header)
        if cols == 0:
            return

        tbl_cfg = self.cfg.get('tables', {})
        border_style = tbl_cfg.get('border_style', 'full_grid')
        header_double_border = bool(tbl_cfg.get('header_double_border', True))
        header_shading = tbl_cfg.get('header_shading', None)
        repeat_header = bool(tbl_cfg.get('repeat_header', True))
        cant_split = bool(tbl_cfg.get('cant_split', True))
        pad_mm = self._get_table_padding_mm()

        cell_size_pt = self._get_style_cfg('GOST_Table_Cell').get('size_pt', 12.0)

        t = self.doc.add_table(rows=len(rows) + 1, cols=cols)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False

        total_w_mm = self._get_current_text_width_mm()
        col_widths_mm = compute_smart_column_widths_mm(
            header,
            rows,
            total_w_mm,
            cell_size_pt=cell_size_pt,
            h_padding_mm=pad_mm['left'] + pad_mm['right']
        )

        if border_style == 'gost_open':
            tblBorders = parse_xml(f'''
                <w:tblBorders {nsdecls("w")}>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:left w:val="none"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:right w:val="none"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                </w:tblBorders>
            ''')
        else:
            tblBorders = parse_xml(f'''
                <w:tblBorders {nsdecls("w")}>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                </w:tblBorders>
            ''')
        t._tbl.tblPr.append(tblBorders)

        mar_top = int(round(pad_mm['top'] * DXA_PER_MM))
        mar_bot = int(round(pad_mm['bottom'] * DXA_PER_MM))
        mar_left = int(round(pad_mm['left'] * DXA_PER_MM))
        mar_right = int(round(pad_mm['right'] * DXA_PER_MM))
        tblCellMar = parse_xml(f'''
            <w:tblCellMar {nsdecls("w")}>
                <w:top w:w="{mar_top}" w:type="dxa"/>
                <w:left w:w="{mar_left}" w:type="dxa"/>
                <w:bottom w:w="{mar_bot}" w:type="dxa"/>
                <w:right w:w="{mar_right}" w:type="dxa"/>
            </w:tblCellMar>
        ''')
        t._tbl.tblPr.append(tblCellMar)

        # Шапка таблицы
        hdr_row = t.rows[0]
        hdr_trPr = hdr_row._tr.get_or_add_trPr()
        if repeat_header:
            hdr_trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        if cant_split:
            hdr_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        for idx, title in enumerate(header):
            cell = hdr_row.cells[idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.style = self._get_doc_style('GOST_Table_Cell')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next = True
            render_inlines_to_paragraph(p, title, is_heading=True, font_size_pt=cell_size_pt)

            tcPr = cell._tc.get_or_add_tcPr()
            if header_shading:
                fill_hex = str(header_shading).lstrip('#')
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{escape(fill_hex)}"/>'))

            if header_double_border:
                tcBorders = parse_xml(f'''
                    <w:tcBorders {nsdecls("w")}>
                        <w:bottom w:val="double" w:sz="12" w:space="0" w:color="auto"/>
                    </w:tcBorders>
                ''')
                tcPr.append(tcBorders)

        # Строки данных
        for r_idx, r_data in enumerate(rows):
            row = t.rows[r_idx + 1]
            trPr = row._tr.get_or_add_trPr()
            if cant_split:
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            for c_idx in range(cols):
                val = r_data[c_idx] if c_idx < len(r_data) else ""
                cell = row.cells[c_idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                p.style = self._get_doc_style('GOST_Table_Cell')

                col_align = alignments[c_idx] if c_idx < len(alignments) else 'auto'
                if col_align == 'center':
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif col_align == 'right':
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif col_align == 'left':
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    is_centered = _is_centered_table_cell(val)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_centered else WD_ALIGN_PARAGRAPH.LEFT

                render_inlines_to_paragraph(p, val, is_heading=False, font_size_pt=cell_size_pt)

        col_dxa = [int(round(w * DXA_PER_MM)) for w in col_widths_mm]
        total_dxa = sum(col_dxa)

        tblPr = t._tbl.tblPr
        for old_el in tblPr.findall(qn('w:tblW')) + tblPr.findall(qn('w:tblLayout')):
            tblPr.remove(old_el)
        tblPr.insert(0, parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_dxa}" w:type="dxa"/>'))
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

        tblGrid = t._tbl.tblGrid
        if tblGrid is not None:
            tblGrid.clear()
            for w_dxa in col_dxa:
                tblGrid.append(parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w_dxa}"/>'))

        for r in t.rows:
            trPr = r._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="454" w:hRule="atLeast"/>'))
            for c_i, cell in enumerate(r.cells):
                if c_i < len(col_dxa):
                    tcPr = cell._tc.get_or_add_tcPr()
                    for old_tcw in tcPr.findall(qn('w:tcW')):
                        tcPr.remove(old_tcw)
                    tcPr.insert(0, parse_xml(f'<w:tcW {nsdecls("w")} w:w="{col_dxa[c_i]}" w:type="dxa"/>'))


# ==============================================================================
# ФУНКЦИИ APPEND, ИЗМЕНЕНИЯ СТАНДАРТНЫХ СТИЛЕЙ И СКЛЕЙКИ DOCX
# ==============================================================================

def _apply_style_definition(style, s_def: dict, font_specs: dict):
    align_lut = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    f_key = s_def.get('font', 'primary')
    if isinstance(f_key, dict):
        f_data = f_key
    elif f_key in font_specs:
        f_data = font_specs[f_key]
    elif s_def.get('font'):
        # Прямое имя шрифта, например "Arial" или "Times New Roman"
        f_data = {'name': str(f_key)}
    else:
        f_data = font_specs.get('primary', {'name': 'Times New Roman'})

    if 'font' in s_def or 'font_name' in s_def:
        fname = s_def.get('font_name') or f_data['name']
        style.font.name = fname
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        for theme_attr in ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']:
            attr_qn = qn(f'w:{theme_attr}')
            if attr_qn in rFonts.attrib:
                del rFonts.attrib[attr_qn]
        for attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
            rFonts.set(qn(f'w:{attr}'), f_data.get(attr, fname))

    if 'size_pt' in s_def:
        style.font.size = Pt(float(s_def['size_pt']))
    if 'bold' in s_def:
        style.font.bold = bool(s_def['bold'])
    if 'italic' in s_def:
        style.font.italic = bool(s_def['italic'])

    pPr = style.paragraph_format
    if 'alignment' in s_def:
        pPr.alignment = align_lut.get(s_def['alignment'], WD_ALIGN_PARAGRAPH.JUSTIFY)
    if 'line_spacing' in s_def:
        pPr.line_spacing = float(s_def['line_spacing'])
    if 'first_line_indent_mm' in s_def:
        pPr.first_line_indent = Mm(float(s_def['first_line_indent_mm']))
    if 'left_indent_mm' in s_def:
        pPr.left_indent = Mm(float(s_def['left_indent_mm']))
    if 'space_before_pt' in s_def:
        pPr.space_before = Pt(float(s_def['space_before_pt']))
    if 'space_after_pt' in s_def:
        pPr.space_after = Pt(float(s_def['space_after_pt']))
    if 'keep_with_next' in s_def:
        pPr.keep_with_next = bool(s_def['keep_with_next'])
    if 'page_break_before' in s_def:
        pPr.page_break_before = bool(s_def['page_break_before'])


def update_docx_styles(
    existing_docx_path: str,
    style_updates: dict,
    output_path: Optional[str] = None,
    fonts_cfg: Optional[dict] = None,
) -> str:
    """
    Обновляет стандартные и пользовательские стили существующего DOCX (например, 'Обычный' / 'Normal',
    'Заголовок 1' / 'Heading 1'), полностью сохраняя исходный текст, таблицы и локальное форматирование.
    """
    doc = Document(existing_docx_path)
    if 'styles' in style_updates and isinstance(style_updates['styles'], dict):
        styles_map = style_updates['styles']
        font_specs = fonts_cfg or style_updates.get('fonts', {})
    elif any(k in style_updates for k in ('page', 'strip_body_emphasis', 'fonts', 'tables', 'lists')):
        styles_map = {}
        font_specs = fonts_cfg or style_updates.get('fonts', {})
    else:
        styles_map = {k: v for k, v in style_updates.items() if isinstance(v, dict)}
        font_specs = fonts_cfg or {}

    # Индексируем существующие стили документа
    existing_map = {}
    for s in doc.styles:
        if s.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        if s.name:
            existing_map[s.name] = s
            existing_map[s.name.lower()] = s
            canon = STYLE_CANONICAL_MAP.get(s.name.strip().lower())
            if canon:
                existing_map.setdefault(canon, s)
                existing_map.setdefault(canon.lower(), s)
        if s.style_id:
            existing_map[s.style_id] = s
            existing_map[s.style_id.lower()] = s
        aliases_el = s.element.find(qn('w:aliases'))
        if aliases_el is not None:
            for a in (aliases_el.get(qn('w:val')) or '').split(','):
                a_c = a.strip()
                if a_c:
                    existing_map.setdefault(a_c, s)
                    existing_map.setdefault(a_c.lower(), s)

    for raw_name, s_def in styles_map.items():
        canon_name = STYLE_CANONICAL_MAP.get(raw_name.strip().lower(), raw_name)
        target_style = None
        for cand in (raw_name, raw_name.lower(), canon_name, canon_name.lower()):
            if cand in existing_map:
                target_style = existing_map[cand]
                break
        if target_style is None:
            try:
                target_style = doc.styles[canon_name]
            except KeyError:
                target_style = doc.styles.add_style(canon_name, WD_STYLE_TYPE.PARAGRAPH)
        _apply_style_definition(target_style, s_def, font_specs)

    if isinstance(style_updates, dict) and style_updates.get('strip_body_emphasis'):
        for p_obj in doc.paragraphs:
            st_name = p_obj.style.name if p_obj.style else 'Normal'
            if st_name in ('Normal', 'List Paragraph', 'Body Text', 'Обычный'):
                for r_obj in p_obj.runs:
                    r_obj.bold = False
                    r_obj.italic = False

    if isinstance(style_updates, dict) and isinstance(style_updates.get('page'), dict):
        cols_n = int(style_updates['page'].get('columns', 1) or 1)
        col_sp_mm = float(style_updates['page'].get('column_spacing_mm', 5.0) or 5.0)
        if cols_n > 1:
            for sec in doc.sections:
                cols_el = sec._sectPr.find(qn('w:cols'))
                if cols_el is None:
                    cols_el = OxmlElement('w:cols')
                    sec._sectPr.append(cols_el)
                cols_el.set(qn('w:num'), str(cols_n))
                cols_el.set(qn('w:space'), str(int(round(col_sp_mm * 56.692913))))

    out_p = Path(output_path or existing_docx_path).resolve()
    out_p.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_p))
    return str(out_p)


def append_markdown_to_docx(
    existing_docx_path: str,
    md_text: str,
    template_cfg: dict,
    output_path: Optional[str] = None,
    update_existing_styles: bool = False,
    add_page_break: bool = False,
    base_dir: Optional[str] = None,
) -> str:
    """
    Добавляет (append) отформатированный Markdown/LaTeX-текст в конец существующего DOCX.
    По умолчанию использует стили, уже присутствующие в документе, и сохраняет весь исходный текст и форматирование.
    """
    parser = MarkdownASTParser(template_cfg.get('structural_titles', []))
    ast_nodes = parser.parse(md_text)

    formatter = GostDocxFormatter(
        template_cfg,
        existing_docx_path=existing_docx_path,
        preserve_existing_styles=not update_existing_styles,
        update_existing_styles=update_existing_styles,
    )
    if add_page_break and ast_nodes:
        formatter.doc.add_page_break()

    out_target = output_path or existing_docx_path
    formatter.build_document(ast_nodes, out_target, base_dir=base_dir)
    return str(Path(out_target).resolve())


def merge_docx_files(
    base_docx_path: str,
    append_docx_paths: List[str],
    output_path: Optional[str] = None,
    add_page_break: bool = False,
) -> str:
    """
    Склеивает (append) один или несколько DOCX-файлов в конец базового документа,
    сохраняя исходный текст, форматирование, стили целевого документа, таблицы, изображения и гиперссылки.
    """
    dst_doc = Document(base_docx_path)
    dst_body = dst_doc.element.body

    existing_style_ids = {s.style_id for s in dst_doc.styles if s.style_id}
    existing_style_by_name = {s.name.lower(): s.style_id for s in dst_doc.styles if s.name and s.style_id}

    rel_attrs = [
        qn('r:embed'),
        qn('r:link'),
        qn('r:id'),
    ]

    for src_path in append_docx_paths:
        src_doc = Document(src_path)
        src_body = src_doc.element.body

        # 1. Копируем недостающие стили из src_doc (не перезаписывая существующие стили dst_doc)
        # Если стиль совпадает по имени с уже существующим в dst_doc (но имеет другой styleId), запоминаем ремаппинг
        dst_styles_el = dst_doc.styles.element
        style_id_remap: Dict[str, str] = {}
        for s in src_doc.styles:
            if s.style_id in existing_style_ids:
                continue
            if s.name and s.name.lower() in existing_style_by_name:
                if s.style_id:
                    style_id_remap[s.style_id] = existing_style_by_name[s.name.lower()]
                continue
            dst_styles_el.append(copy.deepcopy(s.element))
            if s.style_id:
                existing_style_ids.add(s.style_id)
            if s.name and s.style_id:
                existing_style_by_name[s.name.lower()] = s.style_id

        # 1.5. Переносим определения списков (numbering.xml) с уникальными abstractNumId и numId
        num_id_remap: Dict[str, str] = {}
        try:
            src_num_part = src_doc.part.numbering_part
        except Exception:
            src_num_part = None
        if src_num_part is not None:
            src_num_el = src_num_part.numbering_definitions._numbering
            dst_num_el = dst_doc.part.numbering_part.numbering_definitions._numbering
            dst_aids = [int(el.get(qn('w:abstractNumId'))) for el in dst_num_el.findall(qn('w:abstractNum')) if (el.get(qn('w:abstractNumId')) or '').isdigit()]
            dst_nids = [int(el.get(qn('w:numId'))) for el in dst_num_el.findall(qn('w:num')) if (el.get(qn('w:numId')) or '').isdigit()]
            next_aid = (max(dst_aids) + 1) if dst_aids else 1
            next_nid = (max(dst_nids) + 1) if dst_nids else 1
            aid_remap: Dict[str, str] = {}
            first_num_el = dst_num_el.find(qn('w:num'))
            for abs_el in src_num_el.findall(qn('w:abstractNum')):
                old_aid = abs_el.get(qn('w:abstractNumId'))
                new_abs = copy.deepcopy(abs_el)
                new_aid_str = str(next_aid)
                next_aid += 1
                new_abs.set(qn('w:abstractNumId'), new_aid_str)
                if old_aid is not None:
                    aid_remap[old_aid] = new_aid_str
                if first_num_el is not None:
                    first_num_el.addprevious(new_abs)
                else:
                    dst_num_el.append(new_abs)
            for num_el in src_num_el.findall(qn('w:num')):
                old_nid = num_el.get(qn('w:numId'))
                new_num = copy.deepcopy(num_el)
                new_nid_str = str(next_nid)
                next_nid += 1
                new_num.set(qn('w:numId'), new_nid_str)
                abs_ref = new_num.find(qn('w:abstractNumId'))
                if abs_ref is not None:
                    old_ref = abs_ref.get(qn('w:val'))
                    if old_ref in aid_remap:
                        abs_ref.set(qn('w:val'), aid_remap[old_ref])
                dst_num_el.append(new_num)
                if old_nid is not None:
                    num_id_remap[old_nid] = new_nid_str

        # 2. Вставляем разрыв страницы при необходимости
        dst_sect_pr = dst_body.find(qn('w:sectPr'))
        if add_page_break:
            pb_p = OxmlElement('w:p')
            pb_r = OxmlElement('w:r')
            pb_br = OxmlElement('w:br')
            pb_br.set(qn('w:type'), 'page')
            pb_r.append(pb_br)
            pb_p.append(pb_r)
            if dst_sect_pr is not None:
                dst_sect_pr.addprevious(pb_p)
            else:
                dst_body.append(pb_p)

        # 3. Копируем и переназначаем relationship ID (изображения с уникальными partname, гиперссылки)
        from docx.opc.part import Part
        from docx.opc.packuri import PackURI
        existing_partnames = {str(pt.partname) for pt in dst_doc.part.package.iter_parts()}
        rid_map: Dict[str, str] = {}
        media_counter = 1
        for old_rid, rel in src_doc.part.rels.items():
            try:
                if rel.is_external:
                    new_rid = dst_doc.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                    rid_map[old_rid] = new_rid
                elif 'image' in rel.reltype:
                    tpart = rel.target_part
                    pname_str = str(tpart.partname)
                    if pname_str in existing_partnames:
                        ext = Path(pname_str).suffix or '.png'
                        while f'/word/media/merged_{media_counter}{ext}' in existing_partnames:
                            media_counter += 1
                        new_pname = PackURI(f'/word/media/merged_{media_counter}{ext}')
                        media_counter += 1
                        tpart = Part(new_pname, tpart.content_type, tpart.blob, dst_doc.part.package)
                    existing_partnames.add(str(tpart.partname))
                    new_rid = dst_doc.part.relate_to(tpart, rel.reltype)
                    rid_map[old_rid] = new_rid
                elif 'hyperlink' in rel.reltype:
                    new_rid = dst_doc.part.relate_to(rel.target_part, rel.reltype)
                    rid_map[old_rid] = new_rid
            except Exception:
                pass

        # 4. Копируем элементы тела (параграфы и таблицы), исключая финальный sectPr
        for child in list(src_body):
            if child.tag == qn('w:sectPr'):
                continue
            cloned = copy.deepcopy(child)
            if rid_map or style_id_remap or num_id_remap:
                for elem in cloned.iter():
                    if rid_map:
                        for attr in rel_attrs:
                            old_val = elem.get(attr)
                            if old_val and old_val in rid_map:
                                elem.set(attr, rid_map[old_val])
                    if style_id_remap and elem.tag in (qn('w:pStyle'), qn('w:rStyle'), qn('w:tblStyle')):
                        old_st = elem.get(qn('w:val'))
                        if old_st and old_st in style_id_remap:
                            elem.set(qn('w:val'), style_id_remap[old_st])
                    if num_id_remap and elem.tag == qn('w:numId'):
                        old_nid = elem.get(qn('w:val'))
                        if old_nid and old_nid in num_id_remap:
                            elem.set(qn('w:val'), num_id_remap[old_nid])
            if dst_sect_pr is not None:
                dst_sect_pr.addprevious(cloned)
            else:
                dst_body.append(cloned)

    out_p = Path(output_path or base_docx_path).resolve()
    out_p.parent.mkdir(parents=True, exist_ok=True)
    dst_doc.save(str(out_p))
    return str(out_p)
