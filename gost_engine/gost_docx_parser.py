#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gost_docx_parser.py — Высокоточный синтаксический анализатор Markdown для стандартов ГОСТ / ДГТУ.
Отвечает за:
- Извлечение метаданных YAML Frontmatter
- Двухфазную изоляцию формул, кода и таблиц (Vault)
- Контекстную типографическую нормализацию (Smart Typography: тире, неразрывные пробелы, единицы, инициалы, кавычки)
- Построение абстрактного синтаксического дерева (AST) с сохранением маркеров списков и диаграмм
- Автоматическое связывание соседних блоков (подписи таблиц, рисунков, листингов)
- Парсинг инлайн-разметки (жирный, курсив, зачеркивание, подчеркивание, индексы, сноски, подсветка)
"""

from __future__ import annotations

import re
from typing import List, Dict, Tuple, Any, Optional

NBSP = '\u00A0'

def sanitize_and_repair_latex_escapes(text: str) -> str:
    """Восстанавливает поврежденные управляющие последовательности LaTeX
    (\b -> \x08, \f -> \x0c, \a -> \x07, \v -> \x0b, \t -> TAB, \r -> CR),
    возникающие при записи неэкранированных строк в Python/JSON,
    и удаляет все запрещенные спецификацией XML 1.0 управляющие байты.
    """
    if not text:
        return text
    # 1. Восстановление управляющих байтов, съевших слэш в командах LaTeX:
    # \x07 (BEL, \a) -> \a (\approx, \alpha, \ast, \angle, \arccos, \arcsin, \arctan)
    text = re.sub(r'\x07(?=[a-zA-Z])', r'\\a', text)
    # \x08 (BS, \b) -> \b (\beta, \begin, \bar, \mathbf, \bullet, \bigcup, \bigcap, \binom)
    text = re.sub(r'\x08(?=[a-zA-Z])', r'\\b', text)
    # \x0c (FF, \f) -> \f (\frac, \forall, \flat)
    text = re.sub(r'\x0c(?=[a-zA-Z])', r'\\f', text)
    # \x0b (VT, \v) -> \v (\varepsilon, \varphi, \vartheta, \varrho, \varsigma, \varkappa, \vec, \vdash)
    text = re.sub(r'\x0b(?=[a-zA-Z])', r'\\v', text)
    # \t (TAB) перед известными командами LaTeX на 't' (\text, \times, \theta, \tau, \to, \top, \triangle, \tg, \tan, \tilde, \tag)
    text = re.sub(r'\t(?=(?:ext|imes|heta|au|o\b|op\b|riangle|g\b|anh|an\b|ilde|ag\b))', r'\\t', text)
    # \r (CR без \n) перед известными командами LaTeX на 'r' (\right, \rho, \rangle, \rceil, \rfloor, \rvert)
    text = re.sub(r'\r(?!\n)(?=(?:ight|ho\b|angle|ceil|floor|vert))', r'\\r', text)

    # 2. Удаление любых оставшихся невалидных для XML 1.0 управляющих символов (кроме \t, \n, \r)
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ufffe\uffff]', '', text)
    return text


UNITS = [
    # Время
    r'с', r'сек', r'мс', r'мкс', r'нс', r'пс', r'мин', r'ч', r'сут', r'суток',
    r'дн', r'дней', r'дня', r'день', r'нед', r'недель', r'мес', r'месяцев', r'г', r'лет', r'года',
    r's', r'sec', r'ms', r'us', r'ns', r'min', r'h', r'hr', r'd', r'day', r'days',

    # Геометрия и пространство
    r'нм', r'мкм', r'мм', r'см', r'дм', r'м', r'км',
    r'мм[2²]', r'см[2²]', r'м[2²]', r'км[2²]', r'га',
    r'мм[3³]', r'см[3³]', r'м[3³]', r'мл', r'л',
    r'nm', r'um', r'mm', r'cm', r'dm', r'm', r'km', r'l', r'ml',
    r'рад', r'мрад', r'град', r'об/мин', r'об/с', r'рад/с',

    # Типографика и полиграфия
    r'пт', r'pt', r'px', r'dpi', r'ppi', r'ед\.', r'экз\.', r'чел\.', r'шт\.',

    # Масса и количество вещества
    r'мкг', r'мг', r'г', r'кг', r'ц', r'т', r'моль', r'ммоль',
    r'mcg', r'mg', r'g', r'kg', r't', r'mol', r'mmol',

    # Механика, гидродинамика, термодинамика
    r'Н', r'кН', r'МН', r'Па', r'кПа', r'МПа', r'ГПа', r'бар', r'мбар', r'атм',
    r'Дж', r'кДж', r'МДж', r'кал', r'ккал', r'Вт', r'мВт', r'кВт', r'МВт', r'ГВт', r'л\.с\.',
    r'Вт·ч', r'кВт·ч', r'А·ч', r'мА·ч',
    r'N', r'kN', r'Pa', r'kPa', r'MPa', r'bar', r'J', r'kJ', r'cal', r'kcal', r'W', r'kW', r'MW',

    # Электродинамика и физика
    r'А', r'мА', r'мкА', r'нА', r'пА', r'кА',
    r'В', r'мВ', r'мкВ', r'нВ', r'кВ', r'МВ',
    r'В/А', r'В/В', r'В/мкс', r'В/мс', r'В/с', r'А/В',
    r'Ом', r'мОм', r'мкОм', r'кОм', r'МОм', r'ГОм',
    r'Ф', r'мФ', r'мкФ', r'нФ', r'пФ', r'Гн', r'мГн', r'мкГн', r'нГн',
    r'Тл', r'мТл', r'мкТл', r'Гц', r'кГц', r'МГц', r'ГГц',
    r'лм', r'лк', r'Кд', r'Кл', r'Вб', r'дБ', r'дБм',
    r'A', r'mA', r'uA', r'nA', r'V', r'mV', r'uV', r'nV', r'kV',
    r'Ohm', r'F', r'uF', r'nF', r'pF', r'H', r'mH', r'uH', r'Hz', r'kHz', r'MHz', r'GHz', r'dB',

    # Температура
    r'°C', r'°F', r'К', r'K',

    # Цифровые данные
    r'бит', r'байт', r'Кб', r'Кбайт', r'Мб', r'Мбайт', r'Гб', r'Гбайт', r'Тб', r'Тбайт',
    r'КБ', r'МБ', r'ГБ', r'ТБ', r'бит/с', r'Кбит/с', r'Мбит/с', r'Гбит/с', r'байт/с',
    r'bit', r'byte', r'Kb', r'Mb', r'Gb', r'Tb', r'KB', r'MB', r'GB', r'TB',
    r'bps', r'Kbps', r'Mbps', r'Gbps',

    # Валюты и финансовые показатели
    r'руб\.', r'руб', r'коп\.', r'коп', r'тыс\.\s*руб\.', r'млн\s*руб\.', r'млрд\s*руб\.',
    r'₽', r'\$', r'€', r'¥'
]

UNITS_SORTED = sorted(UNITS, key=len, reverse=True)
UNITS_REGEX_PATTERN = (
    r'((?:\x00MATH\d+\x00|\d+(?:[\.,]\d+)?(?:\s*(?:\.\.\.|…|[-–—±])\s*\d+(?:[\.,]\d+)?)?))'
    r'[ \t\u00A0]+'
    r'(' + '|'.join(UNITS_SORTED) + r')(?![а-яА-ЯёЁa-zA-Z0-9])'
)
UNITS_COMPILED_RE = re.compile(UNITS_REGEX_PATTERN)

# Предлоги и союзы, привязываемые к последующему слову неразрывным пробелом
PREPOSITIONS_AND_CONJUNCTIONS = (
    r'в|во|к|ко|с|со|у|о|об|обо|от|до|из|за|на|по|под|над|при|про|без|для|через|'
    r'и|а|но|да|или|ни|не'
)
PREP_CONJ_RE = re.compile(
    rf'(^|[\s\(\[«„—–])({PREPOSITIONS_AND_CONJUNCTIONS})[ \t]+(?=[а-яА-ЯёЁa-zA-Z0-9«„\(\[\x00])',
    re.IGNORECASE
)

# Сокращения, после которых ставится неразрывный пробел
ABBREV_AFTER_RE = re.compile(
    r'(?i)(?<![а-яА-ЯёЁa-zA-Z])'
    r'(рис\.|табл\.|стр\.|с\.|гл\.|п\.|пп\.|разд\.|см\.|ср\.|г\.|гг\.|'
    r'рисун(?:ок|ка|ку|ком|ке|ки|ков|кам|ками|ках)|'
    r'таблиц(?:а|ы|е|у|ей|ею|ах|ам|ами))[ \t\u00A0]+'
)

# Частые составные сокращения (т. е., т. д., т. п., и др., и пр.)
COMPOUND_ABBREV_RULES = [
    (re.compile(r'(?i)\bт\.[ \t]*е\.'), 'т.' + NBSP + 'е.'),
    (re.compile(r'(?i)\bт\.[ \t]*д\.'), 'т.' + NBSP + 'д.'),
    (re.compile(r'(?i)\bт\.[ \t]*п\.'), 'т.' + NBSP + 'п.'),
    (re.compile(r'(?i)\bи[ \t]+др\.'), 'и' + NBSP + 'др.'),
    (re.compile(r'(?i)\bи[ \t]+пр\.'), 'и' + NBSP + 'пр.'),
]

# Инициалы: "А. М. Смирнов", "А.М. Смирнов", "Смирнов А. М."
INITIALS_DOUBLE_BEFORE_SURNAME_RE = re.compile(
    r'\b([А-ЯЁA-Z]\.)[ \t\u00A0]*([А-ЯЁA-Z]\.)[ \t\u00A0]+([А-ЯЁA-Z][а-яёa-z]{1,})'
)
INITIALS_SINGLE_BEFORE_SURNAME_RE = re.compile(
    r'(?<![А-ЯЁA-Z]\.)\b([А-ЯЁA-Z]\.)[ \t\u00A0]+([А-ЯЁA-Z][а-яёa-z]{1,})'
)
SURNAME_BEFORE_INITIALS_RE = re.compile(
    r'\b([А-ЯЁA-Z][а-яёa-z]{1,})[ \t\u00A0]+([А-ЯЁA-Z]\.)[ \t\u00A0]*([А-ЯЁA-Z]\.)?'
)


def typography_cleanup(text: str) -> str:
    """
    Контекстная типографика (Smart Typography):
    - неразрывные пробелы после предлогов, союзов, инициалов, сокращений и перед единицами измерения;
    - короткое тире (–) в числовых интервалах (10–20);
    - длинное тире (—) в предложениях;
    - парные русские кавычки-елочки («...»).
    Вызывается только для текстовых узлов (блоки кода, инлайн-код и формулы защищены).
    """
    if not text:
        return text

    # 1. Дефисы и двойные дефисы между словами -> длинное тире (—) с неразрывным пробелом слева
    text = re.sub(r'(?<=\S)[ \t\u00A0]+(?:--?|[–—])[ \t\u00A0]+(?=\S)', NBSP + '— ', text)
    # Тире в начале строки (диалог / перечисление в предложении)
    text = re.sub(r'^--[ \t]+', '—' + NBSP, text)

    # 2. Короткое тире (–) в числовых интервалах (10-20 -> 10–20)
    text = re.sub(r'(\d)[ \t]*[-–—][ \t]*(\d)', r'\1–\2', text)

    # 3. Неразрывные пробелы в инициалах (А. М. Смирнов -> А.\u00A0М.\u00A0Смирнов)
    text = INITIALS_DOUBLE_BEFORE_SURNAME_RE.sub(rf'\1{NBSP}\2{NBSP}\3', text)
    text = INITIALS_SINGLE_BEFORE_SURNAME_RE.sub(rf'\1{NBSP}\2', text)
    def _surname_init_repl(m):
        if m.group(3):
            return f"{m.group(1)}{NBSP}{m.group(2)}{NBSP}{m.group(3)}"
        return f"{m.group(1)}{NBSP}{m.group(2)}"
    text = SURNAME_BEFORE_INITIALS_RE.sub(_surname_init_repl, text)

    # 4. Неразрывные пробелы после сокращений (рис. 1, табл. 2, т. е., т. д., т. п.)
    text = ABBREV_AFTER_RE.sub(r'\1' + NBSP, text)
    for pat, repl in COMPOUND_ABBREV_RULES:
        text = pat.sub(repl, text)

    # 5. Неразрывный пробел между числом (или формулой) и единицей измерения
    text = UNITS_COMPILED_RE.sub(r'\1' + NBSP + r'\2', text)

    # 6. Неразрывный пробел перед знаком процента/промилле
    text = re.sub(r'(\d+|\x00MATH\d+\x00)[ \t\u00A0]+([%‰])', r'\1' + NBSP + r'\2', text)

    # 7. Неразрывные пробелы после коротких предлогов и союзов (в два прохода для пар подряд: "и в доме")
    for _ in range(2):
        text = PREP_CONJ_RE.sub(rf'\1\2{NBSP}', text)

    # 8. Замена прямых машинописных кавычек ("...") на парные русские кавычки «...»
    text = re.sub(r'(^|[\s\(\[\{«„])"([^"]+?)"', r'\1«\2»', text)

    return text


MATH_AND_CODE_RE = re.compile(
    r'\$\$([\s\S]+?)\$\$'
    r'|\$((?:[^\$\\]|\\.)+)\$'
    r'|\\\(([\s\S]+?)\\\)'
    r'|\\\[([\s\S]+?)\\\]'
    r'|`([^`]+)`'
)

STYLE_SPAN_RE = re.compile(
    r'(\*{3}[\s\S]+?\*{3}|_{3}[\s\S]+?_{3})'
    r'|(\*{2}[\s\S]+?\*{2}|_{2}[\s\S]+?_{2})'
    r'|(\*[^*\n]+\*|(?<!\w)_[^_\n]+_(?!\w))'
    r'|(~~[\s\S]+?~~)'
    r'|(<u>[\s\S]+?</u>)'
    r'|(==[\s\S]+?==|<mark>[\s\S]+?</mark>)'
    r'|(<kbd>[\s\S]+?</kbd>)'
    r'|(<sub>[\s\S]+?</sub>|~[^\s~]+~)'
    r'|(<sup>[\s\S]+?</sup>|\^[^\s\^]+\^)'
    r'|(\[\^([a-zA-Z0-9_\-]+)\])'
    r'|(\[([^\]]+)\]\(([^)]+)\))'
)

PLACEHOLDER_SPLIT_RE = re.compile(r'(\x00(?:MATH|CODE)\d+\x00)')


def tokenize_inlines(text: str) -> list:
    if not text:
        return []

    vault: Dict[str, Tuple[str, str]] = {}
    counter = 0

    def _stash(m):
        nonlocal counter
        idx = counter
        counter += 1
        if m.group(5) is not None:
            key = f"\x00CODE{idx}\x00"
            vault[key] = ('code', m.group(5))
        else:
            math_val = m.group(1) or m.group(2) or m.group(3) or m.group(4) or ""
            key = f"\x00MATH{idx}\x00"
            vault[key] = ('math', math_val.strip())
        return key

    stashed_text = MATH_AND_CODE_RE.sub(_stash, text)
    cleaned_text = typography_cleanup(stashed_text)

    styled_segments = []
    pos = 0

    for m in STYLE_SPAN_RE.finditer(cleaned_text):
        if m.start() > pos:
            styled_segments.append({
                'text': cleaned_text[pos:m.start()],
                'bold': False, 'italic': False, 'strike': False,
                'underline': False, 'sub': False, 'sup': False,
                'highlight': False, 'kbd': False, 'footnote': None, 'link': None
            })

        matched = m.group(0)
        seg = {
            'text': matched,
            'bold': False, 'italic': False, 'strike': False,
            'underline': False, 'sub': False, 'sup': False,
            'highlight': False, 'kbd': False, 'footnote': None, 'link': None
        }

        if (matched.startswith('***') and matched.endswith('***')) or (matched.startswith('___') and matched.endswith('___')):
            seg['text'] = matched[3:-3]
            seg['bold'] = True
            seg['italic'] = True
        elif (matched.startswith('**') and matched.endswith('**')) or (matched.startswith('__') and matched.endswith('__')):
            seg['text'] = matched[2:-2]
            seg['bold'] = True
        elif (matched.startswith('*') and matched.endswith('*')) or (matched.startswith('_') and matched.endswith('_')):
            seg['text'] = matched[1:-1]
            seg['italic'] = True
        elif matched.startswith('~~') and matched.endswith('~~'):
            seg['text'] = matched[2:-2]
            seg['strike'] = True
        elif matched.startswith('<u>') and matched.endswith('</u>'):
            seg['text'] = matched[3:-4]
            seg['underline'] = True
        elif matched.startswith('==') and matched.endswith('=='):
            seg['text'] = matched[2:-2]
            seg['highlight'] = True
        elif matched.startswith('<mark>') and matched.endswith('</mark>'):
            seg['text'] = matched[6:-7]
            seg['highlight'] = True
        elif matched.startswith('<kbd>') and matched.endswith('</kbd>'):
            seg['text'] = matched[5:-6]
            seg['kbd'] = True
        elif matched.startswith('<sub>') and matched.endswith('</sub>'):
            seg['text'] = matched[5:-6]
            seg['sub'] = True
        elif matched.startswith('~') and matched.endswith('~') and len(matched) > 2:
            seg['text'] = matched[1:-1]
            seg['sub'] = True
        elif matched.startswith('<sup>') and matched.endswith('</sup>'):
            seg['text'] = matched[5:-6]
            seg['sup'] = True
        elif matched.startswith('^') and matched.endswith('^') and len(matched) > 2:
            seg['text'] = matched[1:-1]
            seg['sup'] = True
        elif matched.startswith('[^') and matched.endswith(']'):
            seg['text'] = ''
            seg['footnote'] = matched[2:-1]
        elif matched.startswith('[') and '](' in matched:
            m_l = re.match(r'^\[([^\]]+)\]\(([^)]+)\)$', matched)
            if m_l:
                seg['text'] = m_l.group(1)
                seg['link'] = m_l.group(2)

        styled_segments.append(seg)
        pos = m.end()

    if pos < len(cleaned_text):
        styled_segments.append({
            'text': cleaned_text[pos:],
            'bold': False, 'italic': False, 'strike': False,
            'underline': False, 'sub': False, 'sup': False,
            'highlight': False, 'kbd': False, 'footnote': None, 'link': None
        })

    tokens = []
    for seg in styled_segments:
        if seg['footnote']:
            tokens.append({
                'val': '',
                'bold': False, 'italic': False, 'strike': False, 'underline': False,
                'sub': False, 'sup': False, 'highlight': False, 'kbd': False,
                'code': False, 'math': False, 'footnote': seg['footnote'], 'link': None
            })
            continue

        parts = PLACEHOLDER_SPLIT_RE.split(seg['text'])
        for part in parts:
            if not part:
                continue
            if part in vault:
                kind, val = vault[part]
                tokens.append({
                    'val': val,
                    'bold': seg['bold'],
                    'italic': seg['italic'],
                    'strike': seg['strike'],
                    'underline': seg['underline'],
                    'sub': seg['sub'],
                    'sup': seg['sup'],
                    'highlight': seg['highlight'],
                    'kbd': seg['kbd'],
                    'code': (kind == 'code'),
                    'math': (kind == 'math'),
                    'footnote': None,
                    'link': seg['link']
                })
            else:
                tokens.append({
                    'val': part,
                    'bold': seg['bold'],
                    'italic': seg['italic'],
                    'strike': seg['strike'],
                    'underline': seg['underline'],
                    'sub': seg['sub'],
                    'sup': seg['sup'],
                    'highlight': seg['highlight'],
                    'kbd': seg['kbd'],
                    'code': False,
                    'math': False,
                    'footnote': None,
                    'link': seg['link']
                })
    return tokens


def split_markdown_table_row(line: str) -> List[str]:
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith(r'\|'):
        s = s[:-1]

    cells = []
    cur = []
    in_math = False
    in_code = False
    i = 0
    n = len(s)

    while i < n:
        ch = s[i]
        if ch == '\\' and i + 1 < n:
            if s[i + 1] == '|':
                cur.append('|')
                i += 2
                continue
            cur.append(s[i:i + 2])
            i += 2
            continue
        if ch == '`' and not in_math:
            in_code = not in_code
            cur.append(ch)
            i += 1
            continue
        if ch == '$' and not in_code:
            if i + 1 < n and s[i + 1] == '$':
                in_math = not in_math
                cur.append('$$')
                i += 2
                continue
            in_math = not in_math
            cur.append(ch)
            i += 1
            continue
        if ch == '|' and not in_math and not in_code:
            cells.append("".join(cur).strip())
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1

    cells.append("".join(cur).strip())
    return cells


def unwrap_outer_emphasis(line: str) -> str:
    s = line.strip()
    m = re.match(r'^(?:\*{1,3}|_{1,3})\s*(.+?)\s*(?:\*{1,3}|_{1,3})$', s)
    if m:
        return m.group(1).strip()
    return s


def parse_table_alignment(separator_cells: List[str]) -> List[str]:
    alignments = []
    for c in separator_cells:
        c_clean = c.strip()
        left_c = c_clean.startswith(':')
        right_c = c_clean.endswith(':')
        if left_c and right_c:
            alignments.append('center')
        elif right_c:
            alignments.append('right')
        elif left_c:
            alignments.append('left')
        else:
            alignments.append('auto')
    return alignments


DIAGRAM_LANG_ALIASES = {
    'mermaid': 'mermaid',
    'plantuml': 'plantuml',
    'puml': 'plantuml',
    'dot': 'dot',
    'graphviz': 'dot',
}


def _compute_list_level(raw_line: str) -> int:
    expanded = raw_line.expandtabs(4)
    indent = len(expanded) - len(expanded.lstrip(' '))
    if indent <= 0:
        return 0
    if indent <= 3:
        return 1
    if indent <= 6:
        return 2
    return min(5, 2 + (indent - 4) // 3)


def _normalize_table_caption(text: str) -> Tuple[str, Optional[str]]:
    unwrapped = unwrap_outer_emphasis(text)
    if unwrapped.lower().startswith("продолжение таблицы"):
        norm_tbl = re.sub(
            r'^(Продолжение\s+таблицы)\s*(\d+(?:\.\d+)?)',
            rf'\1{NBSP}\2',
            unwrapped,
            flags=re.IGNORECASE
        ).rstrip('.')
    else:
        norm_tbl = re.sub(
            r'^Таблица\s*(\d+(?:\.\d+)?)\.?\s*[-–—\.:]?\s*',
            rf'Таблица{NBSP}\1{NBSP}–{NBSP}',
            unwrapped,
            flags=re.IGNORECASE
        ).rstrip('.')
    unit_text = None
    m_unit = re.search(r'\((В\s+[^)]+)\)$', norm_tbl)
    if m_unit:
        unit_text = m_unit.group(1)
        norm_tbl = norm_tbl[:m_unit.start()].strip()
    return norm_tbl, unit_text


def _normalize_figure_caption(text: str) -> str:
    unwrapped = unwrap_outer_emphasis(text)
    if re.match(r'^(?:Рисунок|Рис\.?)\s+\d+(?:\.\d+)?', unwrapped, re.IGNORECASE):
        return re.sub(
            r'^(?:Рисунок|Рис\.?)\s*(\d+(?:\.\d+)?)\.?\s*[-–—\.:]?\s*',
            rf'Рисунок{NBSP}\1{NBSP}–{NBSP}',
            unwrapped,
            flags=re.IGNORECASE
        ).rstrip('.')
    return unwrapped.rstrip('.')


def _normalize_listing_caption(text: str) -> str:
    unwrapped = unwrap_outer_emphasis(text)
    if re.match(r'^Листинг\s+(?:\d+|[А-ЯA-Z]\.\d+)', unwrapped, re.IGNORECASE):
        return re.sub(
            r'^Листинг\s*(\d+|[А-ЯA-Z]\.\d+)\.?\s*[-–—\.:]?\s*',
            rf'Листинг{NBSP}\1{NBSP}–{NBSP}',
            unwrapped,
            flags=re.IGNORECASE
        ).rstrip('.')
    return unwrapped.rstrip('.')


class MarkdownASTParser:
    def __init__(self, structural_titles: Optional[List[str]] = None):
        if structural_titles is None:
            self.structural_titles = [
                "аннотация", "содержание", "введение", "заключение",
                "перечень сокращений и условных обозначений",
                "перечень использованных информационных ресурсов",
                "список использованных источников", "список литературы",
                "приложение", "приложение а", "приложение б", "приложение в"
            ]
        else:
            self.structural_titles = [st.strip().lower() for st in structural_titles]

    def extract_frontmatter(self, text: str) -> Tuple[Dict[str, Any], str]:
        meta = {}
        lines = text.splitlines()
        if lines and lines[0].strip() == '---':
            end_idx = -1
            for idx in range(1, len(lines)):
                if lines[idx].strip() == '---':
                    end_idx = idx
                    break
            if end_idx != -1:
                for l in lines[1:end_idx]:
                    if ':' in l:
                        k, v = l.split(':', 1)
                        meta[k.strip()] = v.strip().strip('"').strip("'")
                remaining = "\n".join(lines[end_idx + 1:])
                return meta, remaining
        return meta, text

    def parse(self, text: str) -> List[Dict[str, Any]]:
        text = sanitize_and_repair_latex_escapes(text)
        meta, clean_text = self.extract_frontmatter(text)
        lines = clean_text.splitlines()
        nodes: List[Dict[str, Any]] = []
        i = 0
        n = len(lines)

        while i < n:
            raw = lines[i]
            line = raw.strip()
            if not line:
                i += 1
                continue

            # 1. Горизонтальные разделители ---
            if re.match(r'^(?:-{3,}|\*{3,}|_{3,})$', line):
                i += 1
                continue

            # 2. Блочные формулы $$ ... $$ или \[ ... \]
            if line.startswith('$$') or line.startswith(r'\['):
                end_delim = '$$' if line.startswith('$$') else r'\]'
                f_lines = []
                if line.endswith(end_delim) and len(line) > 4:
                    f_lines.append(line[2:-2].strip())
                    i += 1
                else:
                    f_lines.append(line[2:].strip())
                    i += 1
                    while i < n and not lines[i].strip().endswith(end_delim):
                        f_lines.append(lines[i].strip())
                        i += 1
                    if i < n:
                        f_lines.append(lines[i].strip()[:-2].strip())
                        i += 1
                full_formula = " ".join(f_lines).strip()
                eq_num = None
                m_tag = re.search(r'\\tag\*?\{([^{}]+)\}\s*$', full_formula)
                if m_tag:
                    eq_num = f"({m_tag.group(1).strip('()')})"
                    full_formula = full_formula[:m_tag.start()].strip()
                else:
                    m_num = re.search(r'(?:\\qquad|\\quad|\s+)\((\d+(?:\.\d+)?[а-яa-z]?)\)\s*$', full_formula)
                    if m_num:
                        eq_num = f"({m_num.group(1)})"
                        full_formula = full_formula[:m_num.start()].strip()
                nodes.append({'type': 'formula_block', 'latex': full_formula, 'number': eq_num})
                continue

            # 3. Блоки кода и схем ``` (включая mermaid, plantuml, dot)
            if line.startswith('```'):
                c_lines = []
                lang = line[3:].strip()
                i += 1
                while i < n and not lines[i].strip().startswith('```'):
                    c_lines.append(lines[i].rstrip('\r\n').replace('\u00A0', ' '))
                    i += 1
                if i < n:
                    i += 1
                while c_lines and not c_lines[0].strip():
                    c_lines.pop(0)
                while c_lines and not c_lines[-1].strip():
                    c_lines.pop()

                lang_lower = lang.lower()
                if lang_lower in DIAGRAM_LANG_ALIASES:
                    nodes.append({
                        'type': 'rendered_diagram',
                        'lang': DIAGRAM_LANG_ALIASES[lang_lower],
                        'raw_lang': lang,
                        'code': "\n".join(c_lines),
                        'lines': c_lines,
                        'keep_with_next': False
                    })
                else:
                    box_chars = ('┌', '┐', '└', '┘', '├', '┤', '┬', '┴', '┼', '─', '│', '▼', '▲', '◄', '►', '+--', '-->', '<--')
                    is_diagram = any(any(bc in l for bc in box_chars) for l in c_lines)
                    nodes.append({
                        'type': 'diagram_block' if is_diagram else 'code_block',
                        'lang': lang,
                        'lines': c_lines,
                        'keep_with_next': False
                    })
                continue

            # 4. Сетекст-заголовки
            if i + 1 < n and re.match(r'^[=\-]{3,}$', lines[i + 1].strip()):
                underline = lines[i + 1].strip()
                lvl = 1 if underline.startswith('=') else 2
                h_text = line.rstrip('.')
                style_name, clean_title, outline_level = self._classify_heading(h_text, lvl)
                nodes.append({
                    'type': 'heading',
                    'style': style_name,
                    'text': clean_title,
                    'outline_level': outline_level
                })
                i += 2
                continue

            # 5. ATX-заголовки
            m_h = re.match(r'^(#{1,6})\s+(.*)$', line)
            if m_h:
                lvl = len(m_h.group(1))
                h_text = m_h.group(2).strip()
                h_text = re.sub(r'\s+#+\s*$', '', h_text).rstrip('.')
                style_name, clean_title, outline_level = self._classify_heading(h_text, lvl)
                nodes.append({
                    'type': 'heading',
                    'style': style_name,
                    'text': clean_title,
                    'outline_level': outline_level
                })
                i += 1
                continue

            # 6. Подписи к таблицам
            unwrapped = unwrap_outer_emphasis(line)
            if re.match(r'^Таблица\s+\d+(?:\.\d+)?\s*[-–—\.:]', unwrapped, re.IGNORECASE) or unwrapped.lower().startswith("продолжение таблицы"):
                norm_tbl, unit_text = _normalize_table_caption(unwrapped)
                nodes.append({'type': 'table_caption', 'text': norm_tbl, 'unit': unit_text, 'keep_with_next': True})
                i += 1
                continue

            # 7. Подписи к листингам
            if re.match(r'^Листинг\s+(?:\d+|[А-ЯA-Z]\.\d+)\s*[-–—\.:]', unwrapped, re.IGNORECASE):
                norm_lst = _normalize_listing_caption(unwrapped)
                nodes.append({'type': 'listing_caption', 'text': norm_lst, 'keep_with_next': True})
                i += 1
                continue

            # 8. Таблицы
            if '|' in line and not line.startswith('>'):
                is_tbl_start = False
                if i + 1 < n and '|' in lines[i + 1]:
                    sep_cells = [c.strip() for c in split_markdown_table_row(lines[i + 1])]
                    if all(bool(re.match(r'^:?-{2,}:?$', sc)) for sc in sep_cells if sc):
                        is_tbl_start = True

                if is_tbl_start:
                    tbl_lines = [line]
                    i += 1
                    while i < n and '|' in lines[i].strip():
                        tbl_lines.append(lines[i].strip())
                        i += 1
                    parsed_tbl = self._parse_table(tbl_lines)
                    if parsed_tbl['header']:
                        nodes.append(parsed_tbl)
                    continue

            # 9. Подрисуночные подписи
            if re.match(r'^(?:Рисунок|Рис\.?)\s+\d+(?:\.\d+)?\s*[-–—\.:]', unwrapped, re.IGNORECASE):
                norm_fig = _normalize_figure_caption(unwrapped)
                nodes.append({'type': 'figure_caption', 'text': norm_fig})
                i += 1
                continue

            # 10. Изображения (поддержка угловых скобок, пробелов, заголовков "title", width=...)
            m_img = re.match(r'^!\[(.*?)\]\((.+?)\)(?:\{.*width=([0-9.]+)(cm|mm)?.*\})?$', line)
            if m_img:
                raw_target = m_img.group(2).strip()
                if raw_target.startswith('<'):
                    m_ang = re.match(r"""^<([^>]+)>(?:\s+['"].*['"])?$""", raw_target)
                    img_path = m_ang.group(1).strip() if m_ang else raw_target.strip('<> ')
                else:
                    m_tit = re.match(r"""^(.+?)\s+(['"]).*\2$""", raw_target)
                    img_path = m_tit.group(1).strip() if m_tit else raw_target
                w_val = float(m_img.group(3)) if m_img.group(3) else None
                if w_val is not None and m_img.group(4) == 'mm':
                    w_val = w_val / 10.0
                nodes.append({
                    'type': 'figure_image',
                    'alt': m_img.group(1),
                    'path': img_path,
                    'width_cm': w_val,
                    'keep_with_next': False
                })
                i += 1
                continue

            # 11. Цитаты и Callouts
            if line.startswith('>'):
                quote_lines = []
                while i < n and lines[i].strip().startswith('>'):
                    quote_lines.append(re.sub(r'^>\s?', '', lines[i].strip()))
                    i += 1
                full_quote_text = " ".join(quote_lines).strip()
                m_callout = re.match(r'^\[!(NOTE|WARNING|IMPORTANT|TIP|CAUTION)\]\s*(.*)$', full_quote_text, re.IGNORECASE)
                if m_callout:
                    c_type = m_callout.group(1).upper()
                    c_body = m_callout.group(2).strip()
                    nodes.append({'type': 'callout', 'kind': c_type, 'text': c_body})
                else:
                    nodes.append({'type': 'blockquote', 'text': full_quote_text})
                continue

            # 12. Сноски
            m_fn_def = re.match(r'^\[\^([a-zA-Z0-9_\-]+)\]:\s+(.*)$', line)
            if m_fn_def:
                fn_id = m_fn_def.group(1)
                fn_text = m_fn_def.group(2).strip()
                nodes.append({'type': 'footnote_definition', 'id': fn_id, 'text': fn_text})
                i += 1
                continue

            # 13. Списки (сохраняем исходный символ маркера и уровень вложенности)
            level = _compute_list_level(raw)
            m_task = re.match(r'^([-*+–—])\s+\[([ xX])\]\s+(.*)$', line)
            if m_task:
                is_checked = m_task.group(2).strip().lower() == 'x'
                sym = '☑' if is_checked else '☐'
                nodes.append({
                    'type': 'list_item',
                    'marker': sym,
                    'raw_marker': m_task.group(1),
                    'marker_char': sym,
                    'list_kind': 'task',
                    'level': level,
                    'text': m_task.group(3).strip()
                })
                i += 1
                continue

            m_list = re.match(r'^([-*+–—]|\d+[\.\)]|[а-яa-z]\))\s+(.*)$', line)
            if m_list:
                raw_marker = m_list.group(1)
                item_text = m_list.group(2).strip()
                if raw_marker in ('-', '–', '—'):
                    marker_char = raw_marker
                    list_kind = 'gost_dash'
                    compat_marker = '–'
                elif raw_marker in ('*', '+'):
                    marker_char = raw_marker
                    list_kind = 'bullet'
                    compat_marker = raw_marker
                elif raw_marker.endswith('.'):
                    marker_char = '.'
                    list_kind = 'ordered_dot'
                    compat_marker = raw_marker
                elif raw_marker[0].isdigit() and raw_marker.endswith(')'):
                    marker_char = ')'
                    list_kind = 'ordered_paren'
                    compat_marker = raw_marker
                else:
                    marker_char = ')'
                    list_kind = 'alpha_paren'
                    compat_marker = raw_marker

                nodes.append({
                    'type': 'list_item',
                    'marker': compat_marker,
                    'raw_marker': raw_marker,
                    'marker_char': marker_char,
                    'list_kind': list_kind,
                    'level': level,
                    'text': item_text
                })
                i += 1
                continue

            # 14. Обычный абзац
            nodes.append({'type': 'paragraph', 'text': line})
            i += 1

        return self._link_adjacent_blocks(nodes)

    def _link_adjacent_blocks(self, nodes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Автоматическое связывание соседних блоков:
        - Абзац, начинающийся с 'Таблица N' непосредственно перед таблицей -> table_caption (keep_with_next=True)
        - Абзац непосредственно под изображением/диаграммой -> figure_caption (или связка с последующим figure_caption) + keep_with_next=True на объекте
        - Абзац непосредственно под блоком кода -> listing_caption + keep_with_next=True на блоке кода
        """
        n = len(nodes)
        for idx in range(n):
            cur = nodes[idx]

            # 1. Абзац "Таблица N..." непосредственно перед таблицей
            if cur['type'] == 'paragraph' and idx + 1 < n and nodes[idx + 1]['type'] == 'table':
                unwrapped = unwrap_outer_emphasis(cur['text'])
                if re.match(r'^Таблица\s+\d+(?:\.\d+)?\b', unwrapped, re.IGNORECASE):
                    norm_tbl, unit_text = _normalize_table_caption(unwrapped)
                    nodes[idx] = {
                        'type': 'table_caption',
                        'text': norm_tbl,
                        'unit': unit_text,
                        'keep_with_next': True
                    }
                    continue

            # 2. Изображение / диаграмма и подпись или пояснение под ним
            if cur['type'] in ('figure_image', 'rendered_diagram', 'diagram_block') and idx + 1 < n:
                nxt = nodes[idx + 1]
                if nxt['type'] == 'figure_caption':
                    cur['keep_with_next'] = True
                elif nxt['type'] == 'paragraph':
                    # Если за абзацем сразу идет явная подпись Рисунок N — это подрисуночный пояснительный текст (tc_085)
                    if idx + 2 < n and nodes[idx + 2]['type'] == 'figure_caption':
                        cur['keep_with_next'] = True
                        nxt['keep_with_next'] = True
                    else:
                        unwrapped = unwrap_outer_emphasis(nxt['text'])
                        is_explicit_fig = bool(re.match(r'^(?:Рисунок|Рис\.?|Схема|Диаграмма|График)\b', unwrapped, re.IGNORECASE))
                        is_italic_caption = bool(re.match(r'^(?:\*|_)[^*_]+(?:\*|_)$', nxt['text'].strip()))
                        is_short_caption = len(unwrapped) <= 140 and not unwrapped.endswith((';', ':')) and '.' not in unwrapped[:-1]
                        if is_explicit_fig or is_italic_caption or is_short_caption:
                            cur['keep_with_next'] = True
                            nodes[idx + 1] = {
                                'type': 'figure_caption',
                                'text': _normalize_figure_caption(unwrapped)
                            }

            # 3. Блок кода и подпись листинга под ним
            if cur['type'] == 'code_block' and idx + 1 < n:
                nxt = nodes[idx + 1]
                if nxt['type'] == 'listing_caption':
                    cur['keep_with_next'] = True
                elif nxt['type'] == 'paragraph':
                    unwrapped = unwrap_outer_emphasis(nxt['text'])
                    is_explicit_lst = bool(re.match(r'^Листинг\b', unwrapped, re.IGNORECASE))
                    is_italic_caption = bool(re.match(r'^(?:\*|_)[^*_]+(?:\*|_)$', nxt['text'].strip()))
                    if is_explicit_lst or is_italic_caption:
                        cur['keep_with_next'] = True
                        nodes[idx + 1] = {
                            'type': 'listing_caption',
                            'text': _normalize_listing_caption(unwrapped),
                            'keep_with_next': False
                        }

        return nodes

    def _classify_heading(self, h_text: str, md_level: int) -> Tuple[str, str, int]:
        clean = re.sub(r'[*_`]', '', h_text).strip()
        lower = clean.lower()

        for st in self.structural_titles:
            if lower == st or lower.startswith(st + ' ') or lower.startswith(st + ':'):
                return 'GOST_Heading_Structural', clean, 0

        m_num = re.match(r'^(\d+(?:\.\d+)*)\.?\s+(.*)$', clean)
        if m_num:
            num_str = m_num.group(1)
            title = m_num.group(2).strip()
            parts = num_str.split('.')
            formatted = f"{num_str} {title}"
            if len(parts) == 1:
                return 'GOST_Heading_1', formatted, 0
            elif len(parts) == 2:
                return 'GOST_Heading_2', formatted, 1
            else:
                return 'GOST_Heading_3', formatted, 2

        if md_level <= 2:
            return 'GOST_Heading_Structural', clean, 0
        return 'GOST_Heading_2', clean, 1

    def _parse_table(self, lines: List[str]) -> Dict[str, Any]:
        rows = []
        alignments = []
        for idx, l in enumerate(lines):
            cells = split_markdown_table_row(l)
            if idx == 1 and all(bool(re.match(r'^:?-{2,}:?$', c.strip())) for c in cells if c.strip()):
                alignments = parse_table_alignment(cells)
                continue
            rows.append(cells)
        if not rows:
            return {'type': 'table', 'header': [], 'rows': [], 'alignments': []}
        header = rows[0]
        data_rows = rows[1:]
        return {
            'type': 'table',
            'header': header,
            'rows': data_rows,
            'alignments': alignments
        }
