"""Структура заголовков и нумерация рисунков/таблиц/формул."""

import re

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .config import NUM_MODES
from .docx_utils import make_run, replace_span, rPr_at_position


HEADING_STYLE_RE = re.compile(r'^(?:heading|заголовок)\s*(\d+)$', re.IGNORECASE)
HEADING_MANUAL_NUM_RE = re.compile(r'^[\s\u00a0]*(\d+(?:\.\d+)*)\s*\.?[\s\u00a0]')

# Стили обычного текста: заголовком не считаются даже при заданном уровне.
PLAIN_STYLE_NAMES = {
    'normal', 'обычный', 'body text', 'текст', 'first paragraph',
    'первый абзац', 'list paragraph', 'абзац списка', 'caption', 'название',
    'table caption', 'название таблицы',
}

# Ненумеруемые заголовки: обрывают нумерацию разделов.
UNNUMBERED_HEADING_RE = re.compile(
    r'^[\s\u00a0]*(?:оглавление|содержание|введение|заключение|реферат|'
    r'аннотация|библиограф|список|перечень|references|bibliography|contents|'
    r'introduction|conclusion|abstract|приложение|appendix|термины|'
    r'определения|сокращени|ключевые\s+слова|keywords)',
    re.IGNORECASE)


def _paragraph_style_name(p_el, style_names):
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        return ''
    st = pPr.find(qn('w:pStyle'))
    if st is None:
        return ''
    return style_names.get(st.get(qn('w:val')), '')


def heading_level(p_el, style_names):
    """Уровень заголовка (1, 2, 3…) или None."""
    name = _paragraph_style_name(p_el, style_names).strip()
    m = HEADING_STYLE_RE.match(name)
    if m:
        return int(m.group(1))
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        return None
    ol = pPr.find(qn('w:outlineLvl'))
    if ol is None or name.lower() in PLAIN_STYLE_NAMES:
        return None
    try:
        lvl = int(ol.get(qn('w:val'))) + 1
    except (TypeError, ValueError):
        return None
    return lvl if 1 <= lvl <= 9 else None


def _paragraph_raw_text(p_el):
    return ''.join(t.text or '' for t in p_el.iter(qn('w:t')))


def heading_text_short(p_el, limit=60):
    text = ' '.join(_paragraph_raw_text(p_el).split())
    return text if len(text) <= limit else text[:limit - 1] + '…'


class HeadingPlan:
    """Структура заголовков и расхождения в их нумерации.

    scope – для каждого заголовка (раздел, подраздел, пункт); issues –
    ('wrong' | 'missing' | 'parent' | 'level', w:p, уровень, фактический
    номер, ожидаемый номер).
    """

    MAX_DEPTH = 6

    def __init__(self, doc):
        self.doc = doc
        self.style_names = {}
        try:
            for style in doc.styles:
                try:
                    self.style_names[style.style_id] = style.name or ''
                except Exception:
                    continue
        except Exception:
            pass
        self.scope = {}             # w:p -> (раздел, подраздел, пункт)
        self.issues = []
        self.heading_levels = set()
        self.numbered_levels = set()
        self.count = 0
        self._scan()

    def _scan(self):
        counters = [0] * (self.MAX_DEPTH + 1)   # уровень -> счетчик
        current = [None, None, None]
        for p_el in self.doc.element.body.iter(qn('w:p')):
            level = heading_level(p_el, self.style_names)
            if not level:
                continue
            self.count += 1
            self.heading_levels.add(level)
            text = _paragraph_raw_text(p_el)
            m = HEADING_MANUAL_NUM_RE.match(text)

            if m is None:
                pPr = p_el.find(qn('w:pPr'))
                numPr = pPr.find(qn('w:numPr')) if pPr is not None else None
                if numPr is not None:
                    # Автонумерация Word: номер в тексте не набран.
                    depth = min(level, self.MAX_DEPTH)
                    counters[depth] += 1
                    for k in range(depth + 1, len(counters)):
                        counters[k] = 0
                    expected = tuple(counters[1:depth + 1])
                    for k in range(3):
                        v = expected[k] if k < len(expected) else None
                        current[k] = str(v) if v else None
                    self.scope[p_el] = tuple(current)
                    self.numbered_levels.add(level)
                    continue
                # Заголовок без номера.
                if level > 1 and not UNNUMBERED_HEADING_RE.match(text):
                    # Неясно, начался ли новый раздел: счетчики не трогаем.
                    self.issues.append(('missing', p_el, level, None, None))
                    self.scope[p_el] = tuple(current)
                    continue
                self._reset(counters, current, min(level, 3) - 1)
                self.scope[p_el] = tuple(current)
                continue

            actual = tuple(int(x) for x in m.group(1).split('.'))
            depth = min(len(actual), self.MAX_DEPTH)
            counters[depth] += 1
            for k in range(depth + 1, len(counters)):
                counters[k] = 0
            expected = tuple(counters[1:depth + 1])
            used = expected
            if 0 in expected:
                # Нет номера вышестоящего уровня: ожидаемый неизвестен.
                self.issues.append(('parent', p_el, level, actual, None))
                self._adopt(counters, actual, depth)
                used = actual
            elif len(actual) != level:
                # Глубина номера не совпадает с уровнем заголовка.
                self.issues.append(('level', p_el, level, actual, expected))
                self._adopt(counters, actual, depth)
                used = actual
            elif actual != expected:
                self.issues.append(('wrong', p_el, level, actual, expected))
            for k in range(3):
                v = used[k] if k < len(used) else None
                current[k] = str(v) if v else None
            self.scope[p_el] = tuple(current)
            self.numbered_levels.add(level)

    @staticmethod
    def _adopt(counters, parts, depth):
        """Счетчики по фактическому номеру заголовка."""
        for k in range(depth):
            counters[k + 1] = parts[k]
        for k in range(depth + 1, len(counters)):
            counters[k] = 0

    @staticmethod
    def _reset(counters, current, idx):
        """Заголовок без номера: счетчики с уровня idx обнуляются."""
        for k in range(max(idx, 0), len(counters)):
            counters[k] = 0
        for k in range(max(idx, 0), 3):
            current[k] = None


def fix_heading_numbers(plan):
    """Исправить неверные номера заголовков (недостающие не добавляются)."""
    fixed = 0
    for kind, p_el, _level, _actual, expected in plan.issues:
        if kind != 'wrong' or not expected:
            continue
        paragraph = Paragraph(p_el, plan.doc._body)
        m = HEADING_MANUAL_NUM_RE.match(paragraph.text)
        if m is None:
            continue
        new_num = '.'.join(str(x) for x in expected)
        if m.group(1) == new_num:
            continue
        rPr = rPr_at_position(paragraph, m.start(1))
        if replace_span(paragraph, m.start(1), m.end(1),
                        [make_run(new_num, rPr)]):
            fixed += 1
    return fixed


def log_heading_issues(plan, log):
    """Записать расхождения нумерации заголовков в журнал."""
    shown = hidden = 0
    for kind, p_el, level, actual, expected in plan.issues:
        if shown >= 40:
            hidden += 1
            continue
        shown += 1
        text = heading_text_short(p_el)
        actual_s = '.'.join(str(x) for x in actual) if actual else ''
        if kind == 'wrong':
            log('log_head_wrong', text=text, actual=actual_s,
                expected='.'.join(str(x) for x in expected))
        elif kind == 'missing':
            log('log_head_missing', text=text)
        elif kind == 'parent':
            log('log_head_parent', text=text, actual=actual_s)
        else:
            log('log_head_level', text=text, actual=actual_s, level=level)
    if not plan.issues:
        log('log_head_ok')
    if hidden:
        log('log_head_more', count=hidden)


class ElementNumbering:
    """Номера рисунков/таблиц/формул с учетом структуры документа.

    Режимы: global – 1, 2, 3; section – <раздел>.<номер>; clause –
    <раздел>.<подраздел>.<номер>. Номер раздела берется из автонумерации
    заголовков Word, иначе из номера в тексте заголовка.
    """

    def __init__(self, doc, mode='global', log=None, plan=None):
        self.doc = doc
        self.log = log or (lambda *args, **kwargs: None)
        self.requested = mode if mode in NUM_MODES else 'global'
        self.mode = self.requested
        self.plan = plan if plan is not None else HeadingPlan(doc)
        self.position = {}          # w:p -> порядковый номер в документе
        self.scope = {}             # номер абзаца -> (раздел, подраздел, пункт)
        self.heading_levels = set(self.plan.heading_levels)
        self.numbered_levels = set(self.plan.numbered_levels)
        self.counters = {}          # (префикс, область) -> счетчик
        self._scan()
        self._resolve_mode()

    def _scan(self):
        """Область (раздел, подраздел, пункт) для каждого абзаца."""
        current = (None, None, None)
        for i, p_el in enumerate(self.doc.element.body.iter(qn('w:p'))):
            self.position[p_el] = i
            scope = self.plan.scope.get(p_el)
            if scope is not None:
                current = scope
            self.scope[i] = current

    def _resolve_mode(self):
        if self.mode == 'global':
            return
        if 1 not in self.numbered_levels:
            self.log('log_num_no_headings')
            self.mode = 'global'
            return
        if self.mode == 'clause' and 3 not in self.numbered_levels:
            self.log('log_num_no_clause')
            self.mode = 'section'

    def scope_for(self, paragraph):
        i = self.position.get(paragraph._p)
        if i is None:
            return None, None, None
        return self.scope.get(i, (None, None, None))

    def peek(self, paragraph, prefix):
        """Номер, который получит элемент; счетчик не увеличивается."""
        sec, sub, _clause = self.scope_for(paragraph)
        parts = []
        if self.mode == 'clause' and sec and sub:
            parts = [sec, sub]
        elif self.mode in ('clause', 'section') and sec:
            parts = [sec]
        scope_key = '.'.join(parts)
        n = self.counters.get((prefix, scope_key), 0) + 1
        head = '.'.join(parts)
        num_str = f'{head}.{n}' if head else str(n)
        return num_str, parts, len(parts), scope_key

    def commit(self, prefix, scope_key):
        self.counters[(prefix, scope_key)] = \
            self.counters.get((prefix, scope_key), 0) + 1

