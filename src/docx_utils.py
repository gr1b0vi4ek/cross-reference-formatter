"""Операции с OOXML: runs, поля, закладки."""

import copy
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def make_run(text, rPr=None):
    r = OxmlElement('w:r')
    if rPr is not None:
        r.append(copy.deepcopy(rPr))
    if text:
        t = OxmlElement('w:t')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        r.append(t)
    return r


def make_field_runs(instr, placeholder='1', rPr=None):
    runs = []
    r1 = make_run(None, rPr)
    fc1 = OxmlElement('w:fldChar'); fc1.set(qn('w:fldCharType'), 'begin')
    r1.append(fc1); runs.append(r1)
    r2 = make_run(None, rPr)
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
    it.text = f' {instr} '
    r2.append(it); runs.append(r2)
    r3 = make_run(None, rPr)
    fc2 = OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'), 'separate')
    r3.append(fc2); runs.append(r3)
    runs.append(make_run(placeholder, rPr))
    r5 = make_run(None, rPr)
    fc3 = OxmlElement('w:fldChar'); fc3.set(qn('w:fldCharType'), 'end')
    r5.append(fc3); runs.append(r5)
    return runs


def make_prefix_field(text, rPr=None):
    """Поле QUOTE с префиксом номера: { QUOTE "2." }."""
    if not text:
        return []
    return make_field_runs(f'QUOTE "{text}"', text, rPr)


def make_bookmark_start(name, id_):
    bs = OxmlElement('w:bookmarkStart')
    bs.set(qn('w:id'), str(id_))
    bs.set(qn('w:name'), name)
    return bs


def make_bookmark_end(id_):
    be = OxmlElement('w:bookmarkEnd')
    be.set(qn('w:id'), str(id_))
    return be


def get_run_text_map(paragraph):
    result, cur = [], 0
    for run in paragraph.runs:
        n = len(run.text)
        result.append((cur, cur + n, run))
        cur += n
    return result


def find_overlapping_runs(paragraph, start, end):
    return [(s, e, r) for s, e, r in get_run_text_map(paragraph)
            if e > start and s < end]


def copy_run_formatting(run):
    rPr = run._r.find(qn('w:rPr'))
    return copy.deepcopy(rPr) if rPr is not None else None


def rPr_at_position(paragraph, pos):
    for s, e, run in get_run_text_map(paragraph):
        if s <= pos < e:
            return copy_run_formatting(run)
    return None


def caption_number_rPr(paragraph, num_start=None, strip_size=True):
    """rPr для номера: как у текста подписи, без унаследованного размера."""
    runs = get_run_text_map(paragraph)

    def text_rPrs():
        for s, e, run in runs:
            el = run._r
            if not run.text.strip():
                continue
            if el.find(qn('w:fldChar')) is not None:
                continue
            if el.find(qn('w:instrText')) is not None:
                continue
            yield s, e, run, el.find(qn('w:rPr'))

    best, best_len = None, -1
    for s, e, run, rPr in text_rPrs():
        if rPr is None or rPr.find(qn('w:sz')) is None:
            continue
        if rPr.find(qn('w:vertAlign')) is not None:   # верхний/нижний индекс
            continue
        n = len(run.text.strip())
        if n > best_len:
            best, best_len = rPr, n
    if best is not None:
        return copy.deepcopy(best)

    num_rPr = None
    if num_start is not None:
        for s, e, run, rPr in text_rPrs():
            if s <= num_start < e:
                num_rPr = rPr
                break
    if num_rPr is None:
        return None
    num_rPr = copy.deepcopy(num_rPr)
    if not strip_size:
        return num_rPr if len(num_rPr) else None
    for tag in ('w:sz', 'w:szCs'):
        el = num_rPr.find(qn(tag))
        if el is not None:
            num_rPr.remove(el)
    return num_rPr if len(num_rPr) else None


def replace_span(paragraph, start, end, new_elements):
    overlap = find_overlapping_runs(paragraph, start, end)
    if not overlap:
        return False
    first_start, _, first_run = overlap[0]
    last_start, last_end, last_run = overlap[-1]
    text_before = first_run.text[:start - first_start]
    text_after = last_run.text[end - last_start:]
    rPr = copy_run_formatting(first_run)
    if first_run is last_run:
        first_run.text = text_before
        anchor = first_run._r
        for el in new_elements:
            anchor.addnext(el); anchor = el
        if text_after:
            anchor.addnext(make_run(text_after, rPr))
    else:
        first_run.text = text_before
        last_run.text = text_after
        for _, _, r in overlap[1:-1]:
            p = r._r.getparent()
            if p is not None:
                p.remove(r._r)
        anchor = first_run._r
        for el in new_elements:
            anchor.addnext(el); anchor = el
    return True


def collect_bookmark_names(doc):
    """Имена всех закладок документа."""
    names = set()
    for bs in doc.element.body.iter(qn('w:bookmarkStart')):
        name = bs.get(qn('w:name'))
        if name:
            names.add(name)
    return names


def unique_bookmark_name(base, used_names):
    """Свободное имя закладки: base, base_2, base_3…"""
    if used_names is None:
        return base
    if base not in used_names:
        used_names.add(base)
        return base
    k = 2
    while f'{base}_{k}' in used_names:
        k += 1
    name = f'{base}_{k}'
    used_names.add(name)
    return name


def existing_field_ranges(paragraph):
    """Диапазоны текста абзаца, занятые полями Word."""
    run_map = get_run_text_map(paragraph)
    ranges = []
    depth, start = 0, None
    for s, e, run in run_map:
        for fc in run._r.findall(qn('w:fldChar')):
            ftype = fc.get(qn('w:fldCharType'))
            if ftype == 'begin':
                if depth == 0:
                    start = s
                depth += 1
            elif ftype == 'end':
                if depth > 0:
                    depth -= 1
                    if depth == 0 and start is not None:
                        ranges.append((start, e))
                        start = None
    if start is not None and run_map:
        ranges.append((start, run_map[-1][1]))
    return ranges


def _fields_in_paragraph(paragraph, needle=None):
    """Поля Word в абзаце: start/end по дочерним элементам, instr."""
    children = list(paragraph._p)
    stack, fields = [], []
    for i, ch in enumerate(children):
        if ch.tag == qn('w:fldSimple'):
            instr = ch.get(qn('w:instr')) or ''
            if needle is None or needle in instr:
                fields.append({'start': i, 'end': i, 'start_el': ch,
                               'end_el': ch, 'instr': instr})
            continue
        if ch.tag != qn('w:r'):
            continue
        instr = ''.join(t.text or '' for t in ch.iter(qn('w:instrText')))
        if instr and stack:
            stack[-1]['instr'] += instr
        for fc in ch.findall(qn('w:fldChar')):
            ftype = fc.get(qn('w:fldCharType'))
            if ftype == 'begin':
                stack.append({'start': i, 'start_el': ch, 'instr': ''})
            elif ftype == 'end' and stack:
                f = stack.pop()
                f['end'] = i
                f['end_el'] = ch
                if needle is None or needle in f['instr']:
                    fields.append(f)
    fields.sort(key=lambda f: f['start'])
    return fields


def _bookmark_spans(paragraph, prefix):
    """Закладки prefix* в абзаце: start/end и элементы."""
    children = list(paragraph._p)
    pos = {ch: i for i, ch in enumerate(children)}
    opened, result = {}, []
    for ch in children:
        if ch.tag == qn('w:bookmarkStart'):
            name = ch.get(qn('w:name')) or ''
            if name.startswith(prefix):
                opened[ch.get(qn('w:id'))] = {
                    'name': name, 'start_el': ch, 'end_el': None,
                    'start': pos[ch], 'end': pos[ch]}
        elif ch.tag == qn('w:bookmarkEnd'):
            bm = opened.pop(ch.get(qn('w:id')), None)
            if bm is not None:
                bm['end_el'] = ch
                bm['end'] = pos[ch]
                result.append(bm)
    result.extend(opened.values())
    result.sort(key=lambda b: b['start'])
    return result


def _remove_bookmark_element(bm):
    for key in ('start_el', 'end_el'):
        el = bm.get(key)
        if el is not None and el.getparent() is not None:
            el.getparent().remove(el)


def sync_field_bookmarks(paragraph, needle, prefix, base_names, bm_counter,
                         used_names=None):
    """Одна закладка на каждое поле needle.

    base_names – имена или функция k -> имя для новых закладок.
    Возвращает (names, bm_counter).
    """
    fields = _fields_in_paragraph(paragraph, needle)
    spans = _bookmark_spans(paragraph, prefix)

    names, used = [], set()
    for f in fields:
        found = None
        for j, bm in enumerate(spans):
            if j in used:
                continue
            if bm['start'] <= f['start'] and bm['end'] >= f['end']:
                found = j
                break
        if found is None:
            names.append(None)
        else:
            used.add(found)
            names.append(spans[found]['name'])

    for j, bm in enumerate(spans):
        if j in used or bm['end_el'] is None:
            continue
        _remove_bookmark_element(bm)

    for k, (f, name) in enumerate(zip(fields, names)):
        if name is not None:
            continue
        if callable(base_names):
            base = base_names(k)
        elif base_names and k < len(base_names):
            base = base_names[k]
        else:
            base = f'{prefix}{k + 1:03d}'
        name = unique_bookmark_name(base, used_names)
        _number_group_start(f).addprevious(make_bookmark_start(name, bm_counter))
        f['end_el'].addnext(make_bookmark_end(bm_counter))
        bm_counter += 1
        names[k] = name
    return names, bm_counter


BOOKMARK_NAME_RE = re.compile(r'^(Fig|Tab|Eq|Src)_(\d+)(?:_\d+)*$')

# Поле, которое должна окружать закладка с данным префиксом
BOOKMARK_NEEDLES = {
    'Fig': 'SEQ Figure',
    'Tab': 'SEQ Table',
    'Eq': 'SEQ Equation',
    'Src': 'SEQ Source',
}

PREFIX_RUN_RE = re.compile(r'^[\s\u00a0]*\(?\d+(?:\.\d+)*\.[\s\u00a0]*$')


def _looks_like_number_prefix(text):
    """Текст вида «2.» или «2.1.» – префикс номера."""
    return bool(PREFIX_RUN_RE.match(text or ''))


def _quote_field_start(field):
    """Начало поля QUOTE с префиксом перед данным полем."""
    prev = field['start_el'].getprevious()
    guard = 0
    while prev is not None and guard < 40:
        guard += 1
        if prev.tag == qn('w:fldSimple'):
            instr = prev.get(qn('w:instr')) or ''
            return prev if 'QUOTE' in instr else None
        if prev.tag != qn('w:r'):
            return None
        text = ''.join(t.text or '' for t in prev.findall(qn('w:t')))
        if text.strip():
            return None
        fcs = [fc.get(qn('w:fldCharType'))
               for fc in prev.findall(qn('w:fldChar'))]
        if not fcs:
            prev = prev.getprevious()
            continue
        if 'end' not in fcs:
            return None
        cur = prev.getprevious()
        instr, start_el = '', None
        while cur is not None:
            if cur.tag == qn('w:r'):
                instr = ''.join(t.text or ''
                                for t in cur.findall(qn('w:instrText'))) + instr
                if any(fc.get(qn('w:fldCharType')) == 'begin'
                       for fc in cur.findall(qn('w:fldChar'))):
                    start_el = cur
                    break
            cur = cur.getprevious()
        if start_el is not None and 'QUOTE' in instr:
            return start_el
        return None
    return None


def _number_group_start(field):
    """Первый элемент группы «префикс + поле»."""
    el = _quote_field_start(field)
    if el is None:
        el = field['start_el']
    prev = el.getprevious()
    while prev is not None and prev.tag == qn('w:r'):
        text = ''.join(t.text or '' for t in prev.findall(qn('w:t')))
        if not _looks_like_number_prefix(text):
            break
        el = prev
        prev = el.getprevious()
    return el


def number_prefix(parts):
    """Префикс номера: «2.» для 2.3, пусто для 3."""
    return ('.'.join(parts) + '.') if parts else ''


def bookmark_base_for(prefix, num_str):
    """Имя закладки по номеру: 2.3 -> Fig_002_003."""
    parts = []
    for part in str(num_str).split('.'):
        digits = re.sub(r'\D', '', part) or '0'
        parts.append(digits[-3:].zfill(3))
    return prefix + '_' + '_'.join(parts)


def seq_instruction(seq_name, depth=0):
    """Инструкция SEQ; depth – перезапуск нумерации на заголовке."""
    instr = f'SEQ {seq_name} \\* ARABIC'
    if depth:
        instr += f' \\s {depth}'
    return instr


def _set_field_result(paragraph, field, text):
    """Записать текст в результат поля."""
    children = list(paragraph._p)
    i0 = children.index(field['start_el'])
    i1 = children.index(field['end_el'])
    started, first = False, True
    for ch in children[i0 + 1:i1]:
        if ch.tag != qn('w:r'):
            continue
        types = [fc.get(qn('w:fldCharType'))
                 for fc in ch.findall(qn('w:fldChar'))]
        if types:
            if 'separate' in types and not started:
                started = True
            continue
        if not started:
            continue
        ts = ch.findall(qn('w:t'))
        if not ts:
            if first and text:
                t = OxmlElement('w:t')
                t.text = text
                t.set(qn('xml:space'), 'preserve')
                ch.append(t)
            first = False
            continue
        ts[0].text = text if first else ''
        ts[0].set(qn('xml:space'), 'preserve')
        for extra in ts[1:]:
            extra.text = ''
        first = False
    if first and text:
        # Результата нет: создаем run перед fldChar end.
        model = None
        for ch in reversed(children[i0 + 1:i1]):
            if ch.tag == qn('w:r'):
                model = ch
                break
        new_run = OxmlElement('w:r')
        if model is not None and model.find(qn('w:rPr')) is not None:
            new_run.append(copy.deepcopy(model.find(qn('w:rPr'))))
        t = OxmlElement('w:t')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        new_run.append(t)
        field['end_el'].addprevious(new_run)


def _set_field_instruction(paragraph, field, instr_text):
    """Заменить инструкцию поля."""
    start = field['start_el']
    if start.tag == qn('w:fldSimple'):
        start.set(qn('w:instr'), f' {instr_text} ')
        return
    children = list(paragraph._p)
    i0 = children.index(start)
    i1 = children.index(field['end_el'])
    first = True
    for ch in children[i0 + 1:i1]:
        if ch.tag != qn('w:r'):
            continue
        for it in ch.findall(qn('w:instrText')):
            it.text = f' {instr_text} ' if first else ''
            it.set(qn('xml:space'), 'preserve')
            first = False


def refresh_element_prefix(paragraph, prefix, needle, parts, last, rPr=None,
                           depth=0):
    """Обновить префикс, инструкцию SEQ и результат поля."""
    fields = _fields_in_paragraph(paragraph, needle)
    if not fields:
        return False
    f = fields[0]
    bm = None
    for span in _bookmark_spans(paragraph, prefix + '_'):
        if span['start'] <= f['start'] and span['end'] >= f['end']:
            bm = span
            break
    if bm is None:
        return False

    if rPr is None:
        # Размер берем у текста подписи, а не у прежнего номера.
        rPr = caption_number_rPr(paragraph, strip_size=True)
    if rPr is None:
        # Форматирование прежнего префикса, иначе текста после поля.
        children = list(paragraph._p)
        i_start = children.index(bm['start_el'])
        i_field = children.index(f['start_el'])
        i_end = children.index(f['end_el'])
        for ch in children[i_start + 1:i_field] + children[i_end + 1:]:
            if ch.tag != qn('w:r'):
                continue
            found = ch.find(qn('w:rPr'))
            if found is not None:
                rPr = copy.deepcopy(found)
                break

    # Между закладкой и полем SEQ – прежний префикс: пересобираем заново.
    children = list(paragraph._p)
    i_start = children.index(bm['start_el'])
    i_field = children.index(f['start_el'])
    for ch in children[i_start + 1:i_field]:
        paragraph._p.remove(ch)

    # Закладку могли сдвинуть за префикс – убираем префикс вне закладки.
    prev = bm['start_el'].getprevious()
    for _ in range(2):
        if prev is None or prev.tag != qn('w:r'):
            break
        text = ''.join(t.text or '' for t in prev.findall(qn('w:t')))
        if not _looks_like_number_prefix(text):
            break
        prev_prev = prev.getprevious()
        paragraph._p.remove(prev)
        prev = prev_prev

    text = number_prefix(parts)
    if text:
        anchor = bm['start_el']
        for el in make_prefix_field(text, rPr):
            anchor.addnext(el)
            anchor = el
    seq_name = needle.split()[-1]
    _set_field_instruction(paragraph, f, seq_instruction(seq_name, depth))
    _set_field_result(paragraph, f, last)
    return True


ELEMENT_BM_PREFIXES = ('Fig_', 'Tab_', 'Eq_', 'Src_')

# REF <закладка> [переключатели]
REF_INSTR_RE = re.compile(
    r'^REF\s+([^\s\\]+)((?:\s+\\[a-zA-Z*])*)\s*$', re.IGNORECASE)


def bookmark_numbers(doc):
    """Номер внутри закладки: {имя закладки: '2.1'}."""
    numbers = {}
    for p_el in doc.element.body.iter(qn('w:p')):
        spans = []      # [id закладки, имя, [текст]]
        for ch in p_el:
            if ch.tag == qn('w:bookmarkStart'):
                spans.append([ch.get(qn('w:id')),
                              ch.get(qn('w:name')) or '', []])
                continue
            if ch.tag == qn('w:bookmarkEnd'):
                bid = ch.get(qn('w:id'))
                for span in spans:
                    if span[0] != bid:
                        continue
                    if span[1].startswith(ELEMENT_BM_PREFIXES):
                        text = ''.join(span[2]).strip()
                        if text:
                            numbers[span[1]] = text
                    spans.remove(span)
                    break
                continue
            if spans and ch.tag == qn('w:r'):
                txt = ''.join(t.text or '' for t in ch.findall(qn('w:t')))
                if txt:
                    for span in spans:
                        span[2].append(txt)
    return numbers


def _field_result_info(paragraph, field):
    """Границы результата поля или None, если его нельзя переписать."""
    children = list(paragraph._p)
    try:
        i0 = children.index(field['start_el'])
        i1 = children.index(field['end_el'])
    except ValueError:
        return None
    start = None
    for i in range(i0 + 1, i1):
        ch = children[i]
        if ch.tag != qn('w:r'):
            continue
        types = [fc.get(qn('w:fldCharType'))
                 for fc in ch.findall(qn('w:fldChar'))]
        if start is None:
            if 'separate' in types:
                start = i + 1
            continue
        if 'begin' in types:
            return None
    if start is None:
        return None
    return start, i1


def _set_simple_field_result(el, text):
    """Текст результата простого поля (w:fldSimple)."""
    ts = [t for r in el.findall(qn('w:r')) for t in r.findall(qn('w:t'))]
    if not ts:
        return False
    ts[0].text = text
    ts[0].set(qn('xml:space'), 'preserve')
    for extra in ts[1:]:
        extra.text = ''
    return True


def drop_empty_ref_fields(doc):
    """Удалить пустые поля REF на закладки элементов."""
    removed = 0
    for p_el in list(doc.element.body.iter(qn('w:p'))):
        paragraph = Paragraph(p_el, doc._body)
        for field in list(_fields_in_paragraph(paragraph, 'REF ')):
            m = REF_INSTR_RE.match(' '.join(field['instr'].split()))
            if not m or not m.group(1).startswith(ELEMENT_BM_PREFIXES):
                continue
            el = field['start_el']
            if el.getparent() is None or field['end_el'].getparent() is None:
                continue
            # Вложенное поле внутри – не наш случай.
            inner = el.getnext()
            nested = False
            while inner is not None and inner is not field['end_el']:
                if inner.tag in (qn('w:bookmarkStart'), qn('w:bookmarkEnd')):
                    nested = True
                    break
                if inner.tag == qn('w:fldSimple') or (
                        inner.tag == qn('w:r') and any(
                            fc.get(qn('w:fldCharType')) == 'begin'
                            for fc in inner.findall(qn('w:fldChar')))):
                    nested = True
                    break
                inner = inner.getnext()
            if nested:
                continue
            shown = []
            node = el
            while node is not None:
                if node.tag == qn('w:r'):
                    shown.extend(t.text or '' for t in node.findall(qn('w:t')))
                if node is field['end_el']:
                    break
                node = node.getnext()
            if ''.join(shown).strip():
                continue
            node, stop = el, field['end_el']
            while node is not None:
                nxt = node.getnext()
                parent = node.getparent()
                parent.remove(node)
                if node is stop:
                    break
                node = nxt
            removed += 1
    return removed


def refresh_ref_results(doc):
    """Показать в полях REF актуальные номера закладок."""
    numbers = bookmark_numbers(doc)
    if not numbers:
        return 0
    updated = 0
    for p_el in list(doc.element.body.iter(qn('w:p'))):
        paragraph = Paragraph(p_el, doc._body)
        for field in _fields_in_paragraph(paragraph, 'REF '):
            instr = ' '.join(field['instr'].split())
            m = REF_INSTR_RE.match(instr)
            if not m:
                continue
            # \n, \r, \p и т. п. меняют результат – такой кэш не трогаем.
            if [s for s in m.group(2).split() if s.lower() != '\\h']:
                continue
            new_text = numbers.get(m.group(1))
            if not new_text:
                continue
            if field['start_el'].tag == qn('w:fldSimple'):
                ts = [t for r in field['start_el'].findall(qn('w:r'))
                      for t in r.findall(qn('w:t'))]
                if ''.join(t.text or '' for t in ts) == new_text:
                    continue
                if _set_simple_field_result(field['start_el'], new_text):
                    updated += 1
                continue
            info = _field_result_info(paragraph, field)
            if info is None:
                continue
            children = list(paragraph._p)
            start, end = info
            shown = ''.join(t.text or '' for i in range(start, end)
                            for t in children[i].findall(qn('w:t')))
            if shown == new_text:
                continue
            _set_field_result(paragraph, field, new_text)
            updated += 1
    return updated


def _rename_field_targets(doc, rename):
    """Перенаправить REF и гиперссылки на новые имена закладок."""
    if not rename:
        return
    pattern = re.compile(
        r'(?<![0-9A-Za-z_])(' +
        '|'.join(re.escape(k) for k in sorted(rename, key=len, reverse=True)) +
        r')(?![0-9A-Za-z_])')
    body = doc.element.body
    for instr in body.iter(qn('w:instrText')):
        txt = instr.text or ''
        new_txt = pattern.sub(lambda m: rename[m.group(1)], txt)
        if new_txt != txt:
            instr.text = new_txt
    for fs in body.iter(qn('w:fldSimple')):
        txt = fs.get(qn('w:instr'))
        if txt:
            new_txt = pattern.sub(lambda m: rename[m.group(1)], txt)
            if new_txt != txt:
                fs.set(qn('w:instr'), new_txt)
    for hl in body.iter(qn('w:hyperlink')):
        anchor = hl.get(qn('w:anchor'))
        if anchor in rename:
            hl.set(qn('w:anchor'), rename[anchor])


def renumber_bookmarks(doc, prefixes=None, rename=None):
    """Перенумеровать закладки по порядку: Fig_001, Src_001…

    prefixes – какие префиксы перенумеровывать (None – все). В режиме
    нумерации по разделам имена Fig_002_003 сохраняются, поэтому для них
    функция не вызывается. rename пополняется парами «старое -> новое».
    """
    if rename is None:
        rename = {}
    counters = {}
    seen = set()
    renamed = 0
    for bs in doc.element.body.iter(qn('w:bookmarkStart')):
        name = bs.get(qn('w:name')) or ''
        m = BOOKMARK_NAME_RE.match(name)
        if not m:
            continue
        prefix = m.group(1)
        if prefixes is not None and prefix not in prefixes:
            continue
        counters[prefix] = counters.get(prefix, 0) + 1
        new_name = f'{prefix}_{counters[prefix]:03d}'
        if name in seen:
            # Дубликат: уникальное имя. REF ведут на первое вхождение.
            bs.set(qn('w:name'), new_name)
            renamed += 1
            continue
        seen.add(name)
        if new_name != name:
            rename[name] = new_name
            bs.set(qn('w:name'), new_name)
            renamed += 1
    return renamed


def apply_desired_names(doc, desired, rename=None):
    """Привести имена закладок к номерам элементов (Fig_002_003).

    desired: {текущее имя: нужное имя}. rename пополняется парами
    «старое -> новое» для полей REF.
    """
    if not desired:
        return 0
    if rename is None:
        rename = {}
    starts = list(doc.element.body.iter(qn('w:bookmarkStart')))
    order = {id(bs): i for i, bs in enumerate(starts)}
    by_name = {}
    for bs in starts:
        by_name.setdefault(bs.get(qn('w:name')) or '', []).append(bs)

    pending = []
    for old, base in desired.items():
        for bs in by_name.get(old, []):
            pending.append((bs, old, base))
    pending.sort(key=lambda t: order.get(id(t[0]), 0))

    moving = {id(bs) for bs, _old, _base in pending}
    reserved = {name for name, lst in by_name.items()
                if any(id(bs) not in moving for bs in lst)}
    # Закладки, которые остаются под своим именем, тоже его занимают.
    reserved |= {old for _bs, old, base in pending if base == old}

    renamed = 0
    for bs, old, base in pending:
        if base == old:
            continue
        if base in reserved:
            new = unique_bookmark_name(base, reserved)
        else:
            new = base
            reserved.add(new)
        bs.set(qn('w:name'), new)
        rename.setdefault(old, new)
        renamed += 1
    return renamed


def _enclosing_paragraph(element):
    """Ближайший предок w:p."""
    parent = element.getparent()
    while parent is not None and parent.tag != qn('w:p'):
        parent = parent.getparent()
    return parent


def _remove_element(el):
    if el is not None and el.getparent() is not None:
        el.getparent().remove(el)


def repair_field_bookmarks(doc):
    """Поставить закладки элементов вокруг их полей SEQ.

    Закладка переносится в свой абзац; закладки без поля или повторные
    закладки на одном поле удаляются. Возвращает (moved, removed).
    """
    body = doc.element.body
    ends_by_id = {}
    for be in body.iter(qn('w:bookmarkEnd')):
        bid = be.get(qn('w:id'))
        ends_by_id.setdefault(bid, []).append(be)

    moved = removed = 0
    # Ключ – элемент абзаца: он держит ссылку, поэтому id() не переиспользуется.
    claimed = {}  # w:p -> индексы полей, к которым уже есть закладка

    for bs in list(body.iter(qn('w:bookmarkStart'))):
        m = BOOKMARK_NAME_RE.match(bs.get(qn('w:name')) or '')
        if not m:
            continue
        bid = bs.get(qn('w:id'))
        ends = ends_by_id.get(bid) or []
        be = ends.pop(0) if ends else None
        if be is None:
            _remove_element(bs)
            removed += 1
            continue
        if be.getparent() is None:
            _remove_element(bs)
            removed += 1
            continue

        para = _enclosing_paragraph(bs)
        if para is None:
            para = _enclosing_paragraph(be)
        if para is None:
            _remove_element(bs)
            _remove_element(be)
            removed += 1
            continue

        paragraph = Paragraph(para, doc._body)
        fields = _fields_in_paragraph(paragraph, BOOKMARK_NEEDLES[m.group(1)])
        used = claimed.setdefault(para, set())
        children = list(para)
        end_pos = children.index(be) if be in children else -1

        target = None
        # Поле, внутри которого уже стоит конец закладки.
        for idx, f in enumerate(fields):
            if idx not in used and f['start'] <= end_pos <= f['end']:
                target = (idx, f)
                break
        if target is None:
            for idx, f in enumerate(fields):
                if idx not in used:
                    target = (idx, f)
                    break

        if target is None:
            # Поля нет (или на нем уже есть закладка) – закладка лишняя.
            _remove_element(bs)
            _remove_element(be)
            removed += 1
            continue

        idx, f = target
        used.add(idx)
        # Закладка охватывает и префикс номера («2.» перед SEQ).
        start_el = _number_group_start(f)
        if start_el.getprevious() is not bs:
            start_el.addprevious(bs)
            moved += 1
        if f['end_el'].getnext() is not be:
            f['end_el'].addnext(be)
            moved += 1
    return moved, removed


def get_max_bookmark_id(doc):
    max_id = 0
    for bs in doc.element.body.iter(qn('w:bookmarkStart')):
        try:
            n = int(bs.get(qn('w:id')))
            if n > max_id:
                max_id = n
        except (ValueError, TypeError):
            pass
    return max_id