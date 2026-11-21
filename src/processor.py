"""Основная логика: подписи, формулы, библиография, ссылки."""

import re

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import config, i18n
from .config import NUM_MODES
from .patterns import (
    get_caption_re, get_fig_ref_re, get_tab_ref_re, get_eq_ref_re,
    get_bib_heading_re, normalize_num, NUM_PART,
    SRC_REF_RE, BIB_END_RE, BIB_HEADING_RE_RU, BIB_HEADING_RE_EN,
    TOC_LINE_RE, MATH_SYMBOLS,
)
from .numbering import (
    HeadingPlan, ElementNumbering, fix_heading_numbers, log_heading_issues,
)
from .docx_utils import (
    make_run, make_field_runs, make_prefix_field,
    make_bookmark_start, make_bookmark_end,
    find_overlapping_runs, copy_run_formatting,
    rPr_at_position, replace_span, get_max_bookmark_id,
    collect_bookmark_names, unique_bookmark_name, existing_field_ranges,
    sync_field_bookmarks, renumber_bookmarks, _rename_field_targets,
    apply_desired_names, repair_field_bookmarks, refresh_element_prefix,
    drop_empty_ref_fields, refresh_ref_results, caption_number_rPr,
    number_prefix, bookmark_base_for, seq_instruction,
)


def process_caption(paragraph, bm_counter, seq_name, prefix, numbering, log,
                    used_names=None, desired=None):
    """Пронумеровать подпись рисунка/таблицы.

    desired – «имя закладки -> имя по номеру элемента».
    """
    m = get_caption_re().match(paragraph.text)
    if not m:
        return None, None, bm_counter, None

    label = m.group(1).lower()
    if i18n.current_doc_lang == 'en':
        is_figure = 'fig' in label
    else:
        is_figure = label.startswith('рис')
    seq = 'Figure' if is_figure else 'Table'
    pfx = 'Fig' if is_figure else 'Tab'
    kind = 'figure' if is_figure else 'table'

    num_str, parts, depth, scope_key = numbering.peek(paragraph, pfx)
    bm_base = bookmark_base_for(pfx, num_str)
    last = num_str.rsplit('.', 1)[-1]
    pre = number_prefix(parts)
    needle = f'SEQ {seq}'

    if needle in paragraph._p.xml:
        names, bm_counter = sync_field_bookmarks(
            paragraph, needle, pfx + '_', [bm_base], bm_counter, used_names)
        bm = names[0] if names else None
        if not bm:
            return None, None, bm_counter, None
        rPr = caption_number_rPr(paragraph, m.start(2))
        refresh_element_prefix(paragraph, pfx, needle, parts, last,
                               rPr=rPr, depth=depth)
        if desired is not None:
            desired[bm] = bm_base
        numbering.commit(pfx, scope_key)
        return bm, num_str, bm_counter, kind

    num_start, num_end = m.start(2), m.end(2)
    overlap = find_overlapping_runs(paragraph, num_start, num_end)
    if not overlap:
        return None, None, bm_counter, None
    # Шрифт номера как у текста подписи.
    rPr = caption_number_rPr(paragraph, num_start)

    bm_name = unique_bookmark_name(bm_base, used_names)
    elements = [make_bookmark_start(bm_name, bm_counter)]
    elements.extend(make_prefix_field(pre, rPr))
    elements.extend(make_field_runs(seq_instruction(seq, depth), last, rPr))
    elements.append(make_bookmark_end(bm_counter))
    if not replace_span(paragraph, num_start, num_end, elements):
        return None, None, bm_counter, None
    if desired is not None:
        desired[bm_name] = bm_base
    numbering.commit(pfx, scope_key)
    return bm_name, num_str, bm_counter + 1, kind


def _full_text(paragraph):
    parts = []
    for elem in paragraph._p.iter():
        if elem.tag in (qn('w:t'), qn('m:t')):
            parts.append(elem.text or '')
    return ''.join(parts)


def _find_number(paragraph):
    text_w = paragraph.text
    if text_w:
        m = re.search(r'\(\s*(' + NUM_PART + r')\s*\)\s*$', text_w)
        if m:
            return m.group(1), 'paren', 'wt'
        m = re.search(r'(?<![\wА-Яеа-яе])(' + NUM_PART + r')\.?\s*$', text_w)
        if m:
            return m.group(1), 'bare', 'wt'
    full = _full_text(paragraph)
    if full:
        m = re.search(r'\(\s*(' + NUM_PART + r')\s*\)\s*$', full)
        if m:
            return m.group(1), 'paren', 'mt'
        m = re.search(r'(?<![\wА-Яеа-яе])(' + NUM_PART + r')\.?\s*$', full)
        if m:
            return m.group(1), 'bare', 'mt'
    return None, None, None


def is_formula_paragraph(paragraph):
    if 'oMath' in paragraph._p.xml:
        return True
    st = paragraph.style.name.lower()
    if 'equation' in st or 'формул' in st or 'formula' in st:
        return True
    text = paragraph.text.strip()
    if not text:
        return bool(_find_number(paragraph)[0])
    num, kind, _ = _find_number(paragraph)
    if not num:
        return False
    return kind == 'paren' or any(s in text for s in MATH_SYMBOLS)


def _remove_omml_suffix(paragraph, num, kind):
    if kind == 'paren':
        pattern = r'\(\s*' + re.escape(num) + r'\s*\)\s*$'
    else:
        pattern = re.escape(num) + r'\s*$'
    for mt in reversed(list(paragraph._p.iter(qn('m:t')))):
        text = mt.text or ''
        new_text = re.sub(pattern, '', text)
        if new_text != text:
            mt.text = new_text
            return True
    for mt in reversed(list(paragraph._p.iter(qn('m:t')))):
        text = mt.text or ''
        if num in text:
            mt.text = text.replace(num, '', 1)
            return True
    return False


def process_formula(paragraph, bm_counter, bm_eq, log, used_names=None,
                    numbering=None, desired=None):
    """Пронумеровать формулу в тексте (до или после формулы, «(1)» или «1»)."""
    needle = 'SEQ Equation'
    if needle in paragraph._p.xml:
        num_str, parts, depth, scope_key = numbering.peek(paragraph, 'Eq')
        bm_base = bookmark_base_for('Eq', num_str)
        names, bm_counter = sync_field_bookmarks(
            paragraph, needle, 'Eq_', [bm_base], bm_counter, used_names)
        bm = names[0] if names else None
        if not bm:
            return None, None, bm_counter
        refresh_element_prefix(paragraph, 'Eq', needle, parts,
                               num_str.rsplit('.', 1)[-1], depth=depth)
        if desired is not None:
            desired[bm] = bm_base
        numbering.commit('Eq', scope_key)
        bm_eq[num_str] = bm
        return bm, num_str, bm_counter

    num, kind, source = _find_number(paragraph)
    if not num:
        return None, None, bm_counter

    num_str, parts, depth, scope_key = numbering.peek(paragraph, 'Eq')
    bm_base = bookmark_base_for('Eq', num_str)
    bm_name = unique_bookmark_name(bm_base, used_names)
    last = num_str.rsplit('.', 1)[-1]
    pre = number_prefix(parts)
    instr = seq_instruction('Equation', depth)

    if source == 'wt':
        text_w = paragraph.text
        if kind == 'paren':
            m = re.search(r'\(\s*(' + re.escape(num) + r')\s*\)\s*$', text_w)
        else:
            m = re.search(r'(?<![\wА-Яеа-яе])(' + re.escape(num) + r')\.?\s*$',
                          text_w)
        if not m:
            return None, None, bm_counter
        if kind == 'paren':
            s, e = m.start(1), m.end(1)
        else:
            s, e = m.start(1), m.end(0)
        overlap = find_overlapping_runs(paragraph, s, e)
        if not overlap:
            return None, None, bm_counter
        rPr = copy_run_formatting(overlap[0][2])
        elements = [make_bookmark_start(bm_name, bm_counter)]
        elements.extend(make_prefix_field(pre, rPr))
        elements.extend(make_field_runs(instr, last, rPr))
        elements.append(make_bookmark_end(bm_counter))
        if not replace_span(paragraph, s, e, elements):
            return None, None, bm_counter
        if desired is not None:
            desired[bm_name] = bm_base
        numbering.commit('Eq', scope_key)
        bm_eq[num_str] = bm_name
        return bm_name, num_str, bm_counter + 1

    if not _remove_omml_suffix(paragraph, num, kind):
        return None, None, bm_counter
    parent_oMath = None
    for elem in paragraph._p.iter(qn('m:oMath')):
        parent_oMath = elem
    if parent_oMath is None:
        return None, None, bm_counter
    elements = [make_run(' (' if kind == 'paren' else ' ')]
    elements.append(make_bookmark_start(bm_name, bm_counter))
    elements.extend(make_prefix_field(pre))
    elements.extend(make_field_runs(instr, last))
    elements.append(make_bookmark_end(bm_counter))
    if kind == 'paren':
        elements.append(make_run(')'))
    anchor = parent_oMath
    for el in elements:
        anchor.addnext(el); anchor = el
    if desired is not None:
        desired[bm_name] = bm_base
    numbering.commit('Eq', scope_key)
    bm_eq[num_str] = bm_name
    return bm_name, num_str, bm_counter + 1


def collect_formula_table_rows(doc):
    """Формулы в таблицах: формула слева, номер справа.

    Возвращает {элемент абзаца с номером: (номер, вид)}; документ не меняет.
    """
    rows = {}
    for table in doc.tables:
        try:
            if len(table.columns) < 2:
                continue
        except Exception:
            continue
        for row in table.rows:
            try:
                cells = row.cells
            except Exception:
                continue
            if len(cells) < 2:
                continue
            left, right = cells[0], cells[-1]
            rt = right.text.strip()
            m = re.match(r'^\(\s*(' + NUM_PART + r')\s*\)$', rt)
            if m:
                num_orig, kind = m.group(1), 'paren'
            else:
                m = re.match(r'^(' + NUM_PART + r')$', rt)
                if not m:
                    continue
                num_orig, kind = m.group(1), 'bare'
            has_omml = 'oMath' in left._tc.xml
            has_math = any(c in left.text for c in MATH_SYMBOLS)
            if not (has_omml or has_math):
                continue
            if not right.paragraphs:
                continue
            rows[right.paragraphs[0]._p] = (num_orig, kind)
    return rows


def process_table_formula(paragraph, num_orig, kind, bm_counter, bm_eq,
                          numbering, log, used_names=None, desired=None):
    """Пронумеровать формулу таблицы (номер в правой ячейке)."""
    num_str, parts, depth, scope_key = numbering.peek(paragraph, 'Eq')
    bm_base = bookmark_base_for('Eq', num_str)
    last = num_str.rsplit('.', 1)[-1]
    pre = number_prefix(parts)

    if 'SEQ Equation' in paragraph._p.xml:
        names, bm_counter = sync_field_bookmarks(
            paragraph, 'SEQ Equation', 'Eq_', [bm_base], bm_counter, used_names)
        bm = names[0] if names else None
        if not bm:
            return bm_counter
        refresh_element_prefix(paragraph, 'Eq', 'SEQ Equation', parts, last,
                               depth=depth)
        if desired is not None:
            desired[bm] = bm_base
        numbering.commit('Eq', scope_key)
        bm_eq[num_str] = bm
        return bm_counter

    idx = paragraph.text.rfind(num_orig)
    if idx < 0:
        return bm_counter

    bm_name = unique_bookmark_name(bm_base, used_names)
    rPr = rPr_at_position(paragraph, idx)
    elements = [make_bookmark_start(bm_name, bm_counter)]
    elements.extend(make_prefix_field(pre, rPr))
    elements.extend(make_field_runs(
        seq_instruction('Equation', depth), last, rPr))
    elements.append(make_bookmark_end(bm_counter))
    if not replace_span(paragraph, idx, idx + len(num_orig), elements):
        return bm_counter
    if desired is not None:
        desired[bm_name] = bm_base
    numbering.commit('Eq', scope_key)
    bm_eq[num_str] = bm_name
    return bm_counter + 1


def _has_numPr(paragraph):
    pPr = paragraph._p.find(qn('w:pPr'))
    return pPr is not None and pPr.find(qn('w:numPr')) is not None


def _read_num_format(doc, paragraph):
    pPr = paragraph._p.find(qn('w:pPr'))
    if pPr is None:
        return None
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        return None
    num_id_el = numPr.find(qn('w:numId'))
    ilvl_el = numPr.find(qn('w:ilvl'))
    if num_id_el is None:
        return None
    num_id = num_id_el.get(qn('w:val'))
    ilvl = ilvl_el.get(qn('w:val')) if ilvl_el is not None else '0'
    try:
        numbering = doc.part.numbering_part.element
    except Exception:
        return None
    abstract_id = None
    for num in numbering.findall(qn('w:num')):
        if num.get(qn('w:numId')) == num_id:
            ab = num.find(qn('w:abstractNumId'))
            if ab is not None:
                abstract_id = ab.get(qn('w:val'))
                break
    if abstract_id is None:
        return None
    for absnum in numbering.findall(qn('w:abstractNum')):
        if absnum.get(qn('w:abstractNumId')) == abstract_id:
            for lvl in absnum.findall(qn('w:lvl')):
                if lvl.get(qn('w:ilvl')) == ilvl:
                    lvl_text_el = lvl.find(qn('w:lvlText'))
                    num_fmt_el = lvl.find(qn('w:numFmt'))
                    if lvl_text_el is None:
                        return None
                    fmt = num_fmt_el.get(qn('w:val')) \
                        if num_fmt_el is not None else 'decimal'
                    if fmt != 'decimal':
                        return None
                    lvl_text = lvl_text_el.get(qn('w:val')) or ''
                    marker = '%' + str(int(ilvl) + 1)
                    return lvl_text.split(marker)
    return None


def _replace_numPr_with_seq(doc, paragraph, bm_counter, bm_map, src_pos, log,
                            used_names=None):
    parts = _read_num_format(doc, paragraph)
    if not parts:
        parts = ['', '.\u00a0']
    elif len(parts) > 1 and parts[1].strip() == '.':
        parts[1] = '.\u00a0'

    pPr = paragraph._p.find(qn('w:pPr'))
    numPr = pPr.find(qn('w:numPr'))
    pPr.remove(numPr)
    key = str(src_pos)
    bm_name = unique_bookmark_name(f'Src_{src_pos:03d}', used_names)
    elements = []
    if parts[0]:
        elements.append(make_run(parts[0]))
    elements.append(make_bookmark_start(bm_name, bm_counter))
    elements.extend(make_field_runs('SEQ Source \\* ARABIC', key))
    elements.append(make_bookmark_end(bm_counter))
    if len(parts) > 1 and parts[1]:
        elements.append(make_run(parts[1]))
    first_run = paragraph._p.find(qn('w:r'))
    if first_run is not None:
        for el in elements:
            first_run.addprevious(el)
    else:
        for el in elements:
            paragraph._p.append(el)
    bm_map[key] = bm_name
    return bm_counter + 1


def _find_bib_items(text):
    results = []
    m = re.match(r'^[\s\xa0]*(?:\[\s*(\d+)\s*\]|(\d+)\s*([.)]))([\s\xa0]+)?', text)
    if not m:
        return []

    if m.group(1) is not None:
        num_str = m.group(1)
        prefix = '['
        suffix = ']'
        s, e = m.start(1), m.end(1)
    else:
        num_str = m.group(2)
        prefix = ''
        suffix = m.group(3)
        s, e = m.start(2), m.end(2)

    e_ext = m.end()
    results.append((num_str, s, e_ext, prefix, suffix))

    prev = int(num_str)
    search_from = m.end()
    while True:
        target = prev + 1
        pat = re.compile(
            r'(?:\s{2,}|[\n\t])[\s\xa0]*'
            r'(?:\[\s*' + str(target) + r'\s*\]|' + str(target) + r'\s*([.)]))([\s\xa0]+)?'
        )
        mm = pat.search(text, search_from)
        if not mm:
            break

        num_match = re.search(str(target), mm.group(0))
        if not num_match:
            break
        s = mm.start() + num_match.start()
        e = s + len(str(target))

        if '[' in mm.group(0):
            prefix = '['
            suffix = ']'
        else:
            prefix = ''
            suffix = mm.group(1)

        e_ext = mm.end()
        results.append((str(target), s, e_ext, prefix, suffix))
        prev = target
        search_from = mm.end()
    return results


def process_bibliography(doc, bm_counter, log=print, used_names=None):
    bm_map = {}
    in_bib = bib_started = False
    src_pos = 1
    items = 0

    heading_re = get_bib_heading_re()
    other_heading_re = BIB_HEADING_RE_RU if i18n.current_doc_lang == 'en' else BIB_HEADING_RE_EN

    all_paras = list(doc.paragraphs)
    seen = {id(p._p) for p in all_paras}
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if id(p._p) not in seen:
                        seen.add(id(p._p))
                        all_paras.append(p)

    for para in all_paras:
        text = para.text.strip()
        if not text:
            continue

        if TOC_LINE_RE.search(text) or 'toc' in para.style.name.lower():
            continue

        if not bib_started:
            if heading_re.match(text):
                in_bib = bib_started = True
                log('log_bib_found', text=text[:60])
            continue
        if not in_bib:
            continue

        if BIB_END_RE.match(text) or other_heading_re.match(text):
            in_bib = False
            log('log_bib_end', text=text[:40])
            continue

        if 'SEQ Source' in para._p.xml:
            # По одной закладке на поле SEQ Source: недостающие создаются,
            # лишние удаляются.
            names, bm_counter = sync_field_bookmarks(
                para, 'SEQ Source', 'Src_',
                lambda k: f'Src_{src_pos + k:03d}',
                bm_counter, used_names)
            for bm in names:
                bm_map[str(src_pos)] = bm
                items += 1
                src_pos += 1
            continue

        if _has_numPr(para):
            bm_counter = _replace_numPr_with_seq(
                doc, para, bm_counter, bm_map, src_pos, log,
                used_names=used_names)
            items += 1
            src_pos += 1
            continue

        matches = _find_bib_items(para.text)
        if not matches:
            if len(text) > 20:
                log('log_bib_warn', text=text[:60])
            continue

        plan = []
        for num_str, s, e_ext, prefix, suffix in matches:
            plan.append((s, e_ext, prefix, suffix, num_str, src_pos))
            src_pos += 1

        for s, e_ext, prefix, suffix, num_str, pos in reversed(plan):
            rPr = rPr_at_position(para, s)
            bm_name = unique_bookmark_name(f'Src_{pos:03d}', used_names)

            elements = []
            if prefix:
                elements.append(make_run(prefix))
            elements.append(make_bookmark_start(bm_name, bm_counter))
            elements.extend(make_field_runs('SEQ Source \\* ARABIC', str(pos), rPr))
            elements.append(make_bookmark_end(bm_counter))
            if suffix:
                elements.append(make_run(suffix))
            elements.append(make_run('\u00a0'))

            if replace_span(para, s, e_ext, elements):
                bm_map[str(pos)] = bm_name
                bm_counter += 1
                items += 1
                log('log_bib_item', bm_name=bm_name)

    log('log_bib_total', count=items)
    return bm_map, bm_counter


def _resolve_key(key, bm_map):
    key = str(key).strip()
    if key in bm_map:
        return bm_map[key]
    try:
        num = int(key)
        norm = str(num)
    except ValueError:
        num, norm = None, None
    if norm and norm in bm_map:
        return bm_map[norm]
    for k in bm_map:
        if k.startswith(key + '_'):
            return bm_map[k]
    # «рисунок 3» при нумерации 2.1-стилем: третий рисунок по порядку.
    if num is not None and 1 <= num <= len(bm_map):
        return bm_map[list(bm_map)[num - 1]]
    return None


def _real_key(bm_name, bm_map, default=None):
    """Ключ (номер), под которым закладка записана в карте номеров."""
    for k, v in bm_map.items():
        if v == bm_name:
            return k
    return default


def build_number_refs(num_str, bm_map, rPr, placeholder=None):
    if placeholder is None:
        placeholder = num_str

    norm_str = re.sub(r'\s+and\s+', ',', num_str, flags=re.IGNORECASE)
    norm_str = re.sub(r'\s+и\s+', ',', norm_str)
    normalized = normalize_num(norm_str)

    m = re.match(
        r'^(' + NUM_PART + r')-(' + NUM_PART + r')$',
        normalized)
    if m:
        k1, k2 = m.group(1), m.group(2)
        bm1 = _resolve_key(k1, bm_map)
        bm2 = _resolve_key(k2, bm_map)
        if bm1 and bm2:
            sep = '-'
            sm = re.search(
                NUM_PART + r'(\s*[-\u2013\u2014\u2212]\s*)' + NUM_PART,
                num_str)
            if sm:
                sep = sm.group(1)
            return [
                *make_field_runs(f'REF {bm1} \\h',
                                 _real_key(bm1, bm_map, k1), rPr),
                make_run(sep, rPr),
                *make_field_runs(f'REF {bm2} \\h',
                                 _real_key(bm2, bm_map, k2), rPr),
            ]
        return None

    if ',' in normalized:
        parts = [p.strip() for p in normalized.split(',') if p.strip()]
        bookmarks = [_resolve_key(p, bm_map) for p in parts]
        if parts and all(bookmarks):
            elems = []
            sep = ', '
            sm = re.search(NUM_PART + r'(\s*,\s*)' + NUM_PART, num_str)
            if sm:
                sep = sm.group(1)
            for i, p in enumerate(parts):
                if i > 0:
                    elems.append(make_run(sep, rPr))
                elems.extend(make_field_runs(
                    f'REF {bookmarks[i]} \\h',
                    _real_key(bookmarks[i], bm_map, p), rPr))
            return elems
        return None

    bm = _resolve_key(normalized, bm_map)
    if bm:
        return make_field_runs(
            f'REF {bm} \\h', _real_key(bm, bm_map, placeholder), rPr)
    return None


def parse_source_inner(inner):
    tokens = [t.strip() for t in inner.split(',') if t.strip()]
    items = []
    i = 0
    while i < len(tokens):
        t = tokens[i]
        m = re.match(r'^(' + NUM_PART + r')-(' + NUM_PART + r')$',
                     normalize_num(t))
        if m:
            items.append(('range', (m.group(1), m.group(2)), None))
            i += 1
            continue
        m = re.match(r'^(' + NUM_PART + r')$', normalize_num(t))
        if m:
            pages = None
            if i + 1 < len(tokens) and re.match(
                    r'^[сp]{1,2}\.\s*.+$', tokens[i + 1], re.I):
                pages = tokens[i + 1]
                i += 2
            else:
                i += 1
            items.append(('single', m.group(1), pages))
            continue
        i += 1
    return items


def build_source_refs(inner, bm_map, rPr):
    items = parse_source_inner(inner.strip())
    if not items:
        return None, "cannot parse"
    elems = []

    sep_list = ', '
    sm = re.search(r'\d+(?:\.\d+)?(\s*,\s*)(?:\[\s*)?\d+(?:\.\d+)?', inner)
    if sm:
        sep_list = sm.group(1)

    sep_pages = ', '
    sm = re.search(r'\d+(?:\.\d+)?(\s*,\s*)(?:с\.|p\.)', inner, re.I)
    if sm:
        sep_pages = sm.group(1)

    for idx, (kind, val, pages) in enumerate(items):
        if idx > 0:
            elems.append(make_run(sep_list, rPr))
        if kind == 'range':
            k1, k2 = val
            bm1 = _resolve_key(k1, bm_map)
            bm2 = _resolve_key(k2, bm_map)
            if bm1 and bm2:
                sep_dash = '-'
                sm = re.search(r'\d+(?:\.\d+)?(\s*[-–—−]\s*)\d+(?:\.\d+)?', inner)
                if sm:
                    sep_dash = sm.group(1)
                elems.extend(make_field_runs(f'REF {bm1} \\h', k1, rPr))
                elems.append(make_run(sep_dash, rPr))
                elems.extend(make_field_runs(f'REF {bm2} \\h', k2, rPr))
            else:
                return None, f"no key {k1} or {k2}"
        else:
            bm = _resolve_key(val, bm_map)
            if bm:
                elems.extend(make_field_runs(f'REF {bm} \\h', val, rPr))
                if pages:
                    elems.append(make_run(sep_pages + pages, rPr))
            else:
                return None, f"no key {val}"
    return elems, None


def process_references(paragraph, bm_fig, bm_tab, bm_eq, bm_src,
                       skip_fig_tab_eq=False, log=None):
    text = paragraph.text
    if not text.strip():
        return 0
    if get_caption_re().match(text):
        skip_fig_tab_eq = True

    # Готовые поля Word не трогаем: иначе REF/SEQ вкладываются в них.
    field_ranges = existing_field_ranges(paragraph)

    def in_field(s, e):
        return any(s < fe and fs < e for fs, fe in field_ranges)

    replacements = []

    def collect(regex, bm_map, kind):
        for m in regex.finditer(text):
            num_str = m.group('num')
            s, e = m.start('num'), m.end('num')
            if in_field(s, e):
                continue
            rPr = rPr_at_position(paragraph, s)
            elems = build_number_refs(num_str, bm_map, rPr, placeholder=num_str)
            if elems:
                replacements.append((s, e, elems, kind, num_str))

    if not skip_fig_tab_eq:
        collect(get_fig_ref_re(), bm_fig, 'fig')
        collect(get_tab_ref_re(), bm_tab, 'tab')
        collect(get_eq_ref_re(),  bm_eq,  'eq')

    for m in SRC_REF_RE.finditer(text):
        inner = m.group(1)
        s, e = m.start(1), m.end(1)
        if in_field(s, e):
            continue
        rPr = rPr_at_position(paragraph, s)
        elems, reason = build_source_refs(inner, bm_src, rPr)
        if elems:
            replacements.append((s, e, elems, 'src', inner))

    if not replacements:
        return 0
    replacements.sort(key=lambda x: x[0])
    filtered = [replacements[0]]
    for r in replacements[1:]:
        if r[0] >= filtered[-1][1]:
            filtered.append(r)

    made = 0
    for s, e, elems, kind, num_str in sorted(filtered, key=lambda x: -x[0]):
        if replace_span(paragraph, s, e, elems):
            made += 1
    return made


def process_document(input_path, output_path, log=print, dry_run=False,
                     num_mode=None):
    doc = Document(input_path)
    mode = num_mode if num_mode in NUM_MODES else config.current_num_mode
    moved, removed = repair_field_bookmarks(doc)
    if moved or removed:
        log('log_repaired', moved=moved, removed=removed)
    counter = get_max_bookmark_id(doc) + 1
    used_names = collect_bookmark_names(doc)
    bm_fig, bm_tab, bm_eq, bm_src = {}, {}, {}, {}
    element_paras = set()
    desired = {}

    log('log_start', name="CRF/ФПС", path=input_path)

    # Абзацы в порядке документа, включая ячейки таблиц.
    all_paras = [Paragraph(p_el, doc._body)
                 for p_el in doc.element.body.iter(qn('w:p'))]
    log('log_paragraphs', count=len(all_paras))

    # Нумерация заголовков проверяется до элементов: номера считаются
    # по исправленной структуре.
    log('log_head_check')
    plan = HeadingPlan(doc)
    log('log_head_total', count=plan.count)
    log_heading_issues(plan, log)
    head_fixed = 0
    if any(kind == 'wrong' for kind, *_rest in plan.issues):
        if dry_run:
            log('log_head_keep')
        else:
            head_fixed = fix_heading_numbers(plan)
            log('log_head_fixed', count=head_fixed)

    numbering = ElementNumbering(doc, mode, log=log, plan=plan)
    log(f'log_num_{numbering.mode}')

    log('log_step1a')
    table_formulas = collect_formula_table_rows(doc)

    log('log_step1b')
    n_paras = 0
    caption_re = get_caption_re()
    for para in all_paras:
        n_paras += 1
        text = para.text
        p_el = para._p

        if p_el in table_formulas:
            num_orig, kind = table_formulas[p_el]
            counter = process_table_formula(
                para, num_orig, kind, counter, bm_eq, numbering, log=log,
                used_names=used_names, desired=desired)
            element_paras.add(id(p_el))
            continue

        m_caption = caption_re.match(text)
        if m_caption:
            label = m_caption.group(1).lower()
            if i18n.current_doc_lang == 'en':
                is_fig = 'fig' in label
            else:
                is_fig = label.startswith('рис')
            seq_name = 'Figure' if is_fig else 'Table'
            prefix = 'Fig' if is_fig else 'Tab'

            bm, key, counter, kind = process_caption(
                para, counter, seq_name, prefix, numbering, log,
                used_names=used_names, desired=desired)
            if bm and key:
                if is_fig:
                    bm_fig[key] = bm
                else:
                    bm_tab[key] = bm
                element_paras.add(id(p_el))
            continue

        if is_formula_paragraph(para):
            bm, key, counter = process_formula(
                para, counter, bm_eq, log=log, used_names=used_names,
                numbering=numbering, desired=desired)
            if key is not None and bm:
                element_paras.add(id(p_el))

    log('log_found_stats', figs=len(bm_fig), tabs=len(bm_tab), eqs=len(bm_eq))

    log('log_step2')
    bm_src, counter = process_bibliography(doc, counter, log=log,
                                           used_names=used_names)

    log('log_step3')
    total_refs = 0
    for para in all_paras:
        skip = id(para._p) in element_paras
        total_refs += process_references(
            para, bm_fig, bm_tab, bm_eq, bm_src,
            skip_fig_tab_eq=skip, log=log)
    log('log_total_refs', refs=total_refs)

    # Имена закладок по номерам элементов: Fig_002_003 для рисунка 2.3.
    rename = {}
    apply_desired_names(doc, desired, rename)
    if numbering.mode == 'global':
        renamed = renumber_bookmarks(doc, rename=rename)
    else:
        # В режиме по разделам порядковые имена Fig_001 разрушают иерархию,
        # поэтому пересчитываются только источники (они всегда сквозные).
        renamed = renumber_bookmarks(doc, prefixes=('Src',), rename=rename)
    _rename_field_targets(doc, rename)
    if renamed:
        log('log_renamed', count=renamed)

    if not dry_run:
        # В полях REF показан кэш результата: обновляем его без F9.
        dropped = drop_empty_ref_fields(doc)
        if dropped:
            log('log_ref_dropped', count=dropped)
        refreshed = refresh_ref_results(doc)
        if refreshed:
            log('log_ref_refreshed', count=refreshed)

        try:
            doc.save(output_path)
            log('log_saved', path=output_path)
        except PermissionError:
            raise
    else:
        log('log_not_saved')

    return {
        'figures': len(bm_fig), 'tables': len(bm_tab),
        'equations': len(bm_eq), 'sources': len(bm_src),
        'paragraphs': n_paras, 'refs': total_refs,
        'head_fixed': head_fixed,
    }