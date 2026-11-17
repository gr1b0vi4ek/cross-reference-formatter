"""Низкоуровневые операции с OOXML: runs, fields, bookmarks."""

import copy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


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


def _find_existing_bookmark(paragraph, prefix):
    for bs in paragraph._p.findall('.//' + qn('w:bookmarkStart')):
        name = bs.get(qn('w:name'))
        if name and name.startswith(prefix):
            return name
    return None


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